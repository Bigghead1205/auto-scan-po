
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkcalendar import DateEntry
from pathlib import Path
import threading
import pythoncom
import win32com.client
from m01_email_reader import read_po_emails_and_save_pdfs
from m02_pdf_scan import process_po_pdfs, merge_thread_logs
from m03_send_request_email import send_email_outlook, load_log
import pandas as pd
import os
from config import TEMP_DIR, LOG_DIR, MAX_WORKERS
from datetime import datetime
import concurrent.futures
import time

ENTITY_SHORT_NAMES = {
    "GREEN PLANET DISTRIBUTION CENTRE COMPANY LIMITED": "GREEN PLANET",
    "TECHTRONIC TOOLS (VIETNAM) COMPANY LIMITED": "TTI TOOLS",
    "TECHTRONIC PRODUCTS (VIETNAM) COMPANY LIMITED": "TTI PRODUCTS",
    "TECHTRONIC INDUSTRIES VIETNAM MANUFACTURING COMPANY LIMITED": "TTIVN MFG",
    "TECHTRONIC INDUSTRIES VIETNAM MANUFACTURING COMPANY LIMITED – BRANCH IN DAU GIAY INDUSTRIAL PARK": "TTIVN MFG - CNDG"
}

class POApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PO Classifier Tool")

        self.top_frame = tk.Frame(root)
        self.top_frame.pack(fill="x", padx=10, pady=5)

        self.input_frame = tk.LabelFrame(self.top_frame, text="User Configuration")
        self.input_frame.pack(side="left", fill="both", expand=True, padx=5, pady=5)

        tk.Label(self.input_frame, text="User Account:").grid(row=0, column=0, sticky="e")
        self.user_account_var = tk.StringVar()
        tk.Entry(self.input_frame, textvariable=self.user_account_var, width=40).grid(row=0, column=1, padx=5, pady=2)

        tk.Label(self.input_frame, text="User Email:").grid(row=1, column=0, sticky="e")
        self.user_email_var = tk.StringVar()
        tk.Entry(self.input_frame, textvariable=self.user_email_var, width=40).grid(row=1, column=1, padx=5, pady=2)

        tk.Label(self.input_frame, text="Email Folder Path:").grid(row=2, column=0, sticky="e")
        self.folder_path_var = tk.StringVar()
        tk.Entry(self.input_frame, textvariable=self.folder_path_var, width=40).grid(row=2, column=1, padx=5, pady=2)

        tk.Label(self.input_frame, text="Max Emails per Run:").grid(row=3, column=0, sticky="e")
        tk.Label(self.input_frame, text="From Date:").grid(row=3, column=2, sticky="e")
        self.max_emails_var = tk.IntVar(value=100)
        tk.Entry(self.input_frame, textvariable=self.max_emails_var, width=10).grid(row=3, column=1, sticky="w", padx=5)
        self.from_date_var = tk.StringVar()
        self.from_date_var = tk.StringVar()
        self.date_entry = DateEntry(
            self.input_frame,
            textvariable=self.from_date_var,
            date_pattern="yyyy-mm-dd",
            width=12
        )
        self.date_entry.delete(0, "end")  # ✅ Cho phép trống
        self.date_entry.grid(row=3, column=3, padx=5)
        tk.Button(self.input_frame, text="Clear", command=lambda: self.date_entry.delete(0, "end")).grid(row=3, column=4, padx=5)
        # Cho phép xóa ngày khi người dùng backspace hoặc clear nội dung
        def on_date_focus_out(event):
            if not self.from_date_var.get().strip():
                self.date_entry.delete(0, "end")

        def allow_delete(event):
            if event.keysym in ("BackSpace", "Delete"):
                self.date_entry.delete(0, "end")

        self.date_entry.bind("<FocusOut>", on_date_focus_out)
        self.date_entry.bind("<Key>", allow_delete)

        tk.Label(self.input_frame, text="Output Folder:").grid(row=4, column=0, sticky="e")
        self.output_folder_var = tk.StringVar()
        tk.Entry(self.input_frame, textvariable=self.output_folder_var, width=40).grid(row=4, column=1, padx=5, pady=2)
        tk.Button(self.input_frame, text="Browse", command=self.browse_output_folder).grid(row=4, column=2)
        tk.Button(self.input_frame, text="Fetch Emails", command=self.fetch_emails).grid(row=4, column=3, padx=5)

        self.email_frame = tk.LabelFrame(self.top_frame, text="Send Request Email")
        self.email_frame.pack(side="right", fill="y", padx=5, pady=5)

        tk.Label(self.email_frame, text="Filter by Entity:").pack(side="top", padx=5, anchor="w")
        self.entity_filter_var = tk.StringVar()
        self.entity_filter_combo = ttk.Combobox(self.email_frame, textvariable=self.entity_filter_var)
        self.entity_filter_combo['values'] = ["All"] + list(ENTITY_SHORT_NAMES.values())
        self.entity_filter_combo.current(0)
        self.entity_filter_combo.pack(padx=5, pady=2, fill="x")

        tk.Button(self.email_frame, text="Send Email for Selected", command=self.send_email_selected).pack(pady=2, fill="x", padx=5)

        self.summary_frame = tk.LabelFrame(root, text="Summary of PO Scan")
        self.summary_frame.pack(fill="x", padx=10, pady=5)

        self.summary_text = tk.Text(self.summary_frame, height=8, wrap="word", state="disabled", bg=self.root.cget("bg"), relief="flat")
        self.summary_text.pack(fill="x", padx=10, pady=5)

        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        tk.Label(root, textvariable=self.status_var).pack(pady=2)

        self.email_results = []
        self.output_base_path = None

    def browse_output_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.output_folder_var.set(folder_selected)

    def fetch_emails(self):
        self.output_base_path = Path(self.output_folder_var.get())
        threading.Thread(target=self._fetch_emails_thread).start()

    def _fetch_emails_thread(self):
        pythoncom.CoInitialize()
        try:
            start = time.perf_counter()
            max_emails = self.max_emails_var.get()
            from_date_str = self.from_date_var.get().strip()
            from_date = datetime.strptime(from_date_str, "%Y-%m-%d") if from_date_str else None


            email_account = self.user_email_var.get().strip()
            folder_path_str = self.folder_path_var.get().strip()
            folder_path = [seg.strip() for seg in folder_path_str.split(">") if seg.strip()]

            self.status_var.set("📥 Fetching emails and saving PDFs...")
            temp_dir = self.output_base_path / "temp"

            self.email_results = read_po_emails_and_save_pdfs(
                temp_dir,
                email_account=email_account,
                folder_path=folder_path,
                max_emails=max_emails,
                from_date=from_date
            )

            self.status_var.set("📄 Scanning PDF content...")
            process_po_pdfs(self.email_results, self.output_base_path)

            self.status_var.set("📝 Merging thread logs...")
            log_path, count = merge_thread_logs(self.output_base_path)
            df_log = pd.read_csv(log_path, dtype=str)

            self.status_var.set("📊 Generating summary...")
            cd_needed = df_log[df_log["Need_CDs"] == "Yes"]
            entity_counts = {}
            for _, row in cd_needed.iterrows():
                buyer = row.get("Buyer", "").strip()
                entity = ENTITY_SHORT_NAMES.get(buyer, buyer if buyer else "Unknown")
                entity_counts[entity] = entity_counts.get(entity, 0) + 1

            elapsed = time.perf_counter() - start
            summary = f"✅ Time Elapsed: {elapsed:.1f}s\nPO total: {len(self.email_results)}\nPO CDs Required:\n"
            summary += "\n".join(f"{k}: {v}" for k, v in entity_counts.items()) if entity_counts else "(None)"
            self.summary_text.config(state="normal")
            self.summary_text.delete("1.0", tk.END)
            self.summary_text.insert(tk.END, summary)
            self.summary_text.config(state="disabled")
            self.status_var.set("✅ Done.")

        except Exception as e:
            self.status_var.set(f"Error: {e}")


        except Exception as e:
            self.status_var.set(f"Error: {e}")

    def send_email_selected(self):
        self.output_base_path = Path(self.output_folder_var.get())
        df = load_log(self.output_base_path)
        if df is None:
            messagebox.showerror("Log Missing", "Log file not found.")
            return

        selected_entity = self.entity_filter_var.get().strip().upper()
        df_filtered = df[(df["Need_CDs"] == "Yes") & (df["Email Request Info"] != "Yes")]

        sent = 0
        for _, row in df_filtered.iterrows():
            short_name = ENTITY_SHORT_NAMES.get(row.get("Buyer", ""), "").upper()
            if selected_entity != "ALL" and selected_entity != short_name:
                continue

            result = send_email_outlook(row, self.output_base_path)
            if result:  # only True if mail.Send() succeeds
                sent += 1
                df.loc[(df["PO Number"] == row["PO Number"]), "Email Request Info"] = "Yes"

        # ⏳ Đảm bảo log được cập nhật sau vòng lặp
        log_file = self.output_base_path / "log" / "po_log.csv"
        df.to_csv(log_file, index=False, encoding="utf-8", quoting=1)

        messagebox.showinfo("Done", f"Sent {sent} emails.")
        self.status_var.set(f"Sent {sent} emails.")

if __name__ == "__main__":
    root = tk.Tk()
    app = POApp(root)
    root.geometry("780x520")
    root.mainloop()
