import re
from pathlib import Path
import pandas as pd
import win32com.client  # type: ignore[import]
from jinja2 import Template

from config import TEMPLATE_LOCAL, TEMPLATE_OVERSEA, TEMP_DIR

# --- Email body template ---
EMAIL_BODY_TEMPLATE = Template(
    """
Dear Supplier,

This is an automatic request from our system regarding PO# {{ po }}.

To support customs declaration procedures, please kindly fill in the attached “Cargo Info” file with full and correct product details once the goods are ready and the shipment is being prepared.

If this PO does not involve any item requiring customs clearance, you may kindly ignore this message.

Thank you for your cooperation.

Best regards,

TTIVN Customs Team
    """
)

def is_valid_email(email):
    return isinstance(email, str) and re.match(r"[^@]+@[^@]+\.[^@]+", email.strip())

def load_log(output_base_dir):
    log_file = Path(output_base_dir) / "log" / "po_log.csv"
    if not log_file.exists():
        print("\u26a0\ufe0f Log file không tồn tại.")
        return None
    df = pd.read_csv(log_file, dtype=str)
    df["Email Request Info"] = df.get("Email Request Info", "")
    return df

def filter_po_need_email(df):
    return df[
        (df["Need_CDs"] == "Yes") &
        (df["Email Request Info"] != "Yes") &
        (df["Supplier/Vendor email"].notna())
    ]

def get_attachments(po_number, currency, output_base_dir):
    attachments = []
    template_file = None

    template_path = Path(TEMPLATE_LOCAL if "VND" in currency.upper() else TEMPLATE_OVERSEA)
    if not template_path.exists():
        print(f"❌ Template không tồn tại: {template_path}")
        return attachments, None

    renamed_template = template_path.parent / f"{template_path.stem}_{po_number}{template_path.suffix}"
    try:
        renamed_template.write_bytes(template_path.read_bytes())
        attachments.append(renamed_template)
        template_file = renamed_template
    except Exception as e:
        print(f"❌ Lỗi khi sao chép template: {e}")

    for subdir in (Path(output_base_dir) / "PO_Filtered").iterdir():
        if subdir.is_dir():
            for file in subdir.glob(f"*{po_number}*.pdf"):
                attachments.append(file)
                break
        if len(attachments) >= 2:
            break

    return attachments, template_file

def send_email_outlook(po_row, output_base_dir):
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)

    po_number = po_row["PO Number"]

    raw_to = po_row["Supplier/Vendor email"]
    to_emails = [e.strip() for e in re.split(r"[;/]", raw_to) if is_valid_email(e)]
    if not to_emails:
        print(f"❌ Không có email hợp lệ trong TO cho PO {po_number}: {raw_to}")
        return False
    mail.To = "; ".join(to_emails)

    raw_cc = po_row.get("End-User Email", "")
    cc_emails = [e.strip() for e in re.split(r"[;/]", raw_cc) if is_valid_email(e)]
    mail.CC = "; ".join(cc_emails)

    mail.Subject = f"{po_row['Buyer']}/PO#{po_number}/Cargo info Request"
    mail.Body = EMAIL_BODY_TEMPLATE.render(po=po_number).strip()

    attachments, template_file = get_attachments(po_number, po_row["Currency"], output_base_dir)
    if template_file is None:
        print(f"❌ Không thể đính kèm file template cho PO {po_number}, email sẽ không được gửi.")
        return False

    for file in attachments:
        try:
            if file.exists():
                mail.Attachments.Add(str(file))
        except Exception as e:
            print(f"⚠\ufe0f Không thể đính kèm file: {file} - {e}")

    try:
        mail.Send()
        print(f"✅ Đã gửi email cho PO {po_number}")
    except Exception as e:
        if "moved or deleted" in str(e).lower():
            print(f"✅ Đã gửi email cho PO {po_number} (Outlook đã di chuyển email)")
        else:
            print(f"❌ Lỗi gửi email cho PO {po_number}: {e}")
            return False

    # Clean up temp files
    for file in attachments:
        try:
            if file.exists() and file.parent.resolve() == Path(TEMP_DIR).resolve():
                file.unlink()
        except Exception as e:
            print(f"⚠\ufe0f Không thể xoá file tạm: {file} - {e}")

    return True

def main_send_all(output_base_dir):
    df = load_log(output_base_dir)
    if df is None:
        return

    df_filtered = filter_po_need_email(df)
    if df_filtered.empty:
        print("✅ Không có PO nào cần gửi email.")
        return

    for idx, row in df_filtered.iterrows():
        success = send_email_outlook(row, output_base_dir)
        if success:
            df.loc[(df["PO Number"] == row["PO Number"]), "Email Request Info"] = "Yes"

    log_file = Path(output_base_dir) / "log" / "po_log.csv"
    df.to_csv(log_file, index=False, encoding="utf-8", quoting=1)
    print("📤 Đã cập nhật cột 'Email Request Info' trong log.")

if __name__ == "__main__":
    main_send_all()
