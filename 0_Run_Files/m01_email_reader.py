import win32com.client
import os
import gc
import pythoncom
from datetime import datetime

# Note: PDF scanning and classification have been moved to m02_pdf_scan.process_po_pdfs
# to avoid redundant work. read_po_emails_and_save_pdfs now only downloads PDF
# attachments and collects basic email metadata. See improvement.txt items 1–3, 12.

from utils import resolve_email  # moved to utils.py to avoid duplication

def read_po_emails_and_save_pdfs(save_folder, email_account, folder_path, max_emails: int = 100, from_date: datetime | None = None):
    """
    Download PDF attachments from unread emails in the specified Outlook folder.

    This function no longer scans PDF content; it focuses solely on downloading
    attachments and collecting basic metadata (recipient emails, subject, etc.).
    Scanning and classification are handled in m02_pdf_scan.process_po_pdfs.
    See improvement.txt items 3 and 12.

    Parameters
    ----------
    save_folder : str or Path
        Directory where attachments will be saved.
    email_account : str
        Outlook account to access (e.g. "DinhLong.Bui@ttigroup.com.vn").
    folder_path : list[str]
        Path segments under the account to reach the target folder (e.g. ["CUS", "CUS MACHINE", "ERP PO"]).
    max_emails : int, optional
        Maximum number of unread emails to process.
    from_date : datetime, optional
        Only process emails received on or after this date.

    Returns
    -------
    list[dict]
        A list of dictionaries containing basic metadata for each downloaded PDF.
    """
    import os
    import gc
    import pythoncom
    import win32com.client

    os.makedirs(save_folder, exist_ok=True)

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    folder = outlook.Folders.Item(email_account)
    for name in folder_path:
        folder = folder.Folders(name)
    erp_po_folder = folder

    messages = erp_po_folder.Items
    messages.Sort("[ReceivedTime]", True)

    # Filter by date if provided
    if from_date:
        naive_from_date = from_date.replace(tzinfo=None)
        filtered_messages = [msg for msg in messages if msg.UnRead and msg.ReceivedTime.replace(tzinfo=None) >= naive_from_date]
    else:
        filtered_messages = [msg for msg in messages if msg.UnRead]
    unread_messages = filtered_messages[:max_emails]

    results: list[dict] = []

    for msg in unread_messages:
        try:
            attachments = msg.Attachments
            for i in range(attachments.Count):
                attachment = attachments.Item(i + 1)
                # Only handle PDF attachments
                if not attachment.FileName.lower().endswith(".pdf"):
                    continue

                original_name = attachment.FileName
                base_name, ext = os.path.splitext(original_name)
                save_path = os.path.join(str(save_folder), original_name)
                count = 1
                # Ensure unique filename to avoid overwriting existing files
                while os.path.exists(save_path):
                    save_path = os.path.join(str(save_folder), f"{base_name}_{count}{ext}")
                    count += 1
                attachment.SaveAsFile(save_path)

                # Resolve TO recipients once per message
                to_emails: list[str] = []
                for j in range(msg.Recipients.Count):
                    recipient = msg.Recipients.Item(j + 1)
                    if recipient.Type == 1:  # To
                        email = resolve_email(recipient, outlook)
                        if email and "@" in email:
                            to_emails.append(email)

                to_email_str = " / ".join(to_emails)
                received_time_str = msg.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S")

                results.append({
                    "file_name": os.path.basename(save_path),
                    "to_emails": to_email_str,
                    "pdf_path": save_path,
                    "received_time": received_time_str,
                    "subject": msg.Subject,
                })

        except Exception as e:
            # Wrap each email in try/except to avoid batch failure (improvement 10)
            print(f"❌ Lỗi xử lý email {getattr(msg, 'Subject', '')}: {e}")
        finally:
            try:
                msg.UnRead = False  # mark as read
            except Exception:
                pass
            pythoncom.CoFreeUnusedLibraries()
            gc.collect()

    return results
