import re
import csv
import pdfplumber
import threading
import concurrent.futures
import pandas as pd
import shutil
from pathlib import Path
from datetime import datetime
from config import NON_CDS_SUPPLIER_FILE, MAX_WORKERS

thread_local = threading.local()

def clean_cell(val):
    if isinstance(val, str):
        return val.replace("\n", " ").replace("\r", " ").strip()
    return val

def classify_buyer(buyer_text):
    text = buyer_text.upper()
    if "GREEN PLANET" in text:
        return "GREEN PLANET DISTRIBUTION CENTRE COMPANY LIMITED"
    elif "TECHTRONIC TOOLS" in text:
        return "TECHTRONIC TOOLS (VIETNAM) COMPANY LIMITED"
    elif "TECHTRONIC PRODUCTS" in text:
        return "TECHTRONIC PRODUCTS (VIETNAM) COMPANY LIMITED"
    elif "BRANCH IN DAU GIAY" in text:
        return "TECHTRONIC INDUSTRIES VIETNAM MANUFACTURING COMPANY LIMITED – BRANCH IN DAU GIAY INDUSTRIAL PARK"
    elif "TECHTRONIC INDUSTRIES VIETNAM" in text:
        return "TECHTRONIC INDUSTRIES VIETNAM MANUFACTURING COMPANY LIMITED"
    return "Unknown"

def get_buyer_folder_name(buyer_text):
    text = buyer_text.upper()
    if "GREEN PLANET" in text:
        return "2. GREEN PLANET"
    elif "TECHTRONIC TOOLS" in text:
        return "5. TTI TOOLS"
    elif "TECHTRONIC PRODUCTS" in text:
        return "4. TTI PRODUCTS"
    elif "BRANCH IN DAU GIAY" in text:
        return "3. TTIVN MFG - CNDG"
    elif "TECHTRONIC INDUSTRIES VIETNAM" in text:
        return "1. TTIVN MFG"
    return "Unknown"

def extract_po_number(text):
    match = re.search(r"PO#:\s*(\d{6,})", text, re.IGNORECASE)
    return match.group(1).strip() if match else "Unknown"

def extract_seller_name(text):
    match = re.search(r"SELLER:\s*(.*?)\s*BUYER:", text, re.IGNORECASE | re.DOTALL)
    if match:
        seller_block = match.group(1).strip()
        lines = [line.strip() for line in seller_block.splitlines() if line.strip()]
        return " ".join(lines)
    return "Unknown"

def extract_vat_from_table(text):
    matches = re.findall(r"(\d{1,2})\s*%\s+\d", text)
    return "/".join(sorted({m + "%" for m in matches})) if matches else "Unknown"

def extract_currency_from_table(text):
    currencies = re.findall(r"\b(VND|USD|EUR|JPY)\b", text, re.IGNORECASE)
    valid_currencies = [c.upper() for c in currencies if len(c) == 3]
    return "/".join(sorted(set(valid_currencies))) if valid_currencies else "Unknown"

def extract_uom_from_table(pdf):
    for page in pdf.pages:
        try:
            tables = page.extract_tables()
            for table in tables:
                header = table[0]
                if not header:
                    continue
                for i, h in enumerate(header):
                    if h and "uom" in h.lower():
                        col_index = i
                        uoms = set()
                        for row in table[1:]:
                            if len(row) > col_index:
                                val = row[col_index]
                                if val and len(val.strip()) <= 10:
                                    uoms.add(val.strip())
                        return "/".join(sorted(uoms)) if uoms else "Unknown"
        except:
            continue
    return "Unknown"

def extract_max_unit_price_from_table(pdf):
    prices = []
    for page in pdf.pages:
        try:
            tables = page.extract_tables()
            for table in tables:
                header = table[0]
                if not header:
                    continue
                col_index = None
                for i, h in enumerate(header):
                    if h and "unit" in h.lower() and "price" in h.lower():
                        col_index = i
                        break
                if col_index is None:
                    continue
                for row in table[1:]:
                    try:
                        cell = row[col_index]
                        if cell:
                            val = float(cell.replace(",", "").replace(" ", ""))
                            if val > 0:
                                prices.append(val)
                    except:
                        continue
        except:
            continue
    return max(prices) if prices else 0

def extract_end_user_email(text):
    match = re.search(r"[A-Za-z0-9._%+-]+@ttigroup\.com\.vn", text, re.IGNORECASE)
    return match.group(0).strip() if match else ""

def determine_need_cds(vat: str, currency: str, uom: str, seller: str, max_unit_price: float) -> str:
    """Determine whether a PO requires a customs declaration sheet (CDs).

    This function follows the decision tree defined in ``classify logic.txt``:

    * If currency is not ``VND`` → return ``"Yes"`` (CDs needed).
    * Else (currency == ``VND``):
      - If VAT rates do not include ``0%`` → return ``"No"``.
      - If UOM equals ``UNIT`` or its abbreviations (``UN``, ``UNT``) → return ``"No"``.
      - If the seller appears in the ``non_cds_sellers`` list → return ``"No"``.
      - If ``max_unit_price`` > 30,000,000 → return ``"Yes"``.
      - If UOM does not include either ``PIECE`` or ``SET`` → return ``"No"``.
      - Otherwise → return ``"Yes"``.

    Parameters
    ----------
    vat : str
        VAT rate string, e.g. ``"10%/0%"``.
    currency : str
        Currency code string, e.g. ``"VND"`` or ``"USD"``.
    uom : str
        Unit of measure string from the PO.
    seller : str
        Seller name from the PO.
    max_unit_price : float
        Maximum unit price extracted from the PO.

    Returns
    -------
    str
        ``"Yes"`` if CDs are required, otherwise ``"No"``.
    """
    # Cache the non-CDs supplier list to avoid re-reading the CSV on every call.
    # The list is stored on the module object so it persists across calls.
    global _NON_CDS_SUPPLIER_CACHE
    try:
        cache = _NON_CDS_SUPPLIER_CACHE
    except NameError:
        cache = None

    if cache is None:
        non_cds_sellers: set[str] = set()
        supplier_path = Path(NON_CDS_SUPPLIER_FILE)
        if supplier_path.exists():
            try:
                df_sup = pd.read_csv(supplier_path, dtype=str)
                non_cds_sellers = set(df_sup.iloc[:, 0].str.upper().str.strip())
            except Exception:
                # If reading fails, leave the set empty and continue
                pass
        _NON_CDS_SUPPLIER_CACHE = non_cds_sellers
    else:
        non_cds_sellers = cache

    seller_clean = (seller or "").upper().strip()
    vat_clean = (vat or "").strip()
    currency_clean = (currency or "").strip().upper()
    uom_clean = (uom or "").strip().upper()

    def is_all_non_zero(vat_str: str) -> bool:
        # Split by slash, comma or semicolon and ensure every rate is not 0%
        rates = [v.strip() for v in re.split(r"[\\/,;]", vat_str) if v.strip()]
        return bool(rates) and all(rate != "0%" for rate in rates)

    def uom_contains_any(uom_str: str, valid_uoms: set[str]) -> bool:
        parts = [u.strip().upper() for u in re.split(r"[\\/,;]", uom_str) if u.strip()]
        return any(u in valid_uoms for u in parts)

    # Currency check
    if currency_clean and currency_clean != "VND":
        return "Yes"
    # VAT check
    if vat_clean and is_all_non_zero(vat_clean):
        return "No"
    # UOM = UNIT check
    if uom_clean in {"UNIT", "UN", "UNT"}:
        return "No"
    # Seller exclusion check
    if seller_clean and seller_clean in non_cds_sellers:
        return "No"
    # High unit price check
    try:
        if float(max_unit_price) > 30_000_000:
            return "Yes"
    except Exception:
        pass
    # UOM must contain PIECE or SET
    if not uom_contains_any(uom_clean, {"PIECE", "SET"}):
        return "No"
    # Default to requiring CDs
    return "Yes"

def process_po_pdfs(email_results: list[dict], output_base_dir: Path):
    """
    Scan downloaded PO PDFs, classify whether CDs are needed and update the log.

    This implementation uses a thread pool to parallelize PDF parsing. Results
    are collected in memory and written to disk once at the end, reducing
    contention and I/O overhead (improvement items 1–3). Any errors
    encountered while processing a PDF are recorded in ``log/error.txt``.

    Parameters
    ----------
    email_results : list[dict]
        A list of dictionaries returned by ``read_po_emails_and_save_pdfs``.
    output_base_dir : Path
        The base directory where ``log`` and ``PO_Filtered`` folders reside.
    """
    LOG_DIR = output_base_dir / "log"
    FILTERED_DIR = output_base_dir / "PO_Filtered"
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    FILTERED_DIR.mkdir(parents=True, exist_ok=True)

    thread_id = threading.get_ident()
    log_path = LOG_DIR / f"thread_{thread_id}.csv"
    error_log_path = LOG_DIR / "error.txt"

    COLUMNS = [
        "PO Number", "Buyer", "Seller", "VAT", "Currency", "UOM",
        "Max Unit Price", "Need_CDs", "Supplier/Vendor email", "End-User Email", "ReceivedTime"
    ]
    df_log = pd.read_csv(log_path, dtype=str) if log_path.exists() else pd.DataFrame(columns=COLUMNS)

    # Preload existing PO numbers to detect revisions
    existing_po_numbers = set(df_log["PO Number"].values)

    results: list[dict] = []

    def process_one(res: dict) -> dict | None:
        pdf_path = Path(res.get("pdf_path"))
        if not pdf_path.exists():
            return None
        try:
            with pdfplumber.open(pdf_path) as pdf:
                # Extract text once for reuse
                text = "\n".join(page.extract_text() or "" for page in pdf.pages)
                # Use extracted text for all functions that operate on strings
                uom = extract_uom_from_table(pdf)
                max_unit_price = extract_max_unit_price_from_table(pdf)
                end_user_email = extract_end_user_email(text)

            po_number = extract_po_number(text)
            buyer = classify_buyer(text.splitlines()[0] if text else "")
            seller = extract_seller_name(text)
            vat = extract_vat_from_table(text)
            currency = extract_currency_from_table(text)
            need_cds = determine_need_cds(vat, currency, uom, seller, max_unit_price)

            # Prepare rename destination if needed
            rename_dest: Path | None = None
            if need_cds == "Yes":
                folder_name = get_buyer_folder_name(buyer)
                dest = FILTERED_DIR / folder_name
                dest.mkdir(parents=True, exist_ok=True)
                new_path = dest / pdf_path.name
                if new_path.exists():
                    new_path = new_path.with_name(
                        f"{new_path.stem}_{datetime.now():%Y%m%d%H%M%S}{new_path.suffix}"
                    )
                rename_dest = new_path

            return {
                "po_number": po_number,
                "buyer": buyer,
                "seller": seller,
                "vat": vat,
                "currency": currency,
                "uom": uom,
                "max_unit_price": max_unit_price,
                "need_cds": need_cds,
                "to_emails": res.get("to_emails", ""),
                "received_time": res.get("received_time", ""),
                "end_user_email": end_user_email,
                "pdf_path": pdf_path,
                "rename_dest": rename_dest,
            }
        except Exception as e:
            # Capture any error and log it for troubleshooting (improvement 11)
            with error_log_path.open("a", encoding="utf-8") as err_file:
                err_file.write(f"{pdf_path}: {e}\n")
            return None

    # Parallel processing of PDFs
    with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        future_to_result = {executor.submit(process_one, res): res for res in email_results}
        for future in concurrent.futures.as_completed(future_to_result):
            processed = future.result()
            if processed:
                results.append(processed)

    # Update log based on processed results
    for item in results:
        po_number = item["po_number"] or "Unknown"
        need_cds = item["need_cds"]
        # Remove existing entry if PO number already exists
        if po_number in existing_po_numbers:
            need_cds = "Revised"
            df_log = df_log[df_log["PO Number"] != po_number]
        # Standardize email lists with semicolons
        raw_to = clean_cell(item.get("to_emails", ""))
        to_emails = [e.strip() for e in re.split(r"[;/]", raw_to) if "@" in e]
        to_email_str = "; ".join(to_emails)

        raw_cc = clean_cell(item.get("end_user_email", ""))
        cc_emails = [e.strip() for e in re.split(r"[;/]", raw_cc) if "@" in e]
        cc_email_str = "; ".join(cc_emails)

        new_row = {
            "PO Number": clean_cell(po_number),
            "Buyer": clean_cell(item["buyer"]),
            "Seller": clean_cell(item["seller"]),
            "VAT": clean_cell(item["vat"]),
            "Currency": clean_cell(item["currency"]),
            "UOM": clean_cell(item["uom"]),
            "Max Unit Price": item["max_unit_price"],
            "Need_CDs": clean_cell(need_cds),
            "Supplier/Vendor email": to_email_str,
            "End-User Email": cc_email_str,
            "ReceivedTime": clean_cell(item["received_time"]),
        }
        # Append to DataFrame
        df_log = pd.concat(
            [df_log, pd.DataFrame([new_row], columns=df_log.columns)],
            ignore_index=True,
            sort=False,
            copy=False
        )

        # Rename the PDF if necessary
        dest = item.get("rename_dest")
        if dest:
            try:
                item["pdf_path"].rename(dest)
            except Exception as e:
                # Log rename errors but continue
                with error_log_path.open("a", encoding="utf-8") as err_file:
                    err_file.write(f"Rename error for {item['pdf_path']}: {e}\n")

    # Persist log to CSV once
    df_log.to_csv(log_path, index=False, encoding="utf-8", quoting=csv.QUOTE_NONNUMERIC)

    # Remove temporary files after processing
    temp_folder = output_base_dir / "temp"
    if temp_folder.exists():
        try:
            shutil.rmtree(temp_folder)
        except Exception as e:
            print(f"⚠️ Không thể xóa thư mục tạm: {e}")

def merge_thread_logs(output_base_dir):
    log_dir = Path(output_base_dir) / "log"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_files = list(log_dir.glob("thread_*.csv"))

    final_path = log_dir / "po_log.csv"
    df_final = pd.DataFrame()

    # Đọc log cũ nếu có (read existing final log if any)
    if final_path.exists():
        df_final = pd.read_csv(final_path, dtype=str)

    # Merge new thread logs
    if log_files:
        df_list = [pd.read_csv(f, dtype=str) for f in log_files]
        df_threads = pd.concat(df_list, ignore_index=True, sort=False, copy=False)
        df_all = pd.concat([df_final, df_threads], ignore_index=True)
        df_all.drop_duplicates(subset="PO Number", keep="last", inplace=True)
    else:
        df_all = df_final

    df_all.to_csv(final_path, index=False, encoding="utf-8", quoting=csv.QUOTE_NONNUMERIC)

    for f in log_files:
        f.unlink(missing_ok=True)

    return (str(final_path), len(df_all))
