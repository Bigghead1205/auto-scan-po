"""Utility functions shared across modules.

This module consolidates helper functions that were previously duplicated in
multiple modules (see improvement.txt item 8). Keeping these utilities in a
central place simplifies maintenance and encourages reuse.
"""
from __future__ import annotations

import win32com.client  # type: ignore[import]

def resolve_email(recipient, outlook) -> str | None:
    """Resolve an Outlook recipient to its SMTP address.

    Parameters
    ----------
    recipient : COM object
        The recipient object from an Outlook mail item.
    outlook : COM object
        Outlook namespace used to resolve Exchange recipients.

    Returns
    -------
    str | None
        The resolved SMTP address or ``None`` if it cannot be determined.

    Notes
    -----
    Outlook stores recipients as MAPI objects which may represent Exchange
    addresses (type ``EX``) or simple SMTP addresses. This function attempts
    to extract a usable SMTP address by trying several strategies, including
    querying the AddressEntry, using a fallback recipient, and finally
    retrieving the ``PR_SMTP_ADDRESS`` property via PropertyAccessor.
    """
    try:
        entry = recipient.AddressEntry
        if entry:
            if entry.Type == "EX":
                exch = entry.GetExchangeUser()
                if exch and exch.PrimarySmtpAddress:
                    return exch.PrimarySmtpAddress
            elif entry.Type == "SMTP":
                return entry.Address
        # Fallback: resolve via name
        fallback = outlook.CreateRecipient(recipient.Name)
        if fallback.Resolve():
            entry2 = fallback.AddressEntry
            if entry2.Type == "SMTP":
                return entry2.Address
            elif entry2.Type == "EX":
                exch2 = entry2.GetExchangeUser()
                if exch2 and exch2.PrimarySmtpAddress:
                    return exch2.PrimarySmtpAddress
        # Last resort: use PropertyAccessor to get SMTP address
        pa = recipient.PropertyAccessor
        smtp = pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E")
        if smtp and "@" in smtp:
            return smtp
    except Exception as e:
        # Log errors to console; callers decide how to handle None
        print(f"⚠️ Lỗi resolve email: {e}")
    return None

__all__ = ["resolve_email"]