"""Select one email in classic Outlook, then run this file."""

from pywintypes import com_error
from win32com.client import GetActiveObject


try:
    outlook = GetActiveObject("Outlook.Application")
    explorer = outlook.ActiveExplorer()
    if explorer is None or explorer.Selection.Count == 0:
        raise RuntimeError("Select one email in Outlook first.")

    mail = explorer.Selection.Item(1)
    print("Subject:", mail.Subject)
    print("Reading protected properties (Guard should appear)...")
    print("Sender:", mail.SenderEmailAddress)
    print("Body:", mail.Body[:200].replace("\r", " ").replace("\n", " "))
except com_error as error:
    print(f"Outlook blocked/failed: 0x{error.hresult & 0xFFFFFFFF:08X}")
except RuntimeError as error:
    print(error)
