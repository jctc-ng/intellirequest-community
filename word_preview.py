"""Optional private, local Word-to-PDF preview. Requires Microsoft Word on Windows."""
import base64
import sys
import tempfile
from pathlib import Path

def render_pdf(document_bytes):
    if sys.platform!="win32":
        return None,"High-fidelity preview needs Windows desktop Microsoft Word"
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        return None,"Install pywin32 for Microsoft Word preview"
    with tempfile.TemporaryDirectory(prefix="intellirequest_") as folder:
        original=Path(folder)/"preview.docx"
        target=Path(folder)/"preview.pdf"
        original.write_bytes(document_bytes)
        pythoncom.CoInitialize()
        app=None;doc=None
        try:
            app=win32com.client.DispatchEx("Word.Application")
            app.Visible=False
            app.DisplayAlerts=0
            app.AutomationSecurity=3
            doc=app.Documents.Open(FileName=str(original),ReadOnly=True,
                                   AddToRecentFiles=False,ConfirmConversions=False)
            doc.ExportAsFixedFormat(OutputFileName=str(target),ExportFormat=17)
            return base64.b64encode(target.read_bytes()).decode("ascii"),None
        except Exception as exc:
            return None,"Microsoft Word could not render the document: "+str(exc)
        finally:
            if doc is not None:
                try:doc.Close(SaveChanges=0)
                except Exception:pass
            if app is not None:
                try:app.Quit(SaveChanges=0)
                except Exception:pass
            pythoncom.CoUninitialize()
