"""Local-only HTTP app. No document or case data persisted by server."""
import base64
import json
import re
from email.parser import BytesParser
from email.policy import default
from http.server import BaseHTTPRequestHandler,ThreadingHTTPServer
from pathlib import Path
from engine import generate,preview_html
from word_preview import render_pdf

ROOT=Path(__file__).resolve().parent
MAX_SIZE=12*1024*1024

class Handler(BaseHTTPRequestHandler):
    def headers_no_store(self,size,mime):
        self.send_header("Content-Type",mime)
        self.send_header("Content-Length",str(size))
        self.send_header("Cache-Control","no-store")
        self.send_header("X-Content-Type-Options","nosniff")

    def respond(self,status,body):
        payload=json.dumps(body).encode("utf-8")
        self.send_response(status)
        self.headers_no_store(len(payload),"application/json; charset=utf-8")
        self.end_headers()
        self.wfile.write(payload)

    def do_GET(self):
        if self.path not in ("/","/index.html"):
            self.send_error(404);return
        body=(ROOT/"index.html").read_bytes()
        self.send_response(200)
        self.headers_no_store(len(body),"text/html; charset=utf-8")
        self.send_header("Content-Security-Policy",
                         "default-src 'none'; style-src 'unsafe-inline'; script-src 'unsafe-inline'; "
                         "connect-src 'self'; frame-src blob:; object-src blob:")
        self.end_headers()
        self.wfile.write(body)

    def do_POST(self):
        if self.path!="/preview":
            self.send_error(404);return
        origin=self.headers.get("Origin")
        if origin and origin not in ("http://127.0.0.1:8765","http://localhost:8765"):
            self.respond(403,{"error":"Invalid origin"});return
        try:
            size=int(self.headers.get("Content-Length","0"))
            if size<1 or size>MAX_SIZE:raise ValueError("Request too large or empty")
            content_type=self.headers.get("Content-Type","")
            if not content_type.startswith("multipart/form-data"):
                raise ValueError("Expected form upload")
            raw=self.rfile.read(size)
            msg=BytesParser(policy=default).parsebytes(
                b"Content-Type: "+content_type.encode()+b"\r\nMIME-Version: 1.0\r\n\r\n"+raw)
            fields={};docx=None
            for part in msg.iter_parts():
                match=re.search(r'name="([^"]+)"',part.get("Content-Disposition",""))
                if not match:continue
                key=match.group(1)
                payload=part.get_payload(decode=True)
                if key=="document":docx=payload
                else:fields[key]=payload.decode("utf-8").strip()
            if not docx or not docx.startswith(b"PK"):
                raise ValueError("Select a DOCX file from your computer")
            generated,report=generate(docx,fields)
            pdf,pdf_error=render_pdf(generated)
            result={"docx_base64":base64.b64encode(generated).decode("ascii"),
                    "preview_html":preview_html(generated),
                    "pdf_base64":pdf,"pdf_error":pdf_error,"report":report}
            self.respond(200,result)
        except Exception as e:
            self.respond(400,{"error":str(e)})

if __name__=="__main__":
    print("IntelliRequest Community: http://127.0.0.1:8765")
    ThreadingHTTPServer(("127.0.0.1",8765),Handler).serve_forever()
