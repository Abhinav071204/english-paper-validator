import os
import json
import urllib.parse
import http.server
import socketserver
from email.parser import BytesParser
from email.policy import default

# Import validator functions
from validator import DocumentValidator, auto_correct_document

PORT = int(os.environ.get("PORT", 8000))
DIRECTORY = r"C:\Users\Riyak\.gemini\antigravity\scratch\english-paper-validator"

class ValidatorHTTPRequestHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=DIRECTORY, **kwargs)

    def end_headers(self):
        self.send_header('Cache-Control', 'no-store, no-cache, must-revalidate, max-age=0')
        self.send_header('Pragma', 'no-cache')
        self.send_header('Expires', '0')
        super().end_headers()

    def do_GET(self):
        parsed_url = urllib.parse.urlparse(self.path)
        path = parsed_url.path
        
        # 1. API: List stored files
        if path == "/api/files":
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            
            try:
                files = [f for f in os.listdir(DIRECTORY) if f.endswith(".docx") and not f.startswith("~$") and not f.endswith("_Corrected.docx")]
                self.wfile.write(json.dumps({"files": files}).encode("utf-8"))
            except Exception as e:
                self.wfile.write(json.dumps({"error": str(e)}).encode("utf-8"))
            return
            
        # 2. API: Validate a stored file
        elif path == "/api/validate_stored":
            query = urllib.parse.parse_qs(parsed_url.query)
            filename = query.get("file", [None])[0]
            
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            
            if not filename:
                self.wfile.write(json.dumps({"error": "Missing 'file' parameter"}).encode("utf-8"))
                return
                
            filepath = os.path.join(DIRECTORY, filename)
            if not os.path.exists(filepath):
                self.wfile.write(json.dumps({"error": f"File '{filename}' does not exist"}).encode("utf-8"))
                return
                
            try:
                validator = DocumentValidator(filepath)
                result = validator.run_validation()
                self.wfile.write(json.dumps(result).encode("utf-8"))
            except Exception as e:
                self.wfile.write(json.dumps({"error": str(e)}).encode("utf-8"))
            return

        # 3. API: Correct and download a stored file
        elif path == "/api/correct_stored":
            query = urllib.parse.parse_qs(parsed_url.query)
            filename = query.get("file", [None])[0]
            
            if not filename:
                self.send_response(400)
                self.end_headers()
                return
                
            input_path = os.path.join(DIRECTORY, filename)
            if not os.path.exists(input_path):
                self.send_response(404)
                self.end_headers()
                return
                
            temp_corrected_path = os.path.join(DIRECTORY, "temp_corrected.docx")
            try:
                auto_correct_document(input_path, temp_corrected_path)
                
                # Stream the corrected file bytes
                with open(temp_corrected_path, "rb") as f:
                    file_bytes = f.read()
                    
                corrected_filename = filename.replace(".docx", "_Corrected.docx")
                
                self.send_response(200)
                self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                self.send_header("Content-Disposition", f"attachment; filename=\"{corrected_filename}\"")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.send_header("Content-Length", str(len(file_bytes)))
                self.end_headers()
                
                self.wfile.write(file_bytes)
            except Exception as e:
                self.send_response(500)
                self.send_header("Content-Type", "application/json")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({"error": str(e)}).encode("utf-8"))
            finally:
                if os.path.exists(temp_corrected_path):
                    try: os.remove(temp_corrected_path)
                    except: pass
            return
            
        # Standard file serving for index.html, style.css, app.js, etc.
        return super().do_GET()

    def do_POST(self):
        parsed_url = urllib.parse.urlparse(self.path)
        path = parsed_url.path
        
        # 4. API: Validate uploaded file
        if path == "/api/validate":
            content_type = self.headers.get('Content-Type', '')
            if not content_type.startswith('multipart/form-data'):
                self.send_response(400)
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                self.wfile.write(json.dumps({"error": "Content-Type must be multipart/form-data"}).encode("utf-8"))
                return

            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)
            
            msg_data = b"Content-Type: " + content_type.encode('utf-8') + b"\r\n\r\n" + body
            msg = BytesParser(policy=default).parsebytes(msg_data)
            
            uploaded_file_data = None
            uploaded_filename = "uploaded_doc.docx"
            
            for part in msg.iter_parts():
                filename = part.get_filename()
                if filename:
                    uploaded_filename = filename
                    uploaded_file_data = part.get_payload(decode=True)
                    break
                    
            if not uploaded_file_data:
                self.send_response(400)
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                self.wfile.write(json.dumps({"error": "No file uploaded"}).encode("utf-8"))
                return
                
            temp_path = os.path.join(DIRECTORY, "temp_uploaded.docx")
            try:
                with open(temp_path, "wb") as f:
                    f.write(uploaded_file_data)
                
                validator = DocumentValidator(temp_path)
                result = validator.run_validation()
                result["filename"] = uploaded_filename
                
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps(result).encode("utf-8"))
            except Exception as e:
                self.send_response(500)
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                self.wfile.write(json.dumps({"error": str(e)}).encode("utf-8"))
            finally:
                if os.path.exists(temp_path):
                    try: os.remove(temp_path)
                    except: pass
            return

        # 5. API: Correct uploaded file and download directly
        elif path == "/api/correct":
            content_type = self.headers.get('Content-Type', '')
            if not content_type.startswith('multipart/form-data'):
                self.send_response(400)
                self.end_headers()
                return

            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)
            
            msg_data = b"Content-Type: " + content_type.encode('utf-8') + b"\r\n\r\n" + body
            msg = BytesParser(policy=default).parsebytes(msg_data)
            
            uploaded_file_data = None
            uploaded_filename = "uploaded_doc.docx"
            
            for part in msg.iter_parts():
                filename = part.get_filename()
                if filename:
                    uploaded_filename = filename
                    uploaded_file_data = part.get_payload(decode=True)
                    break
                    
            if not uploaded_file_data:
                self.send_response(400)
                self.end_headers()
                return
                
            temp_path = os.path.join(DIRECTORY, "temp_uploaded.docx")
            temp_corrected_path = os.path.join(DIRECTORY, "temp_corrected.docx")
            
            try:
                with open(temp_path, "wb") as f:
                    f.write(uploaded_file_data)
                
                # Run auto-corrections
                auto_correct_document(temp_path, temp_corrected_path)
                
                with open(temp_corrected_path, "rb") as f:
                    file_bytes = f.read()
                    
                corrected_filename = uploaded_filename.replace(".docx", "_Corrected.docx")
                
                self.send_response(200)
                self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                self.send_header("Content-Disposition", f"attachment; filename=\"{corrected_filename}\"")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.send_header("Content-Length", str(len(file_bytes)))
                self.end_headers()
                
                self.wfile.write(file_bytes)
            except Exception as e:
                self.send_response(500)
                self.send_header("Content-Type", "application/json")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({"error": str(e)}).encode("utf-8"))
            finally:
                if os.path.exists(temp_path):
                    try: os.remove(temp_path)
                    except: pass
                if os.path.exists(temp_corrected_path):
                    try: os.remove(temp_corrected_path)
                    except: pass
            return
            
        self.send_response(404)
        self.end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

def run_server():
    os.makedirs(DIRECTORY, exist_ok=True)
    socketserver.TCPServer.allow_reuse_address = True
    with socketserver.TCPServer(("", PORT), ValidatorHTTPRequestHandler) as httpd:
        print(f"Validator local HTTP server running at http://localhost:{PORT}")
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nShutting down server.")

if __name__ == "__main__":
    run_server()
