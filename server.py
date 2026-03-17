"""
Teams Chat Viewer - Local PST Conversion Server
Run this once and leave it running. The web app will use it automatically.
"""

import os
import sys
import subprocess
import tempfile
import shutil
import zipfile
import json
from pathlib import Path
from http.server import HTTPServer, BaseHTTPRequestHandler
import urllib.parse
import email
import email.policy

PORT = 5000
ALLOWED_ORIGIN = "*"  # local only, fine to be open


def check_readpst():
    """Check if readpst is available."""
    try:
        subprocess.run(["readpst", "--version"], capture_output=True, timeout=5)
        return True
    except (FileNotFoundError, subprocess.TimeoutExpired):
        return False


def check_pypff():
    """Check if pypff (libpff-python) is available."""
    try:
        import pypff
        return True
    except ImportError:
        return False


def convert_with_readpst(pst_path, output_dir):
    """Convert PST to EML files using readpst command-line tool."""
    result = subprocess.run(
        ["readpst", "-e", "-o", output_dir, pst_path],
        capture_output=True, text=True, timeout=120
    )
    emls = list(Path(output_dir).rglob("*.eml"))
    return emls


def convert_with_pypff(pst_path, output_dir):
    """Convert PST to EML files using pypff (libpff-python)."""
    import pypff

    emls = []
    pff_file = pypff.file()
    pff_file.open(pst_path)

    def process_folder(folder, depth=0):
        if depth > 20:
            return
        try:
            for i in range(folder.number_of_sub_messages):
                try:
                    msg = folder.get_sub_message(i)
                    subject = msg.subject or "(no subject)"
                    safe_subject = "".join(c for c in subject if c.isalnum() or c in " -_")[:50]
                    eml_path = os.path.join(output_dir, f"msg_{len(emls)+1}_{safe_subject}.eml")

                    # Build EML content
                    lines = []
                    lines.append(f"From: {msg.sender_name or ''} <{msg.sender_email_address or ''}>")
                    lines.append(f"Subject: {subject}")
                    if msg.delivery_time:
                        lines.append(f"Date: {msg.delivery_time}")
                    lines.append("Content-Type: text/plain; charset=utf-8")
                    lines.append("")
                    body = msg.plain_text_body or msg.html_body or b""
                    if isinstance(body, bytes):
                        body = body.decode("utf-8", errors="replace")
                    lines.append(body)

                    with open(eml_path, "w", encoding="utf-8") as f:
                        f.write("\r\n".join(lines))
                    emls.append(Path(eml_path))
                except Exception:
                    pass

            for i in range(folder.number_of_sub_folders):
                try:
                    sub = folder.get_sub_folder(i)
                    process_folder(sub, depth + 1)
                except Exception:
                    pass
        except Exception:
            pass

    root = pff_file.get_root_folder()
    process_folder(root)
    pff_file.close()
    return emls


def convert_pst(pst_path, output_dir):
    """Try available conversion methods in order of preference."""
    if check_readpst():
        print("  Using: readpst")
        return convert_with_readpst(pst_path, output_dir)
    elif check_pypff():
        print("  Using: pypff")
        return convert_with_pypff(pst_path, output_dir)
    else:
        raise RuntimeError(
            "No PST conversion tool found. "
            "Please run: pip install libpff-python\n"
            "Or install readpst: brew install libpst (Mac) / apt install pst-utils (Linux)"
        )


class Handler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        print(f"  [{self.address_string()}] {format % args}")

    def send_cors(self):
        self.send_header("Access-Control-Allow-Origin", ALLOWED_ORIGIN)
        self.send_header("Access-Control-Allow-Methods", "POST, GET, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_cors()
        self.end_headers()

    def do_GET(self):
        if self.path == "/status":
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_cors()
            self.end_headers()
            status = {
                "running": True,
                "readpst": check_readpst(),
                "pypff": check_pypff(),
                "version": "1.0.0"
            }
            self.wfile.write(json.dumps(status).encode())
        else:
            self.send_response(404)
            self.end_headers()

    def do_POST(self):
        if self.path == "/convert":
            self.handle_convert()
        else:
            self.send_response(404)
            self.end_headers()

    def handle_convert(self):
        try:
            content_type = self.headers.get("Content-Type", "")
            content_length = int(self.headers.get("Content-Length", 0))

            if "multipart/form-data" not in content_type:
                self.send_error_json(400, "Expected multipart/form-data")
                return

            # Parse boundary
            boundary = None
            for part in content_type.split(";"):
                part = part.strip()
                if part.startswith("boundary="):
                    boundary = part[9:].strip('"')
                    break

            if not boundary:
                self.send_error_json(400, "No boundary found in Content-Type")
                return

            body = self.rfile.read(content_length)

            # Extract file from multipart
            bnd = ("--" + boundary).encode()
            parts = body.split(bnd)
            pst_data = None
            filename = "upload.pst"

            for part in parts:
                if b"Content-Disposition" in part and b'name="file"' in part:
                    header_end = part.find(b"\r\n\r\n")
                    if header_end == -1:
                        header_end = part.find(b"\n\n")
                        file_data = part[header_end + 2:].rstrip(b"\r\n--")
                    else:
                        file_data = part[header_end + 4:].rstrip(b"\r\n--")

                    header_section = part[:header_end].decode("utf-8", errors="replace")
                    for line in header_section.split("\r\n"):
                        if 'filename="' in line:
                            fn_start = line.find('filename="') + 10
                            fn_end = line.find('"', fn_start)
                            if fn_end > fn_start:
                                filename = line[fn_start:fn_end]
                    pst_data = file_data
                    break

            if not pst_data:
                self.send_error_json(400, "No file found in request")
                return

            print(f"\n>> Converting: {filename} ({len(pst_data):,} bytes)")

            # Save to temp dir and convert
            with tempfile.TemporaryDirectory() as tmpdir:
                pst_path = os.path.join(tmpdir, filename)
                out_dir = os.path.join(tmpdir, "output")
                os.makedirs(out_dir)

                with open(pst_path, "wb") as f:
                    f.write(pst_data)

                try:
                    eml_files = convert_pst(pst_path, out_dir)
                except RuntimeError as e:
                    self.send_error_json(500, str(e))
                    return

                if not eml_files:
                    self.send_error_json(422, "No messages could be extracted. The PST may be encrypted or empty.")
                    return

                print(f"  Extracted: {len(eml_files)} messages")

                # Zip up all EMLs
                zip_path = os.path.join(tmpdir, "emails.zip")
                with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
                    for eml in eml_files:
                        zf.write(eml, eml.name)

                with open(zip_path, "rb") as f:
                    zip_data = f.read()

            # Send response
            self.send_response(200)
            self.send_header("Content-Type", "application/zip")
            self.send_header("Content-Disposition", f'attachment; filename="converted_emails.zip"')
            self.send_header("Content-Length", str(len(zip_data)))
            self.send_header("X-Message-Count", str(len(eml_files)))
            self.send_cors()
            self.end_headers()
            self.wfile.write(zip_data)

        except Exception as e:
            print(f"  ERROR: {e}")
            import traceback
            traceback.print_exc()
            self.send_error_json(500, f"Server error: {str(e)}")

    def send_error_json(self, code, message):
        body = json.dumps({"error": message}).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.send_cors()
        self.end_headers()
        self.wfile.write(body)


def main():
    print("=" * 55)
    print("  Teams Chat Viewer — PST Conversion Server")
    print("=" * 55)
    print(f"\n  Server running at: http://localhost:{PORT}")
    print(f"  Status check:      http://localhost:{PORT}/status")
    print("\n  Checking tools...")
    print(f"  readpst:   {'✓ found' if check_readpst() else '✗ not found'}")
    print(f"  pypff:     {'✓ found' if check_pypff() else '✗ not found'}")

    if not check_readpst() and not check_pypff():
        print("\n  ⚠  No conversion tool found!")
        print("  Run one of:")
        print("    pip install libpff-python")
        print("    brew install libpst        (Mac)")
        print("    sudo apt install pst-utils (Linux/WSL)")
        print("\n  Server will still start but conversions will fail")
        print("  until a tool is installed.")

    print("\n  Keep this window open while using the app.")
    print("  Press Ctrl+C to stop.\n")
    print("-" * 55)

    server = HTTPServer(("localhost", PORT), Handler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n\n  Server stopped.")


if __name__ == "__main__":
    main()
