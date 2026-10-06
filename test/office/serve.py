"""Static server for the add-in plus POST /report (saves the self-test log).

    python3 test/office/serve.py [port] [report-file]
"""
import http.server, sys

PORT = int(sys.argv[1]) if len(sys.argv) > 1 else 8765
REPORT = sys.argv[2] if len(sys.argv) > 2 else "selftest-report.json"


class Handler(http.server.SimpleHTTPRequestHandler):
    def do_POST(self):
        body = self.rfile.read(int(self.headers.get("Content-Length", 0)))
        with open(REPORT, "wb") as f:
            f.write(body)
        self.send_response(204)
        self.end_headers()

    def end_headers(self):
        self.send_header("Cache-Control", "no-store")
        super().end_headers()


http.server.ThreadingHTTPServer(("127.0.0.1", PORT), Handler).serve_forever()
