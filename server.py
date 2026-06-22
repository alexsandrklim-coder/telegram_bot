import os
import json
from http.server import HTTPServer, SimpleHTTPRequestHandler

PORT = 8080
WEBDIR = os.path.join(os.path.dirname(__file__), "webapp")

class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=WEBDIR, **kwargs)

    def end_headers(self):
        self.send_header("Cache-Control", "no-cache")
        super().end_headers()

if __name__ == "__main__":
    print(f"Serving on http://localhost:{PORT}")
    HTTPServer(("0.0.0.0", PORT), Handler).serve_forever()
