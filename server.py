#!/data/data/com.termux/files/usr/bin/python
import argparse
import json
import os
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer

USER_FILE_NAME = "user"


class MaiTavernHandler(SimpleHTTPRequestHandler):
    def _user_file_path(self):
        return os.path.join(self.directory or os.getcwd(), USER_FILE_NAME)

    def _send_json(self, status: int, payload: dict):
        body = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if self.path == "/api/user":
            user_path = self._user_file_path()
            if not os.path.exists(user_path):
                self._send_json(HTTPStatus.NOT_FOUND, {"error": "user file not found"})
                return

            try:
                with open(user_path, "r", encoding="utf-8") as f:
                    payload = json.load(f)
            except Exception as exc:
                self._send_json(HTTPStatus.INTERNAL_SERVER_ERROR, {"error": f"could not read user file: {exc}"})
                return

            self._send_json(HTTPStatus.OK, payload)
            return

        return super().do_GET()

    def do_PUT(self):
        if self.path != "/api/user":
            self._send_json(HTTPStatus.NOT_FOUND, {"error": "unknown endpoint"})
            return

        try:
            length = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            length = 0

        raw = self.rfile.read(max(length, 0))
        try:
            payload = json.loads(raw.decode("utf-8") if raw else "{}")
            if not isinstance(payload, dict):
                raise ValueError("payload must be a JSON object")
        except Exception as exc:
            self._send_json(HTTPStatus.BAD_REQUEST, {"error": f"invalid JSON payload: {exc}"})
            return

        user_path = self._user_file_path()
        temp_path = f"{user_path}.tmp"

        try:
            with open(temp_path, "w", encoding="utf-8") as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)
            os.replace(temp_path, user_path)
        except Exception as exc:
            self._send_json(HTTPStatus.INTERNAL_SERVER_ERROR, {"error": f"could not write user file: {exc}"})
            return

        self.send_response(HTTPStatus.NO_CONTENT)
        self.end_headers()


def main():
    parser = argparse.ArgumentParser(description="MaiTavern local storage server")
    parser.add_argument("--port", type=int, default=4444)
    parser.add_argument("--bind", default="127.0.0.1")
    args = parser.parse_args()

    directory = os.path.abspath(os.getcwd())

    def handler(*h_args, **h_kwargs):
        return MaiTavernHandler(*h_args, directory=directory, **h_kwargs)

    with ThreadingHTTPServer((args.bind, args.port), handler) as server:
        print(f"[MaiTavern] Serving {directory} at http://{args.bind}:{args.port}/index.html")
        print(f"[MaiTavern] User file endpoint: http://{args.bind}:{args.port}/api/user")
        server.serve_forever()


if __name__ == "__main__":
    main()
