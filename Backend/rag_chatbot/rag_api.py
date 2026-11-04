import json
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

from answer_generator import answer_question
from retrieval_engine import ALLOWED_CATEGORIES


HOST = "127.0.0.1"
PORT = 8765
MAX_REQUEST_SIZE = 16 * 1024


class RAGRequestHandler(BaseHTTPRequestHandler):

    def send_json(self, status_code, data):

        response = json.dumps(data, ensure_ascii=False).encode("utf-8")

        self.send_response(status_code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(response)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()

        self.wfile.write(response)

    def do_GET(self):

        if self.path == "/health":
            self.send_json(200, {
                "status": "ok",
                "service": "SLT Procurement RAG API"
            })
        else:
            self.send_json(404, {
                "error": "Endpoint not found."
            })

    def do_POST(self):

        if self.path != "/api/ask":
            self.send_json(404, {
                "error": "Endpoint not found."
            })
            return

        try:
            content_length = int(
                self.headers.get("Content-Length", "0")
            )

            if content_length <= 0 or content_length > MAX_REQUEST_SIZE:
                self.send_json(400, {
                    "error": "Invalid request size."
                })
                return

            raw_body = self.rfile.read(content_length)
            data = json.loads(raw_body.decode("utf-8"))

            if not isinstance(data, dict):
                self.send_json(400, {
                    "error": "Request must be a JSON object."
                })
                return

            question = data.get("question")
            category = data.get("category")
            top_k = data.get("top_k", 5)

            if not isinstance(question, str) or not question.strip():
                self.send_json(400, {
                    "error": "Please enter a procurement question."
                })
                return

            if len(question) > 5000:
                self.send_json(400, {
                    "error": "Question is too long."
                })
                return

            if category == "":
                category = None

            if category is not None and category not in ALLOWED_CATEGORIES:
                self.send_json(400, {
                    "error": "Invalid procurement category."
                })
                return

            if (
                not isinstance(top_k, int)
                or isinstance(top_k, bool)
                or not 1 <= top_k <= 20
            ):
                self.send_json(400, {
                    "error": "top_k must be between 1 and 20."
                })
                return

            print("\nReceived procurement question:")
            print(question)

            result = answer_question(
                question=question,
                category=category,
                top_k=top_k,
            )

            self.send_json(200, result)

            print("Answer generated successfully.")

        except json.JSONDecodeError:
            self.send_json(400, {
                "error": "Invalid JSON request."
            })

        except Exception as error:
            print("\nRAG API ERROR:")
            print(repr(error))

            self.send_json(500, {
                "error": "The procurement chatbot could not process your question. Please try again."
            })


def main():

    server = ThreadingHTTPServer(
        (HOST, PORT),
        RAGRequestHandler
    )

    print("=" * 60)
    print("SLT PROCUREMENT RAG API")
    print("=" * 60)
    print(f"API address: http://{HOST}:{PORT}")
    print(f"Health check: http://{HOST}:{PORT}/health")
    print(f"Question endpoint: http://{HOST}:{PORT}/api/ask")
    print("\nPress Ctrl+C to stop the server.")

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopping RAG API...")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
