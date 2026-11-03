import json
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

from flask import Blueprint, jsonify, render_template, request


chatbot_bp = Blueprint("chatbot", __name__)

RAG_API_URL = "http://127.0.0.1:8765/api/ask"

ALLOWED_CATEGORIES = [
    "1. Goods & Services (Local or Foreign)",
    "2. Goods (Local & Foreign)",
    "3. Goods & Services (Local)",
    "4. Goods (Local)",
    "5. Services (Local or Foreign)",
]


@chatbot_bp.get("/chatbot")
def chatbot_page():
    return render_template(
        "chatbot.html",
        categories=ALLOWED_CATEGORIES,
    )


@chatbot_bp.post("/chatbot/ask")
def chatbot_ask():

    data = request.get_json(silent=True)

    if not isinstance(data, dict):
        return jsonify({"error": "Invalid request. Please send JSON."}), 400

    question = data.get("question")
    category = data.get("category") or None

    if not isinstance(question, str) or not question.strip():
        return jsonify({"error": "Please enter a procurement question."}), 400

    if len(question) > 5000:
        return jsonify({"error": "Question is too long."}), 400

    if category is not None and category not in ALLOWED_CATEGORIES:
        return jsonify({"error": "Invalid procurement category."}), 400

    payload = json.dumps({
        "question": question.strip(),
        "category": category,
        "top_k": 5,
    }).encode("utf-8")

    api_request = Request(
        RAG_API_URL,
        data=payload,
        headers={"Content-Type": "application/json"},
        method="POST",
    )

    try:
        with urlopen(api_request, timeout=180) as response:
            result = json.loads(response.read().decode("utf-8"))

        return jsonify(result), 200

    except HTTPError as error:
        try:
            error_data = json.loads(error.read().decode("utf-8"))
        except (ValueError, UnicodeDecodeError):
            error_data = {"error": "The RAG service returned an error."}

        return jsonify(error_data), error.code

    except (URLError, TimeoutError):
        return jsonify({
            "error": (
                "Cannot connect to the procurement RAG service. "
                "Please check that the RAG API is running."
            )
        }), 503

    except (ValueError, UnicodeDecodeError):
        return jsonify({
            "error": "The RAG service returned an invalid response."
        }), 502

    except Exception:
        return jsonify({
            "error": "An unexpected error occurred while processing your question."
        }), 500
