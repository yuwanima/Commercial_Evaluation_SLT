
import json
import re
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

from flask import (
    Blueprint,
    jsonify,
    render_template,
    request,
    session,
)

from vendor_validation.submission_processor import process_submission
from vendor_validation.submission_storage import STORAGE_DIR


chatbot_bp = Blueprint("chatbot", __name__)

RAG_API_URL = "http://127.0.0.1:8765/api/ask"

MAX_VENDOR_EVIDENCE_CHARS = 30000
MAX_VENDOR_EVIDENCE_SECTIONS = 20

ALLOWED_CATEGORIES = [
    "1. Goods & Services (Local or Foreign)",
    "2. Goods (Local & Foreign)",
    "3. Goods & Services (Local)",
    "4. Goods (Local)",
    "5. Services (Local or Foreign)",
]


def load_submission_evidence(submission_id):
    """Load evidence only for the submission owned by this session."""

    if not isinstance(submission_id, str):
        return None

    if not re.fullmatch(r"[0-9a-f]{32}", submission_id):
        return None

    if session.get("contractwise_submission_id") != submission_id:
        return None

    submission_dir = (STORAGE_DIR / submission_id).resolve()
    storage_root = STORAGE_DIR.resolve()

    if submission_dir.parent != storage_root:
        return None

    extraction_path = submission_dir / "extraction.json"

    if not extraction_path.is_file():
        return None

    with extraction_path.open("r", encoding="utf-8") as file:
        extraction = json.load(file)

    if extraction.get("submission_id") != submission_id:
        return None

    return extraction



def select_relevant_evidence(question, evidence):
    """Select relevant vendor document sections within a size limit."""

    if not isinstance(evidence, list):
        return []

    question_words = set(
        re.findall(r"[a-zA-Z0-9]+", question.lower())
    )

    stop_words = {
        "what", "which", "when", "where", "does", "this",
        "that", "with", "from", "have", "has", "are", "was",
        "were", "the", "and", "for", "how", "can", "could",
        "would", "should", "about", "please", "tell", "me",
        "is", "it", "of", "to", "in", "on", "a", "an",
    }

    question_words -= stop_words

    candidates = []

    for index, item in enumerate(evidence):
        if not isinstance(item, dict):
            continue

        text = item.get("text")

        if not isinstance(text, str) or not text.strip():
            continue

        if item.get("requires_ocr"):
            continue

        file_name = item.get("file_name")
        location = item.get("location")

        if not isinstance(file_name, str) or not file_name.strip():
            file_name = "Unknown document"

        if not isinstance(location, str) or not location.strip():
            location = "Location unavailable"

        searchable_text = (
            file_name + " " + location + " " + text
        ).lower()

        document_words = set(
            re.findall(r"[a-zA-Z0-9]+", searchable_text)
        )

        score = len(question_words & document_words)

        candidates.append({
            "index": index,
            "score": score,
            "file_name": file_name,
            "location": location,
            "text": text.strip(),
        })

    # Prefer sections matching the question.
    candidates.sort(
        key=lambda item: (-item["score"], item["index"])
    )

    selected = []
    total_chars = 0

    for item in candidates:
        if question_words and item["score"] == 0:
            continue

        remaining = MAX_VENDOR_EVIDENCE_CHARS - total_chars

        if remaining <= 0:
            break

        text = item["text"][:remaining]

        if not text.strip():
            continue

        selected.append({
            "file_name": item["file_name"],
            "location": item["location"],
            "text": text,
        })

        total_chars += len(text)

        if len(selected) >= MAX_VENDOR_EVIDENCE_SECTIONS:
            break

    # If nothing matched, provide a small sample rather than
    # incorrectly claiming that the uploaded document is empty.
    if not selected:
        for item in candidates[:5]:
            remaining = MAX_VENDOR_EVIDENCE_CHARS - total_chars

            if remaining <= 0:
                break

            text = item["text"][:min(2000, remaining)]

            if not text.strip():
                continue

            selected.append({
                "file_name": item["file_name"],
                "location": item["location"],
                "text": text,
            })

            total_chars += len(text)

    return selected



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
        return jsonify({
            "error": "Invalid request. Please send JSON."
        }), 400

    question = data.get("question")
    category = data.get("category") or None

    if not isinstance(question, str) or not question.strip():
        return jsonify({
            "error": "Please enter a procurement question."
        }), 400

    question = question.strip()

    if len(question) > 5000:
        return jsonify({"error": "Question is too long."}), 400

    if category is not None and category not in ALLOWED_CATEGORIES:
        return jsonify({
            "error": "Invalid procurement category."
        }), 400

    # Use the supplied ID or the latest submission in this session.
    submission_id = data.get("submission_id")

    if submission_id is None:
        submission_id = session.get("contractwise_submission_id")

    vendor_evidence = []

    if submission_id is not None:
        extraction = load_submission_evidence(submission_id)

        if extraction is None:
            return jsonify({
                "error": (
                    "The uploaded submission could not be accessed. "
                    "Please upload your documents again."
                )
            }), 404

        vendor_evidence = select_relevant_evidence(
            question,
            extraction.get("evidence", []),
        )

    payload = json.dumps({
        "question": question,
        "category": category,
        "top_k": 5,
        "vendor_evidence": vendor_evidence,
    }, ensure_ascii=False).encode("utf-8")

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
            error_data = {
                "error": "The RAG service returned an error."
            }

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
            "error": (
                "An unexpected error occurred while "
                "processing your question."
            )
        }), 500


@chatbot_bp.post("/chatbot/upload")
def chatbot_upload():
    """Accept and process multiple vendor submission documents."""

    category = request.form.get("category") or None

    if category is not None and category not in ALLOWED_CATEGORIES:
        return jsonify({
            "error": "Invalid procurement category."
        }), 400

    uploaded_files = request.files.getlist("files")

    if not uploaded_files:
        return jsonify({
            "error": "Please select at least one document."
        }), 400

    try:
        result = process_submission(
            uploaded_files,
            category=category,
        )

        # Associate this submission with the current browser session.
        session["contractwise_submission_id"] = result["submission_id"]

        return jsonify({
            "message": "Documents uploaded and processed.",
            "submission_id": result["submission_id"],
            "category": result["category"],
            "status": result["status"],
            "summary": result["summary"],
            "files": [
                {
                    "file_name": item["file_name"],
                    "status": item["status"],
                    "section_count": item["section_count"],
                    **(
                        {"error": item["error"]}
                        if item.get("error")
                        else {}
                    ),
                }
                for item in result["files"]
            ],
            "pages_requiring_ocr": result["pages_requiring_ocr"],
        }), 201

    except ValueError as error:
        return jsonify({"error": str(error)}), 400

    except Exception:
        return jsonify({
            "error": (
                "An unexpected error occurred while processing "
                "the uploaded documents."
            )
        }), 500
