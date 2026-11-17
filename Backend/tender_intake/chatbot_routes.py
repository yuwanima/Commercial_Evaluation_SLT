
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
from vendor_validation.document_intelligence import (
    analyze_workspace_documents,
)

from vendor_validation.tender_workspace import (
    add_submission_to_workspace,
    get_or_create_workspace,
    load_workspace,
    submission_exists,
)
from vendor_validation.document_intelligence import load_existing_intelligence
# Access the existing compliance analyzer in Backend/rag_chatbot.
import sys

BACKEND_DIR = Path(__file__).resolve().parents[1]

if str(BACKEND_DIR) not in sys.path:
    sys.path.insert(0, str(BACKEND_DIR))

from rag_chatbot.compliance_analyzer import (
    analyze_compliance,
    extract_status_declarations,
    get_chart_data,
    process_status_declarations,
)


chatbot_bp = Blueprint("chatbot", __name__)


@chatbot_bp.errorhandler(Exception)
def chatbot_unhandled_error(error):
    """Keep uncaught chatbot route errors in the JSON API format."""
    from werkzeug.exceptions import HTTPException

    if isinstance(error, HTTPException):
        return jsonify({"error": error.description}), error.code

    return jsonify({
        "error": "The chatbot request could not be completed. Please check the Flask server logs."
    }), 500

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


def load_workspace_submissions():
    """Load only validated submissions recorded in this session's workspace."""
    workspace_id = session.get("contractwise_workspace_id")
    if workspace_id:
        workspace = load_workspace(workspace_id)
        if workspace is None:
            return []
        records = workspace.get("submissions", [])
        submissions = []
        for record in records:
            submission_id = record.get("submission_id") if isinstance(record, dict) else None
            if not submission_exists(submission_id):
                continue
            submission_dir = (STORAGE_DIR / submission_id).resolve()
            extraction_path = submission_dir / "extraction.json"
            try:
                with extraction_path.open("r", encoding="utf-8") as source:
                    extraction = json.load(source)
            except (OSError, ValueError):
                continue
            if extraction.get("submission_id") == submission_id:
                submissions.append((submission_id, extraction))
        return submissions
    latest = session.get("contractwise_submission_id")
    extraction = load_submission_evidence(latest) if latest else None
    return [(latest, extraction)] if extraction else []


def is_comparison_request(question):
    comparison = re.search(r"\b(compare|comparison|versus|vs\.?|against|which vendor|best vendor|meets? the tender requirements)\b", question, re.I)
    scope = re.search(r"\b(documents?|vendors?|submissions?|uploaded|workspace|graph|chart|visuali[sz]e|visuali[sz]ation|delivery|technical|specifications?|compliance|price|warranty|security|validity)\b", question, re.I)
    return bool(comparison and scope)


def is_visualization_request(question):
    return bool(re.search(
        r"\b(graph|chart|visuali[sz](?:ation|e)|plot|graphical|diagram|distribution)\b",
        question,
        re.I,
    ))


def make_compliance_visualization(submission_id):
    """Return deterministic chart payload for one session-owned submission."""
    extraction = load_submission_evidence(submission_id)
    if extraction is None:
        return None, ({"error": "The uploaded submission could not be accessed. Please upload your documents again."}, 404)
    evidence = extraction.get("evidence", [])
    if not isinstance(evidence, list) or not evidence:
        return None, ({"error": "No extracted document evidence is available for visualization."}, 422)
    analysis = analyze_compliance(evidence)
    return {
        "success": True,
        "visualization_type": "compliance",
        "title": "Vendor Compliance Analysis",
        "analysis_complete": analysis["analysis_complete"],
        "summary": analysis["summary"],
        "declaration_summary": analysis["declaration_summary"],
        "chart_data": get_chart_data(analysis),
        "clauses": analysis["clauses"],
        "clause_count": analysis["clause_count"],
        "matched_clause_count": analysis["matched_clause_count"],
        "clause_matching_performed": analysis["clause_matching_performed"],
        "declarations": analysis["declarations"],
        "raw_declarations": analysis["raw_declarations"],
        "validated_declarations": analysis["validated_declarations"],
        "ambiguous_declaration_lines": analysis["ambiguous_declaration_lines"],
        "ambiguous_declaration_line_count": analysis["ambiguous_declaration_line_count"],
        "declaration_count": analysis["declaration_count"],
        "declaration_counts_available": analysis["declaration_counts_available"],
        "declaration_counts_complete": analysis["declaration_counts_complete"],
        "warnings": analysis["warnings"],
        "notice": analysis["notice"],
    }, None


def build_submission_comparison(submissions):
    """Build evidence-led, submission-isolated comparison records."""
    results = []
    warnings = []
    metrics = {"delivery_days": []}
    topics = {
        "Technical specifications": r"\b(specification|technical|model|capacity|performance|feature|configuration|standard)\b",
        "Quoted prices and costs": r"\b(price|quoted|quotation|total|amount|cost|currency|unit rate|tax|discount)\b",
        "Delivery period and schedule": r"\b(delivery|deliver|lead[ -]?time|days?|weeks?|months?|schedule)\b",
        "Warranty and support": r"\b(warranty|guarantee|support|maintenance|service period)\b",
        "Bid validity": r"\b(validity|valid for|bid valid|offer valid)\b",
        "Bid/performance security": r"\b(bid security|bid bond|performance security|bank guarantee|security amount)\b",
        "Compliance declarations": r"\b(complied|not complied|partially complied|compliance)\b",
    }

    for index, (submission_id, extraction) in enumerate(submissions, 1):
        extraction = extraction if isinstance(extraction, dict) else {}
        evidence = [item for item in extraction.get("evidence", []) if isinstance(item, dict)]
        try:
            analysis = analyze_compliance(evidence) if evidence else None
        except ValueError:
            analysis = None
        declaration_processing = process_status_declarations(
            extract_status_declarations(evidence)
        )
        statuses = declaration_processing["counts"]
        declaration_count = declaration_processing["total_declarations"]
        ambiguous_lines = declaration_processing["ambiguous_declaration_lines"]
        ambiguous_declaration_lines = declaration_processing["ambiguous_line_count"]
        extraction_errors = extraction.get("extraction_errors") or []
        has_ocr = any(item.get("requires_ocr") for item in evidence)
        counts_available = declaration_processing["counts_available"]
        counts_complete = counts_available and not extraction_errors and not has_ocr and not ambiguous_declaration_lines
        if not evidence:
            warnings.append(f"Submission {index}: no extracted evidence is available; compliance counts are unavailable.")
        elif not declaration_count:
            warnings.append(f"Submission {index}: no unambiguous compliance declarations were extracted; a zero-count graph was suppressed.")
        if ambiguous_declaration_lines:
            warnings.append(f"Submission {index}: {ambiguous_declaration_lines} declaration line(s) contain conflicting statuses and were excluded from counts.")
        if extraction_errors:
            warnings.append(f"Submission {index}: {len(extraction_errors)} file(s) failed extraction; available declaration counts may be incomplete.")
        names = sorted({item.get("file_name", "Unknown document") for item in evidence})
        intelligence = load_existing_intelligence(session.get("contractwise_workspace_id")) if session.get("contractwise_workspace_id") else None
        identified = set()
        ambiguous_identity = False
        if intelligence:
            for doc in intelligence.get("documents", []):
                if doc.get("submission_id") != submission_id or doc.get("analysis_status") != "analyzed":
                    continue
                identity = doc.get("vendor_identity_validation", {})
                if identity.get("accepted"):
                    name = doc.get("analysis", {}).get("primary_vendor_name")
                    if name:
                        identified.add(name)
                elif doc.get("analysis", {}).get("primary_vendor_name"):
                    ambiguous_identity = True
        vendor_identity = next(iter(identified)) if len(identified) == 1 and not ambiguous_identity else "Ambiguous identity; review required" if identified or ambiguous_identity else "Not verified"
        topic_evidence = {}
        for title, pattern in topics.items():
            matches = []
            for item in evidence:
                if re.search(pattern, item.get("text", ""), re.I):
                    matches.append({"file_name": item.get("file_name", "Unknown document"), "location": item.get("location", "Location unavailable"), "text": item.get("text", "")[:700], "requires_ocr": bool(item.get("requires_ocr"))})
            topic_evidence[title] = matches
        delivery_values = []
        for item in topic_evidence["Delivery period and schedule"]:
            for match in re.finditer(r"\b(\d+(?:\.\d+)?)\s*(days?|weeks?)\b", item["text"], re.I):
                value = float(match.group(1)) * (7 if match.group(2).lower().startswith("week") else 1)
                delivery_values.append({"value": value, "source": item})
        if has_ocr:
            warnings.append(f"Submission {index}: OCR-dependent evidence may be incomplete.")
        if analysis and not analysis.get("analysis_complete"):
            warnings.append(f"Submission {index}: clause-to-declaration matching is incomplete; graph counts use extracted declaration records independently.")
        if vendor_identity != "Not verified" and "Ambiguous" in vendor_identity:
            warnings.append(f"Submission {index}: vendor identity is ambiguous and requires review.")
        chart_label = f"Submission {index} — {vendor_identity} ({', '.join(names) or 'documents'})"
        clauses = analysis.get("clauses", []) if analysis else []
        results.append({
            "label": f"Submission {index}",
            "chart_label": chart_label,
            "submission_id": submission_id,
            "vendor_identity": vendor_identity,
            "documents": names,
            "topic_evidence": topic_evidence,
            "delivery_values": delivery_values,
            # Retain the legacy clause-matched counts field for API consumers.
            "counts": {status: sum(1 for clause in clauses if clause.get("vendor_status") == status) for status in ("Complied", "Partially Complied", "Not Complied", "Not Specified")},
            # These are the independent declaration counts used by the comparison graph.
            "declaration_counts": (statuses if counts_available else None),
            "declaration_count": declaration_count if counts_available else None,
            "raw_declarations": declaration_processing["raw_declarations"],
            "validated_declarations": declaration_processing["valid_declarations"],
            "ambiguous_declaration_lines": ambiguous_lines,
            "ambiguous_declaration_line_count": ambiguous_declaration_lines,
            "comparison_counts": ({
                "Complied": statuses["Complied"],
                "Non-Complied": statuses["Not Complied"] + statuses["Partially Complied"],
            } if counts_available else None),
            "counts_available": counts_available,
            "counts_complete": counts_complete,
        })
    if len(results) < 2:
        warnings.append("At least two valid submissions are needed for a cross-submission comparison.")
    notice = (
        "Topic matches are extracted evidence excerpts, not verified values or "
        "compliance decisions. Missing mentions are shown as Not found and are "
        "not proof of non-compliance."
    )
    if len(results) >= 2 and all(len(item["delivery_values"]) == 1 for item in results):
        metrics["delivery_days"] = [{"label": item["label"], "value": item["delivery_values"][0]["value"]} for item in results]
    else:
        warnings.append("Delivery values were not charted because one or more submissions had missing or multiple candidate values.")
    chart_available = len(results) >= 2 and all(item["counts_available"] for item in results)
    compliance_chart = {
        "available": chart_available,
        "complete": chart_available and all(item["counts_complete"] for item in results),
        "labels": [item["chart_label"] for item in results] if chart_available else [],
        "datasets": ([
            {"label": "Complied", "data": [item["comparison_counts"]["Complied"] for item in results], "backgroundColor": "#16A34A"},
            {"label": "Non-Complied", "data": [item["comparison_counts"]["Non-Complied"] for item in results], "backgroundColor": "#DC2626"},
        ] if chart_available else []),
        "x_axis_label": "Submission / Vendor",
        "y_axis_label": "Compliance declarations",
    }
    if len(results) >= 2 and not chart_available:
        warnings.append("Compliance comparison graph was suppressed because at least one submission has unavailable declaration counts.")
    return {"submissions": results, "topics": list(topics), "statuses": ["Complied", "Partially Complied", "Not Complied", "Not Specified"], "metrics": metrics, "compliance_chart": compliance_chart, "warnings": warnings, "notice": notice, "complete": len(results) == len(submissions) and len(results) >= 2 and not warnings}



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

    if is_comparison_request(question):
        submissions = load_workspace_submissions()
        if not submissions:
            return jsonify({"error": "No valid submissions are available in the active tender workspace."}), 400
        workspace_id = session.get("contractwise_workspace_id")
        if workspace_id:
            try:
                from vendor_validation.document_intelligence import analyze_workspace_documents
                analyze_workspace_documents(workspace_id)
            except Exception:
                # Keep evidence comparison available; identity remains explicitly unverified.
                pass
        try:
            comparison = build_submission_comparison(submissions)
            return jsonify({
                "answer": "Comparison uses evidence extracted separately from each submission. Vendor identity and compliance findings require review where flagged.",
                "comparison": comparison,
            }), 200
        except Exception:
            # Keep this API response JSON so fetch clients never receive Flask's HTML error page.
            return jsonify({
                "error": "The document comparison could not be completed. Please check the Flask server logs."
            }), 500

    # Server-side fallback prevents visualization questions from reaching Gemini,
    # even when an older/cached frontend script submits them to /chatbot/ask.
    if is_visualization_request(question):
        submission_id = session.get("contractwise_submission_id")
        if not submission_id:
            return jsonify({"error": "Please upload vendor documents before requesting a compliance graph."}), 400
        try:
            chart_response, failure = make_compliance_visualization(submission_id)
            if failure:
                return jsonify(failure[0]), failure[1]
            chart_response["answer"] = "Vendor compliance results from the uploaded submission."
            return jsonify(chart_response), 200
        except ValueError as error:
            return jsonify({"error": str(error)}), 422
        except Exception:
            return jsonify({"error": "The compliance visualization could not be generated. Please check the Flask server logs."}), 500

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

                # Reuse the current tender workspace or create one.
        workspace = get_or_create_workspace(
            workspace_id=session.get("contractwise_workspace_id"),
            previous_submission_id=session.get(
                "contractwise_submission_id"
            ),
        )

        # Associate the newly uploaded submission with this workspace.
        workspace = add_submission_to_workspace(
            workspace_id=workspace["workspace_id"],
            submission_id=result["submission_id"],
            category=result["category"],
            status=result["status"],
        )

        # Preserve the existing latest-submission session key.
        session["contractwise_submission_id"] = result["submission_id"]

        # Store the persistent workspace reference.
        session["contractwise_workspace_id"] = workspace["workspace_id"]

        # Keep all submission IDs available in the current session.
        session["contractwise_submission_ids"] = [
            item["submission_id"]
            for item in workspace["submissions"]
        ]

        session.modified = True

        return jsonify({
            "message": "Documents uploaded and processed.",
            "submission_id": result["submission_id"],
            "workspace_id": workspace["workspace_id"],
            "workspace_submission_count": len(
                workspace["submissions"]
            ),
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


@chatbot_bp.post("/chatbot/analyze-workspace")
def chatbot_analyze_workspace():
    """
    Analyze all unique documents associated with the active workspace.
    """

    workspace_id = session.get("contractwise_workspace_id")

    if not workspace_id:
        return jsonify({
            "success": False,
            "error": "No active tender workspace was found. Upload documents first.",
        }), 400

    try:
        result = analyze_workspace_documents(
            workspace_id=workspace_id
        )

        response_data = dict(result)
        response_data["success"] = True

        # Hide internal exception details from the frontend.
        response_data["analysis_errors"] = [
            {
                "document_id": item.get("document_id"),
                "file_name": item.get("file_name"),
                "error": "Document analysis failed. Please check the server logs.",
            }
            for item in result.get("analysis_errors", [])
            if isinstance(item, dict)
        ]

        # Remove internal error details from document records.
        response_data["documents"] = [
            {
                key: value
                for key, value in document.items()
                if key != "error"
            }
            for document in result.get("documents", [])
            if isinstance(document, dict)
        ]

        return jsonify(response_data), 200

    except ValueError as error:
        return jsonify({
            "success": False,
            "error": str(error),
        }), 422

    except Exception:
        return jsonify({
            "success": False,
            "error": "Workspace analysis failed. Please check the Flask server logs.",
        }), 500

@chatbot_bp.post("/chatbot/visualize")
def chatbot_visualize():
    """
    Generate structured compliance analysis and chart data
    from the currently uploaded vendor submission.
    """

    data = request.get_json(silent=True)

    if not isinstance(data, dict):
        return jsonify({
            "error": "Invalid visualization request."
        }), 400

    submission_id = data.get("submission_id")

    if submission_id is None:
        submission_id = session.get(
            "contractwise_submission_id"
        )

    if submission_id is None:
        return jsonify({
            "error": (
                "Please upload vendor documents before "
                "requesting a visualization."
            )
        }), 400

    extraction = load_submission_evidence(submission_id)

    if extraction is None:
        return jsonify({
            "error": (
                "The uploaded submission could not be accessed. "
                "Please upload your documents again."
            )
        }), 404

    evidence = extraction.get("evidence", [])

    if not isinstance(evidence, list) or not evidence:
        return jsonify({
            "error": (
                "No extracted document evidence is available "
                "for visualization."
            )
        }), 422

    try:
        # Analyse the complete submission evidence.
        # Do not use select_relevant_evidence here because
        # compliance totals must use the full submission.
        analysis = analyze_compliance(evidence)

        chart_data = get_chart_data(analysis)

        return jsonify({
            "success": True,
            "visualization_type": "compliance",
            "title": "Vendor Compliance Analysis",
            "analysis_complete": analysis["analysis_complete"],
            "summary": analysis["summary"],
            "declaration_summary": analysis["declaration_summary"],
            "chart_data": chart_data,
            "clauses": analysis["clauses"],
            "clause_count": analysis["clause_count"],
            "matched_clause_count": analysis["matched_clause_count"],
            "clause_matching_performed": analysis["clause_matching_performed"],
            "declarations": analysis["declarations"],
            "raw_declarations": analysis["raw_declarations"],
            "validated_declarations": analysis["validated_declarations"],
            "ambiguous_declaration_lines": analysis["ambiguous_declaration_lines"],
            "ambiguous_declaration_line_count": analysis["ambiguous_declaration_line_count"],
            "declaration_count": analysis["declaration_count"],
            "declaration_counts_available": analysis["declaration_counts_available"],
            "declaration_counts_complete": analysis["declaration_counts_complete"],
            "warnings": analysis["warnings"],
            "notice": analysis["notice"],
        }), 200

    except ValueError as error:
        return jsonify({
            "error": str(error)
        }), 422

    except Exception:
        return jsonify({
            "error": (
                "An unexpected error occurred while "
                "generating the compliance visualization."
            )
        }), 500
