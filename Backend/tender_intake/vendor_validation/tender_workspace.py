
import json
import re
import secrets
from datetime import datetime, timezone
from pathlib import Path

from .submission_storage import STORAGE_DIR


WORKSPACE_DIR = STORAGE_DIR.parent / "tender_workspaces"

ID_PATTERN = re.compile(r"^[0-9a-f]{32}$")


def utc_now():
    """Return the current UTC timestamp."""
    return datetime.now(timezone.utc).isoformat()


def write_json_atomically(path, data):
    """Write JSON safely without leaving a partially written file."""

    path.parent.mkdir(parents=True, exist_ok=True)

    temporary_path = path.with_suffix(path.suffix + ".tmp")

    with temporary_path.open("w", encoding="utf-8") as output:
        json.dump(data, output, indent=2, ensure_ascii=False)

    temporary_path.replace(path)


def get_workspace_path(workspace_id):
    """Resolve a workspace ID to a safe workspace JSON path."""

    if (
        not isinstance(workspace_id, str)
        or not ID_PATTERN.fullmatch(workspace_id)
    ):
        raise ValueError("Invalid tender workspace ID.")

    workspace_path = (
        WORKSPACE_DIR / workspace_id / "workspace.json"
    ).resolve()

    workspace_root = WORKSPACE_DIR.resolve()

    if workspace_path.parent.parent != workspace_root:
        raise ValueError("Invalid tender workspace location.")

    return workspace_path


def load_workspace(workspace_id):
    """Load and validate an existing workspace."""

    workspace_path = get_workspace_path(workspace_id)

    if not workspace_path.is_file():
        return None

    with workspace_path.open("r", encoding="utf-8") as source:
        workspace = json.load(source)

    if workspace.get("workspace_id") != workspace_id:
        raise ValueError("Tender workspace ID mismatch.")

    if not isinstance(workspace.get("submissions"), list):
        raise ValueError("Invalid tender workspace records.")

    return workspace


def submission_exists(submission_id):
    """Check that a submission is stored inside the existing storage root."""

    if (
        not isinstance(submission_id, str)
        or not ID_PATTERN.fullmatch(submission_id)
    ):
        return False

    submission_path = (STORAGE_DIR / submission_id).resolve()

    if submission_path.parent != STORAGE_DIR.resolve():
        return False

    return (
        submission_path.is_dir()
        and (submission_path / "manifest.json").is_file()
        and (submission_path / "extraction.json").is_file()
    )


def add_submission_to_workspace(
    workspace_id,
    submission_id,
    category=None,
    status=None,
):
    """Associate an existing uploaded submission with a workspace."""

    workspace = load_workspace(workspace_id)

    if workspace is None:
        raise ValueError("Tender workspace does not exist.")

    if not submission_exists(submission_id):
        raise ValueError("The uploaded submission could not be verified.")

    existing_ids = {
        item.get("submission_id")
        for item in workspace["submissions"]
    }

    if submission_id not in existing_ids:
        workspace["submissions"].append({
            "submission_id": submission_id,
            "added_at": utc_now(),
            "category": category,
            "status": status,
        })

        workspace["updated_at"] = utc_now()

        write_json_atomically(
            get_workspace_path(workspace_id),
            workspace,
        )

    return workspace


def get_or_create_workspace(
    workspace_id=None,
    previous_submission_id=None,
):
    """
    Reuse the current workspace or create a new one.

    If an existing session has a previous submission but no workspace,
    preserve that submission by associating it with the new workspace.
    """

    if workspace_id is not None:
        workspace = load_workspace(workspace_id)

        if workspace is None:
            raise ValueError(
                "The current tender workspace could not be found. "
                "Please start a new workspace."
            )

        return workspace

    new_workspace_id = secrets.token_hex(16)

    workspace = {
        "workspace_id": new_workspace_id,
        "created_at": utc_now(),
        "updated_at": utc_now(),
        "submissions": [],
    }

    workspace_path = get_workspace_path(new_workspace_id)

    write_json_atomically(workspace_path, workspace)

    if (
        previous_submission_id
        and submission_exists(previous_submission_id)
    ):
        workspace = add_submission_to_workspace(
            new_workspace_id,
            previous_submission_id,
        )

    return workspace
