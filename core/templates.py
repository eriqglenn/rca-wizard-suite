"""Template manifest for the RCA Wizard Suite.

Before this module each wizard hard-coded the file name of every official City
template.  Replacing a template therefore meant either keeping the exact old
file name or editing four wizard files.  The manifest
(``Wizard_Files/templates/manifest.json``) is now the single place that says
which .docx file plays which role ("Director approval", "Routing form", ...)
for each wizard:

    {
      "wizards": {
        "mls": {
          "name": "MLS Wizard",
          "roles": [
            {"key": "director_template",
             "label": "Director approval",
             "file": "Director_Approval_Template_Clean.docx",
             "fill": "tokens",
             "expected_tokens": ["{{Application_Number}}", ...],
             "updated": "2026-09-02 10:15",
             "history": [...]}
          ]
        }
      }
    }

``key`` is the wizard's own file-dictionary key (unchanged from the original
code), ``fill`` records how the template is filled:

    tokens   the wizard replaces {{Token}} placeholders
    wording  the wizard matches the template's own placeholder wording
             ("XX Block of XX", "Planner Name", ...) - the MBL and renewal
             wizards work this way

Everything here is safe to call at startup: if the manifest is missing or
damaged it is rebuilt from the built-in defaults, and ``template_path`` always
falls back to the historical file name so a wizard can never fail to start
because of the manifest.
"""

from __future__ import annotations

import json
import re
import shutil
import zipfile
from datetime import datetime
from pathlib import Path

from .paths import TEMPLATES_ROOT, templates_dir

MANIFEST_PATH: Path = TEMPLATES_ROOT / "manifest.json"
PREVIOUS_DIRNAME = "_previous"
TOKEN_RE = re.compile(r"\{\{[^}]+\}\}")

WIZARD_NAMES = {
    "mls": "MLS Wizard",
    "mbl": "MBL Wizard",
    "mls_renewal": "MLS Renewal Wizard",
    "mbl_renewal": "MBL Renewal Wizard",
}

#: (role key used by the wizard's files dict, label, historical file name)
DEFAULT_ROLES: dict[str, list[tuple[str, str, str]]] = {
    "mls": [
        ("template", "Owner notification letter", "MLS_Owner_Notification_Letter.docx"),
        ("protest_template", "Protest letter to owners (PC hearing)", "MLS_Protest_All_Owners.docx"),
        ("hearing_template", "Notice of public hearing (support under 51%)", "MLS_Notice_Public_Hearing_Less_51.docx"),
        ("applicant_template", "Applicant sign letter", "MLS_Applicant_Complete_Signs.docx"),
        ("director_template", "Director approval", "MLS_Director_Approval.docx"),
        ("pc_template", "Planning Commission approval", "PC_Approval_NEW.docx"),
        ("routing_template", "MLS routing form", "MLS_Routing Form.docx"),
        ("staff_report_template", "OneMeeting staff report (Director path)", "MLS_OneMeeting_Staff_Report.docx"),
        ("ordinance_template", "Draft ordinance", "MLS_Ordinance_Template.docx"),
    ],
    "mbl": [
        ("pc_speech_template", "Planning Commission speech", "MBL_PC_Speech.docx"),
        ("pc_draft_template", "Planning Commission draft report", "MBL_PC_Draft_Report.docx"),
        ("template", "Owner notification letter", "MBL_ALL_OWNERS.docx"),
        ("hearing_template", "Notice of public hearing (support under 51%)", "MBL_Notice_Public_Hearing_Less_51.docx"),
        ("protest_template", "Protest letter to owners (PC hearing)", "MBL_Protest_All_Owners.docx"),
        ("reschedule_template", "Notice of PC rescheduling", "MBL_Notice_PC_Rescheduling.docx"),
        ("applicant_template", "Applicant sign letter", "MBL_APPLICANT_COMPLETE_SIGNS.docx"),
        ("director_template", "Director approval", "MBL_Directors_Approval.docx"),
        ("pc_template", "Planning Commission approval", "MBL_PC_Approval.docx"),
        ("routing_template", "MBL routing form", "MBL_Routing_Form.docx"),
        ("staff_report_template", "OneMeeting staff report (Director path)", "MBL_OneMeeting_Staff_Report.docx"),
        ("staff_report_less51_template", "OneMeeting staff report (support under 51%)", "MBL_OneMeeting_Staff_Report_Less51.docx"),
        ("staff_report_protest_template", "OneMeeting staff report (protest)", "MBL_OneMeeting_Staff_Report_Protest.docx"),
        ("ordinance_template", "MBL draft ordinance", "MBL_Ordinance_Template_Clean.docx"),
    ],
    "mls_renewal": [
        ("pc_speech_template", "Planning Commission speech", "MLS_REN_PC_Speech.docx"),
        ("pc_draft_template", "Planning Commission draft report", "MLS_REN_PC_Draft_Report.docx"),
        ("sfr_template", "SFR renewal letter to owners", "MLS_Renewal_SFR_Letters_to_owners.docx"),
        ("no_sfr_template", "No-SFR renewal letter to owners", "MLS_Renewal_NO_SFR_Letters_to_owners.docx"),
        ("protest_template", "Protest letter to owners (PC hearing)", "MLS_REN_Protest_All_Owners.docx"),
        ("applicant_template", "Applicant sign letter", "MLS_Renewal_Applicant_Signs.docx"),
        ("director_template", "Director approval", "MLS_REN_Director_Approval.docx"),
        ("pc_template", "Planning Commission approval", "MLS_REN_PC_Approval.docx"),
        ("routing_template", "Renewal routing form", "MLS_Renewal_Routing_Form.docx"),
        ("staff_report_director_template", "OneMeeting staff report (Director path)", "MLS_REN_OneMeeting_Director_Approval.docx"),
        ("staff_report_pc_template", "OneMeeting staff report (PC path)", "MLS_REN_OneMeeting_PC_Approval.docx"),
        ("ordinance_template", "Draft renewal ordinance", "MLS_RENEWAL_Ordinance_Template.docx"),
    ],
    "mbl_renewal": [
        ("pc_speech_template", "Planning Commission speech", "MBL_REN_PC_Speech.docx"),
        ("pc_draft_template", "Planning Commission draft report", "MBL_REN_PC_Draft_Report.docx"),
        ("owner_template", "Owner renewal letter", "MBL_REN_ALL_OWNERS.docx"),
        ("protest_template", "Protest letter to owners (PC hearing)", "MBL_REN_Protest_All_Owners.docx"),
        ("reschedule_template", "Notice of PC rescheduling", "MBL_REN_Notice_PC_Rescheduling.docx"),
        ("acceptance_template", "Application acceptance letter", "MBL_REN_Application_Acceptance.docx"),
        ("applicant_template", "Applicant sign letter", "MBL_REN_APPLICANT_COMPLETE_SIGNS.docx"),
        ("director_template", "Director approval", "MBL_REN_Directors_Approval.docx"),
        ("pc_template", "Planning Commission approval", "MBL_REN_PC_Approval.docx"),
        ("routing_template", "Renewal routing form", "MBL_REN_Routing_Form.docx"),
        ("staff_report_template", "OneMeeting staff report (Director path)", "MBL_REN_OneMeeting_Staff_Report.docx"),
        ("ordinance_template", "Draft renewal ordinance", "MBL_REN_Ordinance_Template.docx"),
    ],
}


def _now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M")


# ---------------------------------------------------------------- inspection
def inspect_template(path: str | Path) -> dict:
    """Structural facts about one .docx: tokens (body + headers + footers),
    split-run risk, section count, checkbox mechanism.  Never raises; an
    unreadable file returns ``{"error": ...}``."""
    path = Path(path)
    info: dict = {"tokens": [], "split_run_risk": False, "sections": 0,
                  "checkboxes": 0, "error": None}
    try:
        z = zipfile.ZipFile(path)
        body = z.read("word/document.xml").decode("utf-8", errors="replace")
        sources = [body]
        for part in z.namelist():
            if re.match(r"word/(header|footer)\d*\.xml$", part):
                sources.append(z.read(part).decode("utf-8", errors="replace"))
        tokens: set[str] = set()
        for xml in sources:
            tokens.update(TOKEN_RE.findall(xml))
            # tokens that only exist once the runs are joined would not fill
            joined = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml))
            if len(TOKEN_RE.findall(joined)) > len(TOKEN_RE.findall(xml)):
                info["split_run_risk"] = True
        info["tokens"] = sorted(tokens)
        info["sections"] = len(re.findall(r"<w:sectPr", body))
        info["checkboxes"] = body.count("w14:checkbox")
    except Exception as e:  # not a docx, corrupt, locked...
        info["error"] = str(e)
    return info


# ---------------------------------------------------------------- manifest IO
def _default_role(wizard_key: str, key: str, label: str, filename: str) -> dict:
    path = templates_dir(wizard_key) / filename
    tokens = inspect_template(path)["tokens"] if path.exists() else []
    return {
        "key": key,
        "label": label,
        "file": filename,
        "fill": "tokens" if tokens else "wording",
        "expected_tokens": tokens,
        "updated": None,
        "history": [],
    }


def bootstrap_manifest() -> dict:
    """Build the manifest from the built-in defaults and save it."""
    data = {"version": 1, "updated": _now(), "wizards": {}}
    for wkey, roles in DEFAULT_ROLES.items():
        data["wizards"][wkey] = {
            "name": WIZARD_NAMES[wkey],
            "roles": [_default_role(wkey, k, lbl, fn) for k, lbl, fn in roles],
        }
    save_manifest(data)
    return data


def load_manifest() -> dict:
    """Load the manifest, rebuilding or self-healing it when needed."""
    data: dict | None = None
    try:
        if MANIFEST_PATH.exists():
            loaded = json.loads(MANIFEST_PATH.read_text(encoding="utf-8"))
            if isinstance(loaded, dict) and isinstance(loaded.get("wizards"), dict):
                data = loaded
    except Exception:
        data = None
    if data is None:
        return bootstrap_manifest()

    # Self-heal: any wizard/role missing from an older manifest is added.
    changed = False
    for wkey, roles in DEFAULT_ROLES.items():
        wiz = data["wizards"].setdefault(wkey, {"name": WIZARD_NAMES[wkey], "roles": []})
        wiz.setdefault("name", WIZARD_NAMES[wkey])
        wiz.setdefault("roles", [])
        have = {r.get("key") for r in wiz["roles"]}
        for k, lbl, fn in roles:
            if k not in have:
                wiz["roles"].append(_default_role(wkey, k, lbl, fn))
                changed = True
    if changed:
        save_manifest(data)
    return data


def save_manifest(data: dict) -> None:
    data["updated"] = _now()
    MANIFEST_PATH.parent.mkdir(parents=True, exist_ok=True)
    tmp = MANIFEST_PATH.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    tmp.replace(MANIFEST_PATH)


# ---------------------------------------------------------------- queries
def wizard_keys() -> list[str]:
    return list(DEFAULT_ROLES)


def roles(wizard_key: str) -> list[dict]:
    return list(load_manifest()["wizards"].get(wizard_key, {}).get("roles", []))


def get_role(wizard_key: str, role_key: str) -> dict | None:
    for r in roles(wizard_key):
        if r.get("key") == role_key:
            return r
    return None


def default_filename(wizard_key: str, role_key: str) -> str | None:
    for k, _lbl, fn in DEFAULT_ROLES.get(wizard_key, []):
        if k == role_key:
            return fn
    return None


def template_path(wizard_key: str, role_key: str, fallback_filename: str | None = None) -> Path:
    """The .docx to use for one document role.

    Order: the manifest's file if it exists on disk, else the historical file
    name if that exists, else the manifest's file name (so the wizard shows a
    meaningful missing-file error).  Never raises.
    """
    folder = templates_dir(wizard_key)
    fallback = fallback_filename or default_filename(wizard_key, role_key)
    try:
        role = get_role(wizard_key, role_key)
    except Exception:
        role = None
    if role and role.get("file"):
        p = folder / role["file"]
        if p.exists():
            return p
        if fallback and (folder / fallback).exists():
            return folder / fallback
        return p
    return folder / (fallback or f"{role_key}.docx")


def role_status(wizard_key: str, role: dict) -> dict:
    """Health of one role's template for the Template Manager."""
    folder = templates_dir(wizard_key)
    path = folder / role.get("file", "")
    st = {
        "key": role.get("key"),
        "label": role.get("label"),
        "file": role.get("file"),
        "path": path,
        "fill": role.get("fill", "wording"),
        "exists": path.exists(),
        "updated": role.get("updated") or "original",
        "level": "ok",
        "message": "",
        "tokens": [],
        "missing_tokens": [],
        "unexpected_tokens": [],
        "split_run_risk": False,
        "can_restore": bool(role.get("history")),
    }
    if not path.exists():
        st["level"] = "error"
        st["message"] = "File not found in the templates folder"
        return st
    info = inspect_template(path)
    if info["error"]:
        st["level"] = "error"
        st["message"] = f"Cannot read as a Word file: {info['error']}"
        return st
    st["tokens"] = info["tokens"]
    st["split_run_risk"] = info["split_run_risk"]
    expected = set(role.get("expected_tokens") or [])
    found = set(info["tokens"])
    st["missing_tokens"] = sorted(expected - found)
    st["unexpected_tokens"] = sorted(found - expected)
    notes = []
    if st["unexpected_tokens"]:
        st["level"] = "warn"
        notes.append(f"{len(st['unexpected_tokens'])} new {{{{token}}}} the wizard does not know")
    if st["split_run_risk"]:
        st["level"] = "warn"
        notes.append("a token is split across Word runs and will not fill")
    if st["missing_tokens"]:
        notes.append(f"{len(st['missing_tokens'])} token(s) no longer used")
    st["message"] = "; ".join(notes) if notes else (
        f"{len(found)} tokens" if found else "filled by matching wording")
    return st


def wizard_status(wizard_key: str) -> list[dict]:
    return [role_status(wizard_key, r) for r in roles(wizard_key)]


# ---------------------------------------------------------------- changes
def _safe_name(name: str) -> str:
    name = re.sub(r'[\\/:*?"<>|]+', "_", name).strip()
    return name or "template.docx"


def _stamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def replace_template(wizard_key: str, role_key: str, new_file: str | Path) -> dict:
    """Install ``new_file`` as the template for one role.

    The current template is moved to ``templates/<wizard>/_previous`` (never
    deleted), the new file is copied into the templates folder under its own
    name, and the manifest records the change so it can be undone with
    :func:`restore_previous`.  Returns a summary dict.
    """
    new_file = Path(new_file)
    if not new_file.exists():
        raise FileNotFoundError(f"File not found: {new_file}")
    if new_file.suffix.lower() != ".docx":
        raise ValueError("Templates must be Word .docx files. Open the file in "
                         "Word and use Save As -> Word Document (.docx) first.")
    probe = inspect_template(new_file)
    if probe["error"]:
        raise ValueError(f"The file could not be read as a Word document: {probe['error']}")

    data = load_manifest()
    wiz = data["wizards"][wizard_key]
    role = next((r for r in wiz["roles"] if r.get("key") == role_key), None)
    if role is None:
        raise KeyError(f"Unknown template role {role_key!r} for {wizard_key!r}")

    folder = templates_dir(wizard_key)
    folder.mkdir(parents=True, exist_ok=True)
    prev_dir = folder / PREVIOUS_DIRNAME
    prev_dir.mkdir(exist_ok=True)

    target_name = _safe_name(new_file.name)
    # Another role may not share the same file name.
    for other in wiz["roles"]:
        if other is not role and other.get("file") == target_name:
            raise ValueError(
                f"'{target_name}' is already the {other.get('label')} template. "
                "Rename the new file before installing it.")

    old_name = role.get("file")
    old_path = folder / old_name if old_name else None
    archived_as = None
    if old_path and old_path.exists():
        archived_as = f"{_stamp()}__{old_name}"
        shutil.move(str(old_path), str(prev_dir / archived_as))
    # A stale file with the new name (not the current template) is archived too.
    target = folder / target_name
    if target.exists():
        shutil.move(str(target), str(prev_dir / f"{_stamp()}__stale__{target_name}"))
    shutil.copy2(str(new_file), str(target))

    role.setdefault("history", []).append({
        "date": _now(),
        "previous_file": old_name,
        "archived_as": archived_as,
        "new_file": target_name,
        "source": str(new_file),
    })
    role["file"] = target_name
    role["updated"] = _now()
    role["fill"] = "tokens" if probe["tokens"] else role.get("fill", "wording")
    save_manifest(data)
    return {
        "role": role,
        "installed": target,
        "archived": (prev_dir / archived_as) if archived_as else None,
        "tokens": probe["tokens"],
        "split_run_risk": probe["split_run_risk"],
    }


def restore_previous(wizard_key: str, role_key: str) -> dict:
    """Undo the most recent :func:`replace_template` for one role."""
    data = load_manifest()
    wiz = data["wizards"][wizard_key]
    role = next((r for r in wiz["roles"] if r.get("key") == role_key), None)
    if role is None or not role.get("history"):
        raise ValueError("There is no previous version of this template to restore.")
    entry = role["history"][-1]
    folder = templates_dir(wizard_key)
    prev_dir = folder / PREVIOUS_DIRNAME
    if not entry.get("archived_as") or not (prev_dir / entry["archived_as"]).exists():
        raise FileNotFoundError("The previous template file is no longer in the "
                                f"'{PREVIOUS_DIRNAME}' folder, so it cannot be restored.")

    current = folder / role["file"]
    if current.exists():
        prev_dir.mkdir(exist_ok=True)
        shutil.move(str(current), str(prev_dir / f"{_stamp()}__undone__{role['file']}"))
    restored = folder / entry["previous_file"]
    shutil.move(str(prev_dir / entry["archived_as"]), str(restored))

    role["history"].pop()
    role["file"] = entry["previous_file"]
    role["updated"] = role["history"][-1]["date"] if role["history"] else None
    role["fill"] = "tokens" if inspect_template(restored)["tokens"] else "wording"
    save_manifest(data)
    return {"role": role, "restored": restored}


def set_baseline(wizard_key: str, role_key: str) -> None:
    """Record the current template's tokens as the expected set (called after
    a replaced template passes the check)."""
    data = load_manifest()
    role = next((r for r in data["wizards"][wizard_key]["roles"]
                 if r.get("key") == role_key), None)
    if role is None:
        return
    path = templates_dir(wizard_key) / role["file"]
    info = inspect_template(path)
    if not info["error"]:
        role["expected_tokens"] = info["tokens"]
        role["fill"] = "tokens" if info["tokens"] else "wording"
        save_manifest(data)
