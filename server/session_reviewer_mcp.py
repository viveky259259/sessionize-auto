#!/usr/bin/env python3
"""Local stdio MCP server for reviewing session-submission Excel workbooks."""

from __future__ import annotations

import json
import math
import re
import sys
import tempfile
import zipfile
from collections import Counter, defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
XML_NS = "http://www.w3.org/XML/1998/namespace"
NS = {"x": MAIN_NS, "r": REL_NS, "pr": PKG_REL_NS}
STOPWORDS = {
    "a", "an", "and", "are", "as", "at", "be", "by", "for", "from", "how", "in", "is",
    "it", "of", "on", "or", "the", "to", "with", "your", "you", "this", "that", "will",
    "we", "our", "into", "about", "using", "through", "session", "talk", "learn",
}
VAGUE_PHRASES = ("various topics", "many things", "general overview", "some ideas", "best practices")
TOPIC_CATEGORIES = (
    "AI & Machine Learning", "Architecture & State Management", "Backend, Data & Offline",
    "Business, Career & Community", "Developer Tools & Productivity", "Flutter Fundamentals",
    "Gaming & Graphics", "Hardware, IoT & Wearables", "Performance & Scalability",
    "Platform & Native Integration", "Release Engineering & DevOps", "Security & Privacy",
    "Testing & Quality", "UI, UX & Design", "Web, Desktop & Multi-platform", "Other / Needs review",
)
FIELD_ALIASES = {
    "id": ("id", "session id", "proposal id", "submission id", "session code"),
    "title": ("title", "session title", "proposal title", "name"),
    "abstract": ("abstract", "description", "session description", "summary", "proposal", "content"),
    "speaker": ("speaker", "speaker name", "presenter", "presenter name", "submitted by", "author"),
    "tags": ("tags", "tag", "topics", "topic", "track", "category", "categories"),
}


class ToolError(ValueError):
    pass


def text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def normalize(value: Any) -> str:
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9]+", " ", text(value).lower())).strip()


def tokens(value: Any) -> set[str]:
    def stem(word: str) -> str:
        return word[:-1] if len(word) > 3 and word.endswith("s") and not word.endswith("ss") else word

    return {stem(word) for word in normalize(value).split() if len(word) > 2 and word not in STOPWORDS}


def topic_category(title: Any, description: Any = "", format_name: Any = "") -> str:
    value = normalize(f"{title} {description} {format_name}")
    checks = (
        (r"\b(ai|llm|genkit|agent|model|machine learning|artificial intelligence)\b", "AI & Machine Learning"),
        (r"\b(security|secure|privacy|pii|token|auth|authentication|authorization|compliance)\b", "Security & Privacy"),
        (r"\b(test|testing|quality|verify|verification|debug|debugging)\b", "Testing & Quality"),
        (r"\b(release|deploy|deployment|ci|cd|shorebird|ota update|devops)\b", "Release Engineering & DevOps"),
        (r"\b(performance|jank|smooth|scale|scaling|scalable|million users|billion users)\b", "Performance & Scalability"),
        (r"\b(bluetooth|hardware|wearable|watch|healthkit|airplay|iot|sensor)\b", "Hardware, IoT & Wearables"),
        (r"\b(native|android|ios|callkit|ffi|interop|background execution|foreground service|workmanager|live activities)\b", "Platform & Native Integration"),
        (r"\b(game|flame|canvas|rendering|graphics|animation)\b", "Gaming & Graphics"),
        (r"\b(web|wasm|wasmgc|desktop|multi platform|cross platform)\b", "Web, Desktop & Multi-platform"),
        (r"\b(cache|sync|database|storage|backend|supabase|offline|data stream)\b", "Backend, Data & Offline"),
        (r"\b(tooling|analyzer|devtools|package|cli|code generation|productivity)\b", "Developer Tools & Productivity"),
        (r"\b(architecture|monorepo|modular|module|ddd|clean architecture|state management|dependency)\b", "Architecture & State Management"),
        (r"\b(widget|ui|ux|design|layout|theme|foldable)\b", "UI, UX & Design"),
        (r"\b(burnout|happiness|career|community|leadership|keynote|journey|speaker)\b", "Business, Career & Community"),
        (r"\b(dart|flutter)\b", "Flutter Fundamentals"),
    )
    return next((category for pattern, category in checks if re.search(pattern, value)), "Other / Needs review")


def col_number(reference: str) -> int:
    letters = re.match(r"[A-Z]+", reference or "")
    if not letters:
        return 0
    result = 0
    for letter in letters.group(0):
        result = result * 26 + ord(letter) - 64
    return result - 1


def col_letter(index: int) -> str:
    result = ""
    index += 1
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def xml_escape(value: Any) -> str:
    value = str(value)
    return (value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                 .replace('"', "&quot;").replace("'", "&apos;"))


def read_xlsx(source_path: Path, requested_sheet: str | None = None) -> tuple[str, list[str], list[dict[str, str]]]:
    if source_path.suffix.lower() != ".xlsx":
        raise ToolError("Only .xlsx workbooks are supported.")
    if not source_path.is_file():
        raise ToolError(f"Workbook not found: {source_path}")
    try:
        archive = zipfile.ZipFile(source_path)
    except zipfile.BadZipFile as error:
        raise ToolError("The selected file is not a readable .xlsx workbook.") from error
    with archive:
        try:
            workbook = ET.fromstring(archive.read("xl/workbook.xml"))
            relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        except KeyError as error:
            raise ToolError("The workbook is missing required Excel XML parts.") from error
        rel_targets = {
            rel.attrib["Id"]: rel.attrib["Target"]
            for rel in relationships.findall("pr:Relationship", NS)
        }
        sheets: list[tuple[str, str]] = []
        for sheet in workbook.findall("x:sheets/x:sheet", NS):
            rel_id = sheet.attrib.get(f"{{{REL_NS}}}id")
            target = rel_targets.get(rel_id or "")
            if target:
                sheets.append((sheet.attrib.get("name", "Sheet"), "xl/" + target.lstrip("/")))
        if not sheets:
            raise ToolError("No readable worksheets were found.")
        sheet_name, sheet_file = next(((name, path) for name, path in sheets if name == requested_sheet), sheets[0])
        if requested_sheet and sheet_name != requested_sheet:
            raise ToolError("Selected sheet was not found. Available sheets: " + ", ".join(name for name, _ in sheets))
        shared: list[str] = []
        if "xl/sharedStrings.xml" in archive.namelist():
            shared_root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            shared = ["".join(node.itertext()) for node in shared_root.findall("x:si", NS)]
        try:
            root = ET.fromstring(archive.read(sheet_file))
        except KeyError as error:
            raise ToolError(f"Worksheet data could not be read: {sheet_name}") from error
        rows: list[list[str]] = []
        for row in root.findall("x:sheetData/x:row", NS):
            values: dict[int, str] = {}
            for cell in row.findall("x:c", NS):
                index = col_number(cell.attrib.get("r", "A1"))
                cell_type = cell.attrib.get("t")
                if cell_type == "inlineStr":
                    value = "".join(cell.find("x:is", NS).itertext()) if cell.find("x:is", NS) is not None else ""
                else:
                    raw = cell.findtext("x:v", default="", namespaces=NS)
                    value = shared[int(raw)] if cell_type == "s" and raw.isdigit() and int(raw) < len(shared) else raw
                values[index] = value
            if values:
                rows.append([values.get(index, "") for index in range(max(values) + 1)])
    if not rows:
        raise ToolError("The selected worksheet is empty.")
    header_index = infer_header_row(rows)
    headers = [text(value) or f"Column {index + 1}" for index, value in enumerate(rows[header_index])]
    seen: Counter[str] = Counter()
    unique_headers: list[str] = []
    for header in headers:
        seen[header] += 1
        unique_headers.append(header if seen[header] == 1 else f"{header} ({seen[header]})")
    records = [
        {header: row[index] if index < len(row) else "" for index, header in enumerate(unique_headers)}
        for row in rows[header_index + 1:]
        if any(text(value) for value in row)
    ]
    if not records:
        raise ToolError("No proposal rows were found beneath the worksheet header.")
    return sheet_name, unique_headers, records


def infer_header_row(rows: list[list[str]]) -> int:
    best_index, best_score = 0, -1
    known = {normalize(alias) for aliases in FIELD_ALIASES.values() for alias in aliases}
    for index, row in enumerate(rows[:20]):
        values = [normalize(value) for value in row if text(value)]
        score = sum(3 if value in known else 1 if any(alias in value for alias in known) else 0 for value in values)
        score += min(len(values), 10) / 10
        if score > best_score:
            best_index, best_score = index, score
    return best_index


def infer_mapping(headers: list[str], supplied: dict[str, str] | None = None) -> tuple[dict[str, str | None], list[str]]:
    mapping: dict[str, str | None] = {}
    warnings: list[str] = []
    normalized_headers = {normalize(header): header for header in headers}
    for field, aliases in FIELD_ALIASES.items():
        chosen = (supplied or {}).get(field)
        if chosen:
            if chosen not in headers:
                raise ToolError(f"Mapping for {field!r} references an unknown column: {chosen!r}")
            mapping[field] = chosen
            continue
        exact = next((normalized_headers[normalize(alias)] for alias in aliases if normalize(alias) in normalized_headers), None)
        partial = next((header for header in headers if any(normalize(alias) in normalize(header) for alias in aliases)), None)
        mapping[field] = exact or partial
        if field in ("title", "abstract") and not mapping[field]:
            warnings.append(f"Could not infer a {field} column.")
    if not mapping["title"] and not mapping["abstract"]:
        raise ToolError("Could not infer either a title or abstract column. Provide an explicit mapping.")
    return mapping, warnings


def load_state(workspace: str | Path) -> tuple[Path, dict[str, Any]]:
    root = Path(workspace).expanduser().resolve()
    state_file = root / "session-review-state.json"
    if not state_file.is_file():
        raise ToolError("No review workspace exists there. Run import_sessions first.")
    try:
        return root, json.loads(state_file.read_text(encoding="utf-8"))
    except json.JSONDecodeError as error:
        raise ToolError("The review workspace state file is invalid JSON.") from error


def save_state(root: Path, state: dict[str, Any]) -> None:
    root.mkdir(parents=True, exist_ok=True)
    state_file = root / "session-review-state.json"
    with tempfile.NamedTemporaryFile("w", encoding="utf-8", dir=root, delete=False) as handle:
        json.dump(state, handle, ensure_ascii=False, indent=2)
        temp_name = handle.name
    Path(temp_name).replace(state_file)


def session_view(session: dict[str, Any], mapping: dict[str, str | None]) -> dict[str, Any]:
    fields = session["fields"]
    enrichment = session.get("enrichment", {})
    abstract = enrichment["description"] if "description" in enrichment else fields.get(mapping.get("abstract") or "", "")
    return {
        "id": session["id"],
        "title": fields.get(mapping.get("title") or "", ""),
        "abstract": abstract,
        "speaker": enrichment["speaker"] if "speaker" in enrichment else fields.get(mapping.get("speaker") or "", ""),
        "tags": fields.get(mapping.get("tags") or "", ""),
        "sessionize_url": enrichment.get("sessionize_url", ""),
        "format": enrichment.get("format", ""),
        "level": enrichment.get("level", ""),
        "status": enrichment.get("status", ""),
        "topic_category": enrichment.get("topic_category", "Other / Needs review"),
        "fields": fields,
    }


def quality_flags(view: dict[str, Any]) -> list[str]:
    flags: list[str] = []
    title, abstract, speaker = text(view["title"]), text(view["abstract"]), text(view["speaker"])
    if not title:
        flags.append("missing_title")
    if not abstract:
        flags.append("missing_abstract")
    elif len(abstract) < 80:
        flags.append("short_abstract")
    if not speaker:
        flags.append("missing_speaker")
    combined = normalize(title + " " + abstract)
    if any(phrase in combined for phrase in VAGUE_PHRASES):
        flags.append("vague_language")
    return flags


def jaccard(left: set[str], right: set[str]) -> float:
    return len(left & right) / len(left | right) if left or right else 0.0


def similarity(left: dict[str, Any], right: dict[str, Any]) -> tuple[float, list[str]]:
    title_left, title_right = tokens(left["title"]), tokens(right["title"])
    abstract_left, abstract_right = tokens(left["abstract"]), tokens(right["abstract"])
    tags_left, tags_right = tokens(left["tags"]), tokens(right["tags"])
    title_score, abstract_score, tag_score = jaccard(title_left, title_right), jaccard(abstract_left, abstract_right), jaccard(tags_left, tags_right)
    score = 0.62 * title_score + 0.30 * abstract_score + 0.08 * tag_score
    reasons: list[str] = []
    if title_left and title_left == title_right:
        score = max(score, 0.94)
        reasons.append("same normalized title")
    if title_score >= 0.6:
        reasons.append("high title overlap")
    if abstract_score >= 0.45:
        reasons.append("substantial abstract overlap")
    if tag_score >= 0.5:
        reasons.append("shared tags or track")
    return min(score, 1.0), reasons


def analyze(state: dict[str, Any]) -> dict[str, Any]:
    mapping = state["mapping"]
    views = [session_view(session, mapping) for session in state["sessions"]]
    flags = {view["id"]: quality_flags(view) for view in views}
    parents = {view["id"]: view["id"] for view in views}

    def find(item: str) -> str:
        while parents[item] != item:
            parents[item] = parents[parents[item]]
            item = parents[item]
        return item

    def union(left: str, right: str) -> None:
        a, b = find(left), find(right)
        if a != b:
            parents[b] = a

    pairs: dict[tuple[str, str], tuple[float, list[str]]] = {}
    for index, left in enumerate(views):
        for right in views[index + 1:]:
            score, reasons = similarity(left, right)
            if score >= 0.5 and reasons:
                pairs[(left["id"], right["id"])] = (round(score, 3), reasons)
                union(left["id"], right["id"])
    grouped: defaultdict[str, list[str]] = defaultdict(list)
    for view in views:
        grouped[find(view["id"])].append(view["id"])
    clusters: list[dict[str, Any]] = []
    for members in grouped.values():
        if len(members) < 2:
            continue
        pair_details = [
            {"session_ids": list(pair), "similarity": score, "reasons": reasons}
            for pair, (score, reasons) in pairs.items() if pair[0] in members and pair[1] in members
        ]
        if pair_details:
            clusters.append({"id": f"cluster-{len(clusters) + 1}", "session_ids": members, "pairs": pair_details})
    session_clusters = {member: cluster["id"] for cluster in clusters for member in cluster["session_ids"]}
    return {"quality_flags": flags, "clusters": clusters, "session_clusters": session_clusters, "analyzed_at": utc_now()}


def utc_now() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def review_score(review: dict[str, Any]) -> float:
    return round(sum(review["scores"].values()) / 5, 2)


def column_xml(value: Any, reference: str, style: int = 0) -> str:
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f'<c r="{reference}" s="{style}"><v>{value}</v></c>'
    content = xml_escape("" if value is None else value)
    preserve = ' xml:space="preserve"' if content[:1].isspace() or content[-1:].isspace() else ""
    return f'<c r="{reference}" t="inlineStr" s="{style}"><is><t{preserve}>{content}</t></is></c>'


def worksheet_xml(headers: list[str], rows: list[list[Any]]) -> str:
    all_rows = [headers] + rows
    max_cols = max((len(row) for row in all_rows), default=1)
    widths = []
    for index in range(max_cols):
        longest = max((len(str(row[index])) if index < len(row) else 0 for row in all_rows), default=10)
        widths.append(max(10, min(longest + 2, 45)))
    columns = "".join(f'<col min="{i + 1}" max="{i + 1}" width="{width}" customWidth="1"/>' for i, width in enumerate(widths))
    rendered_rows = []
    for row_number, row in enumerate(all_rows, start=1):
        cells = "".join(column_xml(value, f"{col_letter(index)}{row_number}", 1 if row_number == 1 else 0) for index, value in enumerate(row))
        rendered_rows.append(f'<row r="{row_number}">{cells}</row>')
    last_cell = f"{col_letter(max_cols - 1)}{max(1, len(all_rows))}"
    filter_ref = f"A1:{col_letter(max_cols - 1)}1" if headers else "A1:A1"
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="{MAIN_NS}" xmlns:r="{REL_NS}"><dimension ref="A1:{last_cell}"/><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/><cols>{columns}</cols><sheetData>{''.join(rendered_rows)}</sheetData><autoFilter ref="{filter_ref}"/></worksheet>'''


def write_xlsx(destination: Path, sheets: list[tuple[str, list[str], list[list[Any]]]]) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(suffix=".xlsx", dir=destination.parent, delete=False) as temporary:
        temp_path = Path(temporary.name)
    try:
        with zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as archive:
            content_types = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">', '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>', '<Default Extension="xml" ContentType="application/xml"/>', '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>', '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>']
            for index in range(len(sheets)):
                content_types.append(f'<Override PartName="/xl/worksheets/sheet{index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
            archive.writestr("[Content_Types].xml", "".join(content_types) + "</Types>")
            archive.writestr("_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
            sheet_nodes = "".join(f'<sheet name="{xml_escape(name)}" sheetId="{index + 1}" r:id="rId{index + 1}"/>' for index, (name, _, _) in enumerate(sheets))
            archive.writestr("xl/workbook.xml", f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="{MAIN_NS}" xmlns:r="{REL_NS}"><sheets>{sheet_nodes}</sheets></workbook>')
            rels = [f'<Relationship Id="rId{index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{index + 1}.xml"/>' for index in range(len(sheets))]
            rels.append(f'<Relationship Id="rId{len(sheets) + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>')
            archive.writestr("xl/_rels/workbook.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + "".join(rels) + "</Relationships>")
            styles = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="{MAIN_NS}"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="1" borderId="0" applyFont="1" applyFill="1" xfId="0"/></cellXfs></styleSheet>'''
            archive.writestr("xl/styles.xml", styles)
            for index, (_, headers, rows) in enumerate(sheets, start=1):
                archive.writestr(f"xl/worksheets/sheet{index}.xml", worksheet_xml(headers, rows))
        temp_path.replace(destination)
    finally:
        if temp_path.exists():
            temp_path.unlink()


def tool_import(args: dict[str, Any]) -> dict[str, Any]:
    source_path = Path(required(args, "source_path")).expanduser().resolve()
    workspace = Path(args.get("workspace_dir") or source_path.parent / ".session-review").expanduser().resolve()
    sheet, headers, rows = read_xlsx(source_path, args.get("sheet_name"))
    mapping, warnings = infer_mapping(headers, args.get("mapping"))
    sessions = []
    for index, fields in enumerate(rows, start=2):
        identifier = text(fields.get(mapping.get("id") or "")) or f"row-{index}"
        sessions.append({"id": identifier, "source_row": index, "fields": fields, "enrichment": {}})
    ids = [session["id"] for session in sessions]
    if len(ids) != len(set(ids)):
        raise ToolError("Session IDs are not unique. Map a unique ID column or remove duplicate IDs.")
    state = {"schema_version": 1, "source_path": str(source_path), "sheet_name": sheet, "headers": headers, "mapping": mapping, "sessions": sessions, "reviews": {}, "analysis": None, "imported_at": utc_now()}
    save_state(workspace, state)
    return {"workspace_dir": str(workspace), "sheet_name": sheet, "session_count": len(sessions), "mapping": mapping, "warnings": warnings, "sample_sessions": [session_view(session, mapping) for session in sessions[:3]]}


def tool_inspect(args: dict[str, Any]) -> dict[str, Any]:
    root, state = load_state(required(args, "workspace_dir"))
    return {"workspace_dir": str(root), "source_path": state["source_path"], "sheet_name": state["sheet_name"], "session_count": len(state["sessions"]), "headers": state["headers"], "mapping": state["mapping"], "sample_sessions": [session_view(session, state["mapping"]) for session in state["sessions"][:5]]}


def tool_analyze(args: dict[str, Any]) -> dict[str, Any]:
    root, state = load_state(required(args, "workspace_dir"))
    state["analysis"] = analyze(state)
    save_state(root, state)
    flag_counts = Counter(flag for flags in state["analysis"]["quality_flags"].values() for flag in flags)
    return {"workspace_dir": str(root), "quality_flag_counts": dict(flag_counts), "duplicate_cluster_count": len(state["analysis"]["clusters"]), "clusters": state["analysis"]["clusters"]}


def tool_enrich_details(args: dict[str, Any]) -> dict[str, Any]:
    """Persist detail-page data captured from Sessionize or another approved source."""
    root, state = load_state(required(args, "workspace_dir"))
    items = args.get("sessions")
    if not isinstance(items, list) or not items:
        raise ToolError("sessions must be a non-empty list.")
    allowed = {"id", "description", "sessionize_url", "speaker", "format", "level", "status"}
    by_id = {session["id"]: session for session in state["sessions"]}
    updated: set[str] = set()
    for incoming in items:
        if not isinstance(incoming, dict) or set(incoming) - allowed:
            raise ToolError("Each session detail may contain only id, description, sessionize_url, speaker, format, level, and status.")
        identifier = text(incoming.get("id"))
        if identifier not in by_id:
            raise ToolError(f"Unknown session ID: {identifier!r}")
        details = by_id[identifier].setdefault("enrichment", {})
        for name in allowed - {"id"}:
            if name in incoming:
                details[name] = text(incoming[name])
        updated.add(identifier)
    save_state(root, state)
    return {"workspace_dir": str(root), "updated_count": len(updated), "session_count": len(state["sessions"])}


def tool_categorize_topics(args: dict[str, Any]) -> dict[str, Any]:
    root, state = load_state(required(args, "workspace_dir"))
    mapping = state["mapping"]
    counts: Counter[str] = Counter()
    for session in state["sessions"]:
        view = session_view(session, mapping)
        category = topic_category(view["title"], view["abstract"], view["format"])
        session.setdefault("enrichment", {})["topic_category"] = category
        counts[category] += 1
    save_state(root, state)
    return {"workspace_dir": str(root), "session_count": len(state["sessions"]), "categories": dict(counts), "category_list": list(TOPIC_CATEGORIES)}


def tool_batch(args: dict[str, Any]) -> dict[str, Any]:
    _, state = load_state(required(args, "workspace_dir"))
    limit = int(args.get("limit", 10))
    if not 1 <= limit <= 15:
        raise ToolError("limit must be between 1 and 15.")
    status = args.get("status", "unreviewed")
    if status not in ("unreviewed", "reviewed", "all"):
        raise ToolError("status must be unreviewed, reviewed, or all.")
    items = state["sessions"]
    if status == "unreviewed":
        items = [item for item in items if item["id"] not in state["reviews"]]
    elif status == "reviewed":
        items = [item for item in items if item["id"] in state["reviews"]]
    cursor = int(args.get("cursor", 0))
    if cursor < 0:
        raise ToolError("cursor cannot be negative.")
    analysis_data = state.get("analysis") or {"quality_flags": {}, "session_clusters": {}}
    sessions = []
    for session in items[cursor:cursor + limit]:
        view = session_view(session, state["mapping"])
        view["quality_flags"] = analysis_data["quality_flags"].get(session["id"], [])
        view["duplicate_cluster"] = analysis_data["session_clusters"].get(session["id"])
        view["review"] = state["reviews"].get(session["id"])
        sessions.append(view)
    next_cursor = cursor + len(sessions)
    return {"sessions": sessions, "next_cursor": next_cursor if next_cursor < len(items) else None, "remaining": max(0, len(items) - next_cursor)}


def tool_record(args: dict[str, Any]) -> dict[str, Any]:
    root, state = load_state(required(args, "workspace_dir"))
    reviews = args.get("reviews")
    if not isinstance(reviews, list) or not reviews:
        raise ToolError("reviews must be a non-empty list.")
    known_ids = {session["id"] for session in state["sessions"]}
    validated: dict[str, dict[str, Any]] = {}
    required_scores = ("clarity_depth", "audience_fit", "speaker_credibility", "originality", "delivery_readiness")
    for incoming in reviews:
        identifier = text(incoming.get("id"))
        if identifier not in known_ids:
            raise ToolError(f"Unknown session ID: {identifier!r}")
        if identifier in validated:
            raise ToolError(f"Session ID is repeated in this batch: {identifier!r}")
        scores = incoming.get("scores")
        if not isinstance(scores, dict) or set(scores) != set(required_scores):
            raise ToolError("Each review must include exactly the five rubric score keys.")
        if any(not isinstance(scores[key], (int, float)) or isinstance(scores[key], bool) or not 1 <= scores[key] <= 5 for key in required_scores):
            raise ToolError("Every rubric score must be a number from 1 to 5.")
        decision = incoming.get("decision")
        if decision not in ("accept", "maybe", "reject"):
            raise ToolError("decision must be accept, maybe, or reject.")
        review = {"scores": {key: scores[key] for key in required_scores}, "decision": decision, "comment": text(incoming.get("comment")), "average": round(sum(scores.values()) / 5, 2), "reviewed_at": utc_now()}
        validated[identifier] = review
    state["reviews"].update(validated)
    save_state(root, state)
    return {"saved_count": len(validated), "reviewed_total": len(state["reviews"]), "session_count": len(state["sessions"])}


def tool_cluster(args: dict[str, Any]) -> dict[str, Any]:
    _, state = load_state(required(args, "workspace_dir"))
    if not state.get("analysis"):
        raise ToolError("Run analyze_sessions before requesting duplicate clusters.")
    cluster_id = required(args, "cluster_id")
    cluster = next((cluster for cluster in state["analysis"]["clusters"] if cluster["id"] == cluster_id), None)
    if not cluster:
        raise ToolError("Duplicate cluster not found.")
    by_id = {session["id"]: session for session in state["sessions"]}
    sessions = []
    for identifier in cluster["session_ids"]:
        view = session_view(by_id[identifier], state["mapping"])
        view["review"] = state["reviews"].get(identifier)
        sessions.append(view)
    return {"cluster": cluster, "sessions": sessions}


def tool_progress(args: dict[str, Any]) -> dict[str, Any]:
    _, state = load_state(required(args, "workspace_dir"))
    decisions = Counter(review["decision"] for review in state["reviews"].values())
    analysis_data = state.get("analysis") or {"clusters": [], "quality_flags": {}}
    unresolved = [cluster["id"] for cluster in analysis_data["clusters"] if not all(identifier in state["reviews"] for identifier in cluster["session_ids"])]
    return {"session_count": len(state["sessions"]), "reviewed_count": len(state["reviews"]), "remaining_count": len(state["sessions"]) - len(state["reviews"]), "decisions": dict(decisions), "duplicate_cluster_count": len(analysis_data["clusters"]), "unresolved_duplicate_clusters": unresolved, "quality_flagged_sessions": sum(bool(flags) for flags in analysis_data["quality_flags"].values())}


def tool_export(args: dict[str, Any]) -> dict[str, Any]:
    root, state = load_state(required(args, "workspace_dir"))
    destination = Path(args.get("output_path") or root / "reviewed-sessions.xlsx").expanduser().resolve()
    if destination.suffix.lower() != ".xlsx":
        raise ToolError("output_path must end in .xlsx.")
    if destination.exists() and not args.get("overwrite", False):
        raise ToolError("Output file already exists. Choose a new output_path or set overwrite to true.")
    analysis_data = state.get("analysis") or {"quality_flags": {}, "session_clusters": {}, "clusters": []}
    enrichment_columns = ["Sessionize URL", "Session Format", "Session Level", "Session Status", "Topic Category"]
    review_columns = ["Clarity & Depth", "Audience Fit", "Speaker Credibility", "Originality", "Delivery Readiness", "Average", "Decision", "Reviewer Comment", "Quality Flags", "Duplicate Cluster"]
    reviewed_rows: list[list[Any]] = []
    output_sessions: list[tuple[dict[str, Any], dict[str, Any] | None]] = []
    for session in state["sessions"]:
        review = state["reviews"].get(session["id"])
        output_sessions.append((session, review))
        scores = review["scores"] if review else {}
        view = session_view(session, state["mapping"])
        reviewed_rows.append([session["fields"].get(header, "") for header in state["headers"]] + [view["sessionize_url"], view["format"], view["level"], view["status"], view["topic_category"], scores.get("clarity_depth", ""), scores.get("audience_fit", ""), scores.get("speaker_credibility", ""), scores.get("originality", ""), scores.get("delivery_readiness", ""), review.get("average", "") if review else "", review.get("decision", "") if review else "", review.get("comment", "") if review else "", "; ".join(analysis_data["quality_flags"].get(session["id"], [])), analysis_data["session_clusters"].get(session["id"], "")])
    order = {"accept": 0, "maybe": 1, "reject": 2}
    shortlist = sorted(((session, review) for session, review in output_sessions if review and review["decision"] in order), key=lambda item: (order[item[1]["decision"]], -item[1]["average"], normalize(session_view(item[0], state["mapping"])["title"])))
    shortlist_headers = ["Session ID", "Title", "Speaker", "Decision", "Average", "Reviewer Comment", "Duplicate Cluster"]
    shortlist_rows = []
    for session, review in shortlist:
        view = session_view(session, state["mapping"])
        shortlist_rows.append([session["id"], view["title"], view["speaker"], review["decision"], review["average"], review["comment"], analysis_data["session_clusters"].get(session["id"], "")])
    cluster_headers = ["Cluster", "Session ID", "Title", "Speaker", "Decision", "Average", "Similarity evidence"]
    cluster_rows = []
    by_id = {session["id"]: session for session in state["sessions"]}
    for cluster in analysis_data["clusters"]:
        evidence = "; ".join(f"{pair['similarity']}: {', '.join(pair['reasons'])}" for pair in cluster["pairs"])
        for identifier in cluster["session_ids"]:
            session, review = by_id[identifier], state["reviews"].get(identifier)
            view = session_view(session, state["mapping"])
            cluster_rows.append([cluster["id"], identifier, view["title"], view["speaker"], review["decision"] if review else "", review["average"] if review else "", evidence])
    write_xlsx(destination, [("Reviewed Sessions", state["headers"] + enrichment_columns + review_columns, reviewed_rows), ("Shortlist", shortlist_headers, shortlist_rows), ("Duplicate Clusters", cluster_headers, cluster_rows)])
    return {"output_path": str(destination), "reviewed_count": len(state["reviews"]), "session_count": len(state["sessions"]), "partial_review": len(state["reviews"]) < len(state["sessions"])}


def required(args: dict[str, Any], name: str) -> Any:
    value = args.get(name)
    if value is None or value == "":
        raise ToolError(f"{name} is required.")
    return value


TOOL_DEFINITIONS = [
    ("import_sessions", "Import an .xlsx session workbook, infer proposal fields, and create a local review workspace.", {"source_path": {"type": "string"}, "workspace_dir": {"type": "string"}, "sheet_name": {"type": "string"}, "mapping": {"type": "object", "additionalProperties": {"type": "string"}}}, ["source_path"]),
    ("inspect_import", "Inspect imported fields, mappings, and sample sessions.", {"workspace_dir": {"type": "string"}}, ["workspace_dir"]),
    ("enrich_session_details", "Save read-only detail-page data such as submitted descriptions and Sessionize metadata into the local review workspace.", {"workspace_dir": {"type": "string"}, "sessions": {"type": "array", "minItems": 1, "items": {"type": "object"}}}, ["workspace_dir", "sessions"]),
    ("categorize_topics", "Assign a deterministic topic category to every proposal from its title, description, and format.", {"workspace_dir": {"type": "string"}}, ["workspace_dir"]),
    ("analyze_sessions", "Flag incomplete proposals and find conservative duplicate clusters.", {"workspace_dir": {"type": "string"}}, ["workspace_dir"]),
    ("get_review_batch", "Fetch a bounded batch of session proposals for scoring.", {"workspace_dir": {"type": "string"}, "limit": {"type": "integer", "minimum": 1, "maximum": 15}, "cursor": {"type": "integer", "minimum": 0}, "status": {"type": "string", "enum": ["unreviewed", "reviewed", "all"]}}, ["workspace_dir"]),
    ("record_reviews", "Persist a batch of validated rubric scores, decisions, and comments.", {"workspace_dir": {"type": "string"}, "reviews": {"type": "array", "minItems": 1, "items": {"type": "object"}}}, ["workspace_dir", "reviews"]),
    ("get_duplicate_cluster", "Fetch every proposal in a similarity cluster for editorial comparison.", {"workspace_dir": {"type": "string"}, "cluster_id": {"type": "string"}}, ["workspace_dir", "cluster_id"]),
    ("review_progress", "Summarize review decisions and unresolved duplicate clusters.", {"workspace_dir": {"type": "string"}}, ["workspace_dir"]),
    ("export_reviews", "Export scores, flags, comments, shortlist, and duplicate clusters to a new .xlsx workbook.", {"workspace_dir": {"type": "string"}, "output_path": {"type": "string"}, "overwrite": {"type": "boolean"}}, ["workspace_dir"]),
]
TOOL_FUNCTIONS = {"import_sessions": tool_import, "inspect_import": tool_inspect, "enrich_session_details": tool_enrich_details, "categorize_topics": tool_categorize_topics, "analyze_sessions": tool_analyze, "get_review_batch": tool_batch, "record_reviews": tool_record, "get_duplicate_cluster": tool_cluster, "review_progress": tool_progress, "export_reviews": tool_export}


def tool_schema() -> list[dict[str, Any]]:
    return [{"name": name, "description": description, "inputSchema": {"type": "object", "properties": properties, "required": required_fields, "additionalProperties": False}} for name, description, properties, required_fields in TOOL_DEFINITIONS]


def respond(payload: dict[str, Any]) -> None:
    sys.stdout.write(json.dumps(payload, ensure_ascii=False) + "\n")
    sys.stdout.flush()


def main() -> None:
    for raw in sys.stdin:
        request: dict[str, Any] | None = None
        try:
            request = json.loads(raw)
            if not isinstance(request, dict):
                raise ToolError("MCP requests must be JSON objects.")
            method = request.get("method")
            request_id = request.get("id")
            if method == "notifications/initialized":
                continue
            if method == "initialize":
                result = {"protocolVersion": request.get("params", {}).get("protocolVersion", "2024-11-05"), "capabilities": {"tools": {}}, "serverInfo": {"name": "session-reviewer", "version": "0.1.0"}}
            elif method == "tools/list":
                result = {"tools": tool_schema()}
            elif method == "tools/call":
                params = request.get("params", {})
                name = params.get("name")
                if name not in TOOL_FUNCTIONS:
                    raise ToolError(f"Unknown tool: {name}")
                result = {"content": [{"type": "text", "text": json.dumps(TOOL_FUNCTIONS[name](params.get("arguments") or {}), ensure_ascii=False, indent=2)}]}
            else:
                raise ToolError(f"Unsupported MCP method: {method}")
            if request_id is not None:
                respond({"jsonrpc": "2.0", "id": request_id, "result": result})
        except Exception as error:  # Convert all operational errors to MCP tool errors.
            if isinstance(request, dict) and request.get("id") is not None:
                if request.get("method") == "tools/call":
                    respond({"jsonrpc": "2.0", "id": request["id"], "result": {"content": [{"type": "text", "text": str(error)}], "isError": True}})
                else:
                    respond({"jsonrpc": "2.0", "id": request["id"], "error": {"code": -32603, "message": str(error)}})


if __name__ == "__main__":
    main()
