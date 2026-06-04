# =========================================================
# INSTALL IF NEEDED
# =========================================================
# %pip install pymupdf requests json-repair XlsxWriter
import fitz  # PyMuPDF
import requests
import json
import os
import re
import shutil
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
from collections import Counter
from datetime import datetime
from json_repair import repair_json
import xlsxwriter


# =========================================================
# 1. CONFIGURATION
# =========================================================

LM_STUDIO_URL = "http://127.0.0.1:1234/v1/chat/completions"
LM_STUDIO_API_KEY = "lm-studio"
MODEL_NAME = "mrademacher/MedGemma-4B-Instruct-ft-2-GGUF"

# Safer local-model chunk settings.
# Smaller chunks reduce missed table rows and context errors.
ENDPOINT_TEXT_CHUNK_SIZE = 3500
SCHEDULE_TEXT_CHUNK_SIZE = 3500
TEXT_CHUNK_OVERLAP = 100

# Smaller classification batches are safer for MedGemma 4B.
ROWS_PER_CLASSIFICATION_BATCH = 5

# If the TOC cannot be trusted, these limits avoid sending too many pages.
MAX_ENDPOINT_PAGES = 18

# Schedule of Activities/Assessments is usually short.
# The code first uses the TOC boundary:
# schedule section start page -> page before the next major TOC section.
# It also includes relevant schedule/sampling/data-collection tables from List of Tables/bookmarks.
MAX_SCHEDULE_PAGES = 8
MAX_SCHEDULE_SECTION_SPAN = 5
STRICT_TOC_SCHEDULE_SECTION_ONLY = True

MAX_SUPPORTING_PAGES = 20


# =========================================================
# 2. REUSABLE TARGET CONCEPTS
# =========================================================

ENDPOINT_SECTION_TERMS = [
    "study objectives",
    "objectives",
    "objectives and endpoints",
    "endpoints",
    "endpoint",
    "outcome measures",
    "outcome measure",
    "primary endpoint",
    "secondary endpoint",
    "exploratory endpoint",
    "primary objective",
    "secondary objective",
    "exploratory objective"
]

SCHEDULE_SECTION_TERMS = [
    "schedule of assessments",
    "schedule of assessment",
    "schedule of activities",
    "schedule of activity",
    "schedule of events",
    "assessment schedule",
    "data collection schedule",
    "data collection table",
    "study procedures schedule",
    "sampling schedule",
    "visit schedule"
]

SUPPORTING_SECTION_TERMS = [
    "safety reporting",
    "adverse event reporting",
    "serious adverse event reporting",
    "adverse events",
    "safety variables",
    "safety assessments",
    "statistical methods",
    "statistical considerations",
    "statistical analysis",
    "data management",
    "electronic case report form",
    "source data",
    "source documents",
    "good clinical practice",
    "informed consent"
]

NOISE_SECTION_TERMS = [
    "abbreviations",
    "definition of terms",
    "reference list",
    "background",
    "rationale",
    "publication policy",
    "confidentiality"
]

VALID_CATEGORIES = [
    "primary_endpoint_related_safety",
    "primary_endpoint_related_efficacy",
    "secondary_endpoint_related_safety",
    "secondary_endpoint_related_efficacy",
    "additional_or_exploratory_endpoint_related",
    "general_study_operations_related_to_gcp",
    "not_classifiable"
]


# =========================================================
# 3. GCP REFERENCE ACTIVITIES
# =========================================================

GCP_REFERENCE_ACTIVITIES = {
    "ethics_and_subject_protection": [
        "informed consent",
        "re-consent",
        "ethics committee approval",
        "institutional review board approval",
        "subject rights",
        "subject safety",
        "privacy",
        "confidentiality"
    ],
    "eligibility_and_enrollment": [
        "inclusion criteria",
        "exclusion criteria",
        "eligibility assessment",
        "screening",
        "enrollment",
        "randomization",
        "screen failure",
        "withdrawal",
        "early termination"
    ],
    "baseline_and_demographics": [
        "demographics",
        "medical history",
        "surgical history",
        "baseline characteristics",
        "physical examination",
        "vital signs",
        "pregnancy test",
        "concomitant medications",
        "prior medications"
    ],
    "safety_reporting": [
        "adverse event",
        "serious adverse event",
        "device deficiency",
        "adverse device effect",
        "safety reporting",
        "safety follow-up",
        "medical monitor review"
    ],
    "protocol_compliance": [
        "protocol deviation",
        "protocol violation",
        "visit window",
        "missed visit",
        "unscheduled visit",
        "study discontinuation",
        "lost to follow-up"
    ],
    "data_integrity_and_source": [
        "source data",
        "source document",
        "case report form",
        "CRF",
        "electronic data capture",
        "EDC",
        "data entry",
        "query resolution",
        "audit trail",
        "database lock",
        "data management"
    ],
    "monitoring_and_quality": [
        "monitoring",
        "source data verification",
        "quality control",
        "quality assurance",
        "audit",
        "inspection readiness"
    ],
    "essential_records": [
        "essential documents",
        "trial master file",
        "investigator site file",
        "record retention",
        "archiving",
        "version control"
    ],
    "study_product_or_device_accountability": [
        "investigational product accountability",
        "device accountability",
        "device tracking",
        "product storage",
        "product handling",
        "blinding",
        "unblinding"
    ]
}


def build_gcp_reference_text():
    lines = []
    for category, activities in GCP_REFERENCE_ACTIVITIES.items():
        lines.append(f"{category}:")
        for activity in activities:
            lines.append(f"- {activity}")
        lines.append("")
    return "\n".join(lines)


# =========================================================
# 4. GENERAL UTILITIES
# =========================================================

def safe_folder_name(value):
    value = os.path.splitext(os.path.basename(value))[0]
    value = re.sub(r"[^A-Za-z0-9_\-]+", "_", value)
    value = re.sub(r"_+", "_", value)
    return value.strip("_") or "protocol"


def create_output_paths(pdf_path, output_root):
    base_name = safe_folder_name(pdf_path)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    run_folder = os.path.join(output_root, f"{base_name}_analysis_{timestamp}")
    metadata_folder = os.path.join(run_folder, "metadata")

    os.makedirs(run_folder, exist_ok=True)
    os.makedirs(metadata_folder, exist_ok=True)

    output_json_path = os.path.join(run_folder, f"{base_name}_classification_output.json")
    output_excel_path = os.path.join(run_folder, f"{base_name}_classification_output.xlsx")

    return {
        "base_name": base_name,
        "run_folder": run_folder,
        "metadata_folder": metadata_folder,
        "output_json_path": output_json_path,
        "output_excel_path": output_excel_path
    }


def archive_previous_run_folders(output_root, base_name, update_callback):
    archive_root = os.path.join(output_root, "archive")
    os.makedirs(archive_root, exist_ok=True)

    moved = 0

    for item_name in os.listdir(output_root):
        item_path = os.path.join(output_root, item_name)

        if item_name == "archive":
            continue

        if not os.path.isdir(item_path):
            continue

        if item_name.startswith(f"{base_name}_analysis_"):
            destination = os.path.join(archive_root, item_name)

            if os.path.exists(destination):
                destination = os.path.join(
                    archive_root,
                    f"{item_name}_archived_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                )

            shutil.move(item_path, destination)
            moved += 1

    if moved:
        update_callback(f"Archived {moved} previous run folder(s).")
    else:
        update_callback("No previous run folders found to archive.")

    return archive_root


def normalize_text(value):
    value = str(value).lower()
    value = re.sub(r"[^a-z0-9]+", " ", value)
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def contains_any(text, terms):
    text_norm = normalize_text(text)
    return any(normalize_text(term) in text_norm for term in terms)


def count_term_matches(text, terms):
    text_norm = normalize_text(text)
    return sum(1 for term in terms if normalize_text(term) in text_norm)


def split_text(text, chunk_size=6000, overlap=250):
    chunks = []
    start = 0

    while start < len(text):
        end = min(start + chunk_size, len(text))
        chunks.append(text[start:end])

        if end == len(text):
            break

        start = max(0, end - overlap)

    return chunks


# =========================================================
# 5. LM STUDIO CALL AND JSON PARSING
# =========================================================

def call_lm_studio(messages, temperature=0.0, max_tokens=1200):
    payload = {
        "model": MODEL_NAME,
        "messages": messages,
        "temperature": temperature,
        "max_tokens": max_tokens
    }

    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {LM_STUDIO_API_KEY}"
    }

    response = requests.post(
        LM_STUDIO_URL,
        headers=headers,
        json=payload,
        timeout=900
    )

    if response.status_code != 200:
        raise RuntimeError(
            f"LM Studio error.\n"
            f"Status code: {response.status_code}\n"
            f"Response:\n{response.text}"
        )

    data = response.json()

    try:
        return data["choices"][0]["message"]["content"]
    except Exception:
        raise RuntimeError(
            "LM Studio response did not contain choices[0].message.content.\n"
            f"Raw response:\n{json.dumps(data, indent=4)}"
        )


def clean_json_text(text):
    text = str(text).strip()
    text = text.replace("```json", "")
    text = text.replace("```", "")
    text = text.strip()

    start = text.find("{")
    end = text.rfind("}")

    if start == -1 or end == -1 or end <= start:
        raise ValueError("No valid JSON object found in AI response.")

    return text[start:end + 1]


def default_json_for_prefix(prefix):
    if prefix.startswith("toc_section_map"):
        return {
            "endpoint_sections": [],
            "schedule_sections": [],
            "supporting_sections": [],
            "parse_warning": "AI TOC response could not be parsed."
        }

    if prefix.startswith("study_map"):
        return {
            "reference_concepts": [],
            "study_objectives": [],
            "endpoints": {
                "primary": [],
                "secondary": [],
                "additional_or_exploratory": []
            },
            "study_logic_notes": [
                "AI response could not be parsed for study map."
            ]
        }

    if prefix.startswith("schedule_rows"):
        return {
            "schedule_rows": [],
            "parse_warning": "AI response could not be parsed for schedule row extraction."
        }

    if prefix.startswith("classification_batch"):
        return {
            "classified_items": [],
            "parse_warning": "AI response could not be parsed for classification batch."
        }

    return {
        "parse_warning": "AI response could not be parsed."
    }


def parse_ai_json(ai_response, metadata_folder, output_name_prefix):
    os.makedirs(metadata_folder, exist_ok=True)

    raw_output_path = os.path.join(metadata_folder, f"{output_name_prefix}_raw_ai_response.txt")
    repaired_output_path = os.path.join(metadata_folder, f"{output_name_prefix}_repaired_ai_response.json")
    parsed_output_path = os.path.join(metadata_folder, f"{output_name_prefix}_parsed_ai_response.json")
    failed_output_path = os.path.join(metadata_folder, f"{output_name_prefix}_parse_failed.json")

    with open(raw_output_path, "w", encoding="utf-8") as f:
        f.write(str(ai_response))

    try:
        json_text = clean_json_text(ai_response)

        try:
            parsed_data = json.loads(json_text)
        except json.JSONDecodeError:
            repaired_json_text = repair_json(json_text)

            with open(repaired_output_path, "w", encoding="utf-8") as f:
                f.write(repaired_json_text)

            parsed_data = json.loads(repaired_json_text)

        with open(parsed_output_path, "w", encoding="utf-8") as f:
            json.dump(parsed_data, f, indent=4, ensure_ascii=False)

        return parsed_data

    except Exception as e:
        fallback_data = default_json_for_prefix(output_name_prefix)
        fallback_data["parse_error"] = str(e)
        fallback_data["raw_response_file"] = raw_output_path

        with open(failed_output_path, "w", encoding="utf-8") as f:
            json.dump(fallback_data, f, indent=4, ensure_ascii=False)

        print(f"WARNING: Could not parse AI response for {output_name_prefix}.")
        print(f"Raw response saved here: {raw_output_path}")
        return fallback_data


# =========================================================
# 6. PDF READING
# =========================================================

def read_pdf_pages(pdf_path):
    pages = []

    with fitz.open(pdf_path) as doc:
        for page_index, page in enumerate(doc):
            page_number = page_index + 1
            page_text = page.get_text()
            table_text_parts = []

            try:
                tables = page.find_tables()

                for table_index, table in enumerate(tables, start=1):
                    extracted = table.extract()

                    table_text_parts.append(
                        f"\n\n--- EXTRACTED TABLE {table_index} ON PDF PAGE {page_number} ---\n"
                    )

                    for row in extracted:
                        clean_cells = [
                            str(cell).strip() if cell is not None else ""
                            for cell in row
                        ]
                        table_text_parts.append(" | ".join(clean_cells))

            except Exception:
                pass

            if table_text_parts:
                page_text += "\n\n" + "\n".join(table_text_parts)

            pages.append({
                "page_number": page_number,
                "text": page_text
            })

    return pages


def extract_pdf_bookmarks(pdf_path):
    bookmarks = []

    with fitz.open(pdf_path) as doc:
        toc = doc.get_toc(simple=True)

        for item in toc:
            level, title, page_number = item
            bookmarks.append({
                "level": int(level),
                "title": str(title).strip(),
                "page_number": int(page_number)
            })

    return bookmarks


def extract_initial_toc_text(pages, max_pages=10):
    text = ""

    for page in pages[:min(max_pages, len(pages))]:
        text += f"\n\n--- PDF PAGE {page['page_number']} ---\n\n"
        text += page["text"]

    return text


def pages_to_text(pages, page_numbers):
    page_set = set(page_numbers)
    text = ""

    for page in pages:
        if page["page_number"] in page_set:
            text += f"\n\n--- PDF PAGE {page['page_number']} ---\n\n"
            text += page["text"]

    return text


# =========================================================
# 7. TOC AND TARGET PAGE IDENTIFICATION
# =========================================================

def extract_page_number_from_toc_line(line):
    match = re.search(r"(\d{1,4})\s*$", line.strip())

    if match:
        return int(match.group(1))

    return 0


def looks_like_reliable_toc_or_table_line(line):
    stripped = line.strip()

    if not stripped:
        return False

    if re.search(r"\.{3,}\s*\d{1,4}\s*$", stripped):
        return True

    if re.match(r"^\d+(\.\d+)*\.?\s+[A-Za-z].*\s+\d{1,4}$", stripped):
        return True

    if re.match(r"^Table\s+\d+[\-\.\dA-Za-z]*\s+.+\s+\d{1,4}$", stripped, flags=re.IGNORECASE):
        return True

    return False


def clean_toc_title(line):
    title = re.sub(r"\s*\d{1,4}\s*$", "", line).strip()
    title = re.sub(r"\.{2,}", " ", title).strip()
    title = re.sub(r"\s+", " ", title).strip()
    return title


def extract_structural_candidates_from_text(toc_text, total_pages):
    endpoint_candidates = []
    schedule_candidates = []
    supporting_candidates = []

    for raw_line in toc_text.splitlines():
        line = raw_line.strip()

        if not looks_like_reliable_toc_or_table_line(line):
            continue

        page_number = extract_page_number_from_toc_line(line)

        if page_number < 1 or page_number > total_pages:
            continue

        title_without_page = clean_toc_title(line)

        if contains_any(title_without_page, NOISE_SECTION_TERMS):
            continue

        is_table_line = bool(re.match(r"^Table\s+\d+", title_without_page, flags=re.IGNORECASE))

        candidate = {
            "section_title": title_without_page,
            "start_page": page_number,
            "end_page": page_number,
            "source": "written_toc_or_list_of_tables",
            "is_table_line": is_table_line,
            "reason": "Matched reliable written TOC/List of Tables line."
        }

        if contains_any(title_without_page, SCHEDULE_SECTION_TERMS):
            schedule_candidates.append(dict(candidate))
        elif contains_any(title_without_page, ENDPOINT_SECTION_TERMS):
            endpoint_candidates.append(dict(candidate))
        elif contains_any(title_without_page, SUPPORTING_SECTION_TERMS):
            supporting_candidates.append(dict(candidate))

    return endpoint_candidates, schedule_candidates, supporting_candidates


def extract_all_table_entries_from_toc_text(toc_text, total_pages):
    table_entries = []

    for raw_line in toc_text.splitlines():
        line = raw_line.strip()

        if not re.match(r"^Table\s+\d+[\-\.\dA-Za-z]*\s+", line, flags=re.IGNORECASE):
            continue

        page_number = extract_page_number_from_toc_line(line)

        if page_number < 1 or page_number > total_pages:
            continue

        title = clean_toc_title(line)

        table_entries.append({
            "section_title": title,
            "start_page": page_number,
            "end_page": page_number,
            "source": "list_of_tables",
            "is_table_line": True,
            "reason": "Matched List of Tables entry."
        })

    seen = set()
    unique_entries = []

    for entry in table_entries:
        key = (normalize_text(entry["section_title"]), entry["start_page"])

        if key not in seen:
            unique_entries.append(entry)
            seen.add(key)

    return sorted(unique_entries, key=lambda x: x["start_page"])


def infer_table_spans(table_entries, pages, max_span=8):
    if not table_entries:
        return []

    total_pages = len(pages)
    sorted_entries = sorted(table_entries, key=lambda x: x["start_page"])
    spanned_entries = []

    for index, entry in enumerate(sorted_entries):
        start_page = int(entry["start_page"])

        if index + 1 < len(sorted_entries):
            next_start = int(sorted_entries[index + 1]["start_page"])
            end_page = max(start_page, next_start - 1)
        else:
            end_page = min(total_pages, start_page + max_span - 1)

        end_page = min(end_page, start_page + max_span - 1, total_pages)
        refined_end_page = end_page

        for page_number in range(start_page + 1, end_page + 1):
            text_upper = pages[page_number - 1]["text"][:1000].upper()

            if (
                "TABLE OF CONTENTS" in text_upper
                or "LIST OF ABBREVIATIONS" in text_upper
                or re.search(r"\n\s*\d+\.\s+[A-Z][A-Z ]{4,}", "\n" + text_upper)
            ):
                refined_end_page = page_number - 1
                break

        spanned = dict(entry)
        spanned["end_page"] = max(start_page, refined_end_page)
        spanned["reason"] = entry.get("reason", "") + " Span inferred from neighboring List of Tables entries."
        spanned_entries.append(spanned)

    return spanned_entries


def extract_numbered_toc_entries_from_text(toc_text, total_pages):
    entries = []

    for raw_line in toc_text.splitlines():
        line = raw_line.strip()

        if not line:
            continue

        match = re.match(
            r"^(\d+(?:\.\d+)*)\.?\s+(.+?)(?:\.{2,}|\s{2,}|\s+)(\d{1,4})$",
            line
        )

        if not match:
            continue

        section_number = match.group(1).strip()
        title = match.group(2).strip()
        page_number = int(match.group(3))

        if page_number < 1 or page_number > total_pages:
            continue

        title = re.sub(r"\.{2,}", " ", title).strip()
        title = re.sub(r"\s+", " ", title).strip()

        if contains_any(title, NOISE_SECTION_TERMS):
            continue

        level = section_number.count(".") + 1

        entries.append({
            "section_number": section_number,
            "level": level,
            "section_title": title,
            "start_page": page_number,
            "end_page": page_number,
            "source": "numbered_toc",
            "reason": "Matched numbered TOC section."
        })

    unique = []
    seen = set()

    for entry in entries:
        key = (entry["section_number"], normalize_text(entry["section_title"]), entry["start_page"])

        if key not in seen:
            unique.append(entry)
            seen.add(key)

    return unique


def infer_numbered_toc_section_spans(toc_entries, total_pages, max_span=5):
    spanned_entries = []

    for index, entry in enumerate(toc_entries):
        start_page = int(entry["start_page"])
        level = int(entry["level"])
        next_boundary_page = None

        for next_entry in toc_entries[index + 1:]:
            next_start = int(next_entry["start_page"])
            next_level = int(next_entry["level"])

            if next_start > start_page and next_level <= level:
                next_boundary_page = next_start
                break

        if next_boundary_page:
            end_page = next_boundary_page - 1
        else:
            end_page = start_page + max_span - 1

        end_page = min(end_page, start_page + max_span - 1, total_pages)
        end_page = max(start_page, end_page)

        spanned = dict(entry)
        spanned["end_page"] = end_page
        spanned["reason"] = (
            entry.get("reason", "")
            + " End page inferred from next major TOC section and capped by MAX_SCHEDULE_SECTION_SPAN."
        )
        spanned_entries.append(spanned)

    return spanned_entries


def get_strict_schedule_sections_from_numbered_toc(toc_text, total_pages):
    toc_entries = extract_numbered_toc_entries_from_text(toc_text, total_pages)
    spanned_entries = infer_numbered_toc_section_spans(
        toc_entries,
        total_pages,
        max_span=MAX_SCHEDULE_SECTION_SPAN
    )

    schedule_sections = []

    for entry in spanned_entries:
        title = entry.get("section_title", "")

        if contains_any(title, SCHEDULE_SECTION_TERMS):
            schedule_sections.append(entry)

    return schedule_sections, toc_entries


def extract_structural_candidates_from_bookmarks(bookmarks, total_pages):
    endpoint_candidates = []
    schedule_candidates = []
    supporting_candidates = []

    for i, bookmark in enumerate(bookmarks):
        title = bookmark["title"]
        start_page = bookmark["page_number"]
        level = bookmark["level"]

        if start_page < 1 or start_page > total_pages:
            continue

        end_page = start_page

        for next_bookmark in bookmarks[i + 1:]:
            if next_bookmark["level"] <= level and next_bookmark["page_number"] > start_page:
                end_page = next_bookmark["page_number"] - 1
                break

        if end_page < start_page:
            end_page = start_page

        candidate = {
            "section_title": title,
            "start_page": start_page,
            "end_page": end_page,
            "source": "pdf_bookmark",
            "reason": "Matched PDF bookmark."
        }

        if contains_any(title, SCHEDULE_SECTION_TERMS):
            schedule_candidates.append(dict(candidate))
        elif contains_any(title, ENDPOINT_SECTION_TERMS):
            endpoint_candidates.append(dict(candidate))
        elif contains_any(title, SUPPORTING_SECTION_TERMS):
            supporting_candidates.append(dict(candidate))

    return endpoint_candidates, schedule_candidates, supporting_candidates


def normalize_candidate(candidate, total_pages, default_span=2, max_span=8):
    start_page = candidate.get("start_page", 0)
    end_page = candidate.get("end_page", 0)

    try:
        start_page = int(start_page)
    except Exception:
        start_page = 0

    try:
        end_page = int(end_page)
    except Exception:
        end_page = 0

    if start_page < 1 or start_page > total_pages:
        return None

    if end_page < start_page:
        end_page = start_page + default_span

    end_page = min(end_page, start_page + max_span - 1)
    end_page = min(end_page, total_pages)

    clean = dict(candidate)
    clean["start_page"] = start_page
    clean["end_page"] = end_page

    return clean


def add_page_range(page_set, start_page, end_page, total_pages, before=0, after=0):
    for page_number in range(max(1, start_page - before), min(total_pages, end_page + after) + 1):
        page_set.add(page_number)


def candidate_list_to_pages(candidates, total_pages, max_pages, default_span=2, max_span=8, before=0, after=0):
    pages = set()
    clean_candidates = []

    for candidate in candidates:
        clean = normalize_candidate(
            candidate=candidate,
            total_pages=total_pages,
            default_span=default_span,
            max_span=max_span
        )

        if clean is None:
            continue

        clean_candidates.append(clean)
        add_page_range(
            page_set=pages,
            start_page=clean["start_page"],
            end_page=clean["end_page"],
            total_pages=total_pages,
            before=before,
            after=after
        )

    if len(pages) > max_pages:
        sorted_pages = sorted(pages)
        pages = set(sorted_pages[:max_pages])

    return sorted(pages), clean_candidates


def find_header_fallback_pages(pages, terms, max_pages):
    scores = []

    for page in pages:
        page_number = page["page_number"]
        text = page["text"]
        score = count_term_matches(text[:3000], terms)

        if score > 0:
            scores.append((score, page_number))

    scores = sorted(scores, reverse=True)
    selected = set()

    for score, page_number in scores:
        add_page_range(selected, page_number, min(page_number + 2, len(pages)), len(pages), before=0, after=0)

        if len(selected) >= max_pages:
            break

    return sorted(selected)


def build_target_pages(pages, bookmarks, toc_text, toc_ai_data, metadata_folder, base_name):
    total_pages = len(pages)

    endpoint_from_text, schedule_from_text, supporting_from_text = extract_structural_candidates_from_text(toc_text, total_pages)
    endpoint_from_bookmarks, schedule_from_bookmarks, supporting_from_bookmarks = extract_structural_candidates_from_bookmarks(bookmarks, total_pages)

    strict_schedule_sections, numbered_toc_entries = get_strict_schedule_sections_from_numbered_toc(
        toc_text,
        total_pages
    )

    all_table_entries = extract_all_table_entries_from_toc_text(toc_text, total_pages)
    spanned_table_entries = infer_table_spans(all_table_entries, pages, max_span=MAX_SCHEDULE_SECTION_SPAN)

    schedule_table_candidates = [
        entry for entry in spanned_table_entries
        if contains_any(entry.get("section_title", ""), SCHEDULE_SECTION_TERMS)
    ]

    endpoint_from_ai = toc_ai_data.get("endpoint_sections", [])
    schedule_from_ai = toc_ai_data.get("schedule_sections", [])
    supporting_from_ai = toc_ai_data.get("supporting_sections", [])

    endpoint_candidates = endpoint_from_text + endpoint_from_ai + endpoint_from_bookmarks

    # Important update:
    # If a strict schedule section is found, keep it, but ALSO include schedule-related
    # List of Tables entries and PDF bookmarks. This avoids missing sampling/data-collection tables.
    if STRICT_TOC_SCHEDULE_SECTION_ONLY and strict_schedule_sections:
        schedule_candidates = (
            strict_schedule_sections
            + schedule_table_candidates
            + schedule_from_bookmarks
        )
        schedule_source_mode = "strict_numbered_toc_plus_list_of_tables_and_bookmarks"
    else:
        schedule_candidates = (
            schedule_table_candidates
            + schedule_from_text
            + schedule_from_ai
            + schedule_from_bookmarks
        )
        schedule_source_mode = "table_or_ai_or_bookmark_with_fallback"

    supporting_candidates = supporting_from_text + supporting_from_ai + supporting_from_bookmarks

    endpoint_pages, clean_endpoint_candidates = candidate_list_to_pages(
        endpoint_candidates,
        total_pages,
        max_pages=MAX_ENDPOINT_PAGES,
        default_span=2,
        max_span=8,
        before=0,
        after=1
    )

    schedule_pages, clean_schedule_candidates = candidate_list_to_pages(
        schedule_candidates,
        total_pages,
        max_pages=MAX_SCHEDULE_PAGES,
        default_span=MAX_SCHEDULE_SECTION_SPAN,
        max_span=MAX_SCHEDULE_SECTION_SPAN,
        before=0,
        after=0
    )

    supporting_pages, clean_supporting_candidates = candidate_list_to_pages(
        supporting_candidates,
        total_pages,
        max_pages=MAX_SUPPORTING_PAGES,
        default_span=3,
        max_span=8,
        before=0,
        after=1
    )

    if not endpoint_pages:
        endpoint_pages = find_header_fallback_pages(pages, ENDPOINT_SECTION_TERMS, MAX_ENDPOINT_PAGES)

    if not schedule_pages:
        schedule_pages = find_header_fallback_pages(pages, SCHEDULE_SECTION_TERMS, MAX_SCHEDULE_PAGES)

    if not supporting_pages:
        supporting_pages = find_header_fallback_pages(pages, SUPPORTING_SECTION_TERMS, MAX_SUPPORTING_PAGES)

    diagnostics = {
        "endpoint_pages": endpoint_pages,
        "schedule_pages": schedule_pages,
        "supporting_pages": supporting_pages,
        "schedule_source_mode": schedule_source_mode,
        "strict_schedule_sections_from_numbered_toc": strict_schedule_sections,
        "numbered_toc_entries": numbered_toc_entries,
        "all_table_entries": all_table_entries,
        "spanned_table_entries": spanned_table_entries,
        "schedule_table_candidates": schedule_table_candidates,
        "endpoint_candidates": clean_endpoint_candidates,
        "schedule_candidates": clean_schedule_candidates,
        "supporting_candidates": clean_supporting_candidates,
        "note": (
            "Schedule pages are selected from strict numbered TOC when possible, "
            "plus schedule-related List of Tables entries and PDF bookmarks."
        )
    }

    diagnostics_path = os.path.join(metadata_folder, f"{base_name}_target_page_diagnostics.json")

    with open(diagnostics_path, "w", encoding="utf-8") as f:
        json.dump(diagnostics, f, indent=4, ensure_ascii=False)

    return endpoint_pages, schedule_pages, supporting_pages, diagnostics_path


# =========================================================
# 8. STEP 1: STUDY MAP EXTRACTION
# =========================================================

STUDY_MAP_PROMPT = """
You are analyzing the objectives/endpoints/outcome-measures section of a clinical study protocol.

Task:
Extract a unified study reference map only.
Do not classify activities.
Do not classify schedule rows.
Objectives and endpoints must be treated at the same logical level as reference concepts for later activity classification.

Return only valid JSON.
Do not use markdown.
Do not invent information.

Use this exact JSON structure:
{
  "reference_concepts": [
    {
      "concept_id": "",
      "concept_type": "",
      "concept_text": "",
      "hierarchy": "",
      "domain": "",
      "linked_objective_or_endpoint": "",
      "reason": ""
    }
  ],
  "study_logic_notes": [
    ""
  ]
}

Rules:
- Each objective and each endpoint should become a reference_concept.
- concept_type must be one of: objective, endpoint.
- hierarchy must be one of: primary, secondary, additional_or_exploratory, unclear.
- domain must be one of: safety, efficacy, other, unclear.
- Use efficacy for clinical benefit, performance, effectiveness, symptoms, survival, response, score improvement, imaging improvement, quality of life if defined as an outcome, pharmacokinetic outcome, pharmacodynamic outcome, biomarker outcome, or similar benefit/outcome concepts.
- Use safety for adverse events, serious adverse events, adverse device effects, toxicity, tolerability, complications, laboratory safety, vital signs, ECG, physical examination, injection site reactions, pregnancy safety, risk, harm, discontinuation due to adverse events, or similar safety concepts.
- Use other only when the endpoint/objective is clearly neither safety nor efficacy/performance.
- Use unclear only when the protocol wording does not allow classification.
- If an endpoint is listed under a primary efficacy objective, classify that endpoint as hierarchy=primary and domain=efficacy.
- If an endpoint is listed under a secondary safety objective, classify that endpoint as hierarchy=secondary and domain=safety.
- If an endpoint is listed under tertiary, exploratory, additional, descriptive, observational, or other non-primary/non-secondary objective, classify hierarchy=additional_or_exploratory.
- Do not put schedule activities, procedures, eligibility criteria, statistics, or GCP operations in reference_concepts unless the protocol explicitly lists them as objectives or endpoints.
- Do not classify endpoints/objectives as activities.
- Use objectives/endpoints only as reference concepts for later activity classification.
"""


def infer_safety_or_efficacy_from_text(text):
    text_norm = normalize_text(text)

    safety_terms = [
        "safety", "tolerability", "adverse event", "serious adverse event",
        "ae", "sae", "toxicity", "complication", "risk", "harm",
        "laboratory", "lab", "vital signs", "blood pressure", "pulse",
        "temperature", "respiratory rate", "ecg", "electrocardiogram",
        "physical examination", "injection site", "immunogenicity",
        "ada", "nab", "anti peg", "pregnancy", "hypersensitivity"
    ]

    efficacy_terms = [
        "efficacy", "effectiveness", "performance", "success", "response",
        "survival", "clinical outcome", "functional outcome", "improvement",
        "reduction", "change from baseline", "quality of life", "exposure",
        "concentration", "pharmacokinetic", "pharmacodynamics", "pharmacodynamic",
        "biomarker", "auc", "cmax", "tmax", "half life", "clearance",
        "volume of distribution", "dose proportionality", "endpoint"
    ]

    if any(normalize_text(term) in text_norm for term in safety_terms):
        return "safety"

    if any(normalize_text(term) in text_norm for term in efficacy_terms):
        return "efficacy"

    return "unknown"


def normalize_hierarchy(value):
    value = normalize_text(value)

    if "primary" in value:
        return "primary"

    if "secondary" in value:
        return "secondary"

    if any(word in value for word in ["additional", "exploratory", "tertiary", "observational", "descriptive", "other"]):
        return "additional_or_exploratory"

    return "unclear"


def normalize_domain(value, fallback_text=""):
    value = normalize_text(value)

    if "safety" in value:
        return "safety"

    if any(word in value for word in ["efficacy", "effectiveness", "performance", "benefit", "clinical outcome"]):
        return "efficacy"

    if "other" in value:
        return "other"

    inferred = infer_safety_or_efficacy_from_text(fallback_text)

    if inferred in ["safety", "efficacy"]:
        return inferred

    return "unclear"


def concept_from_legacy_objective(objective, concept_number):
    objective_text = str(objective.get("objective_text", "")).strip()
    linked = str(objective.get("linked_endpoint", "")).strip()
    hierarchy = normalize_hierarchy(objective.get("endpoint_hierarchy", objective.get("objective_type", "")))
    domain = normalize_domain(objective.get("safety_or_efficacy", ""), objective_text + " " + linked)

    return {
        "concept_id": f"REF_{concept_number:04d}",
        "concept_type": "objective",
        "concept_text": objective_text,
        "hierarchy": hierarchy,
        "domain": domain,
        "linked_objective_or_endpoint": linked,
        "reason": "Converted from legacy study_objectives structure."
    }


def concept_from_legacy_endpoint(endpoint, hierarchy, concept_number):
    if isinstance(endpoint, dict):
        endpoint_text = str(endpoint.get("endpoint_text", endpoint.get("text", endpoint))).strip()
        linked = str(endpoint.get("linked_objective_or_endpoint", endpoint.get("linked_objective", ""))).strip()
        domain_value = endpoint.get("domain", endpoint.get("safety_or_efficacy", ""))
    else:
        endpoint_text = str(endpoint).strip()
        linked = ""
        domain_value = ""

    domain = normalize_domain(domain_value, endpoint_text + " " + linked)

    return {
        "concept_id": f"REF_{concept_number:04d}",
        "concept_type": "endpoint",
        "concept_text": endpoint_text,
        "hierarchy": normalize_hierarchy(hierarchy),
        "domain": domain,
        "linked_objective_or_endpoint": linked,
        "reason": "Converted from legacy endpoints structure."
    }


def normalize_reference_concept(concept, concept_number):
    concept_text = str(concept.get("concept_text", "")).strip()

    if not concept_text:
        concept_text = str(
            concept.get("objective_text", concept.get("endpoint_text", concept.get("text", "")))
        ).strip()

    concept_type = normalize_text(concept.get("concept_type", ""))

    if concept_type not in ["objective", "endpoint"]:
        if "objective" in normalize_text(concept_text):
            concept_type = "objective"
        else:
            concept_type = "endpoint"

    hierarchy = normalize_hierarchy(concept.get("hierarchy", concept.get("endpoint_hierarchy", "")))
    linked = str(concept.get("linked_objective_or_endpoint", concept.get("linked_endpoint", ""))).strip()
    domain = normalize_domain(concept.get("domain", concept.get("safety_or_efficacy", "")), concept_text + " " + linked)

    concept_id = str(concept.get("concept_id", "")).strip()

    if not concept_id:
        concept_id = f"REF_{concept_number:04d}"

    return {
        "concept_id": concept_id,
        "concept_type": concept_type,
        "concept_text": concept_text,
        "hierarchy": hierarchy,
        "domain": domain,
        "linked_objective_or_endpoint": linked,
        "reason": str(concept.get("reason", "")).strip()
    }


def normalize_study_map_structure(study_map):
    normalized = {
        "reference_concepts": [],
        "study_objectives": [],
        "endpoints": {
            "primary": [],
            "secondary": [],
            "additional_or_exploratory": []
        },
        "study_logic_notes": []
    }

    seen = set()
    concept_number = 1

    def add_concept(concept):
        nonlocal concept_number

        normalized_concept = normalize_reference_concept(concept, concept_number)
        concept_text = normalized_concept.get("concept_text", "")

        if not concept_text:
            return

        key = (
            normalize_text(normalized_concept["concept_type"]),
            normalize_text(normalized_concept["concept_text"]),
            normalized_concept["hierarchy"],
            normalized_concept["domain"]
        )

        if key in seen:
            return

        normalized_concept["concept_id"] = f"REF_{concept_number:04d}"
        normalized["reference_concepts"].append(normalized_concept)
        seen.add(key)
        concept_number += 1

    for concept in study_map.get("reference_concepts", []):
        add_concept(concept)

    for objective in study_map.get("study_objectives", []):
        add_concept(concept_from_legacy_objective(objective, concept_number))

    endpoints = study_map.get("endpoints", {})

    for hierarchy in ["primary", "secondary", "additional_or_exploratory"]:
        for endpoint in endpoints.get(hierarchy, []):
            add_concept(concept_from_legacy_endpoint(endpoint, hierarchy, concept_number))

    for concept in normalized["reference_concepts"]:
        if concept["concept_type"] == "objective":
            normalized["study_objectives"].append({
                "objective_text": concept["concept_text"],
                "objective_type": concept["hierarchy"],
                "endpoint_hierarchy": concept["hierarchy"],
                "safety_or_efficacy": concept["domain"],
                "linked_endpoint": concept["linked_objective_or_endpoint"]
            })

        if concept["concept_type"] == "endpoint" and concept["hierarchy"] in normalized["endpoints"]:
            normalized["endpoints"][concept["hierarchy"]].append({
                "endpoint_text": concept["concept_text"],
                "domain": concept["domain"],
                "linked_objective_or_endpoint": concept["linked_objective_or_endpoint"]
            })

    for note in study_map.get("study_logic_notes", []):
        note = str(note).strip()

        if note and note not in normalized["study_logic_notes"]:
            normalized["study_logic_notes"].append(note)

    return normalized


def extract_study_map(endpoint_text, metadata_folder):
    chunks = split_text(
        endpoint_text,
        chunk_size=ENDPOINT_TEXT_CHUNK_SIZE,
        overlap=TEXT_CHUNK_OVERLAP
    )

    partial_maps = []

    for i, chunk in enumerate(chunks, start=1):
        response = call_lm_studio(
            messages=[
                {"role": "system", "content": STUDY_MAP_PROMPT},
                {"role": "user", "content": chunk}
            ],
            temperature=0.0,
            max_tokens=1600
        )

        parsed = parse_ai_json(response, metadata_folder, f"study_map_chunk_{i}")
        partial_maps.append(parsed)

    merged_raw = {
        "reference_concepts": [],
        "study_objectives": [],
        "endpoints": {
            "primary": [],
            "secondary": [],
            "additional_or_exploratory": []
        },
        "study_logic_notes": []
    }

    def add_unique(target, items):
        seen = set(json.dumps(item, sort_keys=True) for item in target)

        for item in items:
            key = json.dumps(item, sort_keys=True)

            if key not in seen:
                target.append(item)
                seen.add(key)

    for partial in partial_maps:
        add_unique(merged_raw["reference_concepts"], partial.get("reference_concepts", []))
        add_unique(merged_raw["study_objectives"], partial.get("study_objectives", []))

        endpoints = partial.get("endpoints", {})
        add_unique(merged_raw["endpoints"]["primary"], endpoints.get("primary", []))
        add_unique(merged_raw["endpoints"]["secondary"], endpoints.get("secondary", []))
        add_unique(merged_raw["endpoints"]["additional_or_exploratory"], endpoints.get("additional_or_exploratory", []))

        for note in partial.get("study_logic_notes", []):
            if note and note not in merged_raw["study_logic_notes"]:
                merged_raw["study_logic_notes"].append(note)

    return normalize_study_map_structure(merged_raw)


# =========================================================
# 9. STEP 2: SCHEDULE ROW EXTRACTION
# =========================================================

SCHEDULE_ROW_EXTRACTION_PROMPT = """
You are reading schedule tables from a clinical study protocol.

Task:
Extract every readable activity/procedure/assessment row from schedule tables only.

Schedule tables may be called:
- Schedule of Assessments
- Schedule of Activities
- Schedule of Events
- Data Collection Schedule
- Assessment Schedule
- Sampling Schedule
- Visit Schedule

Do not classify activities.
Do not extract endpoints/objectives as activities unless they appear as actual schedule row labels.
Do not summarize the table.
Do not skip readable rows.
Return one schedule_rows entry per readable activity/procedure/assessment row.

Return only valid JSON.
Do not use markdown.

Use this exact JSON structure:
{
  "schedule_rows": [
    {
      "activity_name": "",
      "source_table": "",
      "source_page": "",
      "visit_or_timepoint": "",
      "raw_row_text": "",
      "notes": ""
    }
  ]
}
"""


def extract_schedule_rows(schedule_text, metadata_folder):
    chunks = split_text(
        schedule_text,
        chunk_size=SCHEDULE_TEXT_CHUNK_SIZE,
        overlap=TEXT_CHUNK_OVERLAP
    )

    all_rows = []
    seen = set()

    for i, chunk in enumerate(chunks, start=1):
        response = call_lm_studio(
            messages=[
                {"role": "system", "content": SCHEDULE_ROW_EXTRACTION_PROMPT},
                {"role": "user", "content": chunk}
            ],
            temperature=0.0,
            max_tokens=2200
        )

        parsed = parse_ai_json(response, metadata_folder, f"schedule_rows_chunk_{i}")

        for row in parsed.get("schedule_rows", []):
            activity_name = str(row.get("activity_name", "")).strip()
            raw_row_text = str(row.get("raw_row_text", "")).strip()
            source_page = str(row.get("source_page", "")).strip()

            if not activity_name and not raw_row_text:
                continue

            key = normalize_text(activity_name + " " + raw_row_text + " " + source_page)

            if key not in seen:
                row["row_id"] = f"ROW_{len(all_rows) + 1:04d}"
                all_rows.append(row)
                seen.add(key)

    return all_rows


# =========================================================
# 10. STEP 3: CLASSIFY EXTRACTED SCHEDULE ROWS
# =========================================================

CLASSIFY_ROWS_PROMPT = """
You are classifying extracted schedule activities from a clinical study protocol.

You are given:
1. A study map containing primary, secondary, and additional/exploratory endpoints/objectives.
2. Extracted schedule rows.
3. A GCP reference list.

Critical rule:
Do not classify endpoints or objectives as activities.
Use endpoints/objectives only as reference concepts for classifying activities.
Only classify the extracted schedule rows.

For every input schedule row, return exactly one classified item using the same row_id.

Allowed classification categories:
- primary_endpoint_related_safety
- primary_endpoint_related_efficacy
- secondary_endpoint_related_safety
- secondary_endpoint_related_efficacy
- additional_or_exploratory_endpoint_related
- general_study_operations_related_to_gcp
- not_classifiable

General endpoint-to-activity mapping:
Classify an activity as endpoint-related when it directly supports measurement, assessment, confirmation, follow-up, adjudication, source-data collection, calculation, derivation, statistical analysis, reporting, or interpretation of an endpoint/objective.

Use endpoint hierarchy from the study map:
- If the linked endpoint/objective is primary, classify as primary_endpoint_related_safety or primary_endpoint_related_efficacy.
- If the linked endpoint/objective is secondary, classify as secondary_endpoint_related_safety or secondary_endpoint_related_efficacy.
- If the linked endpoint/objective is additional/exploratory/tertiary/descriptive/observational/other, classify as additional_or_exploratory_endpoint_related.

Determine safety vs efficacy from the endpoint/objective wording:
- Safety-related examples: safety, tolerability, adverse event, serious adverse event, complication, device deficiency, toxicity, laboratory safety, vital signs, ECG, physical examination, risk, harm.
- Efficacy/performance-related examples: efficacy, effectiveness, performance, success, response, survival, improvement, reduction, clinical outcome, functional outcome, pharmacokinetic outcome, pharmacodynamic outcome, exposure, biomarker, quality of life when defined as outcome.

GCP/study-operation classification:
If an activity is required for study conduct, compliance, subject protection, data integrity, or operational tracking but is not directly linked to an endpoint/objective, classify as general_study_operations_related_to_gcp.

If the activity cannot be linked to an endpoint/objective and is not clearly GCP/study operation, classify as not_classifiable.

Never leave classification blank.
Never skip a row.
Never invent schedule activities.

Return only valid JSON.
Do not use markdown.

Use this exact JSON structure:
{
  "classified_items": [
    {
      "row_id": "",
      "activity_name": "",
      "source_table": "",
      "source_page": "",
      "visit_or_timepoint": "",
      "classification": "",
      "linked_endpoint_or_requirement": "",
      "reason_for_classification": ""
    }
  ]
}
"""


def normalize_classification(value):
    value = str(value).strip().lower()
    value = value.replace("-", "_")
    value = value.replace(" ", "_")

    mapping = {
        "primary_endpoint_related_to_safety": "primary_endpoint_related_safety",
        "primary_endpoint_related_safety": "primary_endpoint_related_safety",
        "primary_safety": "primary_endpoint_related_safety",

        "primary_endpoint_related_to_efficacy": "primary_endpoint_related_efficacy",
        "primary_endpoint_related_efficacy": "primary_endpoint_related_efficacy",
        "primary_efficacy": "primary_endpoint_related_efficacy",

        "secondary_endpoint_related_to_safety": "secondary_endpoint_related_safety",
        "secondary_endpoint_related_safety": "secondary_endpoint_related_safety",
        "secondary_safety": "secondary_endpoint_related_safety",

        "secondary_endpoint_related_to_efficacy": "secondary_endpoint_related_efficacy",
        "secondary_endpoint_related_efficacy": "secondary_endpoint_related_efficacy",
        "secondary_efficacy": "secondary_endpoint_related_efficacy",

        "additional_or_exploratory_endpoint_related": "additional_or_exploratory_endpoint_related",
        "exploratory_endpoint_related": "additional_or_exploratory_endpoint_related",
        "additional_endpoint_related": "additional_or_exploratory_endpoint_related",

        "general_study_operations": "general_study_operations_related_to_gcp",
        "general_study_operations_related_to_gcp": "general_study_operations_related_to_gcp",

        "not_classifiable": "not_classifiable"
    }

    return mapping.get(value, "not_classifiable")


def extract_match_terms(text):
    text = str(text)
    text = re.sub(r"[\u2022*]", ",", text)

    chunks = re.split(r"[,;:\n]|\band\b|\bor\b|\(|\)", text, flags=re.IGNORECASE)
    terms = []

    for chunk in chunks:
        cleaned = chunk.strip(" .:-")
        cleaned_norm = normalize_text(cleaned)

        if len(cleaned_norm) < 4:
            continue

        if cleaned_norm in [
            "endpoint", "endpoints", "objective", "objectives", "variables",
            "assessment", "assessments", "profile", "properties", "parameters"
        ]:
            continue

        if len(cleaned_norm.split()) > 12:
            continue

        terms.append(cleaned)

    full_norm = normalize_text(text)

    if 1 <= len(full_norm.split()) <= 12:
        terms.append(text.strip())

    unique = []
    seen = set()

    for term in terms:
        key = normalize_text(term)

        if key and key not in seen:
            unique.append(term)
            seen.add(key)

    return unique


def build_endpoint_reference_index(study_map):
    references = []
    normalized_map = normalize_study_map_structure(study_map)

    for concept in normalized_map.get("reference_concepts", []):
        concept_text = str(concept.get("concept_text", "")).strip()
        linked = str(concept.get("linked_objective_or_endpoint", "")).strip()

        if not concept_text:
            continue

        references.append({
            "concept_id": concept.get("concept_id", ""),
            "concept_type": concept.get("concept_type", ""),
            "hierarchy": concept.get("hierarchy", "unclear"),
            "nature": concept.get("domain", "unclear"),
            "text": concept_text,
            "linked_objective_or_endpoint": linked,
            "terms": extract_match_terms(concept_text + " " + linked)
        })

    return references


def term_matches_activity(term, activity_text):
    term_norm = normalize_text(term)
    activity_norm = normalize_text(activity_text)

    if not term_norm or not activity_norm:
        return False

    if len(term_norm) < 4:
        return False

    if term_norm in activity_norm:
        return True

    if activity_norm in term_norm and len(activity_norm.split()) >= 2:
        return True

    term_tokens = [t for t in term_norm.split() if len(t) >= 4]
    activity_tokens = set(t for t in activity_norm.split() if len(t) >= 4)

    if not term_tokens:
        return False

    overlap = sum(1 for token in term_tokens if token in activity_tokens)

    if len(term_tokens) <= 2:
        return overlap == len(term_tokens)

    return overlap >= max(2, int(len(term_tokens) * 0.6))


def category_from_endpoint_reference(reference):
    hierarchy = reference.get("hierarchy", "")
    nature = reference.get("nature", "unknown")

    if hierarchy == "additional_or_exploratory":
        return "additional_or_exploratory_endpoint_related"

    if hierarchy in ["primary", "secondary"] and nature in ["safety", "efficacy"]:
        return f"{hierarchy}_endpoint_related_{nature}"

    return "not_classifiable"


def deterministic_activity_classification(activity_text, study_map):
    references = build_endpoint_reference_index(study_map)

    for reference in references:
        category = category_from_endpoint_reference(reference)

        if category == "not_classifiable":
            continue

        for term in reference.get("terms", []):
            if term_matches_activity(term, activity_text):
                return {
                    "classification": category,
                    "linked_endpoint_or_requirement": reference.get("text", ""),
                    "reason_for_classification": "Matched activity to endpoint/objective reference map."
                }

    activity_norm = normalize_text(activity_text)

    for gcp_category, activities in GCP_REFERENCE_ACTIVITIES.items():
        for activity in activities:
            if normalize_text(activity) in activity_norm or activity_norm in normalize_text(activity):
                return {
                    "classification": "general_study_operations_related_to_gcp",
                    "linked_endpoint_or_requirement": f"General GCP / study operation: {gcp_category}",
                    "reason_for_classification": "Matched activity to GCP/study-operation reference list."
                }

    return None


def apply_deterministic_classification_review(classified_items, study_map):
    reviewed = []

    for item in classified_items:
        activity_text = " ".join([
            str(item.get("activity_name", "")),
            str(item.get("source_table", "")),
            str(item.get("visit_or_timepoint", "")),
            str(item.get("linked_endpoint_or_requirement", "")),
            str(item.get("reason_for_classification", ""))
        ])

        deterministic = deterministic_activity_classification(activity_text, study_map)

        if deterministic:
            current = normalize_classification(item.get("classification", "not_classifiable"))

            if current in ["not_classifiable", "general_study_operations_related_to_gcp"]:
                item["classification"] = deterministic["classification"]
                item["linked_endpoint_or_requirement"] = deterministic["linked_endpoint_or_requirement"]
                item["reason_for_classification"] = deterministic["reason_for_classification"]

        item["classification"] = normalize_classification(item.get("classification", "not_classifiable"))
        reviewed.append(item)

    return reviewed


def build_endpoint_presence_summary(study_map):
    references = build_endpoint_reference_index(study_map)

    has_safety = any(ref.get("nature") == "safety" for ref in references)
    has_efficacy = any(ref.get("nature") == "efficacy" for ref in references)
    has_other = any(ref.get("nature") == "other" for ref in references)
    has_unclear = any(ref.get("nature") == "unclear" for ref in references)
    has_primary = any(ref.get("hierarchy") == "primary" for ref in references)
    has_secondary = any(ref.get("hierarchy") == "secondary" for ref in references)
    has_additional = any(ref.get("hierarchy") == "additional_or_exploratory" for ref in references)

    notes = []

    if not has_safety:
        notes.append("No safety endpoint/objective was clearly identified in the extracted study map.")

    if not has_efficacy:
        notes.append("No efficacy/performance endpoint/objective was clearly identified in the extracted study map.")

    if has_other:
        notes.append("At least one reference concept was identified as other because it was neither clearly safety nor efficacy/performance.")

    if has_unclear:
        notes.append("At least one reference concept had unclear safety/efficacy domain based on the extracted text.")

    if not has_primary:
        notes.append("No primary endpoint/objective was clearly identified in the extracted study map.")

    if not has_secondary:
        notes.append("No secondary endpoint/objective was clearly identified in the extracted study map.")

    if not has_additional:
        notes.append("No additional/exploratory endpoint/objective was clearly identified in the extracted study map.")

    if not notes:
        notes.append("Safety/efficacy and endpoint hierarchy concepts were identified in the extracted study map.")

    return {
        "has_safety_endpoint_or_objective": has_safety,
        "has_efficacy_or_performance_endpoint_or_objective": has_efficacy,
        "has_other_endpoint_or_objective": has_other,
        "has_unclear_domain_endpoint_or_objective": has_unclear,
        "has_primary_endpoint_or_objective": has_primary,
        "has_secondary_endpoint_or_objective": has_secondary,
        "has_additional_or_exploratory_endpoint_or_objective": has_additional,
        "summary_notes": notes
    }


def classify_schedule_rows(schedule_rows, study_map, metadata_folder):
    classified = []
    gcp_reference_text = build_gcp_reference_text()

    for start in range(0, len(schedule_rows), ROWS_PER_CLASSIFICATION_BATCH):
        batch = schedule_rows[start:start + ROWS_PER_CLASSIFICATION_BATCH]
        batch_number = (start // ROWS_PER_CLASSIFICATION_BATCH) + 1

        user_payload = {
            "study_map": study_map,
            "gcp_reference_list": gcp_reference_text,
            "schedule_rows_to_classify": batch
        }

        response = call_lm_studio(
            messages=[
                {"role": "system", "content": CLASSIFY_ROWS_PROMPT},
                {"role": "user", "content": json.dumps(user_payload, indent=2, ensure_ascii=False)}
            ],
            temperature=0.0,
            max_tokens=1600
        )

        parsed = parse_ai_json(response, metadata_folder, f"classification_batch_{batch_number}")
        batch_items = parsed.get("classified_items", [])

        batch_by_id = {
            str(item.get("row_id", "")).strip(): item
            for item in batch_items
        }

        for row in batch:
            row_id = str(row.get("row_id", "")).strip()
            item = batch_by_id.get(row_id)

            if not item:
                item = {
                    "row_id": row_id,
                    "activity_name": row.get("activity_name", ""),
                    "source_table": row.get("source_table", ""),
                    "source_page": row.get("source_page", ""),
                    "visit_or_timepoint": row.get("visit_or_timepoint", ""),
                    "classification": "not_classifiable",
                    "linked_endpoint_or_requirement": "Not classifiable",
                    "reason_for_classification": "AI did not return classification for this row."
                }

            item["classification"] = normalize_classification(item.get("classification", ""))

            if item["classification"] not in VALID_CATEGORIES:
                item["classification"] = "not_classifiable"

            if not item.get("activity_name"):
                item["activity_name"] = row.get("activity_name", "")

            if not item.get("source_table"):
                item["source_table"] = row.get("source_table", "")

            if not item.get("source_page"):
                item["source_page"] = row.get("source_page", "")

            if not item.get("visit_or_timepoint"):
                item["visit_or_timepoint"] = row.get("visit_or_timepoint", "")

            if not item.get("linked_endpoint_or_requirement"):
                if item["classification"] == "general_study_operations_related_to_gcp":
                    item["linked_endpoint_or_requirement"] = "General GCP / study operation"
                elif item["classification"] == "not_classifiable":
                    item["linked_endpoint_or_requirement"] = "Not classifiable"
                else:
                    item["linked_endpoint_or_requirement"] = "Endpoint/objective from study map"

            if not item.get("reason_for_classification"):
                item["reason_for_classification"] = "Classified using study map and GCP reference."

            classified.append(item)

    classified = apply_deterministic_classification_review(classified, study_map)
    return classified


# =========================================================
# 11. METRICS AND QC
# =========================================================

def calculate_metrics(classified_items):
    counts = Counter()

    for item in classified_items:
        classification = normalize_classification(item.get("classification", "not_classifiable"))
        item["classification"] = classification
        counts[classification] += 1

    total = len(classified_items)
    percentages = {}

    for category in VALID_CATEGORIES:
        percentages[category] = round((counts[category] / total) * 100, 2) if total else 0

    return {
        "total_classified_activities": total,
        "counts": {
            category: counts[category]
            for category in VALID_CATEGORIES
        },
        "percentages": percentages
    }


def build_qc_checks(study_map, schedule_rows, classified_items, endpoint_pages, schedule_pages):
    classified_row_ids = set(str(item.get("row_id", "")) for item in classified_items)
    input_row_ids = set(str(row.get("row_id", "")) for row in schedule_rows)

    missing_rows = sorted(input_row_ids - classified_row_ids)
    extra_rows = sorted(classified_row_ids - input_row_ids)

    endpoints = study_map.get("endpoints", {})
    endpoint_summary = build_endpoint_presence_summary(study_map)

    checks = [
        {
            "check": "endpoint_reference_map_present",
            "status": "PASS" if (
                endpoints.get("primary") or endpoints.get("secondary") or endpoints.get("additional_or_exploratory")
            ) else "WARNING",
            "details": "Primary/secondary/additional endpoint reference concepts were extracted."
        },
        {
            "check": "safety_endpoint_or_objective_detected",
            "status": "PASS" if endpoint_summary["has_safety_endpoint_or_objective"] else "INFO",
            "details": "Safety endpoint/objective detected." if endpoint_summary["has_safety_endpoint_or_objective"] else "No safety endpoint/objective clearly detected in study map."
        },
        {
            "check": "efficacy_or_performance_endpoint_or_objective_detected",
            "status": "PASS" if endpoint_summary["has_efficacy_or_performance_endpoint_or_objective"] else "INFO",
            "details": "Efficacy/performance endpoint/objective detected." if endpoint_summary["has_efficacy_or_performance_endpoint_or_objective"] else "No efficacy/performance endpoint/objective clearly detected in study map."
        },
        {
            "check": "schedule_rows_extracted",
            "status": "PASS" if len(schedule_rows) > 0 else "WARNING",
            "details": f"{len(schedule_rows)} schedule/activity rows extracted."
        },
        {
            "check": "one_classification_per_schedule_row",
            "status": "PASS" if not missing_rows else "WARNING",
            "details": f"Missing classified row IDs: {missing_rows}"
        },
        {
            "check": "no_extra_classification_rows",
            "status": "PASS" if not extra_rows else "WARNING",
            "details": f"Extra classified row IDs: {extra_rows}"
        },
        {
            "check": "endpoint_pages_selected",
            "status": "PASS" if endpoint_pages else "WARNING",
            "details": f"Endpoint pages selected: {endpoint_pages}"
        },
        {
            "check": "schedule_pages_selected",
            "status": "PASS" if schedule_pages else "WARNING",
            "details": f"Schedule pages selected: {schedule_pages}"
        }
    ]

    return checks


# =========================================================
# 12. EXCEL OUTPUT
# =========================================================

def write_excel_output(output_excel_path, analysis_data):
    workbook = xlsxwriter.Workbook(output_excel_path)

    title_format = workbook.add_format({
        "bold": True,
        "font_size": 16,
        "bg_color": "#D9EAF7",
        "border": 1
    })

    header_format = workbook.add_format({
        "bold": True,
        "bg_color": "#1F4E78",
        "font_color": "white",
        "border": 1
    })

    cell_format = workbook.add_format({
        "text_wrap": True,
        "valign": "top",
        "border": 1
    })

    percent_format = workbook.add_format({
        "num_format": "0.00%",
        "border": 1
    })

    metrics = analysis_data.get("metrics", {})

    ws = workbook.add_worksheet("Summary")
    ws.set_column("A:A", 45)
    ws.set_column("B:B", 18)
    ws.set_column("C:C", 18)

    ws.write("A1", "Schedule Activity Classification Summary", title_format)
    ws.write("A3", "Classification Category", header_format)
    ws.write("B3", "Count", header_format)
    ws.write("C3", "Percentage", header_format)

    row = 3
    counts = metrics.get("counts", {})
    percentages = metrics.get("percentages", {})

    for category in VALID_CATEGORIES:
        ws.write(row, 0, category, cell_format)
        ws.write(row, 1, counts.get(category, 0), cell_format)
        ws.write(row, 2, percentages.get(category, 0) / 100, percent_format)
        row += 1

    chart = workbook.add_chart({"type": "pie"})
    chart.add_series({
        "name": "Schedule Activity Classification",
        "categories": ["Summary", 3, 0, row - 1, 0],
        "values": ["Summary", 3, 1, row - 1, 1],
        "data_labels": {"percentage": True, "category": True}
    })
    chart.set_title({"name": "Schedule Activity Classification"})
    chart.set_style(10)
    ws.insert_chart("E3", chart, {"x_scale": 1.4, "y_scale": 1.4})

    endpoint_summary = analysis_data.get("endpoint_presence_summary", {})

    ws.write("A13", "Endpoint Map Status", title_format)
    ws.write("A15", "Safety endpoint/objective detected", header_format)
    ws.write("B15", str(endpoint_summary.get("has_safety_endpoint_or_objective", "")), cell_format)
    ws.write("A16", "Efficacy/performance endpoint/objective detected", header_format)
    ws.write("B16", str(endpoint_summary.get("has_efficacy_or_performance_endpoint_or_objective", "")), cell_format)
    ws.write("A17", "Primary endpoint/objective detected", header_format)
    ws.write("B17", str(endpoint_summary.get("has_primary_endpoint_or_objective", "")), cell_format)
    ws.write("A18", "Secondary endpoint/objective detected", header_format)
    ws.write("B18", str(endpoint_summary.get("has_secondary_endpoint_or_objective", "")), cell_format)
    ws.write("A19", "Additional/exploratory endpoint/objective detected", header_format)
    ws.write("B19", str(endpoint_summary.get("has_additional_or_exploratory_endpoint_or_objective", "")), cell_format)
    ws.write("A20", "Other endpoint/objective domain detected", header_format)
    ws.write("B20", str(endpoint_summary.get("has_other_endpoint_or_objective", "")), cell_format)
    ws.write("A21", "Unclear endpoint/objective domain detected", header_format)
    ws.write("B21", str(endpoint_summary.get("has_unclear_domain_endpoint_or_objective", "")), cell_format)
    ws.write("A23", "Endpoint Map Notes", header_format)
    notes = endpoint_summary.get("summary_notes", [])
    ws.write("B23", " | ".join(str(note) for note in notes), cell_format)

    ws = workbook.add_worksheet("Study Map")
    ws.set_column("A:A", 14)
    ws.set_column("B:B", 18)
    ws.set_column("C:C", 100)
    ws.set_column("D:D", 24)
    ws.set_column("E:E", 20)
    ws.set_column("F:F", 100)
    ws.set_column("G:G", 80)

    headers = [
        "Concept ID",
        "Concept Type",
        "Reference Concept Text",
        "Hierarchy",
        "Safety/Efficacy Domain",
        "Linked Objective or Endpoint",
        "Reason"
    ]

    for col, header in enumerate(headers):
        ws.write(0, col, header, header_format)

    study_map = normalize_study_map_structure(analysis_data.get("study_map", {}))

    for row_idx, concept in enumerate(study_map.get("reference_concepts", []), start=1):
        ws.write(row_idx, 0, concept.get("concept_id", ""), cell_format)
        ws.write(row_idx, 1, concept.get("concept_type", ""), cell_format)
        ws.write(row_idx, 2, concept.get("concept_text", ""), cell_format)
        ws.write(row_idx, 3, concept.get("hierarchy", ""), cell_format)
        ws.write(row_idx, 4, concept.get("domain", ""), cell_format)
        ws.write(row_idx, 5, concept.get("linked_objective_or_endpoint", ""), cell_format)
        ws.write(row_idx, 6, concept.get("reason", ""), cell_format)

    ws = workbook.add_worksheet("Extracted Schedule Rows")
    ws.set_column("A:A", 14)
    ws.set_column("B:B", 45)
    ws.set_column("C:C", 45)
    ws.set_column("D:D", 18)
    ws.set_column("E:E", 45)
    ws.set_column("F:F", 100)

    headers = ["Row ID", "Activity Name", "Source Table", "Source Page", "Visit / Timepoint", "Raw Row Text"]
    for col, header in enumerate(headers):
        ws.write(0, col, header, header_format)

    for row_idx, item in enumerate(analysis_data.get("schedule_rows", []), start=1):
        ws.write(row_idx, 0, item.get("row_id", ""), cell_format)
        ws.write(row_idx, 1, item.get("activity_name", ""), cell_format)
        ws.write(row_idx, 2, item.get("source_table", ""), cell_format)
        ws.write(row_idx, 3, item.get("source_page", ""), cell_format)
        ws.write(row_idx, 4, item.get("visit_or_timepoint", ""), cell_format)
        ws.write(row_idx, 5, item.get("raw_row_text", ""), cell_format)

    ws = workbook.add_worksheet("Classified Activities")
    ws.set_column("A:A", 14)
    ws.set_column("B:B", 45)
    ws.set_column("C:C", 35)
    ws.set_column("D:D", 18)
    ws.set_column("E:E", 35)
    ws.set_column("F:F", 38)
    ws.set_column("G:G", 70)
    ws.set_column("H:H", 80)

    headers = [
        "Row ID",
        "Activity Name",
        "Source Table",
        "Source Page",
        "Visit / Timepoint",
        "Classification",
        "Linked Endpoint or Requirement",
        "Reason"
    ]
    for col, header in enumerate(headers):
        ws.write(0, col, header, header_format)

    for row_idx, item in enumerate(analysis_data.get("classified_items", []), start=1):
        ws.write(row_idx, 0, item.get("row_id", ""), cell_format)
        ws.write(row_idx, 1, item.get("activity_name", ""), cell_format)
        ws.write(row_idx, 2, item.get("source_table", ""), cell_format)
        ws.write(row_idx, 3, item.get("source_page", ""), cell_format)
        ws.write(row_idx, 4, item.get("visit_or_timepoint", ""), cell_format)
        ws.write(row_idx, 5, item.get("classification", ""), cell_format)
        ws.write(row_idx, 6, item.get("linked_endpoint_or_requirement", ""), cell_format)
        ws.write(row_idx, 7, item.get("reason_for_classification", ""), cell_format)

    ws = workbook.add_worksheet("QC Checks")
    ws.set_column("A:A", 42)
    ws.set_column("B:B", 20)
    ws.set_column("C:C", 120)

    headers = ["Check", "Status", "Details"]
    for col, header in enumerate(headers):
        ws.write(0, col, header, header_format)

    for row_idx, check in enumerate(analysis_data.get("qc_checks", []), start=1):
        ws.write(row_idx, 0, check.get("check", ""), cell_format)
        ws.write(row_idx, 1, check.get("status", ""), cell_format)
        ws.write(row_idx, 2, check.get("details", ""), cell_format)

    ws = workbook.add_worksheet("Metadata")
    ws.set_column("A:A", 38)
    ws.set_column("B:B", 120)

    ws.write(0, 0, "Field", header_format)
    ws.write(0, 1, "Value", header_format)

    for row_idx, (key, value) in enumerate(analysis_data.get("metadata", {}).items(), start=1):
        ws.write(row_idx, 0, key, cell_format)
        ws.write(row_idx, 1, str(value), cell_format)

    workbook.close()


# =========================================================
# 13. MAIN PIPELINE
# =========================================================

def run_full_analysis(pdf_path, output_root, update_callback):
    base_name = safe_folder_name(pdf_path)
    archive_root = archive_previous_run_folders(output_root, base_name, update_callback)
    paths = create_output_paths(pdf_path, output_root)

    base_name = paths["base_name"]
    run_folder = paths["run_folder"]
    metadata_folder = paths["metadata_folder"]
    output_json_path = paths["output_json_path"]
    output_excel_path = paths["output_excel_path"]

    os.makedirs(run_folder, exist_ok=True)
    os.makedirs(metadata_folder, exist_ok=True)

    update_callback("Reading PDF...")
    pages = read_pdf_pages(pdf_path)

    update_callback("Reading PDF bookmarks...")
    bookmarks = extract_pdf_bookmarks(pdf_path)

    bookmarks_path = os.path.join(metadata_folder, f"{base_name}_bookmarks.json")
    with open(bookmarks_path, "w", encoding="utf-8") as f:
        json.dump(bookmarks, f, indent=4, ensure_ascii=False)

    update_callback("Extracting TOC/List of Tables text...")
    toc_text = extract_initial_toc_text(pages, max_pages=10)

    toc_text_path = os.path.join(metadata_folder, f"{base_name}_toc_text.txt")
    with open(toc_text_path, "w", encoding="utf-8") as f:
        f.write(toc_text)

    update_callback("Finding endpoint and schedule pages from TOC/bookmarks...")

    # Important update:
    # Skip AI TOC parsing for speed and stability.
    # Deterministic TOC parsing, List of Tables extraction, and PDF bookmarks are used instead.
    toc_ai_data = {
        "endpoint_sections": [],
        "schedule_sections": [],
        "supporting_sections": []
    }

    endpoint_pages, schedule_pages, supporting_pages, targeting_diagnostics_path = build_target_pages(
        pages=pages,
        bookmarks=bookmarks,
        toc_text=toc_text,
        toc_ai_data=toc_ai_data,
        metadata_folder=metadata_folder,
        base_name=base_name
    )

    endpoint_text = pages_to_text(pages, endpoint_pages)
    schedule_text = pages_to_text(pages, schedule_pages)
    supporting_text = pages_to_text(pages, supporting_pages)

    endpoint_text_path = os.path.join(metadata_folder, f"{base_name}_endpoint_text.txt")
    schedule_text_path = os.path.join(metadata_folder, f"{base_name}_schedule_text.txt")
    supporting_text_path = os.path.join(metadata_folder, f"{base_name}_supporting_text.txt")

    with open(endpoint_text_path, "w", encoding="utf-8") as f:
        f.write(endpoint_text)

    with open(schedule_text_path, "w", encoding="utf-8") as f:
        f.write(schedule_text)

    with open(supporting_text_path, "w", encoding="utf-8") as f:
        f.write(supporting_text)

    if not endpoint_text.strip():
        raise RuntimeError("No endpoint/objective text was extracted. Check metadata targeting diagnostics.")

    if not schedule_text.strip():
        raise RuntimeError("No schedule text was extracted. Check metadata targeting diagnostics.")

    update_callback("Step 1: extracting endpoint/objective reference map. This can take several minutes...")
    study_map = extract_study_map(endpoint_text, metadata_folder)

    study_map_path = os.path.join(metadata_folder, f"{base_name}_study_map.json")
    with open(study_map_path, "w", encoding="utf-8") as f:
        json.dump(study_map, f, indent=4, ensure_ascii=False)

    update_callback("Step 2: extracting schedule/activity rows. This can take several minutes...")
    schedule_rows = extract_schedule_rows(schedule_text, metadata_folder)

    schedule_rows_path = os.path.join(metadata_folder, f"{base_name}_schedule_rows.json")
    with open(schedule_rows_path, "w", encoding="utf-8") as f:
        json.dump(schedule_rows, f, indent=4, ensure_ascii=False)

    update_callback("Step 3: classifying schedule/activity rows. This can take several minutes...")
    classified_items = classify_schedule_rows(schedule_rows, study_map, metadata_folder)

    metrics = calculate_metrics(classified_items)
    endpoint_presence_summary = build_endpoint_presence_summary(study_map)

    qc_checks = build_qc_checks(
        study_map=study_map,
        schedule_rows=schedule_rows,
        classified_items=classified_items,
        endpoint_pages=endpoint_pages,
        schedule_pages=schedule_pages
    )

    metadata = {
        "pdf_path": pdf_path,
        "pdf_file_name": os.path.basename(pdf_path),
        "model_name": MODEL_NAME,
        "lm_studio_url": LM_STUDIO_URL,
        "analysis_approach": "toc_to_endpoint_map_and_schedule_activity_classification",
        "run_folder": run_folder,
        "metadata_folder": metadata_folder,
        "archive_root": archive_root,
        "total_pdf_pages": len(pages),
        "endpoint_pages": endpoint_pages,
        "schedule_pages": schedule_pages,
        "supporting_pages": supporting_pages,
        "endpoint_page_count": len(endpoint_pages),
        "schedule_page_count": len(schedule_pages),
        "supporting_page_count": len(supporting_pages),
        "schedule_rows_extracted": len(schedule_rows),
        "classified_activities": len(classified_items),
        "endpoint_presence_summary": endpoint_presence_summary,
        "toc_text_file": toc_text_path,
        "bookmarks_file": bookmarks_path,
        "targeting_diagnostics_file": targeting_diagnostics_path,
        "endpoint_text_file": endpoint_text_path,
        "schedule_text_file": schedule_text_path,
        "supporting_text_file": supporting_text_path,
        "study_map_file": study_map_path,
        "schedule_rows_file": schedule_rows_path,
        "output_json_file": output_json_path,
        "output_excel_file": output_excel_path
    }

    analysis_data = {
        "study_map": study_map,
        "schedule_rows": schedule_rows,
        "classified_items": classified_items,
        "metrics": metrics,
        "endpoint_presence_summary": endpoint_presence_summary,
        "qc_checks": qc_checks,
        "metadata": metadata
    }

    update_callback("Saving JSON output...")
    with open(output_json_path, "w", encoding="utf-8") as f:
        json.dump(analysis_data, f, indent=4, ensure_ascii=False)

    update_callback("Creating Excel output...")
    write_excel_output(output_excel_path, analysis_data)

    update_callback("Analysis complete.")
    return output_excel_path, output_json_path, metadata_folder, run_folder


# =========================================================
# 14. TKINTER GUI
# =========================================================

selected_pdf_path = ""
selected_output_folder = ""


def gui_update_status(message):
    root.after(0, lambda: status_label.config(text=message))


def gui_set_buttons_state(state):
    root.after(0, lambda: browse_pdf_button.config(state=state))
    root.after(0, lambda: browse_output_button.config(state=state))
    root.after(0, lambda: run_button.config(state=state))


def browse_pdf():
    global selected_pdf_path

    file_path = filedialog.askopenfilename(
        title="Select Clinical Study Protocol PDF",
        filetypes=[("PDF files", "*.pdf")]
    )

    if file_path:
        selected_pdf_path = file_path
        pdf_label.config(text=f"Selected PDF:\n{selected_pdf_path}")


def browse_output_folder():
    global selected_output_folder

    folder_path = filedialog.askdirectory(
        title="Select Output Folder"
    )

    if folder_path:
        selected_output_folder = folder_path
        output_label.config(text=f"Selected output folder:\n{selected_output_folder}")


def run_analysis_button_click():
    if not selected_pdf_path:
        messagebox.showwarning(
            "No PDF selected",
            "Please browse and select a protocol PDF first."
        )
        return

    if not selected_output_folder:
        messagebox.showwarning(
            "No output folder selected",
            "Please browse and select an output folder first."
        )
        return

    gui_set_buttons_state("disabled")
    gui_update_status("Starting analysis...")

    def worker():
        try:
            excel_path, json_path, metadata_folder, run_folder = run_full_analysis(
                pdf_path=selected_pdf_path,
                output_root=selected_output_folder,
                update_callback=gui_update_status
            )

            root.after(
                0,
                lambda: messagebox.showinfo(
                    "Analysis complete",
                    f"Analysis completed successfully.\n\n"
                    f"Run folder:\n{run_folder}\n\n"
                    f"Excel file:\n{excel_path}\n\n"
                    f"JSON file:\n{json_path}\n\n"
                    f"Metadata folder:\n{metadata_folder}"
                )
            )

        except Exception as e:
            error_message = str(e)

            root.after(
                0,
                lambda error_message=error_message: messagebox.showerror(
                    "Error",
                    error_message
                )
            )

        finally:
            gui_set_buttons_state("normal")
            gui_update_status("Ready.")

    threading.Thread(target=worker, daemon=True).start()


root = tk.Tk()
root.title("Clinical Protocol Schedule Activity Classifier")
root.geometry("880x500")

title_label = tk.Label(
    root,
    text="Clinical Protocol Schedule Activity Classifier",
    font=("Arial", 18, "bold")
)
title_label.pack(pady=15)

instruction_label = tk.Label(
    root,
    text=(
        "Workflow: TOC -> endpoint/objective reference map -> schedule/activity row extraction -> "
        "classification of each activity against endpoints or GCP. "
    ),
    font=("Arial", 11),
    wraplength=820
)
instruction_label.pack(pady=5)

browse_pdf_button = tk.Button(
    root,
    text="Browse Protocol PDF",
    font=("Arial", 12),
    width=30,
    command=browse_pdf
)
browse_pdf_button.pack(pady=8)

pdf_label = tk.Label(
    root,
    text="No PDF selected.",
    font=("Arial", 10),
    wraplength=820
)
pdf_label.pack(pady=5)

browse_output_button = tk.Button(
    root,
    text="Browse Output Folder",
    font=("Arial", 12),
    width=30,
    command=browse_output_folder
)
browse_output_button.pack(pady=8)

output_label = tk.Label(
    root,
    text="No output folder selected.",
    font=("Arial", 10),
    wraplength=820
)
output_label.pack(pady=5)

run_button = tk.Button(
    root,
    text="Run / OK",
    font=("Arial", 12, "bold"),
    width=30,
    command=run_analysis_button_click
)
run_button.pack(pady=12)

status_label = tk.Label(
    root,
    text="Ready.",
    font=("Arial", 10, "italic"),
    fg="blue"
)
status_label.pack(pady=10)

root.mainloop()
