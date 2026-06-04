# Clinical Protocol Schedule Activity Classifier

A local Python tool that analyzes clinical study protocol PDFs and classifies activities from the Schedule of Assessments / Schedule of Activities according to their relationship to study endpoints, objectives, or general GCP study operations.

The tool uses a local LLM through LM Studio, extracts protocol sections from the Table of Contents and PDF bookmarks, identifies endpoint/objective reference concepts, extracts schedule rows, classifies each activity, and generates Excel/JSON outputs with QC metadata.

---

## Purpose


Clinical study protocols are large, complex, and highly detailed documents. They contain many activities, assessments, procedures, endpoints, operational requirements, and regulatory obligations. Because of this complexity, it can be difficult for study teams to quickly understand which protocol activities are truly critical to the study objectives and which activities are mainly operational or GCP-related.

The purpose of this analysis is to help study teams identify and classify the activities in the Schedule of Assessments / Schedule of Activities according to their relationship to the study’s primary, secondary, and additional or exploratory objectives and endpoints, as well as general GCP study operations.

This helps teams better understand what is critical in the study and focus attention on the activities that directly support endpoint collection, safety evaluation, efficacy assessment, and study execution.

In the early stages of protocol development, this analysis can support protocol complexity assessment by highlighting how many activities are endpoint-critical versus operational. This can help teams discuss whether some activities are necessary, whether the schedule can be simplified, and whether protocol burden can be reduced before study execution begins.

In later stages of clinical trial planning and execution, the analysis can help study teams focus on what matters most, prioritize critical activities, support oversight discussions, and potentially reduce unnecessary operational complexity and trial running costs.


This tool helps classify those schedule activities into categories such as:

* Primary endpoint-related safety
* Primary endpoint-related efficacy
* Secondary endpoint-related safety
* Secondary endpoint-related efficacy
* Additional or exploratory endpoint-related
* General study operations related to GCP
* Not classifiable

The tool treats **objectives and endpoints as reference concepts only**. It does **not** classify objectives or endpoints as activities. Only schedule rows are classified as activities.

---

## Workflow

The analysis follows this workflow:

```text
Protocol PDF
   ↓
Read Table of Contents, bookmarks, and selected pages
   ↓
Identify endpoint/objective pages and schedule pages
   ↓
Extract endpoint/objective reference map
   ↓
Extract Schedule of Assessment / Schedule of Activities rows
   ↓
Classify each schedule activity against endpoint/objective references or GCP operations
   ↓
Generate Excel, JSON, metadata, and QC outputs
```

---

## Key Features

* Local execution using LM Studio
* No cloud API required if using local models
* Tkinter graphical user interface
* User-selected protocol PDF
* User-selected output folder
* Automatic output run folder creation
* Automatic archiving of previous run folders
* Table-aware PDF extraction using PyMuPDF
* TOC/bookmark-based section targeting
* Schedule activity extraction
* Endpoint/objective reference map extraction
* Activity classification against endpoint/objective concepts and GCP references
* Excel output with summary, study map, extracted rows, classified activities, QC checks, and metadata
* JSON output for traceability
* Metadata files for troubleshooting and auditability

---

## Requirements

### Software

* Windows recommended
* Python 3.10 or later
* LM Studio installed and running
* A local model loaded in LM Studio
* LM Studio local server enabled

The local server should be reachable at:

```text
http://127.0.0.1:1234
```

---

## Python Packages

Install the required packages:

```bash
pip install pymupdf requests json-repair XlsxWriter
```

Optional:

```bash
pip install pymupdf_layout
```

`pymupdf_layout` may improve extraction quality for complex PDF layouts, but it is not required.

---

## Recommended Local Models

The model name in the Python code must match the model ID shown in LM Studio.

Example:

```python
MODEL_NAME = "mrademacher/MedGemma-4B-Instruct-ft-2-GGUF"
```

Recommended model is MedGemma-4B-Instruct-ft-2-GGUF



For clinical protocol terminology, MedGemma may better understand terms such as TEAE, SAE, tolerability, efficacy, pharmacokinetics, pharmacodynamics, safety assessments, and endpoints. However, Qwen models may sometimes produce cleaner JSON.

---

## Configuration

Main settings are near the top of the script:

```python
LM_STUDIO_URL = "http://127.0.0.1:1234/v1/chat/completions"
LM_STUDIO_API_KEY = "lm-studio"
MODEL_NAME = "mrademacher/MedGemma-4B-Instruct-ft-2-GGUF"

ENDPOINT_TEXT_CHUNK_SIZE = 3500
SCHEDULE_TEXT_CHUNK_SIZE = 3500
TEXT_CHUNK_OVERLAP = 100
ROWS_PER_CLASSIFICATION_BATCH = 5
```

### Important Settings

| Setting                         | Meaning                                                            |
| ------------------------------- | ------------------------------------------------------------------ |
| `MODEL_NAME`                    | Exact local model ID loaded in LM Studio                           |
| `ENDPOINT_TEXT_CHUNK_SIZE`      | Maximum text size sent to the model for endpoint/objective mapping |
| `SCHEDULE_TEXT_CHUNK_SIZE`      | Maximum text size sent to the model for schedule row extraction    |
| `TEXT_CHUNK_OVERLAP`            | Overlap between text chunks                                        |
| `ROWS_PER_CLASSIFICATION_BATCH` | Number of schedule rows classified per AI request                  |
| `MAX_SCHEDULE_PAGES`            | Maximum schedule pages selected                                    |
| `MAX_SCHEDULE_SECTION_SPAN`     | Maximum span for one detected schedule section                     |

For slower local machines, reduce:

```python
ROWS_PER_CLASSIFICATION_BATCH = 3
SCHEDULE_TEXT_CHUNK_SIZE = 2500
ENDPOINT_TEXT_CHUNK_SIZE = 2500
```

---

## How to Run

1. Start LM Studio.
2. Load the selected local model.
3. Enable the local server.
4. Confirm the server is running at:

```text
http://127.0.0.1:1234
```

5. Run the Python script:

```bash
python protocol_classifier.py
```

6. In the GUI:

   * Click **Browse Protocol PDF**
   * Select the clinical study protocol PDF
   * Click **Browse Output Folder**
   * Select the folder where outputs should be saved
   * Click **Run / OK**

---

## Output Structure

Each run creates a new timestamped folder:

```text
Selected_Output_Folder/
│
├── protocol_name_analysis_YYYYMMDD_HHMMSS/
│   ├── protocol_name_classification_output.xlsx
│   ├── protocol_name_classification_output.json
│   └── metadata/
│       ├── protocol_name_bookmarks.json
│       ├── protocol_name_toc_text.txt
│       ├── protocol_name_target_page_diagnostics.json
│       ├── protocol_name_endpoint_text.txt
│       ├── protocol_name_schedule_text.txt
│       ├── protocol_name_supporting_text.txt
│       ├── protocol_name_study_map.json
│       ├── protocol_name_schedule_rows.json
│       ├── *_raw_ai_response.txt
│       ├── *_parsed_ai_response.json
│       └── *_parse_failed.json
│
└── archive/
```

Previous run folders for the same protocol are moved into the `archive` folder.

---

## Excel Output Sheets

The Excel workbook contains:

### Summary

Classification counts, percentages, pie chart, and endpoint map status.

### Study Map

Unified endpoint/objective reference concepts used for classification.

Columns include:

* Concept ID
* Concept Type
* Reference Concept Text
* Hierarchy
* Safety/Efficacy Domain
* Linked Objective or Endpoint
* Reason

### Extracted Schedule Rows

Rows extracted from the Schedule of Assessments / Schedule of Activities.

Columns include:

* Row ID
* Activity Name
* Source Table
* Source Page
* Visit / Timepoint
* Raw Row Text

### Classified Activities

Final classification for every extracted schedule activity.

Columns include:

* Row ID
* Activity Name
* Source Table
* Source Page
* Visit / Timepoint
* Classification
* Linked Endpoint or Requirement
* Reason

### QC Checks

Quality-control checks, including whether endpoint pages and schedule pages were selected and whether every extracted row was classified.

### Metadata

Run configuration, file paths, selected pages, model name, and output locations.

---

## Classification Logic

The code uses study objectives and endpoints as reference concepts.

Each reference concept is mapped to:

```text
Hierarchy:
- primary
- secondary
- additional_or_exploratory
- unclear

Domain:
- safety
- efficacy
- other
- unclear
```

Schedule activities are classified based on whether they directly support collection, measurement, assessment, confirmation, source-data capture, calculation, follow-up, reporting, or interpretation of an endpoint or objective.

If the activity supports study conduct, compliance, subject protection, safety oversight, data integrity, or regulatory operations but is not directly tied to an endpoint/objective, it is classified as:

```text
general_study_operations_related_to_gcp
```

If the activity cannot be linked to endpoints, objectives, or GCP operations, it is classified as:

```text
not_classifiable
```

---

## GCP Reference Activities

The script includes general GCP/study-operation reference categories, including:

* Informed consent
* Eligibility assessment
* Demographics
* Medical history
* Concomitant medications
* Adverse event reporting
* Serious adverse event reporting
* Device deficiency reporting
* Protocol deviations
* Source data
* CRF/EDC data entry
* Monitoring
* Quality assurance
* Essential documents
* Trial master file
* Product or device accountability

These references help classify activities that are operational or compliance-related rather than endpoint-related.

---

## Troubleshooting

### LM Studio Timeout

Error:

```text
HTTPConnectionPool(host='127.0.0.1', port=1234): Read timed out
```

Meaning:

The local model did not respond before the script timeout.

Try:

* Reduce `ROWS_PER_CLASSIFICATION_BATCH`
* Reduce `SCHEDULE_TEXT_CHUNK_SIZE`
* Reduce `ENDPOINT_TEXT_CHUNK_SIZE`
* Use a faster model
* Increase GPU offload in LM Studio
* Make sure no previous generation is still running

---

### Context Size Exceeded

Error:

```text
Context size has been exceeded
```

Meaning:

The prompt sent to the model is too large for the loaded model context.

Try:

```python
ENDPOINT_TEXT_CHUNK_SIZE = 2500
SCHEDULE_TEXT_CHUNK_SIZE = 2500
ROWS_PER_CLASSIFICATION_BATCH = 3
```

Also increase context length in LM Studio if your system supports it.

---

### Only a Few Schedule Rows Extracted

Check these metadata files:

```text
metadata/*_target_page_diagnostics.json
metadata/*_schedule_text.txt
metadata/*_schedule_rows.json
metadata/schedule_rows_chunk_*_raw_ai_response.txt
```

Possible causes:

1. Schedule pages were not targeted correctly.
2. PDF table extraction did not capture all rows.
3. The model stopped outputting before all rows were extracted.
4. The schedule table is image-based or scanned.
5. The schedule table is too complex for plain text extraction.

Possible fixes:

* Increase schedule row extraction `max_tokens`
* Reduce schedule text chunk size
* Use a better model
* Try OCR if the PDF is scanned
* Manually inspect `*_schedule_text.txt`

---

### PyMuPDF Warnings

Warning:

```text
MuPDF error: format error: No common ancestor in structure tree
```

This is usually caused by complex or malformed PDF structure. The script may still continue.

If reading the PDF becomes too slow, replace the table-aware `read_pdf_pages()` function with a simpler text-only extraction function.

---

## Limitations


Limitations include:

* Complex tables may not extract perfectly.
* Scanned/image-based PDFs may require OCR.
* Local LLMs may omit rows or produce imperfect JSON.
* Classification depends on the clarity of the protocol text.
* Endpoint/objective interpretation should be reviewed by a qualified user.
* Outputs should be validated before use in regulated decision-making.

---

## Recommended Validation


Confirm that:

* The correct endpoint/objective pages were selected.
* The correct Schedule of Activities pages were selected.
* All schedule activities were extracted.
* Objectives/endpoints were not classified as activities.
* Activity classifications are reasonable.
* Not-classifiable rows are justified.

---



## Disclaimer

This software is intended for internal analysis, prototyping, and review support. It is not validated software for regulated clinical trial decision-making. Users are responsible for reviewing, validating, and approving all outputs before operational, regulatory, clinical, or quality use.
