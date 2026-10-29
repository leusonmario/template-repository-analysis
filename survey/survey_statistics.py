#!/usr/bin/env python3
"""Compute all survey statistics used in the GitHub Template Repositories journal extension.

Input
-----
The .xlsx Google Forms/Sheets response workbook used in the study. CSV is also accepted.
The script uses only Python's standard library, so no pandas/openpyxl dependency is required.

Outputs
-------
- survey_report.md                 Human-readable report grouped by Data Collection / RQ1 / RQ2 / RQ3
- survey_statistics.json           Machine-readable full result set
- statistics_long.csv              Long-format counts and percentages
- likert_practices.csv             Five-point Likert + grouped agreement for G1--G7/P1--P5
- qualitative_codes.csv            Frequencies of OQ1--OQ6 codes

Notes
-----
* Percentages for checkbox/multiple-selection questions use the final analytical sample (N=114)
  unless a section explicitly states another denominator.
* P1 adaptation statistics are also computed for the subset excluding respondents who selected
  "created as a template from the beginning" (N=83 in the current dataset).
* Recruitment counts 55,740 -> 29,591 -> 6,652 are external to the response CSV. They are accepted
  as CLI arguments and reported separately. The script does NOT invent an undeliverable-email count.

python survey_statistics.py survey_responses.csv --outdir results
"""

from __future__ import annotations

import argparse
import csv
import json
import re
import io
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter
from pathlib import Path
from typing import Dict, Iterable, List, Mapping, Sequence, Tuple

# -----------------------------------------------------------------------------
# Study-level recruitment numbers (not derivable from the survey response CSV)
# -----------------------------------------------------------------------------
DEFAULT_INITIAL_CANDIDATES = 55_740
DEFAULT_AFTER_BOT_FILTER = 29_591
DEFAULT_CONTACTED = 6_652

# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------

def clean_text(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def norm(text: str) -> str:
    """Normalize whitespace/case for robust header matching."""
    return re.sub(r"\s+", " ", clean_text(text)).casefold()


def pct(n: int, d: int) -> float:
    return round(100.0 * n / d, 1) if d else 0.0


def read_csv(path: Path) -> Tuple[List[str], List[Dict[str, str]]]:
    """Read a CSV export. Kept as an optional fallback."""
    # utf-8-sig handles both plain UTF-8 and UTF-8 with BOM.
    text = path.read_text(encoding="utf-8-sig")
    sample = text[:10000]
    try:
        sniffed = csv.Sniffer().sniff(sample, delimiters=",;\t")
        delimiter = sniffed.delimiter
    except csv.Error:
        delimiter = ","

    # Keep standard CSV quoting semantics. Sniffer can mis-detect doublequote behavior when
    # free-text survey responses contain quotation marks, which can shift later columns.
    reader = csv.DictReader(
        io.StringIO(text), delimiter=delimiter, quotechar='"', doublequote=True,
        skipinitialspace=False
    )
    headers = reader.fieldnames or []
    rows = [dict(r) for r in reader]
    rows = [r for r in rows if any(clean_text(v) for v in r.values())]
    return headers, rows


def _xlsx_col_index(cell_ref: str) -> int:
    """Return a zero-based column index from an Excel cell reference such as AB12."""
    letters = "".join(ch for ch in cell_ref if ch.isalpha()).upper()
    value = 0
    for ch in letters:
        value = value * 26 + (ord(ch) - ord("A") + 1)
    return value - 1


def _xlsx_shared_strings(zf: zipfile.ZipFile) -> List[str]:
    """Read Excel shared strings, including rich-text runs."""
    try:
        data = zf.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    root = ET.fromstring(data)
    ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    strings: List[str] = []
    for si in root.findall(f"{ns}si"):
        strings.append("".join(t.text or "" for t in si.iter(f"{ns}t")))
    return strings


def _xlsx_sheet_path(zf: zipfile.ZipFile, sheet_name: str | None = None) -> Tuple[str, str]:
    """Resolve a worksheet display name to its internal XML path."""
    main_ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    rel_ns = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    pkg_rel_ns = "{http://schemas.openxmlformats.org/package/2006/relationships}"

    wb = ET.fromstring(zf.read("xl/workbook.xml"))
    sheets = wb.find(f"{main_ns}sheets")
    if sheets is None or not list(sheets):
        raise ValueError("The workbook does not contain any worksheets.")

    available: List[Tuple[str, str]] = []
    for sh in sheets:
        name = sh.attrib.get("name", "")
        rid = sh.attrib.get(f"{rel_ns}id", "")
        available.append((name, rid))

    if sheet_name is None:
        chosen_name, chosen_rid = available[0]
    else:
        matches = [(name, rid) for name, rid in available if name == sheet_name]
        if not matches:
            raise ValueError(
                f"Worksheet {sheet_name!r} not found. Available sheets: "
                + ", ".join(name for name, _ in available)
            )
        chosen_name, chosen_rid = matches[0]

    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    target = None
    for rel in rels.findall(f"{pkg_rel_ns}Relationship"):
        if rel.attrib.get("Id") == chosen_rid:
            target = rel.attrib.get("Target")
            break
    if not target:
        raise ValueError(f"Could not resolve worksheet relationship for {chosen_name!r}.")

    # Targets are normally 'worksheets/sheet1.xml'. Normalize to the zip path under xl/.
    target = target.lstrip("/")
    sheet_path = target if target.startswith("xl/") else f"xl/{target}"
    return chosen_name, sheet_path


def read_xlsx(path: Path, sheet_name: str | None = None) -> Tuple[List[str], List[Dict[str, str]]]:
    """Read the response worksheet directly from .xlsx using only the standard library.

    The survey workbook contains plain values; formulas/styles are irrelevant to the statistics.
    Missing Excel cells are preserved as blanks by using each cell's A1 reference.
    """
    main_ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

    with zipfile.ZipFile(path) as zf:
        shared = _xlsx_shared_strings(zf)
        _, sheet_path = _xlsx_sheet_path(zf, sheet_name)
        root = ET.fromstring(zf.read(sheet_path))

    parsed_rows: List[List[str]] = []
    max_cols = 0
    sheet_data = root.find(f"{main_ns}sheetData")
    if sheet_data is None:
        return [], []

    for row_el in sheet_data.findall(f"{main_ns}row"):
        values: Dict[int, str] = {}
        for cell in row_el.findall(f"{main_ns}c"):
            ref = cell.attrib.get("r", "A1")
            col_idx = _xlsx_col_index(ref)
            cell_type = cell.attrib.get("t")

            value = ""
            if cell_type == "inlineStr":
                is_el = cell.find(f"{main_ns}is")
                if is_el is not None:
                    value = "".join(t.text or "" for t in is_el.iter(f"{main_ns}t"))
            else:
                v = cell.find(f"{main_ns}v")
                raw = "" if v is None or v.text is None else v.text
                if cell_type == "s" and raw:
                    try:
                        value = shared[int(raw)]
                    except (ValueError, IndexError):
                        value = raw
                elif cell_type == "b":
                    value = "TRUE" if raw == "1" else "FALSE"
                else:
                    value = raw

            values[col_idx] = value
            max_cols = max(max_cols, col_idx + 1)

        if values:
            row = [""] * max_cols
            for idx, value in values.items():
                if idx >= len(row):
                    row.extend([""] * (idx + 1 - len(row)))
                row[idx] = value
            parsed_rows.append(row)

    if not parsed_rows:
        return [], []

    width = max(len(r) for r in parsed_rows)
    parsed_rows = [r + [""] * (width - len(r)) for r in parsed_rows]
    headers = [clean_text(v) for v in parsed_rows[0]]

    rows: List[Dict[str, str]] = []
    for values in parsed_rows[1:]:
        row = {headers[i]: values[i] for i in range(min(len(headers), len(values))) if headers[i]}
        if any(clean_text(v) for v in row.values()):
            rows.append(row)
    return headers, rows


def read_input(path: Path, sheet_name: str | None = None) -> Tuple[List[str], List[Dict[str, str]]]:
    suffix = path.suffix.lower()
    if suffix == ".xlsx":
        return read_xlsx(path, sheet_name=sheet_name)
    if suffix == ".csv":
        return read_csv(path)
    raise SystemExit(f"Input must be .xlsx or .csv. Got: {path.name}")


def find_header(headers: Sequence[str], *, startswith: str | None = None,
                contains: str | None = None) -> str:
    if (startswith is None) == (contains is None):
        raise ValueError("Specify exactly one of startswith= or contains=")
    matches = []
    for h in headers:
        nh = norm(h)
        if startswith is not None and nh.startswith(norm(startswith)):
            matches.append(h)
        if contains is not None and norm(contains) in nh:
            matches.append(h)
    if len(matches) != 1:
        raise KeyError(
            f"Expected exactly one column for startswith={startswith!r}, contains={contains!r}; "
            f"found {len(matches)}: {matches}"
        )
    return matches[0]


def single_distribution(rows: Sequence[Mapping[str, str]], col: str) -> Counter:
    return Counter(clean_text(r.get(col)) or "<blank>" for r in rows)


def multi_distribution(rows: Sequence[Mapping[str, str]], col: str,
                       options: Mapping[str, str]) -> Counter:
    """Count checkbox options using unique substrings.

    Google Forms exports checkbox answers in one string separated by commas, but many answer
    options themselves contain commas. Naively splitting on commas is therefore incorrect.
    Instead, each known option has a unique identifying substring and is counted by containment.
    """
    out = Counter()
    for r in rows:
        raw = clean_text(r.get(col))
        for label, needle in options.items():
            if needle in raw:
                out[label] += 1
    return out


def code_distribution(
    rows: Sequence[Mapping[str, str]],
    col: str,
    excluded_codes: Iterable[str] = ()
) -> Tuple[int, Counter]:

    out = Counter()
    nonempty = 0
    excluded = {norm(code) for code in excluded_codes}

    for r in rows:
        raw = clean_text(r.get(col))
        if not raw:
            continue

        codes = [
            code.strip()
            for code in raw.split(",")
            if code.strip() and norm(code) not in excluded
        ]

        if not codes:
            continue

        nonempty += 1
        out.update(codes)

    return nonempty, out


def rows_with_any(rows: Sequence[Mapping[str, str]], col: str,
                  needles: Iterable[str]) -> int:
    needles = list(needles)
    return sum(any(n in clean_text(r.get(col)) for n in needles) for r in rows)


def record(long_rows: List[Dict], section: str, question: str, item: str,
           count: int, denominator: int, note: str = "") -> None:
    long_rows.append({
        "section": section,
        "question": question,
        "item": item,
        "count": count,
        "denominator": denominator,
        "percent": pct(count, denominator),
        "note": note,
    })


def md_table(headers: Sequence[str], rows: Sequence[Sequence[object]]) -> str:
    lines = [
        "| " + " | ".join(headers) + " |",
        "| " + " | ".join(["---"] * len(headers)) + " |",
    ]
    for row in rows:
        lines.append("| " + " | ".join(str(x) for x in row) + " |")
    return "\n".join(lines)


# -----------------------------------------------------------------------------
# Survey schema
# -----------------------------------------------------------------------------

def find_exact_header(headers: Sequence[str], name: str) -> str:
    matches = [
        h for h in headers
        if norm(h) == norm(name)
    ]

    if len(matches) != 1:
        raise KeyError(
            f"Expected exactly one column named {name!r}; "
            f"found {len(matches)}: {matches}"
        )

    return matches[0]

def resolve_columns(headers: Sequence[str]) -> Dict[str, str]:
    cols = {
        "consent": find_header(headers, startswith="Do you agree to participate in this research project?"),
        "eligibility": find_header(headers, startswith="Have you ever created, maintained, or contributed to a GitHub template repository?"),
        "roles": find_header(headers, startswith="What was your role in the template repository/repositories?"),
        "motivations": find_header(headers, startswith="What motivated you to create, maintain, or contribute to a GitHub template repository?"),
        "contexts": find_header(headers, startswith="In what context did you create, maintain, or contribute to template repositories?"),
        "domains": find_exact_header(headers, "Template Domains"),
        "audiences": find_header(headers, startswith="Who was the intended audience of the template repository?"),
        "num_templates": find_header(headers, startswith="How many GitHub template repositories have you created, maintained, or contributed to?"),
        "adaptations": find_header(headers, startswith="In case you turned a regular repository into a template"),
        "ai": find_header(headers, startswith="Have you used AI-based tools when creating, maintaining, or contributing to GitHub template repositories?"),
        "maintenance_active": find_header(headers, startswith="How actively do you maintain the template repository/repositories?"),
        "maintenance_aspects": find_header(headers, startswith="Which aspects do you consider important when maintaining a GitHub template repository?"),
        "maintenance_signals": find_header(headers, startswith="When maintaining a template repository, what types of activity do you consider meaningful evidence of maintenance?"),
        "ownership": find_header(headers, startswith="In your experience, does the ownership context influence how a template repository is maintained?"),
        "quality_propagation": find_header(headers, startswith="When maintaining a template repository, do you consider that quality issues in the template may affect projects created from it?"),
        "quality_monitoring": find_header(headers, startswith="Do you actively monitor quality issues in your template repositories?"),
        "followup": find_header(headers, startswith="Would you be willing to participate in a follow-up interview?"),
        "github_frequency": find_header(headers, startswith="How frequently do you use GitHub?"),
        "development_type": find_header(headers, startswith="What type of development do you primarily work on?"),
        "current_position": find_exact_header(headers,"Position",),
        "current_position_raw": find_header(headers, startswith="What is your current position?"),
        "years_position": find_header(headers, startswith="How many years have you been in this position?"),
        "years_dev": find_header(headers, startswith="How many years of experience do you have in software development?"),
        "org_size": find_header(headers, startswith="What is the size of the organization where you currently work?"),
        "gender": find_header(headers, startswith="What is your gender? Optional."),
    }

    # Optional normalized/coded position column from the analyzed sheet.
    #coded_position_matches = [h for h in headers if norm(h) == "position"]
    #if len(coded_position_matches) == 1:
    #    cols["position_coded"] = coded_position_matches[0]

    for oq in ["OQ1", "OQ2", "OQ3", "OQ4", "OQ5", "OQ6"]:
        matches = [h for h in headers if clean_text(h) == oq]
        if len(matches) == 1:
            cols[oq.lower()] = matches[0]

    practice_needles = {
        "G1": "integrate maintenance and quality practices",
        "G2": "provide documentation that explains how to understand, customize, and extend",
        "G3": "avoid contradictory usage instructions",
        "G4": "communicate their evolution through a roadmap or maintenance plan",
        "G5": "organized as a reusable family",
        "G6": "align supported technology versions with template releases",
        "G7": "encourage community or social support",
        "P1": "only be marked as a template if it is designed to be reused",
        "P2": "clearly indicate whether a template is active, inactive, or no longer maintained",
        "P3": "clear and focused purpose",
        "P4": "separated from repositories that mainly host a product, framework, or library implementation",
        "P5": "include enough content and documentation to help users understand and adapt it",
    }
    for code, needle in practice_needles.items():
        cols[code] = find_header(headers, contains=needle)

    return cols


OPTIONS: Dict[str, Dict[str, str]] = {
    "roles": {
        "Creator": "Creator",
        "Maintainer": "Maintainer",
        "Contributor": "Contributor",
        "Organization/team member involved in template design": "Organization/team member involved in template design",
        "User who later contributed improvements": "User who later contributed improvements",
    },
    "motivations": {
        "Reusing project structure across multiple projects": "Reusing project structure across multiple projects",
        "Reducing setup/configuration effort for others": "Reducing setup/configuration effort for others",
        "Standardizing development practices": "Standardizing development practices",
        "Demonstrating a reusable project structure": "Demonstrating a reusable project structure",
        "Sharing best practices with the community": "Sharing best practices with the community",
        "Encouraging consistent tooling, such as testing, linting, or CI/CD": "Encouraging consistent tooling, such as testing, linting, or CI/CD",
        "Supporting a specific technology, framework, or ecosystem": "Supporting a specific technology, framework, or ecosystem",
        "Supporting team onboarding": "Supporting team onboarding",
        "Supporting teaching or training": "Supporting teaching or training",
    },
    "contexts": {
        "Open-source projects": "Open-source projects",
        "Personal projects": "Personal projects",
        "Industry/professional projects": "Industry/professional projects",
        "Internal organizational reuse": "Internal organizational reuse",
        "Experimentation or prototyping": "Experimentation or prototyping",
        "Academic research": "Academic research",
        "Teaching or educational activities": "Teaching or educational activities",
        "Team onboarding": "Team onboarding",
    },
    "domains": {
        # Unique prefixes are sufficient and avoid parsing long choice descriptions that contain commas.
        "Automation/DevOps": "Automation/DevOps:",
        "Communication/Bots": "Communication/Bots:",
        "Data Apps/Visualization": "Data Apps/Visualization:",
        "Data Management": "Data Management:",
        "Education": "Education:",
        "Game Development": "Game Development:",
        "Infrastructure/Cloud": "Infrastructure/Cloud:",
        "Machine Learning/AI": "Machine Learning/AI:",
        "Mobile Development": "Mobile Development:",
        "Plugin/Extensibility Frameworks": "Plugin/Extensibility Frameworks:",
        "Regular Project": "Regular Project:",
        "Research": "Research:",
        "Robotics/IoT": "Robotics/IoT:",
        "Web Development": "Web Development:",
    },
    "audiences": {
        "Myself": "Myself",
        "External developers/community": "External developers/community",
        "Open-source contributors": "Open-source contributors",
        "My team": "My team",
        "My organization/company": "My organization/company",
        "Students": "Students",
        "Researchers": "Researchers",
    },
    "adaptations": {
        "Removed project-specific content": "Removed project-specific content, such as real data, generated files, or local environment settings",
        "Removed sensitive information": "Removed sensitive information, such as API keys, passwords, tokens, or private configuration files",
        "Replaced real configuration files with examples": "Replaced real configuration files with examples, such as .env.example",
        "Replaced fixed names/paths/URLs/project information with placeholders": "Replaced fixed names, paths, URLs, or author/project information with placeholders",
        "Rewrote or updated README": "Rewrote or updated the README to explain how to use the repository as a template",
        "Added setup/installation/configuration instructions": "Added setup, installation, or configuration instructions",
        "Added examples/comments/guidance": "Added examples, comments, or guidance to help users adapt the template",
        "Added supporting files": "Added supporting files, such as .gitignore, LICENSE, CONTRIBUTING.md, or issue/PR templates",
        "Added or updated automation": "Added or updated automation, such as tests, linters, CI/CD workflows, or dependency update tools",
        "No specific changes": "I did not make specific changes before marking it as a template",
        "Created as a template from the beginning": "Not applicable; the repository was created as a template from the beginning",
    },
    "ai": {
        "Generate/improve documentation": "Yes, to generate or improve documentation",
        "Review code/detect issues/suggest improvements": "Yes, to review code, detect issues, or suggest improvements",
        "Improve/refactor template structure": "Yes, to improve or refactor the template structure",
        "Configure CI/CD/tests/linters/dependency management": "Yes, to configure tools such as CI/CD, tests, linters, or dependency management",
        "Generate initial template structure": "Yes, to generate an initial template structure",
        "Create templates for AI-based applications": "Yes, to create templates for AI-based applications, such as LLM, chatbot, RAG, or ML projects",
        "No AI use": "No, I have not used AI tools for template repository work",
        "Not sure": "I am not sure",
    },
    "maintenance_aspects": {
        "Keeping dependencies up to date": "Keeping dependencies up to date",
        "Fixing reported issues": "Fixing reported issues",
        "Updating documentation": "Updating documentation",
        "Monitoring vulnerabilities or security alerts": "Monitoring vulnerabilities or security alerts",
        "Updating supported technology versions": "Updating supported technology versions",
        "Maintaining CI/CD pipelines": "Maintaining CI/CD pipelines",
        "Responding to user questions": "Responding to user questions",
        "Publishing releases or changelogs": "Publishing releases or changelogs",
        "Maintaining tests": "Maintaining tests",
    },
    "maintenance_signals": {
        "Recent commits": "Recent commits",
        "Dependency updates": "Dependency updates",
        "Issue responses": "Issue responses",
        "Security fixes": "Security fixes",
        "Recent releases": "Recent releases",
        "Pull request reviews": "Pull request reviews",
        "Technology/version updates": "Technology/version updates",
        "Documentation updates": "Documentation updates",
        "CI/CD updates": "CI/CD updates",
    },
    "quality_monitoring": {
        "Bugs": "Yes, bugs",
        "Security alerts/dependency risks": "Yes, security alerts or dependency risks",
        "Code smells/maintainability issues": "Yes, code smells or maintainability issues",
        "Vulnerabilities": "Yes, vulnerabilities",
        "Outdated dependencies": "Yes, outdated dependencies",
        "Build/CI failures": "Yes, build or CI failures",
        "Do not actively monitor": "No, I do not actively monitor quality issues",
        "Not applicable": "Not applicable",
    },
    "development_type": {
        "Personal projects": "Personal projects",
        "Open-source projects": "Open-source projects",
        "Industry projects": "Industry projects",
        "Academic research": "Academic research",
    },
}

LIKERT_ORDER = ["Strongly disagree", "Disagree", "Neutral", "Agree", "Strongly agree"]
PRACTICES = ["G1", "G2", "G3", "G4", "G5", "G6", "G7", "P1", "P2", "P3", "P4", "P5"]


def compute_all(headers: Sequence[str], all_rows: Sequence[Mapping[str, str]],
                initial_candidates: int = DEFAULT_INITIAL_CANDIDATES,
                after_bot_filter: int = DEFAULT_AFTER_BOT_FILTER,
                contacted: int = DEFAULT_CONTACTED,
                undeliverable: int | None = None) -> Dict:
    cols = resolve_columns(headers)

    total_submissions = len(all_rows)
    consent_dist = single_distribution(all_rows, cols["consent"])
    eligibility_dist = single_distribution(all_rows, cols["eligibility"])
    analysis_rows = [r for r in all_rows if clean_text(r.get(cols["eligibility"])) == "Yes"]
    n = len(analysis_rows)

    result: Dict = {
        "recruitment": {
            "initial_candidate_records": initial_candidates,
            "after_bot_automation_noncontactable_filter": after_bot_filter,
            "selected_and_contacted_over_50_commits": contacted,
            "undeliverable_emails": undeliverable,
            "delivered_invitations": (contacted - undeliverable) if undeliverable is not None else None,
        },
        "data_collection": {
            "total_submissions": total_submissions,
            "consent": dict(consent_dist),
            "screening": dict(eligibility_dist),
            "final_analytical_sample": n,
        },
        "distributions": {},
        "multi_select": {},
        "likert": {},
        "qualitative_codes": {},
        "derived": {},
    }

    if undeliverable is not None:
        delivered = contacted - undeliverable
        if delivered <= 0:
            raise ValueError("undeliverable must be smaller than contacted")
        result["recruitment"]["submission_response_rate_percent"] = pct(total_submissions, delivered)
        result["recruitment"]["analytical_response_rate_percent"] = pct(n, delivered)

    # Single-choice / categorical distributions.
    singles = [
        "num_templates", "maintenance_active", "ownership", "quality_propagation",
        "followup", "github_frequency", "current_position", "years_position", "years_dev",
        "org_size", "gender",
    ]
    #if "position_coded" in cols:
    #    singles.append("position_coded")
    for key in singles:
        result["distributions"][key] = dict(single_distribution(analysis_rows, cols[key]))

    # Multi-select distributions.
    for key, options in OPTIONS.items():
        result["multi_select"][key] = dict(multi_distribution(analysis_rows, cols[key], options))

    # Derived participant profile statistics.
    gh = Counter(result["distributions"]["github_frequency"])
    ydev = Counter(result["distributions"]["years_dev"])
    pos = Counter(result["distributions"]["current_position"])
    nt = Counter(result["distributions"]["num_templates"])
    result["derived"]["github_daily"] = gh["Daily"]
    result["derived"]["github_frequent"] = gh["Frequently"]
    result["derived"]["github_daily_or_frequent"] = gh["Daily"] + gh["Frequently"]
    result["derived"]["development_experience_5plus_years"] = ydev["5–10 years"] + ydev["5-10 years"] + ydev["More than 10 years"]
    result["derived"]["software_engineer_raw_position"] = pos["Software Engineer"]
    result["derived"]["experience_more_than_one_template"] = n - nt["1"]
    result["derived"]["experience_at_least_five_templates"] = nt["5-10"] + nt["5–10"] + nt["10-20"] + nt["10–20"] + nt["More than 20"]

    # RQ2: any explicit positive quality monitoring category.
    qmon_positive = [needle for label, needle in OPTIONS["quality_monitoring"].items()
                     if label not in {"Do not actively monitor", "Not applicable"}]
    result["derived"]["monitor_at_least_one_quality_issue"] = rows_with_any(
        analysis_rows, cols["quality_monitoring"], qmon_positive
    )

    # RQ3: AI -- options are checkbox/non-exclusive; compute "any positive use" separately.
    ai_positive = [needle for label, needle in OPTIONS["ai"].items()
                   if label not in {"No AI use", "Not sure"}]
    result["derived"]["any_positive_ai_use"] = rows_with_any(analysis_rows, cols["ai"], ai_positive)

    # RQ3/P1: adaptation statistics using the appropriate subset denominator.
    created_from_start_needle = OPTIONS["adaptations"]["Created as a template from the beginning"]
    converted_rows = [r for r in analysis_rows if created_from_start_needle not in clean_text(r.get(cols["adaptations"]))]
    converted_counts = multi_distribution(converted_rows, cols["adaptations"], OPTIONS["adaptations"])
    result["derived"]["p1_conversion_subset_n"] = len(converted_rows)
    result["derived"]["p1_conversion_adaptations"] = dict(converted_counts)

    # Likert practice distributions and grouped agreement.
    for code in PRACTICES:
        c = single_distribution(analysis_rows, cols[code])
        missing_or_other = {k: v for k, v in c.items() if k not in LIKERT_ORDER}
        agree = c["Agree"] + c["Strongly agree"]
        disagree = c["Disagree"] + c["Strongly disagree"]
        result["likert"][code] = {
            "counts": {k: c[k] for k in LIKERT_ORDER},
            "agree_or_strongly_agree": agree,
            "neutral": c["Neutral"],
            "disagree_or_strongly_disagree": disagree,
            "other_or_missing": missing_or_other,
        }

    # Qualitative codes -- exact strings, no silent harmonization.
    excluded_by_question = {
        "oq4": {"NoProblem"},
        "oq5": {"-"},
    }

    for oq in ["oq1", "oq2", "oq3", "oq4", "oq5", "oq6"]:
        if oq in cols:
            nonempty, counts = code_distribution(
                analysis_rows,
                cols[oq],
                excluded_codes=excluded_by_question.get(oq, set())
            )

            result["qualitative_codes"][oq.upper()] = {
                "responses_with_codes": nonempty,
                "counts": dict(counts),
            }

    return result


# -----------------------------------------------------------------------------
# Reporting
# -----------------------------------------------------------------------------

def build_long_rows(stats: Dict) -> List[Dict]:
    n = stats["data_collection"]["final_analytical_sample"]
    rows: List[Dict] = []

    # Data Collection distributions.
    for key in ["github_frequency", "years_dev", "current_position", "num_templates"]:
        for item, count in Counter(stats["distributions"][key]).most_common():
            record(rows, "Data Collection", key, item, count, n)

    for key in ["roles"]:
        for item, count in Counter(stats["multi_select"][key]).most_common():
            record(rows, "Data Collection", key, item, count, n, "multiple-selection")

    # RQ1.
    for key in ["motivations", "contexts", "domains", "audiences"]:
        for item, count in Counter(stats["multi_select"][key]).most_common():
            record(rows, "RQ1", key, item, count, n, "multiple-selection")

    # RQ2.
    for key in ["maintenance_active", "ownership", "quality_propagation"]:
        for item, count in Counter(stats["distributions"][key]).most_common():
            record(rows, "RQ2", key, item, count, n)
    for key in ["maintenance_aspects", "maintenance_signals", "quality_monitoring"]:
        for item, count in Counter(stats["multi_select"][key]).most_common():
            record(rows, "RQ2", key, item, count, n, "multiple-selection")
    record(rows, "RQ2", "quality_monitoring", "At least one explicit positive category",
           stats["derived"]["monitor_at_least_one_quality_issue"], n, "derived")

    # RQ3 AI.
    for item, count in Counter(stats["multi_select"]["ai"]).most_common():
        record(rows, "RQ3", "ai", item, count, n, "multiple-selection")
    record(rows, "RQ3", "ai", "Any positive AI use",
           stats["derived"]["any_positive_ai_use"], n, "derived")

    # RQ3 P1 adaptation subset.
    p1n = stats["derived"]["p1_conversion_subset_n"]
    for item, count in Counter(stats["derived"]["p1_conversion_adaptations"]).most_common():
        if item == "Created as a template from the beginning":
            continue
        record(rows, "RQ3", "P1 adaptations among converted repositories", item, count, p1n,
               "subset excludes 'created as template from the beginning'")

    # RQ3 Likert grouped.
    for code in PRACTICES:
        item = stats["likert"][code]
        record(rows, "RQ3", f"Likert {code}", "Agree + Strongly agree",
               item["agree_or_strongly_agree"], n)
        record(rows, "RQ3", f"Likert {code}", "Neutral", item["neutral"], n)
        record(rows, "RQ3", f"Likert {code}", "Disagree + Strongly disagree",
               item["disagree_or_strongly_disagree"], n)

    return rows


def build_markdown(stats: Dict) -> str:
    n = stats["data_collection"]["final_analytical_sample"]
    dc = stats["data_collection"]
    r = stats["recruitment"]
    d = stats["derived"]

    out: List[str] = []
    out.append("# Survey Statistics Report")
    out.append("")
    out.append("## Recruitment and Data Collection")
    out.append("")
    out.append(f"- Initial candidate contributor records: **{r['initial_candidate_records']:,}**")
    out.append(f"- After bot/automation/non-contactable filtering: **{r['after_bot_automation_noncontactable_filter']:,}**")
    out.append(f"- Selected and contacted (>50 commits): **{r['selected_and_contacted_over_50_commits']:,}**")
    if r["undeliverable_emails"] is not None:
        out.append(f"- Undeliverable emails: **{r['undeliverable_emails']:,}**")
        out.append(f"- Delivered invitations: **{r['delivered_invitations']:,}**")
        out.append(f"- Submission response rate (116 / delivered): **{r['submission_response_rate_percent']:.1f}%**")
        out.append(f"- Analytical response rate ({n} / delivered): **{r['analytical_response_rate_percent']:.1f}%**")
    else:
        out.append("- Undeliverable emails: **not supplied; response rate intentionally not computed**")
    out.append(f"- Survey submissions: **{dc['total_submissions']}**")
    out.append(f"- Screening 'No': **{dc['screening'].get('No', 0)}**")
    out.append(f"- Final analytical sample: **{n}**")
    out.append("")

    out.append("### Participant profile")
    out.append("")
    profile_rows = [
        ["Daily GitHub use", d["github_daily"], f"{pct(d['github_daily'], n):.1f}%"],
        ["Frequent GitHub use", d["github_frequent"], f"{pct(d['github_frequent'], n):.1f}%"],
        ["Daily or frequent GitHub use", d["github_daily_or_frequent"], f"{pct(d['github_daily_or_frequent'], n):.1f}%"],
        ["5+ years software-development experience", d["development_experience_5plus_years"], f"{pct(d['development_experience_5plus_years'], n):.1f}%"],
        ["Software Engineer (raw survey option)", d["software_engineer_raw_position"], f"{pct(d['software_engineer_raw_position'], n):.1f}%"],
        ["Experience with >1 template", d["experience_more_than_one_template"], f"{pct(d['experience_more_than_one_template'], n):.1f}%"],
        ["Experience with >=5 templates", d["experience_at_least_five_templates"], f"{pct(d['experience_at_least_five_templates'], n):.1f}%"],
    ]
    out.append(md_table(["Statistic", "n", "%"], profile_rows))
    out.append("")

    def add_multi_section(title: str, key: str, denom: int = n):
        out.append(f"### {title}")
        out.append("")
        c = Counter(stats["multi_select"][key])
        out.append(md_table(["Option", "n", "%"], [
            [item, count, f"{pct(count, denom):.1f}%"] for item, count in c.most_common()
        ]))
        out.append("")

    out.append("## RQ1 — Purposes, Domains, and Ownership Context")
    out.append("")
    add_multi_section("Roles", "roles")
    add_multi_section("Motivations", "motivations")
    add_multi_section("Development contexts", "contexts")
    add_multi_section("Practitioner-reported template domains", "domains")
    add_multi_section("Intended audiences", "audiences")

    out.append("## RQ2 — Maintenance and Quality")
    out.append("")
    for title, key in [
        ("Maintenance frequency", "maintenance_active"),
        ("Ownership perception", "ownership"),
        ("Perceived downstream quality propagation", "quality_propagation"),
    ]:
        out.append(f"### {title}")
        out.append("")
        c = Counter(stats["distributions"][key])
        out.append(md_table(["Option", "n", "%"], [
            [item, count, f"{pct(count, n):.1f}%"] for item, count in c.most_common()
        ]))
        out.append("")
    add_multi_section("Important maintenance activities", "maintenance_aspects")
    add_multi_section("Meaningful evidence of maintenance", "maintenance_signals")
    add_multi_section("Quality issues actively monitored", "quality_monitoring")
    out.append(f"Participants selecting at least one explicit positive quality-monitoring category: "
               f"**{d['monitor_at_least_one_quality_issue']}/{n} ({pct(d['monitor_at_least_one_quality_issue'], n):.1f}%)**.")
    out.append("")

    out.append("## RQ3 — Guidelines, Pitfalls, Challenges, and AI")
    out.append("")
    out.append("### Likert assessment of G1--G7 and P1--P5")
    out.append("")
    likert_rows = []
    for code in PRACTICES:
        s = stats["likert"][code]
        likert_rows.append([
            code,
            f"{s['agree_or_strongly_agree']} ({pct(s['agree_or_strongly_agree'], n):.1f}%)",
            f"{s['neutral']} ({pct(s['neutral'], n):.1f}%)",
            f"{s['disagree_or_strongly_disagree']} ({pct(s['disagree_or_strongly_disagree'], n):.1f}%)",
        ])
    out.append(md_table(["Practice", "Agree+Strongly agree", "Neutral", "Disagree+Strongly disagree"], likert_rows))
    out.append("")

    p1n = d["p1_conversion_subset_n"]
    created_from_beginning = stats["multi_select"]["adaptations"].get("Created as a template from the beginning", 0)
    out.append("### P1 — Adapting a regular repository for template reuse")
    out.append("")
    out.append(f"Created as a template from the beginning: **{created_from_beginning}/{n} ({pct(created_from_beginning, n):.1f}%)**.")
    out.append(f"Conversion/adaptation subset: **N={p1n}**.")
    out.append("")
    c = Counter(d["p1_conversion_adaptations"])
    out.append(md_table(["Adaptation", "n", "% of conversion subset"], [
        [item, count, f"{pct(count, p1n):.1f}%"]
        for item, count in c.most_common()
        if item != "Created as a template from the beginning"
    ]))
    out.append("")

    add_multi_section("AI-based tool use (checkbox options; non-exclusive)", "ai")
    out.append(f"Participants reporting at least one explicit positive AI use: "
               f"**{d['any_positive_ai_use']}/{n} ({pct(d['any_positive_ai_use'], n):.1f}%)**.")
    out.append("")

    out.append("### Qualitative code frequencies")
    out.append("")
    out.append("Codes are reported exactly as stored in OQ1--OQ6. Counts can exceed the number of responses because one response may receive multiple codes.")
    out.append("")
    for oq, info in stats["qualitative_codes"].items():
        nonempty = info["responses_with_codes"]
        out.append(f"#### {oq} (responses with codes: {nonempty})")
        out.append("")
        c = Counter(info["counts"])
        out.append(md_table(["Code", "n", f"% of {oq} coded responses"], [
            [code, count, f"{pct(count, nonempty):.1f}%"] for code, count in c.most_common()
        ]))
        out.append("")

    return "\n".join(out)


def write_outputs(stats: Dict, outdir: Path) -> None:
    outdir.mkdir(parents=True, exist_ok=True)
    n = stats["data_collection"]["final_analytical_sample"]

    (outdir / "survey_statistics.json").write_text(
        json.dumps(stats, indent=2, ensure_ascii=False), encoding="utf-8"
    )
    (outdir / "survey_report.md").write_text(build_markdown(stats), encoding="utf-8")

    long_rows = build_long_rows(stats)
    with (outdir / "statistics_long.csv").open("w", encoding="utf-8-sig", newline="") as f:
        fieldnames = ["section", "question", "item", "count", "denominator", "percent", "note"]
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        w.writerows(long_rows)

    with (outdir / "likert_practices.csv").open("w", encoding="utf-8-sig", newline="") as f:
        fieldnames = ["practice"] + LIKERT_ORDER + [
            "agree_or_strongly_agree", "agree_percent", "neutral_percent",
            "disagree_or_strongly_disagree", "disagree_percent", "denominator"
        ]
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        for code in PRACTICES:
            s = stats["likert"][code]
            row = {"practice": code, **s["counts"]}
            row.update({
                "agree_or_strongly_agree": s["agree_or_strongly_agree"],
                "agree_percent": pct(s["agree_or_strongly_agree"], n),
                "neutral_percent": pct(s["neutral"], n),
                "disagree_or_strongly_disagree": s["disagree_or_strongly_disagree"],
                "disagree_percent": pct(s["disagree_or_strongly_disagree"], n),
                "denominator": n,
            })
            w.writerow(row)

    with (outdir / "qualitative_codes.csv").open("w", encoding="utf-8-sig", newline="") as f:
        fieldnames = ["question", "code", "count", "responses_with_codes", "percent_of_coded_responses"]
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        for oq, info in stats["qualitative_codes"].items():
            denom = info["responses_with_codes"]
            for code, count in Counter(info["counts"]).most_common():
                w.writerow({
                    "question": oq,
                    "code": code,
                    "count": count,
                    "responses_with_codes": denom,
                    "percent_of_coded_responses": pct(count, denom),
                })


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("input", type=Path, help="Google Forms response workbook (.xlsx; .csv also accepted)")
    ap.add_argument("--sheet", default=None, help="Optional worksheet name. Defaults to the first worksheet.")
    ap.add_argument("--outdir", type=Path, default=Path("survey_stats"), help="Output directory (default: survey_stats)")
    ap.add_argument("--initial-candidates", type=int, default=DEFAULT_INITIAL_CANDIDATES)
    ap.add_argument("--after-bot-filter", type=int, default=DEFAULT_AFTER_BOT_FILTER)
    ap.add_argument("--contacted", type=int, default=DEFAULT_CONTACTED)
    ap.add_argument("--undeliverable", type=int, default=None,
                    help="Optional number of undeliverable invitations. If omitted, no response rate is computed.")
    args = ap.parse_args()

    headers, rows = read_input(args.input, sheet_name=args.sheet)
    stats = compute_all(
        headers, rows,
        initial_candidates=args.initial_candidates,
        after_bot_filter=args.after_bot_filter,
        contacted=args.contacted,
        undeliverable=args.undeliverable,
    )
    write_outputs(stats, args.outdir)

    n = stats["data_collection"]["final_analytical_sample"]
    print(f"Survey submissions: {stats['data_collection']['total_submissions']}")
    print(f"Final analytical sample: {n}")
    print(f"Wrote: {args.outdir / 'survey_report.md'}")
    print(f"Wrote: {args.outdir / 'survey_statistics.json'}")
    print(f"Wrote: {args.outdir / 'statistics_long.csv'}")
    print(f"Wrote: {args.outdir / 'likert_practices.csv'}")
    print(f"Wrote: {args.outdir / 'qualitative_codes.csv'}")


if __name__ == "__main__":
    main()
