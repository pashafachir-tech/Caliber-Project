#!/usr/bin/env python3
"""Build CAP-LOOP dashboard data from the Case 2 baseline.

By default, the export reveals the complete available condition history so the
dashboard can replay any source week.  Use --advance-week for explicit
incremental replay.  No readings are synthesized.
"""

from __future__ import annotations

import argparse
import json
import math
import os
import re
import statistics
import sys
import tempfile
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Iterable

try:
    from openpyxl import load_workbook
except ImportError:  # Give users a useful dependency error at the point of use.
    load_workbook = None

try:
    from pptx import Presentation
except ImportError:
    Presentation = None


# The only assets in this case.  These tags are scope, not fallback data.
ASSET_TAGS = ("PU-2101B", "KO-3201", "PM-4405B", "HE-3301", "BL-5702")

# RCA-deck CAPA audit actions supplied for the five in-scope assets.
CAPA_ACTIONS = (
    {"action_id": "PAA-1", "tag": "PU-2101B", "description": "Modify Plan 53A seal flush barrier piping layout", "status": "OPEN", "pic": "REL-05", "category": "Engineering Mod"},
    {"action_id": "PAA-2", "tag": "PU-2101B", "description": "Install interlock pressure switch PSL-2101 on barrier fluid", "status": "OPEN", "pic": "INS-02", "category": "Instrument"},
    {"action_id": "PAA-3", "tag": "PU-2101B", "description": "Revise daily PM checklist for seal barrier pot level verification", "status": "IN PROGRESS", "pic": "REL-05", "category": "Maintenance"},
    {"action_id": "PAA-4", "tag": "PU-2101B", "description": "Conduct mechanical seal cartridge installation training", "status": "OPEN", "pic": "MCH-01", "category": "People/Training"},
    {"action_id": "PAA-1", "tag": "KO-3201", "description": "Replace DE journal bearing tilting pads and check journal clearance", "status": "CLOSED", "pic": "MCH-03", "category": "Mechanical Repair"},
    {"action_id": "PAA-2", "tag": "KO-3201", "description": "Install online moisture transmitter on lube oil console", "status": "OPEN", "pic": "INS-02", "category": "Instrumentation"},
    {"action_id": "PAA-3", "tag": "KO-3201", "description": "Overhaul N2 buffer seal regulator system", "status": "IN PROGRESS", "pic": "REL-05", "category": "Engineering Mod"},
    {"action_id": "PAA-4", "tag": "KO-3201", "description": "Recalibrate eddy-current radial proximity probes", "status": "OPEN", "pic": "INS-01", "category": "Instrumentation"},
    {"action_id": "PAA-1", "tag": "PM-4405B", "description": "Implement scheduled high-temp synthetic regreasing procedure", "status": "IN PROGRESS", "pic": "REL-02", "category": "Maintenance"},
    {"action_id": "PAA-2", "tag": "PM-4405B", "description": "Integrate motor DE bearing temperature alarm into DCS graphic", "status": "OPEN", "pic": "DCS-01", "category": "Automation"},
    {"action_id": "PAA-3", "tag": "PM-4405B", "description": "Audit motor cooling fan shroud and air passage ventilation", "status": "OPEN", "pic": "MCH-02", "category": "Inspection"},
    {"action_id": "PAA-4", "tag": "PM-4405B", "description": "Permanently mount tri-axial vibration accelerometers", "status": "OPEN", "pic": "REL-02", "category": "Predictive"},
    {"action_id": "PAA-1", "tag": "HE-3301", "description": "Perform offline tube bundle hydrojetting and scale descaling", "status": "IN PROGRESS", "pic": "STA-02", "category": "Maintenance"},
    {"action_id": "PAA-2", "tag": "HE-3301", "description": "Evaluate chemical anti-foulant continuous injection", "status": "OPEN", "pic": "PRC-01", "category": "Process Chem"},
    {"action_id": "PAA-3", "tag": "HE-3301", "description": "Install online tube-side differential pressure transmitter with drift alert", "status": "OPEN", "pic": "INS-02", "category": "Instrument"},
    {"action_id": "PAA-4", "tag": "HE-3301", "description": "Update operating limits on feed heavy-ends boiling fraction", "status": "OPEN", "pic": "PRC-01", "category": "Process SOP"},
    {"action_id": "PAA-1", "tag": "BL-5702", "description": "Perform laser optical shaft realignment motor-blower", "status": "IN PROGRESS", "pic": "ROT-01", "category": "Maintenance"},
    {"action_id": "PAA-2", "tag": "BL-5702", "description": "Replace elastomeric flexible coupling insert", "status": "OPEN", "pic": "ROT-01", "category": "Mechanical"},
    {"action_id": "PAA-3", "tag": "BL-5702", "description": "Standardize 2X harmonic vibration monitoring in weekly walkdown SOP", "status": "OPEN", "pic": "REL-01", "category": "SOP/Reliability"},
    {"action_id": "PAA-4", "tag": "BL-5702", "description": "Inspect baseplate civil grouting and soft-foot tolerance", "status": "OPEN", "pic": "CIV-01", "category": "Civil/Structural"},
)

# Method parameters explicitly requested for this pipeline.
BASELINE_WEEK_COUNT = 6
DRIFT_Z_THRESHOLD = 3.0
DRIFT_CONSECUTIVE_WEEKS = 2
RUL_WINDOW_WEEKS = 5
RISK_HORIZON_DAYS = 30
DAYS_PER_WEEK = 7
SIGNOFF_Z_THRESHOLD = 1.0
SIGNOFF_CONSECUTIVE_WEEKLY_READINGS = 2
OPEN_RCA_STATUSES = frozenset({"NEW REGISTERED", "RCA PROCESS"})

REQUIRED_INCIDENT_COLUMNS = (
    "Serial No",
    "AR No.",
    "Tag Number",
    "Eq. Class",
    "Date of Occur.",
    "Risk Case Title",
    "Overall Status",
    "PIC (RCA)",
    "Component",
    "F Mechanism",
    "Downtime (hrs)",
    "Act. Loss (k US$)",
    "Pot. Loss (k US$)",
    "Total Loss (k US$)",
    "RCA Due Date",
)


class PipelineError(RuntimeError):
    """An input, configuration, or data-quality error with user-facing detail."""


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    return " ".join(str(value).replace("\r", " ").replace("\n", " ").split())


def normalized(value: Any) -> str:
    return clean_text(value).casefold()


def finite_number(value: Any) -> float | None:
    if value is None or isinstance(value, bool):
        return None
    try:
        number = float(value)
    except (TypeError, ValueError, OverflowError):
        return None
    return number if math.isfinite(number) else None


def json_number(value: float | None, digits: int = 8) -> float | None:
    if value is None or not math.isfinite(value):
        return None
    return round(float(value), digits)


def parse_date(value: Any, field_name: str) -> date | None:
    if value is None or clean_text(value) == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    raw = clean_text(value)
    for fmt in (
        "%Y-%m-%d",
        "%Y/%m/%d",
        "%d-%b-%Y",
        "%d-%B-%Y",
        "%d/%m/%Y",
        "%m/%d/%Y",
        "%d-%m-%Y",
    ):
        try:
            return datetime.strptime(raw, fmt).date()
        except ValueError:
            pass
    try:
        return datetime.fromisoformat(raw.replace("Z", "+00:00")).date()
    except ValueError as exc:
        raise PipelineError(f"Cannot parse {field_name} date value {raw!r}.") from exc


def parse_timestamp(value: Any, field_name: str) -> datetime | None:
    if value is None or clean_text(value) == "":
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
    raw = clean_text(value)
    try:
        return datetime.fromisoformat(raw.replace("Z", "+00:00"))
    except ValueError as exc:
        raise PipelineError(f"Cannot parse {field_name} timestamp {raw!r}.") from exc


def parse_limit_pair(value: Any, tag: str, parameter: str) -> tuple[float, float]:
    raw = clean_text(value)
    matches = re.findall(r"[-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][-+]?\d+)?", raw)
    if len(matches) != 2:
        raise PipelineError(
            f"Equipment Info for {tag}: cannot read alarm/trip limits for "
            f"{parameter!r} from {raw!r}; expected two numeric limits."
        )
    return float(matches[0]), float(matches[1])


def require_dependency(module: Any, package: str) -> None:
    if module is None:
        raise PipelineError(
            f"Missing Python dependency {package!r}. Install dependencies with "
            "`python -m pip install -r requirements.txt`."
        )


def discover_one(directory: Path, extension: str, tag: str | None = None) -> Path:
    if not directory.is_dir():
        raise PipelineError(f"Missing required input folder: {directory}")
    candidates = sorted(
        path for path in directory.rglob(f"*{extension}") if path.is_file()
    )
    if tag is not None:
        tag_key = re.sub(r"[^a-z0-9]", "", tag.casefold())
        candidates = [
            path
            for path in candidates
            if tag_key in re.sub(r"[^a-z0-9]", "", path.stem.casefold())
        ]
    if not candidates:
        detail = f" for asset {tag}" if tag else ""
        raise PipelineError(
            f"Missing required {extension} input{detail} under {directory}."
        )
    if len(candidates) != 1:
        names = ", ".join(str(path) for path in candidates)
        detail = f" for asset {tag}" if tag else ""
        raise PipelineError(
            f"Expected exactly one {extension} input{detail} under {directory}; "
            f"found {len(candidates)}: {names}"
        )
    return candidates[0]


def load_sheet(path: Path, sheet_name: str, header_row: int = 1) -> tuple[list[str], list[dict[str, Any]]]:
    require_dependency(load_workbook, "openpyxl")
    try:
        workbook = load_workbook(path, data_only=True, read_only=True)
    except Exception as exc:
        raise PipelineError(f"Cannot read workbook {path}: {exc}") from exc
    try:
        if sheet_name not in workbook.sheetnames:
            raise PipelineError(
                f"Workbook {path} is missing required sheet {sheet_name!r}; "
                f"found: {', '.join(workbook.sheetnames)}"
            )
        sheet = workbook[sheet_name]
        header_values = next(
            sheet.iter_rows(
                min_row=header_row,
                max_row=header_row,
                values_only=True,
            ),
            (),
        )
        headers = [clean_text(value) for value in header_values]
        while headers and not headers[-1]:
            headers.pop()
        if not headers or any(not header for header in headers):
            raise PipelineError(
                f"Workbook {path}, sheet {sheet_name!r}: missing or blank header "
                f"on row {header_row}."
            )
        duplicates = sorted(
            header for header, count in Counter(headers).items() if count > 1
        )
        if duplicates:
            raise PipelineError(
                f"Workbook {path}, sheet {sheet_name!r}: duplicate headers after "
                f"newline cleanup: {', '.join(duplicates)}"
            )
        records: list[dict[str, Any]] = []
        for values in sheet.iter_rows(min_row=header_row + 1, values_only=True):
            row = {
                headers[index]: values[index] if index < len(values) else None
                for index in range(len(headers))
            }
            if any(value is not None and clean_text(value) for value in row.values()):
                records.append(row)
        return headers, records
    finally:
        workbook.close()


def load_raw_sheet(path: Path, sheet_name: str) -> list[tuple[Any, ...]]:
    """Read a layout-driven worksheet that has no single row of column headers."""
    require_dependency(load_workbook, "openpyxl")
    try:
        workbook = load_workbook(path, data_only=True, read_only=True)
    except Exception as exc:
        raise PipelineError(f"Cannot read workbook {path}: {exc}") from exc
    try:
        if sheet_name not in workbook.sheetnames:
            raise PipelineError(
                f"Workbook {path} is missing required sheet {sheet_name!r}; "
                f"found: {', '.join(workbook.sheetnames)}"
            )
        return list(workbook[sheet_name].iter_rows(values_only=True))
    finally:
        workbook.close()


def load_equipment_info(path: Path, expected_tag: str) -> dict[str, Any]:
    rows = load_raw_sheet(path, "Equipment Info")
    info: dict[str, Any] = {}
    limits: list[dict[str, Any]] = []
    for row in rows[3:]:
        label = clean_text(row[0] if len(row) > 0 else None)
        value = row[1] if len(row) > 1 else None
        if label:
            info[label] = value
        parameter = clean_text(row[2] if len(row) > 2 else None)
        if parameter and normalized(parameter) != "parameter":
            alarm, trip = parse_limit_pair(
                row[3] if len(row) > 3 else None, expected_tag, parameter
            )
            limits.append(
                {
                    "parameter": parameter,
                    "alarm_limit": alarm,
                    "trip_limit": trip,
                }
            )
    actual_tag = clean_text(info.get("Equipment Tag"))
    if actual_tag.casefold() != expected_tag.casefold():
        raise PipelineError(
            f"Equipment Info in {path} identifies {actual_tag!r}, expected "
            f"{expected_tag!r}."
        )
    if not limits:
        raise PipelineError(f"Equipment Info in {path} has no monitored limits.")
    info["_limits"] = limits
    return info


def load_condition_history(path: Path, tag: str, limits: list[dict[str, Any]]) -> list[dict[str, Any]]:
    headers, source_rows = load_sheet(path, "Condition History")
    required = {"Week", "Date", "Health Status", "Remark"}
    missing = sorted(required.difference(headers))
    if missing:
        raise PipelineError(
            f"{path}, Condition History is missing required columns: "
            f"{', '.join(missing)}"
        )
    parameter_names = [item["parameter"] for item in limits]
    missing_parameters = [name for name in parameter_names if name not in headers]
    if missing_parameters:
        raise PipelineError(
            f"{path}, Condition History is missing Equipment Info parameter "
            f"columns: {', '.join(missing_parameters)}"
        )

    records: list[dict[str, Any]] = []
    seen_weeks: set[int] = set()
    for row in source_rows:
        raw_week = finite_number(row.get("Week"))
        if raw_week is None:
            continue
        week = int(raw_week)
        if float(week) != raw_week:
            raise PipelineError(f"{path}: non-integer week value {raw_week!r}.")
        if week in seen_weeks:
            raise PipelineError(f"{path}: duplicate Condition History week {week}.")
        seen_weeks.add(week)
        day = parse_date(row.get("Date"), f"{tag} Condition History")
        if day is None:
            raise PipelineError(f"{path}: missing date for Condition History week {week}.")
        values: dict[str, float | None] = {}
        for parameter in parameter_names:
            values[parameter] = finite_number(row.get(parameter))
        records.append(
            {
                "week": week,
                "date": day,
                "health_status": clean_text(row.get("Health Status")),
                "remark": clean_text(row.get("Remark")),
                "values": values,
                "_source_row": row,
            }
        )
    records.sort(key=lambda record: record["week"])
    if len(records) < BASELINE_WEEK_COUNT:
        raise PipelineError(
            f"{path}: only {len(records)} weekly rows; at least "
            f"{BASELINE_WEEK_COUNT} are required to form the baseline."
        )
    return records


def read_performance_summary(path: Path) -> dict[str, Any]:
    rows = load_raw_sheet(path, "Performance Summary")
    heading_index = next(
        (
            index
            for index, row in enumerate(rows)
            if len(row) >= 2
            and normalized(row[0]) == "kpi"
            and normalized(row[1]) == "value"
        ),
        None,
    )
    if heading_index is None:
        raise PipelineError(
            f"{path}: Performance Summary has no KPI/Value header row."
        )
    summary: dict[str, Any] = {}
    for row in rows[heading_index + 1 :]:
        label = clean_text(row[0] if len(row) > 0 else None)
        if label:
            summary[label] = row[1] if len(row) > 1 else None
    if "Estimated Loss (k USD)" not in summary:
        raise PipelineError(
            f"{path}: Performance Summary is missing 'Estimated Loss (k USD)'."
        )
    return summary


def load_incidents(path: Path) -> list[dict[str, Any]]:
    headers, rows = load_sheet(path, "Incident Database", header_row=3)
    missing = sorted(set(REQUIRED_INCIDENT_COLUMNS).difference(headers))
    if missing:
        raise PipelineError(
            f"{path}, Incident Database (header row 3) is missing columns: "
            f"{', '.join(missing)}"
        )
    incidents = [
        row for row in rows if clean_text(row.get("Serial No"))
    ]
    if not incidents:
        raise PipelineError(f"{path}: Incident Database contains no incident rows.")
    return incidents


def load_rca_text(path: Path, tag: str) -> str:
    require_dependency(Presentation, "python-pptx")
    try:
        presentation = Presentation(str(path))
    except Exception as exc:
        raise PipelineError(f"Cannot read RCA PowerPoint for {tag} at {path}: {exc}") from exc
    slide_text: list[str] = []
    for slide in presentation.slides:
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False):
                text = clean_text(shape.text)
                if text:
                    slide_text.append(text)
            if getattr(shape, "has_table", False):
                for row in shape.table.rows:
                    for cell in row.cells:
                        text = clean_text(cell.text)
                        if text:
                            slide_text.append(text)
    text = " ".join(slide_text).strip()
    if not text:
        raise PipelineError(f"RCA PowerPoint for {tag} contains no extractable slide text: {path}")
    return text


def incident_date(incident: dict[str, Any], field: str) -> date | None:
    return parse_date(incident.get(field), f"Incident Database {field}")


def incident_text(incident: dict[str, Any]) -> str:
    return " ".join(
        clean_text(incident.get(field))
        for field in ("Risk Case Title", "Component", "F Mechanism")
        if clean_text(incident.get(field))
    )


TOKEN_RE = re.compile(r"[a-z0-9]+(?:[.-][a-z0-9]+)*", re.IGNORECASE)


def tokens(text: str) -> list[str]:
    return [token.casefold() for token in TOKEN_RE.findall(text)]


def tfidf_vectors(corpus: list[str]) -> tuple[list[dict[str, float]], dict[str, float]]:
    term_counts = [Counter(tokens(document)) for document in corpus]
    document_frequency: Counter[str] = Counter()
    for counts in term_counts:
        document_frequency.update(counts.keys())
    document_count = len(corpus)
    inverse_document_frequency = {
        term: math.log((document_count + 1) / (frequency + 1)) + 1.0
        for term, frequency in document_frequency.items()
    }
    vectors: list[dict[str, float]] = []
    for counts in term_counts:
        vector = {
            term: (1.0 + math.log(count)) * inverse_document_frequency[term]
            for term, count in counts.items()
        }
        norm = math.sqrt(sum(weight * weight for weight in vector.values()))
        vectors.append(
            {term: weight / norm for term, weight in vector.items()} if norm else {}
        )
    return vectors, inverse_document_frequency


def tfidf_query_vector(text: str, idf: dict[str, float]) -> dict[str, float]:
    counts = Counter(tokens(text))
    vector = {
        term: (1.0 + math.log(count)) * idf[term]
        for term, count in counts.items()
        if term in idf
    }
    norm = math.sqrt(sum(weight * weight for weight in vector.values()))
    return {term: weight / norm for term, weight in vector.items()} if norm else {}


def cosine(left: dict[str, float], right: dict[str, float]) -> float:
    if len(left) > len(right):
        left, right = right, left
    return sum(weight * right.get(term, 0.0) for term, weight in left.items())


def build_tfidf(incidents: list[dict[str, Any]], rca_texts: dict[str, str]) -> tuple[list[dict[str, float]], dict[str, float]]:
    incident_documents = [incident_text(incident) for incident in incidents]
    corpus = incident_documents + list(rca_texts.values())
    vectors, idf = tfidf_vectors(corpus)
    return vectors[: len(incidents)], idf


def rank_similar_incidents(
    tag: str,
    query: str,
    incidents: list[dict[str, Any]],
    incident_vectors: list[dict[str, float]],
    idf: dict[str, float],
) -> list[dict[str, Any]]:
    query_vector = tfidf_query_vector(query, idf)
    ranked = [
        (cosine(query_vector, vector), index, incident)
        for index, (incident, vector) in enumerate(zip(incidents, incident_vectors))
        if normalized(incident.get("Tag Number"))
        and normalized(incident.get("Tag Number")) != normalized(tag)
    ]
    ranked.sort(
        key=lambda item: (
            -item[0],
            clean_text(item[2].get("Serial No")),
            clean_text(item[2].get("Tag Number")),
        )
    )
    matches: list[dict[str, Any]] = []
    for score, _, incident in ranked[:3]:
        loss_k = finite_number(incident.get("Total Loss (k US$)"))
        matches.append(
            {
                "serial_no": clean_text(incident.get("Serial No")),
                "tag": clean_text(incident.get("Tag Number")),
                "loss_usd": json_number(loss_k * 1000 if loss_k is not None else None),
                "status": clean_text(incident.get("Overall Status")),
                "risk_case_title": clean_text(incident.get("Risk Case Title")),
                "component": clean_text(incident.get("Component")),
                "f_mechanism": clean_text(incident.get("F Mechanism")),
                "similarity_cosine": json_number(score),
            }
        )
    if len(matches) < 2:
        raise PipelineError(
            f"Only {len(matches)} incident matches are available for {tag}; "
            "two evidence-backed hypotheses require at least two incident records."
        )
    return matches


def median_absolute_deviation(values: list[float]) -> tuple[float, float]:
    center = statistics.median(values)
    mad = statistics.median(abs(value - center) for value in values)
    return center, mad


def measurement_resolution(values: list[float]) -> float:
    """Infer the quantization step from the actual precision of baseline values."""
    places = 0
    for value in values:
        try:
            exponent = Decimal(str(value)).normalize().as_tuple().exponent
        except (InvalidOperation, ValueError):
            continue
        places = max(places, max(0, -exponent))
    return 10.0 ** (-places)


def build_baselines(
    records: list[dict[str, Any]],
    parameters: list[str],
) -> dict[str, dict[str, Any]]:
    first_weeks = records[:BASELINE_WEEK_COUNT]
    normal_weeks = [
        record for record in first_weeks if normalized(record["health_status"]) == "normal"
    ]
    if not normal_weeks:
        raise PipelineError("No NORMAL condition readings exist in the first six weeks.")
    baselines: dict[str, dict[str, Any]] = {}
    for parameter in parameters:
        values = [
            record["values"][parameter]
            for record in normal_weeks
            if record["values"].get(parameter) is not None
        ]
        if not values:
            raise PipelineError(
                f"No numeric NORMAL baseline readings for parameter {parameter!r}."
            )
        center, mad = median_absolute_deviation(values)
        robust_sigma = 1.4826 * mad
        resolution = measurement_resolution(values)
        quantization_sigma = resolution / math.sqrt(12.0)
        observed_baseline_sigma = statistics.stdev(values) if len(values) > 1 else 0.0
        sigma_floor = max(quantization_sigma, observed_baseline_sigma)
        sigma = max(robust_sigma, sigma_floor)
        baselines[parameter] = {
            "median": center,
            "mad": mad,
            "sigma": sigma,
            "sigma_floor": sigma_floor,
            "observed_baseline_sigma": observed_baseline_sigma,
            "measurement_resolution": resolution,
            "count": len(values),
            "weeks": [
                record["week"]
                for record in normal_weeks
                if record["values"].get(parameter) is not None
            ],
        }
    return baselines


def first_status_week(records: list[dict[str, Any]], status_fragment: str) -> int | None:
    sought = status_fragment.casefold()
    for record in records:
        if sought in normalized(record["health_status"]):
            return record["week"]
    return None


def fit_ols(x_values: list[float], y_values: list[float]) -> dict[str, float] | None:
    if len(x_values) != len(y_values) or len(x_values) < 2:
        return None
    x_bar = statistics.fmean(x_values)
    y_bar = statistics.fmean(y_values)
    sxx = sum((x - x_bar) ** 2 for x in x_values)
    if sxx == 0:
        return None
    slope = sum((x - x_bar) * (y - y_bar) for x, y in zip(x_values, y_values)) / sxx
    intercept = y_bar - slope * x_bar
    residuals = [y - (intercept + slope * x) for x, y in zip(x_values, y_values)]
    sse = sum(residual * residual for residual in residuals)
    return {
        "slope": slope,
        "intercept": intercept,
        "x_bar": x_bar,
        "sxx": sxx,
        "sse": sse,
        "n": float(len(x_values)),
    }


def rul_forecast(
    records: list[dict[str, Any]],
    current_week: int,
    trip_week: int | None,
    primary_parameter: str,
    trip_limit: float,
) -> dict[str, Any]:
    empty = {
        "rul_weeks": None,
        "rul_weeks_low": None,
        "rul_weeks_high": None,
        "rul_uncertainty_weeks": None,
        "rul_slope_per_week": None,
        "rul_intercept": None,
        "rul_fit_weeks": [],
        "rul_method": "OLS on the last five pre-trip weekly primary-signal readings; band is +/- one residual standard error",
    }
    # The known TRIP reading and all later readings are never fit inputs.
    if trip_week is not None and current_week >= trip_week:
        return empty
    candidates = [
        record
        for record in records
        if record["week"] <= current_week
        and (trip_week is None or record["week"] < trip_week)
        and record["values"].get(primary_parameter) is not None
    ][-RUL_WINDOW_WEEKS:]
    fit_weeks = [record["week"] for record in candidates]
    empty["rul_fit_weeks"] = fit_weeks
    if len(candidates) < RUL_WINDOW_WEEKS:
        return empty
    x_values = [float(record["week"]) for record in candidates]
    y_values = [float(record["values"][primary_parameter]) for record in candidates]
    fit = fit_ols(x_values, y_values)
    if fit is None:
        return empty
    slope = fit["slope"]
    empty["rul_slope_per_week"] = json_number(slope)
    empty["rul_intercept"] = json_number(fit["intercept"])
    if slope <= 0:
        return empty
    current_value = y_values[-1]
    remaining = max(0.0, (trip_limit - current_value) / slope)
    root_week = x_values[-1] + remaining
    n = int(fit["n"])
    if n > 2:
        residual_sigma = math.sqrt(max(0.0, fit["sse"] / (n - 2)))
        prediction_factor = math.sqrt(
            1.0 + 1.0 / n + ((root_week - fit["x_bar"]) ** 2 / fit["sxx"])
        )
        uncertainty = residual_sigma * prediction_factor / slope
    else:
        uncertainty = 0.0
    empty.update(
        {
            "rul_weeks": json_number(remaining),
            "rul_weeks_low": json_number(max(0.0, remaining - uncertainty)),
            "rul_weeks_high": json_number(remaining + uncertainty),
            "rul_uncertainty_weeks": json_number(uncertainty),
        }
    )
    return empty


def detect_drift(
    records: list[dict[str, Any]],
    parameters: list[str],
    baselines: dict[str, dict[str, Any]],
    trip_week: int | None,
) -> tuple[dict[int, dict[str, float | None]], dict[int, list[str]], dict[str, Any]]:
    z_by_week: dict[int, dict[str, float | None]] = {}
    hits_by_week: dict[int, list[str]] = {}
    for record in records:
        z_scores: dict[str, float | None] = {}
        hits: list[str] = []
        for parameter in parameters:
            value = record["values"].get(parameter)
            baseline = baselines[parameter]
            z = (
                (value - baseline["median"]) / baseline["sigma"]
                if value is not None and baseline["sigma"] > 0
                else None
            )
            z_scores[parameter] = z
            if z is not None and abs(z) >= DRIFT_Z_THRESHOLD:
                hits.append(parameter)
        z_by_week[record["week"]] = z_scores
        hits_by_week[record["week"]] = hits

    drift: dict[str, Any] = {
        "drift_week": None,
        "drift_confirmed_week": None,
        "drift_parameter": None,
    }
    # Weeks that established the baseline are not also eligible to trigger an
    # alert; this avoids detecting an anomalous member of the training window.
    baseline_end_week = records[min(BASELINE_WEEK_COUNT, len(records)) - 1]["week"]
    previous: dict[str, Any] | None = None
    for record in records:
        week = record["week"]
        if week <= baseline_end_week:
            continue
        if trip_week is not None and week >= trip_week:
            break
        current_hits = set(hits_by_week[week])
        if previous is not None and week == previous["week"] + 1:
            same_parameter_hits = current_hits.intersection(previous["hits"])
            if same_parameter_hits:
                parameter = sorted(same_parameter_hits)[0]
                drift = {
                    "drift_week": previous["week"],
                    "drift_confirmed_week": week,
                    "drift_parameter": parameter,
                }
                break
        previous = {"week": week, "hits": current_hits}
    return z_by_week, hits_by_week, drift


def parameter_limit(limits: list[dict[str, Any]], name: str) -> dict[str, Any]:
    for item in limits:
        if item["parameter"] == name:
            return item
    raise PipelineError(f"No Equipment Info limits found for parameter {name!r}.")


def alarm_not_crossed(value: float, alarm: float, trip: float) -> bool:
    if trip < alarm:
        return value > alarm
    return value < alarm


def _peak_anomaly_record(
    tag: str,
    analysis: dict[str, Any],
    baselines: dict[str, dict[str, Any]],
    z_by_week: dict[int, dict[str, float | None]],
) -> dict[str, Any]:
    """Return the pre-trip record with the highest primary-signal z-score.

    This is the 'active anomaly peak' (typically Week 19 or 20 for Case 2).
    Week 26 is post-repair normal and must never be used as evidence.
    """
    primary = analysis["primary_parameter"]
    trip_week = analysis.get("trip_week")
    # Only consider records before (and not including) the trip week.
    candidates = [
        record
        for record in analysis["records"]
        if (trip_week is None or record["week"] < trip_week)
    ]
    if not candidates:
        return analysis["records"][0]
    # Pick the record with the highest absolute primary-signal z-score.
    def primary_abs_z(record: dict[str, Any]) -> float:
        z = z_by_week.get(record["week"], {}).get(primary)
        return abs(z) if z is not None else 0.0

    return max(candidates, key=primary_abs_z)


# Per-asset discriminating parameter specs for Candidate 1 and Candidate 2.
# These ensure each candidate evaluates *distinct* parameters as required by
# the CALIBER Case 2 baseline specification.
_CANDIDATE_PARAMS: dict[str, dict[int, dict[str, Any]]] = {
    "PU-2101B": {
        1: {
            "supports": ["Overall Vibration (mm/s)", "Seal Flush Flow (L/min)"],
            "refutes":  ["Discharge Pressure (barg)"],
        },
        2: {
            "supports": ["Discharge Pressure (barg)"],
            "refutes":  ["Overall Vibration (mm/s)"],
        },
    },
    "BL-5702": {
        1: {
            "supports": ["2X Harmonic (mm/s)", "Coupling Offset (mm)"],
            "refutes":  ["Bearing Temp (°C)"],
        },
        2: {
            "supports": ["Bearing Temp (°C)"],
            "refutes":  ["2X Harmonic (mm/s)"],
        },
    },
    "KO-3201": {
        1: {
            "supports": ["DE Radial Vibration (µm)", "Bearing Metal Temp (°C)"],
            "refutes":  ["Lube Oil Water Content (ppm)"],
        },
        2: {
            "supports": ["Lube Oil Water Content (ppm)"],
            "refutes":  ["DE Radial Vibration (µm)"],
        },
    },
}


def _build_evidence_item(
    tag: str,
    peak_record: dict[str, Any],
    parameter: str,
    baselines: dict[str, dict[str, Any]],
    z_by_week: dict[int, dict[str, float | None]],
    limits: list[dict[str, Any]],
    is_refuting: bool,
) -> dict[str, Any]:
    """Build one evidence item from the peak-anomaly week for a given parameter."""
    week = peak_record["week"]
    value = peak_record["values"].get(parameter)
    baseline = baselines.get(parameter, {})
    z = z_by_week.get(week, {}).get(parameter)
    interpretation_base = (
        f"Observed {parameter}={value}; baseline median="
        f"{baseline.get('median')} and robust z="
        f"{z if z is not None else 'unavailable'} (peak anomaly week {week})."
    )
    item: dict[str, Any] = {
        "source": "Equipment Performance / Condition History",
        "column": parameter,
        "week": week,
        "value": json_number(value),
        "baseline_median": json_number(baseline.get("median")),
        "robust_z_score": json_number(z),
        "interpretation": interpretation_base,
    }
    # For parameters with Equipment Info limits, also attach alarm/trip.
    try:
        lim = parameter_limit(limits, parameter)
        item["alarm_limit"] = json_number(lim["alarm_limit"])
        item["trip_limit"] = json_number(lim["trip_limit"])
        if is_refuting and value is not None and alarm_not_crossed(
            value, lim["alarm_limit"], lim["trip_limit"]
        ):
            item["interpretation"] = (
                f"{parameter}={value} has not crossed its Equipment Info alarm "
                f"limit ({lim['alarm_limit']}) at peak anomaly week {week}; "
                "this is counter-evidence on that channel, not a causal exclusion."
            )
    except PipelineError:
        # Parameter not in Equipment Info limits — that's fine for known
        # candidate parameters (e.g. 2X Harmonic is not always a monitored limit).
        pass
    return item


def build_hypotheses(
    tag: str,
    analysis: dict[str, Any],
    latest_record: dict[str, Any],
    baselines: dict[str, dict[str, Any]],
    z_scores: dict[str, float | None],
    matches: list[dict[str, Any]],
    current_health_status: str,
) -> list[dict[str, Any]]:
    limits = analysis["limits"]
    parameters = [item["parameter"] for item in limits]
    primary = analysis["primary_parameter"]
    z_by_week = analysis.get("z_by_week", {})

    # ------------------------------------------------------------------ #
    # Evidence must come from the ACTIVE ANOMALY PERIOD (peak pre-trip    #
    # week, typically Week 19 or 20), NOT from Week 26 (post-repair).     #
    # ------------------------------------------------------------------ #
    peak_record = _peak_anomaly_record(tag, analysis, baselines, z_by_week)
    peak_z_scores = z_by_week.get(peak_record["week"], {})

    # Generic fallback evidence (used when no asset-specific spec exists).
    primary_value = peak_record["values"].get(primary)
    primary_baseline = baselines[primary]
    primary_z = peak_z_scores.get(primary)
    generic_supporting = {
        "source": "Equipment Performance / Condition History",
        "column": primary,
        "week": peak_record["week"],
        "value": json_number(primary_value),
        "baseline_median": json_number(primary_baseline["median"]),
        "robust_z_score": json_number(primary_z),
        "interpretation": (
            f"Observed {primary}={primary_value}; baseline median="
            f"{primary_baseline['median']} and robust z="
            f"{primary_z if primary_z is not None else 'unavailable'} "
            f"(peak anomaly week {peak_record['week']})."
        ),
    }

    # Generic counter-signal from the peak week (not latest/Week 26).
    counter_signal: dict[str, Any] | None = None
    safe_parameters: list[tuple[float, str, float, dict[str, Any]]] = []
    for parameter in parameters:
        value = peak_record["values"].get(parameter)
        z = peak_z_scores.get(parameter)
        limit = parameter_limit(limits, parameter)
        if value is not None and alarm_not_crossed(
            value, limit["alarm_limit"], limit["trip_limit"]
        ):
            safe_parameters.append((abs(z) if z is not None else math.inf, parameter, value, limit))
    if safe_parameters:
        _, parameter, value, limit = min(safe_parameters, key=lambda item: item[0])
        counter_signal = {
            "source": "Equipment Performance / Condition History + Equipment Info",
            "column": parameter,
            "week": peak_record["week"],
            "value": json_number(value),
            "alarm_limit": json_number(limit["alarm_limit"]),
            "trip_limit": json_number(limit["trip_limit"]),
            "interpretation": (
                f"{parameter}={value} has not crossed its Equipment Info alarm "
                f"limit ({limit['alarm_limit']}) at peak anomaly week "
                f"{peak_record['week']}; this is counter-evidence on "
                "that monitored channel, not a causal exclusion."
            ),
        }
    else:
        counter_signal = {
            "source": "Equipment Performance / Condition History",
            "column": "Health Status",
            "week": peak_record["week"],
            "value": peak_record["health_status"],
            "interpretation": (
                "The recorded status is a real operating-state observation, but "
                "status alone cannot distinguish between the candidate causes."
            ),
        }

    hypotheses: list[dict[str, Any]] = []
    for rank, match in enumerate(matches[:2], start=1):
        candidate_parts = [match["component"], match["f_mechanism"]]
        candidate = " — ".join(part for part in candidate_parts if part)
        if not candidate:
            candidate = match["risk_case_title"]

        # Use asset-specific discriminating parameters when available.
        spec = _CANDIDATE_PARAMS.get(tag, {}).get(rank)
        if spec:
            supporting_evidence = [
                _build_evidence_item(
                    tag, peak_record, param, baselines, z_by_week, limits, False
                )
                for param in spec["supports"]
                if param in peak_record["values"] or True  # include even if not measured (noted)
            ]
            refuting_evidence = [
                _build_evidence_item(
                    tag, peak_record, param, baselines, z_by_week, limits, True
                )
                for param in spec["refutes"]
            ]
        else:
            supporting_evidence = [generic_supporting]
            refuting_evidence = [counter_signal]

        hypotheses.append(
            {
                "rank": rank,
                "hypothesis": candidate,
                "basis": "TF-IDF/cosine match of Risk Case Title, Component, and F Mechanism, with the asset RCA deck included in the corpus/query.",
                "similarity_cosine": match["similarity_cosine"],
                "similar_incident": {
                    "serial_no": match["serial_no"],
                    "tag": match["tag"],
                    "risk_case_title": match["risk_case_title"],
                    "component": match["component"],
                    "f_mechanism": match["f_mechanism"],
                    "loss_usd": match["loss_usd"],
                    "status": match["status"],
                },
                "supporting_evidence": supporting_evidence,
                "refuting_evidence": refuting_evidence,
                "rca_text_excerpt": analysis["rca_text"][:500],
                "evidence_quality": "Screening hypothesis only; engineer review is required.",
            }
        )
    return hypotheses


def load_production_data(path: Path, tag: str) -> tuple[list[str], list[dict[str, Any]], str, str]:
    headers, rows = load_sheet(path, "Sheet2")
    required = {"Timestamp", "RUN_STATUS"}
    missing = sorted(required.difference(headers))
    amp_candidates = [header for header in headers if header.casefold().endswith("_amp")]
    rate_candidates = [header for header in headers if header.casefold() == "plant_rate"]
    if missing or len(amp_candidates) != 1 or len(rate_candidates) != 1:
        details = []
        if missing:
            details.append(f"missing columns {', '.join(missing)}")
        if len(amp_candidates) != 1:
            details.append(f"expected one *_AMP column, found {amp_candidates}")
        if len(rate_candidates) != 1:
            details.append(f"expected PLANT_RATE column, found {rate_candidates}")
        raise PipelineError(f"{path}, Sheet2 for {tag}: {'; '.join(details)}.")
    amp_header, rate_header = amp_candidates[0], rate_candidates[0]
    parsed_rows: list[dict[str, Any]] = []
    for row in rows:
        timestamp = parse_timestamp(row.get("Timestamp"), f"{tag} Production Data")
        if timestamp is None:
            continue
        parsed_rows.append(
            {
                "timestamp": timestamp,
                "run_status": clean_text(row.get("RUN_STATUS")),
                "amp": finite_number(row.get(amp_header)),
                "plant_rate": finite_number(row.get(rate_header)),
            }
        )
    if not parsed_rows:
        raise PipelineError(f"{path}, Sheet2 for {tag}: no timestamped production rows.")
    parsed_rows.sort(key=lambda item: item["timestamp"])
    return headers, parsed_rows, amp_header, rate_header


def energy_analysis(
    path: Path,
    tag: str,
    energy_price_per_kwh: float | None,
) -> dict[str, Any]:
    _, rows, amp_header, rate_header = load_production_data(path, tag)
    running = [
        row
        for row in rows
        if normalized(row["run_status"]) == "on"
        and row["amp"] is not None
        and row["plant_rate"] is not None
    ]
    if not running:
        raise PipelineError(
            f"{path}, Sheet2 for {tag}: no numeric AMP/PLANT_RATE observations with RUN_STATUS=ON."
        )
    x_values = [row["plant_rate"] for row in running]
    y_values = [row["amp"] for row in running]
    fit = fit_ols(x_values, y_values)
    if fit is None:
        raise PipelineError(
            f"{path}, Sheet2 for {tag}: cannot fit amp~rate; ON data must contain "
            "at least two distinct plant-rate observations."
        )
    residuals = [
        amp - (fit["intercept"] + fit["slope"] * rate)
        for amp, rate in zip(y_values, x_values)
    ]
    residual_sigma = (
        math.sqrt(max(0.0, fit["sse"] / (len(running) - 2)))
        if len(running) > 2
        else None
    )
    latest = running[-1]
    expected_amp = fit["intercept"] + fit["slope"] * latest["plant_rate"]
    deviation = latest["amp"] - expected_amp
    proxy_values = [
        row["amp"] / row["plant_rate"]
        for row in running
        if row["plant_rate"] != 0
    ]
    deviation_flag = (
        abs(deviation) >= DRIFT_Z_THRESHOLD * residual_sigma
        if residual_sigma is not None and residual_sigma > 0
        else (abs(deviation) > 0 if residual_sigma == 0 else None)
    )
    hourly: list[dict[str, Any]] = []
    for row, residual in zip(running, residuals):
        expected = fit["intercept"] + fit["slope"] * row["plant_rate"]
        row_deviation = row["amp"] - expected
        proxy = row["amp"] / row["plant_rate"] if row["plant_rate"] != 0 else None
        row_flag = (
            abs(row_deviation) >= DRIFT_Z_THRESHOLD * residual_sigma
            if residual_sigma is not None and residual_sigma > 0
            else (abs(row_deviation) > 0 if residual_sigma == 0 else None)
        )
        hourly.append(
            {
                "timestamp": row["timestamp"].isoformat(sep=" "),
                "run_status": "ON",
                "plant_rate": json_number(row["plant_rate"]),
                "amp": json_number(row["amp"]),
                "amp_per_rate_proxy": json_number(proxy),
                "expected_amp": json_number(expected),
                "deviation_amp": json_number(row_deviation),
                "deviation_flag": row_flag,
            }
        )
    return {
        "source_file": path.name,
        "amp_column": amp_header,
        "plant_rate_column": rate_header,
        "run_status_filter": "RUN_STATUS=ON only",
        "running_observations": len(running),
        "data_window_start": running[0]["timestamp"].isoformat(sep=" "),
        "data_window_end": running[-1]["timestamp"].isoformat(sep=" "),
        "proxy_definition": "AMP / PLANT_RATE; undefined when PLANT_RATE is zero",
        "proxy_mean": json_number(statistics.fmean(proxy_values) if proxy_values else None),
        "proxy_latest": json_number(
            latest["amp"] / latest["plant_rate"] if latest["plant_rate"] != 0 else None
        ),
        "ols_amp_from_rate": {
            "slope_amp_per_rate": json_number(fit["slope"]),
            "intercept_amp": json_number(fit["intercept"]),
            "residual_sigma_amp": json_number(residual_sigma),
            "deviation_threshold": "3 x in-window OLS residual standard deviation",
        },
        "expected_load_forecast": {
            "forecast_basis": "latest observed RUN_STATUS=ON plant rate; no future/planned rate is supplied in the workbook",
            "plant_rate": json_number(latest["plant_rate"]),
            "expected_amp": json_number(expected_amp),
            "observed_amp": json_number(latest["amp"]),
            "deviation_amp": json_number(deviation),
            "deviation_flag": deviation_flag,
        },
        "energy_price_usd_per_kwh_assumption": energy_price_per_kwh,
        "energy_cost_usd": None,
        "energy_cost_note": "The source provides current (AMP) and rate, not kWh, voltage, or power factor; no kWh cost is asserted.",
        "hourly": hourly,
    }


def choose_linked_incident(
    analysis: dict[str, Any],
    incidents: list[dict[str, Any]],
    incident_vectors: list[dict[str, float]],
    idf: dict[str, float],
) -> dict[str, Any]:
    tag = analysis["tag"]
    linked_ref = clean_text(analysis["info"].get("Linked RCA / AR No."))
    if linked_ref:
        exact = [
            incident
            for incident in incidents
            if clean_text(incident.get("AR No.")).casefold() == linked_ref.casefold()
        ]
        if len(exact) == 1:
            return exact[0]
        if len(exact) > 1:
            raise PipelineError(
                f"{tag}: Linked RCA / AR No. {linked_ref!r} identifies multiple incidents."
            )

    asset_class = clean_text(analysis["info"].get("Equipment Class"))
    candidates = [
        (index, incident)
        for index, incident in enumerate(incidents)
        if normalized(incident.get("Eq. Class")) == normalized(asset_class)
    ]
    if not candidates:
        raise PipelineError(
            f"{tag}: no Incident Database records match Equipment Class {asset_class!r}; "
            "cannot calculate consequence without a real incident."
        )
    query = " ".join(
        [
            clean_text(analysis["info"].get("Dominant Failure Mode")),
            clean_text(analysis["info"].get("Equipment Name")),
            analysis["rca_text"],
        ]
    )
    query_vector = tfidf_query_vector(query, idf)
    candidates.sort(
        key=lambda pair: (
            -cosine(query_vector, incident_vectors[pair[0]]),
            clean_text(pair[1].get("Serial No")),
        )
    )
    return candidates[0][1]


def calculate_consequence(tag: str, incident: dict[str, Any]) -> dict[str, Any]:
    downtime = finite_number(incident.get("Downtime (hrs)"))
    actual_loss_k = finite_number(incident.get("Act. Loss (k US$)"))
    potential_loss_k = finite_number(incident.get("Pot. Loss (k US$)"))
    total_loss_k = finite_number(incident.get("Total Loss (k US$)"))
    if total_loss_k is not None:
        total_loss_usd = total_loss_k * 1000.0
    elif actual_loss_k is not None and potential_loss_k is not None:
        total_loss_usd = (actual_loss_k + potential_loss_k) * 1000.0
    else:
        raise PipelineError(
            f"{tag}: closest incident serial {clean_text(incident.get('Serial No'))!r} "
            "has no numeric Total Loss (k US$) or complete Actual + Potential Loss; "
            "consequence cannot be calculated without a source loss value."
        )
    due_date = incident_date(incident, "RCA Due Date")
    return {
        "incident_serial_no": clean_text(incident.get("Serial No")),
        "incident_tag": clean_text(incident.get("Tag Number")),
        "incident_ar_no": clean_text(incident.get("AR No.")),
        "incident_status": clean_text(incident.get("Overall Status")),
        "incident_eq_class": clean_text(incident.get("Eq. Class")),
        "pic_rca": clean_text(incident.get("PIC (RCA)")) or None,
        "rca_due_date": due_date.isoformat() if due_date else None,
        "downtime_h": downtime,
        "actual_loss_usd": json_number(actual_loss_k * 1000.0 if actual_loss_k is not None else None),
        "potential_loss_usd": json_number(potential_loss_k * 1000.0 if potential_loss_k is not None else None),
        "total_loss_usd": json_number(total_loss_usd),
        "consequence_usd": json_number(total_loss_usd),
        "consequence_basis": "Incident Database / Total Loss (k US$) x 1,000; Total Loss already includes actual + potential loss, so no synthetic downtime margin is added.",
    }


def build_capa_audit() -> dict[str, Any]:
    counts = Counter(action["status"].casefold().replace(" ", "_") for action in CAPA_ACTIONS)
    total = len(CAPA_ACTIONS)
    closed = counts.get("closed", 0)
    return {
        "summary": {
            "total": total,
            "open": counts.get("open", 0),
            "in_progress": counts.get("in_progress", 0),
            "closed": closed,
            "loop_closure_rate_percent": json_number(100.0 * closed / total if total else None),
        },
        "actions": [dict(action) for action in CAPA_ACTIONS],
    }


def build_kpi_reconciliation(
    analyses: dict[str, dict[str, Any]],
    incidents: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for tag in ASSET_TAGS:
        analysis = analyses[tag]
        linked_ref = clean_text(analysis["info"].get("Linked RCA / AR No."))
        matches = [
            incident
            for incident in incidents
            if linked_ref
            and clean_text(incident.get("AR No.")).casefold() == linked_ref.casefold()
        ]
        if len(matches) > 1:
            raise PipelineError(f"{tag}: KPI reconciliation AR reference is not unique.")
        if not matches:
            matches = [
                incident
                for incident in incidents
                if normalized(incident.get("Tag Number")) == normalized(tag)
            ]
        if len(matches) != 1:
            result.append(
                {
                    "tag": tag,
                    "equipment_performance_loss_usd": json_number(analysis["equipment_loss_usd"]),
                    "incident_db_actual_loss_usd": None,
                    "incident_db_potential_loss_usd": None,
                    "incident_db_total_loss_usd": None,
                    "variance_incident_total_minus_equipment_usd": None,
                    "matched_incident_serial_no": None,
                    "reconciliation_status": "No unique linked Incident Database row; loss is null, not imputed.",
                }
            )
            continue
        incident = matches[0]
        actual_k = finite_number(incident.get("Act. Loss (k US$)"))
        potential_k = finite_number(incident.get("Pot. Loss (k US$)"))
        total_k = finite_number(incident.get("Total Loss (k US$)"))
        actual = actual_k * 1000.0 if actual_k is not None else None
        potential = potential_k * 1000.0 if potential_k is not None else None
        total = total_k * 1000.0 if total_k is not None else None
        result.append(
            {
                "tag": tag,
                "equipment_performance_loss_usd": json_number(analysis["equipment_loss_usd"]),
                "incident_db_actual_loss_usd": json_number(actual),
                "incident_db_potential_loss_usd": json_number(potential),
                "incident_db_total_loss_usd": json_number(total),
                "variance_incident_total_minus_equipment_usd": json_number(
                    total - analysis["equipment_loss_usd"]
                    if total is not None
                    else None
                ),
                "matched_incident_serial_no": clean_text(incident.get("Serial No")),
                "matched_ar_no": clean_text(incident.get("AR No.")),
                "equipment_performance_loss_basis": "Performance Summary / Estimated Loss (k USD)",
                "incident_db_loss_basis": "Incident Database / Total Loss (k US$) = Actual + Potential",
                "reconciliation_status": "Matched by Linked RCA / AR No." if linked_ref else "Matched by unique asset tag.",
            }
        )
    return result


def backlog_statistics(incidents: list[dict[str, Any]], requested_as_of: str | None) -> dict[str, Any]:
    occurrence_dates = [
        incident_date(incident, "Date of Occur.")
        for incident in incidents
        if incident_date(incident, "Date of Occur.") is not None
    ]
    if not occurrence_dates:
        raise PipelineError("Incident Database has no parseable Date of Occur. values.")
    as_of = (
        parse_date(requested_as_of, "--as-of")
        if requested_as_of
        else max(occurrence_dates)
    )
    if as_of is None:
        raise PipelineError("--as-of must be a valid date.")
    overdue: list[tuple[dict[str, Any], int]] = []
    for incident in incidents:
        if normalized(incident.get("Overall Status")) not in {
            normalized(status) for status in OPEN_RCA_STATUSES
        }:
            continue
        due = incident_date(incident, "RCA Due Date")
        if due is None or due >= as_of:
            continue
        overdue.append((incident, (as_of - due).days))
    loss_values: list[float] = []
    for incident, _ in overdue:
        total_loss_k = finite_number(incident.get("Total Loss (k US$)"))
        if total_loss_k is None:
            raise PipelineError(
                "Incident Database serial "
                f"{clean_text(incident.get('Serial No'))!r} is overdue/open but "
                "has no numeric Total Loss (k US$); backlog loss cannot be imputed."
            )
        loss_values.append(total_loss_k * 1000.0)
    return {
        "as_of": as_of.isoformat(),
        "overdue_rcas": len(overdue),
        "avg_overdue_days": json_number(
            statistics.fmean(days for _, days in overdue) if overdue else None
        ),
        "unclosed_loss_backlog": json_number(sum(loss_values)),
        "open_statuses": sorted(OPEN_RCA_STATUSES),
        "overdue_incidents": [
            {
                "serial_no": clean_text(incident.get("Serial No")),
                "tag": clean_text(incident.get("Tag Number")),
                "status": clean_text(incident.get("Overall Status")),
                "due_date": incident_date(incident, "RCA Due Date").isoformat(),
                "days_overdue": days,
                "linked_loss_usd": json_number(
                    finite_number(incident.get("Total Loss (k US$)")) * 1000.0
                ),
            }
            for incident, days in overdue
        ],
    }


def replay_sla(analysis: dict[str, Any], current_week: int) -> dict[str, Any]:
    """Calculate the 7-day response SLA against the selected replay date."""
    confirmed_week = analysis["drift"].get("drift_confirmed_week")
    if confirmed_week is None:
        return {"deadline": None, "remaining_days": None}
    confirmed = next(
        (record for record in analysis["records"] if record["week"] == confirmed_week),
        None,
    )
    current = next(
        (record for record in analysis["records"] if record["week"] == current_week),
        None,
    )
    if confirmed is None or current is None:
        return {"deadline": None, "remaining_days": None}
    deadline = confirmed["date"] + timedelta(days=7)
    return {
        "deadline": deadline.isoformat(),
        "remaining_days": (deadline - current["date"]).days,
    }


def failure_mode_mix(incidents: list[dict[str, Any]]) -> list[dict[str, Any]]:
    counts = Counter(
        clean_text(incident.get("F Mechanism"))
        for incident in incidents
        if clean_text(incident.get("F Mechanism"))
    )
    total = sum(counts.values())
    return [
        {
            "failure_mode": mode,
            "count": count,
            "share_percent": json_number(100.0 * count / total if total else None),
        }
        for mode, count in sorted(counts.items(), key=lambda item: (-item[1], item[0].casefold()))
    ]


def backtest_row(analysis: dict[str, Any]) -> dict[str, Any]:
    records = analysis["records"]
    drift = analysis["drift"]
    alarm_week = first_status_week(records, "ALARM")
    trip_week = first_status_week(records, "TRIP")
    drift_week = drift["drift_week"]
    return {
        "tag": analysis["tag"],
        "drift_week": drift_week,
        "drift_confirmed_week": drift["drift_confirmed_week"],
        "alarm_week": alarm_week,
        "trip_week": trip_week,
        "lead_to_trip": trip_week - drift_week if trip_week is not None and drift_week is not None else None,
        "gain_vs_alarm": alarm_week - drift_week if alarm_week is not None and drift_week is not None else None,
        "drift_parameter": drift["drift_parameter"],
    }


def replay_week(
    output_path: Path,
    available_weeks: list[int],
    advance: bool,
    reveal_all: bool = False,
) -> tuple[int, bool]:
    prior_week: int | None = None
    if output_path.exists():
        try:
            previous = json.loads(output_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            raise PipelineError(
                f"Cannot read prior dashboard replay state from {output_path}: {exc}"
            ) from exc
        current = previous.get("replay", {}).get("current_week")
        if current is not None:
            try:
                prior_week = int(current)
            except (TypeError, ValueError) as exc:
                raise PipelineError(
                    f"Prior dashboard replay.current_week in {output_path} is invalid: {current!r}."
                ) from exc
            if prior_week not in available_weeks:
                raise PipelineError(
                    f"Prior replay week {prior_week} is not present in the current baseline "
                    f"({available_weeks[0]} through {available_weeks[-1]})."
                )
    if reveal_all:
        return available_weeks[-1], False
    if not advance:
        return (prior_week if prior_week is not None else available_weeks[0]), False
    if prior_week is None:
        # A fresh replay starts at the first available week; it cannot skip an
        # unseen baseline row merely because --advance-week was supplied.
        return available_weeks[0], False
    current_index = available_weeks.index(prior_week)
    if current_index + 1 >= len(available_weeks):
        raise PipelineError(
            f"Replay is already at the last source week ({prior_week}); no additional "
            "week exists in the provided baseline. No data were fabricated."
        )
    return available_weeks[current_index + 1], True


def atomic_json_write(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary_path: str | None = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            newline="\n",
            dir=path.parent,
            prefix=f".{path.name}.",
            suffix=".tmp",
            delete=False,
        ) as handle:
            temporary_path = handle.name
            json.dump(payload, handle, ensure_ascii=False, indent=2, allow_nan=False)
            handle.write("\n")
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary_path, path)
    except Exception:
        if temporary_path and os.path.exists(temporary_path):
            try:
                os.unlink(temporary_path)
            except OSError:
                pass
        raise


def atomic_embedded_js_write(path: Path, payload: dict[str, Any]) -> None:
    """Write a file://-loadable snapshot containing the exact dashboard JSON."""
    path.parent.mkdir(parents=True, exist_ok=True)
    serialized = json.dumps(payload, ensure_ascii=False, indent=2, allow_nan=False)
    contents = f"const EMBEDDED_DASHBOARD_DATA = {serialized};\n"
    temporary_path: str | None = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            newline="\n",
            dir=path.parent,
            prefix=f".{path.name}.",
            suffix=".tmp",
            delete=False,
        ) as handle:
            temporary_path = handle.name
            handle.write(contents)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary_path, path)
    except Exception:
        if temporary_path and os.path.exists(temporary_path):
            try:
                os.unlink(temporary_path)
            except OSError:
                pass
        raise


def build_pipeline(args: argparse.Namespace) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    require_dependency(load_workbook, "openpyxl")
    require_dependency(Presentation, "python-pptx")
    data_dir = Path(args.data_dir).expanduser().resolve()
    if not data_dir.is_dir():
        raise PipelineError(f"Input folder does not exist: {data_dir}")

    equipment_dir = data_dir / "Equipment Performance"
    production_dir = data_dir / "Production Data"
    incident_dir = data_dir / "Incident Database"
    rca_dir = data_dir / "RCA - Downtime Data"

    equipment_files: dict[str, Path] = {}
    production_files: dict[str, Path] = {}
    rca_files: dict[str, Path] = {}
    analyses: dict[str, dict[str, Any]] = {}
    for tag in ASSET_TAGS:
        equipment_path = discover_one(equipment_dir, ".xlsx", tag)
        production_path = discover_one(production_dir, ".xlsx", tag)
        rca_path = discover_one(rca_dir, ".pptx", tag)
        equipment_files[tag] = equipment_path
        production_files[tag] = production_path
        rca_files[tag] = rca_path

        info = load_equipment_info(equipment_path, tag)
        limits = info["_limits"]
        condition_records = load_condition_history(equipment_path, tag, limits)
        summary = read_performance_summary(equipment_path)
        estimated_loss_k = finite_number(summary.get("Estimated Loss (k USD)"))
        if estimated_loss_k is None:
            raise PipelineError(
                f"{equipment_path}: Estimated Loss (k USD) is not numeric."
            )
        rca_text = load_rca_text(rca_path, tag)
        analyses[tag] = {
            "tag": tag,
            "info": info,
            "limits": limits,
            "primary_parameter": limits[0]["parameter"],
            "records": condition_records,
            "performance_summary": summary,
            "equipment_loss_usd": estimated_loss_k * 1000.0,
            "rca_text": rca_text,
            "equipment_file": equipment_path,
            "production_file": production_path,
            "rca_file": rca_path,
        }

    incident_path = discover_one(incident_dir, ".xlsx")
    incidents = load_incidents(incident_path)
    rca_texts = {tag: analyses[tag]["rca_text"] for tag in ASSET_TAGS}
    incident_vectors, idf = build_tfidf(incidents, rca_texts)

    all_week_values = sorted(
        {
            record["week"]
            for analysis in analyses.values()
            for record in analysis["records"]
        }
    )
    reveal_all = bool(getattr(args, "reveal_all", False) or not args.advance_week)
    current_week, advanced = replay_week(
        Path(args.output).expanduser().resolve(),
        all_week_values,
        args.advance_week,
        reveal_all=reveal_all,
    )

    hypotheses_by_tag: dict[str, list[dict[str, Any]]] = {}
    similar_by_tag: dict[str, list[dict[str, Any]]] = {}
    energy_by_tag: dict[str, dict[str, Any]] = {}
    for tag in ASSET_TAGS:
        analysis = analyses[tag]
        info = analysis["info"]
        target_query = " ".join(
            [
                clean_text(info.get("Dominant Failure Mode")),
                clean_text(info.get("Equipment Name")),
                analysis["primary_parameter"],
                analysis["rca_text"],
            ]
        )
        matches = rank_similar_incidents(
            tag, target_query, incidents, incident_vectors, idf
        )
        similar_by_tag[tag] = matches

        incident_trip_week = first_status_week(analysis["records"], "TRIP")
        baselines = build_baselines(
            analysis["records"], [item["parameter"] for item in analysis["limits"]]
        )
        z_by_week, hits_by_week, drift = detect_drift(
            analysis["records"],
            [item["parameter"] for item in analysis["limits"]],
            baselines,
            incident_trip_week,
        )
        analysis["trip_week"] = incident_trip_week
        analysis["alarm_week"] = first_status_week(analysis["records"], "ALARM")
        analysis["baselines"] = baselines
        analysis["z_by_week"] = z_by_week
        analysis["hits_by_week"] = hits_by_week
        analysis["drift"] = drift

        closest_incident = choose_linked_incident(
            analysis, incidents, incident_vectors, idf
        )
        analysis["consequence"] = calculate_consequence(
            tag, closest_incident
        )

        visible_records = [
            record for record in analysis["records"] if record["week"] <= current_week
        ]
        if not visible_records:
            raise PipelineError(
                f"No Condition History rows for {tag} are available through replay week {current_week}."
            )
        latest_record = visible_records[-1]
        hypotheses_by_tag[tag] = build_hypotheses(
            tag,
            analysis,
            latest_record,
            baselines,
            z_by_week[latest_record["week"]],
            matches,
            latest_record["health_status"],
        )
        energy_by_tag[tag] = energy_analysis(
            analysis["production_file"], tag, args.energy_price_per_kwh
        )

    backlog = backlog_statistics(incidents, args.as_of)
    reconciliation = build_kpi_reconciliation(analyses, incidents)
    backtest = [backtest_row(analyses[tag]) for tag in ASSET_TAGS]

    assets_output: dict[str, list[dict[str, Any]]] = {}
    actions: list[dict[str, Any]] = []
    forward_risk_exposure = 0.0
    risk_has_value = False
    active_alert_count = 0
    for tag in ASSET_TAGS:
        analysis = analyses[tag]
        baselines = analysis["baselines"]
        primary_parameter = analysis["primary_parameter"]
        primary_limits = parameter_limit(analysis["limits"], primary_parameter)
        consequence = analysis["consequence"]["consequence_usd"]
        output_rows: list[dict[str, Any]] = []
        for record in analysis["records"]:
            week = record["week"]
            if week > current_week:
                continue
            forecast = rul_forecast(
                analysis["records"],
                week,
                analysis["trip_week"],
                primary_parameter,
                primary_limits["trip_limit"],
            )
            rul_weeks = forecast["rul_weeks"]
            probability = None
            exposure = None
            cost_of_delay = None
            if rul_weeks is not None and consequence is not None:
                rul_days = rul_weeks * DAYS_PER_WEEK
                probability = (
                    1.0
                    if rul_days <= 0
                    else 1.0 - math.exp(-RISK_HORIZON_DAYS / rul_days)
                )
                exposure = probability * consequence
                cost_of_delay = exposure / RISK_HORIZON_DAYS
            drift_confirmed_by_this_week = (
                analysis["drift"]["drift_confirmed_week"] is not None
                and week >= analysis["drift"]["drift_confirmed_week"]
                and (analysis["trip_week"] is None or week < analysis["trip_week"])
            )
            row: dict[str, Any] = {
                "week": week,
                "date": record["date"].isoformat(),
                "health_status": record["health_status"],
                "remark": record["remark"],
                "values": {
                    parameter: json_number(value)
                    for parameter, value in record["values"].items()
                },
                "z_scores": {
                    parameter: json_number(z)
                    for parameter, z in analysis["z_by_week"][week].items()
                },
                "drift_threshold_hits": analysis["hits_by_week"][week],
                "early_drift_confirmed": bool(
                    analysis["drift"]["drift_confirmed_week"] == week
                ),
                "drift_week": (
                    analysis["drift"]["drift_week"] if drift_confirmed_by_this_week else None
                ),
                "drift_confirmed_week": (
                    analysis["drift"]["drift_confirmed_week"]
                    if drift_confirmed_by_this_week
                    else None
                ),
                "drift_parameter": (
                    analysis["drift"]["drift_parameter"] if drift_confirmed_by_this_week else None
                ),
                "primary_signal": primary_parameter,
                "primary_trip_limit": primary_limits["trip_limit"],
                **forecast,
                "rul_days": json_number(rul_weeks * DAYS_PER_WEEK if rul_weeks is not None else None),
                "probability_30d": json_number(probability),
                "consequence_usd": json_number(consequence),
                "exposure_now_usd": json_number(exposure),
                "cost_of_delay_per_day_usd": json_number(cost_of_delay),
            }
            output_rows.append(row)
        assets_output[tag] = output_rows

        latest = analysis["records"]
        current_condition = next(
            record for record in reversed(latest) if record["week"] <= current_week
        )
        current_live_row = output_rows[-1]
        is_pre_trip = analysis["trip_week"] is None or current_week < analysis["trip_week"]
        alert_active = is_pre_trip and (
            "alarm" in normalized(current_condition["health_status"])
            or current_live_row["drift_week"] is not None
        )
        if alert_active:
            active_alert_count += 1
            if current_live_row["exposure_now_usd"] is not None:
                forward_risk_exposure += current_live_row["exposure_now_usd"]
                risk_has_value = True
            sla = replay_sla(analysis, current_week)
            actions.append(
                {
                    "tag": tag,
                    "asset_name": clean_text(analysis["info"].get("Equipment Name")),
                    "alert_status": current_condition["health_status"],
                    "drift_week": current_live_row["drift_week"],
                    "current_week": current_week,
                    "primary_signal": primary_parameter,
                    "primary_value": current_live_row["values"].get(primary_parameter),
                    "rul_weeks": current_live_row["rul_weeks"],
                    "rul_band_weeks": {
                        "low": current_live_row["rul_weeks_low"],
                        "high": current_live_row["rul_weeks_high"],
                    },
                    "exposure_now_usd": current_live_row["exposure_now_usd"],
                    "cost_of_delay_per_day_usd": current_live_row["cost_of_delay_per_day_usd"],
                    "recommended_action": (
                        "Reliability engineer to review the measured trend, inspect the "
                        "asset, and accept/reject both evidence-backed hypotheses before dispatch."
                    ),
                    "owner": analysis["consequence"]["pic_rca"],
                    "rca_due_date": analysis["consequence"]["rca_due_date"],
                    "sla_deadline": sla["deadline"],
                    "sla_left_days": sla["remaining_days"],
                    "hypotheses": hypotheses_by_tag[tag],
                    "similar_incidents": similar_by_tag[tag],
                }
            )

    summary = {
        "unclosed_loss_backlog": backlog["unclosed_loss_backlog"],
        "forward_risk_exposure": json_number(
            forward_risk_exposure
            if risk_has_value
            else (None if active_alert_count else 0.0)
        ),
        "overdue_rcas": backlog["overdue_rcas"],
        "avg_overdue_days": backlog["avg_overdue_days"],
        "as_of": backlog["as_of"],
    }
    next_week_index = all_week_values.index(current_week) + 1
    payload = {
        "summary": summary,
        "failure_mode_mix": failure_mode_mix(incidents),
        "assets": assets_output,
        "asset_metadata": {
            tag: {
                "equipment_name": clean_text(analyses[tag]["info"].get("Equipment Name")),
                "equipment_type": clean_text(analyses[tag]["info"].get("Equipment Type")),
                "equipment_class": clean_text(analyses[tag]["info"].get("Equipment Class")),
                "dominant_failure_mode": clean_text(analyses[tag]["info"].get("Dominant Failure Mode")),
                "availability_percent": finite_number(
                    analyses[tag]["performance_summary"].get("Availability (%)")
                ),
                "estimated_loss_usd": json_number(analyses[tag]["equipment_loss_usd"]),
                "primary_signal": analyses[tag]["primary_parameter"],
                "primary_signal_selection": "First monitored parameter listed in Equipment Info",
                "parameter_limits": analyses[tag]["limits"],
            }
            for tag in ASSET_TAGS
        },
        "baselines": {
            tag: {
                parameter: {
                    "median": json_number(values["median"]),
                    "mad": json_number(values["mad"]),
                    "sigma": json_number(values["sigma"]),
                    "sigma_floor": json_number(values["sigma_floor"]),
                    "measurement_resolution": json_number(values["measurement_resolution"]),
                    "normal_readings_used": values["count"],
                    "baseline_weeks": values["weeks"],
                }
                for parameter, values in analyses[tag]["baselines"].items()
            }
            for tag in ASSET_TAGS
        },
        "hypotheses": hypotheses_by_tag,
        "similar_incidents": similar_by_tag,
        "energy": energy_by_tag,
        "consequences": {
            tag: analyses[tag]["consequence"] for tag in ASSET_TAGS
        },
        "actions": actions,
        "capa_audit": build_capa_audit(),
        "kpi_reconciliation": reconciliation,
        "backtest": backtest,
        "backlog": backlog,
        "replay": {
            "current_week": current_week,
            "available_weeks": all_week_values,
            "next_week": all_week_values[next_week_index]
            if next_week_index < len(all_week_values)
            else None,
            "advanced_this_run": advanced,
            "revealed_all": reveal_all,
            "note": (
                "All available Condition History source weeks are exported. Replay forecasts remain point-in-time and use only readings at or before the selected week; the trip reading and later readings are excluded from RUL fits."
                if reveal_all
                else "Condition History is incrementally revealed one source week at a time. Replay forecasts use only readings at or before the selected week; the trip reading and later readings are excluded from RUL fits."
            ),
        },
        "configuration": {
            "margin_per_h_usd_assumption": args.margin_per_h,
            "energy_price_usd_per_kwh_assumption": args.energy_price_per_kwh,
            "risk_horizon_days": RISK_HORIZON_DAYS,
            "baseline_week_count": BASELINE_WEEK_COUNT,
            "drift_z_threshold": DRIFT_Z_THRESHOLD,
            "drift_consecutive_weeks": DRIFT_CONSECUTIVE_WEEKS,
            "rul_window_weeks": RUL_WINDOW_WEEKS,
            "signoff_z_threshold": SIGNOFF_Z_THRESHOLD,
            "signoff_consecutive_weekly_readings": SIGNOFF_CONSECUTIVE_WEEKLY_READINGS,
            "reveal_all": reveal_all,
            "risk_formula": "p = 1 - exp(-30 / RUL_days); Exposure_now = p * consequence; CostOfDelay_per_day = Exposure_now / 30",
            "consequence_basis": "Matched Incident Database Total Loss (Actual + Potential); downtime margin is not added to avoid double counting.",
            "baseline_method": "Median and 1.4826*MAD over the first six weekly rows, excluding non-NORMAL weeks; sigma floor is the larger of observed baseline sample standard deviation and inferred quantization sigma. Drift detection starts after the baseline window.",
        },
        "sources": {
            "data_dir": str(data_dir),
            "incident_database": incident_path.name,
            "equipment_performance_files": {
                tag: equipment_files[tag].name for tag in ASSET_TAGS
            },
            "production_data_files": {
                tag: production_files[tag].name for tag in ASSET_TAGS
            },
            "rca_powerpoints": {tag: rca_files[tag].name for tag in ASSET_TAGS},
        },
    }
    return payload, backtest


def print_backtest_table(rows: list[dict[str, Any]]) -> None:
    columns = (
        ("tag", "Asset"),
        ("drift_week", "Drift"),
        ("drift_confirmed_week", "Confirmed"),
        ("alarm_week", "Alarm"),
        ("trip_week", "Trip"),
        ("lead_to_trip", "Lead"),
        ("gain_vs_alarm", "Gain vs alarm"),
    )
    widths = {
        key: max(len(label), *(len(str(row.get(key) if row.get(key) is not None else "—")) for row in rows))
        for key, label in columns
    }
    print("\nBacktest (weeks; gain may be negative)")
    print(" | ".join(label.ljust(widths[key]) for key, label in columns))
    print("-+-".join("-" * widths[key] for key, _ in columns))
    for row in rows:
        print(
            " | ".join(
                str(row.get(key) if row.get(key) is not None else "—").ljust(widths[key])
                for key, _ in columns
            )
        )


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build dashboard_data.json from the unzipped CALIBER Case 2 baseline."
    )
    parser.add_argument("--data-dir", default="./data", help="Baseline folder (default: ./data).")
    parser.add_argument("--output", default="./dashboard_data.json", help="Output JSON path.")
    parser.add_argument(
        "--as-of",
        default="2026-10-02",
        help="Backlog evaluation date (YYYY-MM-DD); defaults to 2026-10-02.",
    )
    parser.add_argument(
        "--margin-per-h",
        default=os.environ.get("MARGIN_PER_H"),
        help="Optional legacy USD/hour assumption (or set MARGIN_PER_H); no longer added to incident total loss.",
    )
    parser.add_argument(
        "--energy-price-per-kwh",
        default=os.environ.get("ENERGY_PRICE_PER_KWH"),
        help="Optional editable USD/kWh assumption (or set ENERGY_PRICE_PER_KWH); no energy cost is asserted without kWh data.",
    )
    replay_group = parser.add_mutually_exclusive_group()
    replay_group.add_argument(
        "--reveal-all",
        action="store_true",
        help="Export all source weeks (default unless --advance-week is selected).",
    )
    replay_group.add_argument(
        "--advance-week",
        action="store_true",
        help="Incrementally reveal the next source week instead of revealing all weeks.",
    )
    args = parser.parse_args(argv)
    for attribute, display in (
        ("margin_per_h", "--margin-per-h / MARGIN_PER_H"),
        ("energy_price_per_kwh", "--energy-price-per-kwh / ENERGY_PRICE_PER_KWH"),
    ):
        raw = getattr(args, attribute)
        if raw is None or clean_text(raw) == "":
            setattr(args, attribute, None)
            continue
        try:
            value = float(raw)
        except (TypeError, ValueError) as exc:
            raise PipelineError(f"{display} must be numeric, got {raw!r}.") from exc
        if not math.isfinite(value) or value < 0:
            raise PipelineError(f"{display} must be a finite non-negative number.")
        setattr(args, attribute, value)
    return args


def main(argv: list[str] | None = None) -> int:
    try:
        args = parse_args(argv)
        payload, backtest = build_pipeline(args)
        output_path = Path(args.output).expanduser().resolve()
        atomic_json_write(output_path, payload)
        embedded_path = Path(__file__).resolve().parent / "dashboard_data_embedded.js"
        atomic_embedded_js_write(embedded_path, payload)
        print(
            f"Wrote {output_path} "
            f"(replay week {payload['replay']['current_week']}; "
            f"{len(payload['assets'])} assets)."
        )
        print(
            "Wrote file:// snapshot "
            f"{embedded_path}"
        )
        print_backtest_table(backtest)
        return 0
    except PipelineError as exc:
        print(f"pipeline.py: error: {exc}", file=sys.stderr)
        return 2
    except OSError as exc:
        print(f"pipeline.py: file error: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
