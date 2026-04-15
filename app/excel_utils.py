from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Iterable

from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook


def _cell_to_str(v: Any) -> str:
    if v is None:
        return ""
    return str(v).strip()


def _safe_sheet_name(wb: Workbook, preferred: str | None = None) -> str:
    if preferred and preferred in wb.sheetnames:
        return preferred
    return wb.sheetnames[0]


def read_headers_xlsx(path: str, sheet_name: str | None = None) -> list[str]:
    wb = load_workbook(filename=path, read_only=True, data_only=True)
    ws = wb[_safe_sheet_name(wb, sheet_name)]
    headers: list[str] = []
    for cell in ws[1]:
        headers.append(_cell_to_str(cell.value))
    wb.close()
    headers = [h for h in headers if h != ""]
    return headers


@dataclass(frozen=True)
class ColumnProfile:
    header: str
    samples: list[str]


@dataclass(frozen=True)
class SourcePreview:
    headers: list[str]
    rows: list[list[str]]
    row_count: int
    profiles: list[ColumnProfile]


def read_source_preview_xlsx(
    path: str,
    sheet_name: str | None = None,
    preview_rows: int = 20,
    sample_limit_per_col: int = 50,
) -> SourcePreview:
    wb = load_workbook(filename=path, read_only=True, data_only=True)
    ws = wb[_safe_sheet_name(wb, sheet_name)]

    raw_headers = [_cell_to_str(c.value) for c in ws[1]]
    # Keep columns up to last non-empty header (common Excel pattern).
    last_idx = 0
    for i, h in enumerate(raw_headers, start=1):
        if h != "":
            last_idx = i
    headers = raw_headers[:last_idx] if last_idx else []
    headers = [h for h in headers if h != ""]

    preview: list[list[str]] = []
    samples: list[list[str]] = [[] for _ in range(len(headers))]

    row_count = 0
    for row in ws.iter_rows(min_row=2, max_col=len(headers), values_only=True):
        row_count += 1
        row_str = [_cell_to_str(v) for v in row]
        if len(preview) < preview_rows:
            preview.append(row_str)
        for ci, val in enumerate(row_str):
            if val != "" and len(samples[ci]) < sample_limit_per_col:
                samples[ci].append(val)

    wb.close()

    profiles = [ColumnProfile(header=headers[i], samples=samples[i]) for i in range(len(headers))]
    return SourcePreview(headers=headers, rows=preview, row_count=row_count, profiles=profiles)


def write_filled_template(
    template_path: str,
    source_path: str,
    mapping: dict[str, dict[str, str]],
    *,
    template_sheet_name: str | None = None,
    source_sheet_name: str | None = None,
) -> bytes:
    """
    mapping format per target header:
      { "type": "source"|"constant"|"blank", "value": <source header or constant> }
    """
    twb = load_workbook(filename=template_path, read_only=False, data_only=False)
    tws = twb[_safe_sheet_name(twb, template_sheet_name)]

    # Read template headers from row 1 (keep empty cells as empty; we only map non-empty headers).
    template_headers_full = [_cell_to_str(c.value) for c in tws[1]]
    template_headers = [h for h in template_headers_full if h != ""]

    swb = load_workbook(filename=source_path, read_only=True, data_only=True)
    sws = swb[_safe_sheet_name(swb, source_sheet_name)]

    source_headers_raw = [_cell_to_str(c.value) for c in sws[1]]
    last_idx = 0
    for i, h in enumerate(source_headers_raw, start=1):
        if h != "":
            last_idx = i
    source_headers_raw = source_headers_raw[:last_idx] if last_idx else []
    source_headers = [h for h in source_headers_raw if h != ""]

    source_index = {h: i for i, h in enumerate(source_headers)}  # 0-based within trimmed cols

    # Validate mapping covers all template headers.
    missing = [h for h in template_headers if h not in mapping]
    if missing:
        raise ValueError(f"Missing mappings for template headers: {missing}")

    # Build column order in template based on non-empty headers positions
    template_col_positions: list[tuple[int, str]] = []
    for col_idx_1based, cell in enumerate(tws[1], start=1):
        h = _cell_to_str(cell.value)
        if h != "":
            template_col_positions.append((col_idx_1based, h))

    out_row = 2
    for src_row in sws.iter_rows(min_row=2, max_col=len(source_headers), values_only=True):
        # For each template column, write deterministically.
        for col_idx_1based, th in template_col_positions:
            spec = mapping[th]
            mtype = spec.get("type", "")
            val = ""
            if mtype == "source":
                src_h = spec.get("value", "")
                if src_h not in source_index:
                    raise ValueError(f"Mapped source header not found: {src_h}")
                v = src_row[source_index[src_h]] if source_index[src_h] < len(src_row) else None
                val = v
            elif mtype == "constant":
                val = spec.get("value", "")
            elif mtype == "blank":
                val = ""
            else:
                raise ValueError(f"Invalid mapping type for '{th}': {mtype}")

            tws.cell(row=out_row, column=col_idx_1based).value = val

        out_row += 1

    # Add audit sheets
    if "Mapping" in twb.sheetnames:
        del twb["Mapping"]
    if "Errors" in twb.sheetnames:
        del twb["Errors"]

    ws_map = twb.create_sheet("Mapping")
    ws_map.append(["Template Header", "Mapping Type", "Source/Constant Value"])
    for th in template_headers:
        spec = mapping[th]
        ws_map.append([th, spec.get("type", ""), spec.get("value", "")])

    # We don't compute row-level errors yet (strict mapping gate prevents most issues),
    # but keep the sheet for future validations (types, required columns, etc.).
    ws_err = twb.create_sheet("Errors")
    ws_err.append(["Row", "Template Header", "Issue", "Value"])

    from io import BytesIO

    bio = BytesIO()
    twb.save(bio)
    bio.seek(0)
    swb.close()
    twb.close()
    return bio.read()

