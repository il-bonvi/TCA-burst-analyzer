from __future__ import annotations

from io import BytesIO
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ──────────────────────────────────────────────────────────────────────────────
# COLOR PALETTE  ←  edit these to change all colors in the export
# Format: "FFRRGGBB"  (FF = fully opaque, then hex RGB)
# ──────────────────────────────────────────────────────────────────────────────

# Background / chrome
BG_DARK   = "FF0D0F14"   # very dark title bars
BG_MID    = "FF1A1D24"   # sheet title bars
BG_HEAD   = "FF252932"   # column header rows
BG_WHITE  = "FFFFFFFF"   # pure white cell bg (used for "All Bursts" alt rows)

# Metric accent colors  (match the HTML app)
C_WATT    = "FFC084FC"   # purple — power / watt values
C_HR      = "FFEF4444"   # red    — heart rate
C_CAD     = "FF4DA6FF"   # blue   — cadence
C_TIME    = "FF3CCF7E"   # green  — time / duration
C_DELTA   = "FF3CCF7E"   # green  — delta above threshold
C_MAX     = "FF9555CC"   # deep purple — max power
C_MIN     = "FFD9A8FF"   # light lavender — min power

# Text colors
T_WHITE   = "FFFFFFFF"
T_DARK    = "FF0D0F14"
T_MUTED   = "FF6B7280"

# Border
BORDER_C  = "FFCCCCCC"
BORDER_H  = "FF444444"   # header borders (darker)


# ──────────────────────────────────────────────────────────────────────────────
# Helpers
# ──────────────────────────────────────────────────────────────────────────────

def _hex_to_argb(hex_color: str) -> str:
    return "FF" + hex_color.lstrip("#").upper()


def _lighten(hex_color: str, factor: float) -> str:
    """Blend hex_color toward white by factor (0=unchanged, 1=white)."""
    h = hex_color.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    r = round(r + (255 - r) * factor)
    g = round(g + (255 - g) * factor)
    b = round(b + (255 - b) * factor)
    return f"FF{r:02X}{g:02X}{b:02X}"


def _fill(argb: str) -> PatternFill:
    return PatternFill("solid", fgColor=argb)


def _border(color: str = BORDER_C) -> Border:
    s = Side(style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)


def _center(wrap: bool = False) -> Alignment:
    return Alignment(horizontal="center", vertical="center", wrap_text=wrap)


def _left(indent: int = 1) -> Alignment:
    return Alignment(horizontal="left", vertical="center", indent=indent)


def _font(
    color: str = T_DARK,
    bold: bool = False,
    size: int = 10,
    name: str = "Arial",
) -> Font:
    return Font(bold=bold, size=size, color=color, name=name)


def _fmt_time(seconds: float) -> str:
    total = int(round(seconds))
    h = total // 3600
    m = (total % 3600) // 60
    s = total % 60
    if h > 0:
        return f"{h}:{m:02d}:{s:02d}"
    return f"{m:02d}:{s:02d}"


def _safe_avg(values: list) -> float:
    filtered = [v for v in values if v is not None]
    return sum(filtered) / len(filtered) if filtered else 0.0


def _write_header_row(ws, row: int, columns: list[tuple[str, int]]) -> None:
    """Write a styled header row and set column widths."""
    for col_idx, (label, width) in enumerate(columns, start=1):
        cell = ws.cell(row=row, column=col_idx, value=label)
        cell.font  = _font(color=T_WHITE, bold=True, size=10)
        cell.fill  = _fill(BG_HEAD)
        cell.alignment = _center()
        cell.border = _border(BORDER_H)
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.row_dimensions[row].height = 20


def _write_title(ws, text: str, span: str, fill_argb: str, row: int = 1) -> None:
    ws.merge_cells(span)
    cell = ws[span.split(":")[0]]
    cell.value = text
    cell.font  = _font(color=T_WHITE, bold=True, size=12)
    cell.fill  = _fill(fill_argb)
    cell.alignment = _left(indent=1)
    ws.row_dimensions[row].height = 26


# ──────────────────────────────────────────────────────────────────────────────
# Main builder
# ──────────────────────────────────────────────────────────────────────────────

def build_excel(
    all_results: list[dict[str, Any]],
    min_dur: int = 4,
) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)

    # ── SHEET 1: SUMMARY ──────────────────────────────────
    ws_sum = wb.create_sheet("Summary")

    ws_sum.merge_cells("A1:F1")
    c = ws_sum["A1"]
    c.value = "⚡  Burst Analysis — Summary"
    c.font  = _font(color=T_WHITE, bold=True, size=14)
    c.fill  = _fill(BG_DARK)
    c.alignment = _left()
    ws_sum.row_dimensions[1].height = 30

    SUM_COLS = [
        ("Threshold",    16),
        ("Bursts #",     11),
        ("Total Time",   14),
        ("Avg Power",    13),
        ("Avg HR",       12),
        ("Avg Cadence",  14),
    ]
    _write_header_row(ws_sum, 2, SUM_COLS)

    for res in all_results:
        threshold = res["threshold"]
        color     = res.get("color", "#888888")
        bursts    = res.get("bursts", [])
        n = len(bursts)
        if n == 0:
            continue

        total_dur  = sum(b["duration"] for b in bursts)
        avg_power  = _safe_avg([b["avg_power"] for b in bursts])
        avg_hr_v   = [b["avg_hr"] for b in bursts if b.get("avg_hr")]
        avg_cad_v  = [b["avg_cadence"] for b in bursts if b.get("avg_cadence")]
        avg_hr     = _safe_avg(avg_hr_v)  if avg_hr_v  else None
        avg_cad    = _safe_avg(avg_cad_v) if avg_cad_v else None

        bg_thr = _hex_to_argb(color)
        bg_row = _lighten(color, 0.82)

        row_num = ws_sum.max_row + 1
        values  = [
            f"≥ {round(threshold)} W",
            n,
            _fmt_time(total_dur),
            f"{avg_power:.0f} W",
            f"{avg_hr:.0f} bpm"  if avg_hr  else "—",
            f"{avg_cad:.0f} rpm" if avg_cad else "—",
        ]
        # Per-column colors matching the HTML palette
        text_colors = [T_DARK, T_DARK, C_TIME, C_WATT, C_HR, C_CAD]
        bolds       = [True,   True,   True,   True,   True, True]

        for col_idx, (val, tc, bold) in enumerate(zip(values, text_colors, bolds), start=1):
            cell = ws_sum.cell(row=row_num, column=col_idx, value=val)
            # Threshold column: use the threshold accent color as bg
            bg = bg_thr if col_idx == 1 else bg_row
            fg = T_WHITE if col_idx == 1 else tc
            cell.fill      = _fill(bg)
            cell.font      = _font(color=fg, bold=bold, size=10)
            cell.alignment = _center()
            cell.border    = _border()
        ws_sum.row_dimensions[row_num].height = 19

    ws_sum.freeze_panes = "A3"

    # ── SHEETS 2…N: PER-THRESHOLD DURATION GRID ───────────
    for res in all_results:
        threshold      = res["threshold"]
        color          = res.get("color", "#888888")
        bursts         = res.get("bursts", [])
        duration_counts: dict = res.get("duration_counts", {})

        if not duration_counts:
            for b in bursts:
                dur = round(b["duration"])
                duration_counts[dur] = duration_counts.get(dur, 0) + 1

        sheet_name = f"≥{round(threshold)}W"[:31]
        ws = wb.create_sheet(sheet_name)

        argb   = _hex_to_argb(color)
        light  = _lighten(color, 0.70)
        vlight = _lighten(color, 0.87)

        # Title + summary stats
        avg_pow_all = _safe_avg([b["avg_power"] for b in bursts]) if bursts else 0
        avg_hr_all  = _safe_avg([b["avg_hr"]  for b in bursts if b.get("avg_hr")])  or None
        avg_cad_all = _safe_avg([b["avg_cadence"] for b in bursts if b.get("avg_cadence")]) or None
        total_time  = sum(b["duration"] for b in bursts)

        title_txt = (
            f"≥{round(threshold)} W  ·  {len(bursts)} bursts  ·  "
            f"Tot: {_fmt_time(total_time)}  ·  "
            f"Avg Power: {avg_pow_all:.0f}W  ·  "
            f"Avg HR: {avg_hr_all:.0f} bpm" if avg_hr_all else
            f"≥{round(threshold)} W  ·  {len(bursts)} bursts  ·  "
            f"Tot: {_fmt_time(total_time)}  ·  "
            f"Avg Power: {avg_pow_all:.0f}W"
        )
        _write_title(ws, title_txt, "A1:H1", argb)

        GRID_COLS = [
            ("Duration (s)",   13),
            ("Count",          10),
            ("Total Time",     13),
            ("Avg Power (W)",  14),
            ("Avg HR (bpm)",   14),
            ("Avg Cad (rpm)",  14),
            ("Tot ≥ Count",    13),
            ("Tot ≥ Time",     13),
        ]
        _write_header_row(ws, 2, GRID_COLS)

        sorted_durs = sorted(
            int(d) for d in duration_counts.keys() if int(d) >= min_dur
        )

        for dur in sorted_durs:
            count = duration_counts.get(dur, duration_counts.get(str(dur), 0))
            bd    = [b for b in bursts if round(b["duration"]) == dur]

            total_d  = sum(b["duration"] for b in bd)
            avg_pw   = _safe_avg([b["avg_power"] for b in bd])
            hr_v     = [b["avg_hr"]       for b in bd if b.get("avg_hr")]
            cad_v    = [b["avg_cadence"]  for b in bd if b.get("avg_cadence")]
            avg_hr   = _safe_avg(hr_v)  if hr_v  else None
            avg_cad  = _safe_avg(cad_v) if cad_v else None

            cum_b   = [b for b in bursts if round(b["duration"]) >= dur]
            cum_cnt = len(cum_b)
            cum_dur = sum(b["duration"] for b in cum_b)

            row_vals   = [f"{dur}s", count, _fmt_time(total_d),
                          f"{avg_pw:.0f}", f"{avg_hr:.0f}" if avg_hr else "—",
                          f"{avg_cad:.0f}" if avg_cad else "—",
                          cum_cnt, _fmt_time(cum_dur)]

            # bg: accent-light for key cols, very-light otherwise
            bg_map   = {2: light, 3: light, 7: light, 8: light}
            # text color per column
            tc_map   = {
                1: _hex_to_argb(color),   # duration → threshold accent
                2: _hex_to_argb(color),   # count
                3: C_TIME,                # total time → green
                4: C_WATT,                # avg power → purple
                5: C_HR,                  # avg HR → red
                6: C_CAD,                 # avg cadence → blue
                7: _hex_to_argb(color),   # cum count
                8: C_TIME,                # cum time
            }
            bold_map = {1, 2, 3, 4, 5, 6, 7, 8}

            row_num = ws.max_row + 1
            for col_idx, val in enumerate(row_vals, start=1):
                cell = ws.cell(row=row_num, column=col_idx, value=val)
                cell.fill      = _fill(bg_map.get(col_idx, vlight))
                cell.font      = _font(
                    color=tc_map.get(col_idx, T_DARK),
                    bold=(col_idx in bold_map),
                    size=10,
                )
                cell.alignment = _center()
                cell.border    = _border()
            ws.row_dimensions[row_num].height = 18

        ws.freeze_panes = "A3"

    # ── LAST SHEET: ALL BURSTS DETAIL ─────────────────────
    ws_all = wb.create_sheet("All Bursts")

    _write_title(ws_all, "📋  Tutti i Burst — Dettaglio Completo", "A1:J1", BG_MID)

    ALL_COLS = [
        ("Threshold",     14),
        ("#",              6),
        ("Start",         10),
        ("Duration (s)",  13),
        ("Avg W",         11),
        ("Max W",         11),
        ("Min W",         11),
        ("Delta W",       11),
        ("HR avg",        11),
        ("Cad avg",       11),
    ]
    _write_header_row(ws_all, 2, ALL_COLS)

    for res in all_results:
        threshold = res["threshold"]
        color     = res.get("color", "#888888")
        vlight    = _lighten(color, 0.88)
        argb_thr  = _hex_to_argb(color)

        for b in res.get("bursts", []):
            hour = b.get("hour", "")
            if not hour or hour == "undefined":
                hour = _fmt_time(b.get("start_time", 0))

            row_vals = [
                f"≥{round(threshold)}W",
                b["rank"],
                hour,
                b["duration"],
                round(b["avg_power"]),
                b["max_power"],
                b["min_power"],
                round(b.get("delta_above", 0), 1),
                b["avg_hr"]       if b.get("avg_hr")       else "—",
                b["avg_cadence"]  if b.get("avg_cadence")  else "—",
            ]
            # text colors per column
            tc_all = {
                1: argb_thr,  # threshold label → accent
                2: T_MUTED,   # rank
                3: T_DARK,    # start time
                4: C_WATT,    # duration (power context)
                5: C_WATT,    # avg W → purple
                6: C_MAX,     # max W → deep purple
                7: C_MIN,     # min W → light lavender
                8: C_DELTA,   # delta → green
                9: C_HR,      # HR → red
                10: C_CAD,    # cad → blue
            }
            bold_all = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10}

            row_num = ws_all.max_row + 1
            for col_idx, val in enumerate(row_vals, start=1):
                cell = ws_all.cell(row=row_num, column=col_idx, value=val)
                cell.fill      = _fill(vlight)
                cell.font      = _font(
                    color=tc_all.get(col_idx, T_DARK),
                    bold=(col_idx in bold_all),
                    size=10,
                )
                cell.alignment = _center()
                cell.border    = _border()
            ws_all.row_dimensions[row_num].height = 17

    ws_all.freeze_panes = "A3"

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()