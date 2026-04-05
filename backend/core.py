from __future__ import annotations

from io import BytesIO
from typing import Any

from fitparse import FitFile

from backend.utils import avg, fmt_time, to_number, to_timestamp


def parse_fit_records(file_bytes: bytes) -> list[dict[str, Any]]:
    fit = FitFile(BytesIO(file_bytes))
    records: list[dict[str, Any]] = []

    for record in fit.get_messages("record"):
        values: dict[str, Any] = {field.name: field.value for field in record}

        power = to_number(values.get("power"))
        if power is None:
            continue

        rec = {
            "timestamp": to_timestamp(values.get("timestamp")),
            "power": power,
            "heartrate": to_number(values.get("heart_rate")),
            "cadence": to_number(values.get("cadence")),
            "distance": to_number(values.get("distance")),
            "altitude": to_number(values.get("enhanced_altitude") or values.get("altitude")),
        }
        records.append(rec)

    if not records:
        raise ValueError("Nessun dato di potenza trovato.")

    t0 = records[0]["timestamp"] or 0
    for idx, rec in enumerate(records):
        ts = rec["timestamp"]
        rec["time_sec"] = (ts - t0) if ts is not None else idx

    return records


def _find_and_merge_burst_segments(
    records: list[dict[str, Any]], threshold: float, merge_gap: int
) -> list[dict[str, int]]:
    """Find all segments above threshold and merge nearby ones.
    
    Returns list of segments: [{"s": start_idx, "e": end_idx}, ...]
    """
    segs: list[dict[str, int]] = []
    in_burst = False
    burst_start = 0

    # Find all segments above threshold
    for i, rec in enumerate(records):
        if rec["power"] >= threshold and not in_burst:
            in_burst = True
            burst_start = i
        elif rec["power"] < threshold and in_burst:
            in_burst = False
            segs.append({"s": burst_start, "e": i - 1})

    if in_burst:
        segs.append({"s": burst_start, "e": len(records) - 1})

    # Merge nearby segments
    merged: list[dict[str, int]] = []
    for seg in segs:
        if not merged:
            merged.append(seg.copy())
            continue

        last = merged[-1]
        if records[seg["s"]]["time_sec"] - records[last["e"]]["time_sec"] <= merge_gap:
            last["e"] = seg["e"]
        else:
            merged.append(seg.copy())

    return merged


def detect_bursts(records: list[dict[str, Any]], threshold: float, min_dur: int, merge_gap: int) -> list[dict[str, Any]]:
    merged = _find_and_merge_burst_segments(records, threshold, merge_gap)

    bursts: list[dict[str, Any]] = []
    for idx, seg in enumerate(merged):
        s = seg["s"]
        e = seg["e"]

        duration = records[e]["time_sec"] - records[s]["time_sec"]
        if duration < min_dur:
            continue

        slice_records = records[s : e + 1]
        powers = [r["power"] for r in slice_records]
        avg_power = avg(powers)

        half = len(powers) // 2
        first_half = avg(powers[:half]) if half else avg_power
        second_half = avg(powers[half:]) if powers[half:] else avg_power

        hrs = [r["heartrate"] for r in slice_records if r.get("heartrate") and r["heartrate"] > 0]
        cads = [r["cadence"] for r in slice_records if r.get("cadence") and r["cadence"] > 0]

        bursts.append(
            {
                "rank": idx + 1,
                "seg_start": s,
                "seg_end": e,
                "start_time": records[s]["time_sec"],
                "end_time": records[e]["time_sec"],
                "hour": fmt_time(records[s]["time_sec"]),
                "duration": round(duration, 1),
                "avg_power": round(avg_power, 1),
                "max_power": max(powers),
                "min_power": min(powers),
                "delta_above": round(avg_power - threshold, 1),
                "fatigue_idx": round((second_half / first_half), 3) if first_half > 0 else 1,
                "avg_hr": round(avg(hrs)) if hrs else None,
                "avg_cadence": round(avg(cads)) if cads else None,
            }
        )

    return bursts


def count_bursts_by_exact_duration(records: list[dict[str, Any]], threshold: float, merge_gap: int) -> dict[int, int]:
    """Count bursts by exact duration (no minimum duration filter).
    
    Returns dict {duration_sec: count}.
    """
    merged = _find_and_merge_burst_segments(records, threshold, merge_gap)

    # Count by exact duration
    duration_counts: dict[int, int] = {}
    for seg in merged:
        s = seg["s"]
        e = seg["e"]
        duration = round(records[e]["time_sec"] - records[s]["time_sec"])
        duration_counts[duration] = duration_counts.get(duration, 0) + 1
    
    return duration_counts


def analyze_records(
    records: list[dict[str, Any]],
    thresholds: list[dict[str, Any]],
    min_dur: int,
    merge_gap: int,
) -> list[dict[str, Any]]:
    ordered = sorted(thresholds, key=lambda t: float(t.get("watt", 0)))
    results: list[dict[str, Any]] = []

    for thr in ordered:
        watt = float(thr.get("watt", 0))
        color = thr.get("color")
        bursts = detect_bursts(records, watt, min_dur, merge_gap)
        duration_counts = count_bursts_by_exact_duration(records, watt, merge_gap)
        results.append({
            "threshold": watt, 
            "color": color, 
            "bursts": bursts,
            "duration_counts": duration_counts  # NEW: mappa delle durate
        })

    return results
