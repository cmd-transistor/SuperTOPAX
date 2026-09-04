"""Topax IDVG plotter.

Reads per-die Topax CSV files, plots AbsMeasDrainCurrent (and optionally
AbsMeasSourceCurrent) vs GateVoltage on a log-Y axis, and assembles the
resulting charts into a 4-up PowerPoint deck grouped by device type and wafer.

Quick user customization (edit the constants in the USER CONFIGURATION block
below, or override on the command line):
    * DEFAULT_TOPAX_DIR: folder that holds the per-die CSVs (recursively scanned).
    * DEFAULT_STRUCTURES: exact structure names to include by default.
        To add another structure, append its exact name string to this list.
    * DEFAULT_WAFERS: wafer numbers to include by default.
    * TARGET_CURRENT_NA: ATOPT leakage target current (in nA) per device type.
        Add a new device-type prefix (e.g. "NL") + target here to support it.
    * VDS_COLOR_MAP / SOURCE_VDS_COLOR_MAP: point colors keyed by DrainVoltage.
        Add or edit entries to control coloring for extra Vds values.

See parse_args() at the bottom for every CLI flag, grouped by purpose.
"""

from __future__ import annotations

import argparse
import csv
import math
import os
import re
from concurrent.futures import ProcessPoolExecutor
from datetime import datetime
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import matplotlib

matplotlib.use("Agg")

import matplotlib.pyplot as plt  # noqa: E402
from matplotlib.ticker import AutoMinorLocator, FormatStrFormatter  # noqa: E402
from pptx import Presentation  # noqa: E402
from pptx.util import Inches, Pt  # noqa: E402


# ---------------------------------------------------------------------------
# Filename parser
# ---------------------------------------------------------------------------
# Matches names like:
#   D520FEJ_w363_DieX-1Y2_MFF_NS-108-A36-1X1-IXI-LVT-CMOS_TC31_1.csv
#   D617F5X_w375_DieX-1Y-1_CRITMOS_FS_ALL_NS-108-A36-1X1-IXI-HVT-CMOS_TC31_1.csv
# The package token (e.g. MFF, CRITMOS_FS_ALL) may itself contain underscores,
# so we lazily consume it and anchor the structure by requiring it to start
# with a supported device-type prefix (NS or PS).
# Groups: lot, wafer, die_x, die_y, package, structure.
FILENAME_RE = re.compile(
    r"^(?P<lot>[^_]+)_w(?P<wafer>\d+)_DieX(?P<die_x>-?\d+)Y(?P<die_y>-?\d+)"
    r"_(?P<package>.+?)_(?P<structure>(?:NS|PS)-.+)\.csv$",
    re.IGNORECASE,
)


# ---------------------------------------------------------------------------
# USER CONFIGURATION - defaults you are most likely to edit
# ---------------------------------------------------------------------------

# Root folder that holds Topax CSVs (scanned recursively).
DEFAULT_TOPAX_DIR = Path(
    r"C:\Users\aoberoi\OneDrive - Intel Corporation\1281\WW36\TPX Rework"
)

# Where PNGs (temporary) and the PowerPoint deck are saved.
DEFAULT_OUTPUT_DIR = DEFAULT_TOPAX_DIR / "topax_idvg_output"

# Structures included when --structures is not passed.
# Append exact structure names here (or pass --structures on the CLI).
DEFAULT_STRUCTURES = [
    # "NS-108-A36-1X1-IXI-LVT-CMOS_TC31_1",
    # "PS-108-A36-1X1-IXI-LVT-CMOS_TC31_1",
    "PS-108-A36-1X1-IXI-HVT-UNI_TC31_1",
    "PS-108-A36-1X1-IXI-HVT-CMOS_TC31_1",
]

# Wafers included when --wafers is not passed.
DEFAULT_WAFERS = [250, 353]

# ATOPT leakage target lines drawn on each chart, in nA, keyed by the first
# token of the structure name (e.g. "NS", "PS"). Add device-type prefixes here
# if new families are introduced.
TARGET_CURRENT_NA = {
    "NS": 0.64,
    "PS": 0.32,
}


# ---------------------------------------------------------------------------
# Point coloring
# ---------------------------------------------------------------------------
# Drain-current colors. ISHI and ISLO sweeps override the Vds-based color.
ISHI_COLOR = "red"
ISLO_COLOR = "green"
REST_COLOR = "blue"
VDS_COLOR_MAP = {
    0.65: "blue",
    -0.65: "purple",
    0.05: "orange",
    -0.05: "cyan",
}

# Source-current colors (only used when --plot-both-currents is set).
# Kept visually distinct from the drain palette so overlays remain readable.
SOURCE_ISHI_COLOR = "firebrick"
SOURCE_ISLO_COLOR = "seagreen"
SOURCE_REST_COLOR = "steelblue"
SOURCE_VDS_COLOR_MAP = {
    0.65: "navy",
    -0.65: "indigo",
    0.05: "darkorange",
    -0.05: "teal",
}

# Source (Is), gate (Ig), and body (Ib) current overlays are drawn as connected
# lines rather than scatter points, in a single color each, distinguished by dash style.
SOURCE_LINE_COLOR = "seagreen"
SOURCE_LINESTYLE = (0, (1, 3))   # dotted
GATE_LINE_COLOR = "red"
GATE_LINESTYLE = (0, (1, 1))     # densely dotted
BODY_LINE_COLOR = "purple"
BODY_LINESTYLE = (0, (3, 1, 1, 1))   # densely dashdotted



@dataclass(frozen=True)
class FileMeta:
    path: Path
    lot: str
    wafer: int
    die_x: int
    die_y: int
    package: str
    structure: str
    device_type: str
    # TPX-only: several structures live in one CSV (keyed by the 'Device'
    # column), so rows must be filtered to this value. None for Topax files.
    tpx_device: str | None = None


@dataclass
class PlotRecord:
    gate_v: list[float]
    current_abs: list[float]
    colors: list[str]
    labels: list[str]
    markers: list[str]
    drain_fit_gate_v: list[float]
    drain_fit_current_abs: list[float]
    source_gate_v: list[float]
    source_current_abs: list[float]
    source_dv: list[float]
    gate_i_gate_v: list[float]
    gate_i_current_abs: list[float]
    gate_i_dv: list[float]
    body_i_gate_v: list[float]
    body_i_current_abs: list[float]
    body_i_dv: list[float]
    included_count: int
    plotted_point_count: int
    total_count: int


@dataclass(frozen=True)
class FitResult:
    slope: float
    intercept: float
    fit_points: int
    x_at_window_high: float
    x_at_window_low: float
    x_at_target: float


def parse_csv_list(value: str) -> list[str]:
    return [item.strip() for item in value.split(",") if item.strip()]


def parse_csv_floats(value: str) -> list[float]:
    parsed: list[float] = []
    for token in parse_csv_list(value):
        parsed.append(float(token))
    return parsed


def parse_csv_ints(value: str) -> list[int]:
    parsed: list[int] = []
    for token in parse_csv_list(value):
        parsed.append(int(token))
    return parsed


def nearly_equal(a: float, b: float, tol: float = 1e-6) -> bool:
    return abs(a - b) <= tol


def discover_files(topax_dir: Path) -> list[FileMeta]:
    metas: list[FileMeta] = []

    for csv_path in topax_dir.rglob("*.csv"):
        match = FILENAME_RE.match(csv_path.name)
        if not match:
            continue

        structure = match.group("structure")
        device_type = structure.split("-", 1)[0].upper()
        if device_type not in {"NS", "PS"}:
            continue

        metas.append(
            FileMeta(
                path=csv_path,
                lot=match.group("lot"),
                wafer=int(match.group("wafer")),
                die_x=int(match.group("die_x")),
                die_y=int(match.group("die_y")),
                package=match.group("package"),
                structure=structure,
                device_type=device_type,
            )
        )

    return sorted(
        metas,
        key=lambda m: (m.device_type, m.lot, m.wafer, m.structure, m.die_x, m.die_y),
    )


# TPX CSVs (Process,Product,LotNumber,Wafer,DieXY,TimeStamp,Device,VD,VG,ID,...)
# hold every structure for a die in one file, distinguished by the 'Device'
# column, and use a die-coordinate string like 'X-1Y4' instead of DieX/Y columns.
TPX_REQUIRED_COLUMNS = {"LotNumber", "Wafer", "DieXY", "Device", "VG", "VD", "ID"}
TPX_DIE_XY_RE = re.compile(r"^X(-?\d+)Y(-?\d+)$", re.IGNORECASE)


def discover_tpx_files(topax_dir: Path) -> list[FileMeta]:
    metas: list[FileMeta] = []

    for csv_path in topax_dir.rglob("*.csv"):
        with csv_path.open("r", encoding="utf-8-sig", newline="") as handle:
            reader = csv.DictReader(handle)
            if not TPX_REQUIRED_COLUMNS.issubset(set(reader.fieldnames or [])):
                continue

            seen: set[tuple[str, int, int, int]] = set()
            for row in reader:
                device = (row.get("Device") or "").strip()
                device_type = device.split("-", 1)[0].upper()
                if device_type not in {"NS", "PS"}:
                    continue

                die_match = TPX_DIE_XY_RE.match((row.get("DieXY") or "").strip())
                if die_match is None:
                    continue
                try:
                    wafer = int((row.get("Wafer") or "").strip())
                except ValueError:
                    continue

                die_x, die_y = int(die_match.group(1)), int(die_match.group(2))
                key = (device, wafer, die_x, die_y)
                if key in seen:
                    continue
                seen.add(key)

                metas.append(
                    FileMeta(
                        path=csv_path,
                        lot=(row.get("LotNumber") or "").strip(),
                        wafer=wafer,
                        die_x=die_x,
                        die_y=die_y,
                        package="TPX",
                        structure=device,
                        device_type=device_type,
                        tpx_device=device,
                    )
                )

    return sorted(
        metas,
        key=lambda m: (m.device_type, m.lot, m.wafer, m.structure, m.die_x, m.die_y),
    )


def structure_selected(structure: str, requested: set[str]) -> bool:
    return structure in requested


def lot_selected(lot: str, lot_prefixes: Iterable[str]) -> bool:
    lot_prefixes = list(lot_prefixes)
    if not lot_prefixes:
        return True
    lot_upper = lot.upper()
    return any(lot_upper.startswith(prefix.upper()) for prefix in lot_prefixes)


def row_style(sweep_id: str, drain_v: float, color_by_vds: bool, current_kind: str) -> tuple[str, str, str]:
    if current_kind == "source":
        ishi_color = SOURCE_ISHI_COLOR
        islo_color = SOURCE_ISLO_COLOR
        rest_color = SOURCE_REST_COLOR
        vds_color_map = SOURCE_VDS_COLOR_MAP
        marker = "^"
        label_prefix = "Source"
    else:
        ishi_color = ISHI_COLOR
        islo_color = ISLO_COLOR
        rest_color = REST_COLOR
        vds_color_map = VDS_COLOR_MAP
        marker = "o"
        label_prefix = "Drain"

    sweep_upper = sweep_id.upper()
    if sweep_upper.startswith("ISHI"):
        return ishi_color, f"{label_prefix} ISHI*", marker
    if sweep_upper.startswith("ISLO"):
        return islo_color, f"{label_prefix} ISLO*", marker

    if color_by_vds:
        chosen = vds_color_map.get(drain_v, rest_color)
        return chosen, f"{label_prefix} VDS={drain_v:g}V", marker

    return rest_color, f"{label_prefix} Other", marker


def should_keep_vds(drain_v: float, device_type: str, magnitudes: set[float], compress_sign: bool) -> bool:
    abs_v = abs(drain_v)
    if not any(nearly_equal(abs_v, mag) for mag in magnitudes):
        return False

    if not compress_sign:
        return True

    if device_type == "NS" and drain_v < 0:
        return False
    if device_type == "PS" and drain_v > 0:
        return False
    return True


def load_plot_record(
    meta: FileMeta,
    include_vds_magnitudes: set[float],
    keep_ishi: bool,
    keep_islo: bool,
    pmos_flip_gate: bool,
    color_by_vds: bool,
    compress_sign: bool,
    plot_both_currents: bool,
    plot_gate_current: bool,
    plot_body_current: bool,
    fit_vds_magnitude: float,
) -> PlotRecord:
    gate_v: list[float] = []
    current_abs: list[float] = []
    colors: list[str] = []
    labels: list[str] = []
    markers: list[str] = []
    drain_fit_gate_v: list[float] = []
    drain_fit_current_abs: list[float] = []
    source_gate_v: list[float] = []
    source_current_abs: list[float] = []
    source_dv: list[float] = []
    gate_i_gate_v: list[float] = []
    gate_i_current_abs: list[float] = []
    gate_i_dv: list[float] = []
    body_i_gate_v: list[float] = []
    body_i_current_abs: list[float] = []
    body_i_dv: list[float] = []

    total_count = 0
    included_count = 0
    plotted_point_count = 0

    is_tpx = meta.tpx_device is not None
    col_gate, col_drain_v, col_drain_i, col_source_i = (
        ("VG", "VD", "ID", "IS") if is_tpx else
        ("GateVoltage", "DrainVoltage", "AbsMeasDrainCurrent", "AbsMeasSourceCurrent")
    )
    col_gate_i = "IG" if is_tpx else "AbsMeasGateCurrent"
    col_body_i = "IB" if is_tpx else "AbsMeasBodyCurrent"

    with meta.path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        fieldnames = set(reader.fieldnames or [])
        required = {col_gate, col_drain_v, col_drain_i}
        if is_tpx:
            required.add("Device")
        else:
            required.add("SweepID")
        if plot_both_currents:
            required.add(col_source_i)
        missing = required - fieldnames
        if missing:
            raise ValueError(f"{meta.path.name} missing required columns: {sorted(missing)}")
        has_sweep_id = "SweepID" in fieldnames
        # Gate/body current aren't guaranteed to exist in every file/format; skip gracefully if absent.
        has_gate_i = plot_gate_current and col_gate_i in fieldnames
        has_body_i = plot_body_current and col_body_i in fieldnames

        for row in reader:
            if is_tpx and (row.get("Device") or "").strip() != meta.tpx_device:
                continue

            total_count += 1

            sweep_id = (row.get("SweepID") or "").strip() if has_sweep_id else ""
            try:
                gv = float((row.get(col_gate) or "").strip())
                dv = float((row.get(col_drain_v) or "").strip())
                ids_raw = float((row.get(col_drain_i) or "").strip())
                # TPX current columns are signed; Topax's AbsMeasDrainCurrent is already absolute.
                ids_abs = abs(ids_raw) if is_tpx else ids_raw
            except ValueError:
                continue

            iss_abs: float | None = None
            if plot_both_currents:
                try:
                    iss_raw = float((row.get(col_source_i) or "").strip())
                    iss_abs = abs(iss_raw) if is_tpx else iss_raw
                except ValueError:
                    iss_abs = None

            igg_abs: float | None = None
            if has_gate_i:
                try:
                    igg_raw = float((row.get(col_gate_i) or "").strip())
                    igg_abs = abs(igg_raw) if is_tpx else igg_raw
                except ValueError:
                    igg_abs = None

            ibb_abs: float | None = None
            if has_body_i:
                try:
                    ibb_raw = float((row.get(col_body_i) or "").strip())
                    ibb_abs = abs(ibb_raw) if is_tpx else ibb_raw
                except ValueError:
                    ibb_abs = None

            keep_for_sweep = False
            if has_sweep_id:
                sweep_upper = sweep_id.upper()
                if keep_ishi and sweep_upper.startswith("ISHI"):
                    keep_for_sweep = True
                if keep_islo and sweep_upper.startswith("ISLO"):
                    keep_for_sweep = True

            keep_for_vds = should_keep_vds(dv, meta.device_type, include_vds_magnitudes, compress_sign)
            if not (keep_for_sweep or keep_for_vds):
                continue

            if meta.device_type == "PS" and pmos_flip_gate:
                gv = -gv

            plotted_in_row = False

            if ids_abs > 0 and math.isfinite(ids_abs):
                color, label, marker = row_style(sweep_id, dv, color_by_vds, current_kind="drain")
                gate_v.append(gv)
                current_abs.append(ids_abs)
                colors.append(color)
                labels.append(label)
                markers.append(marker)
                # Restrict SS fit points to the requested Vds magnitude only.
                if nearly_equal(abs(dv), fit_vds_magnitude):
                    drain_fit_gate_v.append(gv)
                    drain_fit_current_abs.append(ids_abs)
                plotted_point_count += 1
                plotted_in_row = True

            if plot_both_currents and iss_abs is not None and iss_abs > 0 and math.isfinite(iss_abs):
                source_gate_v.append(gv)
                source_current_abs.append(iss_abs)
                source_dv.append(dv)
                plotted_point_count += 1
                plotted_in_row = True

            if igg_abs is not None and igg_abs > 0 and math.isfinite(igg_abs):
                gate_i_gate_v.append(gv)
                gate_i_current_abs.append(igg_abs)
                gate_i_dv.append(dv)
                plotted_point_count += 1
                plotted_in_row = True

            if ibb_abs is not None and ibb_abs > 0 and math.isfinite(ibb_abs):
                body_i_gate_v.append(gv)
                body_i_current_abs.append(ibb_abs)
                body_i_dv.append(dv)
                plotted_point_count += 1
                plotted_in_row = True

            if plotted_in_row:
                included_count += 1

    # Per-VDS grouping/sorting for source & gate current happens in save_single_plot,
    # since a single die can carry multiple VDS branches (e.g. 0.05V and 0.65V).
    return PlotRecord(
        gate_v=gate_v,
        current_abs=current_abs,
        colors=colors,
        labels=labels,
        markers=markers,
        drain_fit_gate_v=drain_fit_gate_v,
        drain_fit_current_abs=drain_fit_current_abs,
        source_gate_v=source_gate_v,
        source_current_abs=source_current_abs,
        source_dv=source_dv,
        gate_i_gate_v=gate_i_gate_v,
        gate_i_current_abs=gate_i_current_abs,
        gate_i_dv=gate_i_dv,
        body_i_gate_v=body_i_gate_v,
        body_i_current_abs=body_i_current_abs,
        body_i_dv=body_i_dv,
        included_count=included_count,
        plotted_point_count=plotted_point_count,
        total_count=total_count,
    )


def build_plot_title(
    meta: FileMeta,
    vds_magnitudes: set[float],
    keep_ishi: bool,
    keep_islo: bool,
    plot_both_currents: bool,
    plot_gate_current: bool,
    plot_body_current: bool,
) -> str:
    vds_text = ",".join(f"{v:g}" for v in sorted(vds_magnitudes))
    sweep_parts: list[str] = []
    if keep_ishi:
        sweep_parts.append("ISHI*")
    if keep_islo:
        sweep_parts.append("ISLO*")
    sweep_text = ", ".join(sweep_parts) if sweep_parts else "none"
    current_parts = ["Drain"]
    if plot_both_currents:
        current_parts.append("Source")
    if plot_gate_current:
        current_parts.append("Gate")
    if plot_body_current:
        current_parts.append("Body")
    current_text = "+".join(current_parts)

    return (
        f"{meta.lot} w{meta.wafer} Die({meta.die_x},{meta.die_y}) {meta.structure} | "
        f"Filter: |VDS| in [{vds_text}] + {sweep_text} | Current: {current_text}"
    )


def linear_fit_to_target(
    gate_v: list[float],
    id_abs: list[float],
    fit_current_high: float,
    fit_current_low: float,
    target_current: float,
) -> FitResult | None:
    if fit_current_low <= 0 or fit_current_high <= 0 or target_current <= 0:
        return None

    low = min(fit_current_low, fit_current_high)
    high = max(fit_current_low, fit_current_high)

    fit_x: list[float] = []
    fit_log_y: list[float] = []
    for x_val, y_val in zip(gate_v, id_abs):
        if low <= y_val <= high and math.isfinite(x_val) and math.isfinite(y_val):
            fit_x.append(x_val)
            fit_log_y.append(math.log10(y_val))

    n = len(fit_x)
    if n < 2:
        return None

    sx = sum(fit_x)
    sy = sum(fit_log_y)
    sxx = sum(x * x for x in fit_x)
    sxy = sum(x * y for x, y in zip(fit_x, fit_log_y))

    denom = n * sxx - sx * sx
    if abs(denom) < 1e-20:
        return None

    slope = (n * sxy - sx * sy) / denom
    if abs(slope) < 1e-12:
        return None

    intercept = (sy - slope * sx) / n
    x_at_high = (math.log10(high) - intercept) / slope
    x_at_low = (math.log10(low) - intercept) / slope
    x_at_target = (math.log10(target_current) - intercept) / slope

    return FitResult(
        slope=slope,
        intercept=intercept,
        fit_points=n,
        x_at_window_high=x_at_high,
        x_at_window_low=x_at_low,
        x_at_target=x_at_target,
    )


def apply_axes_style(
    ax: plt.Axes,
    meta: FileMeta,
    y_limits_mode: str,
    y_min: float,
    y_max: float,
    x_limits_mode: str,
    x_min: float,
    x_max: float,
) -> None:
    ax.set_yscale("log")
    ax.set_xlabel("GateVoltage (V)")
    ax.set_ylabel("Abs Current (A)")

    if x_limits_mode == "fixed":
        ax.set_xlim(x_min, x_max)
        major_ticks = [round(x_min + 0.2 * i, 10) for i in range(int(round((x_max - x_min) / 0.2)) + 1)]
        ax.set_xticks(major_ticks)
        ax.xaxis.set_minor_locator(AutoMinorLocator(4))

    if y_limits_mode == "fixed":
        ax.set_ylim(y_min, y_max)

    ax.grid(True, which="major", axis="both", linestyle="-", linewidth=0.6, alpha=0.5)
    ax.grid(True, which="minor", axis="both", linestyle=":", linewidth=0.4, alpha=0.4)

    target_na = TARGET_CURRENT_NA[meta.device_type]
    target_a = target_na * 1e-9
    ax.axhline(target_a, color="black", linestyle="--", linewidth=1.2, label=f"ATOPT {target_na} nA")
    ax.axvline(0.0, color="gray", linestyle="--", linewidth=1.0, label="Vg = 0 V")


def save_single_plot(
    meta: FileMeta,
    record: PlotRecord,
    out_dir: Path,
    y_limits_mode: str,
    y_min: float,
    y_max: float,
    x_limits_mode: str,
    x_min: float,
    x_max: float,
    include_vds_magnitudes: set[float],
    keep_ishi: bool,
    keep_islo: bool,
    plot_both_currents: bool,
    plot_gate_current: bool,
    plot_body_current: bool,
    enable_fit_to_atopt: bool,
    fit_current_high: float,
    fit_current_low: float,
    fit_vds_magnitude: float,
) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)

    file_stub = f"{meta.lot}_w{meta.wafer}_DieX{meta.die_x}Y{meta.die_y}_{meta.structure}"
    safe_stub = re.sub(r"[^A-Za-z0-9_.-]+", "_", file_stub)
    out_path = out_dir / f"{safe_stub}.png"

    fig, ax1 = plt.subplots(figsize=(8, 5), dpi=160)
    grouped_points: dict[tuple[str, str, str], tuple[list[float], list[float]]] = {}
    for x_val, y_val, color, label, marker in zip(
        record.gate_v,
        record.current_abs,
        record.colors,
        record.labels,
        record.markers,
    ):
        key = (label, color, marker)
        if key not in grouped_points:
            grouped_points[key] = ([], [])
        grouped_points[key][0].append(x_val)
        grouped_points[key][1].append(y_val)

    for label, color, marker in sorted(grouped_points.keys()):
        xs, ys = grouped_points[(label, color, marker)]
        ax1.scatter(xs, ys, c=color, s=12, alpha=0.85, linewidths=0, marker=marker, label=label)

    # Group source/gate current by VDS branch (a die can carry e.g. both 0.05V
    # and 0.65V sweeps) so each branch is drawn as its own sorted, non-zigzag line.
    def _group_by_dv(gv_list: list[float], i_list: list[float], dv_list: list[float]) -> dict[float, tuple[list[float], list[float]]]:
        groups: dict[float, tuple[list[float], list[float]]] = {}
        for gv_val, i_val, dv_val in zip(gv_list, i_list, dv_list):
            key = round(dv_val, 3)
            groups.setdefault(key, ([], []))[0].append(gv_val)
            groups.setdefault(key, ([], []))[1].append(i_val)
        for key, (xs, ys) in groups.items():
            order = sorted(range(len(xs)), key=lambda i: xs[i])
            groups[key] = ([xs[i] for i in order], [ys[i] for i in order])
        return groups

    source_groups = _group_by_dv(record.source_gate_v, record.source_current_abs, record.source_dv)
    gate_i_groups = _group_by_dv(record.gate_i_gate_v, record.gate_i_current_abs, record.gate_i_dv)
    body_i_groups = _group_by_dv(record.body_i_gate_v, record.body_i_current_abs, record.body_i_dv)

    for dv_key in sorted(source_groups.keys()):
        xs, ys = source_groups[dv_key]
        color = SOURCE_VDS_COLOR_MAP.get(dv_key, SOURCE_LINE_COLOR)
        ax1.plot(xs, ys, color=color, linestyle=SOURCE_LINESTYLE, linewidth=1.2, label=f"Source Is VDS={dv_key:g}V")
    for dv_key in sorted(gate_i_groups.keys()):
        xs, ys = gate_i_groups[dv_key]
        ax1.plot(xs, ys, color=GATE_LINE_COLOR, linestyle=GATE_LINESTYLE, linewidth=1.2)
    for dv_key in sorted(body_i_groups.keys()):
        xs, ys = body_i_groups[dv_key]
        ax1.plot(xs, ys, color=BODY_LINE_COLOR, linestyle=BODY_LINESTYLE, linewidth=1.2)

    ax1.set_title(build_plot_title(meta, include_vds_magnitudes, keep_ishi, keep_islo, plot_both_currents, plot_gate_current, plot_body_current), fontsize=9)

    apply_axes_style(ax1, meta, y_limits_mode, y_min, y_max, x_limits_mode, x_min, x_max)

    target_a = TARGET_CURRENT_NA[meta.device_type] * 1e-9
    fit_result: FitResult | None = None
    if enable_fit_to_atopt:
        fit_result = linear_fit_to_target(
            gate_v=record.drain_fit_gate_v,
            id_abs=record.drain_fit_current_abs,
            fit_current_high=fit_current_high,
            fit_current_low=fit_current_low,
            target_current=target_a,
        )
        if fit_result is not None:
            x1 = fit_result.x_at_window_high
            x2 = fit_result.x_at_target
            y1 = fit_current_high
            y2 = target_a
            ax1.plot([x1, x2], [y1, y2], color="magenta", linestyle="-.", linewidth=1.2)
            ax1.scatter([fit_result.x_at_target], [target_a], color="magenta", s=20, marker="x")

    # Secondary y-axis: drain/source curves (same Amps units) on a linear scale, so
    # subthreshold (log) and linear-region (linear) characteristics both show.
    # Gate (Ig) and body (Ib) currents are leakage-only and stay log-scale (ax1) only.
    ax2 = ax1.twinx()
    for label, color, marker in sorted(grouped_points.keys()):
        xs, ys = grouped_points[(label, color, marker)]
        ax2.scatter(xs, ys, c=color, s=12, alpha=0.85, linewidths=0, marker=marker)
    for dv_key in sorted(source_groups.keys()):
        xs, ys = source_groups[dv_key]
        color = SOURCE_VDS_COLOR_MAP.get(dv_key, SOURCE_LINE_COLOR)
        ax2.plot(xs, ys, color=color, linestyle=SOURCE_LINESTYLE, linewidth=1.2)
    ax2.set_ylabel("Linear Current (A)")
    ax2.set_ylim(bottom=0)
    ax2.yaxis.set_major_formatter(FormatStrFormatter("%.2e"))
    ax2.grid(False)

    legend_handles: list[plt.Line2D] = []
    legend_labels: list[str] = []

    for label, color, marker in sorted(grouped_points.keys()):
        legend_handles.append(
            plt.Line2D([0], [0], marker=marker, color="w", markerfacecolor=color, markersize=6, label=label)
        )
        legend_labels.append(label)

    for dv_key in sorted(source_groups.keys()):
        color = SOURCE_VDS_COLOR_MAP.get(dv_key, SOURCE_LINE_COLOR)
        label = f"Source Is VDS={dv_key:g}V"
        legend_handles.append(
            plt.Line2D([0], [0], color=color, linestyle=SOURCE_LINESTYLE, linewidth=1.2, label=label)
        )
        legend_labels.append(label)
    if gate_i_groups:
        legend_handles.append(
            plt.Line2D([0], [0], color=GATE_LINE_COLOR, linestyle=GATE_LINESTYLE, linewidth=1.2, label="Gate (Ig)")
        )
        legend_labels.append("Gate (Ig)")
    if body_i_groups:
        legend_handles.append(
            plt.Line2D([0], [0], color=BODY_LINE_COLOR, linestyle=BODY_LINESTYLE, linewidth=1.2, label="Body (Ib)")
        )
        legend_labels.append("Body (Ib)")

    legend_handles.append(
        plt.Line2D([0], [0], color="black", linestyle="--", linewidth=1.2, label="ATOPT target")
    )
    legend_labels.append("ATOPT target")
    legend_handles.append(
        plt.Line2D([0], [0], color="gray", linestyle="--", linewidth=1.0, label="Vg = 0 V")
    )
    legend_labels.append("Vg = 0 V")
    if enable_fit_to_atopt:
        legend_handles.append(
            plt.Line2D([0], [0], color="magenta", linestyle="-.", linewidth=1.2, label=f"Drain fit (|Vds|={fit_vds_magnitude:g}) -> ATOPT")
        )
        legend_labels.append(f"Drain fit (|Vds|={fit_vds_magnitude:g}) -> ATOPT")
    ax1.legend(legend_handles, legend_labels, fontsize=7, loc="best")

    fig.tight_layout()
    fig.savefig(out_path)
    plt.close(fig)
    return out_path


def make_ppt(image_entries: list[tuple[FileMeta, Path]], ppt_path: Path) -> Path:
    if not image_entries:
        return ppt_path

    prs = Presentation()
    blank_layout = prs.slide_layouts[6]

    # 4 plots per slide in a 2x2 layout
    positions = [
        (Inches(0.3), Inches(1.0)),
        (Inches(5.1), Inches(1.0)),
        (Inches(0.3), Inches(4.1)),
        (Inches(5.1), Inches(4.1)),
    ]
    pic_width = Inches(4.4)
    pic_height = Inches(2.9)

    sorted_entries = sorted(
        image_entries,
        key=lambda item: (
            item[0].device_type,
            item[0].lot,
            item[0].wafer,
            item[0].structure,
            item[0].die_x,
            item[0].die_y,
        ),
    )

    current_group: tuple[str, str, int] | None = None
    slide = None
    tile_index = 0

    def new_slide(meta: FileMeta):
        nonlocal slide, tile_index
        slide = prs.slides.add_slide(blank_layout)
        tile_index = 0

        title_box = slide.shapes.add_textbox(Inches(0.3), Inches(0.2), Inches(12.7), Inches(0.6))
        title_tf = title_box.text_frame
        title_tf.clear()
        run = title_tf.paragraphs[0].add_run()
        run.text = f"{meta.device_type} | {meta.lot} w{meta.wafer} | Topax IDVG"
        run.font.size = Pt(20)

    for meta, image_path in sorted_entries:
        group_key = (meta.device_type, meta.lot, meta.wafer)
        if slide is None or group_key != current_group or tile_index >= 4:
            current_group = group_key
            new_slide(meta)

        left, top = positions[tile_index]
        slide.shapes.add_picture(str(image_path), left, top, width=pic_width, height=pic_height)

        caption = slide.shapes.add_textbox(left, top + pic_height, pic_width, Inches(0.35))
        tf = caption.text_frame
        tf.clear()
        run = tf.paragraphs[0].add_run()
        run.text = f"Die({meta.die_x},{meta.die_y}) {meta.structure}"
        run.font.size = Pt(9)

        tile_index += 1

    ppt_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        prs.save(ppt_path)
        return ppt_path
    except PermissionError:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        fallback_path = ppt_path.with_name(f"{ppt_path.stem}_{stamp}{ppt_path.suffix}")
        prs.save(fallback_path)
        return fallback_path


def delete_generated_plots(image_entries: list[tuple[FileMeta, Path]]) -> int:
    deleted = 0
    seen: set[Path] = set()
    for _, image_path in image_entries:
        if image_path in seen:
            continue
        seen.add(image_path)
        if image_path.exists():
            image_path.unlink()
            deleted += 1

    plot_dirs = {img.parent for _, img in image_entries}
    for plot_dir in plot_dirs:
        if plot_dir.exists() and not any(plot_dir.iterdir()):
            try:
                plot_dir.rmdir()
            except OSError:
                # On OneDrive/Windows an empty folder can remain locked briefly.
                pass

    return deleted


@dataclass(frozen=True)
class RenderConfig:
    include_vds_magnitudes: frozenset[float]
    keep_ishi: bool
    keep_islo: bool
    pmos_flip_gate: bool
    color_by_vds: bool
    compress_sign: bool
    plot_both_currents: bool
    plot_gate_current: bool
    plot_body_current: bool
    plot_dir: Path
    y_limits_mode: str
    y_min: float
    y_max: float
    x_limits_mode: str
    x_min: float
    x_max: float
    enable_fit_to_atopt: bool
    fit_current_high: float
    fit_current_low: float
    fit_vds_magnitude: float


def render_one_file(meta: FileMeta, cfg: RenderConfig) -> dict[str, str] | None:
    record = load_plot_record(
        meta=meta,
        include_vds_magnitudes=set(cfg.include_vds_magnitudes),
        keep_ishi=cfg.keep_ishi,
        keep_islo=cfg.keep_islo,
        pmos_flip_gate=cfg.pmos_flip_gate,
        color_by_vds=cfg.color_by_vds,
        compress_sign=cfg.compress_sign,
        plot_both_currents=cfg.plot_both_currents,
        plot_gate_current=cfg.plot_gate_current,
        plot_body_current=cfg.plot_body_current,
        fit_vds_magnitude=cfg.fit_vds_magnitude,
    )

    if record.included_count == 0:
        return None

    image_path = save_single_plot(
        meta=meta,
        record=record,
        out_dir=cfg.plot_dir,
        y_limits_mode=cfg.y_limits_mode,
        y_min=cfg.y_min,
        y_max=cfg.y_max,
        x_limits_mode=cfg.x_limits_mode,
        x_min=cfg.x_min,
        x_max=cfg.x_max,
        include_vds_magnitudes=set(cfg.include_vds_magnitudes),
        keep_ishi=cfg.keep_ishi,
        keep_islo=cfg.keep_islo,
        plot_both_currents=cfg.plot_both_currents,
        plot_gate_current=cfg.plot_gate_current,
        plot_body_current=cfg.plot_body_current,
        enable_fit_to_atopt=cfg.enable_fit_to_atopt,
        fit_current_high=cfg.fit_current_high,
        fit_current_low=cfg.fit_current_low,
        fit_vds_magnitude=cfg.fit_vds_magnitude,
    )

    return {
        "lot": meta.lot,
        "wafer": str(meta.wafer),
        "die_x": str(meta.die_x),
        "die_y": str(meta.die_y),
        "structure": meta.structure,
        "device_type": meta.device_type,
        "file": str(meta.path),
        "points_kept": str(record.included_count),
        "points_plotted": str(record.plotted_point_count),
        "points_total": str(record.total_count),
        "plot_path": str(image_path),
    }


def _worker(args_tuple: tuple[FileMeta, RenderConfig]) -> dict[str, str] | None:
    meta, cfg = args_tuple
    return render_one_file(meta, cfg)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Plot Topax IDVG (AbsMeasDrainCurrent, optionally AbsMeasSourceCurrent) "
            "vs GateVoltage and assemble a 4-up PowerPoint deck."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )

    # ------------------------------------------------------------------
    # Input / output paths
    # ------------------------------------------------------------------
    io_group = parser.add_argument_group("Input / output paths")
    io_group.add_argument(
        "--topax-dir",
        type=Path,
        default=DEFAULT_TOPAX_DIR,
        help="Root folder of Topax CSV files (scanned recursively).",
    )
    io_group.add_argument(
        "--output-dir",
        type=Path,
        default=DEFAULT_OUTPUT_DIR,
        help="Directory where PNGs, summary CSV, and PowerPoint deck are written.",
    )

    # ------------------------------------------------------------------
    # Data selection (which lot / wafer / structure to plot)
    # ------------------------------------------------------------------
    sel_group = parser.add_argument_group("Data selection")
    sel_group.add_argument(
        "--lot-prefixes",
        default=None,
        help=(
            "Comma-separated lot-name prefixes to include (case-insensitive). "
            "If omitted, inferred from the --topax-dir path (its parent folder name)."
        ),
    )
    sel_group.add_argument(
        "--wafers",
        default=",".join(str(w) for w in DEFAULT_WAFERS),
        help="Comma-separated wafer numbers to include, e.g. '363,366,367'.",
    )
    sel_group.add_argument(
        "--structures",
        default=",".join(DEFAULT_STRUCTURES),
        help=(
            "Comma-separated exact structure names to include. Add or remove "
            "structures here (or edit DEFAULT_STRUCTURES at the top of the file)."
        ),
    )
    sel_group.add_argument(
        "--packages",
        default="",
        help=(
            "Optional comma-separated test-package tokens to include (e.g. "
            "'MFF' or 'CRITMOS_FS_ALL'). Matched by exact substring (case-insensitive) "
            "against the package portion of the filename. Empty means no package filter."
        ),
    )
    sel_group.add_argument(
        "--exclude-packages",
        default="",
        help=(
            "Optional comma-separated package substrings to exclude (case-insensitive), "
            "applied after --packages."
        ),
    )
    sel_group.add_argument(
        "--max-files",
        type=int,
        default=0,
        help="Optional cap on files (useful for debug runs). 0 disables the cap.",
    )

    # ------------------------------------------------------------------
    # Row-level filtering (VDS values, sweep IDs)
    # ------------------------------------------------------------------
    filt_group = parser.add_argument_group("Row filtering")
    filt_group.add_argument(
        "--include-vds-magnitudes",
        default="0.65,0.05",
        help="Comma-separated |VDS| magnitudes to keep, e.g. '0.65' or '0.65,0.05'.",
    )
    filt_group.add_argument(
        "--compress-vds-sign-by-device",
        action="store_true",
        default=True,
        help="Keep only positive VDS for NMOS and negative VDS for PMOS.",
    )
    filt_group.add_argument(
        "--no-compress-vds-sign-by-device",
        dest="compress_vds_sign_by_device",
        action="store_false",
        help="Keep both signs of the specified |VDS| magnitudes.",
    )
    filt_group.add_argument("--keep-ishi", action="store_true", default=True, help="Include SweepID starting with 'ISHI'.")
    filt_group.add_argument("--no-keep-ishi", dest="keep_ishi", action="store_false", help="Exclude ISHI* sweeps.")
    filt_group.add_argument("--keep-islo", action="store_true", default=True, help="Include SweepID starting with 'ISLO'.")
    filt_group.add_argument("--no-keep-islo", dest="keep_islo", action="store_false", help="Exclude ISLO* sweeps.")

    # ------------------------------------------------------------------
    # Plot appearance (axes, fit, currents plotted)
    # ------------------------------------------------------------------
    plot_group = parser.add_argument_group("Plot appearance")
    plot_group.add_argument(
        "--pmos-flip-gate",
        action="store_true",
        default=True,
        help="For PMOS structures, plot -VG so both device types trend the same direction.",
    )
    plot_group.add_argument(
        "--no-pmos-flip-gate",
        dest="pmos_flip_gate",
        action="store_false",
        help="Plot raw VG for PMOS.",
    )
    plot_group.add_argument(
        "--color-by-vds",
        action="store_true",
        default=True,
        help="Color non-ISHI/ISLO points by their DrainVoltage value.",
    )
    plot_group.add_argument(
        "--single-rest-color",
        dest="color_by_vds",
        action="store_false",
        help="Use a single color for all non-ISHI/ISLO points.",
    )
    plot_group.add_argument(
        "--plot-both-currents",
        action="store_true",
        default=True,
        help="Overlay AbsMeasSourceCurrent (Is) as a dotted line with distinct color.",
    )
    plot_group.add_argument(
        "--no-plot-both-currents",
        dest="plot_both_currents",
        action="store_false",
        help="Do not overlay source current (Is).",
    )
    plot_group.add_argument(
        "--plot-gate-current",
        action="store_true",
        default=True,
        help="Overlay gate current (Ig) as a densely-dotted red line, when the column is present.",
    )
    plot_group.add_argument(
        "--no-plot-gate-current",
        dest="plot_gate_current",
        action="store_false",
        help="Do not overlay gate current (Ig).",
    )
    plot_group.add_argument(
        "--plot-body-current",
        action="store_true",
        default=True,
        help="Overlay body current (Ib) as a densely-dashdotted purple line, when the column is present.",
    )
    plot_group.add_argument(
        "--no-plot-body-current",
        dest="plot_body_current",
        action="store_false",
        help="Do not overlay body current (Ib).",
    )
    plot_group.add_argument("--x-limits", choices=["fixed", "auto"], default="fixed", help="Use fixed X limits or matplotlib auto.")
    plot_group.add_argument("--x-min", type=float, default=-0.5, help="X-axis minimum when --x-limits fixed.")
    plot_group.add_argument("--x-max", type=float, default=1.0, help="X-axis maximum when --x-limits fixed.")
    plot_group.add_argument("--y-limits", choices=["fixed", "auto"], default="fixed", help="Use fixed Y limits or matplotlib auto.")
    plot_group.add_argument("--y-min", type=float, default=1e-11, help="Y-axis minimum (log10) when --y-limits fixed.")
    plot_group.add_argument("--y-max", type=float, default=1e-3, help="Y-axis maximum (log10) when --y-limits fixed.")
    plot_group.add_argument(
        "--enable-fit-to-atopt",
        action="store_true",
        default=True,
        help="Draw a semilog fit between --fit-current-low and --fit-current-high, extrapolated to the ATOPT target.",
    )
    plot_group.add_argument("--disable-fit-to-atopt", dest="enable_fit_to_atopt", action="store_false", help="Skip the fit overlay.")
    plot_group.add_argument("--fit-current-high", type=float, default=1e-7, help="Upper current bound (A) of the fit window.")
    plot_group.add_argument("--fit-current-low", type=float, default=1e-9, help="Lower current bound (A) of the fit window.")
    plot_group.add_argument(
        "--fit-vds-magnitude",
        type=float,
        default=0.65,
        help="|Vds| whose drain points are used for the SS/ATOPT fit. Only points at this Vds contribute to the fit.",
    )

    # ------------------------------------------------------------------
    # Output behavior (PPT, timestamp, cleanup, parallelism)
    # ------------------------------------------------------------------
    out_group = parser.add_argument_group("Output behavior")
    out_group.add_argument("--skip-ppt", action="store_true", help="Do not generate PowerPoint (keep only PNGs and summary).")
    out_group.add_argument(
        "--timestamp-ppt",
        action="store_true",
        default=True,
        help="Append a YYYYMMDD_HHMMSS timestamp to the PowerPoint filename to avoid overwriting older decks.",
    )
    out_group.add_argument(
        "--no-timestamp-ppt",
        dest="timestamp_ppt",
        action="store_false",
        help="Always save as Topax_IDVG_4up.pptx and overwrite any existing file with that name.",
    )
    out_group.add_argument(
        "--delete-plots-after-ppt",
        action="store_true",
        default=True,
        help="Delete intermediate PNGs after the PowerPoint is written.",
    )
    out_group.add_argument(
        "--keep-plots",
        dest="delete_plots_after_ppt",
        action="store_false",
        help="Keep intermediate PNGs on disk after the deck is generated.",
    )
    out_group.add_argument(
        "--workers",
        type=int,
        default=max(1, (os.cpu_count() or 4) - 1),
        help="Number of parallel worker processes. Use 1 to force serial execution.",
    )

    return parser.parse_args()



def main() -> None:
    args = parse_args()

    topax_dir: Path = args.topax_dir
    if not topax_dir.exists():
        raise FileNotFoundError(f"Topax directory not found: {topax_dir}")

    # ------------------------------------------------------------------
    # 1) Parse CLI filters into simple containers used by discovery.
    # ------------------------------------------------------------------
    if args.lot_prefixes:
        lot_prefixes = parse_csv_list(args.lot_prefixes)
    else:
        # No filter: accept every lot found under --topax-dir (folder depth
        # varies between Topax's .../<LOT>/Topax layout and TPX's deeper tree).
        lot_prefixes = []
        print("No --lot-prefixes given; matching all lots found under --topax-dir.")
    wafers = set(parse_csv_ints(args.wafers))
    structures = set(parse_csv_list(args.structures))
    include_vds_magnitudes = set(parse_csv_floats(args.include_vds_magnitudes))
    include_packages = [p.upper() for p in parse_csv_list(args.packages)]
    exclude_packages = [p.upper() for p in parse_csv_list(args.exclude_packages)]

    # ------------------------------------------------------------------
    # 2) Discover CSVs and apply the lot / wafer / structure / package filters.
    #    Both Topax (filename-encoded structure) and TPX (Device-column
    #    encoded structure) layouts are scanned; each file only matches one.
    # ------------------------------------------------------------------
    all_files = discover_files(topax_dir) + discover_tpx_files(topax_dir)
    filtered_files = [
        meta
        for meta in all_files
        if lot_selected(meta.lot, lot_prefixes)
        and meta.wafer in wafers
        and structure_selected(meta.structure, structures)
        and (
            not include_packages
            or any(token in meta.package.upper() for token in include_packages)
        )
        and (
            not exclude_packages
            or not any(token in meta.package.upper() for token in exclude_packages)
        )
    ]

    if args.max_files > 0:
        filtered_files = filtered_files[: args.max_files]

    if not filtered_files:
        raise ValueError("No files matched your lot/wafer/structure/package filters.")

    # ------------------------------------------------------------------
    # 3) Resolve output paths. PPT filename is timestamped by default
    #    unless --no-timestamp-ppt is passed.
    # ------------------------------------------------------------------
    plot_dir = args.output_dir / "plots"
    summary_path = args.output_dir / "summary.csv"
    if args.timestamp_ppt:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        ppt_path = args.output_dir / f"Topax_IDVG_4up_{stamp}.pptx"
    else:
        ppt_path = args.output_dir / "Topax_IDVG_4up.pptx"

    image_entries: list[tuple[FileMeta, Path]] = []
    summary_rows: list[dict[str, str]] = []

    args.output_dir.mkdir(parents=True, exist_ok=True)
    plot_dir.mkdir(parents=True, exist_ok=True)

    # ------------------------------------------------------------------
    # 4) Build the immutable render config passed to each worker process.
    # ------------------------------------------------------------------
    cfg = RenderConfig(
        include_vds_magnitudes=frozenset(include_vds_magnitudes),
        keep_ishi=args.keep_ishi,
        keep_islo=args.keep_islo,
        pmos_flip_gate=args.pmos_flip_gate,
        color_by_vds=args.color_by_vds,
        compress_sign=args.compress_vds_sign_by_device,
        plot_both_currents=args.plot_both_currents,
        plot_gate_current=args.plot_gate_current,
        plot_body_current=args.plot_body_current,
        plot_dir=plot_dir,
        y_limits_mode=args.y_limits,
        y_min=args.y_min,
        y_max=args.y_max,
        x_limits_mode=args.x_limits,
        x_min=args.x_min,
        x_max=args.x_max,
        enable_fit_to_atopt=args.enable_fit_to_atopt,
        fit_current_high=args.fit_current_high,
        fit_current_low=args.fit_current_low,
        fit_vds_magnitude=args.fit_vds_magnitude,
    )

    workers = max(1, int(args.workers))
    file_meta_by_plot: dict[Path, FileMeta] = {}

    # ------------------------------------------------------------------
    # 5) Render one PNG per CSV, serially or in a process pool.
    # ------------------------------------------------------------------
    if workers == 1:
        for meta in filtered_files:
            row = render_one_file(meta, cfg)
            if row is None:
                continue
            summary_rows.append(row)
            plot_path = Path(row["plot_path"])
            file_meta_by_plot[plot_path] = meta
            image_entries.append((meta, plot_path))
    else:
        tasks = [(meta, cfg) for meta in filtered_files]
        with ProcessPoolExecutor(max_workers=workers) as pool:
            for meta, row in zip(filtered_files, pool.map(_worker, tasks, chunksize=4)):
                if row is None:
                    continue
                summary_rows.append(row)
                plot_path = Path(row["plot_path"])
                file_meta_by_plot[plot_path] = meta
                image_entries.append((meta, plot_path))

    if not image_entries:
        raise ValueError("Files matched filters, but no data points passed VDS/Sweep filters.")

    # ------------------------------------------------------------------
    # 6) Write the per-file summary CSV.
    # ------------------------------------------------------------------
    with summary_path.open("w", newline="", encoding="utf-8") as handle:
        fieldnames = [
            "lot",
            "wafer",
            "die_x",
            "die_y",
            "structure",
            "device_type",
            "file",
            "points_kept",
            "points_plotted",
            "points_total",
            "plot_path",
        ]
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(summary_rows)

    # ------------------------------------------------------------------
    # 7) Assemble the PowerPoint deck and optionally clean up PNGs.
    # ------------------------------------------------------------------
    deleted_plot_count = 0
    saved_ppt_path = ppt_path
    if not args.skip_ppt:
        saved_ppt_path = make_ppt(image_entries, ppt_path)
        if args.delete_plots_after_ppt:
            deleted_plot_count = delete_generated_plots(image_entries)

    print(f"Topax input folder: {topax_dir}")
    print(f"Matched files: {len(filtered_files)}")
    print(f"Generated plots: {len(image_entries)}")
    print(f"Plot folder: {plot_dir}")
    if not args.skip_ppt:
        print(f"PowerPoint: {saved_ppt_path}")
    if deleted_plot_count > 0:
        print(f"Deleted plot PNG files after PPT generation: {deleted_plot_count}")
    print(f"Summary CSV: {summary_path}")


if __name__ == "__main__":
    main()
