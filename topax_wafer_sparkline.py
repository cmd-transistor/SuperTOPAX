"""Wafer-level sparkline plots for Topax IDVG.

Renders, per wafer, a coordinate-accurate grid of tiny log(Id) vs Vg
sparklines - one small chart per die - overlaying the four NS VT flavors
(HVT/LVT/SVT/UVT) at Vds = +0.65 V only.

Outputs: one PowerPoint deck with one slide per wafer.

Requires the shared helpers in topax_idvg_plotter.py (kept next to this file).
"""

from __future__ import annotations

import argparse
import csv
import math
import os
from concurrent.futures import ProcessPoolExecutor
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import matplotlib

matplotlib.use("Agg")

import matplotlib.pyplot as plt  # noqa: E402
from pptx import Presentation  # noqa: E402
from pptx.util import Inches, Pt  # noqa: E402

from topax_idvg_plotter import (  # noqa: E402
    DEFAULT_TOPAX_DIR,
    DEFAULT_OUTPUT_DIR,
    FileMeta,
    GATE_LINE_COLOR,
    GATE_LINESTYLE,
    SOURCE_LINESTYLE,
    TARGET_CURRENT_NA,
    discover_files,
    discover_tpx_files,
    lot_selected,
    nearly_equal,
    parse_csv_ints,
    parse_csv_list,
)


# ---------------------------------------------------------------------------
# USER CONFIGURATION
# ---------------------------------------------------------------------------
VT_COLORS = {"HVT": "black", "LVT": "blue", "SVT": "green", "UVT": "red"}
VT_ORDER = ["HVT", "SVT", "LVT", "UVT"]

# Fixed axes shared across every sparkline.
X_MIN, X_MAX = -0.5, 1.0
Y_MIN, Y_MAX = 1e-11, 1e-3

# Tick marks shown on every sparkline's axes.
X_TICK_STEP = 0.2
X_TICKS = [round(X_MIN + X_TICK_STEP * i, 10) for i in range(int(round((X_MAX - X_MIN) / X_TICK_STEP)) + 1)]
X_LABELED_TICKS = {X_MIN, 0.0, X_MAX}
Y_TICKS = [1e-11, 1e-9, 1e-7, 1e-5]
Y_TICK_LABELS = ["1e-11", "1e-9", "1e-7", "1e-5"]

DEFAULT_VDS_MAGNITUDE = 0.65
DEFAULT_STRUCTURE_FILTER = "PS-108-A36"
# Device type ('NS' or 'PS') the structure filter/curves apply to; keep in sync
# with the prefix used in DEFAULT_STRUCTURE_FILTER / --structure-substring.
DEFAULT_DEVICE_TYPE = "PS"

# Second reference line used to mark the VTCC extraction current.
VTCC_CURRENT_A = 1e-7
VTCC_COLOR = "magenta"


@dataclass(frozen=True)
class WaferSparklineConfig:
    vds_magnitude: float
    structure_substring: str
    device_type: str
    include_packages: tuple[str, ...]
    exclude_packages: tuple[str, ...]
    plot_dir: Path
    overlay_vt_flavors: bool
    plot_source_current: bool
    plot_gate_current: bool


def extract_vt_flavor(structure: str) -> str | None:
    """Return the HVT/LVT/SVT/UVT token found in a structure name, or None."""
    for token in structure.split("-"):
        upper = token.upper()
        if upper in VT_COLORS:
            return upper
    return None


def read_die_curve(
    meta: FileMeta, vds_magnitude: float, current_kind: str = "drain",
) -> tuple[list[float], list[float]]:
    gv: list[float] = []
    ii: list[float] = []
    target_dv = vds_magnitude if meta.device_type == "NS" else -vds_magnitude

    is_tpx = meta.tpx_device is not None
    col_gate, col_drain_v = ("VG", "VD") if is_tpx else ("GateVoltage", "DrainVoltage")
    col_current = {
        "drain": "ID" if is_tpx else "AbsMeasDrainCurrent",
        "source": "IS" if is_tpx else "AbsMeasSourceCurrent",
        "gate": "IG" if is_tpx else "AbsMeasGateCurrent",
    }[current_kind]

    with meta.path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        if col_current not in (reader.fieldnames or []):
            return [], []
        for row in reader:
            if is_tpx and (row.get("Device") or "").strip() != meta.tpx_device:
                continue
            try:
                x = float((row.get(col_gate) or "").strip())
                dv = float((row.get(col_drain_v) or "").strip())
                y_raw = float((row.get(col_current) or "").strip())
                # TPX current columns are signed; Topax's Abs* columns are already absolute.
                y = abs(y_raw) if is_tpx else y_raw
            except ValueError:
                continue

            if not nearly_equal(dv, target_dv):
                continue
            if y <= 0 or not math.isfinite(y):
                continue

            gv.append(x)
            ii.append(y)

    order = sorted(range(len(gv)), key=lambda i: gv[i])
    gv_sorted = [gv[i] for i in order]
    ii_sorted = [ii[i] for i in order]
    return gv_sorted, ii_sorted


def group_files_by_wafer_die_vt(
    files: list[FileMeta], device_type: str
) -> dict[int, dict[tuple[int, int], dict[str, FileMeta]]]:
    result: dict[int, dict[tuple[int, int], dict[str, FileMeta]]] = {}
    for meta in files:
        if meta.device_type != device_type:
            continue
        vt = extract_vt_flavor(meta.structure)
        if vt is None:
            continue
        wafer_map = result.setdefault(meta.wafer, {})
        die_map = wafer_map.setdefault((meta.die_x, meta.die_y), {})
        die_map.setdefault(vt, meta)  # first occurrence wins if duplicates
    return result


def wafer_grid_extents(die_map: dict[tuple[int, int], dict[str, FileMeta]]) -> tuple[int, int, int, int]:
    xs = [k[0] for k in die_map.keys()]
    ys = [k[1] for k in die_map.keys()]
    return min(xs), max(xs), min(ys), max(ys)


def render_wafer_png(
    wafer: int,
    die_map: dict[tuple[int, int], dict[str, FileMeta]],
    cfg: WaferSparklineConfig,
) -> Path:
    x_min, x_max, y_min, y_max = wafer_grid_extents(die_map)
    n_cols = x_max - x_min + 1
    n_rows = y_max - y_min + 1

    # Small tile per die; total figure grows with wafer die extents.
    # Fixed PNG width (inches); height scales with grid aspect + fixed padding for title/legend.
    tile_w, tile_h = 0.9, 0.75
    fig_w = 9.0
    if n_cols > 0:
        grid_h = fig_w * (n_rows / n_cols) * (tile_h / tile_w)
    else:
        grid_h = 4.5
    fig_h = max(4.5, grid_h + 1.2)

    fig, axes = plt.subplots(
        n_rows,
        n_cols,
        figsize=(fig_w, fig_h),
        dpi=150,
        sharex=True,
        sharey=True,
        squeeze=False,
    )

    device_type = cfg.device_type
    target_a = TARGET_CURRENT_NA[device_type] * 1e-9

    for row_i in range(n_rows):
        die_y = y_max - row_i  # top of figure = highest die_y
        for col_i in range(n_cols):
            die_x = x_min + col_i
            ax = axes[row_i][col_i]
            ax.set_xlim(X_MIN, X_MAX)
            ax.set_ylim(Y_MIN, Y_MAX)
            ax.set_yscale("log")

            ax.set_xticks(X_TICKS)
            ax.set_xticklabels(
                [f"{v:g}" if v in X_LABELED_TICKS else "" for v in X_TICKS], fontsize=4
            )
            ax.set_yticks(Y_TICKS)
            ax.set_yticklabels(Y_TICK_LABELS, fontsize=4)
            ax.tick_params(axis="both", which="major", length=2, labelsize=4, pad=1)
            ax.tick_params(axis="both", which="minor", length=0)
            ax.grid(True, which="major", axis="both", linestyle="-", linewidth=0.3, alpha=0.4, color="lightgray")
            for spine in ax.spines.values():
                spine.set_linewidth(0.4)
                spine.set_color("lightgray")

            per_die = die_map.get((die_x, die_y))
            if per_die is None:
                ax.set_facecolor("#f7f7f7")
                continue

            ax.axhline(target_a, color="red", linestyle="--", linewidth=0.6, alpha=0.8)
            ax.axhline(VTCC_CURRENT_A, color=VTCC_COLOR, linestyle="--", linewidth=0.6, alpha=0.8)
            ax.axvline(0.0, color="gray", linestyle="--", linewidth=0.5, alpha=0.8)

            flavors_to_plot = VT_ORDER if cfg.overlay_vt_flavors else [next(iter(per_die.keys()))]
            for vt in flavors_to_plot:
                meta = per_die.get(vt)
                if meta is None:
                    continue
                gv, ii = read_die_curve(meta, cfg.vds_magnitude, "drain")
                if gv:
                    ax.plot(gv, ii, color=VT_COLORS[vt], linewidth=0.7)

                if cfg.plot_source_current:
                    gv_s, ii_s = read_die_curve(meta, cfg.vds_magnitude, "source")
                    if gv_s:
                        ax.plot(gv_s, ii_s, color=VT_COLORS[vt], linestyle=SOURCE_LINESTYLE, linewidth=0.5)

                if cfg.plot_gate_current:
                    gv_g, ii_g = read_die_curve(meta, cfg.vds_magnitude, "gate")
                    if gv_g:
                        ax.plot(gv_g, ii_g, color=GATE_LINE_COLOR, linestyle=GATE_LINESTYLE, linewidth=0.5)

            ax.text(
                0.02,
                0.98,
                f"({die_x},{die_y})",
                transform=ax.transAxes,
                fontsize=5,
                color="dimgray",
                verticalalignment="top",
            )

    lots = sorted({meta.lot for per_die in die_map.values() for meta in per_die.values()})
    lot_str = "/".join(lots)
    current_parts = ["Id solid"]
    if cfg.plot_source_current:
        current_parts.append("Is dotted")
    if cfg.plot_gate_current:
        current_parts.append("Ig red densely-dotted")
    current_note = ", ".join(current_parts)
    fig.suptitle(
        f"{lot_str} | Wafer w{wafer} - IDVG sparklines @ |Vds|={cfg.vds_magnitude:g}V ({cfg.structure_substring}) | {current_note}",
        fontsize=11,
        y=0.995,
    )

    legend_handles = [
        plt.Line2D([0], [0], color=color, linewidth=1.2, label=vt)
        for vt, color in VT_COLORS.items()
    ]
    if cfg.plot_source_current:
        legend_handles.append(
            plt.Line2D([0], [0], color="black", linestyle=SOURCE_LINESTYLE, linewidth=1.2, label="Source (Is)")
        )
    if cfg.plot_gate_current:
        legend_handles.append(
            plt.Line2D([0], [0], color=GATE_LINE_COLOR, linestyle=GATE_LINESTYLE, linewidth=1.2, label="Gate (Ig)")
        )
    target_na = TARGET_CURRENT_NA[device_type]
    legend_handles.append(
        plt.Line2D([0], [0], color="red", linestyle="--", linewidth=1.0, label=f"ATOPT {target_na:g} nA")
    )
    legend_handles.append(
        plt.Line2D([0], [0], color=VTCC_COLOR, linestyle="--", linewidth=1.0, label=f"VTCC {VTCC_CURRENT_A:.0e} A")
    )
    legend_handles.append(
        plt.Line2D([0], [0], color="gray", linestyle="--", linewidth=1.0, label="Vg = 0 V")
    )
    fig.legend(
        handles=legend_handles,
        loc="lower center",
        ncol=len(legend_handles),
        fontsize=7,
        frameon=False,
        bbox_to_anchor=(0.5, 0.005),
    )

    fig.tight_layout(rect=[0, 0.05, 1, 0.97])
    fig.subplots_adjust(top=0.965, hspace=0.15, wspace=0.15)

    cfg.plot_dir.mkdir(parents=True, exist_ok=True)
    out_path = cfg.plot_dir / f"wafer_w{wafer}_sparklines.png"
    fig.savefig(out_path)
    plt.close(fig)
    return out_path


def _render_worker(args_tuple: tuple[int, dict[tuple[int, int], dict[str, FileMeta]], WaferSparklineConfig]) -> tuple[int, Path]:
    wafer, die_map, cfg = args_tuple
    return wafer, render_wafer_png(wafer, die_map, cfg)


def add_fitted_picture(slide, image_path: Path, box_left, box_top, box_width, box_height) -> None:
    pic = slide.shapes.add_picture(str(image_path), 0, 0)
    original_w = pic.width
    original_h = pic.height
    if original_w <= 0 or original_h <= 0:
        pic.left = box_left
        pic.top = box_top
        return

    scale = min(box_width / original_w, box_height / original_h)
    new_w = int(original_w * scale)
    new_h = int(original_h * scale)

    pic.width = new_w
    pic.height = new_h
    pic.left = int(box_left + (box_width - new_w) / 2)
    pic.top = int(box_top + (box_height - new_h) / 2)


def build_ppt(
    wafer_images: list[tuple[int, Path]],
    wafer_lots: dict[int, str],
    ppt_path: Path,
    cfg: WaferSparklineConfig,
) -> Path:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    for wafer, image_path in sorted(wafer_images, key=lambda t: t[0]):
        slide = prs.slides.add_slide(blank_layout)

        lot_str = wafer_lots.get(wafer, "")
        title_text = (
            f"{lot_str} | w{wafer} - IDVG Sparklines @ |Vds|={cfg.vds_magnitude:g}V ({cfg.structure_substring})"
            if lot_str
            else f"w{wafer} - IDVG Sparklines @ |Vds|={cfg.vds_magnitude:g}V ({cfg.structure_substring})"
        )
        title_box = slide.shapes.add_textbox(Inches(0.0), Inches(0.0), prs.slide_width, Inches(0.5))
        tf = title_box.text_frame
        tf.clear()
        run = tf.paragraphs[0].add_run()
        run.text = title_text
        run.font.size = Pt(14)

        content_left = Inches(0.25)
        content_top = Inches(0.5)
        content_w = prs.slide_width - Inches(0.5)
        content_h = Inches(6.5)

        add_fitted_picture(
            slide=slide,
            image_path=image_path,
            box_left=content_left,
            box_top=content_top,
            box_width=content_w,
            box_height=content_h,
        )

    ppt_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        prs.save(ppt_path)
        return ppt_path
    except PermissionError:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        fallback = ppt_path.with_name(f"{ppt_path.stem}_{stamp}{ppt_path.suffix}")
        prs.save(fallback)
        return fallback


def delete_pngs(images: list[tuple[int, Path]]) -> int:
    deleted = 0
    dirs: set[Path] = set()
    for _, p in images:
        if p.exists():
            p.unlink()
            deleted += 1
        dirs.add(p.parent)
    for d in dirs:
        if d.exists() and not any(d.iterdir()):
            try:
                d.rmdir()
            except OSError:
                pass
    return deleted


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Wafer-level sparkline plots for Topax IDVG (log Id vs Vg at Vds=0.65V).",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )

    # Input / output paths
    io_group = parser.add_argument_group("Input / output paths")
    io_group.add_argument("--topax-dir", type=Path, default=DEFAULT_TOPAX_DIR, help="Root folder of Topax CSVs (recursive).")
    io_group.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT_DIR, help="Where the PowerPoint deck is written.")

    # Data selection
    sel_group = parser.add_argument_group("Data selection")
    sel_group.add_argument("--lot-prefixes", default="D617F5X", help="Comma-separated lot-name prefixes.")
    sel_group.add_argument(
        "--wafers",
        default="250,281,353,375",
        help="Comma-separated wafer numbers. Use 'all' to include every wafer discovered.",
    )
    sel_group.add_argument(
        "--structure-substring",
        default=DEFAULT_STRUCTURE_FILTER,
        help="Substring the structure name must contain (case-sensitive), e.g. 'NS-108-A36'.",
    )
    sel_group.add_argument(
        "--device-type",
        choices=["NS", "PS"],
        default=DEFAULT_DEVICE_TYPE,
        help="Device type the structure filter/curves apply to (must match --structure-substring's prefix).",
    )
    sel_group.add_argument(
        "--packages",
        default="",
        help=(
            "Comma-separated package tokens that must appear in the filename package field "
            "(case-insensitive). TPX files always carry package='TPX', so leave empty to include them."
        ),
    )
    sel_group.add_argument(
        "--exclude-packages",
        default="QREL",
        help="Comma-separated package substrings to exclude (case-insensitive).",
    )

    # Plot options
    plot_group = parser.add_argument_group("Plot options")
    plot_group.add_argument("--vds-magnitude", type=float, default=DEFAULT_VDS_MAGNITUDE, help="|Vds| to plot in each sparkline.")
    plot_group.add_argument(
        "--overlay-vt-flavors",
        action="store_true",
        default=True,
        help="Overlay HVT/LVT/SVT/UVT curves in each die's sparkline.",
    )
    plot_group.add_argument(
        "--no-overlay-vt-flavors",
        dest="overlay_vt_flavors",
        action="store_false",
        help="Plot only whichever single VT flavor is present per die.",
    )
    plot_group.add_argument(
        "--plot-source-current",
        action="store_true",
        default=True,
        help="Overlay source current (Is) as a dotted line, same color as its VT flavor.",
    )
    plot_group.add_argument(
        "--no-plot-source-current",
        dest="plot_source_current",
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

    # Output behavior
    out_group = parser.add_argument_group("Output behavior")
    out_group.add_argument(
        "--timestamp-ppt",
        action="store_true",
        default=False,
        help="Append a YYYYMMDD_HHMMSS timestamp to the PowerPoint filename.",
    )
    out_group.add_argument(
        "--no-timestamp-ppt",
        dest="timestamp_ppt",
        action="store_false",
        help="Save the deck as Topax_Wafer_Sparklines.pptx and overwrite in place.",
    )
    out_group.add_argument(
        "--keep-pngs",
        action="store_true",
        default=False,
        help="Keep per-wafer PNG files after the deck is generated (deleted by default).",
    )
    out_group.add_argument(
        "--workers",
        type=int,
        default=max(1, (os.cpu_count() or 4) - 1),
        help="Parallel worker processes for rendering wafer PNGs.",
    )

    return parser.parse_args()


def main() -> None:
    args = parse_args()

    topax_dir: Path = args.topax_dir
    if not topax_dir.exists():
        raise FileNotFoundError(f"Topax directory not found: {topax_dir}")

    lot_prefixes = parse_csv_list(args.lot_prefixes)
    if args.wafers.strip().lower() == "all":
        wafer_filter: set[int] | None = None
    else:
        wafer_filter = set(parse_csv_ints(args.wafers))
    include_pkgs = [p.upper() for p in parse_csv_list(args.packages)]
    exclude_pkgs = [p.upper() for p in parse_csv_list(args.exclude_packages)]

    all_files = discover_files(topax_dir) + discover_tpx_files(topax_dir)
    filtered = [
        m
        for m in all_files
        if lot_selected(m.lot, lot_prefixes)
        and (wafer_filter is None or m.wafer in wafer_filter)
        and args.structure_substring in m.structure
        and m.device_type == args.device_type
        and (not include_pkgs or any(t in m.package.upper() for t in include_pkgs))
        and (not exclude_pkgs or not any(t in m.package.upper() for t in exclude_pkgs))
    ]

    if not filtered:
        raise ValueError("No files matched your lot/wafer/structure/package filters.")

    grouped = group_files_by_wafer_die_vt(filtered, args.device_type)
    if not grouped:
        raise ValueError("Files matched but no VT flavor could be extracted from structure names.")

    wafer_lots: dict[int, str] = {}
    for wafer, die_map in grouped.items():
        lots = sorted({meta.lot for per_die in die_map.values() for meta in per_die.values()})
        wafer_lots[wafer] = "/".join(lots)

    plot_dir = args.output_dir / "wafer_sparkline_pngs"
    if args.timestamp_ppt:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        ppt_path = args.output_dir / f"Topax_Wafer_Sparklines_{stamp}.pptx"
    else:
        ppt_path = args.output_dir / "Topax_Wafer_Sparklines.pptx"

    args.output_dir.mkdir(parents=True, exist_ok=True)
    plot_dir.mkdir(parents=True, exist_ok=True)

    cfg = WaferSparklineConfig(
        vds_magnitude=args.vds_magnitude,
        structure_substring=args.structure_substring,
        device_type=args.device_type,
        include_packages=tuple(include_pkgs),
        exclude_packages=tuple(exclude_pkgs),
        plot_dir=plot_dir,
        overlay_vt_flavors=args.overlay_vt_flavors,
        plot_source_current=args.plot_source_current,
        plot_gate_current=args.plot_gate_current,
    )

    workers = max(1, int(args.workers))
    wafer_images: list[tuple[int, Path]] = []

    tasks = [(wafer, die_map, cfg) for wafer, die_map in grouped.items()]

    if workers == 1 or len(tasks) == 1:
        for t in tasks:
            wafer_images.append(_render_worker(t))
    else:
        with ProcessPoolExecutor(max_workers=workers) as pool:
            for result in pool.map(_render_worker, tasks):
                wafer_images.append(result)

    saved_ppt = build_ppt(wafer_images, wafer_lots, ppt_path, cfg)

    deleted = 0
    if not args.keep_pngs:
        deleted = delete_pngs(wafer_images)

    print(f"Topax input folder: {topax_dir}")
    print(f"Wafers rendered: {sorted(w for w, _ in wafer_images)}")
    print(f"PowerPoint: {saved_ppt}")
    if deleted:
        print(f"Deleted per-wafer PNG files: {deleted}")


if __name__ == "__main__":
    main()
