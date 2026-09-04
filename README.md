# SuperTOPAX

Unified Topax/TPX E-Test tooling: extraction, IDVG plotting, and wafer sparkline generation.

## Contents

- `Topax_Extractor_GUI.py` — main GUI. Extracts Topax/TPX `.tgz` data from Intel E-Test
  servers into per-die CSVs. Includes **Plot IDVG** and **Wafer Sparkline** buttons that
  launch the two scripts below as independent processes (each uses its own settings —
  edit the constants near the top of each script, or run them directly with CLI flags).
- `topax_idvg_plotter.py` — reads per-die CSVs and renders per-die AbsMeasDrainCurrent
  (and optionally Source/Gate/Body current) vs GateVoltage charts, assembled into a
  4-up PowerPoint deck.
- `topax_wafer_sparkline.py` — renders a per-wafer grid of small Id/Is/Ig sparklines
  (one tile per die), assembled into a PowerPoint deck with one slide per wafer.

## Requirements

- Python 3.10+
- `pandas`, `PyUber` (Intel internal, for the extractor's database query)
- `matplotlib`, `python-pptx` (for the plotting scripts)
- `tkinter` (bundled with standard Python on Windows)

## Running

```
python Topax_Extractor_GUI.py
```

Or run either plotting script directly for full CLI control:

```
python topax_idvg_plotter.py --help
python topax_wafer_sparkline.py --help
```
