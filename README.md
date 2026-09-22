# vCIJ_tests

Vanadium single-crystal elastic-constant (Cij) analysis: ultrasonic travel-time
measurements, finite-strain and Cook's-method fits, uncertainty propagation, and the
figures for the paper.

Run everything with **Python 3.11** (`...\Programs\Python\Python311\python.exe`) —
that's where `pandas`, `openpyxl` and `scipy` are installed.

## Layout

| folder | what's in it |
|---|---|
| `apps/` | the GUIs — see below |
| `scripts/` | standalone one-off plot scripts, run directly |
| `tools/` | live workbook maintenance (Excel COM / openpyxl) |
| `workbooks/` | the two live analysis workbooks |
| `data/` | raw measurement inputs |
| `docs/` | papers and guides |
| `figures/` | all generated figures land here |
| `archive/` | retired scripts, probes and superseded workbooks |

## Apps

| launch | what it does |
|---|---|
| `py apps/vplot/vplot_gui.py` | main Cij / moduli / V-V0 plotting GUI |
| `py apps/WaveformPlotter.py` | waveform viewer for .xlsx / .csv |
| `py apps/Baosheng_PEO.py` | pulse-echo-overlap travel-time measurement |
| `py apps/plot_digitizer_gui.py` | pull data points off a figure image |
| `py apps/make_gif_gui.py` | build an animated GIF from an image sequence |

`apps/vplot/` is one unit: `vplot_gui.py` imports `vplot_common.py` and the four
`plot_*.py` render modules by flat import, so those six files must stay in the same
folder. Each `plot_*.py` also runs standalone and writes its figure to `figures/`.

## Data sources

The vplot GUI and `scripts/plot_hpcat_ni_v.py` read **`VCIJplotdata.xlsx`**, which
lives **outside this repo** at `C:\Users\bgulick\Desktop\V_single_figures\`. Override
with the `VPLOT_XLSX` environment variable, or browse to it from the GUI.

Everything else reads from `data/` or `workbooks/` by a path relative to the script, so
the scripts work from any working directory.

## Workbooks — read this before editing

- **`workbooks/Vanadium_Cij_finite_strain_all_equations_GlobalMin - Copy.xlsx`** is the
  *main* finite-strain fit workbook, despite the " - Copy" in its name. It holds 3
  embedded charts, so edit it **only via Excel COM** — saving it with openpyxl silently
  drops the charts.
- **`workbooks/V_Cij.xlsx`** is the per-point Cij / uncertainty analysis workbook.
- The ~264 `#REF!` cells in `V_Cij.xlsx` are **pre-existing and expected**. They are the
  σ and ρV² columns for the one *reconstructed* (omitted) velocity on each combo sheet.
  Don't guess-repair them; a real fix needs a modeling decision about that velocity.

## Tools

All default to the right workbook via `$PSScriptRoot`, so run them from anywhere.

| tool | when to run it |
|---|---|
| `tools/_solveraid.ps1` | **after ANY re-fit** — the standard errors it writes are only valid at the current converged optimum |
| `tools/_make_plotdata.ps1` | rebuild the `Plot Data` sheet; `-Src "<sheet>"` to change the source combo |
| `tools/_uncert_gen_step1.py` → `tools/_apply_spec.ps1` | regenerate the per-point σ block (writes a `.tsv` spec, then applies it via COM) |
| `tools/_uncert_gen_step3.py` → `tools/_apply_spec.ps1` | regenerate the combined-final rollups |
| `tools/_build_katahara.ps1` | rebuild the `Katahara extrap` sheet (backs up to `archive/workbooks/` first) |

## Archive

Nothing here is needed to reproduce current results, but it's kept in place and in git
history:

- `archive/probes/` — throwaway `_inspect*.py` workbook probes (they reference a
  workbook name that no longer exists) plus unzipped `.xlsx` internals and text dumps.
- `archive/build/` — the original one-shot pipeline that built the GlobalMin workbook.
- `archive/workbooks/` — superseded and timestamped backup workbooks.
