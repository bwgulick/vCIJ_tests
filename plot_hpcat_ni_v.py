"""
HPCAT travel-time and length vs pressure for nickel and vanadium.

Data source: VCIJplotdata.xlsx, sheet "HPCAT" (side-by-side blocks; header row 1).
  Nickel     cols A-D : Pressure Ni, L Ni, 2TP Ni, 2TS Ni      (rows 2-7)
  (spacer)   col  E
  Vanadium   cols F-I : P V,        L V,  2TP V, 2TS V          (rows 2-8)

Produces four figures:
  hpcat_tt_nickel.png    : two-way travel time vs P, 2TS (top) / 2TP (bottom)
  hpcat_tt_vanadium.png  : same, vanadium
  hpcat_length_nickel.png   : length vs pressure
  hpcat_length_vanadium.png : length vs pressure

Run:  py plot_hpcat_ni_v.py
"""
import os
import openpyxl
import numpy as np
import matplotlib.pyplot as plt

XLSX = os.environ.get("VPLOT_XLSX", r"C:\Users\bgulick\Downloads\VCIJplotdata.xlsx")
SHEET = "HPCAT"

# element identity colors (colorblind-safe; consistent with the repo's other plots)
COL = {"Nickel": "#2a78d6", "Vanadium": "#eb6834"}


def load():
    """Return {element: dict(P, L, TS, TP)} as float arrays, blanks dropped."""
    wb = openpyxl.load_workbook(XLSX, data_only=True)
    ws = wb[SHEET]
    #                       P, L, 2TP, 2TS  (0-based column indices)
    blocks = {"Nickel": (0, 1, 2, 3), "Vanadium": (5, 6, 7, 8)}
    out = {}
    for name, (cp, cl, ctp, cts) in blocks.items():
        P, L, TS, TP = [], [], [], []
        for row in ws.iter_rows(min_row=2, values_only=True):
            p = row[cp]
            if p is None:
                continue
            P.append(float(p)); L.append(float(row[cl]))
            TP.append(float(row[ctp])); TS.append(float(row[cts]))
        out[name] = dict(P=np.array(P), L=np.array(L),
                         TS=np.array(TS), TP=np.array(TP))
    return out


def figure_travel_time(name, d):
    color = COL[name]
    fig, (ax_s, ax_p) = plt.subplots(
        2, 1, sharex=True, figsize=(7, 7),
        gridspec_kw={"hspace": 0.08}, constrained_layout=True)

    ax_s.plot(d["P"], d["TS"], "-o", color=color, mec="black", mew=0.6, ms=8, lw=1.6)
    ax_p.plot(d["P"], d["TP"], "-o", color=color, mec="black", mew=0.6, ms=8, lw=1.6)

    ax_s.set_ylabel("2·T$_S$  (ns)", fontsize=12)      # S-wave on top
    ax_p.set_ylabel("2·T$_P$  (ns)", fontsize=12)      # P-wave on bottom
    ax_p.set_xlabel("Pressure (GPa)", fontsize=12)

    for ax in (ax_s, ax_p):
        ax.margins(0.06)
    ax_s.set_title(f"{name} — two-way travel time vs pressure", fontsize=13)

    out = f"hpcat_tt_{name.lower()}.png"
    fig.savefig(out, dpi=200, bbox_inches="tight")
    plt.close(fig)
    print(f"{len(d['P'])} points ->", out)


def figure_length(name, d):
    color = COL[name]
    fig, ax = plt.subplots(figsize=(7, 5))
    ax.plot(d["P"], d["L"], "-o", color=color, mec="black", mew=0.6, ms=8, lw=1.6)
    ax.set_xlabel("Pressure (GPa)", fontsize=12)
    ax.set_ylabel("Length (µm)", fontsize=12)
    ax.set_title(f"{name} — length vs pressure", fontsize=13)
    ax.margins(0.06)

    fig.tight_layout()
    out = f"hpcat_length_{name.lower()}.png"
    fig.savefig(out, dpi=200, bbox_inches="tight")
    plt.close(fig)
    print(f"{len(d['P'])} points ->", out)


if __name__ == "__main__":
    data = load()
    for name in ("Nickel", "Vanadium"):
        figure_travel_time(name, data[name])
        figure_length(name, data[name])
