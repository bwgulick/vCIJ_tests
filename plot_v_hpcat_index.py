"""
Length vs acquisition index over the final 7000 PSI ramp + the 7000/5000/3000
PSI heat-cycle holds. Each hold keeps one fixed hue (color = pressure
identity); within a hold, the shade goes light (cooler) -> dark (hotter),
normalized to that hold's own temperature range, so the heat-up/cool-down
excursion is visible as a light-dark-light sweep along the dashed connector.
Temperature is also printed at each point for exact values.
"""
import colorsys
import openpyxl
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.colors import LinearSegmentedColormap, Normalize, to_rgba

XLSX, SHEET = "v_hpcat.xlsx", "July (2)"
C_PSI, C_T, C_LEN = 1, 3, 10          # 0-based: Pressure(PSI), Temp calib(C), Distance length(um)
START = 7                              # first row of the final 7000 PSI ramp + heat cycling

wb = openpyxl.load_workbook(XLSX, data_only=True)
ws = wb[SHEET]
P, T, L = [], [], []
for row in ws.iter_rows(min_row=2, values_only=True):
    psi, t, length = row[C_PSI], row[C_T], row[C_LEN]
    if psi is None or length is None or t is None:
        continue
    P.append(float(psi)); T.append(float(t)); L.append(float(length))
P, T, L = np.array(P), np.array(T), np.array(L)
P, T, L = P[START:], T[START:], L[START:]
x = np.arange(START, START + len(P))

# one fixed hue per pressure hold (categorical identity); anything past the
# top 3 holds folds into a neutral "other" bucket rather than adding a 4th
# hue (scatter color only stays pairwise-distinguishable for 3 series)
HUES = ["#2a78d6", "#eb6834", "#1baf7a"]      # blue (7000), orange (5000), aqua (3000)
OTHER = "#898781"                              # muted gray for stray readings

uP = sorted(set(P), reverse=True)              # highest pressure first
hue_of = {p: HUES[i] for i, p in enumerate(uP[:len(HUES)])}


def light_dark_ramp(hex_color, l_light=0.88, l_dark=0.24):
    """Light-tint -> base hue -> dark-shade colormap, same hue throughout."""
    r, g, b = (int(hex_color[i:i + 2], 16) / 255 for i in (1, 3, 5))
    h, _, s = colorsys.rgb_to_hls(r, g, b)
    light = colorsys.hls_to_rgb(h, l_light, s * 0.85)
    dark = colorsys.hls_to_rgb(h, l_dark, min(1.0, s * 1.15))
    return LinearSegmentedColormap.from_list(f"ramp_{hex_color}", [light, (r, g, b), dark])


ramps = {p: light_dark_ramp(c) for p, c in hue_of.items()}

# shade each point by temperature, normalized within its own hold so every
# hold shows its full light(cool)->dark(hot) excursion regardless of how hot
# that particular hold actually got
pt_c = [None] * len(P)
for p in uP:
    idx = np.where(P == p)[0]
    if p in ramps:
        tt = T[idx]
        vmin, vmax = tt.min(), tt.max()
        norm = Normalize(vmin=vmin, vmax=vmax if vmax > vmin else vmin + 1.0)
        for i, col in zip(idx, ramps[p](norm(tt))):
            pt_c[i] = col
    else:
        for i in idx:
            pt_c[i] = to_rgba(OTHER)

fig, ax = plt.subplots(figsize=(11, 6))
ax.plot(x, L, "--", color="0.5", lw=0.9, zorder=1)      # dashed connector, sheet order
ax.scatter(x, L, c=pt_c, s=70, edgecolor="black", linewidth=0.6, zorder=3)

# label each point with its temperature
for xi, li, ti in zip(x, L, T):
    ax.annotate(f"{ti:.0f}", (xi, li), textcoords="offset points",
                xytext=(4, 5), fontsize=8, color="0.15")

ax.set_xlabel("Acquisition index", fontsize=12)
ax.set_ylabel("Length (µm)", fontsize=12)
ax.set_title("Vanadium length vs step (hue = pressure, shade = temperature, label = °C)",
             fontsize=12)
ax.grid(alpha=0.3)
ax.margins(0.04)

# pressure legend (base hue per hold; gray for stray/other readings)
handles = [plt.Line2D([], [], marker="o", ls="", mec="black", mew=0.6,
                      mfc=hue_of.get(p, OTHER),
                      label=f"{int(p)}" if p in hue_of else f"{int(p)} (single reading)")
           for p in uP]
ax.legend(handles=handles, title="PSI", fontsize=8, title_fontsize=9,
          loc="center left", bbox_to_anchor=(1.01, 0.5), frameon=False)

fig.text(0.5, -0.01,
         "Shade = temperature within each pressure hold (light = cooler → dark = hotter); "
         "numbers label °C.",
         ha="center", fontsize=8, color="0.35")

fig.tight_layout()
out = "v_hpcat_index.png"
fig.savefig(out, dpi=200, bbox_inches="tight")
print(f"{len(P)} points ->", out)
