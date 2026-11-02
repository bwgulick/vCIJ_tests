"""
Build every diamond figure into Diamond/figures/ (PDF + PNG).

    python Diamond/run_all.py

Runs the diamond_eos self-test, reads and checks the workbook once, then draws
Figs 1-6 from that one data set.  Any check failure aborts before saving.
"""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

import diamond_eos as de
import dplot_common as dc
import fig1_vv0, fig2_cij_tri, fig3_anisotropy, fig4_moduli, fig5_engineering, fig6_velocities

FIGS = [fig1_vv0, fig2_cij_tri, fig3_anisotropy, fig4_moduli, fig5_engineering,
        fig6_velocities]


def main():
    print("diamond_eos self-test")
    if not de._selftest():
        raise SystemExit("ABORT: diamond_eos self-test failed")
    d = dc.load()
    for mod in FIGS:
        fig = mod.main(d)
        plt.close(fig)
    print(f"done -> {dc.figures_dir()}")


if __name__ == "__main__":
    main()
