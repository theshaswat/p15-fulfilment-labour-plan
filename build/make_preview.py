"""Draws docs/dashboard.png from the same model that generates the workbook."""
import sys, datetime as dt
sys.path.insert(0, "build")
import matplotlib; matplotlib.use("Agg")
import matplotlib.pyplot as plt, matplotlib.dates as md
from build_xlsx import model, summary, START, N, SCEN

INK, COPPER, PAPER, RULE = "#22303C", "#B8642B", "#FAF7F1", "#E3DDD0"
plt.rcParams.update({"font.family": "Helvetica", "axes.edgecolor": RULE, "axes.labelcolor": INK, "xtick.color": INK, "ytick.color": INK})
r = model(1.0); days = [START + dt.timedelta(i) for i in range(N)]
fig, ax = plt.subplots(2, 1, figsize=(10, 6.4), facecolor=PAPER, gridspec_kw={"height_ratios": [3, 2], "hspace": .38})
a = ax[0]; a.set_facecolor(PAPER)
a.plot(days, [x["J"] for x in r], color=COPPER, lw=2.2, label="Required headcount")
a.axhline(r[0]["K"], color=INK, lw=1.6, ls=(0, (5, 3)), label="Permanent roster on duty (daily average)")
a.set_title("Headcount required against the permanent roster, Base scenario", loc="left", fontsize=11.5, color=INK, fontweight="bold")
a.set_ylim(0, 300); a.yaxis.grid(True, color=RULE); a.set_axisbelow(True)
a.legend(frameon=False, loc="lower left", ncol=2, fontsize=9)
b = ax[1]; b.set_facecolor(PAPER)
b.bar(days, [x["Q"] for x in r], color=COPPER, width=.75)
b.set_title("Flex hires needed per day", loc="left", fontsize=11.5, color=INK, fontweight="bold")
b.yaxis.grid(True, color=RULE); b.set_axisbelow(True)
for x in ax:
    for s in ("top", "right"): x.spines[s].set_visible(False)
    x.xaxis.set_major_formatter(md.DateFormatter("%d %b")); x.xaxis.set_major_locator(md.WeekdayLocator(byweekday=0))
    x.tick_params(labelsize=8.5)
fig.savefig("docs/dashboard.png", dpi=170, bbox_inches="tight", facecolor=PAPER)
