from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils.datetime import from_excel

import matplotlib.dates as mdates
import matplotlib.pyplot as plt


SOURCE = Path(
    "/Users/USER/Documents/Codex/2026-06-21/new-chat-3/work/qc_2026_input/2026/断头趋势.xlsx"
)
OUT = Path(
    "/Users/USER/Documents/Codex/2026-06-21/new-chat-3/work/spreadsheets/qcc_trend_source_excel/qa/trend_chart_from_source.png"
)


def read_rows():
    wb = load_workbook(SOURCE, data_only=True)
    ws = wb["Sheet1"]
    rows = []
    for r in range(6, 43):
        raw_date = ws.cell(r, 1).value
        date = from_excel(raw_date).date() if isinstance(raw_date, (int, float)) else raw_date
        if date.isoformat() < "2026-05-16":
            continue
        rows.append(
            {
                "date": date,
                "breaks": ws.cell(r, 2).value,
                "ccut": ws.cell(r, 3).value,
            }
        )
    return rows


rows = read_rows()
dates = [r["date"] for r in rows]
breaks = [r["breaks"] for r in rows]
ccut = [r["ccut"] for r in rows]

plt.rcParams["font.sans-serif"] = [
    "PingFang SC",
    "Heiti TC",
    "Arial Unicode MS",
    "Noto Sans CJK SC",
    "DejaVu Sans",
]
plt.rcParams["axes.unicode_minus"] = False

fig, ax = plt.subplots(figsize=(12, 5.4), dpi=180)
fig.patch.set_facecolor("white")
ax.set_facecolor("white")

ax.plot(dates, breaks, color="#B73E3E", marker="o", linewidth=2.4, markersize=4.5, label="断头数")
ax.plot(dates, ccut, color="#1F6FAE", marker="o", linewidth=2.2, markersize=4.2, label="C切数")
ax.axhline(90, color="#C45B5B", linewidth=1.5, linestyle=(0, (5, 3)), label="断头目标90")
ax.axhline(50, color="#4D8DCC", linewidth=1.5, linestyle=(0, (5, 3)), label="C切目标50")

node = mdates.datestr2num("2026-05-25")
ax.axvline(node, color="#13795B", linewidth=2, linestyle=":")
ax.text(node, 184, "0.60%整体执行", color="#13795B", fontsize=10, va="top", ha="left")

ax.set_title("近一个月断头/C切趋势", fontsize=15, pad=14)
ax.set_ylabel("数量（个/日）", fontsize=11)
ax.set_ylim(0, max(max(breaks), 200) + 10)
ax.grid(True, axis="y", color="#E5E7EB", linewidth=0.8)
ax.grid(True, axis="x", color="#F0F2F5", linewidth=0.6)

ax.xaxis.set_major_locator(mdates.DayLocator(interval=3))
ax.xaxis.set_major_formatter(mdates.DateFormatter("%-m/%-d"))
ax.tick_params(axis="x", labelsize=9)
ax.tick_params(axis="y", labelsize=9)
ax.legend(loc="upper right", ncol=4, frameon=False, fontsize=9)

for spine in ["top", "right"]:
    ax.spines[spine].set_visible(False)
ax.spines["left"].set_color("#CBD5E1")
ax.spines["bottom"].set_color("#CBD5E1")

OUT.parent.mkdir(parents=True, exist_ok=True)
fig.tight_layout()
fig.savefig(OUT, bbox_inches="tight")
plt.close(fig)
print(OUT)
