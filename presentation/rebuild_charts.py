import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import os

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "charts")

ACCENT = "#6fa8d6"
AMBER = "#e6a94a"
NEUTRAL = "#6d7982"
CAUTION = "#e08a5c"
TEXT = "#aab4bc"
SPINE = "#3a4650"
GRID = "#1c242b"

plt.rcParams["font.family"] = ["DejaVu Sans", "Arial", "Helvetica"]

WIDE_W, WIDE_H = 12.13, 3.7  # matches the wide-slide slot exactly -- no stretch


def style_ax(ax):
    ax.patch.set_alpha(0)
    ax.tick_params(colors=TEXT, labelsize=10)
    ax.spines[["top", "right"]].set_visible(False)
    for sp in ["left", "bottom"]:
        ax.spines[sp].set_color(SPINE)
    ax.xaxis.label.set_color(TEXT)
    ax.yaxis.label.set_color(TEXT)


# ---------------------------------------------------------------- story3_chart_b
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(WIDE_W, WIDE_H), dpi=200)
fig.patch.set_alpha(0)

fraud = {"Western Europe": 2.61, "Canada": 3.02}
fraud_avg = 2.1
colors1 = [ACCENT, ACCENT]
bars1 = ax1.barh(list(fraud.keys()), list(fraud.values()), color=colors1, height=0.5)
ax1.axvline(fraud_avg, color=NEUTRAL, linestyle="--", linewidth=1.2)
ax1.text(fraud_avg, 0.97, f"avg {fraud_avg}%", color=NEUTRAL, fontsize=9, ha="center",
         va="top", transform=ax1.get_xaxis_transform())
ax1.set_xlabel("Fraud rate (%)")
ax1.invert_yaxis()
style_ax(ax1)

late = {"Central Africa": 57.96, "South Asia": 56.31, "Canada": 48.80}
late_avg = 54.6
colors2 = [CAUTION if v > late_avg else NEUTRAL for v in late.values()]
ax2.barh(list(late.keys()), list(late.values()), color=colors2, height=0.5)
ax2.axvline(late_avg, color=NEUTRAL, linestyle="--", linewidth=1.2)
ax2.text(late_avg, 0.97, f"avg {late_avg}%", color=NEUTRAL, fontsize=9, ha="center",
         va="top", transform=ax2.get_xaxis_transform())
ax2.set_xlabel("Late-delivery rate (%)")
ax2.invert_yaxis()
style_ax(ax2)

plt.tight_layout(pad=1.2)
plt.savefig(os.path.join(OUT, "story3_chart_b.png"), transparent=True, dpi=200)
plt.close(fig)

# ---------------------------------------------------------------- story4_chart_a
import pandas as pd
import numpy as np
import matplotlib.colors as mcolors

df4 = pd.read_csv(os.path.join(HERE, "..", "DataCoSupplyChainDataset.csv"), encoding="latin1",
                   usecols=["Order Region", "order date (DateOrders)", "Order Id"])
df4["order date (DateOrders)"] = pd.to_datetime(df4["order date (DateOrders)"])
df4["month"] = df4["order date (DateOrders)"].dt.to_period("M")
df4 = df4[df4["month"] <= "2017-12"]  # excludes the partial Jan 2018 month, matches "of 36" elsewhere in deck

pivot = df4.groupby(["Order Region", "month"])["Order Id"].nunique().unstack(fill_value=0)
active = (pivot > 0).astype(int)
first_active = active.apply(lambda r: r[r == 1].index.min(), axis=1)
active = active.loc[first_active.sort_values().index]

fig, ax = plt.subplots(figsize=(WIDE_W, WIDE_H), dpi=200)
fig.patch.set_alpha(0)
cmap = mcolors.ListedColormap([GRID, ACCENT])
ax.imshow(active.values, aspect="auto", interpolation="nearest", cmap=cmap, vmin=0, vmax=1)
ax.set_yticks(np.arange(len(active.index)))
ax.set_yticklabels(active.index, fontsize=7.6, color=TEXT)
month_labels = [str(m) for m in active.columns]
step = 2
ax.set_xticks(np.arange(0, len(month_labels), step))
ax.set_xticklabels([month_labels[i] for i in range(0, len(month_labels), step)],
                    rotation=90, fontsize=7, color=TEXT)
ax.set_xlabel("Month active (blue = ≥1 order)", color=TEXT, fontsize=9.5)
for spine in ax.spines.values():
    spine.set_visible(False)
ax.tick_params(length=0)
plt.tight_layout(pad=0.8)
plt.savefig(os.path.join(OUT, "story4_chart_a.png"), transparent=True, dpi=200)
plt.close(fig)

# ---------------------------------------------------------------- headline_chart (slide 3)
years = ["2015", "2016", "2017"]
sales_m = [12.34, 12.30, 11.81]
profit_m = [1.32, 1.31, 1.30]

fig, ax = plt.subplots(figsize=(7.6, 4.1), dpi=200)
fig.patch.set_alpha(0)
x = range(len(years))
w = 0.34
b1 = ax.bar([i - w / 2 for i in x], sales_m, width=w, color=ACCENT, label="Total sales ($M)")
b2 = ax.bar([i + w / 2 for i in x], profit_m, width=w, color=AMBER, label="Profit ($M)")
for rect in list(b1) + list(b2):
    h = rect.get_height()
    ax.text(rect.get_x() + rect.get_width() / 2, h + 0.18, f"{h:.2f}", ha="center", va="bottom",
            fontsize=9.5, color=TEXT)
ax.set_xticks(list(x))
ax.set_xticklabels(years)
ax.set_ylim(0, max(sales_m) * 1.22)
leg = ax.legend(loc="upper right", frameon=False, fontsize=10, labelcolor=TEXT)
style_ax(ax)
plt.tight_layout(pad=1.0)
plt.savefig(os.path.join(OUT, "headline_chart.png"), transparent=True, dpi=200)
plt.close(fig)

print("done")
