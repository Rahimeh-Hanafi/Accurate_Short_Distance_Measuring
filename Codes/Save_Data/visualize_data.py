import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

# ---- Read the Excel file ----
df = pd.read_excel("MeasuredRanges.xlsx")

# ---- Calculate statistics ----
stats = {
    "Min": df.min(),
    "Max": df.max(),
    "Median": df.median(),
    "Average": df.mean(),
    "Range of Changes": df.max() - df.min()
}

# Put statistics into a DataFrame
stats_df = pd.DataFrame(stats)

# Replace index with actual distances (10, 20, 30, ...)
stats_df.index = range(10, len(stats_df) * 10 + 10, 10)

# Save stats to another Excel file
stats_df.to_excel("MeasuredRanges_Stats.xlsx")

# ---- Function to set X and Y axis ticks for normal charts ----
def set_axis_ticks(ax, y_data):
    ax.set_xticks(stats_df.index)  # X-axis distances
    y_min = int(min(y_data) // 10 * 10)
    y_max = int(max(y_data) // 10 * 10 + 10)
    ax.set_yticks(range(y_min, y_max + 1, 10))

# ---- Plot individual line charts ----
for stat_name in stats:
    fig, ax = plt.subplots()
    ax.plot(
        stats_df.index,
        stats_df[stat_name],
        marker="o",
        linestyle="-",
        label=stat_name
    )
    ax.set_title(f"{stat_name} of Measured Values by Distances")
    ax.set_xlabel("Distance in mm")
    ax.set_ylabel(f"{stat_name} in mm")
    ax.minorticks_on()
    ax.grid(which='major', linestyle='-', linewidth=0.8, alpha=0.7)
    ax.grid(which='minor', linestyle='--', linewidth=0.5, alpha=0.5)
    
    # ---- Special Y-axis for "Range of Changes" ----
    if stat_name == "Range of Changes":
        ax.set_xticks(stats_df.index)
        ax.set_yticks(range(1, 6))  # 1, 2, 3, ..., 5
    else:
        set_axis_ticks(ax, stats_df[stat_name])
    
    ax.legend()
    fig.tight_layout()
    fig.savefig(f"{stat_name}.png")
    plt.close(fig)

# ---- Combined chart for Min, Max, Median, Average ----
fig, ax = plt.subplots()
for stat_name in ["Min", "Max", "Median", "Average"]:
    ax.plot(
        stats_df.index,
        stats_df[stat_name],
        marker="o",
        linestyle="-",
        label=stat_name
    )
ax.set_title("Statistics of Measured Values by Distances")
ax.set_xlabel("Distance in mm")
ax.set_ylabel("Measured Value in mm")
ax.minorticks_on()
ax.grid(which='major', linestyle='-', linewidth=0.8, alpha=0.7)
ax.grid(which='minor', linestyle='--', linewidth=0.5, alpha=0.5)
y_data = stats_df[["Min","Max","Median","Average"]].values.flatten()
set_axis_ticks(ax, y_data)
ax.legend()
fig.tight_layout()
fig.savefig("All_Stats_Combined.png")
plt.close(fig)

print("✅ Charts saved with correct Y-axis for 'Range of Changes'")
