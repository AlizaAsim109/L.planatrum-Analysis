CELL1
!pip install python-docx

CELL2

import re
import docx
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.patches import Patch

# ----------------------------------------------------------------------
# 0. CONFIG
# ----------------------------------------------------------------------
DOCX_PATH = "modules.docx"     # <-- change path if needed
OUTDIR = "."                       # <-- where PNGs are saved

# Curated neurometabolic module set: module_id -> neuroactive axis.
# This is the manual curation step (KEGG has no such grouping) and is the
# single place to edit if the module selection or axis grouping changes.
NEURO_MODULES = {
    # Serotonergic / Kynurenine axis
    "M00022": "Serotonergic/\nKynurenine",   # Shikimate pathway
    "M00023": "Serotonergic/\nKynurenine",   # Tryptophan biosynthesis
    "M00115": "Serotonergic/\nKynurenine",   # NAD biosynthesis (Asp->Quin)
    "M00912": "Serotonergic/\nKynurenine",   # NAD biosynthesis (Trp->Quin)
    "M00038": "Serotonergic/\nKynurenine",   # Kynurenine pathway

    # Dopaminergic / Catecholamine axis
    "M00017": "Dopaminergic/\nCatecholamine",  # Methionine biosynthesis
    "M00035": "Dopaminergic/\nCatecholamine",  # Methionine degradation
    "M00024": "Dopaminergic/\nCatecholamine",  # Phenylalanine biosynthesis
    "M00025": "Dopaminergic/\nCatecholamine",  # Tyrosine biosynthesis (HPP)
    "M00040": "Dopaminergic/\nCatecholamine",  # Tyrosine biosynthesis (Arogenate)
    "M00034": "Dopaminergic/\nCatecholamine",  # Methionine salvage
    "M00044": "Dopaminergic/\nCatecholamine",  # Tyrosine degradation

    # GABAergic / Glutamatergic axis
    "M00027": "GABAergic/\nGlutamatergic",   # GABA shunt
    "M00028": "GABAergic/\nGlutamatergic",   # Ornithine biosynthesis
    "M00015": "GABAergic/\nGlutamatergic",   # Proline biosynthesis
    "M00118": "GABAergic/\nGlutamatergic",   # Glutathione biosynthesis
    "M00131": "GABAergic/\nGlutamatergic",   # Inositol phosphate metabolism

    # Neuroactive cofactors
    "M00125": "Neuroactive\nCofactors",  # Riboflavin/FAD biosynthesis
    "M00140": "Neuroactive\nCofactors",  # C1-unit interconversion
    "M00120": "Neuroactive\nCofactors",  # Coenzyme A biosynthesis
    "M00126": "Neuroactive\nCofactors",  # THF biosynthesis
    "M00842": "Neuroactive\nCofactors",  # BH4 biosynthesis
    "M00843": "Neuroactive\nCofactors",  # L-threo-BH4 biosynthesis
    "M00116": "Neuroactive\nCofactors",  # Menaquinone (K2) biosynthesis
    "M00122": "Neuroactive\nCofactors",  # Cobalamin (B12) biosynthesis
    "M00124": "Neuroactive\nCofactors",  # Pyridoxal-P (B6) biosynthesis

    # Gut-Brain axis
    "M00579": "Gut-Brain\nAxis",  # Acetate production
    "M00100": "Gut-Brain\nAxis",  # Sphingosine degradation
    "M00090": "Gut-Brain\nAxis",  # Phosphatidylcholine biosynthesis
}

# Fixed axis order (top-to-bottom in Fig4, left-to-right in Fig1B)
AXIS_ORDER = [
    "Serotonergic/\nKynurenine",
    "Dopaminergic/\nCatecholamine",
    "GABAergic/\nGlutamatergic",
    "Neuroactive\nCofactors",
    "Gut-Brain\nAxis",
]

STATUS_ORDER = ["Complete", "Near-complete", "Incomplete"]
STATUS_COLORS = {
    "Complete": "#1B6B4E",
    "Near-complete": "#C9971C",
    "Incomplete": "#A3352B",
}

# Short display names for module descriptions in Fig4 (KEGG's own module names
# are long; these are shortened for axis labels/legibility only)
SHORT_NAME_OVERRIDES = {
    "M00022": "Shikimate pathway",
    "M00023": "Tryptophan biosynthesis",
    "M00115": "NAD biosynthesis (Asp\u2192Quin)",
    "M00912": "NAD biosynthesis (Trp\u2192Quin)",
    "M00038": "Kynurenine pathway",
    "M00017": "Methionine biosynthesis",
    "M00035": "Methionine degradation",
    "M00024": "Phenylalanine biosynthesis",
    "M00025": "Tyrosine biosynthesis (HPP)",
    "M00040": "Tyrosine biosynthesis (Aro)",
    "M00034": "Methionine salvage",
    "M00044": "Tyrosine degradation",
    "M00027": "GABA shunt",
    "M00028": "Ornithine biosynthesis",
    "M00015": "Proline biosynthesis",
    "M00118": "Glutathione biosynthesis",
    "M00131": "Inositol phosphate metab.",
    "M00125": "Riboflavin/FAD biosynthesis",
    "M00140": "C1-unit interconversion",
    "M00120": "Coenzyme A biosynthesis",
    "M00126": "THF biosynthesis",
    "M00842": "BH4 biosynthesis",
    "M00843": "L-threo-BH4 biosynthesis",
    "M00116": "Menaquinone (K2) biosynthesis",
    "M00122": "Cobalamin (B12) biosynthesis",
    "M00124": "Pyridoxal-P (B6) biosynthesis",
    "M00579": "Acetate production",
    "M00100": "Sphingosine degradation",
    "M00090": "Phosphatidylcholine biosynth.",
}

# ----------------------------------------------------------------------
# 1. PARSE THE KEGG MODULE DOCX
# ----------------------------------------------------------------------
LINE_PATTERN = re.compile(
    r"^(M\d+)\s*(.+?)\s*\((\d+)\)\s*\((.+?)\s+(\d+)/(\d+)\)$"
)


def parse_kegg_modules(docx_path=DOCX_PATH):
    """Parse every KEGG module bullet line into a tidy DataFrame:
    module_id, description, ko_hits, kegg_status, blocks_present,
    blocks_total, completion_ratio (%).
    """
    d = docx.Document(docx_path)
    rows = []
    for p in d.paragraphs:
        text = p.text.replace("\xa0", " ").strip()
        m = LINE_PATTERN.match(text)
        if m:
            mod_id, desc, ko_hits, kegg_status, x, y = m.groups()
            rows.append({
                "module_id": mod_id,
                "description": desc,
                "ko_hits": int(ko_hits),
                "kegg_status": kegg_status.strip(),
                "blocks_present": int(x),
                "blocks_total": int(y),
                "completion_ratio": 100 * int(x) / int(y),
            })
    df = pd.DataFrame(rows).drop_duplicates("module_id").set_index("module_id")
    return df


def kegg_status_to_reconstruction(kegg_status):
    """KEGG's 4-way completeness label -> 3-tier reconstruction status."""
    if kegg_status == "complete":
        return "Complete"
    if kegg_status == "1 block missing":
        return "Near-complete"
    # "2 blocks missing" and "incomplete" both collapse to Incomplete
    return "Incomplete"


def build_neurometabolic_table(docx_path=DOCX_PATH):
    all_modules = parse_kegg_modules(docx_path)
    sub = all_modules.loc[list(NEURO_MODULES.keys())].copy()
    sub["axis"] = sub.index.map(NEURO_MODULES)
    sub["status"] = sub["kegg_status"].map(kegg_status_to_reconstruction)
    sub["short_name"] = sub.index.map(lambda i: SHORT_NAME_OVERRIDES.get(i, sub.loc[i, "description"]))

    sub["axis"] = pd.Categorical(sub["axis"], categories=AXIS_ORDER, ordered=True)
    sub = sub.sort_values(["axis", "completion_ratio"], ascending=[True, False])
    return sub


# ----------------------------------------------------------------------
# FIGURE 1 — Donut (overall status) + stacked bar (status by axis)
# ----------------------------------------------------------------------
def fig1_status_distribution(df, outpath=f"{OUTDIR}/fig1_status_distribution.png"):
    fig, (axA, axB) = plt.subplots(1, 2, figsize=(15, 8),
                                    gridspec_kw={"width_ratios": [1, 1.3]})

    # --- Panel A: donut ---
    status_counts = df["status"].value_counts().reindex(STATUS_ORDER)
    colors = [STATUS_COLORS[s] for s in STATUS_ORDER]
    wedges, _, autotexts = axA.pie(
        status_counts.values, colors=colors, startangle=90, counterclock=False,
        autopct=lambda p: f"{p:.1f}%", pctdistance=0.78,
        wedgeprops={"width": 0.45, "edgecolor": "white", "linewidth": 2},
        textprops={"color": "white", "fontweight": "bold", "fontsize": 13},
    )
    axA.set_title(f"Reconstruction Status Distribution\nof Neurometabolic Modules (n={len(df)})",
                   fontsize=15, fontweight="bold")
    legend_labels = [f"{s}\n(n={status_counts[s]})" for s in STATUS_ORDER]
    axA.legend(wedges, legend_labels, loc="upper center",
               bbox_to_anchor=(0.5, -0.02), ncol=3, frameon=False, fontsize=11)
    axA.text(-1.35, 1.15, "A", fontsize=20, fontweight="bold")

    # --- Panel B: stacked bar by axis ---
    counts = (df.groupby(["axis", "status"], observed=True).size()
              .unstack(fill_value=0)
              .reindex(index=AXIS_ORDER, columns=STATUS_ORDER, fill_value=0))

    bottom = np.zeros(len(AXIS_ORDER))
    for status in STATUS_ORDER:
        vals = counts[status].values
        bars = axB.bar(AXIS_ORDER, vals, bottom=bottom, color=STATUS_COLORS[status],
                        label=status, edgecolor="white", linewidth=1)
        for bar, v in zip(bars, vals):
            if v > 0:
                axB.text(bar.get_x() + bar.get_width() / 2,
                          bar.get_y() + bar.get_height() / 2,
                          f"{int(v)}", ha="center", va="center",
                          color="white", fontweight="bold", fontsize=12)
        bottom += vals

    axB.set_ylabel("Number of modules", fontsize=12)
    axB.set_title("Module Reconstruction Status\nby Neuroactive Metabolic Axis",
                   fontsize=15, fontweight="bold")
    axB.set_ylim(0, bottom.max() * 1.25)
    axB.legend(loc="upper right", frameon=True, fontsize=11)
    axB.set_facecolor("#F2F0EA")
    axB.yaxis.grid(True, color="white", linewidth=1.2)
    axB.set_axisbelow(True)
    for spine in ["top", "right"]:
        axB.spines[spine].set_visible(False)
    axB.text(-0.9, bottom.max() * 1.2, "B", fontsize=20, fontweight="bold")

    fig.text(0.5, -0.02,
              "Figure 1. (A) Donut chart showing the distribution of "
              f"{len(df)} neurometabolic KEGG modules by reconstruction completeness.\n"
              "(B) Stacked bar chart of module status across five neuroactive metabolic "
              "axes. Numbers within bars indicate module count.",
              ha="center", fontsize=11, style="italic")

    fig.tight_layout()
    fig.savefig(outpath, dpi=200, bbox_inches="tight")
    plt.show()
    return status_counts, counts


# ----------------------------------------------------------------------
# FIGURE 4 — Sorted horizontal bar heatmap of every module's completion ratio
# ----------------------------------------------------------------------
def fig4_heatmap(df, outpath=f"{OUTDIR}/fig4_heatmap.png"):
    df = df.copy()
    n = len(df)
    y_pos = np.arange(n)[::-1]  # top-to-bottom = first row at top

    fig, ax = plt.subplots(figsize=(13, 0.42 * n + 2))

    cmap = plt.get_cmap("RdYlGn")
    ratios = df["completion_ratio"].values

    # background track (full width, light gray) + colored bar (proportional)
    ax.barh(y_pos, [100] * n, color="#E8E4DC", height=0.65, zorder=1)
    bar_colors = [cmap(0.15 + 0.85 * r / 100) for r in ratios]
    ax.barh(y_pos, ratios, color=bar_colors, height=0.65, zorder=2)

    for yp, r in zip(y_pos, ratios):
        ax.text(min(r, 100) + 1.5, yp, f"{r:.0f}%", va="center", fontsize=9, color="black")

    # status dots to the right of the bars
    dot_x = 108
    for yp, status in zip(y_pos, df["status"]):
        ax.scatter(dot_x, yp, color=STATUS_COLORS[status], s=110,
                   edgecolor="black", linewidth=0.5, zorder=3, clip_on=False)

    # y-axis labels: "M00022 - Shikimate pathway"
    labels = [f"{mid} - {row.short_name}" for mid, row in df.iterrows()]
    ax.set_yticks(y_pos)
    ax.set_yticklabels(labels, fontsize=9)

    # axis-band background shading + rotated axis names on the left margin
    axis_palette = {
        "Serotonergic/\nKynurenine":    "#E7EEF8",
        "Dopaminergic/\nCatecholamine": "#FBEAE3",
        "GABAergic/\nGlutamatergic":    "#E9F5EC",
        "Neuroactive\nCofactors":       "#F1EAF7",
        "Gut-Brain\nAxis":              "#FBF3E4",
    }
    axis_text_color = {
        "Serotonergic/\nKynurenine":    "#3A5A8C",
        "Dopaminergic/\nCatecholamine": "#B5651D",
        "GABAergic/\nGlutamatergic":    "#3E8E5B",
        "Neuroactive\nCofactors":       "#7B5EA7",
        "Gut-Brain\nAxis":              "#B08900",
    }
    for axis in AXIS_ORDER:
        idx = np.where(df["axis"].values == axis)[0]
        if len(idx) == 0:
            continue
        y_top = y_pos[idx].max() + 0.5
        y_bot = y_pos[idx].min() - 0.5
        ax.axhspan(y_bot, y_top, xmin=0, xmax=1.0, color=axis_palette[axis],
                   zorder=0, alpha=0.6)
        ax.text(-33, (y_top + y_bot) / 2, axis.replace("\n", " "), rotation=90,
                va="center", ha="center", fontsize=10, fontweight="bold",
                color=axis_text_color[axis])

    ax.set_xlim(0, 112)
    ax.set_ylim(-1, n)
    ax.set_xlabel("Completion ratio", fontsize=12)
    ax.set_xticks([0, 25, 50, 75, 100])
    ax.set_xticklabels(["0%", "25%", "50%", "75%", "100%"])
    for spine in ["top", "right", "left"]:
        ax.spines[spine].set_visible(False)

    legend_handles = [Patch(facecolor=STATUS_COLORS[s], label=s) for s in STATUS_ORDER]
    ax.legend(handles=legend_handles, title="Reconstruction status",
              loc="lower right", fontsize=9, title_fontsize=10, framealpha=0.95)

    ax.set_title(
        "Figure 4. Neurometabolic Module Reconstruction Heatmap\n"
        f"All {n} modules sorted by neuroactive axis and completion ratio. "
        "Coloured dots: reconstruction status.",
        fontsize=13, fontweight="bold")

    fig.tight_layout()
    fig.savefig(outpath, dpi=200, bbox_inches="tight")
    plt.show()


# ----------------------------------------------------------------------
# RUN EVERYTHING
# ----------------------------------------------------------------------
if __name__ == "__main__":
    neuro_df = build_neurometabolic_table(DOCX_PATH)
    print(neuro_df[["description", "kegg_status", "completion_ratio", "status", "axis"]])

    fig1_status_distribution(neuro_df)
    fig4_heatmap(neuro_df)

    print("\nBoth figures saved to:", OUTDIR)