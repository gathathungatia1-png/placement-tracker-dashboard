"""
Placement tracker - field monitoring dashboard
==============================================

Reads the placement tracker workbook (one sheet per product type) and shows
WHAT has been placed in the field, WHEN it was placed and WHERE it was placed.

Run from the VS Code terminal:

    py -m streamlit run stove_placement_dashboard.py --server.address localhost

Requires: pandas, seaborn, streamlit, openpyxl
    py -m pip install pandas seaborn streamlit openpyxl

Where the data comes from
-------------------------
* On your laptop: the tracker workbook. Its full path is kept in a small text
  file next to this script, "tracker_path.txt", so that the path is not part
  of the code (the code is public on GitHub). Edit that file when the
  workbook moves or is renamed.
* Online, or when the workbook cannot be found: "placements_public.csv" next
  to this script - the cleaned records with personal details removed. Create
  it with the "Download public data file" button in the sidebar of the local
  dashboard, then upload it to GitHub to update the online dashboard.
"""

from __future__ import annotations

import math
import numbers
import re
from pathlib import Path

import matplotlib.dates as mdates
import matplotlib.pyplot as plt
import pandas as pd
import seaborn as sns
import streamlit as st
from matplotlib import font_manager
from matplotlib.colors import LinearSegmentedColormap
from matplotlib.ticker import MaxNLocator

# --------------------------------------------------------------------------
# SETTINGS - edit these to suit the workbook
# --------------------------------------------------------------------------

APP_DIR = Path(__file__).resolve().parent
TRACKER_PATH_FILE = APP_DIR / "tracker_path.txt"        # laptop only, not on GitHub
PUBLIC_DATA_FILE = APP_DIR / "placements_public.csv"    # no personal details

# The public data file never contains phone numbers or contact persons, and
# participant names are removed too - except the names of institutions
# (schools, children's homes...) from these product sheets.
PUBLIC_NAME_PRODUCTS = ("Institutional",)
# ...unless the "institution" is entered as a person ("Madam ...", "Mr ...").
PERSONAL_TITLE = re.compile(r"\s*(madam|mr|mrs|ms|miss|dr|mama|baba|pastor)\b", re.I)

# Sheets whose name contains any of these words are summaries, not product sheets.
# Every other sheet with a "No. / Location / Status" header row is read as a product.
NON_PRODUCT_SHEETS = ("dashboard", "all placements", "qc report", "source log")

# The product sheets do not all use the same column headings. Each dashboard
# column is filled from the first of these headings that the sheet has.
COLUMN_ALIASES = {
    "Model": ("Stove Type", "Stove Model"),
    "Description": ("Stove Description", "Body Material"),
    "Stove ID": ("Stove ID",),
    "Sensor ID": ("Sensor ID",),
    "Participant Name": ("Participant Name", "Institution"),
    "Phone": ("Phone",),
    "Contact Person": ("Contact Person",),
    "Location": ("Location",),
    "Status": ("Status",),
    "Customer Type": ("Customer Type", "Institution Type", "Segment"),
    "Study Type": ("Study Type", "Study"),
    "Placement Date": ("Placement Date (Stove)", "Placement Date", "Placed"),
}

NOT_RECORDED = "Not recorded"
OTHER_AREA = "Other / unmapped"

# The Location column is free text ("Hilton-Ruiru", "Tumba market- Machakos"...).
# To answer "where", each location is grouped into an area using these keywords.
# A keyword matches the start of a word, and the FIRST rule that matches wins,
# so the order matters. Add or edit freely.
AREA_RULES: list[tuple[str, tuple[str, ...]]] = [
    ("Machakos", ("machakos", "syokimau", "mlologo", "mlolongo", "manza")),
    ("Nairobi", ("nairobi", "utawala", "mathare", "mukuru", "karen", "kasarani",
                 "zimmerman", "njiru", "eastleigh")),
    ("Eastern Bypass", ("bypass",)),
    ("Ruiru", ("ruiru", "wataalam", "rainbow", "kimbo", "hilton", "gitambaya",
               "kitambaya", "gitaa mbaya", "gwa kairu", "matopeni", "masaku",
               "kihunguro", "toll", "mugutha", "kamakis", "gitothua", "matangi",
               "bpl", "burn manufacturing")),
    ("Githurai", ("githurai", "githu 45", "mwihoko")),
    ("Kahawa", ("kahawa", "ku model")),
    ("Juja", ("juja", "gachororo")),
    ("Thika", ("thika",)),
    ("Gatundu", ("gatundu",)),
    ("Githunguri", ("githunguri", "komothai", "ikinu", "ikiinu")),
    ("Kikuyu", ("kikuyu", "gikambura", "gitaru", "karai", "kamangu")),
    ("Kajiado", ("kajiado", "ngong", "rongai", "kitengela")),
    ("Murang'a", ("muranga", "murang'a", "gatanga", "kenol", "sabasaba")),
    ("Mombasa", ("mombasa", "bombolulu", "miritini")),
    ("Kakamega", ("kakamega", "mumias")),
    ("Kitui", ("kitui",)),
    ("Nakuru", ("nakuru", "njoro")),
    ("Makueni", ("makueni",)),
    ("Kericho", ("kericho",)),
    ("Siaya", ("siaya",)),
    ("Baringo", ("baringo",)),
    ("Embu", ("embu",)),
    ("Wajir", ("wajir",)),
    ("Bomet", ("bomet",)),
    ("Kiambu (other)", ("kiambu", "kiamabu", "kimabu")),
]

# Text found in the Location column that is not a place (visit notes etc.).
NOT_A_LOCATION = ("not recorded", "any day", "wedne", "tueday", "tuesday",
                  "unavailable", "unreachable", "stolen")

# The Status column is also free text. Anything not listed here is shown as
# "Unknown / other" (e.g. "Unknown", "Saturday", "not reachable").
STATUS_RULES = {
    "In the field": ("active", "running", "placed"),
    "Out of service": ("failed", "inactive", "no not in service", "picked for repair"),
    "Not placed": ("not placed", "to be placed"),
}
STATUS_OTHER = "Unknown / other"

# --------------------------------------------------------------------------
# COLOURS - one fixed colour per product, in light and dark versions
# --------------------------------------------------------------------------

SERIES_LIGHT = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100",
                "#e87ba4", "#008300", "#6250d6", "#e34948"]
SERIES_DARK = ["#3987e5", "#d95926", "#199e70", "#c98500",
               "#d55181", "#008300", "#9085e9", "#e66767"]

THEMES = {
    "light": dict(surface="#fcfcfb", ink="#0b0b0b", ink2="#52514e", muted="#898781",
                  grid="#e1e0d9", axis="#c3c2b7", empty="#f3f2ee", series=SERIES_LIGHT,
                  ramp=["#cde2fb", "#6da7ec", "#2a78d6", "#0d366b"]),
    "dark": dict(surface="#1a1a19", ink="#f0efec", ink2="#c3c2b7", muted="#898781",
                 grid="#2c2c2a", axis="#383835", empty="#232321", series=SERIES_DARK,
                 ramp=["#0d366b", "#1c5cab", "#3987e5", "#9ec5f4"]),
}


def current_theme() -> dict:
    """Match the charts to Streamlit's light / dark mode."""
    try:
        kind = st.context.theme.type
    except Exception:
        kind = None
    return THEMES["dark" if kind == "dark" else "light"]


def product_colours(products: list[str], theme: dict) -> dict[str, str]:
    """Colour follows the product (sheet order), never its rank in a chart."""
    series = theme["series"]
    return {p: (series[i] if i < len(series) else theme["muted"])
            for i, p in enumerate(products)}


# --------------------------------------------------------------------------
# LOADING AND CLEANING
# --------------------------------------------------------------------------

def _to_text(value) -> str | None:
    """Excel cell -> tidy text (10766.0 -> '10766', blanks -> None)."""
    if value is None or value is pd.NaT:
        return None
    if isinstance(value, float) and math.isnan(value):
        return None
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    text = re.sub(r"\s+", " ", str(value)).strip()
    return text or None


def _to_date(value) -> pd.Timestamp:
    """Dates arrive as real dates, as text, or as Excel day numbers (45728)."""
    if isinstance(value, numbers.Real) and not isinstance(value, bool):
        if math.isnan(value) or not 20000 < value < 80000:
            return pd.NaT
        return pd.Timestamp("1899-12-30") + pd.Timedelta(days=float(value))
    return pd.to_datetime(value, errors="coerce")


def _find_header_row(preview: pd.DataFrame) -> int | None:
    """A product sheet has a title row, then a header row with these columns."""
    for i, row in preview.iterrows():
        cells = {str(v).strip().lower() for v in row.dropna()}
        if {"no.", "location", "status"} <= cells:
            return int(i)
    return None


def _cookware_model(row: pd.Series) -> str | None:
    """Cookware has no model column: build one from brand, item and size."""
    item = (_to_text(row.get("Item")) or "").strip(" ,")
    item = re.sub(r"\s*,[\s,]*", ", ", item).replace(", &", " &")
    if re.fullmatch(r"(to be placed|\(not recorded\)|\d+hh|wants .*)?", item.lower()):
        item = ""
    brand = _to_text(row.get("Brand")) or ""
    if brand.lower() in item.lower():
        brand = ""
    return " ".join(p for p in (brand, item, _to_text(row.get("Size"))) if p) or None


def _read_product_sheet(workbook: pd.ExcelFile, sheet: str) -> pd.DataFrame | None:
    header_row = _find_header_row(workbook.parse(sheet, header=None, nrows=10))
    if header_row is None:
        return None
    raw = workbook.parse(sheet, header=header_row)
    raw.columns = [re.sub(r"\s+", " ", str(c)).strip() for c in raw.columns]
    raw = raw.dropna(how="all", subset=[c for c in raw.columns if c != "No."])

    tidy = pd.DataFrame(index=raw.index)
    # Sheet names look like "🔥 Charcoal" - keep just the word.
    tidy["Product"] = re.sub(r"^[^\w]+", "", sheet).strip()
    for column, headings in COLUMN_ALIASES.items():
        source = next((h for h in headings if h in raw.columns), None)
        tidy[column] = raw[source] if source else None
    if "Item" in raw.columns:
        tidy["Model"] = raw.apply(_cookware_model, axis=1)
    return tidy


def _area_from_location(location: str | None) -> str:
    if not location:
        return NOT_RECORDED
    text = location.lower()
    if any(word in text for word in NOT_A_LOCATION):
        return NOT_RECORDED
    for area, keywords in AREA_RULES:
        if any(re.search(rf"\b{re.escape(word)}", text) for word in keywords):
            return area
    return OTHER_AREA


def _status_group(status: str | None) -> str:
    text = (status or "").lower()
    for group, values in STATUS_RULES.items():
        if text in values:
            return group
    return STATUS_OTHER


def _tidy_model_names(models: pd.Series) -> pd.Series:
    """Merge spellings that differ only by case, spaces or punctuation
    ('Ecoa char' / 'Ecoa Char', 'Pro200' / 'Pro 200', 'MM J' / 'MMJ', 'R-120' / 'R120').
    The most common spelling is the one shown."""
    key = models.str.lower().str.replace(r"[^a-z0-9]", "", regex=True)
    key = key.mask(key == "")
    most_common = models.groupby(key).agg(lambda s: s.value_counts().idxmax())
    return key.map(most_common).fillna(NOT_RECORDED)


def _one_row_per_stove(data: pd.DataFrame) -> pd.DataFrame:
    """Count each stove once: rows that repeat a stove ID are merged.

    The tracker was consolidated from several source files, so the same stove
    can be listed two or three times. For each stove ID (ignoring spaces,
    hyphens and case) the first row is kept and any blank cells are filled
    from the later rows. A status the dashboard recognises ("Active",
    "Failed"...) is preferred over "Unknown" or a visit note.

    IDs are matched within a product sheet only: the Charcoal and Wood sheets
    reuse a few numbers for different stoves in different homes.
    """
    data = data.assign(
        _row=range(len(data)),
        _key=data["Stove ID"].map(
            lambda s: re.sub(r"[^A-Z0-9]", "", s.upper()) or None if s else None),
        _known_status=data["Status"].where(
            data["Status"].map(_status_group) != STATUS_OTHER),
    )
    with_id = data[data["_key"].notna()]
    groups = with_id.groupby(["Product", "_key"], sort=False)
    merged = groups.first()                       # first non-blank value per column
    merged["Tracker rows"] = groups.size()
    merged = merged.reset_index()

    without_id = data[data["_key"].isna()].assign(**{"Tracker rows": 1})
    result = pd.concat([merged, without_id]).sort_values("_row", ignore_index=True)
    result["Status"] = result["_known_status"].fillna(result["Status"])
    return result.drop(columns=["_row", "_key", "_known_status"])


RECORD_COLUMNS = ["Product", *COLUMN_ALIASES, "Tracker rows"]
# Left out of the public data file. "Description" holds material specifications
# and supplier names; no chart uses it, so it stays on the laptop as well.
PRIVATE_COLUMNS = ["Phone", "Contact Person", "Description"]


def _read_workbook(path: str) -> pd.DataFrame:
    """Read every product sheet into one table, one row per product placed."""
    frames = []
    with pd.ExcelFile(path) as workbook:
        for sheet in workbook.sheet_names:
            if any(word in sheet.lower() for word in NON_PRODUCT_SHEETS):
                continue
            frame = _read_product_sheet(workbook, sheet)
            if frame is not None:
                frames.append(frame)
    if not frames:
        return pd.DataFrame()

    data = pd.concat(frames, ignore_index=True)
    for column in COLUMN_ALIASES:
        if column != "Placement Date":
            data[column] = data[column].map(_to_text)

    # Phone numbers typed as numbers lose their leading zero in Excel.
    data["Phone"] = data["Phone"].map(
        lambda p: f"0{p}" if p and re.fullmatch(r"[17]\d{8}", p) else p)

    # 1970-01-01 is an empty-cell placeholder, not a real placement date.
    dates = pd.to_datetime(data["Placement Date"].map(_to_date), errors="coerce")
    dates = dates.dt.normalize()
    data["Placement Date"] = dates.where(dates >= pd.Timestamp("2000-01-01"))

    return _one_row_per_stove(data)[RECORD_COLUMNS]


def public_copy(data: pd.DataFrame) -> pd.DataFrame:
    """The records without personal details, for the online dashboard."""
    public = data[[c for c in RECORD_COLUMNS if c not in PRIVATE_COLUMNS]].copy()
    is_institution = (
        public["Product"].isin(PUBLIC_NAME_PRODUCTS)
        & ~public["Participant Name"].fillna("").str.match(PERSONAL_TITLE))
    public["Participant Name"] = public["Participant Name"].where(is_institution)
    return public


def public_file_bytes(data: pd.DataFrame, source_name: str, saved: pd.Timestamp) -> bytes:
    """The public data file: one note line, then the records as CSV."""
    note = (f"# Public copy of {source_name} (saved {saved:%d %b %Y}); "
            "participant names, contact details and product descriptions removed\n")
    return (note + public_copy(data).to_csv(index=False, date_format="%Y-%m-%d",
                                            lineterminator="\n")).encode("utf-8")


def public_file_note(path: Path) -> str:
    with open(path, encoding="utf-8-sig") as file:
        return file.readline().lstrip("# ").strip()


def _read_public_file(path: str) -> pd.DataFrame:
    data = pd.read_csv(path, skiprows=1, dtype=str, encoding="utf-8-sig")
    for column in RECORD_COLUMNS:
        if column not in data.columns:
            data[column] = None                   # the private columns
    for column in COLUMN_ALIASES:
        if column != "Placement Date":
            data[column] = data[column].map(_to_text)
    data["Placement Date"] = pd.to_datetime(data["Placement Date"], errors="coerce")
    data["Tracker rows"] = pd.to_numeric(data["Tracker rows"]).fillna(1).astype(int)
    return data[RECORD_COLUMNS]


@st.cache_data(show_spinner="Reading the placement tracker...")
def load_placements(path: str, modified: float) -> pd.DataFrame:
    """The records plus the columns the dashboard works out for itself.

    `modified` (the file's last-modified time) is only there so the cache
    refreshes automatically whenever the file is saved.
    """
    is_public_file = path.lower().endswith(".csv")
    data = _read_public_file(path) if is_public_file else _read_workbook(path)
    if data.empty:
        return data

    data["Model"] = _tidy_model_names(data["Model"])
    data["Area"] = data["Location"].map(_area_from_location)
    data["Status Group"] = data["Status"].map(_status_group)
    age_days = (pd.Timestamp.today().normalize() - data["Placement Date"]).dt.days
    data["Years in field"] = (age_days / 365.25).round(1)

    columns = ["Product", "Model", "Description", "Stove ID", "Sensor ID",
               "Participant Name", "Phone", "Contact Person", "Location", "Area",
               "Status", "Status Group", "Customer Type", "Study Type",
               "Placement Date", "Years in field", "Tracker rows"]
    return data[columns]


# --------------------------------------------------------------------------
# CHART HELPERS
# --------------------------------------------------------------------------

# Segoe UI on Windows; the online server (Linux) does not have it.
CHART_FONT = ("Segoe UI" if any(f.name == "Segoe UI" for f in font_manager.fontManager.ttflist)
              else "DejaVu Sans")


def apply_chart_style(theme: dict) -> None:
    sns.set_theme(style="white", rc={
        "figure.facecolor": theme["surface"], "axes.facecolor": theme["surface"],
        "savefig.facecolor": theme["surface"],
        "text.color": theme["ink"], "axes.labelcolor": theme["ink2"],
        "xtick.color": theme["muted"], "ytick.color": theme["muted"],
        "xtick.labelcolor": theme["ink2"], "ytick.labelcolor": theme["ink2"],
        "axes.edgecolor": theme["axis"], "grid.color": theme["grid"],
        "grid.linewidth": 0.8, "axes.linewidth": 0.8,
        "font.family": CHART_FONT, "font.size": 10,
        "axes.labelsize": 10, "xtick.labelsize": 9.5, "ytick.labelsize": 9.5,
        "legend.fontsize": 9.5, "figure.dpi": 150,
        "patch.edgecolor": "none", "patch.linewidth": 0,
    })


def new_figure(width: float, height: float, grid: str | None = None):
    fig, ax = plt.subplots(figsize=(width, height))
    sns.despine(ax=ax, left=(grid == "x"), bottom=(grid == "y"))
    if grid:
        ax.grid(axis=grid)
        ax.set_axisbelow(True)
    return fig, ax


def show(fig, alt: str) -> None:
    fig.tight_layout()
    st.pyplot(fig, alt=alt)
    plt.close(fig)


def legend_on_top(ax, theme: dict) -> None:
    """Product names above the plot, in one row (two when there are many)."""
    legend = ax.get_legend()
    if legend is None:
        return
    entries = len(legend.get_texts())
    sns.move_legend(ax, "lower left", bbox_to_anchor=(0, 1.0), title=None,
                    ncol=entries if entries <= 4 else math.ceil(entries / 2),
                    frameon=False, handlelength=1.1, columnspacing=1.4,
                    labelcolor=theme["ink2"], borderaxespad=0.2)


def shorten(text: str, limit: int = 36) -> str:
    return text if len(text) <= limit else text[: limit - 1].rstrip() + "…"


# --------------------------------------------------------------------------
# CHARTS
# --------------------------------------------------------------------------

def chart_by_product(df: pd.DataFrame, colours: dict, theme: dict) -> None:
    counts = (df["Product"].value_counts().rename_axis("Product")
              .reset_index(name="Placed"))
    fig, ax = new_figure(6.2, 0.5 * len(counts) + 1.0, grid="x")
    sns.barplot(data=counts, y="Product", x="Placed", hue="Product", palette=colours,
                legend=False, width=0.62, saturation=1, linewidth=0, ax=ax)
    for bars in ax.containers:
        ax.bar_label(bars, fmt="{:,.0f}", padding=4, color=theme["ink"])
    ax.set(xlabel="Products placed", ylabel="")
    ax.margins(x=0.1)
    ax.xaxis.set_major_locator(MaxNLocator(integer=True, nbins=6))
    show(fig, "Bar chart of products placed, by product type")


def chart_top_models(df: pd.DataFrame, colours: dict, theme: dict, top_n: int) -> None:
    counts = (df.groupby(["Product", "Model"], observed=True).size()
              .reset_index(name="Placed")
              .sort_values("Placed", ascending=False).head(top_n))
    # The same model name can appear under two products - keep the bars separate.
    shared = counts["Model"].duplicated(keep=False)
    counts["Label"] = counts["Model"].map(shorten)
    counts.loc[shared, "Label"] = counts["Label"] + " (" + counts["Product"] + ")"
    present = [p for p in colours if p in set(counts["Product"])]

    fig, ax = new_figure(6.2, 0.3 * len(counts) + 1.3, grid="x")
    sns.barplot(data=counts, y="Label", x="Placed", hue="Product", hue_order=present,
                palette=colours, dodge=False, width=0.66, saturation=1, linewidth=0,
                legend=len(present) > 1, ax=ax)
    for bars in ax.containers:
        ax.bar_label(bars, fmt="{:,.0f}", padding=4, color=theme["ink"])
    ax.set(xlabel="Products placed", ylabel="")
    ax.margins(x=0.1)
    ax.xaxis.set_major_locator(MaxNLocator(integer=True, nbins=6))
    legend_on_top(ax, theme)
    show(fig, "Bar chart of the most-placed models")


def chart_placements_over_time(dated: pd.DataFrame, colours: dict, theme: dict,
                               grain: str) -> None:
    freq = {"Month": "M", "Quarter": "Q", "Year": "Y"}[grain]
    periods = dated["Placement Date"].dt.to_period(freq)
    # Every period between the first and last placement gets a slot, so quiet
    # spells show up as gaps instead of being squeezed out of the axis.
    timeline = pd.period_range(periods.min(), periods.max(), freq=freq)
    position = {p: i for i, p in enumerate(timeline)}
    plot_data = dated.assign(Slot=periods.map(position).astype(int))
    present = [p for p in colours if p in set(plot_data["Product"])]

    fig, ax = new_figure(6.6, 3.5, grid="y")
    sns.histplot(data=plot_data, x="Slot", hue="Product", hue_order=present,
                 palette=colours, multiple="stack", discrete=True, shrink=0.82,
                 alpha=1, edgecolor=theme["surface"], linewidth=0.8,
                 legend=len(present) > 1, ax=ax)

    # Label a readable subset of periods, aligned to the calendar.
    n = len(timeline)
    if grain == "Month":
        step = next(s for s in (1, 2, 3, 6, 12) if n / s <= 12 or s == 12)
        ticks = [i for i, p in enumerate(timeline) if (p.month - 1) % step == 0]
        labels = [timeline[i].strftime("%b %Y") for i in ticks]
    elif grain == "Quarter":
        step = next(s for s in (1, 2, 4) if n / s <= 12 or s == 4)
        ticks = [i for i, p in enumerate(timeline) if (p.quarter - 1) % step == 0]
        labels = [f"Q{timeline[i].quarter} {timeline[i].year}" for i in ticks]
    else:
        ticks = list(range(n))
        labels = [str(p.year) for p in timeline]
    ax.set_xticks(ticks, labels, rotation=35 if len(ticks) > 6 else 0,
                  ha="right" if len(ticks) > 6 else "center")
    ax.set_xlim(-0.7, n - 0.3)
    ax.set(xlabel="", ylabel=f"Placed per {grain.lower()}")
    ax.yaxis.set_major_locator(MaxNLocator(integer=True, nbins=6))
    legend_on_top(ax, theme)
    show(fig, f"Stacked bar chart of products placed per {grain.lower()}")


def chart_cumulative(dated: pd.DataFrame, colours: dict, theme: dict) -> None:
    daily = (dated.groupby(["Product", "Placement Date"], observed=True).size()
             .reset_index(name="Placed"))
    daily["Total placed"] = daily.groupby("Product", observed=True)["Placed"].cumsum()
    # Carry each line through to the latest placement so they all end together.
    last_day = dated["Placement Date"].max()
    tails = daily.groupby("Product", observed=True).tail(1).assign(
        **{"Placement Date": last_day})
    running = (pd.concat([daily, tails])
               .drop_duplicates(["Product", "Placement Date"], keep="first"))
    present = [p for p in colours if p in set(running["Product"])]

    fig, ax = new_figure(6.6, 3.5, grid="y")
    sns.lineplot(data=running, x="Placement Date", y="Total placed", hue="Product",
                 hue_order=present, palette=colours, drawstyle="steps-post",
                 linewidth=2, estimator=None, legend=len(present) > 1, ax=ax)

    # Label where each line ends, nudging labels apart if they would overlap.
    ends = tails.sort_values("Total placed")
    top = ends["Total placed"].max()
    min_gap, last_y = top * 0.07, -math.inf
    for _, row in ends.iterrows():
        y = max(row["Total placed"], last_y + min_gap)
        ax.annotate(f"{row['Total placed']:,}", (last_day, y), xytext=(6, 0),
                    textcoords="offset points", va="center", color=theme["ink"])
        last_y = y

    first_day = dated["Placement Date"].min()
    span = max((last_day - first_day).days, 30)
    ax.set_xlim(first_day - pd.Timedelta(days=span * 0.02),
                last_day + pd.Timedelta(days=span * 0.09))
    ax.set_ylim(bottom=0)
    locator = mdates.AutoDateLocator(minticks=3, maxticks=7)
    ax.xaxis.set_major_locator(locator)
    ax.xaxis.set_major_formatter(mdates.ConciseDateFormatter(locator))
    ax.set(xlabel="", ylabel="Placed to date")
    ax.yaxis.set_major_locator(MaxNLocator(integer=True, nbins=6))
    legend_on_top(ax, theme)
    show(fig, "Line chart of the running total of products placed, by product type")


def chart_by_place(placed: pd.DataFrame, place_col: str, places: list[str],
                   colours: dict, theme: dict) -> None:
    plot_data = placed[placed[place_col].isin(places)].copy()
    plot_data[place_col] = pd.Categorical(plot_data[place_col], categories=places,
                                          ordered=True)
    present = [p for p in colours if p in set(plot_data["Product"])]

    fig, ax = new_figure(6.4, 0.31 * len(places) + 1.3, grid="x")
    sns.histplot(data=plot_data, y=place_col, hue="Product", hue_order=present,
                 palette=colours, multiple="stack", discrete=True, shrink=0.7,
                 alpha=1, edgecolor=theme["surface"], linewidth=0.8,
                 legend=len(present) > 1, ax=ax)
    totals = plot_data[place_col].value_counts()
    for i, place in enumerate(places):
        ax.annotate(f"{totals[place]:,}", (totals[place], i), xytext=(4, 0),
                    textcoords="offset points", va="center", color=theme["ink"])
    ax.set_yticks(range(len(places)), [shorten(p) for p in places])
    ax.set_ylim(len(places) - 0.4, -0.6)          # biggest at the top
    ax.set(xlabel="Products placed", ylabel="")
    ax.margins(x=0.1)
    ax.xaxis.set_major_locator(MaxNLocator(integer=True, nbins=6))
    legend_on_top(ax, theme)
    show(fig, f"Stacked bar chart of products placed by {place_col.lower()}")


def chart_place_by_year(placed: pd.DataFrame, place_col: str, places: list[str],
                        theme: dict) -> None:
    plot_data = placed[placed[place_col].isin(places)]
    year = plot_data["Placement Date"].dt.year
    column = year.astype("Int64").astype("string").fillna("No date")
    grid = pd.crosstab(plot_data[place_col], column).reindex(places).fillna(0).astype(int)
    years = sorted(c for c in grid.columns if c != "No date")
    grid = grid[years + (["No date"] if "No date" in grid.columns else [])]
    grid.index = [shorten(p) for p in grid.index]

    fig, ax = plt.subplots(figsize=(6.4, 0.31 * len(places) + 1.3))
    ax.set_facecolor(theme["empty"])
    sns.heatmap(grid, mask=grid == 0, annot=True, fmt="d", vmin=0,
                cmap=LinearSegmentedColormap.from_list("placements", theme["ramp"]),
                linewidths=2, linecolor=theme["surface"], cbar=False, ax=ax)
    ax.set(xlabel="Year placed", ylabel="")
    ax.tick_params(length=0)
    plt.setp(ax.get_yticklabels(), rotation=0)
    show(fig, f"Grid of products placed by {place_col.lower()} and year")


# --------------------------------------------------------------------------
# PAGE
# --------------------------------------------------------------------------

st.set_page_config(page_title="Placement tracker", page_icon=":material/location_on:",
                   layout="wide")

# The workbook on this laptop if there is one, otherwise the public data file.
workbook_path = None
if TRACKER_PATH_FILE.exists():
    workbook_path = Path(TRACKER_PATH_FILE.read_text(encoding="utf-8-sig").strip().strip('"'))
if workbook_path is not None and workbook_path.exists():
    data_path = workbook_path
elif PUBLIC_DATA_FILE.exists():
    data_path = PUBLIC_DATA_FILE
    if workbook_path is not None:
        st.warning(f"`{workbook_path}`\n\nShowing the public data file instead. "
                   f"Correct the path in {TRACKER_PATH_FILE.name} to use the workbook.",
                   title="Could not find the tracker workbook")
else:
    st.error(f"Put the full path of the tracker workbook in {TRACKER_PATH_FILE.name}, "
             f"or add {PUBLIC_DATA_FILE.name}, next to this script.",
             title="No data to show")
    st.stop()
is_public = data_path == PUBLIC_DATA_FILE

try:
    placements = load_placements(str(data_path), data_path.stat().st_mtime)
except PermissionError:
    st.error("Close it in Excel (or wait for OneDrive to finish syncing) and "
             "reload this page.", title="The workbook is locked by another program")
    st.stop()

if placements.empty:
    st.error("A product sheet needs a header row with 'No.', 'Location' and "
             "'Status' columns.", title="No product sheets were found")
    st.stop()

theme = current_theme()
apply_chart_style(theme)
all_products = list(dict.fromkeys(placements["Product"]))     # sheet order
colours = product_colours(all_products, theme)

# ---- Sidebar filters ------------------------------------------------------

with st.sidebar:
    st.header("Filters")

    chosen_products = st.multiselect("Product", all_products, default=all_products)
    in_products = placements[placements["Product"].isin(chosen_products)]

    chosen_models = st.multiselect(
        "Model", sorted(in_products["Model"].unique()), placeholder="All models")
    chosen_areas = st.multiselect(
        "Area", sorted(placements["Area"].unique()), placeholder="All areas")

    status_groups = [g for g in [*STATUS_RULES, STATUS_OTHER]
                     if g in set(placements["Status Group"])]
    chosen_status = st.multiselect(
        "Status", status_groups,
        default=[g for g in status_groups if g != "Not placed"],
        help="Records marked 'Not Placed' or 'To be placed' in the tracker are "
             "left out by default.")

    first_date = placements["Placement Date"].min().date()
    last_date = placements["Placement Date"].max().date()
    chosen_dates = st.date_input("Placed between", value=(first_date, last_date),
                                 min_value=first_date, max_value=last_date,
                                 format="DD/MM/YYYY")
    include_undated = st.checkbox("Include products with no placement date", value=True)

    search = st.text_input(
        "Search", placeholder=("Institution, stove ID or sensor ID" if is_public
                               else "Participant, stove ID or sensor ID"))

    st.divider()
    modified = pd.Timestamp(data_path.stat().st_mtime, unit="s", tz="UTC").tz_convert(
        "Africa/Nairobi")
    tracker_rows = int(placements["Tracker rows"].sum())
    merged_away = tracker_rows - len(placements)
    source_note = (public_file_note(data_path) if is_public else
                   f"Source: {data_path.name}  \nLast saved: {modified:%d %b %Y, %H:%M}")
    st.caption(
        f"{source_note}  \n"
        f"{tracker_rows:,} tracker rows, {len(placements):,} products"
        + (f" ({merged_away:,} rows repeated a stove ID and were merged)"
           if merged_away else ""))
    if not is_public:
        st.download_button(
            "Download public data file",
            public_file_bytes(placements, data_path.name, modified),
            file_name=PUBLIC_DATA_FILE.name, mime="text/csv",
            icon=":material/download:",
            help="These records without participant names, contact details or "
                 "product descriptions. Upload it to GitHub to update the online "
                 "dashboard.")
        if st.button("Reload workbook", icon=":material/refresh:"):
            st.cache_data.clear()
            st.rerun()

# ---- Apply filters --------------------------------------------------------

view = in_products
if chosen_models:
    view = view[view["Model"].isin(chosen_models)]
if chosen_areas:
    view = view[view["Area"].isin(chosen_areas)]
view = view[view["Status Group"].isin(chosen_status)]

if isinstance(chosen_dates, tuple) and len(chosen_dates) == 2:
    start, end = (pd.Timestamp(d) for d in chosen_dates)
    in_range = view["Placement Date"].between(start, end)
    undated = view["Placement Date"].isna()
    view = view[in_range | (undated & include_undated)]
elif not include_undated:
    view = view[view["Placement Date"].notna()]

if search.strip():
    needle = search.strip().lower()
    haystack = (view[["Participant Name", "Stove ID", "Sensor ID"]]
                .fillna("").agg(" ".join, axis=1).str.lower())
    view = view[haystack.str.contains(needle, regex=False)]

# ---- Header and headline numbers ------------------------------------------

st.title("Placement tracker")
st.caption("Products placed in the field for monitoring - what, when and where.")

if view.empty:
    st.info("No placements match the current filters.")
    st.stop()

dated = view[view["Placement Date"].notna()]
placed = view[view["Area"] != NOT_RECORDED]          # has a usable location

k1, k2, k3, k4, k5 = st.columns([1, 1, 1, 1, 1.6])
k1.metric("Products placed", f"{len(view):,}",
          help="Each stove ID is counted once per product sheet.")
k2.metric("In the field", f"{(view['Status Group'] == 'In the field').sum():,}",
          help="Status is 'Active', 'Running' or 'Placed' in the tracker.")
k3.metric("Models", view.loc[view["Model"] != NOT_RECORDED, "Model"].nunique())
k4.metric("Areas", placed["Area"].nunique())
k5.metric("Latest placement",
          f"{dated['Placement Date'].max():%d %b %Y}" if len(dated) else "-")

# ---- What -----------------------------------------------------------------

st.header("What has been placed")
left, right = st.columns(2)
with left:
    st.subheader("By product type")
    chart_by_product(view, colours, theme)
with right:
    st.subheader("Most-placed models")
    chart_top_models(view[view["Model"] != NOT_RECORDED], colours, theme, top_n=12)

# ---- When -----------------------------------------------------------------

st.header("When they were placed")
if dated.empty:
    st.info("None of the selected products has a placement date.")
else:
    missing = len(view) - len(dated)
    grain = st.segmented_control("Group by", ["Month", "Quarter", "Year"],
                                 default="Month", required=True)
    left, right = st.columns(2)
    with left:
        st.subheader(f"Placements per {grain.lower()}")
        chart_placements_over_time(dated, colours, theme, grain)
    with right:
        st.subheader("Running total")
        chart_cumulative(dated, colours, theme)
    st.caption(
        f"First placement {dated['Placement Date'].min():%d %b %Y}, "
        f"latest {dated['Placement Date'].max():%d %b %Y}. "
        + (f"{missing:,} of {len(view):,} selected products have no placement date "
           "in the tracker and are not shown in these two charts." if missing else ""))

# ---- Where ----------------------------------------------------------------

st.header("Where they were placed")
if placed.empty:
    st.info("None of the selected products has a usable location.")
else:
    c1, c2 = st.columns([2, 3], vertical_alignment="bottom")
    level = c1.segmented_control(
        "Show", ["Area", "Exact location"], default="Area", required=True,
        help="Areas group the free-text locations from the tracker using the "
             "AREA_RULES list at the top of the script.")
    place_col = "Area" if level == "Area" else "Location"
    n_places = placed[place_col].nunique()
    top_n = (c2.slider("Number of places to show", 5, min(n_places, 30),
                       min(n_places, 12)) if n_places > 5 else n_places)
    places = placed[place_col].value_counts().head(top_n).index.tolist()

    left, right = st.columns(2)
    with left:
        st.subheader(f"By {level.lower()}")
        chart_by_place(placed, place_col, places, colours, theme)
    with right:
        st.subheader(f"{level} by year placed")
        chart_place_by_year(placed, place_col, places, theme)

    no_location = len(view) - len(placed)
    unmapped = placed.loc[placed["Area"] == OTHER_AREA, "Location"]
    notes = []
    if no_location:
        notes.append(f"{no_location:,} selected products have no usable location "
                     "and are not shown here.")
    if len(unmapped):
        notes.append(f"{len(unmapped):,} have a location that is not yet assigned "
                     f"to an area ({', '.join(sorted(unmapped.unique()))}).")
    if notes:
        st.caption(" ".join(notes))

# ---- Records --------------------------------------------------------------

st.header("Placement records")
table_columns = ["Product", "Model", "Placement Date", "Location", "Area",
                 "Status", "Years in field", "Participant Name", "Stove ID",
                 "Sensor ID", "Customer Type", "Study Type", "Description",
                 "Tracker rows"]
# The public data file has no contact details, and only institutions are named.
if not is_public and st.checkbox("Show contact details (phone, contact person)"):
    table_columns[8:8] = ["Phone", "Contact Person"]

table = view[table_columns].sort_values("Placement Date", ascending=False,
                                        na_position="last")
if is_public:
    table = (table.drop(columns="Description")
                  .rename(columns={"Participant Name": "Institution"}))
st.dataframe(
    table, hide_index=True, height=420, alt="Table of the selected placement records",
    column_config={
        "Placement Date": st.column_config.DateColumn(format="DD MMM YYYY"),
        "Years in field": st.column_config.NumberColumn(format="%.1f"),
        "Tracker rows": st.column_config.NumberColumn(
            help="How many rows in the tracker carried this stove ID. More than "
                 "1 means repeated rows were merged into this record."),
    })
st.download_button(
    "Download these records (CSV)",
    table.to_csv(index=False, date_format="%Y-%m-%d").encode("utf-8-sig"),
    file_name="placements_filtered.csv", mime="text/csv",
    icon=":material/download:")
