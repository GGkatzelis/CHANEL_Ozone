import io
import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime, timedelta

import matplotlib
import numpy as np
import pandas as pd
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.lines import Line2D
from matplotlib.ticker import MaxNLocator
from matplotlib.figure import Figure

try:
    from PIL import Image
except ImportError:
    Image = None

if os.name == "nt":
    import ctypes
    from ctypes import wintypes

    CF_DIB = 8
    GMEM_MOVEABLE = 0x0002
    GMEM_ZEROINIT = 0x0040
    GHND = GMEM_MOVEABLE | GMEM_ZEROINIT

    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32

    # declare argument and result types for the Win32 API calls we use
    kernel32.GlobalAlloc.argtypes = (wintypes.UINT, ctypes.c_size_t)
    kernel32.GlobalAlloc.restype = wintypes.HGLOBAL
    kernel32.GlobalLock.argtypes = (wintypes.HGLOBAL,)
    kernel32.GlobalLock.restype = wintypes.LPVOID
    kernel32.GlobalUnlock.argtypes = (wintypes.HGLOBAL,)
    kernel32.GlobalFree.argtypes = (wintypes.HGLOBAL,)

    user32.OpenClipboard.argtypes = (wintypes.HWND,)
    user32.OpenClipboard.restype = wintypes.BOOL
    user32.EmptyClipboard.restype = wintypes.BOOL
    user32.SetClipboardData.argtypes = (wintypes.UINT, wintypes.HANDLE)
    user32.SetClipboardData.restype = wintypes.HANDLE
    user32.CloseClipboard.restype = wintypes.BOOL

    def _copy_image_to_clipboard(image):
        output = io.BytesIO()
        image.convert("RGB").save(output, format="BMP")
        data = output.getvalue()[14:]
        output.close()

        size = len(data)
        hglob = kernel32.GlobalAlloc(GHND, size)
        if not hglob:
            raise RuntimeError("GlobalAlloc failed")

        ptr = kernel32.GlobalLock(hglob)
        if not ptr:
            kernel32.GlobalFree(hglob)
            raise RuntimeError("GlobalLock failed")

        # copy data into the global memory block
        ctypes.memmove(ctypes.c_void_p(ptr), data, size)
        kernel32.GlobalUnlock(hglob)

        if not user32.OpenClipboard(None):
            kernel32.GlobalFree(hglob)
            raise RuntimeError("OpenClipboard failed")
        try:
            user32.EmptyClipboard()
            if not user32.SetClipboardData(CF_DIB, hglob):
                raise RuntimeError("SetClipboardData failed")
            # ownership of hglob transferred to the clipboard on success
            hglob = None
        finally:
            user32.CloseClipboard()
            if hglob:
                kernel32.GlobalFree(hglob)
else:
    def _copy_image_to_clipboard(image):
        raise RuntimeError("Clipboard image copy is only supported on Windows")

CSV_PATH = r"Z:\VOC\Campaigns\2024 CHANEL\Ozone\CHANEL_DataUpload_AMSandTracegas_forTS.csv"
EXCEL_PATH = r"Z:\VOC\Campaigns\2024 CHANEL\Ozone\updated_file.xlsx"
OH_PATH = r"Z:\VOC\Campaigns\2024 CHANEL\Ozone\OHreactivity.csv"
TVOC_PATH = r"Z:\VOC\Campaigns\2024 CHANEL\Ozone\CHANEL_DataUpload_kOH.csv"
KOH_MAX_GAP = pd.Timedelta("30min")
BIN_WIDTH = pd.Timedelta("10min")
# marker area (pt^2) of the 10-min scatter points at roof open and at the longest time after it
POINT_SIZE_MAX = 140
POINT_SIZE_MIN = 12
DAY_EDGE_CMAP = matplotlib.colormaps["tab10"]
# chosen edge colours per day; days not listed here fall back to DAY_EDGE_CMAP
DAY_EDGE_COLORS = {
    "2024-07-10": "#9B3F5C",
    "2024-07-11": "#6EA27E",
    "2024-07-22": "#D3968C",
    "2024-07-29": "#988CB7",
    "2024-07-30": "#8096AD",
}
DAY_EDGE_WIDTH = 1.4
# experiment description shown in the figure legends instead of the date
EXPERIMENT_LABELS = {
    "2024-07-08": "VCPs, medium NO",
    "2024-07-09": "VCPs, high NO",
    "2024-07-10": "gasoline, high NO",
    "2024-07-11": "gasoline, medium NO",
    "2024-07-15": "cooking, medium NO",
    "2024-07-17": "diesel, medium NO",
    "2024-07-18": "diesel, high NO",
    "2024-07-22": "LA Anthropogenic, medium NO",
    "2024-07-29": "LA Anthropogenic+Biogenic, medium NO",
    "2024-07-30": "Global Anthropogenic, medium NO",
}
ROX_PATH =r"Z:\VOC\Campaigns\2024 CHANEL\Ozone\CHANEL_DataUpload_ROx.csv"
ROX_COLUMNS = ["OH_LIF", "HO2_LIF", "RO2_LIF"]
ROX_EXPORT_NAME = "merged_Rox_koh_nox_ozone_10min.csv"
# figure name -> (x column, x label, colour column, colour label)
OPR_10MIN_FIGURES = {
    f"OPR vs. {x_name} (colored by {species})": (x_col, x_label, f"{species}_LIF", f"{species} (molecules cm-3)")
    for x_name, x_col, x_label in [("TVOC", "tvoc_s1", "TVOC (s-1)"), ("NOx", "nox", "NOx (ppb)")]
    for species in ["HO2", "RO2", "OH"]
}
OPR_10MIN_FIGURES["OPR vs. NOx (colored by TVOC)"] = ("nox", "NOx (ppb)", "tvoc_s1", "TVOC (s-1)")
TARGET_DATES = [
    "2024-07-08", "2024-07-09", "2024-07-10", "2024-07-11",
    "2024-07-15", "2024-07-17", "2024-07-18", "2024-07-22",
    "2024-07-29", "2024-07-30",
]
OZONE_COLUMN = "O3_TwoBTech_1014"
NOX_COLUMN = "NOx_CLD_TR43"
NO_COLUMN = "NO_CLD_TR43"
NO2_COLUMN = "NO2_CLD_TR43"
# CSV column -> short name used throughout the app, plotted on the secondary axis in tab 1
NITROGEN_COLUMNS = {NOX_COLUMN: "nox", NO_COLUMN: "no", NO2_COLUMN: "no2"}
NITROGEN_STYLES = [
    ("nox", "NOx", "#9467bd"),
    ("no", "NO", "#8c564b"),
    ("no2", "NO2", "#e377c2"),
]
ENDTIME_COLUMN = "Endtime_UTC"
REACTIONTIME_COLUMN = "reactiontime_UTC"


def find_datetime_column(columns):
    lowered = [c.lower() for c in columns]
    for key in ["datetime", "date_time", "date/time", "date time", "time", "date"]:
        if key in lowered:
            return columns[lowered.index(key)]
    # If date and time are stored separately, return both columns as a tuple.
    date_cols = [c for c in columns if any(k in c.lower() for k in ["date", "day"])]
    time_cols = [c for c in columns if any(k in c.lower() for k in ["time", "hour"])]
    if date_cols and time_cols:
        return date_cols[0], time_cols[0]
    return None


def find_ozone_column(columns):
    lowered = [c.lower() for c in columns]
    for key in ["o3", "ozone", "ozone concentration", "o3 concentration"]:
        if key in lowered:
            return columns[lowered.index(key)]
    for c in columns:
        if "o3" in c.lower() or "ozone" in c.lower():
            return c
    # fallback by numeric columns
    return None


def load_dataset(path):
    if not os.path.exists(path):
        raise FileNotFoundError(f"CSV file not found: {path}")

    df = pd.read_csv(path, low_memory=False)
    if df.empty:
        raise ValueError("CSV file is empty")

    dt_col = find_datetime_column(df.columns)
    if dt_col is None:
        raise ValueError("Could not find a date or datetime column in the CSV")

    if isinstance(dt_col, tuple):
        date_col, time_col = dt_col
        timestamp_series = (
            df[date_col].astype(str).str.strip() + " " + df[time_col].astype(str).str.strip()
        )
    else:
        timestamp_series = df[dt_col].astype(str).str.replace(r"\s+", " ", regex=True).str.strip()

    timestamps = pd.to_datetime(timestamp_series, errors="coerce")
    if timestamps.isna().all():
        timestamps = pd.to_datetime(
            timestamp_series,
            format="%m/%d/%Y %I:%M:%S %p",
            errors="coerce",
        )

    if timestamps.isna().all():
        raise ValueError("Could not parse any timestamps from the CSV")

    ozone_col = OZONE_COLUMN if OZONE_COLUMN in df.columns else find_ozone_column(df.columns)
    if ozone_col is None:
        raise ValueError(f"Could not find ozone column '{OZONE_COLUMN}' in the CSV")

    columns = [ozone_col]
    present_nitrogen = [c for c in NITROGEN_COLUMNS if c in df.columns]
    columns.extend(present_nitrogen)

    df = df[columns].copy()
    df.insert(0, "__timestamp", timestamps)
    df = df.rename(columns={ozone_col: "ozone", **NITROGEN_COLUMNS})
    for name in (NITROGEN_COLUMNS[c] for c in present_nitrogen):
        df[name] = pd.to_numeric(df[name], errors="coerce")

    df = df.sort_values("__timestamp").reset_index(drop=True)
    df = df.dropna(subset=["__timestamp"])

    return df


def load_excel_dataset(path):
    if not os.path.exists(path):
        raise FileNotFoundError(f"Excel file not found: {path}")

    df = pd.read_excel(path)
    if df.empty:
        raise ValueError("Excel file is empty")

    if "Date" not in df.columns:
        raise ValueError("Excel file missing required 'Date' column")
    if ENDTIME_COLUMN not in df.columns:
        raise ValueError(f"Excel file missing required '{ENDTIME_COLUMN}' column")
    if REACTIONTIME_COLUMN not in df.columns:
        raise ValueError(f"Excel file missing required '{REACTIONTIME_COLUMN}' column")

    df["__date"] = pd.to_datetime(df["Date"], errors="coerce").dt.date
    df[ENDTIME_COLUMN] = pd.to_datetime(df[ENDTIME_COLUMN], errors="coerce")
    df[REACTIONTIME_COLUMN] = pd.to_datetime(df[REACTIONTIME_COLUMN], errors="coerce")

    injection_columns = [c for c in df.columns if c.lower().endswith("_injectiontime")]
    if not injection_columns:
        raise ValueError("Excel file has no columns ending with '_injectiontime'")

    for col in injection_columns:
        df[col] = pd.to_datetime(df[col], errors="coerce")

    df = df[["__date", ENDTIME_COLUMN, REACTIONTIME_COLUMN] + injection_columns].dropna(subset=["__date"])

    return df


def load_oh_reactivity(path):
    if not os.path.exists(path):
        raise FileNotFoundError(f"OH reactivity file not found: {path}")

    df = pd.read_csv(path)
    if df.empty:
        raise ValueError("OH reactivity file is empty")
    if "Date" not in df.columns or "kOH_LP_sync_avg" not in df.columns:
        raise ValueError("OH reactivity file must contain 'Date' and 'kOH_LP_sync_avg' columns")

    df["__date"] = pd.to_datetime(df["Date"].astype(str).str.replace("_", "-"), format="%Y-%m-%d", errors="coerce").dt.date
    df = df[["__date", "kOH_LP_sync_avg"]].dropna(subset=["__date"])
    return df.set_index("__date")


def compute_tab2_data(df, df_excel, df_oh):
    target_dates = [
        "2024-07-08", "2024-07-09", "2024-07-10", "2024-07-11",
        "2024-07-15", "2024-07-17", "2024-07-18", "2024-07-22",
        "2024-07-29", "2024-07-30",
    ]

    rows = []
    for date_string in target_dates:
        date_obj = datetime.fromisoformat(date_string).date()
        day_df = df[df["__timestamp"].dt.date == date_obj]
        if day_df.empty:
            continue

        excel_day = df_excel[df_excel["__date"] == date_obj]
        if excel_day.empty:
            continue

        roof_open_times = excel_day[REACTIONTIME_COLUMN].dropna().unique()
        if len(roof_open_times) == 0:
            continue

        roof_open = min(roof_open_times)
        background_start = roof_open - timedelta(hours=1)
        background_period = day_df[(day_df["__timestamp"] >= background_start) & (day_df["__timestamp"] <= roof_open)]
        if background_period.empty:
            continue
        background_ozone = background_period["ozone"].mean()

        peak_rows = []
        for endtime in excel_day[ENDTIME_COLUMN].dropna().unique():
            before_end = day_df[day_df["__timestamp"] < endtime].dropna(subset=["ozone"])
            if not before_end.empty:
                idxmax = before_end["ozone"].idxmax()
                peak_rows.append(before_end.loc[idxmax])
        if not peak_rows:
            continue
        peak_row = max(peak_rows, key=lambda r: r["ozone"])
        peak_ozone = peak_row["ozone"]
        peak_time = peak_row["__timestamp"]
        ozone_increase = peak_ozone - background_ozone
        elapsed_hr = (peak_time - roof_open).total_seconds() / 3600.0
        if elapsed_hr <= 0:
            continue
        experiment_opr = ozone_increase / elapsed_hr

        nox_avg = None
        if "nox" in day_df.columns:
            nox_period = day_df[(day_df["__timestamp"] >= roof_open) & (day_df["__timestamp"] <= peak_time)]
            if not nox_period.empty:
                nox_avg = nox_period["nox"].mean()
        if pd.isna(nox_avg):
            continue

        if date_obj not in df_oh.index:
            continue
        koh = df_oh.loc[date_obj, "kOH_LP_sync_avg"]
        if pd.isna(koh):
            continue

        rows.append({
            "date": date_obj,
            "background_ozone": background_ozone,
            "peak_ozone": peak_ozone,
            "ozone_increase": ozone_increase,
            "kOH_LP_sync_avg": koh,
            "nox_avg": nox_avg,
            "experiment_opr": experiment_opr,
        })

    return pd.DataFrame(rows)


def load_tvoc_dataset(path):
    if not os.path.exists(path):
        raise FileNotFoundError(f"TVOC file not found: {path}")

    df = pd.read_csv(path, low_memory=False)
    if df.empty:
        raise ValueError("TVOC file is empty")
    if "time" not in df.columns or "kOH" not in df.columns:
        raise ValueError("TVOC file must contain 'time' and 'kOH' columns")

    df["__timestamp"] = pd.to_datetime(df["time"], errors="coerce")
    if df["__timestamp"].isna().all():
        raise ValueError("Could not parse any timestamps from the TVOC file")

    df["kOH"] = pd.to_numeric(df["kOH"], errors="coerce")
    return df.dropna(subset=["__timestamp"]).sort_values("__timestamp").reset_index(drop=True)


def load_rox_dataset(path):
    if not os.path.exists(path):
        raise FileNotFoundError(f"ROx file not found: {path}")

    df = pd.read_csv(path, usecols=lambda c: c == "time" or c in ROX_COLUMNS, low_memory=False)
    if df.empty:
        raise ValueError("ROx file is empty")
    missing = [c for c in ["time"] + ROX_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"ROx file missing required columns: {', '.join(missing)}")

    df["__timestamp"] = pd.to_datetime(df["time"], errors="coerce")
    if df["__timestamp"].isna().all():
        raise ValueError("Could not parse any timestamps from the ROx file")

    for col in ROX_COLUMNS:
        df[col] = pd.to_numeric(df[col], errors="coerce")
        # negative concentrations are below detection limit noise; drop the value, keep the row's other species
        df.loc[df[col] < 0, col] = float("nan")
    df = df[["__timestamp"] + ROX_COLUMNS].dropna(subset=["__timestamp"])
    return df.sort_values("__timestamp").reset_index(drop=True)


def _resample_10min(df, columns):
    series = df.set_index("__timestamp")[columns]
    # files parse at different datetime resolutions; align them so the 10-min bins join exactly
    series.index = series.index.astype("datetime64[ns]")
    return series.resample(BIN_WIDTH).mean()


def _interpolate_koh(tvoc_df, times):
    # kOH is sampled only every ~6-14 min, so many 10-min windows hold no sample; interpolate linearly
    # in time instead, but not across gaps longer than KOH_MAX_GAP (instrument off, between days)
    koh = tvoc_df[["__timestamp", "kOH"]].dropna()
    t = koh["__timestamp"].astype("datetime64[ns]").astype("int64").to_numpy()
    v = koh["kOH"].to_numpy()
    x = pd.DatetimeIndex(times).astype("datetime64[ns]").asi8
    values = np.interp(x, t, v, left=np.nan, right=np.nan)

    hi = np.clip(np.searchsorted(t, x), 1, len(t) - 1)
    gap = t[hi] - t[hi - 1]
    exact = t[hi] == x
    values[(gap > KOH_MAX_GAP.value) & ~exact] = np.nan
    return values


def compute_tab2_10min_data(df, df_excel, tvoc_df, rox_df=None):
    rows = []
    # ROx is averaged into the same 10-min windows as ozone and NOx; kOH is interpolated to the window centres
    rox_10min = _resample_10min(rox_df, ROX_COLUMNS) if rox_df is not None else None

    for date_string in TARGET_DATES:
        date_obj = datetime.fromisoformat(date_string).date()
        day_df = df[df["__timestamp"].dt.date == date_obj].copy()
        if day_df.empty:
            continue

        excel_day = df_excel[df_excel["__date"] == date_obj]
        if excel_day.empty:
            continue

        roof_open_times = excel_day[REACTIONTIME_COLUMN].dropna().unique()
        if len(roof_open_times) == 0:
            continue

        roof_open = min(roof_open_times)
        # mean ozone in the hour before roof open, treated as the ozone value at roof open
        background_start = roof_open - timedelta(hours=1)
        background_period = day_df[(day_df["__timestamp"] >= background_start) & (day_df["__timestamp"] <= roof_open)]
        background_ozone = background_period["ozone"].mean()

        peak_rows = []
        for endtime in excel_day[ENDTIME_COLUMN].dropna().unique():
            before_end = day_df[day_df["__timestamp"] < endtime].dropna(subset=["ozone"])
            if not before_end.empty:
                idxmax = before_end["ozone"].idxmax()
                peak_rows.append(before_end.loc[idxmax])
        if not peak_rows:
            continue
        peak_row = max(peak_rows, key=lambda r: r["ozone"])
        peak_time = peak_row["__timestamp"]
        if peak_time <= roof_open:
            continue

        if "nox" not in day_df.columns:
            continue

        resampled = _resample_10min(day_df, ["ozone", "nox"]).dropna(how="all")
        resampled = resampled[(resampled.index > roof_open) & (resampled.index <= peak_time)]
        if resampled.empty:
            continue

        # OPR is the local slope to the previous bin that has ozone: (O3_i - O3_prev) / (t_i - t_prev), with
        # times at bin centres. Empty bins are skipped, so the step becomes 20, 30, ... min instead of losing
        # the point. The first bin is referenced to the background ozone at roof open.
        ozone = resampled["ozone"].dropna()
        times = pd.Series(ozone.index + BIN_WIDTH / 2, index=ozone.index)
        prev_ozone = ozone.shift(1)
        prev_time = times.shift(1)
        if not ozone.empty:
            prev_ozone.iloc[0] = background_ozone  # NaN if there was no background data
            prev_time.iloc[0] = roof_open
        resampled["opr_ppbv_h"] = (ozone - prev_ozone) / ((times - prev_time) / pd.Timedelta("1h"))

        resampled["kOH"] = _interpolate_koh(tvoc_df, resampled.index + BIN_WIDTH / 2)
        if rox_10min is not None:
            resampled = resampled.join(rox_10min, how="left")
        merged = resampled.rename_axis("__timestamp").reset_index()

        for _, row in merged.iterrows():
            if pd.isna(row["kOH"]) or pd.isna(row["opr_ppbv_h"]) or pd.isna(row["nox"]):
                continue

            # bins are labelled by their start, but the averaged values represent the bin centre
            elapsed_hr = (row["__timestamp"] + BIN_WIDTH / 2 - roof_open).total_seconds() / 3600.0
            if elapsed_hr <= 0:
                continue

            out_row = {
                "date": date_obj,
                "timestamp": row["__timestamp"],
                "elapsed_h": elapsed_hr,
                "opr_ppbv_h": row["opr_ppbv_h"],
                "tvoc_s1": row["kOH"],
                "nox": row["nox"],
            }
            if rox_df is not None:
                out_row["ozone"] = row["ozone"]
                for col in ROX_COLUMNS:
                    out_row[col] = row[col]
            rows.append(out_row)

    return pd.DataFrame(rows)


def get_available_dates(df):
    if df is None or df.empty:
        return []
    dates = df["__timestamp"].dt.date.dropna().unique()
    return [d.isoformat() for d in sorted(dates)]


def filter_day(df, selected_date):
    if not selected_date:
        return df.iloc[0:0]
    date_obj = datetime.fromisoformat(selected_date).date()
    return df[df["__timestamp"].dt.date == date_obj]


def create_tab1(frame):
    file_label = ttk.Label(frame, text=f"Ozone CSV file: {CSV_PATH}")
    file_label.pack(anchor="w", padx=10, pady=(10, 0))

    control_frame = ttk.Frame(frame)
    control_frame.pack(fill="x", padx=10, pady=10)

    ttk.Label(control_frame, text="Select date:").grid(row=0, column=0, sticky="w")
    date_var = tk.StringVar()
    date_combo = ttk.Combobox(control_frame, textvariable=date_var, state="readonly", width=24)
    date_combo.grid(row=0, column=1, padx=(8, 0), sticky="w")

    refresh_button = ttk.Button(control_frame, text="Plot day")
    refresh_button.grid(row=0, column=2, padx=(10, 0), sticky="w")

    clipboard_button = ttk.Button(control_frame, text="Copy figure to clipboard")
    clipboard_button.grid(row=0, column=3, padx=(10, 0), sticky="w")

    message_label = ttk.Label(frame, text="", foreground="red")
    message_label.pack(anchor="w", padx=10)

    fig = Figure(figsize=(8, 4), dpi=100)
    ax = fig.add_subplot(111)
    # nitrogen species are an order of magnitude smaller than ozone, so give them their own axis
    ax_n = ax.twinx()
    ax.set_title("Ozone and NOx time series")
    ax.set_xlabel("Time")
    ax.set_ylabel("Ozone")
    ax_n.set_ylabel("NO / NO2 / NOx")
    fig.tight_layout()

    canvas = FigureCanvasTkAgg(fig, master=frame)
    canvas.get_tk_widget().pack(fill="both", expand=True, padx=10, pady=(0, 10))

    try:
        df = load_dataset(CSV_PATH)
        df_excel = load_excel_dataset(EXCEL_PATH)
        dates = get_available_dates(df)
        if not dates:
            raise ValueError("No valid data rows found after parsing timestamps")
        date_combo["values"] = dates
        date_combo.current(0)
    except Exception as exc:
        message_label.config(text=str(exc))
        refresh_button.config(state="disabled")
        return

    def draw_plot():
        selected = date_var.get()
        if not selected:
            messagebox.showwarning("Select date", "Please choose a date first.")
            return
        day_df = filter_day(df, selected)
        if day_df.empty:
            messagebox.showinfo("No data", f"No ozone data found for {selected}.")
            return

        ax.clear()
        ax_n.clear()
        # clear() restores the twin's background patch, which would hide the ozone line beneath it,
        # and resets its y-label back to the left-hand side
        ax_n.patch.set_visible(False)
        ax_n.yaxis.set_label_position("right")
        ax_n.yaxis.tick_right()
        ax.plot(day_df["__timestamp"], day_df["ozone"], marker="o", linestyle="-", color="#1f77b4", label="Ozone")

        for name, label, color in NITROGEN_STYLES:
            if name in day_df.columns and day_df[name].notna().any():
                ax_n.plot(day_df["__timestamp"], day_df[name], linestyle="-", linewidth=1.2,
                          color=color, label=label)

        selected_date = datetime.fromisoformat(selected).date()
        excel_day = df_excel[df_excel["__date"] == selected_date]
        if not excel_day.empty:
            endtime_labeled = False
            peak_labeled = False
            injection_cols = [c for c in excel_day.columns if c.lower().endswith("_injectiontime")]
            for endtime in excel_day[ENDTIME_COLUMN].dropna().unique():
                ax.axvline(endtime, color="#ff7f0e", linestyle="--", linewidth=2,
                           label="Endtime_UTC" if not endtime_labeled else None)
                endtime_labeled = True

                before = day_df[day_df["__timestamp"] < endtime].copy()
                before = before.dropna(subset=["ozone"])
                if not before.empty:
                    idxmax = before["ozone"].idxmax()
                    peak_row = before.loc[idxmax]
                    ax.scatter([peak_row["__timestamp"]], [peak_row["ozone"]], color="red", s=60, zorder=5,
                               label="Peak before Endtime" if not peak_labeled else None)
                    peak_labeled = True

            injection_times = []
            for _, row in excel_day.iterrows():
                for col in injection_cols:
                    if pd.notna(row[col]):
                        injection_times.append(row[col])
            if injection_times:
                earliest_injection = min(injection_times) - timedelta(hours=2)
                ax.axvline(earliest_injection, color="gray", linestyle="--", linewidth=2,
                           label="earliest injection time_UTC")

            roof_open_labeled = False
            for reaction_time in excel_day[REACTIONTIME_COLUMN].dropna().unique():
                ax.axvline(reaction_time, color="green", linestyle="--", linewidth=2,
                           label="roof open" if not roof_open_labeled else None)
                roof_open_labeled = True

        ax.set_title(f"Ozone and NOx time series for {selected}")
        ax.set_xlabel("Time")
        ax.set_ylabel("Ozone")
        ax_n.set_ylabel("NO / NO2 / NOx")
        ax.grid(True, linestyle="--", alpha=0.4)
        handles, labels = ax.get_legend_handles_labels()
        n_handles, n_labels = ax_n.get_legend_handles_labels()
        if handles or n_handles:
            ax.legend(handles + n_handles, labels + n_labels, loc="upper left", fontsize=8)
        fig.autofmt_xdate(rotation=25)
        canvas.draw()

    def copy_figure_to_clipboard():
        if Image is None:
            messagebox.showerror(
                "Missing dependency",
                "Pillow is required to copy figures to the clipboard. Install it with: pip install pillow",
            )
            return

        if os.name != "nt":
            messagebox.showerror(
                "Unsupported platform",
                "Clipboard image copy is supported only on Windows.",
            )
            return

        buffer = io.BytesIO()
        fig.savefig(buffer, format="png")
        buffer.seek(0)
        image = Image.open(buffer)

        try:
            _copy_image_to_clipboard(image)
            messagebox.showinfo("Copied", "Figure image copied to clipboard.")
        except Exception as exc:
            messagebox.showerror("Clipboard error", str(exc))

    refresh_button.config(command=draw_plot)
    clipboard_button.config(command=copy_figure_to_clipboard)
    draw_plot()


def build_gui():
    root = tk.Tk()
    root.title("CHANEL Ozone Explorer")
    root.geometry("980x640")

    notebook = ttk.Notebook(root)
    notebook.pack(fill="both", expand=True)

    tab1 = ttk.Frame(notebook)
    notebook.add(tab1, text="Tab 1: Ozone day plot")
    create_tab1(tab1)

    tab2 = ttk.Frame(notebook)
    notebook.add(tab2, text="Tab 2: OPR Figures")
    create_tab2(tab2)

    tab3 = ttk.Frame(notebook)
    notebook.add(tab3, text="Tab 3")
    ttk.Label(tab3, text="Future tab content goes here.").pack(padx=16, pady=16)

    root.mainloop()


def create_tab2(frame):
    file_label = ttk.Label(frame, text=f"OH reactivity file: {OH_PATH}")
    file_label.pack(anchor="w", padx=10, pady=(10, 0))

    file_label2 = ttk.Label(frame, text=f"TVOC file: {TVOC_PATH}")
    file_label2.pack(anchor="w", padx=10, pady=(0, 0))

    file_label3 = ttk.Label(frame, text=f"ROx file: {ROX_PATH}")
    file_label3.pack(anchor="w", padx=10, pady=(0, 10))

    control_frame = ttk.Frame(frame)
    control_frame.pack(fill="x", padx=10, pady=10)

    ttk.Label(control_frame, text="Select figure:").grid(row=0, column=0, sticky="w")
    figure_var = tk.StringVar()
    figure_combo = ttk.Combobox(
        control_frame,
        textvariable=figure_var,
        state="readonly",
        values=[
            "Ozone enhancement for each experiment",
            "OPR for each experiment",
            "OPR vs. TVOC (colored by NOx)",
            "NOx vs. TVOC",
            *OPR_10MIN_FIGURES,
        ],
        width=40,
    )
    figure_combo.grid(row=0, column=1, padx=(8, 0), sticky="w")
    figure_combo.current(0)

    message_label = ttk.Label(frame, text="", foreground="red")
    message_label.pack(anchor="w", padx=10)

    button_frame = ttk.Frame(frame)
    button_frame.pack(fill="x", padx=10, pady=(0, 8))
    clipboard_button = ttk.Button(button_frame, text="Copy figure to clipboard")
    clipboard_button.pack(side="left")
    download_button = ttk.Button(button_frame, text="Download merged ROx/kOH/NOx/ozone 10-min CSV")
    download_button.pack(side="left", padx=(10, 0))

    fig = Figure(figsize=(8, 4), dpi=100)
    ax = fig.add_subplot(111)
    canvas = FigureCanvasTkAgg(fig, master=frame)
    canvas.get_tk_widget().pack(fill="both", expand=True, padx=10, pady=(0, 10))
    current_cbar = None
    # hover tooltip state for the 10-min scatter figures; reset on every redraw
    hover_scatter = None
    hover_labels = []
    hover_annot = None

    try:
        df = load_dataset(CSV_PATH)
        df_excel = load_excel_dataset(EXCEL_PATH)
        df_oh = load_oh_reactivity(OH_PATH)
        tvoc_df = load_tvoc_dataset(TVOC_PATH)
        rox_df = load_rox_dataset(ROX_PATH)
        tab2_data = compute_tab2_data(df, df_excel, df_oh)
        tab2_10min_data = compute_tab2_10min_data(df, df_excel, tvoc_df, rox_df)
    except Exception as exc:
        message_label.config(text=str(exc))
        download_button.config(state="disabled")
        return

    def window_labels(data):
        starts = pd.to_datetime(data["timestamp"])
        return [f"{s:%Y-%m-%d %H:%M}–{s + timedelta(minutes=10):%H:%M}" for s in starts]

    # 10-min points shrink with time since roof open; one scale for all figures so sizes compare across them
    size_max_h = tab2_10min_data["elapsed_h"].max() if not tab2_10min_data.empty else 1.0

    def point_sizes(elapsed_h):
        fraction = (np.asarray(elapsed_h, dtype=float) / size_max_h).clip(0, 1)
        return POINT_SIZE_MAX - (POINT_SIZE_MAX - POINT_SIZE_MIN) * fraction

    def add_size_legend():
        hours = [h for h in MaxNLocator(nbins=4).tick_values(0, size_max_h) if 0 < h <= size_max_h]
        handles = [
            Line2D([], [], linestyle="", marker="o", markersize=np.sqrt(point_sizes(h)),
                   markerfacecolor="lightgray", markeredgecolor="black", markeredgewidth=0.5)
            for h in hours
        ]
        ax.legend(handles, [f"{h:g} h" for h in hours], title="Time since roof open",
                  loc="best", fontsize=8, title_fontsize=8, labelspacing=1.0, borderpad=0.8)

    # one fixed edge colour per experiment day, so a day keeps its colour in every figure
    day_colors = {
        datetime.fromisoformat(d).date(): DAY_EDGE_COLORS.get(d, DAY_EDGE_CMAP(i % DAY_EDGE_CMAP.N))
        for i, d in enumerate(TARGET_DATES)
    }

    def day_edge_colors(data):
        return [day_colors.get(d, "black") for d in data["date"]]

    def add_day_legend(data):
        # placed in the strip reserved right of the colour bar, so it never covers data
        days = [d for d in day_colors if d in set(data["date"])]
        handles = [
            Line2D([], [], linestyle="", marker="o", markersize=8, markerfacecolor="white",
                   markeredgecolor=day_colors[d], markeredgewidth=2.0)
            for d in days
        ]
        legend = fig.legend(handles, [EXPERIMENT_LABELS.get(d.isoformat(), d.isoformat()) for d in days],
                            title="Experiment", loc="center right", fontsize=8, title_fontsize=8, frameon=False)
        # fraction of the figure width the legend needs, so the plot can be squeezed to make room for it
        width = legend.get_window_extent(fig.canvas.get_renderer()).width
        return min(width / fig.bbox.width + 0.01, 0.5)

    def draw_tab2():
        nonlocal current_cbar, ax, hover_scatter, hover_labels, hover_annot
        fig.clear()
        ax = fig.add_subplot(111)
        current_cbar = None
        hover_scatter = None
        hover_labels = []
        hover_annot = None

        selected = figure_var.get().strip()
        day_legend_data = None

        if selected in OPR_10MIN_FIGURES:
            x_col, x_label, c_col, c_label = OPR_10MIN_FIGURES[selected]
            plot_data = tab2_10min_data.dropna(subset=[x_col, "opr_ppbv_h", c_col]) if not tab2_10min_data.empty else tab2_10min_data
            if plot_data.empty:
                message_label.config(text=f"No valid 10-minute OPR data with {c_label} available.")
                canvas.draw()
                return
            message_label.config(text="")
            ax.set_title(selected)
            ax.set_xlabel(x_label)
            ax.set_ylabel("OPR (ppbv/h)")
            scatter = ax.scatter(
                plot_data[x_col],
                plot_data["opr_ppbv_h"],
                c=plot_data[c_col],
                cmap="YlOrRd",
                s=point_sizes(plot_data["elapsed_h"]),
                edgecolor=day_edge_colors(plot_data),
                linewidth=DAY_EDGE_WIDTH,
            )
            current_cbar = fig.colorbar(scatter, ax=ax)
            current_cbar.set_label(c_label)
            add_size_legend()
            day_legend_data = plot_data
            hover_scatter, hover_labels = scatter, window_labels(plot_data)

        elif "Ozone enhancement" in selected:
            if tab2_data.empty:
                message_label.config(text="No valid experiment-level data available.")
                canvas.draw()
                return
            message_label.config(text="")
            ax.set_title("Ozone enhancement for each experiment")
            ax.set_xlabel("Injected TVOC (s-1)")
            ax.set_ylabel("Ozone enhancement (ppb)")
            scatter = ax.scatter(
                tab2_data["kOH_LP_sync_avg"],
                tab2_data["ozone_increase"],
                c=tab2_data["nox_avg"],
                cmap="YlOrRd",
                s=120,
                edgecolor="black",
                linewidth=0.5,
            )
            for _, row in tab2_data.iterrows():
                ax.annotate(
                    row["date"].isoformat(),
                    (row["kOH_LP_sync_avg"], row["ozone_increase"]),
                    textcoords="offset points",
                    xytext=(5, 5),
                    ha="left",
                    fontsize=8,
                )
            current_cbar = fig.colorbar(scatter, ax=ax)
            current_cbar.set_label("NOx avg (ppb)")

        elif "OPR for each experiment" in selected:
            if tab2_data.empty:
                message_label.config(text="No valid experiment-level OPR data available.")
                canvas.draw()
                return
            message_label.config(text="")
            ax.set_title("OPR for each experiment")
            ax.set_xlabel("Injected TVOC (s-1)")
            ax.set_ylabel("Experiment OPR (ppbv/h)")
            scatter = ax.scatter(
                tab2_data["kOH_LP_sync_avg"],
                tab2_data["experiment_opr"],
                c=tab2_data["nox_avg"],
                cmap="YlOrRd",
                s=120,
                edgecolor="black",
                linewidth=0.5,
            )
            for _, row in tab2_data.iterrows():
                ax.annotate(
                    row["date"].isoformat(),
                    (row["kOH_LP_sync_avg"], row["experiment_opr"]),
                    textcoords="offset points",
                    xytext=(5, 5),
                    ha="left",
                    fontsize=8,
                )
            current_cbar = fig.colorbar(scatter, ax=ax)
            current_cbar.set_label("NOx avg (ppb)")

        elif selected == "OPR vs. TVOC (colored by NOx)":
            if tab2_10min_data.empty:
                message_label.config(text="No valid 10-minute OPR data available.")
                canvas.draw()
                return
            message_label.config(text="")
            ax.set_title("OPR vs. TVOC (colored by NOx)")
            ax.set_xlabel("TVOC (s-1)")
            ax.set_ylabel("OPR (ppbv/h)")
            scatter = ax.scatter(
                tab2_10min_data["tvoc_s1"],
                tab2_10min_data["opr_ppbv_h"],
                c=tab2_10min_data["nox"],
                cmap="YlOrRd",
                s=point_sizes(tab2_10min_data["elapsed_h"]),
                edgecolor=day_edge_colors(tab2_10min_data),
                linewidth=DAY_EDGE_WIDTH,
            )
            current_cbar = fig.colorbar(scatter, ax=ax)
            current_cbar.set_label("NOx (ppb)")
            add_size_legend()
            day_legend_data = tab2_10min_data
            hover_scatter, hover_labels = scatter, window_labels(tab2_10min_data)

        elif "NOx vs. TVOC" in selected:
            if tab2_10min_data.empty:
                message_label.config(text="No valid 10-minute OPR data available.")
                canvas.draw()
                return
            message_label.config(text="")
            ax.set_title("NOx vs. TVOC")
            ax.set_xlabel("TVOC (s-1)")
            ax.set_ylabel("NOx (ppb)")
            scatter = ax.scatter(
                tab2_10min_data["tvoc_s1"],
                tab2_10min_data["nox"],
                c=tab2_10min_data["opr_ppbv_h"],
                cmap="YlOrRd",
                s=point_sizes(tab2_10min_data["elapsed_h"]),
                edgecolor=day_edge_colors(tab2_10min_data),
                linewidth=DAY_EDGE_WIDTH,
            )
            current_cbar = fig.colorbar(scatter, ax=ax)
            current_cbar.set_label("OPR (ppbv/h)")
            add_size_legend()
            day_legend_data = tab2_10min_data
            hover_scatter, hover_labels = scatter, window_labels(tab2_10min_data)

            x_min = 0
            x_max = 100
            y_min = 0
            y_max = 8
            for ratio, label in [(0.05, "1:20"), (0.1, "1:10"), (0.2, "1:5"), (1.0, "1:1")]:
                y_vals = [ratio * x for x in [x_min, x_max]]
                ax.plot([x_min, x_max], y_vals, linestyle="--", color="gray", linewidth=1.0, alpha=0.7)
                x_label = min(x_max * 0.85, y_max / ratio * 0.85)
                y_label = ratio * x_label
                if y_label > y_max * 0.95:
                    y_label = y_max * 0.95
                ax.text(x_label, y_label, label, fontsize=8, color="gray", ha="left", va="bottom")

            ax.set_xscale("linear")
            ax.set_yscale("linear")
            ax.set_xlim(0, 100)
            ax.set_ylim(0, 8)

        # only the per-experiment figures keep a grid; the 10-min scatter figures are drawn without one
        if "for each experiment" in selected:
            ax.grid(True, linestyle="--", alpha=0.4)
        if day_legend_data is not None:
            legend_fraction = add_day_legend(day_legend_data)
            fig.tight_layout(rect=(0, 0, 1 - legend_fraction, 1))
        else:
            fig.tight_layout()
        if hover_scatter is not None:
            hover_annot = ax.annotate(
                "",
                xy=(0, 0),
                xytext=(10, 10),
                textcoords="offset points",
                fontsize=8,
                bbox=dict(boxstyle="round", fc="white", ec="gray", alpha=0.9),
            )
            hover_annot.set_visible(False)
        canvas.draw()

    def on_hover(event):
        if hover_scatter is None or hover_annot is None:
            return
        hit, info = (False, {})
        if event.inaxes is hover_scatter.axes:
            hit, info = hover_scatter.contains(event)
        if hit:
            # markers overlap, so of all points under the cursor pick the one whose centre is closest
            offsets = hover_scatter.get_offsets()
            hits = info["ind"]
            pixels = hover_scatter.axes.transData.transform(offsets[hits])
            idx = hits[np.argmin(np.hypot(pixels[:, 0] - event.x, pixels[:, 1] - event.y))]
            hover_annot.xy = offsets[idx]
            hover_annot.set_text(hover_labels[idx])
            hover_annot.set_visible(True)
            canvas.draw_idle()
        elif hover_annot.get_visible():
            hover_annot.set_visible(False)
            canvas.draw_idle()

    def copy_figure_to_clipboard():
        if Image is None:
            messagebox.showerror(
                "Missing dependency",
                "Pillow is required to copy figures to the clipboard. Install it with: pip install pillow",
            )
            return

        if os.name != "nt":
            messagebox.showerror(
                "Unsupported platform",
                "Clipboard image copy is supported only on Windows.",
            )
            return

        buffer = io.BytesIO()
        fig.savefig(buffer, format="png")
        buffer.seek(0)
        image = Image.open(buffer)

        try:
            _copy_image_to_clipboard(image)
            messagebox.showinfo("Copied", "Figure image copied to clipboard.")
        except Exception as exc:
            messagebox.showerror("Clipboard error", str(exc))

    def download_merged_csv():
        if tab2_10min_data.empty:
            messagebox.showinfo("No data", "No 10-minute merged data to save.")
            return
        path = filedialog.asksaveasfilename(
            title="Save merged 10-minute data",
            initialdir=os.path.dirname(ROX_PATH),
            initialfile=ROX_EXPORT_NAME,
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if not path:
            return
        export = tab2_10min_data.rename(columns={"tvoc_s1": "kOH"})
        try:
            export.to_csv(path, index=False)
            messagebox.showinfo("Saved", f"Saved {len(export)} rows to:\n{path}")
        except Exception as exc:
            messagebox.showerror("Save error", str(exc))

    clipboard_button.config(command=copy_figure_to_clipboard)
    download_button.config(command=download_merged_csv)
    canvas.mpl_connect("motion_notify_event", on_hover)
    figure_var.trace_add('write', lambda *args: draw_tab2())
    draw_tab2()


if __name__ == "__main__":
    build_gui()
