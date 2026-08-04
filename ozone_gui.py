import io
import os
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime, timedelta

import pandas as pd
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
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
OZONE_COLUMN = "O3_TwoBTech_1014"
NOX_COLUMN = "NOx_CLD_TR43"
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

    df["__timestamp"] = pd.to_datetime(timestamp_series, errors="coerce")
    if df["__timestamp"].isna().all():
        df["__timestamp"] = pd.to_datetime(
            timestamp_series,
            format="%m/%d/%Y %I:%M:%S %p",
            errors="coerce",
        )

    if df["__timestamp"].isna().all():
        raise ValueError("Could not parse any timestamps from the CSV")

    ozone_col = OZONE_COLUMN if OZONE_COLUMN in df.columns else find_ozone_column(df.columns)
    if ozone_col is None:
        raise ValueError(f"Could not find ozone column '{OZONE_COLUMN}' in the CSV")

    columns = ["__timestamp", ozone_col]
    if NOX_COLUMN in df.columns:
        columns.append(NOX_COLUMN)

    df = df[columns].copy()
    df = df.rename(columns={ozone_col: "ozone", NOX_COLUMN: "nox"})
    if "nox" in df.columns:
        df["nox"] = pd.to_numeric(df["nox"], errors="coerce")

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
        })

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
    ax.set_title("Ozone time series")
    ax.set_xlabel("Time")
    ax.set_ylabel("Ozone")
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
        ax.plot(day_df["__timestamp"], day_df["ozone"], marker="o", linestyle="-", color="#1f77b4", label="Ozone")

        selected_date = datetime.fromisoformat(selected).date()
        excel_day = df_excel[df_excel["__date"] == selected_date]
        if not excel_day.empty:
            first_label = True
            injection_cols = [c for c in excel_day.columns if c.lower().endswith("_injectiontime")]
            for endtime in excel_day[ENDTIME_COLUMN].dropna().unique():
                ax.axvline(endtime, color="#ff7f0e", linestyle="--", linewidth=2,
                           label="Endtime_UTC" if first_label else None)
                first_label = False

                before = day_df[day_df["__timestamp"] < endtime].copy()
                before = before.dropna(subset=["ozone"])
                if not before.empty:
                    idxmax = before["ozone"].idxmax()
                    peak_row = before.loc[idxmax]
                    ax.scatter([peak_row["__timestamp"]], [peak_row["ozone"]], color="red", s=60, zorder=5,
                               label="Peak before Endtime" if first_label else None)
                    first_label = False

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

        ax.set_title(f"Ozone time series for {selected}")
        ax.set_xlabel("Time")
        ax.set_ylabel("Ozone")
        ax.grid(True, linestyle="--", alpha=0.4)
        if not excel_day.empty:
            ax.legend()
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
    notebook.add(tab2, text="OPR figures")
    create_tab2(tab2)

    tab3 = ttk.Frame(notebook)
    notebook.add(tab3, text="Tab 3")
    ttk.Label(tab3, text="Future tab content goes here.").pack(padx=16, pady=16)

    root.mainloop()


def create_tab2(frame):
    file_label = ttk.Label(frame, text=f"OH reactivity file: {OH_PATH}")
    file_label.pack(anchor="w", padx=10, pady=(10, 0))

    fig = Figure(figsize=(8, 4), dpi=100)
    ax = fig.add_subplot(111)
    ax.set_title("OPR figure: Ozone enhancement vs kOH_LP_sync_avg")
    ax.set_xlabel("kOH_LP_sync_avg (s-1)")
    ax.set_ylabel("Ozone enhancement (ppb)")
    fig.tight_layout()

    canvas = FigureCanvasTkAgg(fig, master=frame)
    canvas.get_tk_widget().pack(fill="both", expand=True, padx=10, pady=10)

    try:
        df = load_dataset(CSV_PATH)
        df_excel = load_excel_dataset(EXCEL_PATH)
        df_oh = load_oh_reactivity(OH_PATH)
        tab2_data = compute_tab2_data(df, df_excel, df_oh)
        if tab2_data.empty:
            raise ValueError("No valid Tab 2 data could be computed for the selected dates")
    except Exception as exc:
        ttk.Label(frame, text=str(exc), foreground="red").pack(padx=10, pady=10)
        return

    scatter = ax.scatter(
        tab2_data["kOH_LP_sync_avg"],
        tab2_data["ozone_increase"],
        c=tab2_data["nox_avg"],
        cmap="viridis",
        s=120,
        edgecolor="black",
        linewidth=0.5,
    )
    for _, row in tab2_data.iterrows():
        ax.annotate(row["date"].isoformat(), (row["kOH_LP_sync_avg"], row["ozone_increase"]),
                    textcoords="offset points", xytext=(5, 5), ha="left", fontsize=8)

    cbar = fig.colorbar(scatter, ax=ax)
    cbar.set_label("NOx avg (ppb)")
    ax.grid(True, linestyle="--", alpha=0.4)
    canvas.draw()


if __name__ == "__main__":
    build_gui()
