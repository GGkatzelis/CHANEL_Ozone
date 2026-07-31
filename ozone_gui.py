import os
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime

import pandas as pd
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

CSV_PATH = r"Z:\VOC\Campaigns\2024 CHANEL\Ozone\CHANEL_DataUpload_AMSandTracegas_forTS.csv"
OZONE_COLUMN = "O3_TwoBTech_1014"


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

    df = df[["__timestamp", ozone_col]].copy()
    df = df.rename(columns={ozone_col: "ozone"})
    df = df.sort_values("__timestamp").reset_index(drop=True)
    df = df.dropna(subset=["__timestamp"])

    return df


def get_available_dates(df):
    df = df.copy()
    df["__day"] = df["__timestamp"].dt.date
    dates = sorted(df["__day"].unique())
    return [d.isoformat() for d in dates]


def filter_day(df, day_iso):
    day = datetime.fromisoformat(day_iso).date()
    return df[df["__timestamp"].dt.date == day].copy()


def create_tab1(frame):
    file_label = ttk.Label(frame, text=f"CSV file: {CSV_PATH}")
    file_label.pack(anchor="w", padx=10, pady=(10, 0))

    control_frame = ttk.Frame(frame)
    control_frame.pack(fill="x", padx=10, pady=10)

    ttk.Label(control_frame, text="Select date:").grid(row=0, column=0, sticky="w")
    date_var = tk.StringVar()
    date_combo = ttk.Combobox(control_frame, textvariable=date_var, state="readonly", width=24)
    date_combo.grid(row=0, column=1, padx=(8, 0), sticky="w")

    refresh_button = ttk.Button(control_frame, text="Plot day")
    refresh_button.grid(row=0, column=2, padx=(10, 0), sticky="w")

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
        ax.plot(day_df["__timestamp"], day_df["ozone"], marker="o", linestyle="-", color="#1f77b4")
        ax.set_title(f"Ozone time series for {selected}")
        ax.set_xlabel("Time")
        ax.set_ylabel("Ozone")
        ax.grid(True, linestyle="--", alpha=0.4)
        fig.autofmt_xdate(rotation=25)
        canvas.draw()

    refresh_button.config(command=draw_plot)
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
    notebook.add(tab2, text="Tab 2")
    ttk.Label(tab2, text="Future tab content goes here.").pack(padx=16, pady=16)

    tab3 = ttk.Frame(notebook)
    notebook.add(tab3, text="Tab 3")
    ttk.Label(tab3, text="Future tab content goes here.").pack(padx=16, pady=16)

    root.mainloop()


if __name__ == "__main__":
    build_gui()
