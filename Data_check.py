import ozone_gui as g
import os

out_dir = r"Z:\VOC\Campaigns\2024 CHANEL\Ozone"
os.makedirs(out_dir, exist_ok=True)

df = g.load_dataset(g.CSV_PATH)
df_excel = g.load_excel_dataset(g.EXCEL_PATH)
tvoc_df = g.load_tvoc_dataset(g.TVOC_PATH)

merged = g.compute_tab2_10min_data(df, df_excel, tvoc_df)
print(merged.head(20))
print(merged.columns)
print(len(merged))

# Save the 10-minute resampled kOH time series before it is merged with ozone/NOx.
tvoc_10min = (
    tvoc_df.set_index("__timestamp")[['kOH']]
    .resample('10min')
    .mean()
    .dropna()
    .reset_index()
)
print(tvoc_10min.head(20))
print(len(tvoc_10min))

tvoc_10min.to_csv(os.path.join(out_dir, "tvoc_koh_10min.csv"), index=False)
#merged.to_csv(os.path.join(out_dir, "merged_koh_nox_ozone_10min.csv"), index=False)