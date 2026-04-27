import sys
import pandas as pd

file_a = sys.argv[1]
file_b = sys.argv[2]

print(f"Reading {file_a}...")
df_a = pd.read_excel(file_a)

print(f"Reading {file_b}...")
df_b = pd.read_excel(file_b)

def profile_column(series):
    """Compute stats for a single column."""
    stats = {}

    #Info for all data types
    stats["dtype"] = str(series.dtype)
    stats["total_rows"] = len(series)
    stats["null_count"] = int(series.isna().sum())
    stats["blank_count"] = int((series.astype(str).str.strip() == "").sum()) - stats["null_count"]
    stats["null_pct"] = round(stats["null_count"]/stats["total_rows"] * 100,2) if stats["total_rows"] > 0 else 0
    stats["unique_values"] = int(series.nunique())

    #Additional info for numeric datatypes
    if pd.api.types.is_numeric_dtype(series):
        stats["min"] = series.min()
        stats["max"] = series.max()
        stats["mean"] = round(float(series.mean()), 4)
        stats["median"] = series.median()
        stats["std_dev"] = round(float(series.std()),4 ) if len(series.dropna()) > 1 else None

        #Outliers using quartiles
        q1 = series.quantile(0.25)
        q3 = series.quantile(0.75)
        iqr = q3 - q1
        lower = q1 - 1.5 * iqr
        upper = q3 + 1.5 * iqr
        stats["outlier_count"] = int(((series < lower) | (series > upper)).sum())
        stats["q1"] = q1
        stats["q3"] = q3
    else:
        stats ["min"] = None
        stats ["max"] = None
        stats ["mean"] = None
        stats ["median"] = None
        stats ["std_dev"] = None
        stats ["outlier_count"] = None
        stats ["q1"] = None
        stats ["q3"] = None

    #Most common value (any data type)
    if not series.dropna().empty:
        stats["most_common_value"] = series.value_counts().index[0]
        stats["most_common_count"] = int(series.value_counts().iloc[0])
    else:
        stats["most_common_value"] = None
        stats["most_common_count"] = 0

    return stats

def profile_dataframe(df, label):
    """Profiling each column."""
    results = []
    for col in df.columns:
        stats = profile_column(df[col])
        stats["source"] = label
        stats["column_name"] = col
        results.append(stats)
    return results

def build_summary(df_a, df_b, file_a, file_b):
    "Build a high-level comparison of both files"
    cols_a = set(df_a.columns)
    cols_b = set(df_b.columns)

    summary_data = {
        "Metric": [
            "File Name",
            "Row Count",
            "Column Count",
            "Total Nulls",
            "Total Null %",
            "Columns Only in This File",
            "Common Columns",
            "Numeric Columns",
            "Text/Other Columns",
        ],
        "File_A": [
            file_a,
            len(df_a),
            len(df_a.columns),
            int(df_a.isna().sum().sum()),
            round(float(df_a.isna().sum().sum()) / (df_a.shape[0] * df_a.shape[1]) * 100, 2) if df_a.size > 0 else 0,
            ", ".join(sorted(cols_a - cols_b)) or "None",
            len(cols_a & cols_b),
            sum(pd.api.types.is_numeric_dtype(df_a[c]) for c in df_a.columns),
            sum(not pd.api.types.is_numeric_dtype(df_a[c]) for c in df_a.columns),
        ],
        "File_B": [
            file_b,
            len(df_b),
            len(df_b.columns),
            int(df_b.isna().sum().sum()),
            round(float(df_b.isna().sum().sum()) / (df_b.shape[0] * df_b.shape[1]) * 100, 2) if df_b.size > 0 else 0,
            ", ".join(sorted(cols_b - cols_a)) or "None",
            len(cols_a & cols_b),
            sum(pd.api.types.is_numeric_dtype(df_b[c]) for c in df_b.columns),
            sum(not pd.api.types.is_numeric_dtype(df_b[c]) for c in df_b.columns),
        ],
    }
    return pd.DataFrame(summary_data)

print("Profiling File A...")
profile_a = profile_dataframe(df_a, "File_A")
print("Profiling File B...")
profile_b = profile_dataframe(df_b, "File_B")

all_profiles = pd.DataFrame(profile_a + profile_b)

# Reorder columns
col_order = ["source", "column_name", "dtype", "total_rows", "null_count",
               "blank_count", "null_pct", "unique_values", "min", "max",
               "mean", "median", "std_dev", "q1", "q3", "outlier_count",
               "most_common_value", "most_common_count"]
all_profiles = all_profiles[col_order]

# Build summary
summary_df = build_summary(df_a,df_b,file_a,file_b)

# Write to Excel
output_file = "comparison_output.xlsx"

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    summary_df.to_excel(writer, sheet_name="Summary", index=False)
    all_profiles.to_excel(writer, sheet_name="Column Detail", index=False)

print(f"\n Done! Output written to: {output_file}")
