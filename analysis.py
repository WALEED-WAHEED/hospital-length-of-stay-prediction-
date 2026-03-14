# -*- coding: utf-8 -*-
"""
MN5812 Machine Learning & Predictive Analytics
NHS Hospital Length of Stay (LOS) Prediction
---------------------------------------------
Runs end-to-end: load raw data -> blend -> engineer features ->
EDA -> model comparison -> forecast current patients.

Target variable : LOS_Days (continuous, fractional days)
Best model found: Linear Regression  RMSE=0.591  R2=0.779
"""

import os
import warnings
warnings.filterwarnings("ignore")

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")   # no display needed -- writing PNGs to disk
import matplotlib.pyplot as plt
import seaborn as sns
from scipy import stats

from sklearn.model_selection import train_test_split, cross_val_score
from sklearn.preprocessing import StandardScaler
from sklearn.linear_model import LinearRegression, Ridge
from sklearn.tree import DecisionTreeRegressor
from sklearn.ensemble import RandomForestRegressor, GradientBoostingRegressor
from sklearn.metrics import mean_squared_error, mean_absolute_error, r2_score
import joblib

# output folders / constants
CHARTS_DIR   = "charts"
os.makedirs(CHARTS_DIR, exist_ok=True)

RANDOM_STATE = 42
TEST_SIZE    = 0.20
TODAY        = pd.Timestamp("2026-03-12")

# NHS colour scheme for charts
NHS_BLUE  = "#005EB8"
NHS_DARK  = "#003087"
NHS_LIGHT = "#41B6E6"
PALETTE   = [NHS_BLUE, NHS_LIGHT, "#AE2573", "#007F3B", "#ED8B00", "#768692"]

print("=" * 70)
print("MN5812  |  Hospital Length of Stay Prediction")
print("=" * 70)


# =========================================================================
# 1. LOAD DATA
# =========================================================================
print("\n[1/6] Reading source files...")

patient_info  = pd.read_excel("Patient Information.xlsx")
surgical_info = pd.read_excel("Surgical Information.xlsx")
current_pts   = pd.read_excel("Current Patients.xlsx")
icd10_codes   = pd.read_excel("ICD-10 Codes.xlsx")

print(f"  Patient Information  : {patient_info.shape[0]:>5} rows")
print(f"  Surgical Information : {surgical_info.shape[0]:>5} rows")
print(f"  Current Patients     : {current_pts.shape[0]:>5} rows")
print(f"  ICD-10 Codes         : {icd10_codes.shape[0]:>5} rows")


# =========================================================================
# 2. MERGE DATASETS
# =========================================================================
print("\n[2/6] Joining tables...")

# Inner join keeps only patients who actually had surgery. Ten patients
# in the admin file have no matching surgery record -- they fall out here,
# which is expected and fine.
df = surgical_info.merge(patient_info, on="Patient ID", how="inner")
print(f"  Surgical JOIN Patient: {df.shape[0]} rows")

# Left join so all 982 surgical records survive even if a code were somehow
# missing from the lookup (none are, but safer than inner).
df = df.merge(icd10_codes, on="ICD-10 Code", how="left")
print(f"  After ICD-10 lookup : {df.shape[0]} rows")


# =========================================================================
# 3. FEATURE ENGINEERING
# =========================================================================
print("\n[3/6] Building features...")

# Date of Birth comes in as a plain string ('1/1/1945') from Excel.
# Everything else is already datetime64 so we just coerce to be safe.
df["Date of Birth"] = pd.to_datetime(df["Date of Birth"], dayfirst=False)
for col in ["Hospital Admission Date", "Surgery End Datetime",
            "First Ambulation", "Hospital Release"]:
    df[col] = pd.to_datetime(df[col])

# --- TARGET: LOS in fractional days ---
# Using total seconds / 86400 rather than .days so we keep sub-day
# precision from the timestamps.
df["LOS_Days"] = (
    (df["Hospital Release"] - df["Hospital Admission Date"])
    .dt.total_seconds() / 86400
)

# --- Age at admission (whole years) ---
df["Age"] = (df["Hospital Admission Date"] - df["Date of Birth"]).dt.days // 365

# --- Hours to first ambulation (turned out to be the dominant feature) ---
# This is the gap between surgery end and when the patient first walked.
# Clinically, earlier mobilisation = shorter stay, which the data confirms
# very strongly (r = 0.785).
df["Hours_till_Ambulation"] = (
    (df["First Ambulation"] - df["Surgery End Datetime"])
    .dt.total_seconds() / 3600
)

# Surgery_Type: first word of the ICD-10 description
# e.g. "Removal of Autologous Tissue Substitute..." -> "Removal"
# Happens to match the column in Current Patients exactly, so no manual map needed.
df["Surgery_Type"] = df["Description"].str.split().str[0]

# Month and day-of-week -- might capture seasonal/staffing effects on discharge
df["Admission_Month"]     = df["Hospital Admission Date"].dt.month
df["Admission_DayOfWeek"] = df["Hospital Admission Date"].dt.dayofweek  # 0=Mon

# Age bands -- only for EDA grouping, not in the model
df["Age_Band"] = pd.cut(
    df["Age"],
    bins=[0, 40, 55, 65, 75, 120],
    labels=["<40", "40-54", "55-64", "65-74", "75+"]
)

# --- Data quality checks ---
n_dupes = df.duplicated(subset="Patient ID").sum()
n_nulls = df.isnull().sum().sum()
print(f"  Duplicates on Patient ID : {n_dupes}")
print(f"  Total null cells          : {n_nulls}")
# Both are 0 -- no imputation needed in training data.

# Flag IQR outliers in LOS but don't drop them.
# The one record flagged (7.41 days) is plausible for complex spinal work.
Q1, Q3 = df["LOS_Days"].quantile([0.25, 0.75])
IQR = Q3 - Q1
df["LOS_Outlier"] = (
    (df["LOS_Days"] < Q1 - 1.5 * IQR) | (df["LOS_Days"] > Q3 + 1.5 * IQR)
)
print(f"  IQR outliers (flagged, not removed): {df['LOS_Outlier'].sum()}")

# --- Encoding ---
# One-hot rather than label encode because neither Gender nor Surgery_Type
# has a meaningful order. drop_first avoids the dummy trap in linear models.
df_encoded = pd.get_dummies(df, columns=["Gender", "Surgery_Type"], drop_first=True)

FEATURE_COLS = (
    ["Age", "Hours_till_Ambulation", "Admission_Month", "Admission_DayOfWeek"]
    + [c for c in df_encoded.columns
       if c.startswith("Gender_") or c.startswith("Surgery_Type_")]
)
TARGET_COL = "LOS_Days"

print(f"  Features: {FEATURE_COLS}")

# Standardise numerics so linear models aren't scale-sensitive.
# Trees don't care but we use the same feature matrix for all models.
scaler = StandardScaler()
df_encoded[["Age_scaled", "Hours_scaled", "Month_scaled", "DayOfWeek_scaled"]] = (
    scaler.fit_transform(
        df_encoded[["Age", "Hours_till_Ambulation",
                    "Admission_Month", "Admission_DayOfWeek"]]
    )
)

X = df_encoded[FEATURE_COLS].astype(float)
y = df_encoded[TARGET_COL]

print(f"\n  Training dataset: {X.shape[0]} rows x {X.shape[1]} features")
print(f"  LOS: mean={y.mean():.2f}d  std={y.std():.2f}  "
      f"range=[{y.min():.2f}, {y.max():.2f}]")


# =========================================================================
# 4. EXPLORATORY DATA ANALYSIS
# =========================================================================
print("\n[4/6] EDA...")

# Summary stats
print("\n  Descriptive stats:")
num_cols = ["LOS_Days", "Age", "Hours_till_Ambulation",
            "Admission_Month", "Admission_DayOfWeek"]
print(df[num_cols].describe().round(3).to_string())

skewness = df["LOS_Days"].skew()
print(f"\n  LOS skewness: {skewness:.4f}  "
      f"({'right' if skewness > 0 else 'left'}-skewed)")

# Chart 1: LOS distribution + Q-Q
fig, axes = plt.subplots(1, 2, figsize=(12, 5))
fig.suptitle("Length of Stay -- Distribution", fontsize=14,
             fontweight="bold", color=NHS_DARK)

axes[0].hist(df["LOS_Days"], bins=40, color=NHS_BLUE,
             edgecolor="white", alpha=0.85)
axes[0].axvline(df["LOS_Days"].mean(), color="red", linestyle="--",
                label=f"Mean {df['LOS_Days'].mean():.2f}d")
axes[0].axvline(df["LOS_Days"].median(), color="orange", linestyle="--",
                label=f"Median {df['LOS_Days'].median():.2f}d")
axes[0].set_xlabel("LOS (days)"); axes[0].set_ylabel("Frequency")
axes[0].set_title("Histogram"); axes[0].legend()

stats.probplot(df["LOS_Days"], dist="norm", plot=axes[1])
axes[1].set_title("Q-Q Plot")
axes[1].get_lines()[0].set(color=NHS_BLUE, markersize=3)
axes[1].get_lines()[1].set(color="red")

plt.tight_layout()
plt.savefig(f"{CHARTS_DIR}/01_LOS_distribution.png", dpi=150, bbox_inches="tight")
plt.close()
print("  Saved: 01_LOS_distribution.png")

# Chart 2: Correlation heatmap
corr = df[num_cols].corr()
fig, ax = plt.subplots(figsize=(7, 6))
sns.heatmap(corr, annot=True, fmt=".2f", cmap="Blues",
            linewidths=0.5, ax=ax, annot_kws={"size": 10})
ax.set_title("Pearson Correlations -- Numeric Features",
             fontweight="bold", color=NHS_DARK)
plt.tight_layout()
plt.savefig(f"{CHARTS_DIR}/02_correlation_heatmap.png", dpi=150, bbox_inches="tight")
plt.close()
print("  Saved: 02_correlation_heatmap.png")

print("\n  Correlations with LOS_Days:")
print(corr["LOS_Days"].drop("LOS_Days").sort_values(ascending=False).to_string())

# Chart 3: Box plots by surgery type, age band, gender
fig, axes = plt.subplots(1, 3, figsize=(16, 6))
fig.suptitle("LOS by Patient / Procedure Category",
             fontsize=14, fontweight="bold", color=NHS_DARK)

order_surg = (df.groupby("Surgery_Type")["LOS_Days"]
              .median().sort_values(ascending=False).index)
sns.boxplot(data=df, x="Surgery_Type", y="LOS_Days",
            order=order_surg, palette=PALETTE, ax=axes[0])
axes[0].set_title("Surgery Type"); axes[0].set_xlabel("")
axes[0].tick_params(axis="x", rotation=30)

sns.boxplot(data=df, x="Age_Band", y="LOS_Days",
            palette=PALETTE, ax=axes[1])
axes[1].set_title("Age Band"); axes[1].set_xlabel("")

sns.boxplot(data=df, x="Gender", y="LOS_Days",
            palette=[NHS_BLUE, NHS_LIGHT], ax=axes[2])
axes[2].set_title("Gender"); axes[2].set_xlabel("")

for ax in axes:
    ax.set_ylabel("LOS (days)")

plt.tight_layout()
plt.savefig(f"{CHARTS_DIR}/03_LOS_boxplots.png", dpi=150, bbox_inches="tight")
plt.close()
print("  Saved: 03_LOS_boxplots.png")

# Chart 4: LOS by ICD-10 code
fig, ax = plt.subplots(figsize=(14, 6))
order_icd = (df.groupby("Description")["LOS_Days"]
             .median().sort_values(ascending=False).index)
short_labels = [d[:45] + "..." if len(d) > 45 else d for d in order_icd]

bp = ax.boxplot(
    [df[df["Description"] == d]["LOS_Days"].values for d in order_icd],
    labels=short_labels, patch_artist=True, notch=False
)
for patch, col in zip(bp["boxes"],
                      plt.cm.Blues(np.linspace(0.4, 0.9, len(order_icd)))):
    patch.set_facecolor(col)
ax.set_xticklabels(short_labels, rotation=45, ha="right", fontsize=7)
ax.set_ylabel("LOS (days)")
ax.set_title("LOS by ICD-10 Procedure", fontweight="bold", color=NHS_DARK)
plt.tight_layout()
plt.savefig(f"{CHARTS_DIR}/04_LOS_by_ICD10.png", dpi=150, bbox_inches="tight")
plt.close()
print("  Saved: 04_LOS_by_ICD10.png")

# Chart 5: Procedure frequency
fig, axes = plt.subplots(1, 2, figsize=(16, 6))
fig.suptitle("Procedure Counts", fontsize=14, fontweight="bold", color=NHS_DARK)

dx_counts = df["Surgery_Type"].value_counts()
axes[0].barh(dx_counts.index, dx_counts.values, color=NHS_BLUE)
axes[0].set_xlabel("Count"); axes[0].set_title("Surgery Type")
for i, v in enumerate(dx_counts.values):
    axes[0].text(v + 2, i, str(v), va="center", fontsize=9)

icd_counts = df["Description"].str[:40].value_counts().head(10)
axes[1].barh(range(len(icd_counts)), icd_counts.values, color=NHS_LIGHT)
axes[1].set_yticks(range(len(icd_counts)))
axes[1].set_yticklabels(icd_counts.index, fontsize=8)
axes[1].set_xlabel("Count"); axes[1].set_title("Top 10 ICD-10 Procedures")

plt.tight_layout()
plt.savefig(f"{CHARTS_DIR}/05_procedure_frequency.png", dpi=150, bbox_inches="tight")
plt.close()
print("  Saved: 05_procedure_frequency.png")

# Chart 6: Hours-to-ambulation vs LOS scatter (the money chart)
fig, ax = plt.subplots(figsize=(8, 5))
sc = ax.scatter(df["Hours_till_Ambulation"], df["LOS_Days"],
                c=df["Age"], cmap="Blues", alpha=0.6, edgecolors="none", s=20)
plt.colorbar(sc, ax=ax, label="Age (years)")

m, b_coef = np.polyfit(df["Hours_till_Ambulation"], df["LOS_Days"], 1)
xs = np.linspace(df["Hours_till_Ambulation"].min(),
                 df["Hours_till_Ambulation"].max(), 100)
r_val = df["Hours_till_Ambulation"].corr(df["LOS_Days"])
ax.plot(xs, m * xs + b_coef, "r--", linewidth=1.5,
        label=f"Fit (r={r_val:.2f})")
ax.set_xlabel("Hours till First Ambulation")
ax.set_ylabel("LOS (days)")
ax.set_title("Ambulation Time vs LOS", fontweight="bold", color=NHS_DARK)
ax.legend()
plt.tight_layout()
plt.savefig(f"{CHARTS_DIR}/06_ambulation_vs_LOS.png", dpi=150, bbox_inches="tight")
plt.close()
print("  Saved: 06_ambulation_vs_LOS.png")


# =========================================================================
# 5. MODEL TRAINING & EVALUATION
# =========================================================================
print("\n[5/6] Training models...")

# 80/20 split with a fixed seed so results are reproducible
X_train, X_test, y_train, y_test = train_test_split(
    X, y, test_size=TEST_SIZE, random_state=RANDOM_STATE
)
print(f"  Train: {X_train.shape[0]} rows  |  Test: {X_test.shape[0]} rows")

models = {
    "Linear Regression": LinearRegression(),
    "Ridge Regression":  Ridge(alpha=1.0, random_state=RANDOM_STATE),
    "Decision Tree":     DecisionTreeRegressor(max_depth=6, random_state=RANDOM_STATE),
    "Random Forest":     RandomForestRegressor(n_estimators=200, max_depth=10,
                                               random_state=RANDOM_STATE, n_jobs=-1),
    "Gradient Boosting (sklearn)": GradientBoostingRegressor(
                                       n_estimators=200, learning_rate=0.05,
                                       max_depth=4, random_state=RANDOM_STATE),
}

results = {}
for name, model in models.items():
    model.fit(X_train, y_train)
    preds = model.predict(X_test)

    rmse  = np.sqrt(mean_squared_error(y_test, preds))
    mae   = mean_absolute_error(y_test, preds)
    r2    = r2_score(y_test, preds)
    cv_r2 = cross_val_score(model, X, y, cv=5, scoring="r2")  # 5-fold on full data

    results[name] = {
        "Model":      name,
        "RMSE":       round(rmse, 4),
        "MAE":        round(mae,  4),
        "R2":         round(r2,   4),
        "CV_R2_mean": round(cv_r2.mean(), 4),
        "CV_R2_std":  round(cv_r2.std(),  4),
    }
    print(f"  {name:<34} RMSE={rmse:.4f}  R2={r2:.4f}  "
          f"CV_R2={cv_r2.mean():.4f}+/-{cv_r2.std():.4f}")

results_df = pd.DataFrame(list(results.values())).set_index("Model")
print("\n  Model comparison:")
print(results_df.to_string())

# Pick winner by rank sum: best (lowest) RMSE + best (highest) CV R2.
# Using both guards against a model that just got lucky on the test split.
results_df["Rank"] = (
    results_df["RMSE"].rank(ascending=True)
    + results_df["CV_R2_mean"].rank(ascending=False)
)
best_model_name = results_df["Rank"].idxmin()
best_model      = models[best_model_name]
best_metrics    = results[best_model_name]

print(f"\n  Winner: {best_model_name}")
print(f"  RMSE={best_metrics['RMSE']}  MAE={best_metrics['MAE']}  "
      f"R2={best_metrics['R2']}  CV_R2={best_metrics['CV_R2_mean']}+/-{best_metrics['CV_R2_std']}")

joblib.dump(best_model, "best_model.pkl")
print("  Saved: best_model.pkl")

# Feature importances: for tree models this comes for free; for linear
# models we use absolute coefficient values as a rough proxy.
y_pred_best = best_model.predict(X_test)

if hasattr(best_model, "feature_importances_"):
    importances = best_model.feature_importances_
else:
    importances = np.abs(best_model.coef_)

feat_imp = pd.Series(importances, index=FEATURE_COLS).sort_values(ascending=False)
print("\n  Top 5 predictors:")
print(feat_imp.head(5).to_string())

# Chart 7: Feature importances
fig, ax = plt.subplots(figsize=(9, 5))
colors = [NHS_BLUE if i < 5 else "#768692" for i in range(len(feat_imp))]
feat_imp.plot(kind="barh", ax=ax, color=colors[::-1])
ax.set_xlabel("Importance")
ax.set_title(f"Feature Importances  ({best_model_name})",
             fontweight="bold", color=NHS_DARK)
ax.invert_yaxis()
plt.tight_layout()
plt.savefig(f"{CHARTS_DIR}/07_feature_importances.png", dpi=150, bbox_inches="tight")
plt.close()
print("  Saved: 07_feature_importances.png")

# Chart 8: Actual vs predicted
fig, ax = plt.subplots(figsize=(7, 7))
ax.scatter(y_test, y_pred_best, alpha=0.5, color=NHS_BLUE, s=20, edgecolors="none")
mn = min(y_test.min(), y_pred_best.min())
mx = max(y_test.max(), y_pred_best.max())
ax.plot([mn, mx], [mn, mx], "r--", linewidth=1.5, label="Perfect")
ax.set_xlabel("Actual LOS (days)")
ax.set_ylabel("Predicted LOS (days)")
ax.set_title(f"Actual vs Predicted  ({best_model_name})",
             fontweight="bold", color=NHS_DARK)
ax.legend()
ax.text(0.05, 0.92, f"R2 = {best_metrics['R2']}",
        transform=ax.transAxes, fontsize=11, color=NHS_DARK)
plt.tight_layout()
plt.savefig(f"{CHARTS_DIR}/08_actual_vs_predicted.png", dpi=150, bbox_inches="tight")
plt.close()
print("  Saved: 08_actual_vs_predicted.png")

# Chart 9: Residuals
residuals = y_test - y_pred_best
fig, axes = plt.subplots(1, 2, figsize=(12, 5))
axes[0].scatter(y_pred_best, residuals, alpha=0.5,
                color=NHS_BLUE, s=15, edgecolors="none")
axes[0].axhline(0, color="red", linestyle="--")
axes[0].set_xlabel("Predicted LOS")
axes[0].set_ylabel("Residual")
axes[0].set_title("Residuals vs Fitted")

axes[1].hist(residuals, bins=35, color=NHS_LIGHT, edgecolor="white")
axes[1].axvline(0, color="red", linestyle="--")
axes[1].set_xlabel("Residual"); axes[1].set_ylabel("Frequency")
axes[1].set_title("Residual Distribution")
axes[1].text(0.65, 0.90, f"Skew: {residuals.skew():.3f}",
             transform=axes[1].transAxes, fontsize=9, color=NHS_DARK)

fig.suptitle(f"Residuals  ({best_model_name})", fontsize=13,
             fontweight="bold", color=NHS_DARK)
plt.tight_layout()
plt.savefig(f"{CHARTS_DIR}/09_residuals.png", dpi=150, bbox_inches="tight")
plt.close()
print("  Saved: 09_residuals.png")

# Chart 10: Side-by-side metric comparison across all models
fig, axes = plt.subplots(1, 3, figsize=(15, 5))
fig.suptitle("Model Comparison", fontsize=14, fontweight="bold", color=NHS_DARK)

for i, (metric, label) in enumerate(
    [("RMSE", "RMSE (days)"), ("MAE", "MAE (days)"), ("CV_R2_mean", "CV R2")]
):
    vals = results_df[metric]
    bars = axes[i].bar(range(len(vals)), vals, color=PALETTE[:len(vals)])
    axes[i].set_xticks(range(len(vals)))
    axes[i].set_xticklabels(vals.index, rotation=30, ha="right", fontsize=8)
    axes[i].set_ylabel(label); axes[i].set_title(metric)
    for bar in bars:
        h = bar.get_height()
        axes[i].text(bar.get_x() + bar.get_width() / 2, h + 0.002,
                     f"{h:.3f}", ha="center", va="bottom", fontsize=8)

plt.tight_layout()
plt.savefig(f"{CHARTS_DIR}/10_model_comparison.png", dpi=150, bbox_inches="tight")
plt.close()
print("  Saved: 10_model_comparison.png")


# =========================================================================
# 6. FORECAST CURRENT PATIENTS
# =========================================================================
print("\n[6/6] Forecasting current patients...")

cp = current_pts.copy()
print(f"  Patients to score: {len(cp)}")

# Column names in Current Patients have spaces; rename to match training schema
cp.rename(columns={
    "Surgery Type":       "Surgery_Type",
    "Hours till Ambulation": "Hours_till_Ambulation",
}, inplace=True)

# Admission timing is unknown for current patients, so we impute with
# training set mode. Not ideal but documented and consistent with the report.
cp["Admission_Month"]     = int(df["Admission_Month"].mode()[0])
cp["Admission_DayOfWeek"] = int(df["Admission_DayOfWeek"].mode()[0])

cp_enc = pd.get_dummies(cp, columns=["Gender", "Surgery_Type"], drop_first=True)

# Any dummy column present in training but absent here (e.g. a gender category
# not in Current Patients) needs to be added as all-zeros.
for col in FEATURE_COLS:
    if col not in cp_enc.columns:
        cp_enc[col] = 0

X_current   = cp_enc[FEATURE_COLS].astype(float)
predictions = np.clip(best_model.predict(X_current), 0, None)

# Build the output table
output = cp[["Gender", "Age", "Surgery_Type", "Hours_till_Ambulation"]].copy()
output.insert(0, "Patient_ID", [f"CUR-{i+1:03d}" for i in range(len(cp))])
output["Predicted_LOS_Days"] = np.round(predictions, 2)
output["Predicted_Discharge_Date"] = [
    (TODAY + pd.Timedelta(days=float(p))).strftime("%Y-%m-%d")
    for p in predictions
]
output["Risk_Category"] = pd.cut(
    predictions,
    bins=[-np.inf, 3, 7, np.inf],
    labels=["Short (<3 days)", "Medium (3-7 days)", "Long (>7 days)"]
)

print("\n  Predictions:")
print(output.to_string(index=False))

output.to_excel("predictions_output.xlsx", index=False)
print("\n  Saved: predictions_output.xlsx")


# =========================================================================
# SUMMARY
# =========================================================================
avg_los   = output["Predicted_LOS_Days"].mean()
n_current = len(output)

print("\n" + "=" * 70)
print("SUMMARY")
print("=" * 70)
print(f"  Best model        : {best_model_name}")
print(f"  RMSE              : {best_metrics['RMSE']} days")
print(f"  R2                : {best_metrics['R2']}")
print(f"  CV R2 (5-fold)    : {best_metrics['CV_R2_mean']} +/- {best_metrics['CV_R2_std']}")
print(f"  Current patients  : {n_current}")
print(f"  Avg predicted LOS : {avg_los:.2f} days")
print(f"  Charts            : {len(os.listdir(CHARTS_DIR))}")
print("=" * 70)

# expose for external use if this module is imported
ANALYSIS_RESULTS = {
    "best_model_name": best_model_name,
    "best_rmse":       best_metrics["RMSE"],
    "best_mae":        best_metrics["MAE"],
    "best_r2":         best_metrics["R2"],
    "best_cv_r2":      best_metrics["CV_R2_mean"],
    "n_current":       n_current,
    "avg_los":         round(avg_los, 2),
    "top_features":    feat_imp.head(5).index.tolist(),
    "output_df":       output,
}
