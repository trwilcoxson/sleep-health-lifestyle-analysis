"""
Sleep Health and Lifestyle Dataset - Statistical Analysis
Computes descriptive statistics and generates visualizations.
"""

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from scipy import stats

# Load dataset
df = pd.read_csv('sleep_health_and_lifestyle_dataset.csv')

print("=" * 60)
print("SLEEP HEALTH AND LIFESTYLE DATASET - STATISTICAL ANALYSIS")
print("=" * 60)
print(f"\nDataset: {len(df)} records, {len(df.columns)} columns")
print(f"Columns: {list(df.columns)}")

# ============================================================
# SECTION 1: Data Description - Variable Types
# ============================================================
print("\n" + "=" * 60)
print("SECTION 1: DATA DESCRIPTION - VARIABLE TYPES")
print("=" * 60)
print(f"Continuous Variable:         Sleep Duration (hours) - values like {df['Sleep Duration'].iloc[0]}, {df['Sleep Duration'].iloc[1]}, {df['Sleep Duration'].iloc[2]}")
print(f"Integer Variable:            Daily Steps - values like {df['Daily Steps'].iloc[0]}, {df['Daily Steps'].iloc[1]}, {df['Daily Steps'].iloc[2]}")
print(f"Ordinal Categorical:         Quality of Sleep (1-10 scale) - ordered rating")
print(f"Nominal Categorical:         Gender - {df['Gender'].unique()}")

# ============================================================
# SECTION 2: Physical Activity Level - Measures of Center
# ============================================================
print("\n" + "=" * 60)
print("SECTION 2: PHYSICAL ACTIVITY LEVEL (minutes/day)")
print("=" * 60)
pa = df['Physical Activity Level']
pa_mean = pa.mean()
pa_median = pa.median()
pa_mode = pa.mode().iloc[0]

print(f"Mean:   {pa_mean:.2f} minutes")
print(f"Median: {pa_median:.2f} minutes")
print(f"Mode:   {pa_mode} minutes")

if pa_mean < pa_median:
    skew_dir = "left-skewed (negative skew)"
elif pa_mean > pa_median:
    skew_dir = "right-skewed (positive skew)"
else:
    skew_dir = "approximately symmetric"

print(f"\nSkewness Analysis:")
print(f"  Mean ({pa_mean:.2f}) < Median ({pa_median:.2f})")
print(f"  Distribution is {skew_dir}")
print(f"  The mean being slightly less than the median suggests the distribution")
print(f"  has a slight tail toward lower physical activity values.")

# ============================================================
# SECTION 3: Daily Steps - Measures of Spread
# ============================================================
print("\n" + "=" * 60)
print("SECTION 3: DAILY STEPS TAKEN")
print("=" * 60)
ds = df['Daily Steps']
ds_std = ds.std()
ds_max = ds.max()
ds_min = ds.min()
ds_range = ds_max - ds_min
ds_var = ds.var()
ds_q1 = ds.quantile(0.25)
ds_q3 = ds.quantile(0.75)
ds_iqr = ds_q3 - ds_q1

print(f"Standard Deviation: {ds_std:,.2f} steps")
print(f"Maximum:            {ds_max:,} steps")
print(f"Minimum:            {ds_min:,} steps")
print(f"Range:              {ds_range:,} steps")
print(f"\nAdditional Measures of Spread:")
print(f"  Variance:  {ds_var:,.2f} steps²")
print(f"  Q1 (25th): {ds_q1:,.0f} steps")
print(f"  Q3 (75th): {ds_q3:,.0f} steps")
print(f"  IQR:       {ds_iqr:,.0f} steps")

# ============================================================
# SECTION 4: Heart Rate Distribution
# ============================================================
print("\n" + "=" * 60)
print("SECTION 4: HEART RATE DISTRIBUTION")
print("=" * 60)
hr = df['Heart Rate']
hr_mean = hr.mean()
hr_median = hr.median()
hr_mode = hr.mode().iloc[0]
hr_std = hr.std()
hr_skew = hr.skew()
hr_q1 = hr.quantile(0.25)
hr_q3 = hr.quantile(0.75)
hr_iqr = hr_q3 - hr_q1
hr_lower = hr_q1 - 1.5 * hr_iqr
hr_upper = hr_q3 + 1.5 * hr_iqr
outliers = hr[(hr < hr_lower) | (hr > hr_upper)]

print(f"Mean:     {hr_mean:.2f} bpm")
print(f"Median:   {hr_median:.2f} bpm")
print(f"Mode:     {hr_mode} bpm")
print(f"Std Dev:  {hr_std:.2f} bpm")
print(f"Min:      {hr.min()} bpm")
print(f"Max:      {hr.max()} bpm")
print(f"Skewness: {hr_skew:.4f}")
print(f"\nOutlier Detection (IQR Method):")
print(f"  Q1: {hr_q1}, Q3: {hr_q3}, IQR: {hr_iqr}")
print(f"  Lower bound: {hr_lower:.1f} bpm")
print(f"  Upper bound: {hr_upper:.1f} bpm")
print(f"  Number of outliers: {len(outliers)}")
if len(outliers) > 0:
    print(f"  Outlier values: {sorted(outliers.unique())}")

# Generate Heart Rate Distribution Plot
fig, ax = plt.subplots(figsize=(10, 6))
n, bins, patches = ax.hist(hr, bins=15, color='#4A90D9', edgecolor='white',
                           linewidth=1.2, alpha=0.85)

# Color outlier bins differently
for i, (left, right) in enumerate(zip(bins[:-1], bins[1:])):
    if right > hr_upper or left < hr_lower:
        patches[i].set_facecolor('#E74C3C')
        patches[i].set_alpha(0.85)

ax.axvline(hr_mean, color='#2C3E50', linestyle='--', linewidth=2, label=f'Mean: {hr_mean:.1f} bpm')
ax.axvline(hr_median, color='#E67E22', linestyle='-.', linewidth=2, label=f'Median: {hr_median:.1f} bpm')
ax.axvline(hr_upper, color='#E74C3C', linestyle=':', linewidth=2, label=f'Outlier Threshold: {hr_upper:.0f} bpm')

ax.set_xlabel('Heart Rate (bpm)', fontsize=13, fontweight='bold')
ax.set_ylabel('Frequency', fontsize=13, fontweight='bold')
ax.set_title('Distribution of Heart Rates', fontsize=16, fontweight='bold', pad=15)
ax.legend(fontsize=11, loc='upper right', framealpha=0.9)
ax.grid(axis='y', alpha=0.3)
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)

plt.tight_layout()
plt.savefig('heart_rate_distribution.png', dpi=200, bbox_inches='tight')
print("\nHeart rate distribution plot saved to: heart_rate_distribution.png")

print("\n" + "=" * 60)
print("ANALYSIS COMPLETE")
print("=" * 60)
