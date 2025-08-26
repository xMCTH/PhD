"""
Excel-Based Voxel Heatmap Visualizer
------------------------------------

This script loads voxel data from an Excel workbook (sheet `"all"`) and generates 
slice-wise heatmaps of metabolite intensities across the XY plane for each z-layer.  

Main Features:
1. Interactive file selection:
   - Prompts user to select an Excel file (`.xlsx` or `.xls`).
   - Reads the `"all"` sheet, using **row 4** (index=3) as the header row 
     (i.e. data starts at row 5 in Excel).

2. Coordinate parsing:
   - Identifies the voxel coordinate column (e.g., `"Coord_p"`, `"Coord"`).
   - Splits into integer `x`, `y`, `z` values for spatial mapping.

3. Intensity source selection:
   - Detects candidate columns containing `"Height"` and `"Area"`.
   - User chooses which intensity metric to use:
       • Height  
       • Area  
       • Default (Height if available, otherwise Area)

4. Color scaling:
   - User is prompted for `vmin` and `vmax` (default: 0–10).
   - Actual min/max of the data are computed.
   - If user values exceed actual data ranges, they are automatically adjusted 
     to “nice” rounded limits using custom rounding helpers.
   - Out-of-range values and NaNs are masked and displayed as **white**.

5. Heatmap visualization:
   - Loops over all unique z-slices.
   - Builds a full XY grid per slice (ensuring gaps are filled with NaN).
   - Displays 2D heatmaps using Seaborn with consistent colormap scaling.
   - Each slice shows only values within `[vmin, vmax]`; others appear white.

6. Output:
   - Heatmaps are displayed interactively with colorbars labeled by intensity source.
   - Each slice is shown in a separate figure.

Dependencies:
- Python standard library (`sys`, `math`, `tkinter`)
- Third-party libraries:
  • pandas  
  • numpy  
  • matplotlib  
  • seaborn  

Usage:
- Run the script in Python.
- Select an Excel file with an `"all"` sheet structured as exported by the 
  metabolite-processing workflow.
- Choose color scaling and intensity source when prompted.
- Slice-wise heatmaps will be displayed for inspection.

"""


import sys
import math
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from matplotlib.colors import ListedColormap
import tkinter as tk
from tkinter import filedialog

# --- Select Excel file ---
root = tk.Tk()
root.withdraw()  # Hide main window
excel_file = filedialog.askopenfilename(
    title="Select Excel File",
    filetypes=[("Excel Files", "*.xlsx;*.xls")]
)
if not excel_file:
    print("No Excel file selected.")
    sys.exit()

# --- User-defined settings ---
try:
    vmin = input("Enter minimum intensity value (e.g. 0) [default 0]: ").strip()
    vmin = float(vmin) if vmin != "" else 0.0
except ValueError:
    print("Invalid minimum value entered; using 0.")
    vmin = 0.0

try:
    vmax = input("Enter maximum intensity value (e.g. 10) [default 10]: ").strip()
    vmax = float(vmax) if vmax != "" else 10.0
except ValueError:
    print("Invalid maximum value entered; using 10.")
    vmax = 10.0
# -----------------------------

# Read the "all" sheet using row 4 (index 3) as the header (so data starts on row 5)
df = pd.read_excel(excel_file, sheet_name='all', header=3)

# Clean column names
df.columns = df.columns.str.strip()

print("Loaded columns:", df.columns.tolist())

# Find the coordinate column (e.g. 'Coord_p' or any column containing 'coord')
coord_cols = [c for c in df.columns if 'coord' in c.lower()]
if not coord_cols:
    raise RuntimeError("No column containing 'Coord' found in 'all' sheet headers.")
coord_col = coord_cols[0]
print(f"Using coordinate column: '{coord_col}'")

# Split coordinate into x,y,z (strip spaces)
coords = df[coord_col].astype(str).str.strip().str.split('_', expand=True)
if coords.shape[1] < 3:
    raise RuntimeError(f"Coordinate values in '{coord_col}' do not split into 3 parts.")
df['x'] = coords[0].astype(int)
df['y'] = coords[1].astype(int)
df['z'] = coords[2].astype(int)

# Detect candidate columns
height_cols = [c for c in df.columns if 'height' in c.lower()]
area_cols = [c for c in df.columns if c.lower() == 'area' or 'area' in c.lower()]

print("\nDetected candidate columns:")
print(f"  Height-like columns: {height_cols}")
print(f"  Area-like columns:   {area_cols}")

# Prompt user to choose which source to use
prompt = (
    "\nChoose intensity source:\n"
    "  [H]eight  - use Height column (if available)\n"
    "  [A]rea    - use Area column (if available)\n"
    "  [D]efault - prefer Height if present, otherwise Area\n"
    "Type H, A, or D (default D): "
)
choice = input(prompt).strip().lower() or 'd'

int_col = None
if choice.startswith('h'):
    if height_cols:
        int_col = height_cols[0]
    elif area_cols:
        print("Height not found — falling back to Area.")
        int_col = area_cols[0]
    else:
        raise RuntimeError("Neither Height nor Area found in sheet.")
elif choice.startswith('a'):
    if area_cols:
        int_col = area_cols[0]
    elif height_cols:
        print("Area not found — falling back to Height.")
        int_col = height_cols[0]
    else:
        raise RuntimeError("Neither Area nor Height found in sheet.")
else:  # default preference: Height then Area
    if height_cols:
        int_col = height_cols[0]
    elif area_cols:
        int_col = area_cols[0]
    else:
        raise RuntimeError("Neither Height nor Area found in sheet.")

print(f"Using intensity column: '{int_col}'\n")

# Prepare intensity column (numeric)
df = df.copy()
df['intensity'] = pd.to_numeric(df[int_col], errors='coerce')

# Compute actual min/max (ignoring NaNs)
has_values = df['intensity'].notna().any()
if has_values:
    actual_min = float(np.nanmin(df['intensity'].values))
    actual_max = float(np.nanmax(df['intensity'].values))
else:
    actual_min = 0.0
    actual_max = 0.0

def round_up_nice(x):
    """Round x up to a 'nice' number (1 significant digit: e.g. 123->200, 78->80, 0.023->0.03)."""
    if x == 0:
        return 0.0
    sign = 1 if x > 0 else -1
    ax = abs(x)
    exp = math.floor(math.log10(ax))
    mag = 10 ** exp
    return sign * (math.ceil(ax / mag) * mag)

def round_down_nice(x):
    """Round x down to a 'nice' number (1 significant digit: e.g. 123->100, 78->70, 0.023->0.02)."""
    if x == 0:
        return 0.0
    sign = 1 if x > 0 else -1
    ax = abs(x)
    exp = math.floor(math.log10(ax))
    mag = 10 ** exp
    return sign * (math.floor(ax / mag) * mag)

# Adjust vmax if user set it larger than actual_max
if has_values and vmax > actual_max:
    new_vmax = round_up_nice(actual_max)
    # ensure new_vmax >= actual_max (it should by construction); handle edge-case when round_up returns 0
    if new_vmax < actual_max:
        new_vmax = actual_max
    print(f"User vmax ({vmax}) > actual max ({actual_max}). Adjusting colorbar vmax to {new_vmax}.")
    vmax = new_vmax

# Adjust vmin if user set it smaller than actual_min
if has_values and vmin < actual_min:
    new_vmin = round_down_nice(actual_min)
    # If rounding down produced the same value as actual_min (or incorrectly 0), allow small adjustments:
    # ensure new_vmin <= actual_min
    if new_vmin > actual_min:
        new_vmin = actual_min
    print(f"User vmin ({vmin}) < actual min ({actual_min}). Adjusting colorbar vmin to {new_vmin}.")
    vmin = new_vmin

# Safety: if vmin > vmax after adjustments, swap them and warn
if vmin > vmax:
    print(f"Warning: vmin ({vmin}) > vmax ({vmax}) after adjustments. Swapping.")
    vmin, vmax = vmax, vmin

print(f"Final color limits: vmin={vmin}, vmax={vmax}")

# Get unique z slices
z_slices = sorted(df['z'].unique())

# For consistent axis coverage across slices, compute full x,y ranges now
xs = np.arange(df['x'].min(), df['x'].max() + 1)
ys = np.arange(df['y'].min(), df['y'].max() + 1)

# Define colormap and set NaN/masked values to white
cmap = plt.get_cmap('viridis')
new_colors = cmap(np.linspace(0, 1, 256))
new_cmap = ListedColormap(new_colors)
white = (1.0, 1.0, 1.0, 1.0)
# set_bad is supported for ListedColormap objects; NaN/masked will show as white
new_cmap.set_bad(white)

# Plot each z slice as heatmap
for z in z_slices:
    slice_df = df[df['z'] == z]

    # Pivot to 2D grid (rows = y, cols = x)
    intensity_grid = slice_df.pivot(index='y', columns='x', values='intensity')

    # Reindex to ensure the full grid and sorted ascending axes (so row 0 is top when not inverted)
    intensity_grid = intensity_grid.reindex(index=ys, columns=xs).sort_index(axis=0).sort_index(axis=1)

    # Ensure numeric
    intensity_grid = intensity_grid.astype(float)

    # Mask values outside chosen vmin/vmax and also the NaNs
    out_of_range_mask = (intensity_grid < vmin) | (intensity_grid > vmax) | intensity_grid.isna()
    # Make a copy where out-of-range are set to NaN so they render as white
    plot_data = intensity_grid.mask(out_of_range_mask)

    plt.figure(figsize=(8, 6))
    sns.heatmap(
        plot_data,
        cmap=new_cmap,
        cbar_kws={'label': f'{int_col}'},
        linewidths=0.5,
        vmin=vmin,
        vmax=vmax,
        mask=plot_data.isna(),
        square=False
    )
    plt.title(f'{int_col} Map at z = {z} (values outside [{vmin}, {vmax}] shown white)')
    plt.xlabel('x')
    plt.ylabel('y')
    # NOTE: Removed plt.gca().invert_yaxis() so that y=0 is displayed at the top-left (row 0 at top)
    plt.tight_layout()
    plt.show()
