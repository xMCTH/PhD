"""
Excel-based Intensity Heatmap Visualization
-------------------------------------------

This script loads voxel-wise concentration levels from an Excel file (produced by 
the Phantom vs. Subject Processing Script or a similar pipeline) and generates heatmaps 
for each z-slice.

Workflow:
1. Prompts the user to select an Excel file (`.xlsx` or `.xls`).
   - Expects an "all" worksheet with headers starting on row 4 (index 3).
2. Reads voxel coordinates from the `Coord_p` column (formatted as `x_y_z`) and extracts x, y, z.
3. Uses the `c [mM]` column as intensity (renamed internally to `intensity`).
4. Asks the user to specify minimum (`vmin`) and maximum (`vmax`) intensity thresholds.
   - Values outside this range are masked and displayed as white.
5. Iterates through all unique z-slices and creates a 2D intensity map (x vs. y):
   - Rows = y-coordinates, Columns = x-coordinates
   - Color scale = viridis (with masked values shown in white)
   - One heatmap per z-slice
6. Displays each heatmap with axes labeled (x, y) and a colorbar labeled "Intensity (mM)".

Key Features:
- Consistent x,y grid across all z-slices for alignment.
- Flexible intensity thresholding for visualization control.
- Missing or invalid values are shown as white.
- Heatmaps generated using `seaborn.heatmap`.

Dependencies:
- Python standard library (`tkinter`)
- Third-party libraries:
  • pandas  
  • numpy  
  • matplotlib  
  • seaborn  

Usage:
- Run the script in Python.
- Select an Excel file containing an "all" sheet with voxel data.
- Enter min/max intensity thresholds when prompted.
- Inspect heatmaps for each z-slice.

"""


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
    exit()

# --- User-defined settings ---
vmin = float(input("Enter minimum intensity value (e.g. 0): ") or 0)
vmax = float(input("Enter maximum intensity value (e.g. 10): ") or 10)
# -----------------------------

# Read the "all" sheet using row 4 (index 3) as the header
df = pd.read_excel(excel_file, sheet_name='all', header=3)

# Clean column names
df.columns = df.columns.str.strip()

# Check loaded columns
print("Loaded columns:", df.columns.tolist())

# Split 'Coord_p' into x, y, z
coords = df['Coord_p'].astype(str).str.split('_', expand=True)
df['x'] = coords[0].astype(int)
df['y'] = coords[1].astype(int)
df['z'] = coords[2].astype(int)

# Rename intensity column
df.rename(columns={'c [mM]': 'intensity'}, inplace=True)

# Ensure 'intensity' is numeric
df['intensity'] = pd.to_numeric(df['intensity'], errors='coerce')

# Get unique z slices
z_slices = sorted(df['z'].unique())

# For consistent axis coverage across slices, compute full x,y ranges now
xs = np.arange(df['x'].min(), df['x'].max() + 1)
ys = np.arange(df['y'].min(), df['y'].max() + 1)

# Define colormap with white for masked/NaN values
cmap = plt.get_cmap('viridis')
new_colors = cmap(np.linspace(0, 1, 256))
new_cmap = ListedColormap(new_colors)
# Use set_bad so NaN/masked entries render white
new_cmap.set_bad((1.0, 1.0, 1.0, 1.0))

# Plot each z slice as heatmap
for z in z_slices:
    slice_df = df[df['z'] == z]

    # Pivot to 2D grid (rows = y, cols = x)
    intensity_grid = slice_df.pivot(index='y', columns='x', values='intensity')

    # Ensure full grid coverage and ascending order so row 0 maps to the top row visually (no invert)
    intensity_grid = intensity_grid.reindex(index=ys, columns=xs).sort_index(axis=0).sort_index(axis=1)

    # Replace empty strings with NaN and ensure float dtype
    intensity_grid = intensity_grid.replace('', np.nan).astype(float)

    # Mask values outside chosen vmin/vmax and also the NaNs
    out_of_range_mask = (intensity_grid < vmin) | (intensity_grid > vmax) | intensity_grid.isna()
    # Make a copy where out-of-range are set to NaN so they render as white (set_bad)
    plot_data = intensity_grid.mask(out_of_range_mask)

    plt.figure(figsize=(8, 6))
    sns.heatmap(
        plot_data,
        cmap=new_cmap,
        cbar_kws={'label': 'Intensity (mM)'},
        linewidths=0.5,
        vmin=vmin,
        vmax=vmax,
        mask=plot_data.isna(),  # ensure NaNs are masked (and show as white)
        square=False
    )
    plt.title(f'Intensity Map at z = {z} (values outside [{vmin}, {vmax}] shown white)')
    plt.xlabel('x')
    plt.ylabel('y')
    # NOTE: removed plt.gca().invert_yaxis() so y=0 displays at the top (row 0 → top-left)
    plt.tight_layout()
    plt.show()
