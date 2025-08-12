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

# Define colormap with white for masked values
cmap = plt.get_cmap('viridis')
new_colors = cmap(np.linspace(0, 1, 256))
white = np.array([1, 1, 1, 1])
new_cmap = ListedColormap(new_colors)
new_cmap.set_over(white)
new_cmap.set_under(white)

# Plot each z slice as heatmap
for z in z_slices:
    slice_df = df[df['z'] == z]

    # Pivot to 2D grid
    intensity_grid = slice_df.pivot(index='y', columns='x', values='intensity')

    # Replace invalid entries
    intensity_grid = intensity_grid.replace('', np.nan).astype(float)

    plt.figure(figsize=(8, 6))
    sns.heatmap(
        intensity_grid,
        cmap=new_cmap,
        cbar_kws={'label': 'Intensity (mM)'},
        linewidths=0.5,
        vmin=vmin,
        vmax=vmax
    )
    plt.title(f'Intensity Map at z = {z}')
    plt.xlabel('x')
    plt.ylabel('y')
    plt.gca().invert_yaxis()
    plt.tight_layout()
    plt.show()
