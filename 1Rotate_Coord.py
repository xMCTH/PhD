"""
Voxel Coordinate Rotation Script
--------------------------------

This script allows you to rotate voxel coordinate data stored in a text file of fitted MRSI data.
It can be used, if for a profile correction of a subject measurement, the phantom measurement is not rotated the same way in the raster.
Workflow:
1. Prompts the user to select an input `.txt` file containing voxel coordinate blocks.
   - Each voxel block is expected to have a "Coord" header line, a data line with
     coordinates in the format "X_Y_Z" followed by tab-separated fields, and an end line.
2. Asks the user for a rotation angle (90, 180, or 270 degrees).
   - Rotation is applied around the Z-axis in a 3D grid of size GRID_X × GRID_Y × GRID_Z.
3. Applies the chosen rotation to each voxel’s (x, y, z) coordinates while preserving
   the rest of the data.
4. Prompts the user to choose a save location and writes the rotated voxel data to
   a new `.txt` file.

Dependencies:
- Python standard library (`tkinter`, `os`)

Usage:
- Run the script, choose an input file, specify a rotation angle, and save the rotated output.

"""


import tkinter as tk
from tkinter import filedialog, simpledialog
import os

# Grid dimensions — update as needed
GRID_X, GRID_Y, GRID_Z = 32, 32, 8

def rotate_coord(x, y, z, rotation):
    if rotation == 90:
        return GRID_Y - 1 - y, x, z
    elif rotation == 180:
        return GRID_X - 1 - x, GRID_Y - 1 - y, z
    elif rotation == 270:
        return y, GRID_X - 1 - x, z
    else:
        raise ValueError("Unsupported rotation angle")

def main():
    # GUI root
    root = tk.Tk()
    root.withdraw()

    # Select input file
    input_path = filedialog.askopenfilename(
        title="Select Input TXT File",
        filetypes=[("Text Files", "*.txt")]
    )
    if not input_path:
        print("No input file selected.")
        return

    # Ask for rotation angle
    rotation = simpledialog.askinteger("Rotation", "Enter rotation angle (90, 180, 270):", minvalue=90, maxvalue=270)
    if rotation not in [90, 180, 270]:
        print("Invalid rotation angle.")
        return

    # Parse and rotate
    with open(input_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    output_lines = []
    i = 0
    while i < len(lines):
        if lines[i].strip().startswith("Coord"):
            # Start of voxel block
            header_line = lines[i].strip()
            data_line = lines[i + 1].strip()
            end_line = lines[i + 2].strip()

            parts = data_line.split('\t')
            coord_str = parts[0]

            try:
                x_str, y_str, z_str = coord_str.strip().split("_")
                x, y, z = int(x_str), int(y_str), int(z_str)
                new_x, new_y, new_z = rotate_coord(x, y, z, rotation)
                new_coord_str = f"{new_x}_{new_y}_{new_z}"
                parts[0] = new_coord_str
                rotated_line = "\t".join(parts)

                output_lines.extend([header_line + '\n', rotated_line + '\n', end_line + '\n'])
            except Exception as e:
                print(f"Error processing line {i+1}: {e}")
                # Keep original lines if parsing fails
                output_lines.extend([lines[i], lines[i + 1], lines[i + 2]])

            i += 3  # Advance to next block
        else:
            # Copy any unrelated line
            output_lines.append(lines[i])
            i += 1

    # Save to new file
    output_path = filedialog.asksaveasfilename(
        title="Save Rotated Output As",
        defaultextension=".txt",
        filetypes=[("Text Files", "*.txt")]
    )
    if not output_path:
        print("No output file selected.")
        return

    with open(output_path, 'w', encoding='utf-8') as f:
        f.writelines(output_lines)

    print(f"Rotation complete. Output saved to: {output_path}")

if __name__ == "__main__":
    main()
