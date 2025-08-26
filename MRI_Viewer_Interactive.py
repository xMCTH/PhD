"""
Interactive DICOM Viewer with Brightness Control
------------------------------------------------

This script loads DICOM image files from a specified folder and displays them one by one
with an interactive brightness slider. It allows easy visualization and adjustment of
grayscale intensity of MRI data.

Workflow:
1. Define the path to a folder containing `.dcm` DICOM files (`dcm_folder`).
2. Iterates through all DICOM files in the folder.
3. For each file:
   - Reads metadata (Patient Name, Modality, Study Date).
   - Extracts the pixel array from the DICOM object.
   - Displays the image using Matplotlib in grayscale.
   - Provides a slider to interactively adjust brightness in real time.
4. Normalizes image intensities to [0,1] for consistent visualization.

Key Features:
- Interactive brightness adjustment using Matplotlib’s `Slider` widget.
- Automatic DICOM metadata extraction and display in the console.
- Compatible with PyCharm or other IDEs by explicitly setting
  the Matplotlib backend to `TkAgg`.

Dependencies:
- Python standard library (`os`)
- Third-party libraries:
  • numpy
  • matplotlib
  • pydicom

Usage:
- Update the `dcm_folder` variable with the path to your DICOM folder.
- Run the script in Python.
- Scroll through all `.dcm` files in the folder.
- Adjust brightness with the slider to enhance visibility.

"""


import matplotlib
matplotlib.use('TkAgg')  # Use TkAgg backend for interactive sliders in PyCharm

import pydicom
import matplotlib.pyplot as plt
from matplotlib.widgets import Slider
import os
import numpy as np

# Set the path to your DICOM folder
dcm_folder = r"C:\Users\marc_\OneDrive - Universitaet Bern\PhD\DATA\Results\JR_Liver"

def show_image_with_brightness_slider(image, title):
    # Normalize image to [0,1]
    img_norm = image.astype(float)
    img_norm = np.maximum(img_norm, 0) / img_norm.max()

    fig, ax = plt.subplots(figsize=(6, 6))
    plt.subplots_adjust(bottom=0.25)  # Make space for slider

    # Initial display
    img_display = ax.imshow(img_norm, cmap='gray')
    ax.set_title(title)
    ax.axis('off')

    # Slider axis
    ax_brightness = plt.axes([0.25, 0.1, 0.5, 0.03])
    slider = Slider(ax_brightness, 'Brightness', 0.1, 8.0, valinit=1.0)

    # Update function for slider
    def update(val):
        brightness = slider.val
        new_img = np.clip(img_norm * brightness, 0, 1)
        img_display.set_data(new_img)
        fig.canvas.draw_idle()

    slider.on_changed(update)
    plt.show()

# Loop through all files in the folder
for filename in os.listdir(dcm_folder):
    if filename.endswith(".dcm"):
        dcm_path = os.path.join(dcm_folder, filename)

        # Load the DICOM file
        ds = pydicom.dcmread(dcm_path)

        # Print basic metadata
        print(f"File: {filename}")
        print("  Patient Name:", ds.get("PatientName", "Unknown"))
        print("  Modality:", ds.get("Modality", "Unknown"))
        print("  Study Date:", ds.get("StudyDate", "Unknown"))

        # Get image data
        image = ds.pixel_array

        # Show image with brightness slider
        show_image_with_brightness_slider(image, f'MR Image: {filename}')
