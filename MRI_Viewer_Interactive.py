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
