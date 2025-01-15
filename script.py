import pyautogui
import tkinter as tk
from tkinter import filedialog
from PIL import Image, ImageTk
from docx import Document
from docx.shared import Inches
from pynput import keyboard
import threading
import os
from datetime import datetime

# Global Variables
screenshot_path = "screenshot.png"
cropped_path = "cropped_screenshot.png"
word_file_path = "screenshots.docx"

# Capture Screenshot
def capture_screenshot():
    global screenshot_path
    timestamp = datetime.now().strftime('%Y_%m_%d_%H:%M:%S')
    screenshot_path = f"temp/screenshot_{timestamp}.png"
    screenshot = pyautogui.screenshot()
    screenshot.save(screenshot_path)
    open_crop_gui(screenshot_path)

# Crop GUI
def open_crop_gui(image_path):
    root = tk.Tk()
    root.title("Crop Screenshot")

    img = Image.open(image_path)
    tk_image = ImageTk.PhotoImage(img)

    canvas = tk.Canvas(root, width=img.width, height=img.height)
    canvas.pack()
    canvas.create_image(0, 0, anchor=tk.NW, image=tk_image)

    crop_coords = []

    def on_mouse_press(event):
        crop_coords.clear()
        crop_coords.append((event.x, event.y))

    def on_mouse_release(event):
        crop_coords.append((event.x, event.y))
        crop_image(image_path, crop_coords)
        root.destroy()

    canvas.bind("<ButtonPress-1>", on_mouse_press)
    canvas.bind("<ButtonRelease-1>", on_mouse_release)

    root.mainloop()

# Crop Image
def crop_image(filepath, crop_coords):
    global cropped_path
    (x1, y1), (x2, y2) = crop_coords
    with Image.open(filepath) as img:
        cropped_img = img.crop((x1, y1, x2, y2))
        cropped_path = filepath.replace(".png", "_cropped.png")
        cropped_img.save(cropped_path)
    append_to_word(cropped_path)

# Append to Word File
def append_to_word(image_path):
    if not os.path.exists(word_file_path):
        doc = Document()
    else:
        doc = Document(word_file_path)
    doc.add_picture(image_path, width=Inches(6))
    doc.save(word_file_path)
    print(f"Saved to {word_file_path}")

# Hotkey Listener
def hotkey_listener():
    def on_activate():
        print("Hotkey pressed: Capturing screenshot...")
        capture_screenshot()

    with keyboard.GlobalHotKeys({
        '<ctrl>+<shift>+x': on_activate
    }) as h:
        h.join()

# Main
if __name__ == "__main__":
    threading.Thread(target=hotkey_listener, daemon=True).start()
    print("Application is running. Listening for hotkey...")
    while True:
        pass  # Keep the script running
