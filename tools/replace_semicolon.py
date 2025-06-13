# === Replace ; with newline in a text file ===
import tkinter as tk
from tkinter import filedialog

# Hide main tkinter window
root = tk.Tk()
root.withdraw()

# Prompt user to select input file
input_path = filedialog.askopenfilename(title="Select text file", filetypes=[("Text Files", "*.txt")])
if not input_path:
    raise Exception("No file selected.")

# Prompt user to select output location
output_path = filedialog.asksaveasfilename(title="Save new file as", defaultextension=".txt", filetypes=[("Text Files", "*.txt")])
if not output_path:
    raise Exception("No output file specified.")

# Read, replace, and save
with open(input_path, 'r', encoding='utf-8') as infile:
    content = infile.read()

new_content = content.replace(';', '\n')

with open(output_path, 'w', encoding='utf-8') as outfile:
    outfile.write(new_content)

print(f"✅ New file saved to: {output_path}")
