# trim.py
from PIL import Image

input_path = "amazon_home.png"
output_path = "amazon_home_cropped.png"

image = Image.open(input_path)

# (left, top, right, bottom)
crop_box = (0, 0, 2500, 500)

cropped = image.crop(crop_box)
cropped.save(output_path)

print("✅ Cropped section saved:", output_path)
