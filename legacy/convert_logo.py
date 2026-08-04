from PIL import Image

# Convert SILVINA V08.png to icon
img = Image.open("assets/SILVINA V08.png")
img.save("assets/silvina_v08.ico", format="ICO", sizes=[(256, 256)])
print("✅ Logo SILVINA V08 convertido a .ico")
