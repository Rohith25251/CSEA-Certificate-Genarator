import fitz
import os

pdf_path = r"D:\Projects\CSEA Certificate Generator\test_output_certs\Preethika_sri_K_Certificate.pdf"
img_out = r"D:\Projects\CSEA Certificate Generator\test_output_certs\Preethika_sri_K_Certificate.png"

doc = fitz.open(pdf_path)
page = doc[0]
pix = page.get_pixmap(dpi=150)
pix.save(img_out)

print("Saved preview image to:", img_out)
print("Dimensions:", pix.width, "x", pix.height)
