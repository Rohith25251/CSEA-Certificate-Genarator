import os
from generator import generate_single_native_pdf, get_active_pptx_template_path

template_path = get_active_pptx_template_path()
print(f"Using template: {template_path}")

output_pdf = os.path.abspath(os.path.join(os.path.dirname(__file__), "generated_pdfs", "Test_Preserved_Style.pdf"))

replacements = {
    "Name": "Preethika sri K",
    "Roll Number": "25CSR220",
    "Roll Number ": "25CSR220",
    "Date": "25.07.2026",
    "Date ": "25.07.2026",
    "event_name": "DATASET TO DECISION Workshop"
}

success = generate_single_native_pdf(template_path, replacements, output_pdf)
print(f"PDF Generation Result: {success}")
print(f"Saved at: {output_pdf}")
