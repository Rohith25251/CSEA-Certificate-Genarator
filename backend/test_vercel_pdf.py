import os
import sys

# Add backend directory to sys.path
sys.path.append(os.path.dirname(__file__))

import generator

# Simulate Vercel/Linux environment by forcing has_win32com to False
generator.has_win32com = False

def run_test():
    print("[Test] Simulating Vercel environment (forcing has_win32com = False)...")
    
    # 1. Find a PPTX template
    template_path = generator.get_active_pptx_template_path()
    if not template_path or not os.path.exists(template_path):
        print(f"[Test] No PPTX template found at '{template_path}'. Please upload/save a template first.")
        return
        
    print(f"[Test] Using template: {template_path}")
    
    # 2. Define test replacements
    replacements = {
        "Name": "John Doe Test Vercel",
        "Roll Number": "25CSR000",
        "Date": "24-Aug-2026",
        "Event Name": "Vercel Fallback PDF Test",
        "college_name": "Kongu Engineering College"
    }
    
    # 3. Output path
    output_pdf = os.path.join(os.path.dirname(__file__), "test_vercel_fallback_output.pdf")
    if os.path.exists(output_pdf):
        try:
            os.remove(output_pdf)
        except Exception:
            pass
        
    # 4. Generate PDF
    print("[Test] Running generate_single_native_pdf...")
    success = generator.generate_single_native_pdf(template_path, replacements, output_pdf)
    
    if success and os.path.exists(output_pdf):
        print(f"[Test] SUCCESS! PDF rendered successfully at: {output_pdf}")
        print(f"[Test] File size: {os.path.getsize(output_pdf)} bytes")
    else:
        print("[Test] FAILURE! PDF rendering returned False or output file not found.")

if __name__ == "__main__":
    run_test()
