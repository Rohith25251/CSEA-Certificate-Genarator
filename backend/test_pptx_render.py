import os
import sys

def test_win32com():
    try:
        import win32com.client
        print("win32com is installed!")
        try:
            ppt = win32com.client.Dispatch("PowerPoint.Application")
            print("PowerPoint COM Application successfully initialized!")
            ppt.Quit()
            return True
        except Exception as e:
            print("PowerPoint COM initialization failed:", e)
            return False
    except ImportError:
        print("win32com module not installed.")
        return False

def test_soffice():
    import subprocess
    try:
        res = subprocess.run(["soffice", "--version"], capture_output=True, text=True)
        print("soffice version:", res.stdout.strip())
        return True
    except Exception as e:
        print("soffice not found:", e)
        return False

print("Testing PPTX rendering engines:")
has_ppt = test_win32com()
has_lo = test_soffice()
