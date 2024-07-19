import os
import win32com.client
import pythoncom
import pywintypes

def convert_ppt_to_pdf(ppt_path, pdf_path):
    try:
        pythoncom.CoInitialize()  # initialize the COM library
        powerpoint = win32com.client.Dispatch("Powerpoint.Application")
        deck = powerpoint.Presentations.Open(ppt_path)
        deck.SaveAs(pdf_path, 32)  # 32 represents the PDF format
        deck.Close()
        powerpoint.Quit()
    except pywintypes.com_error as e:
        print(f"Error occurred while converting PPT to PDF: {str(e)}")
    finally:
        pythoncom.CoUninitialize()  # uninitialize the COM library