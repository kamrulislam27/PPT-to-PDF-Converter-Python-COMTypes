import os
import comtypes.client

def ppt_to_pdf(source_folder, target_folder):
    if not os.path.exists(target_folder):
        os.makedirs(target_folder)

    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
    powerpoint.Visible = 1

    for file in os.listdir(source_folder):
        if file.endswith((".ppt", ".pptx")):
            ppt_file = os.path.join(source_folder, file)
            pdf_file = os.path.join(target_folder, os.path.splitext(file)[0] + ".pdf")
            
            presentation = powerpoint.Presentations.Open(ppt_file)
            presentation.SaveAs(pdf_file, 32)  # 32 = PDF format
            presentation.Close()

    powerpoint.Quit()


# Example usage:
source = r"D:\cxdownload"
target = r"D:\cxdownload\Viona"
ppt_to_pdf(source, target)
