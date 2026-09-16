import win32com.client as win32
import os, sys

path = os.path.abspath("data/outputs/2025-FY_Business_Review_USD.pptx")
outdir = os.path.abspath("data/outputs/_deck_preview")
os.makedirs(outdir, exist_ok=True)

ppt = win32.gencache.EnsureDispatch("PowerPoint.Application")
ppt.Visible = True
pres = ppt.Presentations.Open(path, WithWindow=True)
pres.SaveAs(outdir, 18)  # 18 = ppSaveAsPNG -> exports each slide as Slide1.PNG etc.
n = pres.Slides.Count
pres.Close()
ppt.Quit()
print("Exported", n, "slide images to", outdir)
