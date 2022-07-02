import os, shutil, glob
from pathlib import Path
from svglib.svglib import svg2rlg
from reportlab.graphics import renderPDF
from PyPDF2 import PdfMerger

def bulk_convert(from_glob, output):
    os.makedirs("temp", exist_ok=True)
    for svg_file in glob.glob(from_glob):
        drawing = svg2rlg(svg_file)
        renderPDF.drawToFile(drawing, f"temp/{Path(svg_file).name}.pdf")
        
    merger = PdfMerger()
    for pdf_file in glob.glob("temp/*.pdf"):
        merger.append(pdf_file)

    merger.write(output)
    merger.close()
    shutil.rmtree("temp")

if __name__ == "__main__":
    bulk_convert("output_rockets/*.svg", "output.pdf")