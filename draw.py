import openpyxl
import re
from xls2xlsx import XLS2XLSX
import drawSvg as draw

range_regex = re.compile(r"!\$(.)\$(\d{3}):\$(.)\$(\d{3})")

letters = [letter for letter in "ABCDEFGHIJKLMNOPQRSTUVXYZ"]

chart_series = {
    "aileron": {"type": "line", "color": "red"},
    "aileron2": {"type": "line", "color": "red"},
    "fuselage": {"type": "line", "color": "red"},
    "fuselage2": {"type": "line", "color": "red"},
    "Cone": {"type": "line", "color": "red"},
    "Cone1": {"type": "line", "color": "red"},
    "canard": {"type": "line", "color": "red"},
    "canard2": {"type": "line", "color": "red"}
}


class StabDrawing():
    
    def __init__(self, path, width, height):
        book = openpyxl.load_workbook(path, data_only=True)
        self.sheet = book["Stabilito"]
        self.d = draw.Drawing(width, height, origin=(-int(width/2), -height), displayInline=False)
        self.stroke_width = 3
        
        self.series = dict()
        for serie in self.sheet._charts[0].series:
            if serie.title.value in chart_series:
                x_points = self.range_reduce(serie.xVal.numRef.f)
                y_points = self.range_reduce(serie.yVal.numRef.f)
                self.series[serie.title.value.lower()] = list(zip(x_points, y_points))
                
        if len(self.series) == 0:
            raise Exception("No series")
                
    def range_reduce(self, a1_range):
        result = re.findall(range_regex, a1_range)[0]
        x1, y1 = self.a1_to_xy(f"{result[0]}{result[1]}")
        x2, y2 = self.a1_to_xy(f"{result[2]}{result[3]}")
        points = [self.xy_to_a1(x, y) for x in range(x1, x2+1) for y in range(y1, y2+1)]
        return list(points)
    
    def xy_to_a1(self, x, y):
        return f"{letters[x]}{y}"
    
    def a1_to_xy(self, a1):
        x = letters.index(a1[0])
        y = int(a1[1:])
        return x, y
        
    def get_points(self, points):
        output = list()
        for x, y in points:
            output.append((self.sheet[x].value, self.sheet[y].value))
            
        return output

    def draw_lines(self, points):
        points = self.get_points(points)
        for i in range(1, len(points)):
            self.d.append(draw.Line(*points[i-1], *points[i], stroke="red", stroke_width=self.stroke_width))

    def draw_arcs(self, points):
        points = self.get_points(points)
        for i in range(1, len(points)):
            self.d.append(draw.Arc(*points[i-1], *points[i], stroke="red", stroke_width=self.stroke_width))
        
    def draw(self, path):
        for serie in self.series.values():
            self.draw_lines(serie)
        self.d.saveSvg(path)

def main():
    from glob import glob
    from tqdm import tqdm
    import os
    from pathlib import Path
    os.makedirs("outputs", exist_ok=True)
    files = [file for file in glob("cache/*.xlsx") if not "_temp." in file]
    for file in tqdm(files):
        try:
            StabDrawing(Path(file), 2000, 6000).draw(f"outputs/{Path(file).stem}.svg")
        except Exception as e:
            print(f"Error on file {file}: {e}")
        finally:
            pass
    

if __name__ == '__main__':
    main()
    #StabDrawing("cache/432_stabtraj_v3-4.2-2 Beyond.xlsx", 2000,6000).draw("test3.csv")
