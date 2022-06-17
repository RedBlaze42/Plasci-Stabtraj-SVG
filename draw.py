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

usual_distribution = {
    0: "fuselage",
    1: "aileron",
    2: "fuselage2",
    3: "aileron2",
    6: "canard",
    7: "canard2",
    12: "Cone",
    13: "Cone1"
}

class StabDrawing():
    
    def __init__(self, path, width, height):
        book = openpyxl.load_workbook(path, data_only=True)
        self.sheet = book["Stabilito"]
        self.d = draw.Drawing(width, height, origin=(-int(width/2), -height), displayInline=False)
        self.stroke_width = 3
        
        self.series = dict()
        for i, serie in enumerate(self.sheet._charts[0].series):
            serie_title = None
            if serie.title is None:
                if i in usual_distribution:
                    serie_title = usual_distribution[i]            
            elif serie.title.value in chart_series:
                serie_title = serie.title.value.lower()
            
            if serie_title is not None:
                x_points = self.range_reduce(serie.xVal.numRef.f)
                y_points = self.range_reduce(serie.yVal.numRef.f)
                self.series[serie_title] = list(zip(x_points, y_points))
                
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
        for i in range(1, len(points)):
            self.d.append(draw.Line(*points[i-1], *points[i], stroke="red", stroke_width=self.stroke_width))

    def draw_arcs(self, points):
        for i in range(1, len(points)):
            self.d.append(draw.Arc(*points[i-1], *points[i], stroke="red", stroke_width=self.stroke_width))
        
    def draw(self, path):
        polygons = [self.get_points(serie) for serie in self.series.values()]
        for polygon in polygons:
            self.draw_lines(polygon)
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
    #StabDrawing("cache/532_StabtrajOgma.xlsx", 2000,6000).draw("test3.svg")
    #StabDrawing("cache/488_StabTraj-Irydium.xlsx", 2000,6000).draw("test1.svg")
    #StabDrawing("cache/424_StabTraj-MF26-Leofly-Polaris.xlsx", 2000,6000).draw("test2.svg")
    #StabDrawing("cache/390_stabtrajkuntur.xlsx", 2000,6000).draw("test4.svg")
