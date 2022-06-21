import openpyxl
import pyclipper
import re
import drawSvg as draw
from math import sqrt, pow

range_regex = re.compile(r"!\$(.)\$(\d{3}):\$(.)\$(\d{3})")

letters = [letter for letter in "ABCDEFGHIJKLMNOPQRSTUVXYZ"]

chart_series = [
    "aileron",
    "aileron2",
    "fuselage",
    "fuselage2",
    "Cone",
    "Cone1",
    "canard",
    "canard2"
]

fins_series = [
    "aileron",
    "aileron2",
    "canard",
    "canard2"
]

motor_series = [
    "fuselage",
    "fuselage2"
]

fin_body_contacts = [0, 3, 4]

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

def post_process_fins(series_polys):
    for fins_serie in fins_series:
        if fins_serie not in series_polys: continue
        sign = 1 if series_polys[fins_serie][2][0] > 0 else -1
        for contact_index in fin_body_contacts:
            if len(series_polys[fins_serie]) >= contact_index+1:
                series_polys[fins_serie][contact_index][0] -= sign*0.1
    return series_polys

def post_process_motor(series_polys):
    for motor_serie in motor_series:
        serie = series_polys[motor_serie]
        for i in range(1, len(serie)):
            if serie[i][1] > serie[i-1][1]:
                serie[i][1] = serie[i-1][1]
            
    return series_polys

class StabDrawing():
    
    def __init__(self, path, width, height):
        self.post_process_funcs = [post_process_fins, post_process_motor]
        book = openpyxl.load_workbook(path, data_only=True)
        try:
            self.sheet = book["Stabilito"]
        except KeyError:
            self.sheet = book["stabilito"]
        self.d = draw.Drawing(width, height, origin=(-int(width/2), -height), displayInline=False)
        self.stroke_width = 3
        self.path = path
        
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

    def draw_polygon(self, points):
        for i in range(1, len(points)):
            self.d.append(draw.Line(*points[i-1], *points[i], stroke="red", stroke_width=self.stroke_width))

    def draw(self, path):
        series_polys = dict()
        for serie_name, serie_points in self.series.items():
            points = self.get_points(serie_points)
            if serie_name.lower() in ["cone", "cone1"]:
                points.append((0, points[-1][1]))
            series_polys[serie_name] = points
        
        series_polys = self.clean_series(series_polys)
        
        for post_process_func in self.post_process_funcs:
            series_polys = post_process_func(series_polys)
        
        series_polys = self.clean_series(series_polys)
                
        polygons = list(series_polys.values())
        outline = self.union(polygons)
        if len(outline) != 1:
            raise Exception(f"Error on union")
        self.draw_polygon(outline[0])
        
        self.d.saveSvg(path)
    
    def clean_series(self, series):
        clean_series = dict()
        for poly_name, polygon in series.items():
            clean_poly = self.clean_polygon(polygon)
            if clean_poly is not None:
                clean_series[poly_name] = clean_poly
        return clean_series
    
    def clean_polygon(self, points):
        if len(points) == 0: return None 
        origin = points[0]
        if not any(point != origin for point in points): return None
        polygon = [list(points[0])]
        for i in range(1, len(points)):
            if self.distance(points[i-1], points[i]) > 0:
                polygon.append(list(points[i]))
        return polygon            
    
    def distance(self, point1, point2):
        return sqrt(pow(point1[0]-point2[0], 2) + pow(point1[1]-point2[1], 2))
    
    def union(self, polygons):
        pc = pyclipper.Pyclipper()
        for polygon in polygons:
            if polygon[0] != polygon[-1]:
                polygon.append(polygon[0])
            pc.AddPath(polygon, pyclipper.PT_CLIP, True)
        output = pc.Execute(pyclipper.CT_UNION, pyclipper.PFT_NONZERO, pyclipper.PFT_NONZERO)
        output[0].append(output[0][0])
        return output

def main():
    from glob import glob
    from tqdm import tqdm
    import os
    from pathlib import Path
    os.makedirs("errors", exist_ok=True)
    os.makedirs("outputs", exist_ok=True)
    files = [file for file in glob("cache/*.xlsx") if not "_temp." in file]
    for file in tqdm(files):
        try:
            StabDrawing(Path(file), 2000, 6000).draw(f"outputs/{Path(file).stem}.svg")
        except Exception as e:
            print(f"Error on file {file}: {e}")
            os.rename(file, Path("errors")/Path(file).name)
        finally:
            pass # Présent pour pouvoir facilement commenter la gestion des erreurs

if __name__ == '__main__':
    main()
