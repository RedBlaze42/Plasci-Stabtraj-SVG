import openpyxl
import pyclipper
import re
import drawSvg as draw
from svg_text import Text
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
    
    def __init__(self, path, width, height, name, font_path="fonts/nasalization-rg.otf"):
        self.post_process_funcs = [post_process_fins, post_process_motor]
        self.font_path = font_path
        book = openpyxl.load_workbook(path, data_only=True)
        try:
            self.sheet = book["Stabilito"]
        except KeyError:
            self.sheet = book["stabilito"]
        self.d = draw.Drawing(width, height, origin=(-int(width/2), -height), displayInline=False)
        self.stroke_width = 3
        self.path = path
        self.name = name
        
        self.series = dict()
        for i, serie in enumerate(self.sheet._charts[0].series):
            serie_title = None
            if serie.title is None:
                if i in usual_distribution:
                    serie_title = usual_distribution[i].lower()
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

    def draw_polygon(self, points, color="red"):
        for i in range(1, len(points)):
            self.d.append(draw.Line(*points[i-1], *points[i], stroke=color, stroke_width=self.stroke_width, fill="none"))

    def draw(self, path):
        self.series_polys = dict()
        for serie_name, serie_points in self.series.items(): # TODO Ajouter dans post-process
            points = self.get_points(serie_points)
            if serie_name.lower() in ["cone", "cone1"]:
                points.append((0, points[-1][1]))
            self.series_polys[serie_name] = points
        
        self.series_polys = self.clean_series(self.series_polys)
        
        for post_process_func in self.post_process_funcs:
            self.series_polys = post_process_func(self.series_polys)
        
        self.series_polys = self.clean_series(self.series_polys)
                
        polygons = list(self.series_polys.values())
        outline = self.union(polygons)
        if len(outline) != 1:
            raise Exception(f"Too much polygons after union")
        self.outline = outline[0]
        self.draw_polygon(self.outline)
        
        self.draw_extras()
        
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
    
    def draw_extras(self):
        # Fin lines
        for fin_serie_name in fins_series:
            if fin_serie_name not in self.series_polys: continue
            fin_serie = self.series_polys[fin_serie_name]
            self.draw_polygon([fin_serie[0], fin_serie[-2]], color="lime")
        
        # Project name
        base = min(point[1] for point in self.series_polys["fuselage"])
        base_fairing = min(point[1] for point in self.series_polys["cone"])
        text_y = -int((max(base, base_fairing) - min(base, base_fairing))/2-base_fairing)
        min_diam = min(point[0] for point in (self.series_polys["fuselage"]+self.series_polys["fuselage2"]) if point[0] > 0)
        
        text = Text(self.name.upper(), min_diam*0.5, self.font_path, (0, text_y), rotate_angle=90, rotate_origin=(0, text_y))
        text_bbox = text.get_bbox()
        
        #if text_bbox[3] - text_bbox[1]: # TODO Modifier la taille de police dynamiquement
        #    text.set_font_size(min_diam*0.5)
        #    text_bbox = text.get_bbox()
        
        text.offset[0] -= int(text_bbox[2]/2)
        text.offset[1] -= int(text_bbox[3]/2)
        text.rotation_origin = text.offset
        text.draw(self.d)

def main():
    from glob import glob
    from tqdm import tqdm
    import os, json
    from pathlib import Path
    from svg2pdf import bulk_convert
    os.makedirs("errors", exist_ok=True)
    os.makedirs("outputs", exist_ok=True)
    files = [file for file in glob("cache/*.xlsx") if not "_temp." in file]
    errors = list()
    with open("cache/project_list.json", "r", encoding="utf-8") as f:
        project_data = {project["id"]: project for project in json.load(f)["project_list"]}
    
    for file in tqdm(files):
        project = project_data[Path(file).name.split("_")[0]]
        try:
            StabDrawing(Path(file), 2000, 6000, project["name"]).draw(f"outputs/{Path(file).stem}.svg")
        except Exception as e:
            errors.append((file, e))
            os.rename(file, Path("errors")/Path(file).name)
        finally:
            pass # Présent pour pouvoir facilement commenter la gestion des erreurs
        
    for error in errors:
        print(f" - Erreur sur le fichier {Path(error[0]).name}: ({type(error[1]).__name__}) {error[1]}")
        
    bulk_convert("outputs/*.svg", "output_rockets.pdf")

if __name__ == '__main__':
    main()
