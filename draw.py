import openpyxl
import pyclipper
import re
import drawSvg as draw
from svg_text import Text
from math import sqrt, pow

range_regex = re.compile(r"!\$(.)\$(\d{3}):\$(.)\$(\d{3})")

letters = [letter for letter in "ABCDEFGHIJKLMNOPQRSTUVXYZ"]
max_text_width = 0.75
text_height_percentage = 0.75

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

body_series = [
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
    for motor_serie in body_series:
        serie = series_polys[motor_serie]
        for i in range(1, len(serie)):
            if serie[i][1] > serie[i-1][1]:
                serie[i][1] = serie[i-1][1]
            
    return series_polys

class StabDrawing():
    
    def __init__(self, path, width, height, name, notch_width, font_path="fonts/nasalization-rg.otf", stroke_width=0.01):
        self.post_process_funcs = [post_process_fins, post_process_motor, self.get_body_base_points]
        self.font_path = font_path
        book = openpyxl.load_workbook(path, data_only=True)
        try:
            self.sheet = book["Stabilito"]
        except KeyError:
            self.sheet = book["stabilito"]
        self.d = draw.Drawing(width, height, origin=(-int(width/2), -height), displayInline=False)
        self.width, self.height = width, height
        self.stroke_width = stroke_width
        self.path = path
        self.name = name
        self.notch_width = notch_width
        
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

    def draw_lines(self, lines, color="red"):
        for line in lines:
            self.d.append(draw.Line(*line[0], *line[1], stroke=color, stroke_width=self.stroke_width, fill="none"))

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
        outlines = self.union(polygons)
        if len(outlines) != 1:
            raise Exception(f"Too much polygons after union")
        outline = outlines[0]
        
        lines = list()
        for i in range(1, len(outline)-1):
            if outline[i-1] in self.body_base_points and outline[i] in self.body_base_points: continue
            lines.append([outline[i-1], outline[i]])
        lines.append([lines[-1][1], lines[0][0]])
        self.draw_lines(lines)
        
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
            self.draw_lines([[fin_serie[0], fin_serie[-2]]], color="lime")
        
        # Project name
        base = min(point[1] for point in self.series_polys["fuselage"])
        top = max(point[1] for point in self.series_polys["fuselage"])
        text_y = -int((max(base, top) - min(base, top))/2-top)
        
        min_diam = max(point[0] for point in self.series_polys["fuselage"])
        text_y_range = text_y - max_text_width*(top - base)/2, text_y + max_text_width*(top - base)/2
        text_y_size = text_y_range[1] - text_y_range[0]

        points = [(abs(point[0]), point[1]) for point in self.series_polys["fuselage"]]
        for point in points:
            if point[0] < min_diam and text_y_range[0] < point[1] < text_y_range[1]:
                min_diam = point[0]
        
        if min_diam == max(point[0] for point in self.series_polys["fuselage"]):
            min_diam, min_diam_y = 0, self.height
            for point in points:
                if point[1] < text_y and point[0] > min_diam and point[1] < min_diam_y:
                    min_diam = point[0]
        
        text = Text(self.name.upper(), min_diam*text_height_percentage, self.font_path, (0, text_y), rotate_angle=90, rotate_origin=(0, text_y))
        text_bbox = text.get_bbox()
        
        font_size_modifier = min_diam*text_height_percentage
        while text_bbox[2]-text_bbox[0] > text_y_size or text_bbox[3]-text_bbox[1] > 2*min_diam*text_height_percentage:
            font_size_modifier -= 1
            text.set_font_size(font_size_modifier)
            text_bbox = text.get_bbox()
        
        text.offset[0] -= int(text_bbox[2]/2)
        text.offset[1] -= int(text_bbox[3]/2)
        text.rotation_origin = text.offset
        text.draw(self.d)
        
        # Notch
        base_height = self.body_base_points[0][1]
        for point in self.body_base_points:
            if point[0] < 0:
                self.draw_lines([[point, [-self.notch_width/2, base_height]]])
            else:
                self.draw_lines([[point, [self.notch_width/2, base_height]]])
        self.draw_lines([[[-self.notch_width/2, base_height], [self.notch_width/2, base_height]]], color="blue")
        
        
    def get_body_base_points(self, series_polys):
        self.body_base_points = list()
        for body_serie in body_series:
            base_point = 0, 0
            for x, y in series_polys[body_serie]:
                if abs(y) > abs(base_point[1]):
                    base_point = x, y
                elif abs(y) == abs(base_point[1]) and abs(x) > abs(base_point[0]):
                    base_point = x, y

            self.body_base_points.append(list(int(coord) for coord in base_point))
        if len(self.body_base_points) != 2:
            raise Exception("Too many base points")
        return series_polys

def draw_worker(file, project, base_config):
    from merger import get_scale, mm_per_pix
    from pathlib import Path
    
    try:
        output_path = f"output_rockets/{Path(file).stem}.svg"
        drawing = StabDrawing(Path(file), 2000, 6000, project["name"], 2)
        drawing.draw(output_path)
        scale = get_scale(output_path, base_config["rectangle_size"])
        drawing.notch_width = mm_per_pix*base_config["notch_size"]/scale
        drawing.draw(output_path)
    except Exception as e:
        return e, file
    finally:
        pass
    
    return None

def main():
    from glob import glob
    from tqdm import tqdm
    import os, json
    from pathlib import Path
    from svg2pdf import bulk_convert
    from concurrent.futures import ProcessPoolExecutor, as_completed
    
    os.makedirs("errors", exist_ok=True)
    os.makedirs("output_rockets", exist_ok=True)
    files = [file for file in glob("cache/*.xlsx") if not "_temp." in file]
    errors = list()
    
    with open("config.json") as f:
        bases = json.load(f)["bases"]
    
    with open("cache/project_list.json", "r", encoding="utf-8") as f:
        project_data = {project["id"]: project for project in json.load(f)["project_list"]}
    
    pool = ProcessPoolExecutor(max_workers=3)
    futures = list()
    for file in files:
        project = project_data[Path(file).name.split("_")[0]]
        futures.append(pool.submit(draw_worker, file, project, bases[project["type"]]))

    progress_bar = tqdm(total=len(futures), desc="Dessin des fusées")
    progress_bar.set_postfix({"errors": len(errors)})
    for future in as_completed(futures):
        progress_bar.update(1)
        result = future.result()
        if result is not None:
            e, file = result
            errors.append((file, e))
            os.rename(file, Path("errors")/Path(file).name)
            progress_bar.set_postfix({"errors": len(errors)})

    progress_bar.close()
    
    for error in errors:
        print(f" - Erreur sur le fichier {Path(error[0]).name}: ({type(error[1]).__name__}) {error[1]}")
        
    bulk_convert("output_rockets/*.svg", "output_rockets.pdf")

if __name__ == '__main__':
    main()
