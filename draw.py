import openpyxl
import drawSvg as draw

chart_series = ["aileron", "aileron2", "fuselage", "fuselage2", "Cone", "Cone1"]

class StabDrawing():
    
    def __init__(self, path, width, height):
        book = openpyxl.load_workbook(path, data_only=True)
        self.sheet = book["Stabilito"]
        self.d = draw.Drawing(width, height, origin=(-int(width/2), -height), displayInline=False)
        self.stroke_width = 3
        self.offset = (0, 0)
        
        self.series = list()
        for serie in self.sheet._charts[0].series:
            if serie.title.value in chart_series:
                self.series.append((serie.title.value, serie.xVal.numRef.f, serie.yVal.numRef.f))
        
        print(self.series)
        
    def offset_points(self, points):
        letters = "ABCDEFGHIJKLMNOPQRSTUVXYZ"
        output = list()
        for point in points:
            point_export = list()
            for coord in point:
                x = letters[letters.index(coord[0])+self.offset[0]]
                y = int(coord[1:])+self.offset[1]
                point_export.append(f"{x}{y}")
            output.append(point_export)
            
        return output
        
    def get_points(self, points):
        points = self.offset_points(points)
        output = list()
        for x, y in points:
            output.append((self.sheet[y].value, self.sheet[x].value))
            
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
        fin_coords = [
            ["C133", "D133"],
            ["C134", "D134"],
            ["C135", "D135"],
            ["C136", "D136"],
            ["C137", "D137"]
        ]
        fin2_coords = [
            ["C133", "E133"],
            ["C134", "E134"],
            ["C135", "E135"],
            ["C136", "E136"],
            ["C137", "E137"]
        ]
        tube_coords = [
            ["C132", "E132"],
            ["C131", "E131"],
            ["C130", "E130"],
            ["C129", "E129"],
            ["C128", "E128"],
            ["C127", "E127"],
            ["C126", "E126"],
        ]
        tube2_coords = [
            ["C132", "D132"],
            ["C131", "D131"],
            ["C130", "D130"],
            ["C129", "D129"],
            ["C128", "D128"],
            ["C127", "D127"],
            ["C126", "D126"],
            ["C126", "D126"],
        ]
        fairing_coords = [
            ["C179", "D179"],
            ["C178", "D178"],
            ["C177", "D177"],
            ["C176", "D176"],
            ["C175", "D175"],
            ["C174", "D174"],
        ]
        fairing2_coords = [
            ["C179", "E179"],
            ["C178", "E178"],
            ["C177", "E177"],
            ["C176", "E176"],
            ["C175", "E175"],
            ["C174", "E174"],
        ]
        
        self.draw_lines(fin_coords)
        self.draw_lines(fin2_coords)
        self.draw_lines(tube_coords)
        self.draw_lines(tube2_coords)
        self.draw_lines(fairing_coords)
        self.draw_lines(fairing2_coords)
        self.d.saveSvg(path)

def main():
    from glob import glob
    from tqdm import tqdm
    import os
    from pathlib import Path
    os.makedirs("outputs", exist_ok=True)
    for file in tqdm(glob("cache/*.xlsx")):
        StabDrawing(file, 2000, 6000).draw(f"outputs/{Path(file).stem}")
    

if __name__ == '__main__':
    #main()
    #StabDrawing("cache/2644.xlsx", 600, 1300).draw("test.svg")
    StabDrawing("cache/2237.xlsx", 2000, 6000).draw("test2.svg")