import math
import freetype

import warnings
warnings.simplefilter("ignore")
import drawSvg as draw

load_flags = freetype.FT_LOAD_DEFAULT | freetype.FT_LOAD_NO_BITMAP

class Text():
    
    def __init__(self, text, font_size, font_name, position, rotate_origin=(0, 0), kern_margin=15, path_kwargs=None, vertical=False):
        self.text = text
        self.font_size = font_size
        self.vertical = vertical
        self.scale = 1024/self.font_size
        self.font = freetype.Face(font_name)
        self.font.set_pixel_sizes(32, 32)
        
        self.offset = list(position)
        self.rotate_angle = math.radians(90) if not vertical else 0
        self.rotate_origin = rotate_origin
        self.kern_margin = kern_margin
        
        if path_kwargs is None:
            self.path_kwargs = dict(stroke_width=2, stroke=None, fill='black')
        else:
            self.path_kwargs = path_kwargs
        
    def set_font_size(self, font_size):
        self.font_size = font_size
        self.scale = 1024/self.font_size
    
    def draw(self, drawing):
        for letter in self.text:
            if letter == " ":
                if self.vertical:
                    self.offset[1] += self.font_size
                else:
                    self.offset[0] += self.font_size
            else:
                drawing.append(self.draw_letter(letter))
                cbox = self.font.glyph.outline.get_cbox()
                if self.vertical:
                    self.offset[1] += cbox.yMax/self.scale + self.kern_margin
                else:
                    self.offset[0] += cbox.xMax/self.scale + self.kern_margin
                
    def get_bbox(self):
        bboxes = list()
        offset = [0, 0]
        for letter in self.text:
            if letter == " ":
                if self.vertical:
                    offset[1] += self.font_size
                else:
                    offset[0] += self.font_size
            else:
                self.font.load_char(letter, load_flags)
                cbox = self.font.glyph.outline.get_cbox()
                bboxes.append([cbox.xMin/self.scale+offset[0], cbox.yMin/self.scale+offset[1], cbox.xMax/self.scale+offset[0], cbox.yMax/self.scale+offset[1]])
                if self.vertical:
                    offset[1] += cbox.yMax/self.scale + self.kern_margin
                else:
                    offset[0] += cbox.xMax/self.scale + self.kern_margin
        
        output = list()
        for i in range(4):
            if i < 2:
                f = min
            else:
                f = max
            output.append(f(bbox[i] for bbox in bboxes))
        
        if not self.vertical:
            output[0], output[1], output[2], output[3] = output[1], output[0], output[3], output[2]
        return output
        
    def draw_letter(self, letter):
        p = draw.Path(**self.path_kwargs)
        self.font.load_char(letter, load_flags)
        if self.vertical:
            old_offset_x = self.offset[0]
            cbox = self.font.glyph.outline.get_cbox()
            self.offset[0] -= (cbox.xMax+cbox.xMin)/(self.scale*2)
            
        self.font.glyph.outline.decompose(p, move_to=self.move_to, line_to=self.line_to, conic_to=self.conic_to, cubic_to=self.cubic_to)
        
        if self.vertical:
            self.offset[0] = old_offset_x
        return p
    
    def move_to(self, a, ctx):
        points = self.transform_points([a])
        ctx.M(*self.reduce_points(points))

    def line_to(self, a, ctx):
        points = self.transform_points([a])
        ctx.L(*self.reduce_points(points))

    def conic_to(self, a, b, ctx):
        points = self.transform_points([a, b])
        ctx.Q(*self.reduce_points(points))

    def cubic_to(self, a, b, c, ctx):
        points = self.transform_points([a, b, c])
        ctx.C(*self.reduce_points(points))
        
    def transform_points(self, points):
        output = list()
        for point in points:
            point = [point.x/self.scale+self.offset[0], point.y/self.scale-self.offset[1]]
            
            point = self.rotate_point(point)
            output.append([point[0], point[1]])
        return output
    
    def rotate_point(self, point):
        ox, oy = self.rotate_origin
        px, py = point
        qx = ox + math.cos(self.rotate_angle) * (px - ox) - math.sin(self.rotate_angle) * (py - oy)
        qy = oy + math.sin(self.rotate_angle) * (px - ox) + math.cos(self.rotate_angle) * (py - oy)
        return (qx, qy)
    
    def reduce_points(self, points):
        return [coord for point in points for coord in point]
