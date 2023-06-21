import svgutils
from svgpathtools import svg2paths
import re
from pathlib import Path

mm_per_pix = 2.8346
margin_mm = 0.4

viewbox_regex = re.compile(r'viewBox="(\d+?\.?\d*?) (\d+?\.?\d*?) (\d+?\.?\d*?) (\d+?\.?\d*?)"')
scale_regex = re.compile(r'scale\((\d+\.?\d*) (\d+\.?\d*)\)')
edit_regex = re.compile(r'(<svg.*?)>')
margin_px = margin_mm*mm_per_pix

def svg_bbox(path):
    paths, attributes = svg2paths(path)
    xmin, xmax, ymin, ymax = 0, 0, 0, 0
    for path in paths:
        path_bbox = path.bbox()
        xmin = min(xmin, path_bbox[0])
        xmax = max(xmax, path_bbox[1])
        ymin = min(ymin, path_bbox[2])
        ymax = max(ymax, path_bbox[3])
        
    return xmin, xmax, ymin, ymax

def get_scale(svg_path, target_size):
    bbox = svg_bbox(svg_path)
    orig_dims = bbox[1] - bbox[0], bbox[3] - bbox[2]
    
    inverse_scales = orig_dims[1]/(target_size[0]/mm_per_pix), orig_dims[0]/(target_size[1]/mm_per_pix)
    inverse_scale = max(*inverse_scales)
    new_dims = orig_dims[0]*mm_per_pix/inverse_scale, orig_dims[1]*mm_per_pix/inverse_scale
    
    scale = (mm_per_pix**2)/(inverse_scale)
    return scale

def apply_svg(base_data, rocket_path, output_path):
    base_path, rectangle_coords, target_size = base_data["path"], base_data["rectangle_coords"], base_data["rectangle_size"]
    bbox = svg_bbox(rocket_path)
    orig_dims = bbox[1] - bbox[0], bbox[3] - bbox[2]
    inverse_scales = orig_dims[1]/(target_size[0]/mm_per_pix), orig_dims[0]/(target_size[1]/mm_per_pix)
    inverse_scale = max(*inverse_scales)
    new_dims = orig_dims[0]*mm_per_pix/inverse_scale, orig_dims[1]*mm_per_pix/inverse_scale
    
    scale = (mm_per_pix**2)/(inverse_scale)
    rocket = svgutils.compose.SVG(rocket_path)
    rocket.rotate(90, 0,0)
    
    if base_path is not None:
        translate_x = orig_dims[1] + (rectangle_coords[0] + 0.5*(target_size[0] - new_dims[1]))*mm_per_pix/scale
        translate_y = orig_dims[0]/2 + (rectangle_coords[1] + 0.5*(target_size[1] - new_dims[0]))*mm_per_pix/scale
    else:
        translate_x = orig_dims[1]
        translate_y = orig_dims[0]/2
        
    rocket.move(f"{translate_x}mm", f"{translate_y}mm")
    rocket.scale(scale)

    if base_path is not None:
        figure = svgutils.transform.fromfile(base_path)
    else:
        figure = svgutils.transform.fromfile("bases/empty.svg")
    figure.append(rocket)
    
    figure.save(output_path)
    
    if base_path is not None:
        set_svg_size(output_path)
    else:
        set_svg_size_viewbox(output_path, [0, 0, target_size[0]*mm_per_pix, target_size[1]*mm_per_pix])

def set_svg_size(path):
    with open(path, "r") as f:
        svg_data = f.read()
    svg_viewbox = re.findall(viewbox_regex, svg_data)[0]
    svg_viewbox = [float(coord) for coord in svg_viewbox]
    svg_data = re.sub(edit_regex, f'\g<1> width="{svg_viewbox[2]-svg_viewbox[0]}" height="{svg_viewbox[3]-svg_viewbox[1]}">', svg_data)    
    with open(path, "w") as f:
        f.write(svg_data)
        
def set_svg_size_viewbox(path, viewbox):
    with open(path, "r") as f:
        svg_data = f.read()
    svg_data = re.sub(viewbox_regex, f'viewbox="{viewbox[0]} {viewbox[1]} {viewbox[2]} {viewbox[3]}"', svg_data)
    svg_data = re.sub(edit_regex, f'\g<1> width="{viewbox[2]-viewbox[0]}" height="{viewbox[3]-viewbox[1]}">', svg_data)
    with open(path, "w") as f:
        f.write(svg_data)

def merge_platters(cards, platter_dims, output_dir, prefix=""):
    platter_number = 0
    
    current_coords = [0, 0]
    svg_elements = list()
    
    for card_path in cards:
        card = svgutils.compose.SVG(card_path)
        card_dims = (card.width, card.height)
        
        if (current_coords[0]+card_dims[0])/mm_per_pix > platter_dims[0]:
            current_coords[1] += card_dims[1] + margin_px
            current_coords[0] = 0
        
        if (current_coords[1]+card_dims[1])/mm_per_pix > platter_dims[1]:
            platter = svgutils.compose.Figure(*(dim*mm_per_pix for dim in platter_dims), *svg_elements)
            platter.save(Path(output_dir)/f"{prefix}_{platter_number}.svg")
            current_coords = [0, 0]
            svg_elements = list()
            platter_number += 1
        
        translate_coords = current_coords
        card.move(*translate_coords)
        svg_elements.append(card)
        current_coords[0] = current_coords[0] + card_dims[0] + margin_px
        

    platter = svgutils.compose.Figure(*(dim*mm_per_pix for dim in platter_dims), *svg_elements)
    platter.save(Path(output_dir)/f"{prefix}_{platter_number}.svg")
 
def main():
    import glob, json, os
    from pathlib import Path
    os.makedirs("output_cards", exist_ok=True)
    os.makedirs("output_platters", exist_ok=True)
    
    with open("cache/project_list.json", "r", encoding="utf-8") as f:
        project_types = {project["id"]: project["type"] for project in json.load(f)["project_list"]}
    with open("config.json", "r", encoding="utf-8") as f:
        config = json.load(f)
    base_data = config["bases"]
    minif_copies, fusex_copies = config["minif_copies"], config["fusex_copies"]
        
    for rocket in glob.glob("output_rockets/*.svg"):
        rocket_type = project_types[Path(rocket).name.split("_")[0]]
        rocket_base_data = base_data[rocket_type]
        apply_svg(rocket_base_data, rocket, rocket.replace("output_rockets","output_cards").replace(".svg", f"_{rocket_type}.svg"))
    
    platter_dims = (601, 301)
    list_minif = [card_path for card_path in glob.glob("output_cards/*.svg") if card_path.endswith("minif.svg")]*minif_copies
    list_fusex = [card_path for card_path in glob.glob("output_cards/*.svg") if card_path.endswith("fusex.svg")]*fusex_copies
    list_minif.sort()
    list_fusex.sort()
    merge_platters(list_fusex, platter_dims, "output_platters", prefix="fusex")
    merge_platters(list_minif, platter_dims, "output_platters", prefix="minif")
    print(f"Plateau(x) nécéssaires: {len(glob.glob('output_platters/*.svg'))}")

if __name__ == "__main__":
    main()