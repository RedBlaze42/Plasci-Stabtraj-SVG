import svgutils
from svgpathtools import svg2paths
from math import ceil

mm_per_pix = 2.82

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

def merge_svg(base_data, rocket_path, output_path):
    base_path, rectangle_coords, target_size = base_data["path"], base_data["rectangle_coords"], base_data["rectangle_size"]
    bbox = svg_bbox(rocket_path)
    orig_dims = bbox[1] - bbox[0], bbox[3] - bbox[2]
    inverse_scales = orig_dims[1]/(target_size[0]/mm_per_pix), orig_dims[0]/(target_size[1]/mm_per_pix)
    inverse_scale = max(*inverse_scales)
    
    scale = (mm_per_pix**2)/(inverse_scale)
    rocket = svgutils.compose.SVG(rocket_path)
    rocket.rotate(90, 0,0)
    rocket.move(f"{orig_dims[1]+rectangle_coords[0]/scale}mm", f"{orig_dims[0]+rectangle_coords[1]/scale}mm")
    rocket.scale(scale)

    figure = svgutils.transform.fromfile(base_path)
    figure.append(rocket)
    figure.save(output_path)
    
def main():
    import glob, json, os
    from pathlib import Path
    from tqdm import tqdm
    os.makedirs("output_cards", exist_ok=True)
    with open("cache/project_list.json", "r", encoding="utf-8") as f:
        project_types = {project["id"]: project["type"] for project in json.load(f)["project_list"]}
    with open("config.json", "r", encoding="utf-8") as f:
        base_data = json.load(f)["bases"]
        
    for rocket in tqdm(glob.glob("outputs/*.svg")):
        rocket_base_data = base_data[project_types[Path(rocket).name.split("_")[0]]]
        merge_svg(rocket_base_data, rocket, rocket.replace("outputs","output_cards"))
        
if __name__ == "__main__":
    main()