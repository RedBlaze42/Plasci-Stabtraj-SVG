import glob, json
from pathlib import Path

with open("cache/project_list.json", "r", encoding="utf-8") as f:
    project_list = json.load(f)["project_list"]
    project_ids = {int(project["id"]) for project in project_list}

file_list = glob.glob("cache/*.xlsx")
file_ids = {int(Path(file).name.split("_")[0]) for file in file_list}

missing_files = [project_id for project_id in project_ids if project_id not in file_ids]

if len(missing_files) > 0:
    print(f"Missing {len(missing_files)} files for project IDs:")
    for missing_id in missing_files:
        print(f" - {missing_id}")