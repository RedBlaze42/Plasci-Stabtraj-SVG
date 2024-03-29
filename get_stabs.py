import requests, json
import os, time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from bs4 import BeautifulSoup
from tqdm import tqdm
from subprocess import Popen
from sanitize_filename import sanitize

max_workers = 3

with open("config.json") as f:
    config = json.load(f)

os.makedirs("cache", exist_ok=True)
os.makedirs("errors", exist_ok=True)
os.makedirs("cache/raw_files", exist_ok=True)
session = requests.Session()
session.headers = {"Cookie": config["cookies"], "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/102.0.0.0 Safari/537.36"}
    
def get_path_from_id(project_id):
    return Path("cache")/Path(f"{project_id}.xlsx")

def get_projects():
    url = f"https://www.planete-sciences.org/espace/scae/index.php?p=api&key={config['api_key']}"
    req = session.get(url)
    req.raise_for_status()
    
    output = list()
    for project in req.json():
        if project["launch_year"] is not None and int(project["launch_year"]) == int(config["launch_year"]) and project["type"] in ["minif", "fusex"] and project["status"] == "wip":
            output.append(project)
    
    return output

def convert_workbook(from_path, dest_path):
    process = Popen(f".\\convert.vbs \"{from_path}\" \"{dest_path}\"", shell=True)
    
    while process.poll() == None:
        time.sleep(1)
    
    if process.poll() != 0:
        os.rename(from_path, Path("errors") / Path(from_path).name)
        return False, from_path
    else:
        os.rename(from_path, Path("cache/raw_files") / Path(from_path).name)
        return True, from_path

def get_url_from_id(project_id):
    return f"https://www.planete-sciences.org/espace/scae/edit_project&id={project_id}"

def get_project_details(project_id):
    url = get_url_from_id(project_id)
    req = session.get(url)
    req.raise_for_status()
    soup = BeautifulSoup(req.text, features="lxml")
    
    try:
        if soup.find("select", {"id": "project__campaign"}).find("option", {"selected": True}).text != config["campaign"]: return "Invalid campaign", None
        
        rce3_msg = soup.select(".border-rce3")
        stabratjs = soup.find("h3", text="StabTraj's").parent.findAll("a")
        last_stab = sorted((stab for stab in stabratjs), key=lambda x: int(x["href"].split("=")[-1]) if x["href"] != "#" else 0, reverse=True)[0]
        last_stab_text = last_stab.parent.parent.select_one(".project-document-filename").text
        
        output = {
            "project_name": soup.find("input", {"id": "project__name"})["value"],
            "club_name": soup.find("select", {"id": "project__club"}).find("option", {"selected": True}).text,
            "stabtraj_url": "https://www.planete-sciences.org/" + last_stab["href"],
            "project_id": int(project_id),
            "rce3": rce3_msg[-1].text if len(rce3_msg) > 0  else None
        }
        output["stabtraj_id"] = int(output['stabtraj_url'].split('=')[-1])
        
        stab_path = Path("cache")/Path(sanitize(f"{project_id}_{last_stab_text.replace(' ', '')}"))
        if not stab_path.exists():
            with open(stab_path, "wb") as f:
                stab_req = session.get(output["stabtraj_url"])
                stab_req.raise_for_status()
                f.write(stab_req.content)
            
            if str(stab_path).endswith(".xlsx"):
                temp_path = str(stab_path).replace(".xlsx", "_temp.xlsx")
                os.rename(stab_path, temp_path)
                new_path = str(stab_path.absolute())
                stab_path = Path(temp_path)
            else:
                new_path = str(stab_path.absolute())+"x"
            
            return output, (stab_path.absolute(), new_path)
        
    except AttributeError as e:
        print(e)
        return None, None
    finally:
        pass
    
    return output, None

def main():
    projects = get_projects()
    projects_details = list()
    missing_projects = list()
    converter = ThreadPoolExecutor(max_workers=max_workers)
    futures = list()
    
    progress_bar = tqdm(total=len(projects), desc="Téléchargement des stabtraj")
    progress_bar.set_postfix({"errors": 0})
    for project in projects:
        progress_bar.update(1)
        project_details, converter_args = get_project_details(project["id"])

        if project_details is None:
            print(f"Error on project {project['name']}")
            missing_projects.append(project)
            projects.remove(project)
            progress_bar.set_postfix({"errors": len(missing_projects)})
        elif project_details == "Invalid campaign":
            continue
        else:
            projects_details.append(project_details)
        
        if converter_args is not None:
            futures.append(converter.submit(convert_workbook, *converter_args))

    progress_bar = tqdm(total=len(futures), desc="Conversion des stabtraj")
    progress_bar.set_postfix({"errors": 0})
    errors = list()
    for future in as_completed(futures):
        result = future.result()
        if not result[0]: errors.append(result[1])
        progress_bar.update(1)
        progress_bar.set_postfix({"errors": len(errors)})
    progress_bar.close()
        
    with open(Path("cache")/Path("project_list.json"), "w", encoding="utf-8") as f:
        json.dump({"project_details":projects_details, "project_list": projects, "missing_projects": missing_projects}, f, ensure_ascii=False)
    
    if len(missing_projects) > 0:
        print("Erreurs de téléchargement:")
        for error in missing_projects:
            print(f" - {error['name']}:{get_url_from_id(error['id'])}")
    
    if len(errors) > 0:
        print("Erreurs de conversion:")
        for error in errors:
            print(f" - {Path(error).name}")
        

if __name__ == "__main__":
    main()