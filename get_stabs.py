import requests, json
import os
from pathlib import Path
from bs4 import BeautifulSoup
from tqdm import tqdm

with open("config.json") as f:
    config = json.load(f)

os.makedirs("cache", exist_ok=True)
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
        if project["launch_year"] == "2022" and project["type"] in ["minif", "fusex"] and project["status"] == "wip":
            output.append(project)
    
    return output

def get_project_details(project_id):
    url = f"https://www.planete-sciences.org/espace/scae/edit_project&id={project_id}"
    req = session.get(url)
    req.raise_for_status()
    soup = BeautifulSoup(req.text, features="lxml")
    try:
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
        
        stab_path = Path("cache")/Path(f"{project_id}_{last_stab_text.replace(' ', '')}")
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
            
            os.system(f"convert.vbs \"{stab_path.absolute()}\" \"{new_path}\"")
            
    #except AttributeError as e:
    #    print(e)
    #    return None
    finally:
        pass

    return output

def main():
    projects = get_projects()
    projects_details = list()
    missing_projects = list()
               
    for project in tqdm(projects):
        project_details = get_project_details(project["id"])

        if project_details is not None:
            projects_details.append(project_details)
        else:
            print(f"Error on project {project['name']}")
            missing_projects.append(project)
            projects.remove(project)
            
    with open(Path("cache")/Path("project_list.json"), "w", encoding="utf-8") as f:
        json.dump({"project_details":projects_details, "project_list": projects, "missing_projects": missing_projects}, f, ensure_ascii=False)
        

if __name__ == "__main__":
    main()