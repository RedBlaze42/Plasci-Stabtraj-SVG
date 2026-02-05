"""
SCAE Message Scraper

This script scrapes messages from SCAE project manager pages.
It downloads project HTML pages, extracts messages, and saves them to JSON files
with project metadata.

Usage:
    python scrape_messages.py

Configuration:
    Create a config_messages.json file based on config_messages_example.json
"""

import requests
import json
import os
from pathlib import Path
from bs4 import BeautifulSoup
from tqdm import tqdm
from sanitize_filename import sanitize
import argparse
from datetime import datetime


class SCAEMessageScraper:
    """Scraper for SCAE project manager messages"""
    
    def __init__(self, config_path="config_messages.json"):
        """Initialize the scraper with configuration"""
        with open(config_path) as f:
            self.config = json.load(f)
        
        # Create output directories
        os.makedirs(self.config.get("output_dir", "messages"), exist_ok=True)
        os.makedirs(self.config.get("html_cache_dir", "cache/html"), exist_ok=True)
        
        # Setup session with authentication
        self.session = requests.Session()
        self.session.headers = {
            "Cookie": self.config["cookies"],
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/102.0.0.0 Safari/537.36"
        }
    
    def get_projects(self):
        """Fetch and filter projects from SCAE API"""
        url = f"https://www.planete-sciences.org/espace/scae/index.php?p=api&key={self.config['api_key']}"
        req = self.session.get(url)
        req.raise_for_status()
        
        projects = []
        filters = self.config.get("filters", {})
        
        for project in req.json():
            # Apply filters
            if filters.get("launch_year") and project.get("launch_year"):
                if int(project["launch_year"]) != int(filters["launch_year"]):
                    continue
            
            if filters.get("types") and project.get("type"):
                if project["type"] not in filters["types"]:
                    continue
            
            if filters.get("status") and project.get("status"):
                if project["status"] != filters["status"]:
                    continue
            
            projects.append(project)
        
        return projects
    
    def get_project_url(self, project_id):
        """Generate project page URL"""
        return f"https://www.planete-sciences.org/espace/scae/edit_project&id={project_id}"
    
    def download_project_html(self, project_id):
        """Download HTML page for a project"""
        url = self.get_project_url(project_id)
        html_cache_path = Path(self.config.get("html_cache_dir", "cache/html")) / f"{project_id}.html"
        
        # Check cache if enabled
        if self.config.get("use_cache", True) and html_cache_path.exists():
            with open(html_cache_path, "r", encoding="utf-8") as f:
                return f.read()
        
        # Download HTML
        req = self.session.get(url)
        req.raise_for_status()
        
        # Save to cache
        with open(html_cache_path, "w", encoding="utf-8") as f:
            f.write(req.text)
        
        return req.text
    
    def extract_messages(self, html_content):
        """Extract messages from project HTML page"""
        soup = BeautifulSoup(html_content, features="lxml")
        messages = []
        
        # Extract messages with .border-rce3 class (RCE3 messages)
        rce3_messages = soup.select(".border-rce3")
        for msg in rce3_messages:
            message_data = {
                "type": "rce3",
                "content": msg.get_text(strip=True),
                "html": str(msg)
            }
            messages.append(message_data)
        
        # Look for other message containers
        # Check for general message containers
        message_containers = soup.select(".message, .comment, .note")
        for container in message_containers:
            if container not in rce3_messages:  # Avoid duplicates
                message_data = {
                    "type": "general",
                    "content": container.get_text(strip=True),
                    "html": str(container)
                }
                messages.append(message_data)
        
        # Check for project notes or comments sections
        notes_section = soup.find("textarea", {"id": lambda x: x and "note" in x.lower()})
        if notes_section and notes_section.get_text(strip=True):
            messages.append({
                "type": "note",
                "content": notes_section.get_text(strip=True),
                "html": str(notes_section)
            })
        
        return messages
    
    def extract_project_metadata(self, html_content, project_api_data):
        """Extract metadata from project page"""
        soup = BeautifulSoup(html_content, features="lxml")
        metadata = {
            "project_id": project_api_data.get("id"),
            "api_data": project_api_data,
            "scraped_at": datetime.now().isoformat()
        }
        
        try:
            # Extract project name
            name_input = soup.find("input", {"id": "project__name"})
            if name_input:
                metadata["project_name"] = name_input.get("value", "")
            
            # Extract club name
            club_select = soup.find("select", {"id": "project__club"})
            if club_select:
                selected_option = club_select.find("option", {"selected": True})
                if selected_option:
                    metadata["club_name"] = selected_option.get_text(strip=True)
            
            # Extract campaign
            campaign_select = soup.find("select", {"id": "project__campaign"})
            if campaign_select:
                selected_option = campaign_select.find("option", {"selected": True})
                if selected_option:
                    metadata["campaign"] = selected_option.get_text(strip=True)
            
            # Extract project type
            metadata["project_type"] = project_api_data.get("type")
            metadata["status"] = project_api_data.get("status")
            metadata["launch_year"] = project_api_data.get("launch_year")
            
        except Exception as e:
            print(f"Error extracting metadata for project {project_api_data.get('id')}: {e}")
        
        return metadata
    
    def scrape_project_messages(self, project):
        """Scrape messages from a single project"""
        try:
            project_id = project["id"]
            
            # Download HTML
            html_content = self.download_project_html(project_id)
            
            # Extract messages
            messages = self.extract_messages(html_content)
            
            # Extract metadata
            metadata = self.extract_project_metadata(html_content, project)
            
            # Prepare output
            output = {
                "metadata": metadata,
                "messages": messages,
                "message_count": len(messages)
            }
            
            return output, None
        
        except Exception as e:
            return None, str(e)
    
    def save_project_json(self, project_data, project_id):
        """Save project data to JSON file"""
        output_dir = Path(self.config.get("output_dir", "messages"))
        project_name = project_data["metadata"].get("project_name", f"project_{project_id}")
        
        # Sanitize filename
        filename = sanitize(f"{project_id}_{project_name}.json")
        output_path = output_dir / filename
        
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(project_data, f, ensure_ascii=False, indent=2)
        
        return output_path
    
    def scrape_all_projects(self):
        """Main method to scrape all projects"""
        print("Fetching project list from SCAE API...")
        projects = self.get_projects()
        print(f"Found {len(projects)} projects matching filters")
        
        results = {
            "successful": [],
            "failed": []
        }
        
        progress_bar = tqdm(total=len(projects), desc="Scraping project messages")
        
        for project in projects:
            project_id = project["id"]
            project_name = project.get("name", f"Project {project_id}")
            
            progress_bar.set_postfix({"current": project_name[:30]})
            
            project_data, error = self.scrape_project_messages(project)
            
            if project_data:
                # Save to JSON
                output_path = self.save_project_json(project_data, project_id)
                results["successful"].append({
                    "project_id": project_id,
                    "project_name": project_name,
                    "output_path": str(output_path),
                    "message_count": project_data["message_count"]
                })
            else:
                results["failed"].append({
                    "project_id": project_id,
                    "project_name": project_name,
                    "error": error
                })
            
            progress_bar.update(1)
        
        progress_bar.close()
        
        # Print summary
        print(f"\n{'='*60}")
        print(f"Scraping completed!")
        print(f"{'='*60}")
        print(f"Successful: {len(results['successful'])}")
        print(f"Failed: {len(results['failed'])}")
        
        if results['failed']:
            print("\nFailed projects:")
            for failed in results['failed']:
                print(f"  - {failed['project_name']} (ID: {failed['project_id']})")
                print(f"    Error: {failed['error']}")
        
        # Save summary
        summary_path = Path(self.config.get("output_dir", "messages")) / "scraping_summary.json"
        with open(summary_path, "w", encoding="utf-8") as f:
            json.dump(results, f, ensure_ascii=False, indent=2)
        
        print(f"\nSummary saved to: {summary_path}")
        
        return results


def main():
    """Main entry point"""
    parser = argparse.ArgumentParser(description="Scrape messages from SCAE project manager")
    parser.add_argument(
        "--config",
        default="config_messages.json",
        help="Path to configuration file (default: config_messages.json)"
    )
    
    args = parser.parse_args()
    
    # Check if config exists
    if not os.path.exists(args.config):
        print(f"Error: Configuration file '{args.config}' not found!")
        print("Please create it based on config_messages_example.json")
        return 1
    
    # Run scraper
    scraper = SCAEMessageScraper(config_path=args.config)
    scraper.scrape_all_projects()
    
    return 0


if __name__ == "__main__":
    exit(main())
