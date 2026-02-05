#!/usr/bin/env python3
"""
Demo script for SCAE Message Scraper

This script demonstrates how to use the SCAE message scraper without
actually connecting to the API. It shows the expected workflow and output.
"""

import json
from pathlib import Path


def show_demo():
    """Show a demonstration of the scraper workflow"""
    
    print("=" * 70)
    print("SCAE Message Scraper - Demo")
    print("=" * 70)
    print()
    
    print("Step 1: Configuration")
    print("-" * 70)
    print("Create config_messages.json with your credentials:")
    print()
    
    config_example = {
        "api_key": "YOUR_SCAE_API_KEY",
        "cookies": "PHPSESSID=YOUR_SESSION_ID",
        "filters": {
            "launch_year": 2022,
            "types": ["minif", "fusex"],
            "status": "wip"
        },
        "output_dir": "messages",
        "html_cache_dir": "cache/html",
        "use_cache": True
    }
    
    print(json.dumps(config_example, indent=2))
    print()
    
    print("Step 2: Run the scraper")
    print("-" * 70)
    print("  $ python scrape_messages.py")
    print()
    print("Or with a custom config:")
    print("  $ python scrape_messages.py --config my_config.json")
    print()
    
    print("Step 3: Expected output")
    print("-" * 70)
    print("The scraper will:")
    print("  1. Fetch project list from SCAE API")
    print("  2. Filter projects based on your criteria")
    print("  3. Download HTML pages for each project")
    print("  4. Extract messages from the pages")
    print("  5. Save JSON files with project data and messages")
    print()
    
    print("Output example:")
    print()
    
    output_example = {
        "metadata": {
            "project_id": 123,
            "project_name": "Phoenix Rocket",
            "club_name": "Rocket Science Club",
            "campaign": "C'Space",
            "project_type": "fusex",
            "status": "wip",
            "launch_year": "2022",
            "scraped_at": "2022-06-15T14:30:00.000000"
        },
        "messages": [
            {
                "type": "rce3",
                "content": "Please submit your stability calculations by next week.",
                "html": "<div class='border-rce3'>...</div>"
            },
            {
                "type": "rce3",
                "content": "Great work on the motor selection!",
                "html": "<div class='border-rce3'>...</div>"
            }
        ],
        "message_count": 2
    }
    
    print(json.dumps(output_example, indent=2))
    print()
    
    print("Step 4: Output files")
    print("-" * 70)
    print("After running, you'll find:")
    print("  - messages/123_Phoenix_Rocket.json (one per project)")
    print("  - messages/scraping_summary.json (overall summary)")
    print("  - cache/html/123.html (cached HTML pages)")
    print()
    
    print("Step 5: Next steps (Future: LLM integration)")
    print("-" * 70)
    print("The JSON files can be used with llama-cpp to:")
    print("  - Identify projects that need responses")
    print("  - Suggest action plans to help projects")
    print("  - Analyze communication patterns")
    print("  - Track project progress")
    print()
    
    print("=" * 70)
    print("For more information, see README_MESSAGES.md")
    print("=" * 70)


if __name__ == "__main__":
    show_demo()
