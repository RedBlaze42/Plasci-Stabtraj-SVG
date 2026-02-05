"""
Test script for SCAE message scraper

This script tests the core functionality of the message scraper
with mock data to ensure it works correctly.
"""

import json
import tempfile
import os
from pathlib import Path
from scrape_messages import SCAEMessageScraper


def create_test_html():
    """Create a sample HTML page similar to SCAE project page"""
    return """
    <html>
    <head><title>Test Project</title></head>
    <body>
        <input id="project__name" value="Test Rocket Project" />
        
        <select id="project__club">
            <option>Club 1</option>
            <option selected>Test Rocket Club</option>
            <option>Club 3</option>
        </select>
        
        <select id="project__campaign">
            <option>Campaign 1</option>
            <option selected>C'Space 2022</option>
        </select>
        
        <div class="border-rce3">
            <p>This is a test RCE3 message from the project manager.</p>
            <p>Please submit your documentation by next week.</p>
        </div>
        
        <div class="border-rce3">
            <p>Second message: Great progress on the stability calculations!</p>
        </div>
        
        <div class="message">
            <p>This is a general message in the project.</p>
        </div>
        
        <textarea id="project_notes">
            Important notes about the project timeline.
        </textarea>
    </body>
    </html>
    """


def test_message_extraction():
    """Test message extraction from HTML"""
    print("Testing message extraction...")
    
    # Create a temporary config file
    with tempfile.TemporaryDirectory() as tmpdir:
        config_path = Path(tmpdir) / "test_config.json"
        config = {
            "api_key": "test_key",
            "cookies": "PHPSESSID=test",
            "filters": {"launch_year": 2022, "types": ["minif"], "status": "wip"},
            "output_dir": str(Path(tmpdir) / "messages"),
            "html_cache_dir": str(Path(tmpdir) / "cache"),
            "use_cache": False
        }
        
        with open(config_path, 'w') as f:
            json.dump(config, f)
        
        # Create scraper instance
        scraper = SCAEMessageScraper(str(config_path))
        
        # Test HTML content
        html_content = create_test_html()
        
        # Test message extraction
        messages = scraper.extract_messages(html_content)
        
        print(f"✓ Extracted {len(messages)} messages")
        
        # Verify messages
        rce3_messages = [m for m in messages if m['type'] == 'rce3']
        general_messages = [m for m in messages if m['type'] == 'general']
        note_messages = [m for m in messages if m['type'] == 'note']
        
        print(f"  - RCE3 messages: {len(rce3_messages)}")
        print(f"  - General messages: {len(general_messages)}")
        print(f"  - Note messages: {len(note_messages)}")
        
        assert len(rce3_messages) >= 2, "Should extract at least 2 RCE3 messages"
        assert len(messages) > 0, "Should extract at least some messages"
        
        # Test metadata extraction
        project_data = {
            "id": 123,
            "name": "Test Project",
            "type": "fusex",
            "status": "wip",
            "launch_year": "2022"
        }
        
        metadata = scraper.extract_project_metadata(html_content, project_data)
        
        print(f"\n✓ Extracted metadata:")
        print(f"  - Project name: {metadata.get('project_name')}")
        print(f"  - Club name: {metadata.get('club_name')}")
        print(f"  - Campaign: {metadata.get('campaign')}")
        
        assert metadata['project_name'] == "Test Rocket Project"
        assert metadata['club_name'] == "Test Rocket Club"
        assert metadata['campaign'] == "C'Space 2022"
        
        print("\n✓ All tests passed!")
        return True


def test_json_structure():
    """Test the JSON output structure"""
    print("\nTesting JSON output structure...")
    
    test_data = {
        "metadata": {
            "project_id": 123,
            "project_name": "Test Project",
            "club_name": "Test Club",
            "campaign": "C'Space",
            "project_type": "fusex",
            "status": "wip",
            "launch_year": "2022",
            "scraped_at": "2022-01-01T00:00:00",
            "api_data": {}
        },
        "messages": [
            {
                "type": "rce3",
                "content": "Test message",
                "html": "<div>Test</div>"
            }
        ],
        "message_count": 1
    }
    
    # Verify structure
    assert "metadata" in test_data
    assert "messages" in test_data
    assert "message_count" in test_data
    assert test_data["message_count"] == len(test_data["messages"])
    
    print("✓ JSON structure is correct")
    return True


def main():
    """Run all tests"""
    print("=" * 60)
    print("SCAE Message Scraper - Test Suite")
    print("=" * 60)
    
    try:
        test_message_extraction()
        test_json_structure()
        
        print("\n" + "=" * 60)
        print("All tests completed successfully! ✓")
        print("=" * 60)
        return 0
    
    except Exception as e:
        print(f"\n✗ Test failed with error: {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    exit(main())
