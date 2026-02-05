# SCAE Message Scraper

This Python script scrapes messages from the SCAE (Spatial Activities Engineering Center) project manager website. It extracts project messages and saves them to JSON files with comprehensive metadata.

## Features

- **API Integration**: Lists projects using the SCAE API
- **Configurable Filters**: Filter projects by launch year, type (minif/fusex), and status
- **HTML Download**: Downloads project pages using authenticated PHPSESSID cookie
- **Message Extraction**: Parses HTML to extract messages (RCE3 messages, notes, comments)
- **JSON Output**: Saves one JSON file per project with metadata and messages
- **Caching**: Optionally cache downloaded HTML to avoid redundant requests
- **Progress Tracking**: Shows real-time progress with tqdm progress bars
- **Error Handling**: Comprehensive error handling with detailed error reporting

## Installation

1. Install required dependencies:
```bash
pip install -r requirements.txt
```

## Configuration

1. Create a configuration file based on the example:
```bash
cp config_messages_example.json config_messages.json
```

2. Edit `config_messages.json` with your credentials and preferences:

```json
{
    "api_key": "YOUR_SCAE_API_KEY",
    "cookies": "PHPSESSID=YOUR_SESSION_ID",
    "filters": {
        "launch_year": 2022,
        "types": ["minif", "fusex"],
        "status": "wip"
    },
    "output_dir": "messages",
    "html_cache_dir": "cache/html",
    "use_cache": true
}
```

### Configuration Options

- **api_key**: Your SCAE API key for accessing the project list
- **cookies**: PHPSESSID cookie value for authenticated access to project pages
- **filters**: Project filtering criteria
  - **launch_year**: Filter by launch year (e.g., 2022)
  - **types**: List of project types to include (e.g., ["minif", "fusex"])
  - **status**: Project status (e.g., "wip" for work in progress)
- **output_dir**: Directory where JSON files will be saved (default: "messages")
- **html_cache_dir**: Directory for caching downloaded HTML (default: "cache/html")
- **use_cache**: Whether to use cached HTML files (default: true)

## Usage

Run the scraper with default configuration:
```bash
python scrape_messages.py
```

Or specify a custom configuration file:
```bash
python scrape_messages.py --config my_config.json
```

## Output Structure

The script creates one JSON file per project with the following structure:

```json
{
  "metadata": {
    "project_id": 123,
    "project_name": "Example Project",
    "club_name": "Example Club",
    "campaign": "C'Space",
    "project_type": "fusex",
    "status": "wip",
    "launch_year": "2022",
    "scraped_at": "2026-02-05T21:00:00.000000",
    "api_data": {
      // Original API response data
    }
  },
  "messages": [
    {
      "type": "rce3",
      "content": "Message content here...",
      "html": "<div class='border-rce3'>...</div>"
    },
    {
      "type": "note",
      "content": "Note content here...",
      "html": "<textarea>...</textarea>"
    }
  ],
  "message_count": 2
}
```

### Message Types

- **rce3**: Messages with the `.border-rce3` CSS class
- **general**: Messages found in general message/comment containers
- **note**: Project notes from textarea fields

## Output Files

After running the scraper, you'll find:

1. **Individual project files**: `messages/{project_id}_{project_name}.json`
2. **Summary file**: `messages/scraping_summary.json` containing statistics and any errors

## Future Use Cases

This script is designed to prepare data for LLM analysis using llama-cpp. The structured JSON output can be used to:

- Identify projects that need responses
- Suggest action plans to help projects
- Analyze communication patterns
- Track project progress through messages

## Integration with Existing Scripts

This script follows the same patterns as `get_stabs.py`:
- Uses the SCAE API for project listing
- Authenticates with PHPSESSID cookie
- Filters projects based on configuration
- Uses BeautifulSoup for HTML parsing
- Includes error handling and progress tracking

## Troubleshooting

### Authentication Issues

If you encounter authentication errors:
1. Ensure your PHPSESSID cookie is valid
2. Check that your API key is correct
3. Verify you have access to the projects you're trying to scrape

### Missing Messages

If messages are not being extracted:
1. Check the HTML structure of the project page
2. Update the CSS selectors in `extract_messages()` method if needed
3. Examine cached HTML files in `cache/html/` for debugging

## License

This script is part of the Plasci-Stabtraj-SVG project.
