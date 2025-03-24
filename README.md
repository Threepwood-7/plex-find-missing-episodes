# plex-find-missing-episodes
A sloppy attempt to list missing episodes from your Plex library using Plex and TVDB APIs. Using AI tools is actually more time consuming than coding the thing yourself, but it was worth a try.
Made some refactoring, further changes will be done manually with no AI.
Before anyone points it out, yes I am aware and use the ARR stack. This tool is to report on media that is managed manually, or if you want to have a quick report of your entire library.

## Installation
Just run `pip install -r requirements.txt` or `S01_install_reqs.cmd`

## Configuration
For now, it's all hard-coded in `pl_report_missing_episodes_claude.py`. Ensure to change:

```
# Configuration - Replace these with your actual values
PLEX_URL = "http://localhost:32400"  # Change to your Plex server URL
PLEX_TOKEN = ""  # Your Plex authentication token
TVDB_APIKEY = ""  # Your TVDB API key
CACHE_EXPIRY_DAYS = 30  # how long to cache the TVDB data
LIBRARY_TITLE_FILTER = re.compile("TV .*", re.IGNORECASE)  # filter show libraries that match the regex here
```

## Running it
Just launch the py script or `S05_launch.cmd`. Let it run. Errors would be redirected to the `stderr` log file.

Missing episodes would be identified by the columns:
- Is Plex Missing (Episode)
- Is Plex Missing (Season)

So if you are looking for seasons that are on disk but you are missing episodes, set the Excel filter to:
- Is Plex Missing (Episode) : true
- Is Plex Missing (Season) : false

In addition, the report lists episodes for which you have duplicate files, which might be intended or not:
- Is Plex Duplicate
