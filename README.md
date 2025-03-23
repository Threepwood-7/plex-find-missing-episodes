# plex-find-missing-episodes
A sloppy attempt to list missing episodes from your Plex library using Plex and TVDB APIs. Using AI tools is actually more time consuming than coding the thing yourself, but it was worth a try.

## Installation
Just run `pip install -r requirements.txt` or `S01_install_reqs.cmd`

## Configuration
For now, it's all hard-coded in `pl_report_missing_episodes_claude.py`. Ensure to change

```
# Configuration - Replace these with your actual values
PLEX_URL = "http://localhost:32400"  # Change to your Plex server URL
PLEX_TOKEN = ""  # Your Plex authentication token
TVDB_APIKEY = ""  # Your TVDB API key
```

## Running it
Just launch the py script or `S05_launch.cmd`. Let it run. Errors would be redirected to the `stderr` log file.
