#!/usr/bin/env python3
import os
import json
import sys
import signal
import time
import unicodedata
import re
from datetime import datetime, timedelta
from plexapi.server import PlexServer
import tvdb_v4_official
import xlsxwriter
from pathlib import Path

# Configuration - Replace these with your actual values
PLEX_URL = "http://localhost:32400"  # Change to your Plex server URL
PLEX_TOKEN = ""  # Your Plex authentication token
TVDB_APIKEY = ""  # Your TVDB API key

# Cache settings
CACHE_DIR = "./cache"
CACHE_EXPIRY_DAYS = 14

# Ensure cache directory exists
os.makedirs(CACHE_DIR, exist_ok=True)

# Initialize workbook
wb = xlsxwriter.Workbook("plex-episodes-report.xlsx")
main_sheet = wb.add_worksheet("Episodes")
not_found_sheet = wb.add_worksheet("TVNTF")
error_sheet = wb.add_worksheet("TVERR")

# Define formats - only header is bold now
header_format = wb.add_format({"bold": True})

# Flag to handle graceful termination
terminate = False


def signal_handler(sig, frame):
    """Handle Ctrl+C to stop processing but still output the report."""
    global terminate
    safe_print("Interrupt received, finishing current show and generating report...")
    terminate = True


signal.signal(signal.SIGINT, signal_handler)


def safe_print(text):
    """Print text to console with Unicode characters converted to printable form"""
    if isinstance(text, str):
        # Replace or remove unprintable characters
        text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    print(text)


def get_cache_filename(tvdb_id):
    """Generate a cache filename based on the TVDB ID"""
    return os.path.join(CACHE_DIR, f"tvdb_{tvdb_id}.json")


def is_cache_valid(cache_file):
    """Check if the cache file exists and is less than CACHE_EXPIRY_DAYS old"""
    if not os.path.exists(cache_file):
        return False

    file_time = datetime.fromtimestamp(os.path.getmtime(cache_file))
    expiry_date = datetime.now() - timedelta(days=CACHE_EXPIRY_DAYS)

    return file_time > expiry_date


def extract_tvdb_id(guids):
    """Extract TVDB ID from Plex guid list"""
    for guid in guids:
        if "tvdb://" in guid.id:
            return guid.id.split("//")[1]
    return None


def setup_worksheets():
    """Set up the worksheets with headers and formatting"""
    # Main sheet headers
    headers = [
        "Plex Library Title",
        "TV Show Title",
        "TV Show Year",
        "Number of Seasons",
        "Season Number",
        "Season Title",
        "Number of Episodes",
        "Episode Number",
        "Episode Aired Date",
        "Is Plex Missing",
        "Is Plex Duplicate",
        "Episode Title",
        "File on Disk",
    ]

    for col, header in enumerate(headers):
        main_sheet.write(0, col, header, header_format)

    # Configure filters and freeze panes
    main_sheet.autofilter(0, 0, 0, len(headers) - 1)
    main_sheet.freeze_panes(1, 0)

    # Setup error sheets
    not_found_sheet.write(0, 0, "Plex Library Title", header_format)
    not_found_sheet.write(0, 1, "TV Show Title", header_format)
    not_found_sheet.write(0, 2, "TV Show Year", header_format)
    not_found_sheet.write(0, 3, "Error Details", header_format)

    error_sheet.write(0, 0, "Plex Library Title", header_format)
    error_sheet.write(0, 1, "TV Show Title", header_format)
    error_sheet.write(0, 2, "TV Show Year", header_format)
    error_sheet.write(0, 3, "Error Details", header_format)

    # Configure filters and freeze panes for error sheets
    not_found_sheet.autofilter(0, 0, 0, 3)
    not_found_sheet.freeze_panes(1, 0)
    error_sheet.autofilter(0, 0, 0, 3)
    error_sheet.freeze_panes(1, 0)


def get_tvdb_data(tvdb_client, show_title, show_year, tvdb_id=None, plex_library=None):
    """Get TV show data from TVDB, either by ID or search"""
    safe_print(f"Processing {show_title} ({show_year}) from library '{plex_library}'")

    if not tvdb_id:
        # Search by title
        safe_print(f"TVDB ID not found in Plex, searching by title: {show_title}")
        try:
            search_results = tvdb_client.search(show_title, type="series")
            if not search_results:
                safe_print(f"No results found on TVDB for {show_title}")
                return None, "Not found on TVDB"

            # Try to find the best match considering year
            best_match = None
            for result in search_results:
                if show_year and str(result.get("year", "")) == str(show_year):
                    best_match = result
                    break

            # If no year match, use the first result
            if not best_match and search_results:
                best_match = search_results[0]

            if best_match:
                tvdb_id = best_match.get("tvdb_id")
                safe_print(f"Found TVDB match: {best_match.get('name')} (ID: {tvdb_id})")
            else:
                safe_print(f"No suitable match found on TVDB for {show_title}")
                return None, "No suitable match found"
        except Exception as e:
            error_msg = f"TVDB search error: {str(e)}"
            safe_print(error_msg)
            return None, error_msg

    cache_file = get_cache_filename(tvdb_id)

    if is_cache_valid(cache_file):
        safe_print(f"Using cached data for {show_title} (TVDB ID: {tvdb_id})")
        try:
            with open(cache_file, "r", encoding="utf-8") as f:
                show_data = json.load(f)
            return show_data, None
        except Exception as e:
            safe_print(f"Error reading cache: {str(e)}")
            # Fall through to fetch new data

    # Fetch new data from TVDB
    try:
        safe_print(f"Fetching series extended data for {show_title} (TVDB ID: {tvdb_id}) from TVDB API")
        series_data = tvdb_client.get_series_extended(tvdb_id)

        # Get details for each season
        show_data = {"series": series_data, "seasons": []}

        for season in series_data.get("seasons", []):
            if season["type"] and season["type"]["type"] and season["type"]["type"] != "official":
                print(f"Skipping details, season type = '{season['type']['type']}' ..")
                continue

            season_num = season.get("number")
            safe_print(f"Fetching episodes for Season {season_num} of {show_title} (TVDB ID: {tvdb_id})")
            try:
                season_extended = tvdb_client.get_season_extended(season.get("id"))
                show_data["seasons"].append(season_extended)
            except Exception as e:
                safe_print(f"Error fetching season {season_num} data: {str(e)}")

        # Save to cache
        with open(cache_file, "w", encoding="utf-8") as f:
            json.dump(show_data, f)

        return show_data, None
    except Exception as e:
        error_msg = f"TVDB API error: {str(e)}"
        safe_print(error_msg)
        return None, error_msg


def process_show(show, plex_library_title, tvdb_client, row_index):
    """Process a single TV show and update the report"""
    show_title = show.title
    show_year = getattr(show, "year", "")

    safe_print(f"\nProcessing show: {show_title} ({show_year}) from library '{plex_library_title}'")

    # Extract TVDB ID from Plex
    tvdb_id = None
    if hasattr(show, "guids") and show.guids:
        tvdb_id = extract_tvdb_id(show.guids)
        if tvdb_id:
            safe_print(f"Found TVDB ID in Plex: {tvdb_id}")

    # Get TVDB data
    tvdb_data, error = get_tvdb_data(tvdb_client, show_title, show_year, tvdb_id, plex_library_title)

    if error:
        # Add to error sheet
        if "Not found" in error:
            sheet = not_found_sheet
            sheet_name = "TVNTF"
        else:
            sheet = error_sheet
            sheet_name = "TVERR"

        # Find the next empty row in the error sheet
        error_row = 1
        # while sheet.get_cell(error_row, 0).value is not None:
        #    error_row += 1

        sheet.write_string(error_row, 0, plex_library_title)
        sheet.write_string(error_row, 1, show_title)
        sheet.write(error_row, 2, show_year)
        sheet.write_string(error_row, 3, error)

        safe_print(f"Added {show_title} to {sheet_name} sheet due to error: {error}")
        return row_index

    # Get Plex episodes
    plex_episodes = {}
    plex_episodes_count = {}

    try:
        # Note: No .refresh() call as specified in requirements
        safe_print(f"Fetching episodes for {show_title} from Plex API")
        episodes = show.episodes()

        for episode in episodes:
            season_num = episode.seasonNumber
            episode_num = episode.index

            if season_num not in plex_episodes:
                plex_episodes[season_num] = {}
                plex_episodes_count[season_num] = {}

            for media_part in episode.iterParts():
                episode_file_path = str(Path(media_part.file.replace("\\?\\", "")).absolute())

                # Track episodes for duplicate detection
                if episode_num not in plex_episodes_count[season_num]:
                    plex_episodes_count[season_num][episode_num] = 1
                else:
                    plex_episodes_count[season_num][episode_num] += 1

                # If this is a duplicate, append to existing entry
                if episode_num in plex_episodes[season_num]:
                    existing_ep = plex_episodes[season_num][episode_num]

                    # Combine file paths
                    if hasattr(episode, "locations") and episode.locations:
                        if hasattr(existing_ep, "combined_locations"):
                            existing_ep.combined_locations.extend([episode_file_path])
                        else:
                            existing_ep.combined_locations = [episode_file_path]
                else:
                    # Store the episode
                    plex_episodes[season_num][episode_num] = episode

                    # Initialize combined_locations if needed
                    if hasattr(episode, "locations") and episode.locations:
                        episode.combined_locations = [episode_file_path]

        safe_print(f"Finished fetching episodes for {show_title} from Plex API")
    except Exception as e:
        safe_print(f"Error fetching Plex episodes: {str(e)}")

    # Process TVDB data and compare with Plex
    series_data = tvdb_data.get("series", {})
    tvdb_seasons = tvdb_data.get("seasons", [])

    # Count total seasons from TVDB
    official_seasons = [s for s in series_data.get("seasons", []) if s.get("type") is None or s.get("type") == "official"]
    num_seasons = len(official_seasons)

    # Process each season from TVDB
    for season_data in tvdb_seasons:
        season_num = season_data.get("number")
        season_name = season_data.get("name", "")
        episodes = season_data.get("episodes", [])

        if not episodes:
            safe_print(f"Skipping Season {season_num} of {show_title} - 'episodes' is missing!")
            continue

        safe_print(f"Processing Season {season_num} of {show_title} ({len(episodes)} episodes)")

        for episode in episodes:
            episode_num = episode.get("number")
            episode_title = episode.get("name", "")

            # Try to get air date
            air_date = episode.get("aired")
            if not air_date:
                air_date = ""

            # Check if this episode exists in Plex
            is_missing = True
            is_duplicate = False
            file_path = ""

            if season_num in plex_episodes and episode_num in plex_episodes[season_num]:
                is_missing = False
                plex_episode = plex_episodes[season_num][episode_num]

                # Check for duplicates
                if season_num in plex_episodes_count and episode_num in plex_episodes_count[season_num]:
                    is_duplicate = plex_episodes_count[season_num][episode_num] > 1

                # Get file path(s)
                if hasattr(plex_episode, "combined_locations") and plex_episode.combined_locations:
                    file_path = "\n".join(plex_episode.combined_locations)
                elif hasattr(plex_episode, "locations") and plex_episode.locations:  # this should never be ?
                    file_path = "\n".join(plex_episode.locations)

            # Write to the main sheet using appropriate write methods
            main_sheet.write_string(row_index, 0, plex_library_title)
            main_sheet.write_string(row_index, 1, show_title)
            main_sheet.write(row_index, 2, show_year)
            main_sheet.write(row_index, 3, num_seasons)
            main_sheet.write(row_index, 4, season_num)
            main_sheet.write_string(row_index, 5, season_name)
            main_sheet.write(row_index, 6, len(episodes))
            main_sheet.write(row_index, 7, episode_num)
            main_sheet.write(row_index, 8, air_date)
            main_sheet.write_boolean(row_index, 9, is_missing)
            main_sheet.write_boolean(row_index, 10, is_duplicate)
            main_sheet.write_string(row_index, 11, episode_title)
            main_sheet.write_string(row_index, 12, file_path)

            row_index += 1

    return row_index


def main():
    """Main function to generate the Plex episodes report"""
    safe_print("Starting Plex TV Shows Episode Report Generator")

    # Set up worksheets
    setup_worksheets()

    try:
        # Connect to Plex
        safe_print(f"Connecting to Plex server at {PLEX_URL}")
        plex = PlexServer(PLEX_URL, PLEX_TOKEN)
        safe_print(f"Successfully connected to Plex server")

        # Connect to TVDB
        safe_print("Connecting to TVDB API")
        tvdb = tvdb_v4_official.TVDB(TVDB_APIKEY)
        safe_print("Successfully connected to TVDB API")

        # Get all TV libraries
        safe_print("Fetching Plex libraries")
        libraries = [section for section in plex.library.sections() if section.type == "show"]
        safe_print(f"Finished fetching Plex libraries")

        if not libraries:
            safe_print("No TV libraries found that start with 'TV '")
            wb.close()
            return

        safe_print(f"Found {len(libraries)} TV libraries to process")

        row_index = 1  # Start after headers

        # Process each library
        for library in libraries:
            library_title = library.title
            safe_print(f"\nProcessing library: {library_title}")

            # Get all shows in this library
            safe_print(f"Fetching all shows from library: {library_title}")
            shows = library.search()
            safe_print(f"Finished fetching shows from library: {library_title}")

            safe_print(f"Found {len(shows)} shows in library: {library_title}")

            for show in shows:
                if terminate:
                    safe_print("Terminating early due to user interrupt")
                    break

                try:
                    row_index = process_show(show, library_title, tvdb, row_index)
                except Exception as e1:
                    safe_print(f"Error (I): {str(e1)}")
                    print(e1)

            if terminate:
                break

    except Exception as e:
        safe_print(f"Error (O): {str(e)}")
        print(e)

    finally:
        # Save the workbook
        safe_print("Saving report to plex-episodes-report.xlsx")
        wb.close()
        safe_print("Report generation complete")


if __name__ == "__main__":
    main()
