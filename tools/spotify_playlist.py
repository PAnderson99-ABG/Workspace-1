import csv
import spotipy
from spotipy.oauth2 import SpotifyClientCredentials
from dotenv import load_dotenv
import os

load_dotenv()

def clean_env(var_name):
    val = os.getenv(var_name, "")
    return val.strip().strip("'").strip('"')

# === CONFIGURATION ===
CLIENT_ID = clean_env("SPOTIPY_CLIENT_ID")
CLIENT_SECRET = clean_env("SPOTIPY_CLIENT_SECRET")
PLAYLIST_URL = 'https://open.spotify.com/playlist/1A0eDtjZCm7mcGJHRTVMLT?si=NOyxZ6gQTKKMjdvJOsSYmA&pi=8RXXpgNXSrqL3'

OUTPUT_FILE = 'spotify_playlist_export.csv'


def extract_playlist_id(url):
    """Extract playlist ID from the URL."""
    return url.split("/")[-1].split("?")[0]


def get_playlist_tracks(sp, playlist_id):
    """Fetch all tracks from the playlist."""
    results = []
    offset = 0

    while True:
        response = sp.playlist_items(playlist_id, offset=offset, fields='items.track.name,items.track.artists,next', additional_types=['track'])
        items = response.get('items', [])

        for item in items:
            track = item.get('track')
            if track:
                name = track['name']
                artists = ", ".join(artist['name'] for artist in track['artists'])
                results.append((name, artists))

        if response['next']:
            offset += len(items)
        else:
            break

    return results


def write_to_csv(data, filename):
    """Write song/artist data to CSV."""
    with open(filename, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(['Track Name', 'Artists'])
        writer.writerows(data)


def main():
    playlist_id = extract_playlist_id(PLAYLIST_URL)

    auth_manager = SpotifyClientCredentials(client_id=CLIENT_ID, client_secret=CLIENT_SECRET)
    sp = spotipy.Spotify(auth_manager=auth_manager)

    print("Fetching playlist data...")
    tracks = get_playlist_tracks(sp, playlist_id)
    print(f"Found {len(tracks)} tracks.")

    write_to_csv(tracks, OUTPUT_FILE)
    print(f"Exported to {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
