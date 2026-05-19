"""Download official and community FIT Profile.xlsx files.

Sources:
  - Official: https://github.com/garmin/fit-sdk-tools/raw/main/Profile.xlsx
  - Community: Google Sheet exported as xlsx (undocumented types/messages)
"""

import os
import sys
import requests

OFFICIAL_URL = (
    "https://github.com/garmin/fit-sdk-tools/raw/main/Profile.xlsx"
)
COMMUNITY_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1x34eRAZ45nbi3U3GyANotgmoQfj0fR49wBxmL-oLogc/export?format=xlsx"
)


def download_file(url: str, dest: str, label: str = "file") -> str:
    print(f"Downloading {label}…")
    print(f"  URL: {url}")
    resp = requests.get(url, timeout=120)
    resp.raise_for_status()
    os.makedirs(os.path.dirname(dest) or ".", exist_ok=True)
    with open(dest, 'wb') as f:
        f.write(resp.content)
    print(f"  Saved: {dest} ({len(resp.content):,} bytes)")
    return dest


def download_profiles(output_dir: str | None = None) -> tuple[str, str]:
    """Download both profile xlsx files. Returns (official_path, community_path)."""
    if output_dir is None:
        output_dir = os.path.dirname(os.path.abspath(__file__))
    os.makedirs(output_dir, exist_ok=True)

    official = download_file(
        OFFICIAL_URL,
        os.path.join(output_dir, "Profile_Official.xlsx"),
        "Official FIT Profile",
    )
    community = download_file(
        COMMUNITY_URL,
        os.path.join(output_dir, "Profile_Community.xlsx"),
        "Community FIT Profile",
    )
    return official, community


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else None
    off, com = download_profiles(out)
    print(f"\nDone.  Official: {off}\n       Community: {com}")
