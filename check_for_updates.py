"""Check GitHub for a newer release of the Game Accessibility Testing Template.

Reads the version number embedded in a local .xlsx, compares it to the latest
GitHub Release of this project, and - if a newer one exists - downloads it to
a NEW file next to the local one. It never overwrites, edits, or deletes the
local file, because that file likely already has a tester's own logged
issues, Game/Mod Information, Game Sections, and Severity Scale entries in it.

Uses only the standard library (no pip install required) so it runs the same
way on Windows, macOS, and Linux with just `python check_for_updates.py`.
"""
import argparse
import json
import os
import sys
import urllib.error
import urllib.request
import xml.etree.ElementTree as ET
import zipfile

REPO_OWNER = "ChrisTavar2022"
REPO_NAME = "Game-Accessibility-Testing-Templit"
CORE_PROPS_NS = "{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}"
DEFAULT_FILENAME = "Game Accessibility Testing templit.xlsx"


def get_local_version(path):
    """Return the cp:version text from an .xlsx's docProps/core.xml, or None."""
    if not os.path.isfile(path):
        return None
    try:
        with zipfile.ZipFile(path) as zf:
            with zf.open("docProps/core.xml") as f:
                root = ET.parse(f).getroot()
    except (KeyError, zipfile.BadZipFile, ET.ParseError):
        return None
    version_el = root.find(f"{CORE_PROPS_NS}version")
    if version_el is not None and version_el.text:
        return version_el.text.strip()
    return None


def parse_version(text):
    """Best-effort "1.2.3" -> (1, 2, 3) for comparison; None if unparsable."""
    parts = text.split(".")
    try:
        return tuple(int(p) for p in parts)
    except ValueError:
        return None


def is_older(local, remote):
    """True if local < remote. Falls back to string inequality if either
    version string isn't cleanly numeric, so an unrecognized local version
    is still treated as "needs updating" rather than silently skipped."""
    local_parsed = parse_version(local)
    remote_parsed = parse_version(remote)
    if local_parsed is not None and remote_parsed is not None:
        return local_parsed < remote_parsed
    return local != remote


def fetch_latest_release():
    url = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/releases/latest"
    request = urllib.request.Request(
        url, headers={"User-Agent": "GameAccessibilityTemplate-UpdateChecker"}
    )
    with urllib.request.urlopen(request, timeout=15) as response:
        return json.loads(response.read().decode("utf-8"))


def download_file(url, dest_path):
    request = urllib.request.Request(
        url, headers={"User-Agent": "GameAccessibilityTemplate-UpdateChecker"}
    )
    with urllib.request.urlopen(request, timeout=60) as response, open(dest_path, "wb") as out:
        out.write(response.read())


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--template-path",
        default=os.path.join(os.path.dirname(os.path.abspath(__file__)), DEFAULT_FILENAME),
        help="Path to your local copy of the workbook "
             "(default: the template next to this script).",
    )
    args = parser.parse_args()
    template_path = args.template_path

    print("Game Accessibility Testing Template - Update Check")
    print("====================================================")
    print()

    print(f"Looking for your local file at:\n  {template_path}")
    local_version = get_local_version(template_path)
    if local_version:
        print(f"Your local version: {local_version}")
    else:
        print("Could not read a version number from that file (it may be missing, "
              "or from before versioning was added).")
        print("Continuing anyway - if GitHub has any release at all, it will be offered as an update.")
    print()

    print("Checking GitHub for the latest release...")
    try:
        release = fetch_latest_release()
    except urllib.error.HTTPError as e:
        if e.code == 404:
            print("No releases have been published yet on GitHub. Nothing to update to.")
        else:
            print(f"Could not reach GitHub to check for updates (HTTP {e.code}).")
            print(f"Check the repository manually: https://github.com/{REPO_OWNER}/{REPO_NAME}")
        return 0
    except (urllib.error.URLError, TimeoutError) as e:
        print(f"Could not reach GitHub to check for updates: {e}")
        print("Check your internet connection and try again, or download the latest "
              f"version manually from:\n  https://github.com/{REPO_OWNER}/{REPO_NAME}")
        return 1

    tag_name = release.get("tag_name", "")
    remote_version = tag_name[1:] if tag_name.startswith("v") else tag_name
    print(f"Latest release on GitHub: {tag_name}")
    print()

    needs_update = True
    if local_version:
        needs_update = is_older(local_version, remote_version)

    if not needs_update:
        print("You're already on the latest version. Nothing to do.")
        return 0

    assets = release.get("assets", [])
    asset = next((a for a in assets if a.get("name", "").lower().endswith(".xlsx")), None)
    if not asset:
        print(f"A newer release ({tag_name}) exists, but it doesn't have an .xlsx file attached.")
        print(f"Check the release page manually:\n  {release.get('html_url', '')}")
        return 0

    dest_dir = os.path.dirname(template_path) or "."
    dest_path = os.path.join(dest_dir, f"Game Accessibility Testing templit ({tag_name}).xlsx")

    print(f"Downloading {asset['name']} ...")
    try:
        download_file(asset["browser_download_url"], dest_path)
    except (urllib.error.URLError, OSError) as e:
        print(f"Download failed: {e}")
        print(f"You can download it by hand from:\n  {release.get('html_url', '')}")
        return 1

    print()
    print("Done. The new version was saved as a NEW file:")
    print(f"  {dest_path}")
    print()
    print("Your existing file was NOT changed:")
    print(f"  {template_path}")
    print()
    body = release.get("body")
    if body:
        print(f"What's new in {tag_name}:")
        print(body)
        print()
    print("Next steps: open both files and copy your logged issues, Game/Mod "
          "Information, Game Sections, and Severity Scale entries from your old "
          "file into the new one. Once everything has moved over and you've "
          "checked it, you can delete or archive the old file.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
