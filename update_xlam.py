import os
import shutil
import argparse
from datetime import datetime
import sys

def get_default_addins_folder():
    """
    Returns the default Microsoft Excel Add-ins folder for the current user.
    Usually something like:
    C:\\Users\\<username>\\AppData\\Roaming\\Microsoft\\AddIns
    """
    user_profile = os.environ.get("USERPROFILE") or os.path.expanduser("~")
    default_path = os.path.join(user_profile, "AppData", "Roaming", "Microsoft", "AddIns")
    return default_path if os.path.exists(default_path) else None


def update_xlam(target_dir, filename="YourAddIn.xlam"):
    """
    Copies an .xlam file from this repository folder to the Excel Add-ins folder,
    overwriting the existing version and backing up the old one.
    """

    repo_dir = os.path.dirname(os.path.abspath(__file__))
    source_path = os.path.join(repo_dir, filename)
    target_path = os.path.join(target_dir, filename)

    # Validate source file
    if not os.path.exists(source_path):
        print(f"❌ Source file not found: {source_path}")
        sys.exit(1)

    # Create target directory if missing
    if not os.path.exists(target_dir):
        os.makedirs(target_dir)
        print(f"📁 Created target directory: {target_dir}")

    # Backup existing file (if any)
    if os.path.exists(target_path):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = f"{target_path}.{timestamp}.bak"
        shutil.copy2(target_path, backup_path)
        print(f"🗂️  Existing file backed up to: {backup_path}")

    # Copy and overwrite
    shutil.copy2(source_path, target_path)
    print(f"✅ Updated: {target_path}")
    print("🎉 XLAM file successfully updated!")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Update a local Excel Add-in (.xlam) file from this repository.")
    parser.add_argument("--target", help="Path to target directory for the .xlam (default: Excel Add-ins folder)")
    parser.add_argument("--filename", default="YourAddIn.xlam", help="Name of the .xlam file in this repo")

    args = parser.parse_args()

    target_dir = args.target or get_default_addins_folder()
    if not target_dir:
        print("❌ Could not find Excel Add-ins folder automatically. Please specify with --target.")
        sys.exit(1)

    update_xlam(target_dir, args.filename)
