#!/usr/bin/env python3
import os
import sys
import requests

# --- Configuration from environment ---
TENANT_ID       = os.getenv("MS_TENANT_ID")
CLIENT_ID       = os.getenv("MS_CLIENT_ID")
CLIENT_SECRET   = os.getenv("MS_CLIENT_SECRET")
SITE_ID         = os.getenv("SHAREPOINT_SITE_ID")
TEMPLATES_PATH  = os.getenv("SHAREPOINT_TEMPLATES_FOLDER")  # e.g. "Shared Documents/Templates"

# Basic sanity check
missing = [k for k in ("MS_TENANT_ID","MS_CLIENT_ID","MS_CLIENT_SECRET","SHAREPOINT_SITE_ID","SHAREPOINT_TEMPLATES_FOLDER")
           if not os.getenv(k)]
if missing:
    print(f"Error: missing env vars: {', '.join(missing)}")
    sys.exit(1)

# 1) Get an app-only token
def get_graph_token():
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "client_id":     CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope":         "https://graph.microsoft.com/.default",
        "grant_type":    "client_credentials",
    }
    resp = requests.post(url, data=data)
    resp.raise_for_status()
    return resp.json()["access_token"]

# 2) List folder children
def list_templates_folder(token):
    """
    Finds the ‘Templates’ folder at the root of the drive and returns
    its children. No more guessing about “Documents” vs “Shared Documents”.
    """
    headers = {"Authorization": f"Bearer {token}"}

    # 1) List root children to find the Templates folder
    root_url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drive/root/children"
    resp = requests.get(root_url, headers=headers)
    resp.raise_for_status()
    children = resp.json().get("value", [])

    tmpl_folder = next(
        (c for c in children if c.get("folder") and c["name"] == "Templates"),
        None
    )
    if not tmpl_folder:
        print("❌ No folder named ‘Templates’ found at drive root.")
        sys.exit(1)

    folder_id = tmpl_folder["id"]

    # 2) List that folder’s contents by its ID
    items_url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drive/items/{folder_id}/children"
    resp = requests.get(items_url, headers=headers)
    resp.raise_for_status()
    return resp.json().get("value", [])


def main():
    print("Fetching Graph token…")
    token = get_graph_token()
    print("Listing templates in:", TEMPLATES_PATH)
    items = list_templates_folder(token)
    if not items:
        print("No items found (check that the folder path is correct and your app has permission).")
    else:
        for item in items:
            kind = "📄" if item.get("file") else "📁"
            print(f"{kind} {item['name']}   (id: {item['id']})")

if __name__ == "__main__":
    main()
