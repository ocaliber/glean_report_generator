import os
import sys
import requests
from flask import Blueprint, request, jsonify

# Blueprint for SharePoint template APIs
templates_bp = Blueprint('templates_bp', __name__)

# Env vars
TENANT_ID      = os.getenv("MS_TENANT_ID")
CLIENT_ID      = os.getenv("MS_CLIENT_ID")
CLIENT_SECRET  = os.getenv("MS_CLIENT_SECRET")
SITE_ID        = os.getenv("SHAREPOINT_SITE_ID")

# sanity check
def _check_env():
    missing = [k for k in ("MS_TENANT_ID","MS_CLIENT_ID","MS_CLIENT_SECRET","SHAREPOINT_SITE_ID") if not os.getenv(k)]
    if missing:
        raise RuntimeError(f"Missing environment variables: {', '.join(missing)}")

# 1) Get app-only token
def get_graph_token():
    _check_env()
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    }
    r = requests.post(url, data=data)
    r.raise_for_status()
    return r.json()["access_token"]

# 2) List folder contents by Graph
#    if path=='/' list drive root
#    else list drive items under path

@templates_bp.route('/api/templates')
def list_templates():
    token = get_graph_token()
    headers = {"Authorization": f"Bearer {token}"}
    path = request.args.get('path', '/')
    if path == '/' or path == '' or path == None:
        # list root children
        url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drive/root/children"
    else:
        # find folder by path then list children
        # Graph supports path syntax
        # ensure no leading slash
        p = path.lstrip('/')
        url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drive/root:/{p}:/children"
    resp = requests.get(url, headers=headers)
    if resp.status_code == 404:
        return jsonify([])
    resp.raise_for_status()
    items = resp.json().get('value', [])
    # return minimal fields
    results = [ { 'id': i['id'], 'name': i['name'], 'folder': 'folder' in i } for i in items ]
    return jsonify(results)
