from flask import Flask, redirect, request, session, url_for
import requests
from requests_oauthlib import OAuth2Session
import os
import traceback
from flask_session import Session

# Allow insecure HTTP for local development
os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'

CLIENT_ID = 'tDTndnUEoM8b4IYfW5OVMkCE'
CLIENT_SECRET = 'f9Uys-XRX5vtB9s93ZlSz_qtEG8Sr-pg84TMjKog_61HlTqlgT0617VYb0zPs5eBls1zYj1h3JYD7KQnZ6YhPA'
REDIRECT_URI = 'http://127.0.0.1:5000/callback'
AUTHORIZATION_BASE_URL = 'https://id.getharvest.com/oauth2/authorize'
TOKEN_URL = 'https://id.getharvest.com/api/v2/oauth2/token'
HARVEST_API_URL = 'https://api.harvestapp.com/v2/'

app = Flask(__name__)
app.secret_key = b'\x03\xfc\xbaOFa\x06;\x1b/\x95\xb8\x0c\x899\x88Ks\xb4\\K%?\xab'

# Session configuration
app.config['SESSION_COOKIE_SECURE'] = False
app.config['SESSION_COOKIE_HTTPONLY'] = True
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'
app.config['SESSION_TYPE'] = 'filesystem'
app.config['SESSION_FILE_DIR'] = os.path.join(os.getcwd(), 'flask_session')

Session(app)  # Use filesystem-based sessions

@app.route('/')
def index():
    harvest = OAuth2Session(CLIENT_ID, redirect_uri=REDIRECT_URI)
    authorization_url, state = harvest.authorization_url(AUTHORIZATION_BASE_URL)
    session['oauth_state'] = state
    return redirect(authorization_url)

@app.route('/callback')
def callback():
    try:
        # Ensure state matches
        if 'oauth_state' not in session or request.args.get('state') != session['oauth_state']:
            return "State mismatch error", 400

        # Exchange authorization code for access token
        token_data = {
            'client_id': CLIENT_ID,
            'client_secret': CLIENT_SECRET,
            'code': request.args.get('code'),
            'grant_type': 'authorization_code',
            'redirect_uri': REDIRECT_URI,
        }

        token_response = requests.post(TOKEN_URL, data=token_data, headers={'User-Agent': 'MyApp (yourname@example.com)'})

        if token_response.status_code not in [200, 201]:
            return f"Error exchanging token: {token_response.text}", 500

        # Store the token in the session
        session['oauth_token'] = token_response.json()

        return redirect(url_for('.profile'))

    except Exception as e:
        traceback.print_exc()
        return f"Error during callback: {e}", 500

@app.route('/profile')
def profile():
    try:
        # Retrieve token from session
        token = session.get('oauth_token')
        if not token:
            return redirect(url_for('index'))

        # Use the access token to fetch user data from Harvest API
        harvest = OAuth2Session(CLIENT_ID, token=token)
        response = harvest.get(f'{HARVEST_API_URL}users/me')

        if response.status_code == 200:
            return response.json()
        else:
            return f"Error fetching profile: {response.status_code}", 500

    except Exception as e:
        traceback.print_exc()
        return f"Error fetching profile: {e}", 500

# Optionally, refresh the token if needed
def refresh_token_if_needed():
    token = session.get('oauth_token')
    if token and 'expires_at' in token and token['expires_at'] < time.time():
        try:
            harvest = OAuth2Session(CLIENT_ID, token=token)
            new_token = harvest.refresh_token(
                TOKEN_URL,
                refresh_token=token['refresh_token'],
                client_id=CLIENT_ID,
                client_secret=CLIENT_SECRET
            )
            session['oauth_token'] = new_token
            return new_token
        except Exception as e:
            return None
    return token

if __name__ == '__main__':
    app.run(debug=True)
