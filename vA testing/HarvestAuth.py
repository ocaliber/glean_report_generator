# Replace these with your Harvest app's client ID and secret
CLIENT_ID = 'tDTndnUEoM8b4IYfW5OVMkCE'
CLIENT_SECRET = 'f9Uys-XRX5vtB9s93ZlSz_qtEG8Sr-pg84TMjKog_61HlTqlgT0617VYb0zPs5eBls1zYj1h3JYD7KQnZ6YhPA'
REDIRECT_URI = 'http://127.0.0.1:5000/callback'
from flask import Flask, request, redirect, session, url_for, render_template_string
import requests
import os

app = Flask(__name__)
app.secret_key = os.urandom(24)


AUTH_URL = 'https://id.getharvest.com/oauth2/authorize'
TOKEN_URL = 'https://id.getharvest.com/api/v2/oauth2/token'
TOKEN_URL = 'https://id.getharvest.com/oauth2/authorize?client_id={'+CLIENT_ID+'}&response_type=token'
USER_URL = 'https://api.harvestapp.com/v2/users/me'

@app.route('/')
def index():
    return 'Welcome to the Harvest API Authenticator! <a href="/login">Login with Harvest</a>'

@app.route('/login')
def login():
    auth_params = {
        'client_id': CLIENT_ID,
        'redirect_uri': REDIRECT_URI,
        'response_type': 'code'
    }
    auth_url = requests.Request('GET', AUTH_URL, params=auth_params).prepare().url
    print("Authorization URL: ", auth_url)  # Debugging: Print the full authorization URL
    return redirect(auth_url)

@app.route('/callback')
def callback():
    code = request.args.get('code')
    if not code:
        return 'Error: No authorization code provided.'

    token_data = {
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'code': code,
        'grant_type': 'authorization_code',
        'redirect_uri': REDIRECT_URI
    }

    token_response = requests.post(TOKEN_URL)

    #token_response = requests.post(TOKEN_URL, data=token_data, headers={'Accept': 'application/json'}, verify=False)

    # Log the full token response for debugging
    print("\n\n")
    print("Token Response Status Code: ", token_response.status_code)
    print("Token Response Headers: ", token_response.headers)
    print("Token Response Content: ", token_response.content)

    try:
        token_response_data = token_response.json()
    except ValueError as e:
        print("Error decoding token response JSON: ", e)
        print("Token Response Content: ", token_response.content)
        return 'Error decoding token response.'

    if token_response.status_code in [200, 201]:
        session['access_token'] = token_response_data['access_token']
        session['refresh_token'] = token_response_data['refresh_token']
    else:
        return f'Error: Failed to retrieve access token. Response: {token_response.content}'

    # Fetch the user's account ID
    headers = {
        'Authorization': f'Bearer {session["access_token"]}'
    }
    user_response = requests.get(USER_URL, headers=headers, verify=False)
    
    # Print response details for debugging
    print("User Response Status Code: ", user_response.status_code)
    print("User Response Headers: ", user_response.headers)
    print("User Response Content: ", user_response.content)
    
    # Handle potential errors and invalid JSON responses
    if user_response.status_code == 200:
        try:
            user_data = user_response.json()
            session['account_id'] = user_data['id']  # Fetch the user's ID
            # Print account ID for debugging
            print("Account ID Retrieved: ", session['account_id'])
        except ValueError as e:
            print("Error decoding user response JSON: ", e)
            return f'Error decoding user response JSON: {e}', 500
    else:
        return f'Failed to fetch user details: {user_response.status_code}', 500

    return redirect(url_for('protected'))

@app.route('/protected')
def protected():
    if 'access_token' not in session or 'account_id' not in session:
        return redirect(url_for('login'))

    headers = {
        'Authorization': f'Bearer {session["access_token"]}',
        'Harvest-Account-Id': str(session['account_id'])  # Convert to string
    }
    
    # Print access token and account ID being used for the API request
    print("Using Access Token: ", session['access_token'])
    print("Using Account ID: ", session['account_id'])

    response = requests.get('https://api.harvestapp.com/v2/projects', headers=headers, verify=False)
    print("Response Status Code: ", response.status_code)
    print("Response Headers: ", response.headers)
    print("Response Content: ", response.content)
    
    if response.status_code == 401:
        # Handle token refresh
        refresh_token_response = requests.post(
            TOKEN_URL,
            data={
                'client_id': CLIENT_ID,
                'client_secret': CLIENT_SECRET,
                'refresh_token': session['refresh_token'],
                'grant_type': 'refresh_token'
            },
            headers={
                'Accept': 'application/json'
            },
            verify=False
        )

        print("Refresh Token Response Status Code: ", refresh_token_response.status_code)
        print("Refresh Token Response Headers: ", refresh_token_response.headers)
        print("Refresh Token Response Content: ", refresh_token_response.content)

        try:
            refresh_token_data = refresh_token_response.json()
        except ValueError as e:
            print("Error decoding refresh token response JSON: ", e)
            print("Refresh Token Response Content: ", refresh_token_response.content)
            return 'Error decoding refresh token response.'

        if refresh_token_response.status_code in [200, 201]:
            session['access_token'] = refresh_token_data['access_token']
            session['refresh_token'] = refresh_token_data['refresh_token']

            # Print new access token after refresh
            print("New Access Token: ", session['access_token'])

            # Retry the request
            headers['Authorization'] = f'Bearer {session["access_token"]}'
            print("Retrying with New Access Token: ", session['access_token'])
            response = requests.get('https://api.harvestapp.com/v2/projects', headers=headers, verify=False)
            print("Retried Response Status Code: ", response.status_code)
            print("Retried Response Headers: ", response.headers)
            print("Retried Response Content: ", response.content)
        else:
            return f'Error: Failed to refresh access token. Response: {refresh_token_response.content}'

    if response.status_code == 200:
        projects = response.json()['projects']
        return render_template_string(PROJECTS_TEMPLATE, projects=projects)
    else:
        return f'Failed to access protected resource: {response.status_code}'

PROJECTS_TEMPLATE = '''
<!DOCTYPE html>
<html>
<head>
    <title>Harvest Projects</title>
</head>
<body>
    <h1>Select a Project</h1>
    <form method="post" action="/process_project">
        <select name="project_id">
            {% for project in projects %}
                <option value="{{ project.id }}">{{ project.name }}</option>
            {% endfor %}
        </select>
        <button type="submit">Submit</button>
    </form>
</body>
</html>
'''

@app.route('/process_project', methods=['POST'])
def process_project():
    project_id = request.form['project_id']
    return f'You selected project ID: {project_id}'

if __name__ == '__main__':
    print("Starting Flask app")
    app.run(debug=True)

