from flask import Flask, render_template, request

app = Flask(__name__)

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/submit', methods=['POST'])
def submit():
    # Get the user-entered information from the form
    user_input = request.form['user_input']
    
    # Process the information (replace with your logic)
    message = f"You entered: {user_input}"
    
    # Return a response to the webpage
    return render_template('result.html', message=message)

if __name__ == '__main__':
    app.run(debug=True)
