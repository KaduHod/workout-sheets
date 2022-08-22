from flask import Flask, request, url_for,render_template


app = Flask(__name__)
# url_for('static', filename='')

@app.route("/")
def index():
    return render_template('index.html', name='Teste var')

@app.route('/excel', methods=['GET'])
def excel():
    return "<p>aqui!</p>"