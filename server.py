from flask import Flask, request, url_for,render_template
from src.sheet import WorksheetMaycao
from src.exercicios import Exercicio
from src.treino import Treino
from flask_cors import CORS, cross_origin
app = Flask(__name__)
cors = CORS(app)
app.config['CORS_HEADERS'] = 'Content-Type'
# url_for('static', filename='')

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html', name='Teste var')

@app.route('/excel', methods=['post'])
@cross_origin()
def excel():
   print(request.form)
   return {"True":"treu"}