from flask import Flask, request, url_for,render_template, send_file
from src.sheet import WorksheetMaycao
from src.exercicios import Exercicio
from src.treino import Treino
from flask_cors import CORS, cross_origin
import json
app = Flask(__name__)
cors = CORS(app)
app.config['CORS_HEADERS'] = 'Content-Type'
# url_for('static', filename='')

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/excel', methods=['post'])
@cross_origin()
def excel():
   args = request.get_json()
   treinos = args['treinos']
   args['treinos'] = []
   for treino in treinos:
       if treino is None :
        continue
       if 'exercicios' not in treino:
        continue
       exercicios = treino['exercicios']
       exerciciosClass = []
       for exercicio in exercicios:
        exerciciosClass.append( Exercicio(exercicio) )
       treino['exercicios'] = []
       treino['exercicios'] = exerciciosClass
       newTreino = Treino(treino)
       args['treinos'].append(newTreino)
   args['workbookName'] = 'Treino.xlsx'
   worksheet = WorksheetMaycao(args)
   worksheet.createSheet()
   worksheet.setDimensions()
   worksheet.appendDadosAluno()
   worksheet.stylesDadosAluno()
   worksheet.appendExercicios()
   worksheet.appendPosTreino()
   worksheet.stylePosTreino()
   worksheet.workbook.save(filename = "Treino.xlsx")
   return json.dumps({'success':True}), 200, {'ContentType':'application/json'} 

@app.route('/download', methods=['get'])
def download():
    return send_file('Treino.xlsx', as_attachment=True)