
from multiprocessing.sharedctypes import Value
import openpyxl
from textwrap import fill
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Color, colors, Side, GradientFill, Alignment, Border, PatternFill

class WorksheetMaycao:
    def __init__(self, args):
        self.nomeAluno = args["nomeAluno"]
        self.workbookName = args["workbookName"]
        self.objetivo = args["objetivo"]
        self.diasDaSemana = args["diasDaSemana"]
        self.observacao = args["observacao"]
        self.dataDeInicio = args["dataDeInicio"]
        self.treinos = args['treinos']
        self.tipoTreino = args['tipoTreino']
        self.workbook = 0
        self.worksheet = 0

    def createSheet(self):
        self.workbook = Workbook()
        self.worksheet = self.workbook.active
        self.worksheet.title = self.workbookName   

    def setDimensions(self):
        self.worksheet.column_dimensions['A'].width = 15 
        self.worksheet.column_dimensions['B'].width = 25 
        self.worksheet.column_dimensions['C'].width = 15 
        self.worksheet.column_dimensions['D'].width = 15 
        self.worksheet.column_dimensions['E'].width = 15 
        self.worksheet.column_dimensions['F'].width = 10 
        self.worksheet.column_dimensions['G'].width = 25 
        self.worksheet.column_dimensions['A'].height = 25
        self.worksheet.column_dimensions['B'].height = 25
        self.worksheet.column_dimensions['C'].height = 25
        self.worksheet.column_dimensions['D'].height = 25
        self.worksheet.column_dimensions['E'].height = 25
        self.worksheet.column_dimensions['F'].height = 25
        self.worksheet.column_dimensions['G'].height = 25

    def appendDadosAluno(self): 
        self.worksheet['A1'] = "Aluno:"
        self.worksheet['A2'] = "Data de Ínicio:"
        self.worksheet['A3'] = "Objetivo:"
        self.worksheet['A4'] = "Dias da semana:"
        self.worksheet['A5'] = "Observação:"
        self.worksheet['A6'] = "Treino:"
        self.worksheet['B1'] = self.nomeAluno
        self.worksheet['B2'] = self.dataDeInicio
        self.worksheet['B3'] = self.objetivo
        self.worksheet['B4'] = self.diasDaSemana
        self.worksheet['B5'] = self.observacao
        self.worksheet['B6'] = self.tipoTreino
    
    def stylesDadosAluno(self):
        key_font = Font(color='00FFFFFF')
        key_fill = PatternFill("solid", fgColor="00969696")
        key_alignment = Alignment(horizontal="right", vertical="center")     
        self.worksheet['A1'].border = Border(top = Side(border_style="thick", color="00000000"),
        left = Side(border_style="thick", color="00000000"))
        self.worksheet['A2'].border = Border(left = Side(border_style="thick", color="00000000"))
        self.worksheet['A3'].border = Border(left = Side(border_style="thick", color="00000000"))
        self.worksheet['A4'].border = Border(left = Side(border_style="thick", color="00000000"))
        self.worksheet['A5'].border = Border(left = Side(border_style="thick", color="00000000"))
        self.worksheet['A6'].border = Border(left = Side(border_style="thick", color="00000000"),
        bottom = Side(border_style="thick", color="00000000"))
        self.worksheet['B1'].border = Border(top = Side(border_style="thick", color="00000000"),
        right = Side(border_style="thick", color="00000000"))
        self.worksheet['B2'].border = Border(right = Side(border_style="thick", color="00000000"))
        self.worksheet['B3'].border = Border(right = Side(border_style="thick", color="00000000"))
        self.worksheet['B4'].border = Border(right = Side(border_style="thick", color="00000000"))
        self.worksheet['B5'].border = Border(right = Side(border_style="thick", color="00000000"))
        self.worksheet['B6'].border = Border(right = Side(border_style="thick", color="00000000"),
        bottom = Side(border_style="thick", color="00000000"))
        self.worksheet['A1'].fill = key_fill
        self.worksheet['A2'].fill = key_fill
        self.worksheet['A3'].fill = key_fill
        self.worksheet['A4'].fill = key_fill
        self.worksheet['A5'].fill = key_fill
        self.worksheet['A6'].fill = key_fill
        self.worksheet['A1'].font = key_font
        self.worksheet['A2'].font = key_font
        self.worksheet['A3'].font = key_font
        self.worksheet['A4'].font = key_font
        self.worksheet['A5'].font = key_font
        self.worksheet['A6'].font = key_font
        self.worksheet['A1'].alignment  = key_alignment 
        self.worksheet['A2'].alignment  = key_alignment 
        self.worksheet['A3'].alignment  = key_alignment 
        self.worksheet['A4'].alignment  = key_alignment 
        self.worksheet['A5'].alignment  = key_alignment 
        self.worksheet['A6'].alignment  = key_alignment 
        value_font = Font(color='00000000', bold=True)
        value_fill = PatternFill("solid", fgColor="00FFFFFF")
        value_alignment = Alignment(horizontal="center", vertical="center")
        self.worksheet['B1'].font = value_font
        self.worksheet['B2'].font = value_font
        self.worksheet['B3'].font = value_font
        self.worksheet['B4'].font = value_font
        self.worksheet['B5'].font = value_font
        self.worksheet['B6'].font = value_font
        self.worksheet['B1'].fill = value_fill
        self.worksheet['B2'].fill = value_fill
        self.worksheet['B3'].fill = value_fill
        self.worksheet['B4'].fill = value_fill
        self.worksheet['B5'].fill = value_fill
        self.worksheet['B6'].fill = value_fill
        self.worksheet['B1'].alignment = value_alignment
        self.worksheet['B2'].alignment = value_alignment
        self.worksheet['B3'].alignment = value_alignment
        self.worksheet['B4'].alignment = value_alignment
        self.worksheet['B5'].alignment = value_alignment
        self.worksheet['B6'].alignment = value_alignment

    def appendExercicios(self):
        startingCelLine = 7
        
        for i, treino in enumerate(self.treinos):
            startingCelLine+= 1
            treinoStartingCelLineCopy = startingCelLine + 1
            columnNumber = str(startingCelLine)
            self.worksheet['A' + columnNumber] = "Nº"
            self.worksheet['B' + columnNumber] = "EXERCÍCIO"
            self.worksheet['C' + columnNumber] = "SÉRIES"
            self.worksheet['D' + columnNumber] = "REPETIÇÕES"
            self.worksheet['E' + columnNumber] = "DESCANSO"
            self.worksheet['F' + columnNumber] = "VÍDEO"
            self.worksheet['G' + columnNumber] = "OBS"

            self.styleExercicioColumns(columnNumber)
            
            for iExercicio, exercicio in enumerate(treino.exercicios):
                startingCelLine+= 1
                columnNumber = str(startingCelLine)
                self.worksheet['A' + columnNumber] = iExercicio + 1
                self.worksheet['B' + columnNumber] = exercicio.nomeExercicio
                self.worksheet['C' + columnNumber] = exercicio.series
                self.worksheet['D' + columnNumber] = exercicio.repeticoes
                self.worksheet['E' + columnNumber] = exercicio.descanso
                self.worksheet['F' + columnNumber].hyperlink = exercicio.linkVideo
                self.worksheet['F' + columnNumber].value = ''
                img = openpyxl.drawing.image.Image('./ytb.png')
                img.width = 20
                img.height = 20
                img.anchor = 'F' + columnNumber
                self.worksheet.add_image(img)
                self.worksheet['G' + columnNumber] = exercicio.observacao
                self.styleExercicioRows(columnNumber)
                treinoEndingCelLineCopy = startingCelLine -1


            self.worksheet.merge_cells('H' + str(treinoStartingCelLineCopy) + ':H' + str(columnNumber))
            self.worksheet['H' + str(treinoStartingCelLineCopy)].alignment = Alignment(horizontal="center", vertical="center")
            self.worksheet['H' + str(treinoStartingCelLineCopy)].fill = PatternFill("solid", fgColor="00333333")
            print('aqui')
            self.worksheet['H' + str(treinoStartingCelLineCopy)].font = Font(color='00FFFF00',size=20)
            self.worksheet['H' + str(treinoStartingCelLineCopy)] = treino.dia       

    def styleExercicioRows(self, columnNumber) : 
        alignment = Alignment(horizontal="center", vertical="center")
        self.worksheet['A' + columnNumber].alignment = alignment
        self.worksheet['B' + columnNumber].alignment = alignment
        self.worksheet['C' + columnNumber].alignment = alignment
        self.worksheet['D' + columnNumber].alignment = alignment
        self.worksheet['E' + columnNumber].alignment = alignment
        self.worksheet['F' + columnNumber].alignment = alignment
        self.worksheet['G' + columnNumber].alignment = alignment
        self.worksheet['H' + columnNumber].alignment = alignment
        
        self.worksheet['H' + columnNumber].font = Font(color='00FFFF00',size=20)

    def styleExercicioColumns(self, columnNumber) :
        fill = PatternFill("solid", fgColor="00333333")
        font = Font(color='00FFFF00')
        alignment = Alignment(horizontal="center", vertical="center")
        self.worksheet['A' + columnNumber].fill = fill
        self.worksheet['B' + columnNumber].fill = fill
        self.worksheet['C' + columnNumber].fill = fill
        self.worksheet['D' + columnNumber].fill = fill
        self.worksheet['E' + columnNumber].fill = fill
        self.worksheet['F' + columnNumber].fill = fill
        self.worksheet['G' + columnNumber].fill = fill
        self.worksheet['A' + columnNumber].font = font
        self.worksheet['B' + columnNumber].font = font
        self.worksheet['C' + columnNumber].font = font
        self.worksheet['D' + columnNumber].font = font
        self.worksheet['E' + columnNumber].font = font
        self.worksheet['F' + columnNumber].font = font
        self.worksheet['G' + columnNumber].font = font
        self.worksheet['H' + columnNumber].font = font
        self.worksheet['A' + columnNumber].alignment = alignment
        self.worksheet['B' + columnNumber].alignment = alignment
        self.worksheet['C' + columnNumber].alignment = alignment
        self.worksheet['D' + columnNumber].alignment = alignment
        self.worksheet['E' + columnNumber].alignment = alignment
        self.worksheet['F' + columnNumber].alignment = alignment
        self.worksheet['G' + columnNumber].alignment = alignment
        self.worksheet['H' + columnNumber].alignment = alignment
                
class Exercicio: 
    def __init__(self, args):
        self.nomeExercicio = args["nomeExercicio"]
        self.series = args["series"]
        self.repeticoes = args["repeticoes"]
        self.descanso = args["descanso"]
        self.linkVideo = args["linkVideo"]
        self.observacao = args["observacao"]

class Treino:
    def __init__(self, args):
        self.exercicios = args['exercicios'] 
        self.dia = args['dia']

        

def main():
    exerciciosSegunda = {
        "dia" : "SEG",
        "exercicios" :[
            Exercicio({
                "nomeExercicio" : "VOADOR",
                "series" : "4",
                "repeticoes" : "8 A 12",
                "descanso" : "1 MIN",
                "linkVideo" : "https://www.youtube.com/watch?v=p0Mql-QHnvY",
                "observacao" : "DROP SET NA ULTIMA SÉRIE"
            }),
            Exercicio({
                "nomeExercicio" : "ELEVAÇÃO LATERAL",
                "series" : "4",
                "repeticoes" : "8 A 12",
                "descanso" : "1 MIN",
                "linkVideo" : "https://www.youtube.com/watch?v=p0Mql-QHnvY",
                "observacao" : "DROP SET NA ULTIMA SÉRIE"
            }),
            Exercicio({
                "nomeExercicio" : "LEG MAQUINA",
                "series" : "3",
                "repeticoes" : "8 A 12",
                "descanso" : "1 MIN",
                "linkVideo" : "https://www.youtube.com/watch?v=p0Mql-QHnvY",
                "observacao" : "DROP SET NA ULTIMA SÉRIE"
            }),
            Exercicio({
                "nomeExercicio" : "EXTENSORA",
                "series" : "3",
                "repeticoes" : "8 A 12",
                "descanso" : "1 MIN",
                "linkVideo" : "https://www.youtube.com/watch?v=p0Mql-QHnvY",
                "observacao" : "DROP SET NA ULTIMA SÉRIE"
            }),
            Exercicio({
                "nomeExercicio" : "GLUTEO POLIA BAIXA",
                "series" : "4",
                "repeticoes" : "8 A 12",
                "descanso" : "1 MIN",
                "linkVideo" : "https://www.youtube.com/watch?v=p0Mql-QHnvY",
                "observacao" : "DROP SET NA ULTIMA SÉRIE"
            }),
            Exercicio({
                "nomeExercicio" : "PANTURRILHA SENTADA",
                "series" : "4",
                "repeticoes" : "8 A 12",
                "descanso" : "1 MIN",
                "linkVideo" : "https://www.youtube.com/watch?v=p0Mql-QHnvY",
                "observacao" : "DROP SET NA ULTIMA SÉRIE"
            }),
        ]
    }

    exerciciosTerca = {
        "dia" : "TER",
        "exercicios" :[
            Exercicio({
                "nomeExercicio" : "VOADOR",
                "series" : "4",
                "repeticoes" : "8 A 12",
                "descanso" : "1 MIN",
                "linkVideo" : "https://www.youtube.com/watch?v=p0Mql-QHnvY",
                "observacao" : "DROP SET NA ULTIMA SÉRIE"
            }),
            Exercicio({
                "nomeExercicio" : "ELEVAÇÃO LATERAL",
                "series" : "4",
                "repeticoes" : "8 A 12",
                "descanso" : "1 MIN",
                "linkVideo" : "https://www.youtube.com/watch?v=p0Mql-QHnvY",
                "observacao" : "DROP SET NA ULTIMA SÉRIE"
            }),
            Exercicio({
                "nomeExercicio" : "LEG MAQUINA",
                "series" : "3",
                "repeticoes" : "8 A 12",
                "descanso" : "1 MIN",
                "linkVideo" : "https://www.youtube.com/watch?v=p0Mql-QHnvY",
                "observacao" : "DROP SET NA ULTIMA SÉRIE"
            }),
            Exercicio({
                "nomeExercicio" : "EXTENSORA",
                "series" : "3",
                "repeticoes" : "8 A 12",
                "descanso" : "1 MIN",
                "linkVideo" : "https://www.youtube.com/watch?v=p0Mql-QHnvY",
                "observacao" : "DROP SET NA ULTIMA SÉRIE"
            }),
            Exercicio({
                "nomeExercicio" : "GLUTEO POLIA BAIXA",
                "series" : "4",
                "repeticoes" : "8 A 12",
                "descanso" : "1 MIN",
                "linkVideo" : "https://www.youtube.com/watch?v=p0Mql-QHnvY",
                "observacao" : "DROP SET NA ULTIMA SÉRIE"
            }),
            Exercicio({
                "nomeExercicio" : "PANTURRILHA SENTADA",
                "series" : "4",
                "repeticoes" : "8 A 12",
                "descanso" : "1 MIN",
                "linkVideo" : "https://www.youtube.com/watch?v=p0Mql-QHnvY",
                "observacao" : "DROP SET NA ULTIMA SÉRIE"
            }),
        ]
    }

    exerciciosQuarta = {
        "dia" : "QUA",
        "exercicios" :[
            Exercicio({
                "nomeExercicio" : "VOADOR",
                "series" : "4",
                "repeticoes" : "8 A 12",
                "descanso" : "1 MIN",
                "linkVideo" : "https://www.youtube.com/watch?v=p0Mql-QHnvY",
                "observacao" : "DROP SET NA ULTIMA SÉRIE"
            }),
            Exercicio({
                "nomeExercicio" : "ELEVAÇÃO LATERAL",
                "series" : "4",
                "repeticoes" : "8 A 12",
                "descanso" : "1 MIN",
                "linkVideo" : "https://www.youtube.com/watch?v=p0Mql-QHnvY",
                "observacao" : "DROP SET NA ULTIMA SÉRIE"
            }),
            Exercicio({
                "nomeExercicio" : "LEG MAQUINA",
                "series" : "3",
                "repeticoes" : "8 A 12",
                "descanso" : "1 MIN",
                "linkVideo" : "https://www.youtube.com/watch?v=p0Mql-QHnvY",
                "observacao" : "DROP SET NA ULTIMA SÉRIE"
            }),
            Exercicio({
                "nomeExercicio" : "EXTENSORA",
                "series" : "3",
                "repeticoes" : "8 A 12",
                "descanso" : "1 MIN",
                "linkVideo" : "https://www.youtube.com/watch?v=p0Mql-QHnvY",
                "observacao" : "DROP SET NA ULTIMA SÉRIE"
            }),
            Exercicio({
                "nomeExercicio" : "GLUTEO POLIA BAIXA",
                "series" : "4",
                "repeticoes" : "8 A 12",
                "descanso" : "1 MIN",
                "linkVideo" : "https://www.youtube.com/watch?v=p0Mql-QHnvY",
                "observacao" : "DROP SET NA ULTIMA SÉRIE"
            }),
            Exercicio({
                "nomeExercicio" : "PANTURRILHA SENTADA",
                "series" : "4",
                "repeticoes" : "8 A 12",
                "descanso" : "1 MIN",
                "linkVideo" : "https://www.youtube.com/watch?v=p0Mql-QHnvY",
                "observacao" : "DROP SET NA ULTIMA SÉRIE"
            }),
        ]
    }

    exerciciosQuinta = {
        "dia" : "QUI",
        "exercicios" :[
            Exercicio({
                "nomeExercicio" : "VOADOR",
                "series" : "4",
                "repeticoes" : "8 A 12",
                "descanso" : "1 MIN",
                "linkVideo" : "https://www.youtube.com/watch?v=p0Mql-QHnvY",
                "observacao" : "DROP SET NA ULTIMA SÉRIE"
            }),
            Exercicio({
                "nomeExercicio" : "ELEVAÇÃO LATERAL",
                "series" : "4",
                "repeticoes" : "8 A 12",
                "descanso" : "1 MIN",
                "linkVideo" : "https://www.youtube.com/watch?v=p0Mql-QHnvY",
                "observacao" : "DROP SET NA ULTIMA SÉRIE"
            }),
            Exercicio({
                "nomeExercicio" : "LEG MAQUINA",
                "series" : "3",
                "repeticoes" : "8 A 12",
                "descanso" : "1 MIN",
                "linkVideo" : "https://www.youtube.com/watch?v=p0Mql-QHnvY",
                "observacao" : "DROP SET NA ULTIMA SÉRIE"
            }),
            Exercicio({
                "nomeExercicio" : "EXTENSORA",
                "series" : "3",
                "repeticoes" : "8 A 12",
                "descanso" : "1 MIN",
                "linkVideo" : "https://www.youtube.com/watch?v=p0Mql-QHnvY",
                "observacao" : "DROP SET NA ULTIMA SÉRIE"
            }),
            Exercicio({
                "nomeExercicio" : "GLUTEO POLIA BAIXA",
                "series" : "4",
                "repeticoes" : "8 A 12",
                "descanso" : "1 MIN",
                "linkVideo" : "https://www.youtube.com/watch?v=p0Mql-QHnvY",
                "observacao" : "DROP SET NA ULTIMA SÉRIE"
            }),
            Exercicio({
                "nomeExercicio" : "PANTURRILHA SENTADA",
                "series" : "4",
                "repeticoes" : "8 A 12",
                "descanso" : "1 MIN",
                "linkVideo" : "https://www.youtube.com/watch?v=p0Mql-QHnvY",
                "observacao" : "DROP SET NA ULTIMA SÉRIE"
            }),
        ]
    }

    exerciciosSexta = {
        "dia" : "SEX",
        "exercicios" :[
            Exercicio({
                "nomeExercicio" : "VOADOR",
                "series" : "4",
                "repeticoes" : "8 A 12",
                "descanso" : "1 MIN",
                "linkVideo" : "https://www.youtube.com/watch?v=p0Mql-QHnvY",
                "observacao" : "DROP SET NA ULTIMA SÉRIE"
            }),
            Exercicio({
                "nomeExercicio" : "ELEVAÇÃO LATERAL",
                "series" : "4",
                "repeticoes" : "8 A 12",
                "descanso" : "1 MIN",
                "linkVideo" : "https://www.youtube.com/watch?v=p0Mql-QHnvY",
                "observacao" : "DROP SET NA ULTIMA SÉRIE"
            }),
            Exercicio({
                "nomeExercicio" : "LEG MAQUINA",
                "series" : "3",
                "repeticoes" : "8 A 12",
                "descanso" : "1 MIN",
                "linkVideo" : "https://www.youtube.com/watch?v=p0Mql-QHnvY",
                "observacao" : "DROP SET NA ULTIMA SÉRIE"
            }),
            Exercicio({
                "nomeExercicio" : "EXTENSORA",
                "series" : "3",
                "repeticoes" : "8 A 12",
                "descanso" : "1 MIN",
                "linkVideo" : "https://www.youtube.com/watch?v=p0Mql-QHnvY",
                "observacao" : "DROP SET NA ULTIMA SÉRIE"
            }),
            Exercicio({
                "nomeExercicio" : "GLUTEO POLIA BAIXA",
                "series" : "4",
                "repeticoes" : "8 A 12",
                "descanso" : "1 MIN",
                "linkVideo" : "https://www.youtube.com/watch?v=p0Mql-QHnvY",
                "observacao" : "DROP SET NA ULTIMA SÉRIE"
            }),
            Exercicio({
                "nomeExercicio" : "PANTURRILHA SENTADA",
                "series" : "4",
                "repeticoes" : "8 A 12",
                "descanso" : "1 MIN",
                "linkVideo" : "https://www.youtube.com/watch?v=p0Mql-QHnvY",
                "observacao" : "DROP SET NA ULTIMA SÉRIE"
            }),
        ]
    }
    
    
    

    treinoA = Treino(exerciciosSegunda)
    treinoB = Treino(exerciciosTerca)
    treinoC = Treino(exerciciosQuarta)
    treinoD = Treino(exerciciosQuinta)
    treinoE = Treino(exerciciosSexta)
    treinos = [treinoA,treinoB,treinoC,treinoD,treinoE]

    args = {
        "workbookName" : "Planejamento.xslx",
        "nomeAluno" : "Carlos",
        "objetivo" : "Ganho de massa muscular",
        "diasDaSemana" : " Indefinido ",
        "observacao" : "2º treino adaptativo",
        "dataDeInicio" : "10/12/2022",
        "treinos" : treinos,
        "tipoTreino" : 'Força'
    }

    sheet = WorksheetMaycao(args) 
    sheet.createSheet()
    sheet.setDimensions()
    sheet.appendDadosAluno()
    sheet.stylesDadosAluno()
    sheet.appendExercicios()
    sheet.workbook.save(filename = "Carlos trieno.xlsx")
    




if __name__ == "__main__":
    main()