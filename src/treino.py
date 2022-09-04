from src.dia import getDia
class Treino:
    def __init__(self, args):
        self.exercicios = args['exercicios'] 
        self.dia = getDia(int(args['dia']))

        