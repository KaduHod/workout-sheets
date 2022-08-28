var treinos = []
var posTreino = {
    exercicios : []
}
function handleSubmit(e){
    e.preventDefault()
}
const verififyArray = value => value.isArray() ? value : [];
const addExercicioButton = document.getElementById('button-add-exercicio')
const addTreinoButton = document.getElementById('button-add-treino')
const form = document.getElementsByTagName('form')[0]
const weekDay = document.getElementById('weekDay')
const addExercicioPosTreino = document.getElementById('add-exercicio-pos-treino')
const setEvent = ({event, callback, target}) => target.addEventListener(event, callback)
const tbody = document.getElementById('list-tBody')
const listPosTreino = document.getElementById('pos-treino-lista')



