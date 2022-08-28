const addExercicioButton = document.getElementById('button-add-exercicio')
const addTreinoButton = document.getElementById('button-add-treino')
const form = document.getElementsByTagName('form')[0]
form.addEventListener('submit', handleSubmit)
addExercicioButton.addEventListener('click', addExercicio)
const verififyArray = value => value.isArray() ? value : [];
let treinos = []
const weekDay = document.getElementById('weekDay')
weekDay.addEventListener('change', handleList)
const tbody = document.getElementById('list-tBody')

function handleSubmit(e){
    e.preventDefault()
    
}
function addExercicio({target}){
    const fields = Form.getFields()
    const exerciciosFields = Form.exerciciosFields({fields})
    if(!exerciciosFields) return;
    const {exercicio, weekday} = Exercicio.mountExercicio({fields : exerciciosFields})
    treinos = Treino.appendExercicio({treinos, index : weekday, exercicio})
    handleList()
}

function handleList(){
    const treinoIndex = weekDay.value
    const treino = treinos[treinoIndex]
    tbody.innerHTML = ''
    treino.exercicios.forEach( ({nomeExercicio, series, repeticoes, link}) => {
        tbody.innerHTML += `
            <tr>
                <td><a href='${link}'>${nomeExercicio}</a></td>
                <td>${series}</td>
                <td>${repeticoes}</td>
            </tr>
        `
    } )
}





