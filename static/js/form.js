
const init = () => {
    setEvent({
        event : 'click',
        callback : addExercicio, 
        target: addExercicioButton
    })
    setEvent({
        event : 'submit',
        callback : handleSubmit, 
        target: form
    })
    setEvent({
        event : 'change',
        callback : handleList, 
        target: weekDay
    })
    setEvent({
        event : 'click',
        callback : PosTreino.handle, 
        target: addExercicioPosTreino
    })
}
init()

function removeExercise({target}){
    const {weekDay, exercicioIndex} = target.dataset
    const toRemove = treinos[weekDay].exercicios[exercicioIndex]
    treinos[weekDay].exercicios = treinos[weekDay].exercicios.filter( exercicio => exercicio !== toRemove )
    handleList()
}

function addExercicio({target}){
    const fields = Form.getFields()
    const exerciciosFields = Form.exerciciosFields({fields})
    if(!exerciciosFields) return;
    const {exercicio, weekday, error} = Exercicio.mountExercicio({fields : exerciciosFields})
    treinos = Treino.appendExercicio({treinos, index : weekday, exercicio})
    if(!error) return handleList();
}

function handlePosTreino({target}){
    
}

function handleList(){
    const treinoIndex = weekDay.value
    const {exercicios} = treinos[treinoIndex] !== undefined ? treinos[treinoIndex] : Treino.createTreino({dia:treinoIndex});
    tbody.innerHTML = ''
    exercicios.forEach( ({nomeExercicio, series, repeticoes, descanso, link}, i) => {
        tbody.innerHTML += `
            <tr>
                <td>${i + 1}</td>
                <td><a href='${link}'>${nomeExercicio}</a></td>
                <td>${series}</td>
                <td>${repeticoes}</td>
                <td>${descanso}</td>
                <td>
                    <button data-week-day='${treinoIndex}' data-exercicio-index='${i}' class="btn btn-danger col remover-exercicio">Remover</button>
                </td>
            </tr>
        `
    })

    treinoFields.innerHTML = ''
    treinoFields.innerHTML += divInputs
    setEvent({
        event : 'click',
        callback : addExercicio, 
        target: document.getElementById('button-add-exercicio')
    })
    const buttonsRemoveExercicio = [...document.getElementsByClassName('remover-exercicio')]
    buttonsRemoveExercicio.forEach( input => {
        setEvent({event:'click', callback: removeExercise, target: input})
    })
}         

let trInputs = `
    <tr>
        <th scope="row"></th>
        <td>
            <input class="form-control mb-2" name="nomeExercicio" type="text" placeholder="Nome do exercício">
            <input class="form-control mb-2" name="linkVideo" type="text" placeholder="Link do vídeo">
            <input class="form-control mb-2" name="observacoes" type="text" placeholder="Observações">

        </td>
        <td>
            <input class="form-control mb-2" name="repeticoes" type="number" placeholder="Repeticoes"require >
            <input class="form-control mb-2" name="series" type="number" placeholder="Séries" require >
            <input class="form-control mb-2" name="descanso" type="text" placeholder="Descanso"require >
        </td>
        <td>
            <button class="btn btn-secondary col" id="button-add-exercicio">Adicionar</button>
        </td>
    </tr>
`;

let divInputs = `
 <div class="row d-flex justify-content-center">
     <div class="col">
        <input class="form-control mb-2" name="nomeExercicio" type="text" placeholder="Nome do exercício">
        <input class="form-control mb-2" name="linkVideo" type="text" placeholder="Link do vídeo">                       
     </div>
     <div class="col">
        <input class="form-control mb-2" name="observacoes" type="text" placeholder="Observações">
        <input class="form-control mb-2" name="repeticoes" type="number" placeholder="Repeticoes"require >            
     </div>
     <div class="col">
        <input class="form-control mb-2" name="series" type="number" placeholder="Séries" require >
        <input class="form-control mb-2" name="descanso" type="text" placeholder="Descanso"require >
     </div>
 </div>
 <div class="row d-flex justify-content-center">
     <button class="btn btn-success col" id="button-add-exercicio">Adicionar</button>
 </div>
`;




