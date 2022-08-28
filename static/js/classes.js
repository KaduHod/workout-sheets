class Worksheet {
    constructor(nome,objetivo,dataInicio,observacao,tipoTreino){
        this.nome = nome
        this.objetivo = objetivo
        this.dataInicio = dataInicio
        this.observacao = observacao
        this.tipoTreino = tipoTreino
        this.treinos = []
        this.posTreino = []
    }

    addTreino(treino){
        this.treino.push(treino)
    }
}

const Form = {
    getFields(){
        const select = [...form.getElementsByTagName('select')]
        const inputs = [...form.getElementsByTagName('input')]
        let fields = inputs.concat(select)
        return fields.map( input => {
            const {name, value} = input
            return {name, value, input}
        })
    },
    exerciciosFields({fields}){
        const names = ['nomeExercicio','weekDay','linkVideo', 'observacoes', 'series', 'repeticoes', 'descanso']
        const verify = str => names.indexOf(str) > -1;
        return fields.filter( ({name}) => verify(name))
    },
    dadosAlunoFields({fields}){
        const names = ['nomeAluno', 'objetivo', 'dataInicio', 'diaDaSemana', 'observacao', 'tipoTreino']
        const verify = str => names.indexOf(str) > -1;
        return fields.filter( ({name}) => verify(name))
    },
    addError(inputs){
        inputs.forEach( input => input.classList.add('error-input') )
    },
    posTreinoFields({fields}){
        const names = ['nomeExercicioPosTreino', 'linkExercicioPosTreino']
        const verify = str => names.indexOf(str) > -1;
        return fields.filter( ({name}) => verify(name))
    },
    verifyInputs(inputs){
        return inputs.filter( ({value}) => value == '')
    }
}

const Treino = {
    appendExercicio({treinos, index, exercicio}){
        if(treinos[index] === undefined){
            treinos[index] = {
                dia : index,
                exercicios : [exercicio]
            }

            return treinos
        }
        treinos[index].exercicios.push(exercicio);
        return treinos;
    },
    removeLastTreino({treino}){
        treino.pop()
        return treino
    },
    createTreino({dia}){
        return {dia, exercicios : []}
    },
    mountTableTreino(){
        if(!treinos.length) return 'Nenhum treino encontrado. Seu frango!'
        const tables = treinos.map( ({dia, exercicios}) => {
            return Treino.setTableTreino({dia, exercicios})
        });
        return tables;
    },
    setTableTreino({dia, exercicios}){
        const nameDia = Semana.getDia(dia)
        const trs = exercicios.reduce( (acc, exercicio, i)=> {
            const {nomeExercicio,linkVideo,series,repeticoes,descanso,observacoes} = exercicio
            acc += `
            <tr>
                <td> ${i + 1} </td>
                <td>
                    <a href='${linkVideo}'> ${nomeExercicio} </a> 
                </td>
                <td> ${series} </td>
                <td> ${repeticoes} </td>
                <td> ${descanso} </td>
                <td> ${observacoes} </td>
            </tr>`
            return acc
        },'')

        const table = `
        <table class="table">
            <h3>${nameDia}</h3>
            <thead>
                <tr>
                    <th scope="col">#</th>
                    <th scope="col">Exercício</th>
                    <th scope="col">Séries</th>
                    <th scope="col">Repeticoes</th>
                    <th scope="col">Descanso</th>
                    <th scope="col">Observações</th>
                </tr>
            </thead>
            <tbody id="list-tBody-modal">
                ${trs}
            </tbody>
        </table>
        `
        return table;
    }
}

const Exercicio = {
    mountExercicio({fields}){
        fields.forEach( ({input}) => input.classList.remove('error-input') )
        const verify = Form.verifyInputs(fields)
        if(verify.length) {
            verify.forEach( ({input}) => {
                Form.addError([input]) 
            });
            return {error : true};
        }
        var weekday = null;
        const exercicio = fields.reduce((acc, {name, value}) => {
            if(name == 'weekDay'){
                weekday = value
                return acc;
            }
            acc[name] = value;
            return acc;
        },{})
        return {exercicio, weekday}
    }
}

const PosTreino = {
    handle(){
        const fields = Form.getFields()
        const posTreinoFields = Form.posTreinoFields({fields})
        const verify = Form.verifyInputs(posTreinoFields)

        if(verify.length){
            const inputs = verify.map(({input}) => input)
            return Form.addError(inputs);
        }
        const [descricao, link] = posTreinoFields
        PosTreino.addExercicio({descricao : descricao.value, link: link.value})
        PosTreino.mountList()
    },
    addExercicio({descricao, link}){
        posTreino.exercicios.push({descricao, link})
    },
    mountList(){
        const list = posTreino.exercicios.reduce( (acc, {descricao, link}, i)=>{
            acc += `
                <li class='row'> 
                    <div class="col-1">
                        <a href='${link}'>${descricao}</a>
                    </div>
                    <div class="col-1">
                        <button class='btn btn-danger remove-exercicio-pos-treino'  data-exercicio-index='${i}' >Remover</button>
                    </div>
                </li>`
            return acc
        },'')
        listPosTreino.innerHTML = list + PosTreino.inputHtml
        PosTreino.setEventsPosTreino()
    },
    setEventsPosTreino(){
        const removes = [...document.getElementsByClassName('remove-exercicio-pos-treino')]
            removes.forEach( input => setEvent({event: 'click', target:input, callback: PosTreino.removeExercico}))
        const addExercicio = document.getElementById('add-exercicio-pos-treino')
        setEvent({
            event : 'click',
            target : addExercicio,
            callback : PosTreino.handle
        })
    },
    inputHtml : `
        <li class="row">
            <div class="col">
                <input name="nomeExercicioPosTreino" class="form-control mx-auto" type="text"       placeholder="Descrição do exercicio"> 
            </div>
            <div class="col">
                <input name="linkExercicioPosTreino" class="form-control mx-auto" type="text"       placeholder="Link do vídeo"> 
            </div>
        </li>
        <li>
            <button class="btn btn-success" id="add-exercicio-pos-treino">Adicionar exercício</button> 
        </li>`
}