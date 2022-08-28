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
    addError(input){
        input.classList.add('error-input')
    }
    // posTreinoFiels({fields}){
        // const names = ['nomeExercicio', 'linkVideo', 'observacoes', 'series', 'repeticoes', 'descanso']
        // const verify = str => names.indexOf(str) > -1;
        // return fields.filter( ({name}) => verify(name))
    // }
}

const Treino = {
    appendExercicio({treinos, index, exercicio}){
        console.log({treinos, index, exercicio})
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
    }
}

const Exercicio = {
    mountExercicio({fields}){
        fields.forEach( ({input}) => input.classList.remove('error-input') )
        const verify = this.verifyInputs(fields)
        if(verify.length) {
            verify.forEach( ({input}) => Form.addError(input) );
            return false;
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
    },
    verifyInputs(fields){
        return fields.filter( ({value}) => value == '')
    }
}