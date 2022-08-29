const treinoModal = new bootstrap.Modal(document.getElementById('treino-modal'), {
    keyboard: false
})
const previewTreino = document.getElementById('preview-treino')
      previewTreino.addEventListener('click', hanldeModal)

const Modal = {
    setInfo({content, place}){
        place.innerHTML = content
    },
    callModal({modal}){
        modal.show()
    }
}

function hanldeModal({target}){
    Modal.setInfo({content: Treino.mountTableTreino() , place : document.getElementById('modal-treino-table')})
    Modal.callModal({modal : treinoModal})
}