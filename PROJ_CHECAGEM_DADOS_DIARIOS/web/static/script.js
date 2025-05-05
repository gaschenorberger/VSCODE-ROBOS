document.addEventListener('DOMContentLoaded', () => {
    const refreshButtons = document.querySelectorAll('.refresh-btn');
    
    refreshButtons.forEach(button => {
        button.addEventListener('click', (event) => {
            const tableId = event.target.dataset.table;
            fetch(`/refresh/${tableId}`)
                .then(response => response.json())
                .then(data => {
                    alert(`Tabela ${tableId} atualizada com sucesso!`);
                    location.reload(); 
                })
                .catch(err => console.error(`Erro ao atualizar: ${err}`));
        });
    });
});

const siteButton = document.getElementById('btn-site');

window.addEventListener('scroll', () => {
    if (window.scrollY > 50) { 
        siteButton.classList.add('hidden');
    } else {
        siteButton.classList.remove('hidden');
    }
});

