document.getElementById('uploadForm').addEventListener('submit', async (event) => {
    event.preventDefault();

    const fileInput = document.getElementById('xmlFile');
    const files = fileInput.files;

    if (files.length === 0) {
        alert('Por favor, selecione pelo menos um arquivo XML.');
        return;
    }

    uploadFile(files);
});

async function uploadFile(files) {
    const formData = new FormData();
    for (const file of files) {
        formData.append('files', file);
    }
    
    document.getElementById('status').textContent = 'Enviando arquivo, por favor aguarde...';
    const progressContainer = document.getElementById('progress-container');
    const progressBar = document.getElementById('progress-bar');
    progressContainer.style.display = 'block';

    try {
        const response = await fetch('https://localhost:7075/api/NFe/convert', {
            method: 'POST',
            body: formData,
        });

        if (response.ok) {
            progressBar.style.width = '100%';
            progressBar.classList.add('bg-sucess');
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'Converted.xlsx';
            a.click();
            window.URL.revokeObjectURL(url);
            document.getElementById('status').textContent = 'Download concluído!';
        } else {
            progressBar.classList.add('bg-danger');
            document.getElementById('status').textContent = `Erro: ${await response.text()}`;
        }
    } catch (error) {
        progressBar.classList.add('bg-danger');
        document.getElementById('status').textContent = `Erro de conexão: ${error.message}`;
    } finally {
        progressContainer.style.display = 'none';
    }
}
