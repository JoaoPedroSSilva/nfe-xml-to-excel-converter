document.getElementById('uploadForm').addEventListener('submit', async (event) => {
    event.preventDefault();

    const fileInput = document.getElementById('xmlFile');
    const file = fileInput.files[0];

    if (!file || !file.name.toLowerCase().endsWith(".xml")) {
        window.alert('Favor, escolher um arquivo NF-e XML válido!');
        return;
    }

    size = file.size / 1024 / 1024;
    filesize =  size.toFixed(2);
    console.log("O arquivo tem:");
    console.log(`${filesize} MB.`);

    var sizeLimit = 4 * 1024 * 1024;
    if(file.size > sizeLimit) {
        window.alert('Tamanho do arquivo excede o limite de 4MB!');
        return;
    }

    uploadFile(file);
});

async function uploadFile(file) {
    const formData = new FormData();
    formData.append('file', file);
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
