document.getElementById('uploadForm').addEventListener('submit', async (event) => {
    event.preventDefault();

    console.log("Testando")
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
    const formData = new FormData();
    formData.append('file', file);

    try {
        const response = await fetch('https://localhost:7075/api/NFe/convert', {
            method: 'POST',
            body: formData,
        });

        if (response.ok) {
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'Converted.xlsx';
            a.click();
            window.URL.revokeObjectURL(url);
            document.getElementById('status').textContent = 'Download concluído!';
        } else {
            const errorText = await response.text();
            document.getElementById('status').textContent = `Erro: ${errorText}`;
        }
    } catch (error) {
        document.getElementById('status').textContent = `Erro de conexão: ${error.message}`;
    }
});