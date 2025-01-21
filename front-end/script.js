document.getElementById('uploadForm').addEventListener('submit', async (event) => {
    event.preventDefault();

    const fileInput = document.getElementById('xmlFile');
    const file = fileInput.files[0];

    if (!file) {
        alert('Favor, escolher um arquivo NF-e XML válido!');
        return;
    }

    const formData = new FormData();
    formData.append('file', file);

    try {
        const response = await fetch('http://localhost:5148/convert', {
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