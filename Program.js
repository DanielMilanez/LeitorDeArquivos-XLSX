document.getElementById('processButton').addEventListener('click', function() {
    const fileInput = document.getElementById('fileInput');
    const file = fileInput.files[0];

    if (!file) {
        alert('Por favor, selecione um arquivo XLSX.');
        return;
    }

    const reader = new FileReader();

    reader.onload = function(event) {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];

        let output = document.getElementById('output');
        output.innerHTML = '';

        // Encontrar o intervalo de células usado no arquivo
        const range = XLSX.utils.decode_range(worksheet['!ref']);

        // Verificar o valor das células na primeira linha (linha 0)
        const firstRow = 0;
        const idHeader = worksheet[XLSX.utils.encode_cell({ r: firstRow, c: 0 })]?.v;
        const nameHeader = worksheet[XLSX.utils.encode_cell({ r: firstRow, c: 1 })]?.v;

        // Se os cabeçalhos estiverem presentes, pule para a próxima linha (linha 1)
        let startRow = firstRow;
        if (idHeader === 'ID' && nameHeader === 'Name') {
            startRow = firstRow + 1;
        }

        // Criar o conteúdo para o arquivo TXT
        let txtContent = '';
        for (let row = startRow; row <= range.e.r; row++) {
            // Ler cada célula da coluna A (índice 0) e B (índice 1)
            const id = worksheet[XLSX.utils.encode_cell({ r: row, c: 0 })]?.v;
            const name = worksheet[XLSX.utils.encode_cell({ r: row, c: 1 })]?.v;

            if (id !== undefined && name !== undefined) {
                output.innerHTML += `<p>ID: ${id}, Name: ${name}</p>`;
                txtContent += `ID: ${id}, Name: ${name}\n`;
            }
        }

        // Criar um link temporário para download do arquivo TXT
        const blob = new Blob([txtContent], { type: 'text/plain' });
        const url = URL.createObjectURL(blob);
        const downloadLink = document.createElement('a');
        downloadLink.href = url;
        downloadLink.download = 'Result.txt';
        downloadLink.textContent = 'Download do arquivo TXT';

        // Adicionar o evento de clique no botão de download
        downloadLink.addEventListener('click', function() {
            // Remover o link de download após o download ser concluído
            setTimeout(() => URL.revokeObjectURL(url), 1000);
        });

        // Adicionar o link de download no elemento com o id "output"
        output.appendChild(downloadLink);
    };

    reader.readAsArrayBuffer(file);
});