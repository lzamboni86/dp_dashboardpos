document.addEventListener('DOMContentLoaded', () => {
    const excelUpload = document.getElementById('excelUpload');
    
    excelUpload.addEventListener('change', function(e) {
        const file = e.target.files[0];
        const reader = new FileReader();
        
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data);
                // Restante do código de processamento...
                console.log('Arquivo processado com sucesso!');
            } catch (error) {
                console.error('Erro ao processar arquivo:', error);
            }
        };
        
        reader.onerror = function(error) {
            console.error('Erro na leitura do arquivo:', error);
        };
        
        reader.readAsArrayBuffer(file);
    });
});