// Event listener para o botão "Selecionar Arquivo Excel"
document.getElementById('select-file').addEventListener('click', function() {
    // Obtendo o número de linhas por arquivo inserido pelo usuário
    let numLinhas = parseInt(document.getElementById('num-linhas').value);
    // Verificando se o número de linhas é válido
    if (isNaN(numLinhas) || numLinhas < 1) {
        alert("Por favor, insira um número válido de linhas por arquivo.");
        return;
    }

    // Criando um input de arquivo para selecionar o arquivo Excel
    let fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.accept = '.xlsx, .xls';
    fileInput.onchange = function(event) {
        let file = event.target.files[0];
        if (!file) return;
        let reader = new FileReader();
        reader.onload = function(e) {
            try {
                let data = new Uint8Array(e.target.result);
                // Lendo o arquivo Excel
                let workbook = XLSX.read(data, {type: 'array'});
                if (workbook.SheetNames.length === 0) {
                    throw new Error("Nenhuma planilha encontrada no arquivo.");
                }
                let sheetName = workbook.SheetNames[0];
                let worksheet = workbook.Sheets[sheetName];
                let json = XLSX.utils.sheet_to_json(worksheet, {header: 1});
                let header = json[0]; // Pegando o cabeçalho
                json.shift(); // Removendo o cabeçalho das linhas de dados
                let totalLinhas = json.length;
                let numPedaços = Math.ceil(totalLinhas / numLinhas);
                let i = 0;
                // Utilizando um intervalo para gerar um arquivo
                let interval = setInterval(function() {
                    if (i >= numPedaços) {
                        clearInterval(interval);
                        alert(`Arquivo dividido em ${numPedaços} partes.`);
                        return;
                    }
                    let inicio = i * numLinhas;
                    let fim = Math.min((i + 1) * numLinhas, totalLinhas);
                    let dadosPedaço = json.slice(inicio, fim);
                    dadosPedaço.unshift(header); // Incluindo o cabeçalho nas linhas de dados
                    let novoWorkbook = XLSX.utils.book_new();
                    let novoWorksheet = XLSX.utils.aoa_to_sheet(dadosPedaço);
                    XLSX.utils.book_append_sheet(novoWorkbook, novoWorksheet, 'Sheet1');
                    let nomeArquivo = `Arquivo_particionado_${i + 1}.xlsx`;
                    let blob = new Blob([s2ab(XLSX.write(novoWorkbook, {bookType: 'xlsx', type: 'binary'}))], {type: "application/octet-stream"});
                    let url = URL.createObjectURL(blob);
                    let link = document.createElement('a');
                    link.href = url;
                    link.download = nomeArquivo;
                    document.body.appendChild(link);
                    link.click();
                    document.body.removeChild(link);
                    i++;
                }, 1000); // Um segundo de atraso -- Sem isso o programa roda muito rápido e compromete a quantidade de planilhas.
            } catch (error) {
                console.error("Erro ao processar o arquivo:", error);
                alert('Ocorreu um erro ao processar o arquivo. Verifique se o arquivo está correto.');
            }
        };
        reader.onerror = function(error) {
            console.error("Erro ao ler o arquivo:", error);
            alert('Ocorreu um erro ao ler o arquivo. Tente novamente.');
        };
        reader.readAsArrayBuffer(file);
    };
    fileInput.click();
});

// Função auxiliar para converter string para array buffer
function s2ab(s) {
    var buf = new ArrayBuffer(s.length);
    var view = new Uint8Array(buf);
    for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
    return buf;
}
