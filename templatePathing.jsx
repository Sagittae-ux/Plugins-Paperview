// Ideia inicial de script de processamento

var pasta = Folder.selectDialog("Selecione a pasta com os PDFs");
if (!pasta) {
    alert("Nenhuma pasta selecionada.");
    exit();
}

// busca de .pdf
var arquivos = pasta.getFiles(/\.pdf$/i);

if (arquivos.length === 0) {
    alert("Nenhum PDF encontrado.");
    exit();
}

// ordenar arquivos de forma consistente
arquivos.sort(function (a, b) {
    return a.name.localeCompare(b.name, "pt-BR", { numeric: true });
});

// header
var csv = "PRODUTO,@IMG\r";

for (var i = 0; i < arquivos.length; i++) {
    var nomeArquivo = arquivos[i].name;

    // remove extensão .pdf
    var produto = nomeArquivo.replace(/\.pdf$/i, "");

    // estabelecer caminho relativo
    var caminho = nomeArquivo;

    csv += produto + ",\"" + caminho + "\"\r";
}

// salva CSV na mesma pasta, necessário para que o arquivo em caminho relativo continue funcionando
var arquivoCSV = new File(pasta.fsName + "/dados.csv");
arquivoCSV.encoding = "UTF-16";
arquivoCSV.open("w");
arquivoCSV.write(csv);
arquivoCSV.close();

alert("CSV gerado com sucesso: " + arquivoCSV.fsName);