// gerar_csv_por_nome.jsx

var raiz = Folder.selectDialog("Selecione a pasta raiz");
if (!raiz) {
    alert("Nenhuma pasta selecionada.");
    exit();
}

// ===============================
// coleta recursiva de PDFs
// ===============================
function coletarPDFs(folder, lista) {
    var itens = folder.getFiles();

    for (var i = 0; i < itens.length; i++) {
        var item = itens[i];

        if (item instanceof Folder) {
            coletarPDFs(item, lista);
        } else if (item instanceof File && /\.pdf$/i.test(item.name)) {
            lista.push(item);
        }
    }
}

var arquivos = [];
coletarPDFs(raiz, arquivos);

if (arquivos.length === 0) {
    alert("Nenhum PDF encontrado.");
    exit();
}

// ===============================
// agrupar por nome base
// ===============================
var grupos = {};
var maxImgs = 0;

// regex remove sufixo _02, _03, etc.
var regexBase = /^(.*?)(?:_\d+)?\.pdf$/i;

for (var i = 0; i < arquivos.length; i++) {
    var nome = arquivos[i].name;
    var match = nome.match(regexBase);

    if (!match) continue;

    var base = match[1];

    if (!grupos[base]) {
        grupos[base] = [];
    }

    grupos[base].push(arquivos[i]);
}

// ordenar dentro de cada grupo e descobrir max colunas
for (var base in grupos) {
    grupos[base].sort(function (a, b) {
        return a.name.localeCompare(b.name, "pt-BR", { numeric: true });
    });

    if (grupos[base].length > maxImgs) {
        maxImgs = grupos[base].length;
    }
}

// ===============================
// cabeçalho dinâmico
// ===============================
var header = "PRODUTO";

for (var i = 0; i < maxImgs; i++) {
    header += ",@IMG" + (i === 0 ? "" : (i + 1));
}

var csv = header + "\r";

// ===============================
// gerar linhas
// ===============================
for (var base in grupos) {

    var linha = base;

    var listaArquivos = grupos[base];

    for (var j = 0; j < maxImgs; j++) {

        if (listaArquivos[j]) {
            var caminhoRelativo = listaArquivos[j].fsName
                .replace(raiz.fsName + "/", "")
                .replace(/\\/g, "/");

            linha += ",\"" + caminhoRelativo + "\"";
        } else {
            linha += ",";
        }
    }

    csv += linha + "\r";
}

// ===============================
// salvar CSV
// ===============================
var arquivoCSV = new File(raiz.fsName + "/@TEMPLATE.csv");
arquivoCSV.encoding = "UTF-16";
arquivoCSV.open("w");
arquivoCSV.write(csv);
arquivoCSV.close();

alert("CSV gerado com sucesso: " + arquivoCSV.fsName);