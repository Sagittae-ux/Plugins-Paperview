// .jsx
// Módulo avulso de processamento de arquivos .csv, a ser instalado no motor ignisCalor.jsx se aprovado pela gerência. 
// Versão 1.0
// Dev: Alyssa Ferreiro @Sagittae-UX

var rootFolder = new Folder(Folder.myDocuments + "/PRODUCAO");

function titleCase(str) {

    if (!str) return str;

    // normaliza espaços e força caixa baixa
    str = str.replace(/^\s+|\s+$/g, "").toLowerCase();

    return str;
}

// Parser CSV completo: retorna array de registros, cada registro é array de campos
function parseCSV(content) {
    var records = [];
    var i = 0;
    var len = content.length;
    var field = "";
    var record = [];
    var inQuotes = false;

    while (i < len) {
        var ch = content[i];

        if (inQuotes) {
            if (ch === '"') {
                // peek next char to see if it's a double quote (escaped quote)
                var next = content[i + 1];
                if (next === '"') {
                    field += '"';
                    i += 2;
                    continue;
                } else {
                    // closing quote
                    inQuotes = false;
                    i++;
                    continue;
                }
            } else {
                field += ch;
                i++;
                continue;
            }
        } else {
            if (ch === '"') {
                inQuotes = true;
                i++;
                continue;
            } else if (ch === ',') {
                record.push(field);
                field = "";
                i++;
                continue;
            } else if (ch === '\r') {
                var next = content[i + 1];
                if (next === '\n') i++;
                record.push(field);
                records.push(record);
                record = [];
                field = "";
                i++;
                continue;
            } else if (ch === '\n') {
                // lone LF
                record.push(field);
                records.push(record);
                record = [];
                field = "";
                i++;
                continue;
            } else {
                field += ch;
                i++;
                continue;
            }
        }
    }

    record.push(field);

    if (!(record.length === 1 && record[0] === "" && records.length > 0)) {
        records.push(record);
    }
    return records;
}

function buildCSV(records) {
    var lines = [];
    for (var r = 0; r < records.length; r++) {
        var cols = records[r];
        var outCols = [];
        for (var c = 0; c < cols.length; c++) {
            var val = cols[c] == null ? "" : String(cols[c]);
            var needsQuote = val.indexOf('"') >= 0 || val.indexOf(',') >= 0 || val.indexOf('\n') >= 0 || val.indexOf('\r') >= 0 || /^\s|\s$/.test(val);
            if (val.indexOf('"') >= 0) val = val.replace(/"/g, '""');
            if (needsQuote) val = '"' + val + '"';
            outCols.push(val);
        }
        lines.push(outCols.join(","));
    }
    return lines.join("\r\n");
}

function processCSV(file) {
    if (!file.exists) return;
    if (!file.open("r")) return;
    var content = file.read();
    file.close();

    var records = parseCSV(content);
    if (!records || records.length < 2) return;

    var headers = records[0];
    var nomeIndex = -1;
    var targetHeaders = {
        "DIGITE_O_NOME_A_SER_IMPRESSO": true,
        "DIGITE_O_NOME_A_SER_IMPRESSO_NA_CAPA": true
    };

    for (var h = 0; h < headers.length; h++) {

        var header = headers[h]
            .replace(/^"|"$/g, "")
            .replace(/^\s+|\s+$/g, "");

        if (targetHeaders[header]) {
            nomeIndex = h;
            break;
        }
    }
    if (nomeIndex === -1) return;

    for (var r = 1; r < records.length; r++) {
        var row = records[r];
        if (!row || row.length === 0) continue;
        var val = row[nomeIndex];
        if (val) {
            row[nomeIndex] = titleCase(val);
        }
        records[r] = row;
    }

    var backup = new File(file.fullName + ".bak");
    try {

        if (backup.exists) backup.remove();
        file.copy(backup);
    } catch (e) {

    }

    if (!file.open("w")) return;
    file.encoding = "UTF-16";
    file.lineFeed = "Windows";
    file.write(buildCSV(records));
    file.close();
}

function scanFolder(folder) {

    if (!folder.exists) return;

    var items = folder.getFiles();
    var csvProcessed = false;

    for (var i = 0; i < items.length; i++) {

        var item = items[i];

        if (item instanceof Folder) {

            // continua descendo na árvore
            scanFolder(item);

        } else if (!csvProcessed && item instanceof File && item.name.match(/\.csv$/i)) {

            processCSV(item);
            csvProcessed = true;
        }
    }
}

scanFolder(rootFolder);