// forsetti.jsx
// Batch CSV → SKU → Template → Data Merge → Limpeza → Exportação
// Versão 2.1 (estabilizada)
// Dev: Alyssa Ferreiro @Sagittae-UX

// O plugin tem como objetivo automatizar a mesclagem de dados a partir de .csvs
// com templates do InDesign, aplicando limpeza e exportação em lote dos arquivos
// gerados, além de renomeá-los conforme padrão específico extraído do conteúdo.
// Alterar o caminho de arquivo das pastas de produção, templates e aplicar 
// o nome do diagramador no módulo de configuração.

(function () {

    // ======================================================
    // CONFIGURAÇÕES - MUDANÇAS AQUI
    // ======================================================

    var exportPreset = "Diagramacao2025";
    var mergeTarget = true;

    var entryFolder = Folder("~/Documents/PRODUCAO");
    var rootFolder = Folder("~/Documents/TEMPLATES");

    var csvTarget = /\.csv$/i;
    var userID = "Alyssa";

    // ======================================================
    // CONTADORES DE ERRO / CONTROLE
    // ======================================================

    var processedFiles = 0;
    var errorCount = 0;

    var missingTemplateCounter = {}; // SKU → [csvs]
    var validResults = {};    // fsName → true

    // ======================================================
    // VALIDAÇÕES INICIAIS
    // ======================================================

    if (!entryFolder.exists || !rootFolder.exists) {
        alert("Erro: Pastas de produção ou templates não encontradas. \nRedefina os caminhos ou cheque a localização.");
        return;
    }

    // ======================================================
    // BUSCA RECURSIVA DE CSVs
    // ======================================================

    function csvCollect(pasta, lista) {
        var itens = pasta.getFiles();
        for (var i = 0; i < itens.length; i++) {
            if (itens[i] instanceof File && csvTarget.test(itens[i].name)) {
                lista.push(itens[i]);
            } else if (itens[i] instanceof Folder) {
                csvCollect(itens[i], lista);
            }
        }
    }

    var csvFiles = [];
    csvCollect(entryFolder, csvFiles);

    if (!csvFiles.length) {
        alert("Nenhum CSV encontrado.");
        return;
    }

    // ======================================================
    // PARSER DE CSV
    // ======================================================

    function parseCSVLine(line) {
        var r = [], c = "", q = false;
        for (var i = 0; i < line.length; i++) {
            var ch = line.charAt(i);
            if (ch === '"') q = !q;
            else if (ch === "," && !q) { r.push(c); c = ""; }
            else c += ch;
        }
        r.push(c);
        return r;
    }

    function targetSKU(csv) {
        csv.encoding = "UTF-16";
        if (!csv.open("r")) return null;

        var txt = csv.read();
        csv.close();

        var csvCell = txt.split(/\r\n|\n|\r/);
        if (csvCell.length < 2) return null;

        var cols = parseCSVLine(csvCell[1]);
        if (cols.length < 2) return null;

        return cols[1].replace(/^\s+|\s+$/g, "");
    }

    // ======================================================
    // MESCLAGEM
    // ======================================================

    function mergeFile(docBase) {

        var root = {};
        for (var i = 0; i < app.documents.length; i++) {
            root[app.documents[i].id] = true;
        }

        try {
            docBase.dataMergeProperties.mergeRecords();
        } catch (_) {
            return null;
        }

        for (var j = 0; j < app.documents.length; j++) {
            if (!root[app.documents[j].id]) {
                return app.documents[j];
            }
        }

        return null;
    }

    // ======================================================
    // REGISTRO DO DIAGRAMADOR
    // ======================================================

    function userIdentifier(doc, nome) {
        var regex = /\bdiagramado_por_NOME\b/;
        for (var i = 0; i < doc.stories.length; i++) {
            if (regex.test(doc.stories[i].contents)) {
                doc.stories[i].contents =
                    doc.stories[i].contents.replace(regex, nome);
            }
        }
    }

    // ======================================================
    // MÓDULO - LIMPEZA DO DOCUMENTO
    // ======================================================

    function fileCleanup(doc) {

        app.findGrepPreferences = NothingEnum.nothing;
        app.changeGrepPreferences = NothingEnum.nothing;

        app.findGrepPreferences.findWhat = "\\\\#";
        app.changeGrepPreferences.changeTo = "";
        doc.changeGrep();

        app.findGrepPreferences.findWhat = "\\\\n";
        app.changeGrepPreferences.changeTo = "\\n";
        doc.changeGrep();

        app.findGrepPreferences = NothingEnum.nothing;
        app.changeGrepPreferences = NothingEnum.nothing;

        var items = doc.allPageItems;
        for (var i = items.length - 1; i >= 0; i--) {
            try {
                if (items[i] instanceof TextFrame &&
                    items[i].contents.replace(/\s+/g, "") === "") {
                    items[i].remove();
                }
            } catch (_) { }
        }
    }

    // ======================================================
    // RENOMEAÇÃO DE ARQUIVO
    // ======================================================

    function serialNumberGen(doc, fallback) {

        var regex = /^\d+\s*-\s*\d+_\d{5,}-[A-Z]{2,}\d*$/;

        for (var s = 0; s < doc.stories.length; s++) {
            var csvCell = doc.stories[s].contents.split(/[\r\n]+/);
            for (var l = 0; l < csvCell.length; l++) {
                var linha = csvCell[l].replace(/^\s+|\s+$/g, "");
                if (regex.test(linha)) {
                    return linha.replace(/[\\\/:*?"<>|]/g, "_");
                }
            }
        }

        return fallback.replace(/[\\\/:*?"<>|]/g, "_");
    }

    // ======================================================
    // MÓDULO - EXPORTAÇÃO EM LOTE
    // ======================================================

    for (var i = 0; i < csvFiles.length; i++) {

        var csv = csvFiles[i];

        // trava de CSV duplicado
        if (validResults[csv.fsName]) {
            continue;
        }
        validResults[csv.fsName] = true;

        var sku = targetSKU(csv);
        if (!sku) {
            errorCount++;
            continue;
        }

        var template = File(rootFolder + "/" + sku + ".indt");
        if (!template.exists) {

            if (!missingTemplateCounter[sku]) {
                missingTemplateCounter[sku] = [];
            }
            missingTemplateCounter[sku].push(csv.name);
            continue;
        }

        var docBase;
        try {
            docBase = app.open(template, false);
            docBase.dataMergeProperties.selectDataSource(csv);
        } catch (_) {
            errorCount++;
            try { docBase.close(SaveOptions.NO); } catch (_) { }
            continue;
        }

        var mergedDocument = mergeFile(docBase);
        if (!mergedDocument) {
            errorCount++;
            docBase.close(SaveOptions.NO);
            continue;
        }

        userIdentifier(mergedDocument, userID);
        fileCleanup(mergedDocument);

        var nomeFinal = serialNumberGen(mergedDocument, sku);
        var pasta = csv.parent;

        try {
            mergedDocument.save(File(pasta + "/" + nomeFinal + ".indd"));
            mergedDocument.exportFile(
                ExportFormat.pdfType,
                File(pasta + "/" + nomeFinal + ".pdf"),
                false,
                app.pdfExportPresets.itemByName(exportPreset)
            );

            processedFiles++;

            try {
                app.doScript(
                    'tell application "Finder"\n' +
                    'if (count of windows) > 0 then close front window\n' +
                    'end tell',
                    ScriptLanguage.applescriptLanguage
                );
            } catch (_) { }

        } catch (_) {
            errorCount++;
        }

        if (mergeTarget) mergedDocument.close(SaveOptions.NO);
        docBase.close(SaveOptions.NO);
    }

    // ======================================================
    // ALERT FINAL
    // ======================================================

    var msg =
        "Batch finalizado.\n\n" +
        "Processados: " + processedFiles + "\n" +
        "Erros: " + errorCount;

    var missingTemplates = false;
    for (var k in missingTemplateCounter) { missingTemplates = true; break; }

    if (missingTemplates) {
        msg += "\n\nTemplates faltando (SKU):";
        for (var k in missingTemplateCounter) {
            msg += "\n- " + k;
        }
    }

    alert(msg);

})();
