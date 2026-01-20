// forsetti.jsx
// Batch CSV → SKU → Template → Data Merge → Limpeza → Exportação
// Versão 2.1
// Dev: Alyssa Ferreiro @Sagittae-UX

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
    // LOG DE PROCESSO  (NOVO)
    // ======================================================

    var logFile = File(entryFolder + "/relatório.txt");

    function log(msg) {
        try {
            logFile.open("a");
            logFile.writeln(msg);
            logFile.close();
        } catch (_) { }
    }

    // cabeçalho do log
    log("\n========================================");
    log("INÍCIO DO LOTE: " + new Date());
    log("Usuário: " + userID);
    log("Pasta de entrada: " + entryFolder.fsName);
    log("========================================\n");

    // ======================================================
    // CONTADORES DE ERRO / CONTROLE
    // ======================================================

    var processedFiles = 0;
    var errorCount = 0;

    var missingTemplateCounter = {}; // SKU → [csvs]
    var pastasProcessadas = {}; // fsName da pasta → true

    // ======================================================
    // VALIDAÇÕES INICIAIS
    // ======================================================

    if (!entryFolder.exists || !rootFolder.exists) {
        alert("Erro: Pastas de produção ou templates não encontradas. \nRedefina os caminhos ou cheque a localização.");
        log("ERRO FATAL: Pastas de produção ou templates não encontradas.");
        return;
    }

    // ======================================================
    // BUSCA RECURSIVA DE CSVs
    // ======================================================

    function csvCollect(pastaRaiz) {

        var resultados = [];

        function percorrer(pasta) {

            var itens = pasta.getFiles();
            var encontrouCSV = false;

            for (var i = 0; i < itens.length; i++) {
                if (itens[i] instanceof File && csvTarget.test(itens[i].name)) {
                    resultados.push(itens[i]);
                    encontrouCSV = true;
                    break;
                }
            }

            if (!encontrouCSV) {
                for (var j = 0; j < itens.length; j++) {
                    if (itens[j] instanceof Folder) {
                        percorrer(itens[j]);
                    }
                }
            }
        }

        percorrer(pastaRaiz);
        return resultados;
    }

    var csvFiles = csvCollect(entryFolder);
    log("CSVs encontrados: " + csvFiles.length);

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
        var pastaPedido = csv.parent.fsName;

        if (pastasProcessadas[pastaPedido]) {
            log("IGNORADO (CSV duplicado na pasta): " + csv.fsName);
            continue;
        }

        pastasProcessadas[pastaPedido] = true;

        log("\n--- Processando CSV: " + csv.fsName);

        var sku = targetSKU(csv);
        if (!sku) {
            errorCount++;
            log("ERRO: Não foi possível ler SKU em " + csv.name);
            continue;
        }

        log("SKU identificado: " + sku);

        var template = File(rootFolder + "/" + sku + ".indt");
        if (!template.exists) {

            if (!missingTemplateCounter[sku]) {
                missingTemplateCounter[sku] = [];
            }
            missingTemplateCounter[sku].push(csv.name);

            log("TEMPLATE FALTANDO: " + sku + " (CSV: " + csv.name + ")");
            continue;
        }

        var docBase;
        try {
            docBase = app.open(template, false);
            docBase.dataMergeProperties.selectDataSource(csv);
        } catch (_) {
            errorCount++;
            log("ERRO: Falha ao abrir template ou associar CSV: " + csv.name);
            try { docBase.close(SaveOptions.NO); } catch (_) { }
            continue;
        }

        var mergedDocument = mergeFile(docBase);
        if (!mergedDocument) {
            errorCount++;
            log("ERRO: Falha na mesclagem: " + csv.name);
            docBase.close(SaveOptions.NO);
            continue;
        }

        try {
            if (mergedDocument.crossReferenceSources.length > 0) {
                mergedDocument.crossReferenceSources.everyItem().update();
                log("Referências cruzadas atualizadas.");
            }
        } catch (_) {
            log("AVISO: Falha ao atualizar referências cruzadas.");
        }

        userIdentifier(mergedDocument, userID);
        fileCleanup(mergedDocument);

        var nomeFinal = serialNumberGen(mergedDocument, sku);
        var pasta = csv.parent;

        try {
            var inddFile = File(pasta + "/" + nomeFinal + ".indd");
            var pdfFile = File(pasta + "/" + nomeFinal + ".pdf");

            mergedDocument.save(inddFile);
            mergedDocument.exportFile(
                ExportFormat.pdfType,
                pdfFile,
                false,
                app.pdfExportPresets.itemByName(exportPreset)
            );

            processedFiles++;

            log("EXPORTADO:");
            log("  INDD → " + inddFile.fsName);
            log("  PDF  → " + pdfFile.fsName);

            try {
                var as =
                    'tell application "Finder"\n' +
                    '    if (count of windows) > 0 then\n' +
                    '        close front window\n' +
                    '    end if\n' +
                    'end tell';

                app.doScript(as, ScriptLanguage.applescriptLanguage);
            } catch (e) { }

        } catch (e) {
            alert("Erro ao salvar/exportar o documento: " + e.message);
        }

        // FECHAR DOCUMENTO MESCLADO
        try {
            mergedDocument.close(SaveOptions.NO);
            log("Documento mesclado fechado.");
        } catch (_) {
            log("AVISO: Falha ao fechar documento mesclado.");
        }

        // FECHAR TEMPLATE BASE
        docBase.close(SaveOptions.NO);
    }

    // ======================================================
    // ALERT FINAL + FECHAMENTO DO LOG
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
            log("RESUMO - TEMPLATE FALTANDO: " + k);
        }
    }

    log("\n========================================");
    log("FIM DO LOTE: " + new Date());
    log("Processados: " + processedFiles);
    log("Erros: " + errorCount);
    log("========================================\n");

    // alert(msg);

})();
