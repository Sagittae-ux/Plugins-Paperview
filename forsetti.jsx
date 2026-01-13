// forsetti.jsx
// Batch CSV → SKU → Template → Data Merge → Limpeza → Exportação
// Versão 2.1 (produção consolidada)
// Dev: Alyssa Ferreiro @Sagittae-UX

(function () {

    // ======================================================
    // CONFIGURAÇÕES
    // ======================================================

    var PDF_PRESET = "Diagramacao2025";
    var FECHAR_MESCLADO = true;

    var PASTA_PRODUCAO = Folder("~/Documents/PRODUCAO");
    var PASTA_TEMPLATES = Folder("~/Documents/TEMPLATES");

    var CSV_REGEX = /\.csv$/i;
    var nomeUsuario = "Alyssa";

    // ======================================================
    // CONTADORES / CONTROLE
    // ======================================================

    var CONT_PROCESSADOS = 0;
    var CONT_ERROS = 0;

    var TEMPLATES_FALTANDO = {}; // SKU → [csvs]
    var CSV_PROCESSADOS = {};    // fsName → true

    // ======================================================
    // VALIDAÇÕES
    // ======================================================

    if (!PASTA_PRODUCAO.exists || !PASTA_TEMPLATES.exists) {
        alert("Pastas de produção ou templates não encontradas.");
        return;
    }

    // ======================================================
    // BUSCA RECURSIVA DE CSVs
    // ======================================================

    function coletarCSVs(pasta, lista) {
        var itens = pasta.getFiles();
        for (var i = 0; i < itens.length; i++) {
            if (itens[i] instanceof File && CSV_REGEX.test(itens[i].name)) {
                lista.push(itens[i]);
            } else if (itens[i] instanceof Folder) {
                coletarCSVs(itens[i], lista);
            }
        }
    }

    var csvFiles = [];
    coletarCSVs(PASTA_PRODUCAO, csvFiles);

    if (!csvFiles.length) {
        alert("Nenhum CSV encontrado.");
        return;
    }

    // ======================================================
    // CSV → SKU (2ª COLUNA)
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

    function extrairSKU(csv) {
        csv.encoding = "UTF-16";
        if (!csv.open("r")) return null;

        var txt = csv.read();
        csv.close();

        var linhas = txt.split(/\r\n|\n|\r/);
        if (linhas.length < 2) return null;

        var cols = parseCSVLine(linhas[1]);
        if (cols.length < 2) return null;

        return cols[1].replace(/^\s+|\s+$/g, "");
    }

    // ======================================================
    // MESCLAGEM SEGURA
    // ======================================================

    function criarDocumentoMesclado(docBase) {

        var antes = {};
        for (var i = 0; i < app.documents.length; i++) {
            antes[app.documents[i].id] = true;
        }

        try {
            docBase.dataMergeProperties.mergeRecords();
        } catch (_) {
            return null;
        }

        for (var j = 0; j < app.documents.length; j++) {
            if (!antes[app.documents[j].id]) {
                return app.documents[j];
            }
        }

        return null;
    }

    // ======================================================
    // REGISTRO DO DIAGRAMADOR
    // ======================================================

    function aplicarNomeDiagramador(doc, nome) {
        var regex = /\bdiagramado_por_NOME\b/;
        for (var i = 0; i < doc.stories.length; i++) {
            if (regex.test(doc.stories[i].contents)) {
                doc.stories[i].contents =
                    doc.stories[i].contents.replace(regex, nome);
            }
        }
    }

    // ======================================================
    // LIMPEZA DO DOCUMENTO
    // ======================================================

    function limparDocumento(doc) {

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
    // NOME FINAL (REGEX NA OP)
    // ======================================================

    function extrairNomeFinal(doc, fallback) {

        var regex = /^\d+\s*-\s*\d+_\d{5,}-[A-Z]{2,}\d*$/;

        for (var s = 0; s < doc.stories.length; s++) {
            var linhas = doc.stories[s].contents.split(/[\r\n]+/);
            for (var l = 0; l < linhas.length; l++) {
                var linha = linhas[l].replace(/^\s+|\s+$/g, "");
                if (regex.test(linha)) {
                    return linha.replace(/[\\\/:*?"<>|]/g, "_");
                }
            }
        }

        return fallback.replace(/[\\\/:*?"<>|]/g, "_");
    }

    // ======================================================
    // LOOP PRINCIPAL
    // ======================================================

    for (var i = 0; i < csvFiles.length; i++) {

        var csv = csvFiles[i];

        // 🔒 trava de CSV duplicado
        if (CSV_PROCESSADOS[csv.fsName]) {
            continue;
        }
        CSV_PROCESSADOS[csv.fsName] = true;

        var sku = extrairSKU(csv);
        if (!sku) {
            CONT_ERROS++;
            continue;
        }

        var template = File(PASTA_TEMPLATES + "/" + sku + ".indt");
        if (!template.exists) {

            if (!TEMPLATES_FALTANDO[sku]) {
                TEMPLATES_FALTANDO[sku] = [];
            }
            TEMPLATES_FALTANDO[sku].push(csv.name);
            continue;
        }

        var docBase;
        try {
            docBase = app.open(template, false);
            docBase.dataMergeProperties.selectDataSource(csv);
        } catch (_) {
            CONT_ERROS++;
            try { docBase.close(SaveOptions.NO); } catch (_) { }
            continue;
        }

        var docMesclado = criarDocumentoMesclado(docBase);
        if (!docMesclado) {
            CONT_ERROS++;
            docBase.close(SaveOptions.NO);
            continue;
        }

        aplicarNomeDiagramador(docMesclado, nomeUsuario);
        limparDocumento(docMesclado);

        var nomeFinal = extrairNomeFinal(docMesclado, sku);
        var pasta = csv.parent;

        try {
            docMesclado.save(File(pasta + "/" + nomeFinal + ".indd"));
            docMesclado.exportFile(
                ExportFormat.pdfType,
                File(pasta + "/" + nomeFinal + ".pdf"),
                false,
                app.pdfExportPresets.itemByName(PDF_PRESET)
            );

            CONT_PROCESSADOS++;

            try {
                app.doScript(
                    'tell application "Finder"\n' +
                    'if (count of windows) > 0 then close front window\n' +
                    'end tell',
                    ScriptLanguage.applescriptLanguage
                );
            } catch (_) { }

        } catch (_) {
            CONT_ERROS++;
        }

        if (FECHAR_MESCLADO) docMesclado.close(SaveOptions.NO);
        docBase.close(SaveOptions.NO);
    }

    // ======================================================
    // ALERT FINAL
    // ======================================================

    var msg =
        "Batch finalizado.\n\n" +
        "Processados: " + CONT_PROCESSADOS + "\n" +
        "Erros: " + CONT_ERROS;

    var temFaltantes = false;
    for (var k in TEMPLATES_FALTANDO) { temFaltantes = true; break; }

    if (temFaltantes) {
        msg += "\n\nTemplates faltando (SKU):";
        for (var k in TEMPLATES_FALTANDO) {
            msg += "\n- " + k;
        }
    }

    alert(msg);

})();
