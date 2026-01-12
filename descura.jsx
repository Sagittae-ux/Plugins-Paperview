// descura.jsx
// Batch CSV (1 por subpasta final) → Data Merge → Limpeza → Exportação
// Versão 2.0 (clean / estável)
// Dev: Alyssa Ferreiro @Sagittae-UX

(function () {

    // ======================================================
    // CONFIGURAÇÕES
    // ======================================================

    var NOME_USUARIO = "Alyssa";
    var PDF_PRESET = "Diagramacao2025";
    var FECHAR_MESCLADO = true;

    var PASTA_CSV = Folder("~/Documents");
    var CSV_REGEX = /\.csv$/i;

    // ======================================================
    // VALIDAÇÕES INICIAIS
    // ======================================================

    if (app.documents.length === 0) {
        alert("Abra o documento base antes de rodar o script.");
        return;
    }

    var docBase = app.activeDocument;

    if (!docBase.dataMergeProperties || !docBase.dataMergeProperties.isValid) {
        alert("Documento não possui Data Merge.");
        return;
    }

    if (!PASTA_CSV.exists) {
        alert("Pasta não encontrada:\n" + PASTA_CSV.fsName);
        return;
    }

    // ======================================================
    // MÓDULO: COLETA DE CSVs (1 por pasta final)
    // ======================================================

    function coletarCSVsPorSubpastas(pastaRaiz) {

        var resultados = [];

        function percorrer(pasta) {
            var itens = pasta.getFiles();
            var achouCSV = false;

            for (var i = 0; i < itens.length; i++) {
                if (itens[i] instanceof File && CSV_REGEX.test(itens[i].name)) {
                    resultados.push(itens[i]);
                    achouCSV = true;
                    break; // apenas 1 CSV por pasta
                }
            }

            if (!achouCSV) {
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

    var csvFiles = coletarCSVsPorSubpastas(PASTA_CSV);

    if (csvFiles.length === 0) {
        alert("Nenhum CSV encontrado.");
        return;
    }

    // ======================================================
    // MÓDULO: MESCLAGEM SEGURA
    // ======================================================

    function criarDocumentoMesclado(doc) {

        var antes = {};
        for (var i = 0; i < app.documents.length; i++) {
            antes[app.documents[i].id] = true;
        }

        try {
            doc.dataMergeProperties.mergeRecords();
        } catch (_) {
            return null;
        }

        for (var j = 0; j < app.documents.length; j++) {
            var d = app.documents[j];
            if (!antes[d.id]) return d;
        }

        return null;
    }

    // ======================================================
    // MÓDULO: LIMPEZA DO DOCUMENTO
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
            var item = items[i];

            try {
                if (item instanceof TextFrame && !item.locked) {

                    if (item.contents.replace(/\s+/g, "") === "") {
                        item.remove();
                        continue;
                    }

                    if (item.overflows) {
                        var t = 0;
                        while (item.overflows && t < 20) {
                            item.geometricBounds = [
                                item.geometricBounds[0],
                                item.geometricBounds[1],
                                item.geometricBounds[2] + 100,
                                item.geometricBounds[3]
                            ];
                            t++;
                        }
                    }
                }

                if (item instanceof Rectangle && !item.locked) {
                    if (item.allGraphics.length === 0) {
                        item.remove();
                    }
                }
            } catch (_) { }
        }
    }

    // ======================================================
    // LOOP BATCH
    // ======================================================

    for (var c = 0; c < csvFiles.length; c++) {

        var csv = csvFiles[c];
        $.writeln("PROCESSANDO: " + csv.fsName);

        try {
            docBase.dataMergeProperties.selectDataSource(csv);
        } catch (_) {
            $.writeln("[ERRO] Data Merge: " + csv.name);
            continue;
        }

        var docMesclado = criarDocumentoMesclado(docBase);

        if (!docMesclado || !docMesclado.isValid || docMesclado.stories.length === 0) {
            $.writeln("[ERRO] Mescla inválida: " + csv.name);
            try { docMesclado.close(SaveOptions.NO); } catch (_) { }
            continue;
        }

        limparDocumento(docMesclado);

        // Nome do arquivo
        var nomeFinal = NOME_USUARIO;
        var regexNome = /^\d+\s*-\s*\d+_\d{5,}-[A-Z]{2,}\d*$/;

        for (var s = 0; s < docMesclado.stories.length; s++) {
            var linhas = docMesclado.stories[s].contents.split(/[\r\n]/);
            for (var l = 0; l < linhas.length; l++) {
                var linha = linhas[l].replace(/^\s+|\s+$/g, "");
                if (regexNome.test(linha)) {
                    nomeFinal = linha;
                    break;
                }
            }
        }

        nomeFinal = nomeFinal.replace(/[\\\/:*?"<>|]/g, "_");

        var pasta = csv.parent;
        var indd = File(pasta + "/" + nomeFinal + ".indd");
        var pdf = File(pasta + "/" + nomeFinal + ".pdf");

        try {
            docMesclado.save(indd);

            var preset = app.pdfExportPresets.itemByName(PDF_PRESET);
            docMesclado.exportFile(ExportFormat.pdfType, pdf, false, preset);
            function fecharJanelaFinder() {
                try {
                    var as =
                        'tell application "Finder"\n' +
                        '    if (count of windows) > 0 then\n' +
                        '        close front window\n' +
                        '    end if\n' +
                        'end tell';
                    app.doScript(as, ScriptLanguage.applescriptLanguage);
                } catch (_) {
                    // silencioso
                }
            }

            fecharJanelaFinder();

        } catch (e) {
            $.writeln("[ERRO] Exportação: " + e);
        }

        if (FECHAR_MESCLADO) {
            docMesclado.close(SaveOptions.NO);
        }

        $.writeln("[OK] Finalizado: " + nomeFinal);
    }

    app.activate();
    $.writeln("=== PROCESSO FINALIZADO ===");

})();
