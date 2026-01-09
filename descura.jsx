// descura.jsx
// Batch CSV (1 por subpasta) → Data Merge → Mesclado → Assinatura → AutoSave
// Versão 1.2
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
    // VALIDA DOCUMENTO BASE
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
    // FUNÇÃO: PRIMEIRO CSV RECURSIVO
    // ======================================================

    function primeiroCSVRecursivo(pasta) {
        var itens = pasta.getFiles();
        for (var i = 0; i < itens.length; i++) {
            if (itens[i] instanceof File && CSV_REGEX.test(itens[i].name)) {
                return itens[i];
            }
            if (itens[i] instanceof Folder) {
                var achado = primeiroCSVRecursivo(itens[i]);
                if (achado) return achado;
            }
        }
        return null;
    }

    // ======================================================
    // COLETA: 1 CSV POR SUBPASTA
    // ======================================================

    var csvFiles = [];
    var subpastas = PASTA_CSV.getFiles(function (f) {
        return f instanceof Folder;
    });

    for (var i = 0; i < subpastas.length; i++) {
        var csv = primeiroCSVRecursivo(subpastas[i]);
        if (csv) {
            csvFiles.push(csv);
            $.writeln("[OK] CSV encontrado: " + csv.fsName);
        } else {
            $.writeln("[INFO] Nenhum CSV em: " + subpastas[i].fsName);
        }
    }

    if (csvFiles.length === 0) {
        alert("Nenhum CSV encontrado nas subpastas.");
        return;
    }

    // ======================================================
    // FUNÇÃO: CRIAR DOCUMENTO MESCLADO
    // ======================================================

    function criarDocumentoMesclado(docBase) {

        // Guarda IDs dos documentos abertos ANTES
        var docsAntes = {};
        for (var i = 0; i < app.documents.length; i++) {
            docsAntes[app.documents[i].id] = true;
        }

        try {
            if (typeof docBase.dataMergeProperties.mergeRecords === "function") {
                docBase.dataMergeProperties.mergeRecords();
            } else {
                docBase.dataMergeProperties.createMergedDocument();
            }
        } catch (e) {
            return null;
        }

        // Procura o documento que NÃO existia antes
        for (var j = 0; j < app.documents.length; j++) {
            var d = app.documents[j];
            if (!docsAntes[d.id]) {
                return d; // ← este é o mesclado
            }
        }

        return null;
    }


    // ======================================================
    // LOOP BATCH
    // ======================================================

    for (var c = 0; c < csvFiles.length; c++) {

        var csv = csvFiles[c];

        $.writeln("====================================");
        $.writeln("PROCESSANDO: " + csv.name);

        try {
            docBase.dataMergeProperties.selectDataSource(csv);
            docBase.dataMergeProperties.updateDataSource();
        } catch (e) {
            $.writeln("[ERRO] Data Merge: " + e);
            continue;
        }

        var docMesclado = null;

        try {
            docMesclado = criarDocumentoMesclado(docBase);
        } catch (e) {
            $.writeln("[ERRO] Falha ao mesclar CSV: " + csv.name);
            continue;
        }

        if (!docMesclado || !docMesclado.isValid || docMesclado.stories.length === 0) {
            $.writeln("[ERRO] Mescla inválida (SKU incompatível?): " + csv.name);
            try {
                if (docMesclado && docMesclado.isValid) {
                    docMesclado.close(SaveOptions.NO);
                }
            } catch (_) { }
            continue;
        }


        // Assinatura
        var marcador = /\bdiagramado_por_NOME\b/;
        try {
            for (var s = 0; s < docMesclado.stories.length; s++) {
                if (marcador.test(docMesclado.stories[s].contents)) {
                    docMesclado.stories[s].contents =
                        docMesclado.stories[s].contents.replace(marcador, NOME_USUARIO);
                    break;
                }
            }
        } catch (e) {
            $.writeln("[AVISO] Falha ao aplicar assinatura: " + csv.name);
        }


        // Nome
        var nomeFinal = NOME_USUARIO;
        var regexNome = /^\d+\s*-\s*\d+_\d{5,}-[A-Z]{2,}\d*$/;

        for (var t = 0; t < docMesclado.stories.length; t++) {
            var linhas = docMesclado.stories[t].contents.split(/[\r\n]/);
            for (var l = 0; l < linhas.length; l++) {
                var linha = linhas[l]
                    .replace(/\s+/g, " ")
                    .replace(/^\s+|\s+$/g, "");
                ;
                if (regexNome.test(linha)) {
                    nomeFinal = linha;
                    break;
                }
            }
        }

        nomeFinal = nomeFinal.replace(/[\\\/:*?"<>|]/g, "_");

        var pastaDestino = csv.parent;
        var indd = File(pastaDestino + "/" + nomeFinal + ".indd");
        var pdf = File(pastaDestino + "/" + nomeFinal + ".pdf");

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
