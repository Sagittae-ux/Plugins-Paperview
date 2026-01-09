// descura.jsx
// Batch CSV → Data Merge robusto → Documento mesclado → Assinatura → AutoSave INDD + PDF
// Versão 1.0
// Dev: Alyssa Ferreiro @Sagittae-UX

(function () {

    // ======================================================
    // CONFIGURAÇÕES
    // ======================================================

    var NOME_USUARIO = "Alyssa";               // <<< ALTERE AQUI
    var PDF_PRESET = "Diagramacao2025";       // nome exato do preset
    var FECHAR_MESCLADO = true;                // true = fecha após exportar

    // ======================================================
    // VALIDA DOCUMENTO BASE
    // ======================================================

    if (app.documents.length === 0) {
        alert("Nenhum documento aberto.");
        return;
    }

    var docBase = app.activeDocument;

    if (!docBase.dataMergeProperties || !docBase.dataMergeProperties.isValid) {
        alert("Documento não possui Data Merge.");
        return;
    }

    // ======================================================
    // SELEÇÃO BATCH DE CSV
    // ======================================================

    var csvFiles = File.openDialog(
        "Selecione UM OU MAIS arquivos CSV",
        "*.csv",
        true
    );

    if (!csvFiles || csvFiles.length === 0) return;

    // ======================================================
    // FUNÇÃO: CRIA DOCUMENTO MESCLADO
    // ======================================================

    function criarDocumentoMesclado(doc) {

        var before = app.documents.length;

        try {
            if (typeof doc.dataMergeProperties.mergeRecords === "function") {
                doc.dataMergeProperties.mergeRecords();
            } else {
                doc.dataMergeProperties.createMergedDocument();
            }
        } catch (e) { }

        if (app.documents.length > before) {
            return app.activeDocument;
        }

        // Fallback por menuAction
        var fragments = [
            "create merged",
            "criar documento mescl",
            "documento mesclado",
            "documento combinado"
        ];

        var actions = app.menuActions;
        for (var i = 0; i < actions.length; i++) {
            var nm = actions[i].name.toLowerCase();
            for (var f = 0; f < fragments.length; f++) {
                if (nm.indexOf(fragments[f]) !== -1) {
                    try {
                        actions[i].invoke();
                        if (app.documents.length > before) {
                            return app.activeDocument;
                        }
                    } catch (_) { }
                }
            }
        }

        return null;
    }

    // ======================================================
    // LOOP BATCH
    // ======================================================

    for (var c = 0; c < csvFiles.length; c++) {

        var csv = csvFiles[c];
        $.writeln("=======================================");
        $.writeln("PROCESSANDO CSV: " + csv.fsName);

        // --------------------------------------------------
        // Atualiza Data Merge
        // --------------------------------------------------

        try {
            docBase.dataMergeProperties.selectDataSource(csv);
            try { docBase.dataMergeProperties.updateDataSource(); } catch (_) { }
        } catch (e) {
            $.writeln("[ERRO] Data Merge: " + e);
            continue;
        }

        // --------------------------------------------------
        // Cria documento mesclado
        // --------------------------------------------------

        var docMesclado = criarDocumentoMesclado(docBase);

        if (!docMesclado) {
            $.writeln("[ERRO] Documento mesclado não criado.");
            continue;
        }

        // --------------------------------------------------
        // Assinatura
        // --------------------------------------------------

        var marcador = /\bdiagramado_por_NOME\b/;
        var nomeBase = NOME_USUARIO;

        for (var s = 0; s < docMesclado.stories.length; s++) {
            if (marcador.test(docMesclado.stories[s].contents)) {
                docMesclado.stories[s].contents =
                    docMesclado.stories[s].contents.replace(marcador, NOME_USUARIO);
                break;
            }
        }

        // --------------------------------------------------
        // Detecta nome por RegEx
        // --------------------------------------------------

        var regexNome = /^\d+\s*-\s*\d+_\d{5,}-[A-Z]{2,}\d*$/;
        var nomeFinal = nomeBase;

        for (var t = 0; t < docMesclado.stories.length; t++) {
            var linhas = docMesclado.stories[t].contents.split(/[\r\n]/);
            for (var l = 0; l < linhas.length; l++) {
                var linha = linhas[l].replace(/^\s+|\s+$/g, "");
                if (regexNome.test(linha)) {
                    nomeFinal = linha;
                    break;
                }
            }
            if (nomeFinal !== nomeBase) break;
        }

        nomeFinal = nomeFinal.replace(/[\\\/:*?"<>|]/g, "_");

        // --------------------------------------------------
        // AutoSave na pasta do CSV
        // --------------------------------------------------

        var pastaDestino = csv.parent;
        var fileINDD = File(pastaDestino + "/" + nomeFinal + ".indd");
        var filePDF = File(pastaDestino + "/" + nomeFinal + ".pdf");

        try {
            docMesclado.save(fileINDD);
        } catch (e) {
            $.writeln("[ERRO] Save INDD: " + e);
            docMesclado.close(SaveOptions.NO);
            continue;
        }

        // --------------------------------------------------
        // Exporta PDF
        // --------------------------------------------------

        var preset = app.pdfExportPresets.itemByName(PDF_PRESET);
        if (!preset.isValid) {
            alert("Preset PDF não encontrado: " + PDF_PRESET);
            return;
        }

        try {
            docMesclado.exportFile(ExportFormat.pdfType, filePDF, false, preset);
        } catch (e) {
            $.writeln("[ERRO] PDF: " + e);
        }

        // --------------------------------------------------
        // Finalização
        // --------------------------------------------------

        if (FECHAR_MESCLADO) {
            docMesclado.close(SaveOptions.NO);
        }

        $.writeln("[OK] Finalizado: " + nomeFinal);
    }

    alert("Batch finalizado.\nVerifique o Console para detalhes.");

})();
