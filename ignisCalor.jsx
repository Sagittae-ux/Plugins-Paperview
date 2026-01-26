// ignisCalor.jsx
// Batch CSV → SKU → Template → Data Merge → Limpeza → Exportação
// Versão 3.0
// Dev: Alyssa Ferreiro @Sagittae-UX

// Esse script automatiza o processo de diagramação em lote no Adobe InDesign,
// atuando como o novo motor central para a tarefa de mesclagem preliminar de
// processamento de pedidos.


// RECURSOS PRINCIPAIS:
// - Busca recursiva de arquivos CSV em subpastas.
// - Identificação automática de SKU a partir do conteúdo do CSV através de parsing da segunda coluna.
// - Associação dinâmica de templates baseados no SKU detectado na pasta raiz com as bases revisadas.
// - Mesclagem de dados utilizando o recurso Data Merge do InDesign.
// - Limpeza automática do documento mesclado (remoção de marcadores, quadros vazios, etc.).
// - Exportação em lote para formatos INDD e PDF com presets definidos.
// - Registro detalhado de processos, erros e SKUs ignorados em um arquivo de log.


// INSTRUÇÕES DE USO:
// 1. Configure as variáveis no início do script:
//    - exportPreset: Nome do preset de exportação PDF padronizado para produção, escrito de forma exata.
//    - entryFolder: Caminho da pasta de produção contendo os pedidos baixados do Magento.
//    - rootFolder: Caminho da pasta de templates contendo os arquivos .indt.
//    - userID: Nome do diagramador para registro no documento exportado.

// 2. Adicione SKUs à lista negra quando necessário. Os arquivos serão catalogados
//    e movidos para a pasta "_IGNORADOS", onde deverão ser processados manualmente.

// 3. Adicione o script à pasta de scripts do Adobe InDesign através da janela de utilitários de scripts.
//    Recomenda-se criar uma subpasta específica para scripts personalizados, e um atalho para fácil execução.

// 4. Abrir o Adobe InDesign, na janela de utilitários de scripts, executar o script.

// 5. Ao final do processamento, um relatório será gerado na pasta de entrada,
//    detalhando o número de arquivos processados, erros encontrados e SKUs ignorados.
//    Após o processamento, checar a pasta "_IGNORADOS" para itens que necessitam de diagramação manual
//    e envie os itens para o fechamento.

// 6. Faça o preflight dos arquivos, atentando-se a erros estéticos.

// 7. Caso faltem templates para determinados SKUs, esses serão listados no relatório final.
//    Ajustar e incluir bases conforme necessário, adicionando o SKU entre aspas e respeitando a sintaxe abaixo:

// ignoredSKUs {
//     "MD890": true,
//     "AB123": true
// };

// Em caso de dúvidas ou necessidade de suporte, entre em contato com Alyssa Ferreiro @Sagittae-UX >:3c


(function () {

    // ======================================================
    // CONFIGURAÇÕES - MUDANÇAS AQUI
    // ======================================================

    var exportPreset = "Diagramacao2025"; //Preset exata padronizada para produção

    var entryFolder = Folder("~/Documents/PRODUCAO"); //Colar o caminho de arquivo da pasta de entrada aqui
    var rootFolder = Folder("~/Documents/TEMPLATES"); //Colar o caminho de arquivo da pasta de templates aqui

    var userID = "Alyssa"; //Nome do diagramador para registro no documento exportado

    var csvTarget = /\.csv$/i;

    // ======================================================
    // LISTA NEGRA - SKUs PARA PROCESSAMENTO MANUAL AQUI
    // ======================================================

    // Referir-se ao passo 5 das instruções acima para sintaxe correta.
    // Lista reservada para itens problemáticos ou que exigem diagramação especial.

    var ignoredSKUs = { //Adicionar itens para a lista negra aqui
        "MD890": true
    };

    // ======================================================
    // LOG DE PROCESSO
    // ======================================================

    var logFile = File(entryFolder + "/relatório.txt"); // Arquivo de log na pasta de entrada

    function log(msg) {
        try {
            logFile.open("a");
            logFile.writeln(msg);
            logFile.close();
        } catch (_) { }
    }

    // Cabeçalho do log
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
    var totalBlacklistedFiles = 0;
    var totalBlacklistedSKUs = 0;


    var missingTemplateCounter = {}; // Indicados no fim do log com o código de SKU
    var outputFolder = {}; // Número de pedidos processados com sucesso
    var blacklistCounter = {}; // Contador de SKUs listados para processamento manual

    // ======================================================
    // PASTA DE IGNORADOS
    // ======================================================   

    var ignoredFolder = Folder(entryFolder + "/_IGNORADOS");
    if (!ignoredFolder.exists) {
        ignoredFolder.create();
    }

    // ======================================================
    // VALIDAÇÕES INICIAIS
    // ======================================================

    // Check para pastas de entrada e templates. Tomar cuidado ao alterar os caminhos
    // pois o script será abortado caso estejam incorretos.
    if (!entryFolder.exists || !rootFolder.exists) {
        alert("Erro: O caminho da pasta de produção ou templates não foi encontrado.\nVerifique as configurações no início do script e cole o caminho de arquivo na linha 'var entryFolder' e 'var rootFolder'.");
        log("Erro: Caminho de pasta inválido.");
        return;
    }

    function isIgnoredSKU(sku) {
        return ignoredSKUs[sku] === true;
    }

    // Engine de processamento separado de itens da lista negra.
    function blacklistFile(sku, csvFile, reason) {

        if (!blacklistCounter[sku]) {
            blacklistCounter[sku] = {
                count: 0,
                files: []
            };
            totalBlacklistedSKUs++;
        }

        blacklistCounter[sku].count++;
        blacklistCounter[sku].files.push(csvFile.name);
        totalBlacklistedFiles++;

        var entryFolder = csvFile.parent;
        var destinoBase = ignoredFolder;
        var targetPath = Folder(destinoBase + "/" + entryFolder.name);


        if (!targetPath.exists) {
            try {
                try { csvFile.close(); } catch (_) { }

                var origem = entryFolder.parent.fsName;
                var destino = ignoredFolder.fsName;

                // Enxerto de AppleScript para mover a pasta no Finder
                var as =
                    'tell application "Finder"\n' +
                    '    if exists POSIX file "' + origem + '" then\n' +
                    '        move POSIX file "' + origem + '" to POSIX file "' + destino + '"\n' +
                    '    end if\n' +
                    'end tell';

                app.doScript(as, ScriptLanguage.applescriptLanguage);

                log("PASTA MOVIDA PARA _IGNORADOS (AppleScript): " + entryFolder.name);

            } catch (e) {
                log("ERRO AO MOVER PASTA PARA _IGNORADOS (AS): " + entryFolder.fsName);
                log("DETALHE: " + e.message);
            }

        } else {
            log("AVISO: Pasta já existe em _IGNORADOS: " + entryFolder.name);
        }
    }

    // ======================================================
    // BUSCA RECURSIVA DE CSVs
    // ======================================================

    function csvCollect(pastaRaiz) {

        var resultados = [];

        function parse(pasta) {

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
                        parse(itens[j]);
                    }
                }
            }
        }

        parse(pastaRaiz);
        return resultados;
    }

    var csvFiles = csvCollect(entryFolder);
    log("CSVs encontrados: " + csvFiles.length);

    // ======================================================
    // PARSER DE CSV
    // ======================================================

    // Importante: essa função assume que o CSV possui o SKU na segunda coluna SEMPRE.
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


    // Módulo de limpeza do documento mesclado, mudanças podem ser feitas aqui
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

    // O RegEx abaixo busca por números de série de OP. Ajustar aqui caso as exportações
    // saiam com o nome de fallback (Nome do diagramador) Usar o site regex101.com para validação.
    function serialNumberGen(doc, fallback) {

        var regex = /^\d{2,}\s*-\s*\d{2,}_\d{5,}-[A-Z0-9]+$/; // <----------Alterar aqui

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
        var entryFolder = csv.parent.fsName;

        if (outputFolder[entryFolder]) {
            log("IGNORADO (CSV duplicado na pasta): " + csv.fsName);
            continue;
        }

        outputFolder[entryFolder] = true;

        log("\n--- Processando CSV");

        var sku = targetSKU(csv);
        if (!sku) {
            errorCount++;
            log("ERRO: Não foi possível ler SKU em " + csv.name);
            continue;
        }

        log("SKU identificado: " + sku);

        // ==========================================
        // LISTA NEGRA DE SKUs (PROCESSO MANUAL)
        // ==========================================

        if (isIgnoredSKU(sku)) {

            log("SKU EM LISTA MANUAL — IGNORADO: " + sku);

            blacklistFile(
                sku,
                csv,
                "SKU marcado para processamento manual"
            );

            continue;
        }

        var template = File(rootFolder + "/" + sku + ".indt");
        if (!template.exists) {

            if (!missingTemplateCounter[sku]) {
                missingTemplateCounter[sku] = [];
            }
            missingTemplateCounter[sku].push(csv.name);

            log("TEMPLATE FALTANDO: " + sku, "→ " + template.fsName);
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

        var exportName = serialNumberGen(mergedDocument, sku);
        var pasta = csv.parent;

        try {
            var inddFile = File(pasta + "/" + exportName + ".indd");
            var pdfFile = File(pasta + "/" + exportName + ".pdf");

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

    for (var s in blacklistCounter) {

        log("SKU: " + s);
        log(" Ocorrências: " + blacklistCounter[s].count);

        for (var j = 0; j < blacklistCounter[s].files.length; j++) {
            log("   - " + blacklistCounter[s].files[j]);
        }

        log("");
    }

    var msg =
        "Batch finalizado.\n\n" +
        "Processados: " + processedFiles + "\n" +
        "Erros: " + errorCount + "\n" +
        "Ignorados (movidos): " + totalBlacklistedFiles;


    var missingTemplates = false;
    for (var k in missingTemplateCounter) { missingTemplates = true; break; }

    if (missingTemplates) {
        for (var k in missingTemplateCounter) {
            msg += "\n- " + k;
        }
    }

    log("\n========================================");
    log("FIM DO LOTE: " + new Date());
    log("Processados: " + processedFiles);
    log("Erros: " + errorCount);
    log("Ignorados (movidos): " + totalBlacklistedFiles);
    if (missingTemplates) {
        log("Templates faltando (SKUs):");
        for (var k in missingTemplateCounter) {
            log(" - " + k)
        }
    } else {
        log("Sem templates faltando.");
    }

    log("========================================\n");

    alert(msg);

})();

// ฅ^•ﻌ•^ฅ - Fim do Script - ฅ^•ﻌ•^ฅ
