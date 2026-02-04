// ignisCalor.jsx
// Batch CSV → SKU → Template → Data Merge → Limpeza → Exportação   
// Versão 4.0
// Dev: Alyssa Ferreiro @Sagittae-UX

// Esse script automatiza o processo de diagramação em lote no Adobe InDesign,
// atuando como o novo motor central para a tarefa de mesclagem preliminar de
// processamento de pedidos de forma compreensiva para facilitar o preflight.


// RECURSOS PRINCIPAIS:
// - Busca recursiva de arquivos CSV em subpastas.
// - Identificação automática de SKU a partir de parsing do conteúdo do CSV.
// - Associação dinâmica de templates baseados no SKU detectado na pasta raiz.
// - Mesclagem de dados utilizando o recurso Data Merge do InDesign.
// - Limpeza automática do documento mesclado (remoção de marcadores, quadros vazios, etc.).
// - Exportação em lote para formatos INDD e PDF com presets definidos.
// - Processamento dinâmico de acordo com condicionais presentes no pedido.
// - Separação de pedidos sem template e marcados como manuais para as pastas relevantes. 
// - Registro detalhado de processos, erros, pedidos manuais e templates faltando em um arquivo de log.


// INSTRUÇÕES DE USO:
// 1. Configure as variáveis no início do script:
//    - exportPreset: Nome do preset de exportação PDF padronizado para produção, escrito de forma exata.
//    - entryFolder: Caminho da pasta de produção contendo os pedidos baixados do Magento.
//    - rootFolder: Caminho da pasta de templates contendo os arquivos .indt.
//    - userID: Nome do diagramador para registro no documento exportado.

// 2. Adicione SKUs problemáticos ou determinados como impossíveis de automatizar quando necessário. 
//    Os arquivos serão catalogados e movidos para a pasta "_MANUAL", onde deverão ser processados manualmente.

// 3. Adicione o script à pasta de scripts do Adobe InDesign através da janela de utilitários de scripts.
//    Recomenda-se criar uma subpasta específica para scripts personalizados, e um atalho para fácil execução.

// 4. Abrir o Adobe InDesign, na janela de utilitários de scripts, executar o script. Recomenda-se
//    a criação de um atalho para facilitar o uso (Editar > Atalhos do Teclado > Área do produto: Scripts).

// 5. Ao final do processamento, um relatório será gerado na pasta de entrada,
//    detalhando o número de arquivos processados, erros encontrados e SKUs movidos.

// 6. Faça o preflight dos arquivos, atentando-se a erros estéticos.

// 7. Caso faltem templates para determinados SKUs, esses serão listados no relatório final. Ao fim do
//    processamento, atentar-se a essa pasta e processar os arquivos restantes com o script após a
//    criação das bases. O mesmo processo se aplica a SKU's marcados para processamento manual.

// 8. Ajustar e incluir bases conforme necessário, adicionando o SKU entre aspas e respeitando a sintaxe abaixo:

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

    var entryFolder = Folder("~/Documents/PRODUCAO"); //Colar o caminho de arquivo da pasta de entrada aqui ou determinar caminho padrão
    var rootFolder = Folder("~/Documents/TEMPLATES"); //Colar o caminho de arquivo da pasta de templates aqui ou determinar caminho padrão

    var userID = "Alyssa"; //Nome do diagramador para registro no documento exportado

    var csvTarget = /\.csv$/i;

    // ======================================================
    // LISTA NEGRA - SKUs PARA PROCESSAMENTO MANUAL AQUI
    // ======================================================

    // Referir-se ao passo 8 das instruções acima para sintaxe correta.
    // Lista reservada para chancelas, itens complexos demais para diagramação automática 
    // ou arquivos antigos não adequados para o script.

    var ignoredSKUs = {
        "MD890": true,
        "MD664": true
    };

    // ======================================================
    // LOG DE PROCESSO
    // ======================================================

    var logFile = File(entryFolder.fsName + "/!relatório.txt");

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

    var missingTemplateCounter = {}; // Indicados no fim do log com o código de SKU
    var outputFolder = {}; // Número de pedidos processados com sucesso
    var blacklistCounter = {}; // Contador de SKUs listados para processamento manual

    // ======================================================
    // CRIAÇÃO DA PASTA DE IGNORADOS / TEMPLATES AUSENTES
    // ======================================================   

    var ignoredFolder = Folder(entryFolder + "/_MANUAL");
    if (!ignoredFolder.exists) {
        ignoredFolder.create();
    }

    var missingBaseFolder = Folder(entryFolder + "/_SEM_BASE");
    if (!missingBaseFolder.exists) {
        missingBaseFolder.create();
    }

    // ======================================================
    // VALIDAÇÕES INICIAIS
    // ======================================================

    // Check para pastas de entrada e templates. Tomar cuidado ao alterar os caminhos
    // pois o script será abortado caso estejam incorretos.
    if (!entryFolder.exists || !rootFolder.exists) {
        alert("Erro: O caminho da pasta de produção ou templates não foi encontrado.\nVerifique as configurações no início do script e cole o caminho de arquivo na line 'var entryFolder' e 'var rootFolder'.");
        log("Erro: Caminho de pasta inválido.");
        return;
    }

    function isIgnoredSKU(sku) {
        return ignoredSKUs[sku] === true;
    }

    // ======================================================
    // BUSCA RECURSIVA DE CSVs
    // ======================================================

    function csvCollect(rootDirectory) {

        var results = [];

        function parse(pasta) {

            var items = pasta.getFiles();
            var csvCheck = false;

            for (var i = 0; i < items.length; i++) {
                if (items[i] instanceof File && csvTarget.test(items[i].name)) {
                    results.push(items[i]);
                    csvCheck = true;
                    break;
                }
            }

            if (!csvCheck) {
                for (var j = 0; j < items.length; j++) {
                    if (items[j] instanceof Folder) {
                        parse(items[j]);
                    }
                }
            }
        }

        parse(rootDirectory);
        return results;
    }

    var orderState = {};
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

        var cols = parseCSVLine(csvCell[1]); // IMPORTANTE: Caso a coluna do CSV mude, trocar aqui, atentando-se ao fato de que arrays começam do 0
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

        for (var s = 0; s < doc.stories.length; s++) {

            var story = doc.stories[s];

            var containers = story.textContainers;
            var skipStory = false;

            for (var c = 0; c < containers.length; c++) {
                try {
                    if (
                        containers[c].locked === true ||
                        (containers[c].itemLayer && containers[c].itemLayer.locked === true)
                    ) {
                        skipStory = true;
                        break;
                    }
                } catch (_) { }
            }

            if (skipStory) {
                continue;
            }

            try {
                var contents = story.contents;
                var idx;

                while ((idx = contents.indexOf("\\#")) !== -1) {

                    // Remove o \#
                    story.characters[idx].remove();
                    story.characters[idx].remove();

                    // Força um backspace lógico: remove 1 caractere antes, se existir
                    if (idx - 1 >= 0) {
                        try {
                            story.characters[idx - 1].remove();
                        } catch (_) { }
                    }

                    contents = story.contents;
                }

            } catch (_) { }
        }

        app.findGrepPreferences.findWhat = "\\\\n";
        app.changeGrepPreferences.changeTo = "\\n";
        doc.changeGrep();

        app.findGrepPreferences = NothingEnum.nothing;
        app.changeGrepPreferences = NothingEnum.nothing;

        // Limpeza de frames vazios, contextual para texto e imagens
        var allFrames = doc.allPageItems;
        for (var k = allFrames.length - 1; k >= 0; k--) {
            var item = allFrames[k];

            try {
                // Remover frames de texto vazios
                if (item instanceof TextFrame && !item.locked) {
                    if (item.contents.replace(/\s+/g, "") === "") {
                        item.remove();
                        continue;
                    }
                }

                // Remover frames sem imagem
                if (item instanceof Rectangle && !item.locked) {
                    if (item.graphics.length === 0 && item.allGraphics.length === 0) {
                        item.remove();
                    }
                }
            } catch (e) {
            }
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
                var line = csvCell[l].replace(/^\s+|\s+$/g, "");
                if (regex.test(line)) {
                    return line.replace(/[\\\/:*?"<>|]/g, "_");
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

        var orderFolder = csv.parent.parent;
        var orderKey = orderFolder.fsName;

        if (!orderState[orderKey]) {
            orderState[orderKey] = {
                folder: orderFolder,
                hasBlacklist: false,
                hasMissingTemplate: false,
                hasExistingINDD: false
            };
        }


        var csv = csvFiles[i];
        // Evita reprocessar pedidos já diagramados
        var existingINDDs = csv.parent.getFiles(function (f) {
            return f instanceof File && /\.indd$/i.test(f.name);
        });

        var skipMerge = false;

        if (existingINDDs.length > 0) {
            log("INDD já existente na pasta: " + csv.parent.name);
            orderState[orderKey].hasExistingINDD = true;
            skipMerge = true; // apenas impede a diagramação
        }


        var orderPath = csv.parent.fsName;

        if (outputFolder[orderPath]) {
            log("IGNORADO (CSV duplicado na pasta): " + csv.fsName);
            continue;
        }
        outputFolder[orderPath] = true;

        log("\n--- Processando CSV");

        var sku = targetSKU(csv);
        if (!sku) {
            errorCount++;
            log("ERRO: Não foi possível ler SKU em " + csv.name);
            continue;
        }

        log("SKU identificado: " + sku);

        // ===== CLASSIFICAÇÃO ANTECIPADA (sempre ocorre) =====
        if (isIgnoredSKU(sku)) {
            orderState[orderKey].hasBlacklist = true;

            if (!blacklistCounter[sku]) {
                blacklistCounter[sku] = { count: 0 };
            }
            blacklistCounter[sku].count++;
            totalBlacklistedFiles++;

            log("SKU em blacklist identificado.");
        }


        var template = File(rootFolder + "/" + sku + ".indt");
        if (!template.exists) {

            if (!missingTemplateCounter[sku]) {
                missingTemplateCounter[sku] = [];
            }
            missingTemplateCounter[sku].push(csv.name);

            log("TEMPLATE FALTANDO: " + sku);
            orderState[orderKey].hasMissingTemplate = true;

            // não mescla, mas já classificou
            continue;
        }

        // Já existe INDD → não mesclar novamente
        if (skipMerge) {
            log("Pedido já diagramado → pulando mescla.");
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

        if (isIgnoredSKU(sku)) {

            log("SKU em blacklist → fluxo manual com INDD salvo.");

            // Nome primeiro
            var exportName = serialNumberGen(mergedDocument, sku);

            // SALVA IMEDIATAMENTE APÓS A MESCLA (sem limpeza)
            try {
                var inddFile = File(csv.parent + "/" + exportName + ".indd");
                mergedDocument.save(inddFile);
                log("INDD SALVO (manual, sem limpeza): " + inddFile.fsName);
            } catch (e) {
                log("ERRO ao salvar INDD manual: " + e.message);
            }

            // AGORA SIM aplica identificação e limpeza
            userIdentifier(mergedDocument, userID);
            fileCleanup(mergedDocument);

            // ===== ORDEM CORRETA DO DATAMERGE =====
            try {
                docBase.dataMergeProperties.removeDataSource();
            } catch (_) { }

            try {
                docBase.close(SaveOptions.NO);
            } catch (_) { }

            try {
                mergedDocument.close(SaveOptions.NO);
            } catch (_) { }

            $.sleep(800);

            log("Pedido marcado para _MANUAL ao final.");

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

            try {
                var as =
                    'tell application "Finder"\n' +
                    '    if (count of windows) > 0 then\n' +
                    '        close front window\n' +
                    '    end if\n' +
                    'end tell';

                app.doScript(as, ScriptLanguage.applescriptLanguage);
            } catch (e) { }

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

            } catch (e) { }

        } catch (e) {
            alert("Erro ao salvar/exportar o documento: " + e.message);
        }

        try {
            docBase.dataMergeProperties.removeDataSource();
        } catch (_) { }

        try {
            docBase.close(SaveOptions.NO);
        } catch (_) { }

        try {
            mergedDocument.close(SaveOptions.NO);
        } catch (_) { }

        csv = null;
        $.sleep(800);

    }

    // ======================================================
    // MOVIMENTAÇÃO FINAL DAS PASTAS (APÓS TODO O PROCESSAMENTO)
    // ======================================================

    for (var key in orderState) {

        var state = orderState[key];

        // Se não tem motivo para mover, não move
        if (!state.hasBlacklist && !state.hasMissingTemplate) {
            continue;
        }

        var destino = null;

        if (state.hasMissingTemplate) {
            destino = missingBaseFolder;
            log("Movendo pedido para _SEM_BASE: " + state.folder.name);

        } else if (state.hasBlacklist) {
            destino = ignoredFolder;
            log("Movendo pedido para _MANUAL: " + state.folder.name);
        }

        if (destino) {

            try {
                // tentativa nativa
                state.folder.move(destino);
                log("Movido via ExtendScript: " + state.folder.name);

            } catch (_) {

                // fallback obrigatório via AppleScript (Finder)
                try {
                    var as =
                        'tell application "Finder"\n' +
                        'move folder (POSIX file "' + state.folder.fsName + '") to folder (POSIX file "' + destino.fsName + '")\n' +
                        'end tell';

                    app.doScript(as, ScriptLanguage.applescriptLanguage);
                    log("Movido via AppleScript: " + state.folder.name);

                } catch (e) {
                    log("ERRO AO MOVER: " + state.folder.name + " → " + e.message);
                }
            }
        }
    }

    // ======================================================
    // ALERT FINAL + FECHAMENTO DO LOG
    // ======================================================

    for (var s in blacklistCounter) {

        log("SKU: " + s);
        log(" Ocorrências: " + blacklistCounter[s].count);
    }

    var msg =
        "Batch finalizado.\n\n" +
        "Processados: " + processedFiles + "\n" +
        "Erros: " + errorCount + "\n" +
        "Arquivos manuais movidos: " + totalBlacklistedFiles + "\n" +
        "Templates faltando:";

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
    log("Arquivos manuais movidos: " + totalBlacklistedFiles);
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