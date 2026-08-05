// ignisCalor.jsx
// Batch CSV → SKU → Template → Data Merge → Limpeza → Exportação → Distribuição   
// Versão 4.1
// Dev: Alyssa Ferreiro @Sagittae-UX

// Esse script foi produzido baseado no sistema provisório de ferramentas de 
// diagramação, consolidando processos de busca de produto, mescla, padronização, revisões corretivas e exportação em um
// motor de processamento específico para a empresa.
// O script possui blocos configuráveis no início que devem ser alterados de acordo com o usuário e à medida que arquivos forem 
// emendados para a lista de processamento manual.
// Em caso de dúvidas ou necessidade de patches, entrar em contato com a dev.

// v4.1
// Recursos adicionais em teste nesse patch:
// - Regularizador de nomes próprios para prevenção do erro de fontes cursivas
//   tendo que ser revisadas manualmente para troca de caixa de título
// - Medidas preventivas contra espaços duplos e espaços antes do começo de frases

//         ♡  ╱|、
//           (˚ˎ 。7
//           |、˜〵       
//           じしˍ,)ノ


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
// - Detecção de caracteres Unicode não capturados pela fonte de fallback implantada para encaminhamento manual. 
// - Fallback para evitar duplicatas de arquivos já processados, tornando o plugin seguro para múltiplos passes.


// INSTRUÇÕES DE USO:

// 1. Adicionar o script a pasta de scripts do Adobe InDesign através de Janela > Utilitários > Scripts.
//    No menu lateral, clique "Revelar no Finder" para encontrar o destino de salvamento correto.

// 2. Configurar as variáveis no início do script:
//    - exportPreset: Nome do preset de exportação PDF padronizado para produção, escrito de forma exata.
//    - entryFolder: Caminho da pasta de produção contendo os pedidos baixados do Magento.
//    - rootFolder: Caminho da pasta de templates contendo os arquivos .indt.
//    - userID: Nome do diagramador para registro no documento exportado.

// 3. Adicionar SKUs problemáticos ou determinados como impossíveis de automatizar sempre que necessário. 
//    Os arquivos serão mesclados, salvos, catalogados e movidos para a pasta "_MANUAL". Pedidos que não
//    possuem template válido na pasta raiz serão enviados para a pasta "_SEM_BASE".
//    Ao fim do processo, conferir as pastas e processar os arquivos restantes.

// 4. Baixar e extrair o lote a ser processado na pasta de input, de nome PRODUCAO.

// 5. Abrir o Adobe InDesign, na janela de utilitários, executar o script. Recomenda-se
//    a criação de um atalho para facilitar o uso (Editar > Atalhos do Teclado > Área do produto: Scripts).

// 6. Ao final do processamento, um relatório será gerado na caixa de diálogo, detalhando o tamanho do lote, erros e
//    templates faltando por SKU. Um relatório detalhado do processamento de cada item individualmente, detalhando cada
//    erro será criado no mesmo local da pasta contendo o lote.

// 7. Caso existam templates faltando, checar os produtos, adicionar à pasta de templates e executar o script novamente. 
//    Diagramar os arquivos marcados para processamento manual utilizando as ferramentas anteriores de revisão e exportação.

// 8. Após o lote estar completamente diagramado, faça o preflight dos arquivos, atentando-se a erros estéticos. 

// 9. Unir os pedidos e encaminhar os arquivos para a pasta de fechamento no dia relevante.

// 10. Tome um segundo do tempo economizado para apreciar que isso já demorou um minuto e meio por item
//     para as pessoas mais experientes e depois nunca mais pense nisso porque agora apreciamos a beleza e perfeição
//     da era das máquinas.


(function ignisCalor() {

    // ======================================================
    // CONFIGURAÇÕES - MUDANÇAS AQUI
    // ======================================================

    var exportPreset = "Diagramacao2025"; //Preset exata padronizada para produção

    var entryFolder = Folder("~/Documents/PRODUCAO"); //Colar o caminho de arquivo da pasta de entrada aqui

    var rootFolder = Folder("~/Documents/TEMPLATES"); //Colar o caminho de arquivo da pasta de templates aqui

    var userID = "Alyssa"; //Nome do diagramador para registro no documento exportado

    // ======================================================
    // LISTA NEGRA - SKUs PARA PROCESSAMENTO MANUAL AQUI
    // ======================================================

    // Lista reservada para chancelas, itens complexos demais para diagramação automática 
    // ou arquivos antigos não adequados para o script.
    // Ajustar e incluir bases conforme necessário, adicionando o SKU exato como
    // aparece no site, entre aspas e separando por vírgula, como no exemplo abaixo:

    // var ignoredSKUs {
    //     "MD####": true,
    //     "CA0000": true
    // };

    var ignoredSKUs = {

    };

    // ======================================================
    // LOG DE PROCESSO
    // ======================================================

    // Configuração do relatório de processo

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
    log("========================================\n");

    // ======================================================
    // CONTADORES DE ERRO / CONTROLE
    // ======================================================

    // Contadores zerados
    var processedFiles = 0;
    var errorCount = 0;
    var totalBlacklistedFiles = 0;

    var missingTemplateCounter = {}; // Indicados no fim do log com o código de SKU
    var outputFolder = {}; // Número de pedidos processados com sucesso
    var blacklistCounter = {}; // Contador de SKUs listados para processamento manual

    // ======================================================
    // CRIAÇÃO DA PASTA DE IGNORADOS / TEMPLATES AUSENTES
    // ======================================================   

    // Pasta de arquivos manuais
    var ignoredFolder = Folder(entryFolder + "/_MANUAL");
    if (!ignoredFolder.exists) {
        ignoredFolder.create();
    }
    // Captura de pedidos sem base ou erros de SKU
    var missingBaseFolder = Folder(entryFolder + "/_SEM_BASE");
    if (!missingBaseFolder.exists) {
        missingBaseFolder.create();
    }

    // ======================================================
    // FUNÇÃO EXPERIMENTAL: REGULARIZAÇÃO DE NOME
    // ======================================================

    // A função abaixo tem como objetivo regularizar a entrada de nomes próprios no sistema,
    // visando evitar o maior erro conhecido da diagramação com a entrada de nomes próprios em
    // caixa alta em pedidos onde a fonte é case sensitive. Manter essa função comentada exceto em ambientes de teste,
    // pendente aprovação da gerência para o uso no fluxo normal de trabalho

    function titleCase(str) {

        if (!str) return str;

        // normaliza espaços e força caixa baixa
        str = str.replace(/^\s+|\s+$/g, "").toLowerCase();

        return str;
    }

    // Parser CSV completo: retorna array de registros
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
                    var next = content[i + 1];
                    if (next === '"') {
                        field += '"';
                        i += 2;
                        continue;
                    } else {
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

    // Criação de novo .csv
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

    // Busca de colunas de nome para regularização da caixa de texto
    function processCSV(file) {
        if (!file.exists) return;
        if (!file.open("r")) return;
        var content = file.read();
        file.close();

        var records = parseCSV(content);
        if (!records || records.length < 2) return;

        // Caso a coluna de nome se manifeste de alguma outra maneira, inclua aqui 
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

        // Cria backup do arquivo .csv processado caso seja necessário usar um .csv limpo ou consultas eventuais
        var backup = new File(file.fullName + ".bak");
        try {

            if (backup.exists) backup.remove();
            file.copy(backup);
        } catch (e) {

        }

        if (!file.open("w")) return;
        file.encoding = "UTF-16"; // Importante manter UTF-16, encoding usado no InDesign
        file.lineFeed = "Windows";
        file.write(buildCSV(records));
        file.close();
    }

    // Função evita que arquivos sejam processados novamente quando o script é executado de novo
    function scanFolder(folder) {

        if (!folder.exists) return;

        var items = folder.getFiles();
        var csvProcessed = false;

        for (var i = 0; i < items.length; i++) {

            var item = items[i];

            if (item instanceof Folder) {
                scanFolder(item);

            } else if (!csvProcessed && item instanceof File && item.name.match(/\.csv$/i)) {

                processCSV(item);
                csvProcessed = true;
            }
        }
    }

    scanFolder(entryFolder);


    // ======================================================
    // VALIDAÇÕES INICIAIS
    // ======================================================

    // Check para pastas de entrada e templates. Tomar cuidado ao alterar os caminhos,
    // o script será abortado caso estejam incorretos.
    if (!entryFolder.exists || !rootFolder.exists) {
        alert("Erro: O caminho da pasta de produção ou templates não foi encontrado.\nVerifique as configurações no início do script e cole o caminho de arquivo na linha 'var entryFolder' e 'var rootFolder'.");
        return;
    }

    // Puxa a lista de SKU's ignorados para comparação com a lista do lote.
    function isIgnoredSKU(sku) {
        return ignoredSKUs[sku] === true;
    }

    // ======================================================
    // BUSCA RECURSIVA DE CSVs
    // ======================================================

    var csvTarget = /\.csv$/i;

    function csvCollect(rootDirectory) {

        var results = []; // Array contendo a quantidade de CSVs individuais

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

    function findTemplate(sku) {
        if (!sku) return null;

        var directTemplate = File(rootFolder.fsName + "/" + sku + ".indt");
        if (directTemplate.exists) {
            return directTemplate;
        }

        function search(folder) {
            var items = folder.getFiles();
            for (var i = 0; i < items.length; i++) {
                var item = items[i];

                if (item instanceof File && item.name.toLowerCase() === sku.toLowerCase() + ".indt") {
                    return item;
                }

                if (item instanceof Folder) {
                    var found = search(item);
                    if (found) {
                        return found;
                    }
                }
            }
            return null;
        }

        return search(rootFolder);
    }

    // ======================================================
    // PARSER DE CSV
    // ======================================================

    // Busca de SKU através da posição no .csv
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

        // Parser encontra a segunda coluna do .csv, que possui o SKU do pedido
        var cols = parseCSVLine(csvCell[1]);
        if (cols.length < 2) return null;

        return cols[1].replace(/^\s+|\s+$/g, "");
    }

    // ======================================================
    // MESCLAGEM
    // ======================================================


    function validateCSVLinks(csvFile) {
        if (!csvFile || !csvFile.exists) return true;

        var folder = csvFile.parent;
        var lines = [];

        try {
            csvFile.encoding = "UTF-16";
            if (!csvFile.open("r")) return true;
            var txt = csvFile.read();
            csvFile.close();
            lines = txt.split(/\r\n|\n|\r/);
        } catch (e) {
            log("ERRO: Não foi possível ler o CSV para validar links: " + csvFile.name);
            return false;
        }

        if (lines.length < 2) return true;

        var header = parseCSVLine(lines[0]);
        var imgColIndex = -1;
        for (var i = 0; i < header.length; i++) {
            if (header[i].indexOf('@IMG') === 0) {
                imgColIndex = i;
                break;
            }
        }

        if (imgColIndex < 0) return true;

        for (var r = 1; r < lines.length; r++) {
            if (!lines[r]) continue;

            var cells = parseCSVLine(lines[r]);
            if (imgColIndex >= cells.length) continue;

            var rawPath = cells[imgColIndex];
            if (!rawPath) continue;

            if (rawPath.indexOf("Desejo") !== -1 || rawPath.indexOf("desejo") !== -1) {
                continue;
            }

            var imgFile = new File(rawPath);
            if (!imgFile.exists) {
                imgFile = new File(folder.absoluteURI + "/" + rawPath);
            }

            if (!imgFile.exists) {
                log("ERRO: Link MISSING - " + rawPath + " (CSV: " + csvFile.name + ")");
                return false;
            }
        }

        return true;
    }


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
    // FUNÇÃO EXPERIMENTAL: FALLBACK PARA UNICODE
    // ======================================================

    // Função experimental para tentar determinar alcances de Unicode
    // Não foi deteminado como confiável para a eliminação de Unicodes problemáticos,
    // mas o impacto no processamento que esse trecho de código foi determinado como vestigial
    // para encorajar sua remoção.

    function unicodeFallback(doc) {

        function badUnicode(code) {

            return (
                (code >= 0x2600 && code <= 0x26FF) ||   // Misc Symbols
                (code >= 0x2700 && code <= 0x27BF) ||   // Dingbats
                (code >= 0x1F300 && code <= 0x1F5FF) || // Misc Pictographs
                (code >= 0x1F600 && code <= 0x1F64F) || // Emoticons
                (code >= 0x1F680 && code <= 0x1F6FF) || // Transport
                (code >= 0x1F900 && code <= 0x1F9FF) || // Supplemental
                (code >= 0x1FA00 && code <= 0x1FAFF)    // Extended Symbols
            );
        }

        for (var s = 0; s < doc.stories.length; s++) {

            var contents = doc.stories[s].contents;

            for (var i = 0; i < contents.length; i++) {

                var code = contents.charCodeAt(i);

                // Suporte a surrogate pairs (emoji acima de U+FFFF)
                if (0xD800 <= code && code <= 0xDBFF && i + 1 < contents.length) {
                    var next = contents.charCodeAt(i + 1);
                    if (0xDC00 <= next && next <= 0xDFFF) {
                        code = ((code - 0xD800) * 0x400) + (next - 0xDC00) + 0x10000;
                        i++;
                    }
                }

                if (badUnicode(code)) {
                    return true;
                }
            }
        }

        return false;
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

            // Busca da chave para célula vazia, determinada em reunião como *Não Desejo
            try {
                var contents = story.contents;
                var idx;
                var token = "*Não Desejo";

                while ((idx = contents.indexOf(token)) !== -1) {

                    // Remove a chave de texto
                    for (var i = idx + token.length - 1; i >= idx; i--) {
                        try {
                            story.characters[i].remove();
                        } catch (_) { }
                    }

                    // Backspace após a limpeza para deletar a linha
                    if (idx - 1 >= 0) {
                        try {
                            story.characters[idx - 1].remove();
                        } catch (_) { }
                    }

                    contents = story.contents;
                }

            } catch (_) { }
        }



        // Espaços em branco sem texto ou quebra de linha = "^\ \h*?$"

        // GREP de substituição
        // Um ou mais espaços em branco que precedem uma linha e não possuem nada antes = ^\ {1,}\b
        // Um ou mais espaços em branco após uma linha que não possuem nada depois = \b\ {1,}$
        // Espaços duplos no meio de palavras = \b\s{2,}\b

        app.findGrepPreferences.findWhat = "\\\\n";
        app.changeGrepPreferences.changeTo = "\\n";
        doc.changeGrep();

        app.findGrepPreferences.findWhat = " {2,}";
        app.changeGrepPreferences.changeTo = " ";
        doc.changeGrep();

        app.findGrepPreferences.findWhat = "^\\s+";
        app.changeGrepPreferences.changeTo = "";
        doc.changeGrep();

        app.findGrepPreferences = NothingEnum.nothing;
        app.changeGrepPreferences = NothingEnum.nothing;

        // Limpeza de frames vazios, contextual para texto e imagens
        var allFrames = doc.allPageItems;
        for (var k = allFrames.length - 1; k >= 0; k--) {
            var item = allFrames[k];

            try {
                if (item instanceof TextFrame && !item.locked) {
                    if (item.contents.replace(/\s+/g, "") === "") {
                        item.remove();
                        continue;
                    }
                }

                // Remover frames sem imagem. Atentar-se a bases com quadros vazios não travados no InDesign
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

        var regex = /^\d{2,}\s*-\s*\d{2,}_\d{5,}-[A-Z0-9]+$/; // Alterar aqui

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
            skipMerge = true;
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

        if (isIgnoredSKU(sku) || orderState[orderKey].hasBlacklist) {
            orderState[orderKey].hasBlacklist = true;

            if (!blacklistCounter[sku]) {
                blacklistCounter[sku] = { count: 0 };
            }
            blacklistCounter[sku].count++;
            totalBlacklistedFiles++;

            log("SKU em blacklist identificado.");
        }

        var template = findTemplate(sku);
        if (!template || !template.exists) {

            if (!missingTemplateCounter[sku]) {
                missingTemplateCounter[sku] = [];
            }
            missingTemplateCounter[sku].push(csv.name);

            log("TEMPLATE FALTANDO: " + sku);
            orderState[orderKey].hasMissingTemplate = true;

            continue;
        }

        if (skipMerge) {
            log("Pedido duplicado → pulando mescla.");
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

        if (!validateCSVLinks(csv)) {
            errorCount++;
            log("ERRO: Links ausentes no CSV, pulando mesclagem: " + csv.name);
            try { docBase.dataMergeProperties.removeDataSource(); } catch (_) { }
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

        if (unicodeFallback(mergedDocument)) {

            orderState[orderKey].hasBlacklist = true;

            if (!blacklistCounter[sku]) {
                blacklistCounter[sku] = { count: 0 };
            }
            blacklistCounter[sku].count++;
            totalBlacklistedFiles++;

            log("Fallback ativado → marcado como manual.");
        }

        if (isIgnoredSKU(sku) || orderState[orderKey].hasBlacklist) {
            log("SKU em blacklist → fluxo manual com INDD salvo.");
            var exportName = serialNumberGen(mergedDocument, sku);
            try {
                var inddFile = File(csv.parent + "/" + exportName + ".indd");
                mergedDocument.save(inddFile);
                log("Arquivo manual detectado: " + inddFile.fsName);
            } catch (e) {
                log("ERRO ao salvar INDD manual: " + e.message);
            }

            userIdentifier(mergedDocument, userID);
            fileCleanup(mergedDocument);
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
    // PÓS PROCESSAMENTO DOS PEDIDOS
    // ======================================================

    // Módulo suplementar para mover as pastas de pedido com arquivos sem template ou 
    // marcados como manuais para pastas extras para serem facilmente encontrados.

    for (var key in orderState) {

        var state = orderState[key];

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
                state.folder.move(destino);
                log("Movido: " + state.folder.name);

            } catch (_) {
                try {
                    var as =
                        'tell application "Finder"\n' +
                        'move folder (POSIX file "' + state.folder.fsName + '") to folder (POSIX file "' + destino.fsName + '")\n' +
                        'end tell';

                    app.doScript(as, ScriptLanguage.applescriptLanguage);
                    log("Movido: " + state.folder.name);

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
        "Lote finalizado ദ്ദി◝ ⩊ ◜.ᐟ\n\n" +
        "Total de pedidos: " + csvFiles.length + "\n" +
        "Processados: " + processedFiles + "\n" +
        "Erros: " + errorCount + "\n" +
        "Enviados para manual: " + totalBlacklistedFiles + "\n" +
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