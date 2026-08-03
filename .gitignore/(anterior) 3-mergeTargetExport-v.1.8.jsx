// mergeTargetExport.jsx
// Script para InDesign que realiza a substituição de um marcador "diagramado_por_NOME" pelo nome do usuário, procura um padrão específico de nome no texto para nomear o arquivo salvo, e exporta o documento como .indd e .pdf usando uma predefinição específica.
// Versão 1.8
// Dev: Alyssa Ferreiro / @Sagittae-UX

// ATENÇÃO: Instruções ao usuário
// O usuário deve alterar apenas a variável "nomeUsuario" abaixo para o nome desejado, e possuir uma predefinição de exportação de .pdf de nome "Diagramacao2025".

(function () {
    var doc;
    try {
        doc = app.activeDocument;
    } catch (e) {
        alert("Nenhum documento aberto.");
        return;
    }

    // Área de alteração de nome do diagramador, faça a alteração apenas aqui:

    var nomeUsuario = "Nick"; // <<<---------------------------------------- altere aqui para mudar o nome
    var nomeBase = "documento"; // fallback inicial

    // --- Passo 1: procurar "diagramado_por_NOME" e executar a assinatura do usuário

    var marcadorRegex = /\bdiagramado_por_NOME\b/;
    var marcadorEncontrado = false;

    for (var i = 0; i < doc.stories.length; i++) {
        var story = doc.stories[i];
        if (marcadorRegex.test(story.contents)) {
            marcadorEncontrado = true;

            // Substitui no texto do documento pelo nome fixo
            story.contents = story.contents.replace(marcadorRegex, nomeUsuario);

            // define como base para nome do arquivo caso regex não seja achado
            nomeBase = nomeUsuario;
            break;
        }
    }

    // --- Passo 2: procurar RegEx do nome final

    var regexNomeArquivo = /^\d+\s*-\s*\d+_\d{5,}-[A-Z]{2,}\d*$/;
    var nomeEncontrado = null;

    for (var j = 0; j < doc.stories.length; j++) {
        var conteudo = doc.stories[j].contents;
        var linhas = conteudo.split(/[\r\n]/);
        for (var k = 0; k < linhas.length; k++) {
            var linha = linhas[k].replace(/^\s+|\s+$/g, ""); // trim
            if (regexNomeArquivo.test(linha)) {
                nomeEncontrado = linha;
                break;
            }
        }
        if (nomeEncontrado) break;
    }

    // Se não achar o RegEx, usa o fallback do nome fixo ou "documento"
    var nomeFinal = nomeEncontrado ? nomeEncontrado : nomeBase;



    try {
        // Voltar para a primeira página do documento
        app.activeWindow.activePage = doc.pages[0];

        // Ajustar zoom para caber na janela
        if (app.activeWindow.layoutWindows.length > 0) {
            app.activeWindow.layoutWindows[0].zoom(ZoomOptions.FIT_SPREAD);
        }
    } catch (e) {
        // Ignora erros de zoom/visualização
    }


    // Normaliza caracteres ilegais
    nomeFinal = nomeFinal.replace(/[\\\/:*?"<>|]/g, "_");

    // Monta nomes de arquivo
    var nomeArquivoINDD = nomeFinal + ".indd";
    var nomeArquivoPDF = nomeFinal + ".pdf";

    // Pergunta a pasta de destino
    var pastaDestino = Folder.selectDialog("Escolha a pasta para salvar o arquivo");
    if (!pastaDestino) return;

    var caminhoINDD = File(pastaDestino + "/" + nomeArquivoINDD);
    var caminhoPDF = File(pastaDestino + "/" + nomeArquivoPDF);

    try {
        // --- Exporta INDD ---
        if (caminhoINDD.exists) {
            if (!confirm("O arquivo " + nomeArquivoINDD + " já existe. Deseja sobrescrever?")) {
                return;
            }
        }
        doc.save(caminhoINDD);

        // --- Exporta PDF usando preset ---
        if (caminhoPDF.exists) {
            if (!confirm("O arquivo " + nomeArquivoPDF + " já existe. Deseja sobrescrever?")) {
                return;
            }
        }

        var presetName = "Diagramacao2025"; // nome exato do preset no InDesign 
        var preset = app.pdfExportPresets.itemByName(presetName);

        if (!preset.isValid) {
            alert("O preset de exportação PDF '" + presetName + "' não foi encontrado.");
            return;
        }

        doc.exportFile(ExportFormat.pdfType, caminhoPDF, false, preset);

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

    // Fechar sem salvar
    doc.close(SaveOptions.NO);

})();
