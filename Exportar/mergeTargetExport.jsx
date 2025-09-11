//Esse plugin tem como objetivo renomear e salvar o arquivo ativo do InDesign em um arquivo mesclado, baseado em um nome fixo e/ou regex encontrado no texto do documento.
// 
// Instruções de uso:
// 1. Abra o InDesign e carregue o documento que deseja processar.
// 2. Insira seu nome na variável "nomeUsuario" no início do script.
// 3. Execute o script. Ele procurará por "diagramado_por_NOME" e substituirá pelo seu nome.
// 4. O script também procurará por uma linha que contenha dígitos seguidos de maiúsculas no formato "123 - 456_78901-ABCD" para usar como nome do arquivo. atenção com o dado final.
// 5. Se não encontrar, usará seu nome ou "documento" como fallback.
// 6. Escolha a pasta onde deseja salvar o arquivo renomeado.
// 7. O script irá salvar o arquivo como .indd e exportar um PDF usando o preset "Diagramação2022".


// !! O usuário deve alterar apenas a variável "nomeUsuario" abaixo para o nome desejado, e salvar outra predefinição no InDesign retirando o hífen de "Diagramação-2022".

(function () {
    var doc;
    try {
        doc = app.activeDocument;
    } catch (e) {
        alert("Nenhum documento aberto.");
        return;
    }

    // Área de alteração de nome do diagramador, faça a alteração apenas aqui:

    var nomeUsuario = "Alyssa"; // <<<---------------------------------------- altere aqui para mudar o nome
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

    var regexNomeArquivo = /^\d+\s*-\s*\d+_\d{5,}-[A-Z]{2,6}$/;
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

    // Normaliza caracteres ilegais
    nomeFinal = nomeFinal.replace(/[\\\/:*?"<>|]/g, "_");

    // Monta nomes de arquivo
    var nomeArquivoINDD = nomeFinal + ".indd";
    var nomeArquivoPDF  = nomeFinal + ".pdf";

    // Pergunta a pasta de destino
    var pastaDestino = Folder.selectDialog("Escolha a pasta para salvar o arquivo");
    if (!pastaDestino) return;

    var caminhoINDD = File(pastaDestino + "/" + nomeArquivoINDD);
    var caminhoPDF  = File(pastaDestino + "/" + nomeArquivoPDF);

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

        var presetName = "Diagramação2022"; // nome exato do preset no InDesign (Predefinição não pode ter caracteres especiais como hífen)
        var preset = app.pdfExportPresets.itemByName(presetName);

        if (!preset.isValid) {
            alert("O preset de exportação PDF '" + presetName + "' não foi encontrado.");
            return;
        }

        doc.exportFile(ExportFormat.pdfType, caminhoPDF, false, preset);

        alert("Documento salvo como:\n" + caminhoINDD.fsName + "\n\nE exportado como PDF:\n" + caminhoPDF.fsName);

    } catch (e) {
        alert("Erro ao salvar/exportar o documento: " + e.message);
    }
})();
