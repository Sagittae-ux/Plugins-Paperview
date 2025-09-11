(function () {
    var doc;
    try {
        doc = app.activeDocument;
    } catch (e) {
        alert("Nenhum documento aberto.");
        return;
    }

    // --- Nome fixo configurável ---
    var nomeUsuario = "Alyssa"; // <<<----------------------- Insira seu nome aqui.
    var nomeBase = "documento"; // fallback inicial

    // --- 1ª parte: procurar "diagramado_por_NOME" ---
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

    // --- 2ª parte: procurar regex do nome final ---
    var regexNomeArquivo = /\b[A-Z]{2,9}\b/;
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

    // Se não achar o regex, usa o fallback do nome fixo ou "documento"
    var nomeFinal = nomeEncontrado ? nomeEncontrado : nomeBase;

    // Normaliza caracteres ilegais
    nomeFinal = nomeFinal.replace(/[\\\/:*?"<>|]/g, "_");

    var nomeArquivo = nomeFinal + ".indd";

    // Pergunta a pasta de destino
    var pastaDestino = Folder.selectDialog("Escolha a pasta para salvar o arquivo");
    if (!pastaDestino) return;

    var caminhoFinal = File(pastaDestino + "/" + nomeArquivo);

    try {
        if (caminhoFinal.exists) {
            if (!confirm("O arquivo " + nomeArquivo + " já existe. Deseja sobrescrever?")) {
                return;
            }
        }
        doc.save(caminhoFinal);
        alert("Documento salvo como: " + caminhoFinal.fsName);
    } catch (e) {
        alert("Erro ao salvar o documento: " + e.message);
    }
})();
