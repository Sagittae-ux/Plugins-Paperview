// Limpeza completa de texto e frames no documento InDesign

// O plugin executa uma limpeza completa de dados redundantes do arquivo, ignorando camadas e frames bloqueados, identificando instâncias literais de texto vazio para solucionar erros de mesclagem de dados, removendo frames de texto e imagem vazios, ajustando frames de texto com overset e limpando quebras de texto vazias.



(function Limpeza () {
    var doc = app.activeDocument;

    // Remover caracatere especial representando célula vazia na tabela de dados .csv
    app.findGrepPreferences = NothingEnum.nothing;
    app.changeGrepPreferences = NothingEnum.nothing;


// Edições automáticas adicionais devem seguir o padrão abaixo e utilizar RegEx para identificação de string alvo. Para clareza de debug e manutenção, cada edição deve ser comentada com descrição do objetivo.

// Exemplo:     

        // app.findGrepPreferences.findWhat = "<stringParaBusca>";
        // app.changeGrepPreferences.changeTo = "<stringRevisada>";
        // doc.changeGrep();

    // -------------------------------------------- Área de edição --------------------------------------------


    // O código busca instâncias literais de caractere e deleta
    app.findGrepPreferences.findWhat = "\\\\#"; //<--------- Inserir valor que representa célula vazia
    app.changeGrepPreferences.changeTo = "";
    doc.changeGrep();

    // Limpeza de espaço duplo por simples
    app.findGrepPreferences.findWhat = " {2,}";
    app.changeGrepPreferences.changeTo = " ";
    doc.changeGrep();

    // Limpeza de quebras de texto
    app.findGrepPreferences.findWhat = "\\r(?=\\r)";
    app.changeGrepPreferences.changeTo = "";
    doc.changeGrep();

    // Limpeza de espaços antes do parágrafo
    app.findGrepPreferences.findWhat = "^\\s+";
    app.changeGrepPreferences.changeTo = "";
    doc.changeGrep();

    // Limpeza de espaços depois do parágrafo
    app.findGrepPreferences.findWhat = "\\s+$";
    app.changeGrepPreferences.changeTo = "";
    doc.changeGrep();

// Substituir string literal "\n" por quebra de parágrafo real
    app.findGrepPreferences.findWhat = "\\\\n"; // regex para literal "\n"
    app.changeGrepPreferences.changeTo = "\\r"; // insere quebra de parágrafo real
    doc.changeGrep();


    // -------------------------------------------- Área de edição --------------------------------------------

    // Reset prefs
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

                // Overset
                if (item.overflows) {
                    var tries = 0;
                    while (item.overflows && tries < 20) {
                        item.geometricBounds = [
                            item.geometricBounds[0],
                            item.geometricBounds[1],
                            item.geometricBounds[2] + 100, // aumenta altura
                            item.geometricBounds[3]
                        ];
                        tries++;
                    }
                }
            }

            // Remover frames sem imagem
            if (item instanceof Rectangle && !item.locked) {
                if (item.graphics.length === 0 && item.allGraphics.length === 0) {
                    item.remove();
                }
            }
        } catch (e) {
            // ignora bloqueados ou erros
        }
    }

})();