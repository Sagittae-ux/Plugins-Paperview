// Atalho simples para revinculação de arquivo .csv sem a necessidade de escolher o arquivo específico

// Dev: Alyssa Ferreiro @Sagittae-UX

(function () {
    if (app.documents.length === 0) {
        alert("Nenhum documento aberto.");
        return;
    }

    var doc = app.activeDocument;
    var links = doc.links;
    var csvLinks = [];

    // Coleta todos os links que terminam com .csv
    for (var i = 0; i < links.length; i++) {
        var name = links[i].name.toLowerCase();
        if (name.match(/\.csv$/)) {
            csvLinks.push(links[i]);
        }
    }

    if (csvLinks.length === 0) {
        alert("Nenhum arquivo .csv vinculado encontrado neste documento.");
        return;
    }

    // Permite ao usuário escolher o novo arquivo CSV
    var newCSV = File.openDialog("Selecione o novo arquivo .csv para revincular:", "*.csv");

    // Revincula todos os .csv encontrados
    var count = 0;
    for (var j = 0; j < csvLinks.length; j++) {
        try {
            csvLinks[j].relink(newCSV);
            csvLinks[j].update();
            count++;
        } catch (e) {
            $.writeln("Erro ao revincular: " + csvLinks[j].name + " (" + e + ")");
        }
    }

})();
