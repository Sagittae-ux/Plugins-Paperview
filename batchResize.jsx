#target illustrator

var canvasWidth = 1080;
var canvasHeight = 1350;

var inputFolder = Folder.selectDialog("Selecione a pasta com os arquivos a serem redimensionados");

if (inputFolder) {

    var files = inputFolder.getFiles(/\.(jpg|jpeg)$/i);

    if (files.length === 0) {
        alert("Erro: Nenhum arquivo encontrado na pasta de output.");
    } else {

        for (var i = 0; i < files.length; i++) {

            var doc = app.open(files[i]);

            var ab = doc.artboards[0];
            ab.artboardRect = [0, canvasHeight, canvasWidth, 0];

            var item = doc.pageItems[0];

            // Trocando para tamanho de canvas
            var currentWidth = item.width;
            var size = (canvasWidth / currentWidth) * 100;

            item.resize(
                size,
                size,
                true,
                true,
                true,
                true,
                size,
                Transformation.CENTER
            );

            // Centralizar no artboard
            var centerX = canvasWidth / 2;
            var centerY = canvasHeight / 2;

            item.position = [
                centerX - (item.width / 2),
                centerY + (item.height / 2)
            ];

            // Exportação
            var baseName = doc.name.replace(/\.[^\.]+$/, "");
            var output = new File(doc.path + "/" + baseName + ".jpg");

            var options = new ExportOptionsJPEG();
            options.qualitySetting = 100;
            options.horizontalScale = 100;
            options.verticalScale = 100;
            options.artBoardClipping = true;
            options.antiAliasing = true;

            doc.exportFile(output, ExportType.JPEG, options);

            doc.close(SaveOptions.DONOTSAVECHANGES);
        }

        alert("Processo concluído: " + files.length + " arquivos processados.");
    }
}