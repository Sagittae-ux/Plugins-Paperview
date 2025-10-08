(function () {
    if (app.selection.length !== 1 || !(app.selection[0] instanceof TextFrame)) {
        alert("Selecione apenas uma caixa de texto.");
        return;
    }

    var tf = app.selection[0];
    var texto = tf.contents;

    // Palavras a ignorar
    var ignorar = ["de", "da", "das", "do", "dos"];

    // Transforma em Title Case com exceções
    var resultado = texto.replace(/\b\w+\b/g, function (palavra) {
        var lower = palavra.toLowerCase();
        return lower.charAt(0).toUpperCase() + lower.slice(1);
    });

    tf.contents = resultado;
})();
