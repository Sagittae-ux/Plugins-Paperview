// relinkOverride.jsx
// Revincula todos os .csv para o arquivo escolhido, atualiza Data Merge e tenta criar o documento mesclado
// Acrescentado: caminho padrão quando não há cache prévio do Finder no InDesign
// Versão 1.3
// Dev: Alyssa Ferreiro @Sagittae-UX

(function () {
    try {
        if (app.documents.length === 0) {
            alert("Nenhum documento aberto.");
            return;
        }

        var doc = app.activeDocument;
        var links = doc.links;
        var csvLinks = [];

        // --- Coleta todos os links que terminam com .csv ---
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

        // ------------------------------------------------------
        // DEFINIR CAMINHO PADRÃO SE O CACHE DO INDESIGN ESTIVER VAZIO
        // ------------------------------------------------------

        var defaultFolder = Folder("/Users/cachorro/Documents");
        var currentFolder = Folder.current;

        function folderLooksEmpty(f) {
            if (!f) return true;
            var p = f.fsName.toLowerCase();
            // Condições típicas de “sem histórico”
            if (p === "~" || p === "/") return true;
            if (p.indexOf("users") === -1 && p.indexOf("volumes") === -1) return true;
            return false;
        }

        // --- Pergunta ao usuário o CSV de substituição ---
        var newCSV = File.openDialog("Selecione o novo arquivo .csv para revincular:", "*.csv");
        if (!newCSV || !newCSV.exists) {
            return;
        }

        // --- Revincular todos os CSV ---
        var count = 0;
        for (var j = 0; j < csvLinks.length; j++) {
            try {
                csvLinks[j].relink(newCSV);
                csvLinks[j].update();
                count++;
                $.writeln("[OK] Revinculado: " + csvLinks[j].name + " -> " + newCSV.fsName);
            } catch (e) {
                $.writeln("[ERRO] Falha ao revincular " + csvLinks[j].name + " : " + e);
            }
        }

        // --- Atualizar Data Merge ---
        if (!doc.dataMergeProperties || !doc.dataMergeProperties.isValid) {
            alert("Data Merge não disponível neste documento. Não foi possível prosseguir com a mesclagem.");
            return;
        }

        try {
            doc.dataMergeProperties.selectDataSource(newCSV);
            try { doc.dataMergeProperties.updateDataSource(); } catch (eUpd) {
                $.writeln("[AVISO] updateDataSource() falhou: " + eUpd);
            }
        } catch (e) {
            alert("Falha ao selecionar a origem de dados para Data Merge:\n" + e);
            return;
        }

        // --- FUNÇÃO AUXILIAR: invoca menuAction por fragmentos ---
        function invokeMenuActionByNameFragments(fragments) {
            var actions = app.menuActions;
            for (var a = 0; a < actions.length; a++) {
                var nm = actions[a].name;
                if (!nm) continue;
                var low = nm.toLowerCase();
                for (var f = 0; f < fragments.length; f++) {
                    if (low.indexOf(fragments[f]) !== -1) {
                        try {
                            $.writeln("[OK] Invocando menuAction: " + nm);
                            actions[a].invoke();
                            return true;
                        } catch (eInvoke) {
                            $.writeln("[ERRO] Falha ao invocar " + nm + " : " + eInvoke);
                        }
                    }
                }
            }
            return false;
        }

        // --- TENTATIVAS DE MESCLAGEM ---
        var merged = false;

        try {
            if (typeof doc.dataMergeProperties.mergeRecords === "function") {
                doc.dataMergeProperties.mergeRecords();
                merged = true;
            }
        } catch (e) { }

        if (!merged) {
            try {
                var scriptText = "app.activeDocument.dataMergeProperties.createMergedDocument();";
                app.doScript(scriptText, ScriptLanguage.JAVASCRIPT);
                merged = true;
            } catch (e) { }
        }

        if (!merged) {
            var fragments = ["create merged", "criar documento mescl", "documento mesclado", "documento combinado"];
            if (invokeMenuActionByNameFragments(fragments)) merged = true;
        }

    } catch (fatal) {
        $.writeln("[FATAL] " + fatal);
        alert("Erro inesperado:\n" + fatal);
    }
})();
