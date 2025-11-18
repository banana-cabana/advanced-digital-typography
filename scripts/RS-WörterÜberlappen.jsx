(function () {
    if (app.documents.length === 0) {
        alert("Bitte öffne ein Dokument.");
        return;
    }
    if (app.selection.length !== 1 || !(app.selection[0] instanceof TextFrame)) {
        alert("Bitte wähle genau eine Textbox aus.");
        return;
    }

    var textFrame = app.selection[0];
    var story = textFrame.parentStory;
    if (story.characters.length === 0) {
        alert("Textbox enthält keinen Text.");
        return;
    }

    app.doScript(function () {
        var baseSize = 25;
        var minSize = 12;
        var maxSize = 50;
        var specialCount = 10;

        // Alle Zeichen auf Grundgröße setzen
        for (var c = 0; c < story.characters.length; c++) {
            try {
                story.characters[c].pointSize = baseSize;
            } catch (e) {
                // Steuerzeichen ignorieren
            }
        }

        // Kandidaten sammeln (nur sichtbare Zeichen)
        var candidates = [];
        for (var i = 0; i < story.characters.length; i++) {
            var ch = story.characters[i];
            var content = ch.contents;
            if (!content || /^\s$/.test(content) || content === "\r" || content === "\n") {
                continue;
            }
            candidates.push(i);
        }
        if (candidates.length === 0) return;

        // Shuffle der Kandidaten
        for (var k = candidates.length - 1; k > 0; k--) {
            var swap = Math.floor(Math.random() * (k + 1));
            var tmp = candidates[k];
            candidates[k] = candidates[swap];
            candidates[swap] = tmp;
        }

        // Bis zu 10 Zeichen zufällig neu skalieren (andere Größe als 25 pt)
        var limit = Math.min(specialCount, candidates.length);
        for (var t = 0; t < limit; t++) {
            var idx = candidates[t];
            var newSize;
            do {
                newSize = Math.floor(Math.random() * (maxSize - minSize + 1)) + minSize;
            } while (newSize === baseSize);
            try {
                story.characters[idx].pointSize = newSize;
            } catch (e2) {
                // Wenn Setzen fehlschlägt, weiter zum nächsten
            }
        }
    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Random 10 Letter Sizes");
})();