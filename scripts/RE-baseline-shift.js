var dlg = app.dialogs.add({name: "Baseline Shift Settings"});
var col = dlg.dialogColumns.add();
col.staticTexts.add({staticLabel: "Minimum baseline shift:"});
var minField = col.textEditboxes.add({editContents: "2"});
col.staticTexts.add({staticLabel: "Maximum baseline shift:"});
var maxField = col.textEditboxes.add({editContents: "8"});
col.staticTexts.add({staticLabel: "Apply to:"});
var targetDropdown = col.dropdowns.add({stringList: ["Letters", "Words"], selectedIndex: 0});
col.staticTexts.add({staticLabel: "Pattern:"});
var patternDropdown = col.dropdowns.add({stringList: ["Alternate", "Random"], selectedIndex: 0});

var result = dlg.show();
if (result === true) {
    var minShift = Number(minField.editContents);
    var maxShift = Number(maxField.editContents);
    var targetType = targetDropdown.stringList[targetDropdown.selectedIndex];
    var patternType = patternDropdown.stringList[patternDropdown.selectedIndex];
    dlg.destroy();

    if (
        app.selection.length === 1 &&
        app.selection[0].constructor.name === "TextFrame"
    ) {
        var textFrame = app.selection[0];
        var text = textFrame.texts[0];
        var items = (targetType === "Words") ? text.words : text.characters;

        app.doScript(function() {
            for (var i = 0; i < items.length; i++) {
                if (patternType === "Alternate") {
                    items[i].baselineShift = (i % 2 === 0) ? minShift : maxShift;
                } else if (patternType === "Random") {
                    items[i].baselineShift = Math.random() < 0.5 ? minShift : maxShift;
                }
            }
        }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Apply Baseline Shift");
    } else {
        alert("Please select a text frame.");
    }
} else {
    dlg.destroy();
}
