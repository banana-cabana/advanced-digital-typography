
// InDesign Script: Assign random font sizes to all words in the selected text box
if (app.documents.length > 0 && app.selection.length > 0 && app.selection[0].constructor.name === "TextFrame") {
    var textFrame = app.selection[0];
    var text = textFrame.contents;
    var words = textFrame.words;

    for (var i = 0; i < words.length; i++) {
        var randomFontSize = Math.floor(Math.random() * 20) + 10; // Random size between 10 and 30
        words[i].pointSize = randomFontSize;
    }

    alert("Random font sizes applied to all words in the selected text box.");
} else {
    alert("Please select a text box with content.");
}
// Create a dialog box for user input
var dialog = app.dialogs.add({ name: "Set Font Size Range" });
with (dialog.dialogColumns.add()) {
    staticTexts.add({ staticLabel: "Minimum Font Size:" });
    var minSizeField = integerEditboxes.add({ editValue: 10 });
    staticTexts.add({ staticLabel: "Maximum Font Size:" });
    var maxSizeField = integerEditboxes.add({ editValue: 30 });
}

// Show the dialog and get user input
if (dialog.show() === true) {
    var minSize = minSizeField.editValue;
    var maxSize = maxSizeField.editValue;
    dialog.destroy();

    if (minSize > 0 && maxSize > minSize) {
        for (var i = 0; i < words.length; i++) {
            var randomFontSize = Math.floor(Math.random() * (maxSize - minSize + 1)) + minSize;
            words[i].pointSize = randomFontSize;
        }
        alert("Random font sizes applied to all words in the selected text box.");
    } else {
        alert("Invalid input. Minimum size must be greater than 0 and less than maximum size.");
    }
} else {
    dialog.destroy();
    alert("Operation canceled.");
}