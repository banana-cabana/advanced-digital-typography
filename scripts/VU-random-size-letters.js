// Randomly changes the text size for the current text frame in InDesign

// Ensure a document is open
if (app.documents.length > 0) {
    var doc = app.activeDocument;

    // Ensure a text frame is selected
    if (app.selection.length > 0 && app.selection[0] instanceof TextFrame) {
        var textFrame = app.selection[0];
        var text = textFrame.texts[0];

        // Create a dialog box for minimum and maximum font sizes
        var dialog = new Window("dialog", "Set Font Size Range");
        dialog.add("statictext", undefined, "Minimum Font Size:");
        var minSizeInput = dialog.add("edittext", undefined, "12");
        minSizeInput.characters = 5;

        dialog.add("statictext", undefined, "Maximum Font Size:");
        var maxSizeInput = dialog.add("edittext", undefined, "20");
        maxSizeInput.characters = 5;

        var confirmButton = dialog.add("button", undefined, "OK");
        var cancelButton = dialog.add("button", undefined, "Cancel");

        confirmButton.onClick = function () {
            dialog.close(1);
        };

        cancelButton.onClick = function () {
            dialog.close(0);
        };

        if (dialog.show() === 1) {
            var minSize = parseInt(minSizeInput.text, 10);
            var maxSize = parseInt(maxSizeInput.text, 10);

            if (isNaN(minSize) || isNaN(maxSize) || minSize <= 0 || maxSize <= 0 || minSize > maxSize) {
                alert("Please enter valid minimum and maximum font sizes.");
            } else {
                // Loop through each character and assign a random font size
                for (var i = 0; i < text.characters.length; i++) {
                    var randomFontSize = Math.floor(Math.random() * (maxSize - minSize + 1)) + minSize;
                    text.characters[i].pointSize = randomFontSize;
                }

                alert("Text sizes have been randomly changed!");
            }
        }
    } else {
        alert("Please select a text frame.");
    }
} else {
        alert("No document is open.");
    }