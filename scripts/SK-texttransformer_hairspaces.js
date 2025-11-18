// InDesign Script: Format Text with Progress Bar and Character Selection

// Function to insert a chosen character after every character in the text
function addCharacters(text, insertChar) {
    var newText = "";
    for (var i = 0; i < text.length; i++) {
        newText += text[i] + insertChar;
    }
    return newText;
}

// Check if a document is open
if (app.documents.length === 0) {
    alert("Please open a document and select at least one text frame.");
} else {
    var doc = app.activeDocument;
    var sel = app.selection;

    if (sel.length === 0) {
        alert("Please select at least one text frame.");
    } else {
        // Dialog for selecting the character to insert
        var dlg = new Window("dialog", "Character Selection");
        dlg.orientation = "column";
        dlg.alignChildren = "left";

        dlg.add("statictext", undefined, "Choose the character to insert after each letter:");

        var dropdown = dlg.add("dropdownlist", undefined, [
            "Hair Space (\\u200A)",
            "Thin Space (\\u2009)",
            "Non-breaking Space (\\u00A0)",
            "En Space (\\u2002)",
            "Em Space (\\u2003)",
            "Custom Character..."
        ]);
        dropdown.selection = 0; // Default: Hair Space

        var customGroup = dlg.add("group");
        customGroup.add("statictext", undefined, "Custom:");
        var customInput = customGroup.add("edittext", undefined, "");
        customInput.characters = 5;
        customInput.enabled = false;

        // Enable text input when "Custom Character..." is chosen
        dropdown.onChange = function() {
            customInput.enabled = (dropdown.selection.index === 5);
        };

        // OK / Cancel buttons
        var btnGroup = dlg.add("group");
        btnGroup.alignment = "right";
        var okBtn = btnGroup.add("button", undefined, "OK", { name: "ok" });
        var cancelBtn = btnGroup.add("button", undefined, "Cancel", { name: "cancel" });

        if (dlg.show() != 1) exit();

        var selected = dropdown.selection.index;
        var insertChar;

        switch (selected) {
            case 0: insertChar = "\u200A"; break; // Hair Space
            case 1: insertChar = "\u2009"; break; // Thin Space
            case 2: insertChar = "\u00A0"; break; // Non-breaking Space
            case 3: insertChar = "\u2002"; break; // En Space
            case 4: insertChar = "\u2003"; break; // Em Space
            case 5:
                insertChar = customInput.text || "\u200A"; // Use custom char, fallback: Hair Space
                break;
        }

        // Create progress bar
        var progressWin = new Window("palette", "Formatting text...");
        progressWin.progressBar = progressWin.add("progressbar", undefined, 0, sel.length);
        progressWin.progressBar.preferredSize.width = 300;
        progressWin.show();

        // Process each selected text frame
        for (var i = 0; i < sel.length; i++) {
            var item = sel[i];
            if (item instanceof TextFrame) {
                var text = item.contents;

                // 1. Justify text
                item.texts[0].justification = Justification.FULLY_JUSTIFIED;

                // 2. Remove paragraph breaks
                text = text.replace(/[\r\n]+/g, "");

                // 3. Add chosen character after each character
                text = addCharacters(text, insertChar);

                item.contents = text;
            }

            // Update progress bar
            progressWin.progressBar.value = i + 1;
        }

        progressWin.close();
        alert("Done! Text has been formatted, paragraphs removed, and the chosen character added after each letter.");
    }
}