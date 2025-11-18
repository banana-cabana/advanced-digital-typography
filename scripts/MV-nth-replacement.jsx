/*
 * N-th Word/Letter Replacement Script
 * Replaces every n-th word or letter in selected text frames with a character style
 */

// Check if InDesign is running and a document is open
if (app.documents.length === 0) {
    alert("Please open a document first.");
    exit();
}

// Main function
function main() {
    // Get all available character styles
    var charStyles = app.activeDocument.characterStyles;
    var charStyleNames = ["[None]"]; // Add None option first
    
    for (var i = 0; i < charStyles.length; i++) {
        charStyleNames.push(charStyles[i].name);
    }
    
    if (charStyleNames.length === 1) {
        alert("No character styles found in this document. Please create at least one character style first.");
        return;
    }
    
    // Create the dialog
    var dialog = app.dialogs.add({name: "N-th Word/Letter Replacement", canCancel: true});
    
    with (dialog.dialogColumns.add()) {
        with (dialogRows.add()) {
            staticTexts.add({staticLabel: "Mode:"});
            var modeDropdown = dropdowns.add({
                stringList: ["Replace Words", "Replace Letters"],
                selectedIndex: 0
            });
        }
        
        // Separator
        with (dialogRows.add()) {
            staticTexts.add({staticLabel: "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━", minWidth: 250});
        }
        
        with (dialogRows.add()) {
            staticTexts.add({staticLabel: "Replace Every N-th:"});
            var nthInput = integerEditboxes.add({
                editValue: 2,
                minValue: 1
            });
        }
        
        with (dialogRows.add()) {
            staticTexts.add({staticLabel: "Start at Position:"});
            var startPosInput = integerEditboxes.add({
                editValue: 1,
                minValue: 1
            });
        }
        
        // Separator
        with (dialogRows.add()) {
            staticTexts.add({staticLabel: "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━", minWidth: 250});
        }
        
        with (dialogRows.add()) {
            staticTexts.add({staticLabel: "Character Style:"});
            var charStyleDropdown = dropdowns.add({
                stringList: charStyleNames,
                selectedIndex: 0
            });
        }
    }
    
    // Show dialog
    var result = dialog.show();
    
    if (result === true) {
        // Get values
        var mode = modeDropdown.selectedIndex; // 0 = words, 1 = letters
        var nthValue = nthInput.editValue;
        var startPos = startPosInput.editValue;
        var selectedCharStyleIndex = charStyleDropdown.selectedIndex;
        var selectedCharStyle = null;
        
        // Get the actual character style object (skip if [None] is selected)
        if (selectedCharStyleIndex > 0) {
            selectedCharStyle = charStyles[selectedCharStyleIndex - 1]; // -1 because we added [None] at index 0
        }
        
        dialog.destroy();
        
        // Process the selection
        if (app.selection.length === 0) {
            alert("Please select at least one text frame.");
            return;
        }
        
        // Start undo group
        app.doScript(function() {
            processSelection(mode, nthValue, startPos, selectedCharStyle);
        }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "N-th Replacement");
        
        alert("Replacement complete!");
    } else {
        dialog.destroy();
    }
}

// Process selected text frames
function processSelection(mode, nthValue, startPos, charStyle) {
    var selection = app.selection;
    
    for (var i = 0; i < selection.length; i++) {
        var item = selection[i];
        
        // Check if it's a text frame or has text content
        if (item.constructor.name === "TextFrame") {
            processTextFrame(item, mode, nthValue, startPos, charStyle);
        } else if (item.constructor.name === "InsertionPoint" || 
                   item.constructor.name === "Character" || 
                   item.constructor.name === "Word" || 
                   item.constructor.name === "TextStyleRange" || 
                   item.constructor.name === "Line" || 
                   item.constructor.name === "Paragraph" || 
                   item.constructor.name === "Text") {
            // If text selection within a frame
            processTextObject(item.parent, mode, nthValue, startPos, charStyle);
        }
    }
}

// Process a text frame
function processTextFrame(textFrame, mode, nthValue, startPos, charStyle) {
    if (textFrame.contents === "") {
        return;
    }
    
    if (mode === 0) {
        // Replace words
        replaceNthWords(textFrame, nthValue, startPos, charStyle);
    } else {
        // Replace letters
        replaceNthLetters(textFrame, nthValue, startPos, charStyle);
    }
}

// Process text object (for text selections)
function processTextObject(textObj, mode, nthValue, startPos, charStyle) {
    if (mode === 0) {
        replaceNthWords(textObj, nthValue, startPos, charStyle);
    } else {
        replaceNthLetters(textObj, nthValue, startPos, charStyle);
    }
}

// Replace every n-th word
function replaceNthWords(textContainer, nthValue, startPos, charStyle) {
    try {
        var allWords = textContainer.words.everyItem().getElements();
        
        if (allWords.length === 0) {
            return;
        }
        
        var counter = 0;
        for (var i = 0; i < allWords.length; i++) {
            counter++;
            
            // Check if this is an n-th position (accounting for start position)
            if (counter >= startPos && (counter - startPos) % nthValue === 0) {
                try {
                    // Apply character style (if one is selected)
                    if (charStyle !== null) {
                        allWords[i].applyCharacterStyle(charStyle);
                    }
                } catch (e) {
                    // Skip if there's an error with this particular word
                }
            }
        }
    } catch (e) {
        // Handle any errors silently or log them
    }
}

// Replace every n-th letter (character)
function replaceNthLetters(textContainer, nthValue, startPos, charStyle) {
    try {
        var allChars = textContainer.characters.everyItem().getElements();
        
        if (allChars.length === 0) {
            return;
        }
        
        var counter = 0;
        for (var i = 0; i < allChars.length; i++) {
            var charContent = allChars[i].contents;
            
            // Only count actual letters (not spaces or punctuation)
            if (charContent.match(/[a-zA-ZäöüÄÖÜßáéíóúÁÉÍÓÚàèìòùÀÈÌÒÙâêîôûÂÊÎÔÛãñõÃÑÕ]/)) {
                counter++;
                
                // Check if this is an n-th position (accounting for start position)
                if (counter >= startPos && (counter - startPos) % nthValue === 0) {
                    try {
                        // Apply character style (if one is selected)
                        if (charStyle !== null) {
                            allChars[i].applyCharacterStyle(charStyle);
                        }
                    } catch (e) {
                        // Skip if there's an error with this particular character
                    }
                }
            }
        }
    } catch (e) {
        // Handle any errors silently or log them
    }
}

// Run the main function
main();
