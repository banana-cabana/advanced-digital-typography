/*
 * Advanced Text Highlighting & Effects
 * Combines multiple text manipulation features:
 * 1. Extract styled/unstyled text to new frame (with choice)
 * 2. Apply stroke gradient across multiple text frames as one continuous gradient
 */

// Check if InDesign is running and a document is open
if (app.documents.length === 0) {
    alert("Please open a document first.");
    exit();
}

// Helper function to get all character style names
function getCharacterStyleNames() {
    var styleNames = [];
    try {
        var doc = app.activeDocument;
        var characterStyles = doc.characterStyles.everyItem().getElements();
        
        for (var i = 0; i < characterStyles.length; i++) {
            var styleName = characterStyles[i].name;
            // Skip [None] and [Ohne] since they're handled by the "Unstyled" option
            if (styleName !== "[None]" && styleName !== "[Ohne]") {
                styleNames.push(styleName);
            }
        }
        
        // If no custom styles found, add a placeholder
        if (styleNames.length === 0) {
            styleNames.push("(No custom character styles found)");
        }
        
    } catch (e) {
        styleNames.push("(Error loading character styles)");
    }
    
    return styleNames;
}

// Main function
function main() {
    // Check if something is selected
    if (app.selection.length === 0) {
        alert("Please select at least one text frame.");
        return;
    }
    
    // Create the main dialog
    var dialog = app.dialogs.add({name: "Advanced Text Highlighting & Effects", canCancel: true});
    
    var col = dialog.dialogColumns.add();
    
    // Title
    var rowTitle = col.dialogRows.add();
    rowTitle.staticTexts.add({staticLabel: "Choose your text effect:", minWidth: 400});
    
    // Radio button group for mode selection
    var rowMode = col.dialogRows.add();
    var modeGroup = rowMode.radiobuttonGroups.add();
    var extractStyledOption = modeGroup.radiobuttonControls.add({staticLabel: "Extract STYLED text (hide unstyled)", checkedState: true});
    
    var rowMode2 = col.dialogRows.add();
    var extractUnstyledOption = modeGroup.radiobuttonControls.add({staticLabel: "Extract UNSTYLED text (hide styled)"});
    
    var rowMode3 = col.dialogRows.add();
    var strokeGradientOption = modeGroup.radiobuttonControls.add({staticLabel: "Apply stroke gradient to text"});
    
    // Separator
    var rowSep1 = col.dialogRows.add();
    rowSep1.staticTexts.add({staticLabel: "----------------------------------------", minWidth: 400});
    
    // Extract options (visible for both extract modes)
    var rowExtractTitle = col.dialogRows.add();
    rowExtractTitle.staticTexts.add({staticLabel: "Extract Options:"});
    
    var rowExtract1 = col.dialogRows.add();
    var moveLayerCheckbox = rowExtract1.checkboxControls.add({
        staticLabel: "Move duplicate to a new layer above",
        checkedState: true
    });
    
    // Separator
    var rowSep2 = col.dialogRows.add();
    rowSep2.staticTexts.add({staticLabel: "----------------------------------------", minWidth: 400});
    
    // Stroke gradient options
    var rowStrokeTitle = col.dialogRows.add();
    rowStrokeTitle.staticTexts.add({staticLabel: "Stroke Gradient Options:"});
    
    var rowStroke1 = col.dialogRows.add();
    rowStroke1.staticTexts.add({staticLabel: "Max stroke (top/bottom):"});
    var maxStrokeInput = rowStroke1.realEditboxes.add({
        editValue: 2.0,
        minValue: 0.1,
        smallNudge: 0.1,
        largeNudge: 0.5
    });
    rowStroke1.staticTexts.add({staticLabel: "pt"});
    
    var rowStroke2 = col.dialogRows.add();
    rowStroke2.staticTexts.add({staticLabel: "Min stroke (middle):"});
    var minStrokeInput = rowStroke2.realEditboxes.add({
        editValue: 0.25,
        minValue: 0.1,
        smallNudge: 0.1,
        largeNudge: 0.5
    });
    rowStroke2.staticTexts.add({staticLabel: "pt"});
    
    var rowStroke3 = col.dialogRows.add();
    var duplicateCheckbox = rowStroke3.checkboxControls.add({
        staticLabel: "Work on duplicate (preserves original)",
        checkedState: true
    });
    
    var rowStroke3a = col.dialogRows.add();
    rowStroke3a.staticTexts.add({staticLabel: "Apply gradient to:"});
    var targetGroup = rowStroke3a.radiobuttonGroups.add();
    var unstyledOption = targetGroup.radiobuttonControls.add({staticLabel: "Unstyled text ([None])", checkedState: true});
    
    var rowStroke3b = col.dialogRows.add();
    rowStroke3b.staticTexts.add({staticLabel: ""});
    var specificStyleOption = targetGroup.radiobuttonControls.add({staticLabel: "Specific character style:"});
    
    var rowStroke3c = col.dialogRows.add();
    rowStroke3c.staticTexts.add({staticLabel: "Style name:"});
    var styleDropdown = rowStroke3c.dropdowns.add({
        stringList: getCharacterStyleNames(),
        selectedIndex: 0
    });
    
    var rowStroke3d = col.dialogRows.add();
    var reverseGradientCheckbox = rowStroke3d.checkboxControls.add({
        staticLabel: "Reverse gradient (thin-thick-thin instead of thick-thin-thick)",
        checkedState: false
    });
    
    var rowStroke4 = col.dialogRows.add();
    var strokeLayerCheckbox = rowStroke4.checkboxControls.add({
        staticLabel: "Move to new layer above",
        checkedState: true
    });
    
    // Separator
    var rowSep3 = col.dialogRows.add();
    rowSep3.staticTexts.add({staticLabel: "----------------------------------------", minWidth: 400});
    
    // Show dialog
    var result = dialog.show();
    
    if (result === true) {
        var mode = "extractStyled"; // default
        if (extractUnstyledOption.checkedState) {
            mode = "extractUnstyled";
        } else if (strokeGradientOption.checkedState) {
            mode = "strokeGradient";
        }
        
        var moveToNewLayer = moveLayerCheckbox.checkedState;
        var maxStroke = maxStrokeInput.editValue;
        var minStroke = minStrokeInput.editValue;
        var workOnDuplicate = duplicateCheckbox.checkedState;
        var strokeLayerMove = strokeLayerCheckbox.checkedState;
        var targetUnstyled = unstyledOption.checkedState;
        var reverseGradient = reverseGradientCheckbox.checkedState;
        var targetStyleName = "";
        var styleListLength = styleDropdown.stringList.length;
        var selectedStyleIndex = styleDropdown.selectedIndex;
        
        if (!targetUnstyled && styleListLength > 0 && selectedStyleIndex >= 0) {
            targetStyleName = styleDropdown.stringList[selectedStyleIndex];
        }
        
        dialog.destroy();
        
        // Validate stroke inputs if needed
        if (mode === "strokeGradient" && maxStroke <= minStroke) {
            alert("Maximum stroke must be greater than minimum stroke!");
            return;
        }
        
        // Validate specific style name if needed
        if (mode === "strokeGradient" && !targetUnstyled && (styleListLength === 0 || targetStyleName.length === 0)) {
            alert("No character styles available or none selected!");
            return;
        }
        
        // Execute based on mode
        switch (mode) {
            case "extractStyled":
                app.doScript(function() {
                    processExtraction(moveToNewLayer, false); // false = show styled, hide unstyled
                }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Extract Styled Text");
                alert("Styled text extraction complete!");
                break;
                
            case "extractUnstyled":
                app.doScript(function() {
                    processExtraction(moveToNewLayer, true); // true = show unstyled, hide styled
                }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Extract Unstyled Text");
                alert("Unstyled text extraction complete!");
                break;
                
            case "strokeGradient":
                app.doScript(function() {
                    processStrokeGradient(maxStroke, minStroke, workOnDuplicate, strokeLayerMove, targetUnstyled, targetStyleName, reverseGradient);
                }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Apply Multi-Frame Stroke Gradient");
                alert("Multi-frame stroke gradient applied!");
                break;
        }
        
    } else {
        dialog.destroy();
    }
}

// Process text extraction (unified for both modes)
function processExtraction(moveToNewLayer, reverseMode) {
    var selection = app.selection;
    var doc = app.activeDocument;
    var newLayer = null;
    
    // Create or get the [Hidden] character style
    var hiddenStyle = getOrCreateHiddenStyle(doc, reverseMode);
    
    if (hiddenStyle === null) {
        alert("Could not create hidden character style. Aborting.");
        return;
    }
    
    // Create new layer if requested
    if (moveToNewLayer) {
        var layerName = reverseMode ? "Unstyled Text Extract" : "Styled Text Extract";
        try {
            newLayer = doc.layers.add({name: layerName});
            newLayer.move(LocationOptions.BEFORE, doc.layers[0]);
        } catch (e) {
            // If layer already exists, use it
            try {
                newLayer = doc.layers.item(layerName);
            } catch (e2) {
                newLayer = null;
            }
        }
    }
    
    var processedCount = 0;
    
    for (var i = 0; i < selection.length; i++) {
        var item = selection[i];
        
        // Check if it's a text frame
        if (item.constructor.name === "TextFrame") {
            extractText(item, hiddenStyle, newLayer, reverseMode);
            processedCount++;
        }
    }
    
    if (processedCount === 0) {
        alert("No text frames were selected.");
    }
}

// Process stroke gradient across multiple frames
function processStrokeGradient(maxStroke, minStroke, workOnDuplicate, moveToNewLayer, targetUnstyled, targetStyleName, reverseGradient) {
    var selection = app.selection;
    var doc = app.activeDocument;
    var newLayer = null;
    
    // Create new layer if requested
    if (moveToNewLayer) {
        try {
            newLayer = doc.layers.add({name: "Multi-Frame Stroke Gradient"});
            newLayer.move(LocationOptions.BEFORE, doc.layers[0]);
        } catch (e) {
            // If layer already exists, use it
            try {
                newLayer = doc.layers.itemByName("Multi-Frame Stroke Gradient");
            } catch (e2) {
                newLayer = null;
            }
        }
    }
    
    // Collect all text frames and their lines
    var allFrameLines = [];
    var targetFrames = [];        for (var i = 0; i < selection.length; i++) {
            var item = selection[i];
            
            // Validate the selection item still exists
            if (!item.isValid) {
                continue;
            }
            
            if (item.constructor.name === "TextFrame") {
                var targetFrame = item;
                
                // Duplicate if requested
                if (workOnDuplicate) {
                    try {
                        targetFrame = item.duplicate();
                        
                        // Validate the duplicate was created
                        if (!targetFrame.isValid) {
                            continue;
                        }
                        
                        // Move to new layer if specified
                        if (newLayer !== null) {
                            targetFrame.itemLayer = newLayer;
                        }
                        
                        targetFrame.bringToFront();
                    } catch (e) {
                        // Skip this frame if duplication fails
                        continue;
                    }
                }
                
                targetFrames.push(targetFrame);
                
                // Get all lines from this frame with validation
                if (targetFrame.isValid && targetFrame.contents !== "") {
                    try {
                        var lines = targetFrame.lines.everyItem().getElements();
                        for (var j = 0; j < lines.length; j++) {
                            if (lines[j].isValid) {
                                allFrameLines.push(lines[j]);
                            }
                        }
                    } catch (e) {
                        // Skip if can't access lines
                    }
                }
            }
        }
    
    if (allFrameLines.length === 0) {
        alert("No text lines found in selected frames.");
        return;
    }
    
    // Apply gradient across ALL lines from ALL frames
    applyMultiFrameStrokeGradient(allFrameLines, maxStroke, minStroke, targetUnstyled, targetStyleName, reverseGradient);
}

// Create or get the hidden character style
function getOrCreateHiddenStyle(doc, reverseMode) {
    var styleName = reverseMode ? "Hidden_Text_Unstyled" : "Hidden_Text_Styled";
    var hiddenStyle = null;
    
    // Check if style already exists
    try {
        hiddenStyle = doc.characterStyles.itemByName(styleName);
        if (hiddenStyle.isValid) {
            return hiddenStyle;
        }
    } catch (e) {
        // Style doesn't exist, will create it
    }
    
    // Create the hidden style with Paper color
    try {
        hiddenStyle = doc.characterStyles.add();
        hiddenStyle.name = styleName;
        
        // Try to get Paper swatch - should always exist in InDesign
        var paperSwatch = doc.swatches.itemByName("Paper");
        
        // Apply Paper color to the character style
        hiddenStyle.fillColor = paperSwatch;
        hiddenStyle.fillTint = 100;
        
        return hiddenStyle;
        
    } catch (e) {
        alert("Error creating style: " + e.message + " at line " + e.line);
        return null;
    }
}

// Extract text to new frame (unified function)
function extractText(originalFrame, hiddenStyle, targetLayer, reverseMode) {
    if (!originalFrame.isValid || originalFrame.contents === "") {
        return;
    }
    
    try {
        // Duplicate the text frame
        var duplicateFrame = originalFrame.duplicate();
        
        // Validate the duplicate was created successfully
        if (!duplicateFrame.isValid) {
            return;
        }
        
        // Move to new layer if specified
        if (targetLayer !== null) {
            duplicateFrame.itemLayer = targetLayer;
        }
        
        // Make sure duplicate is on top
        duplicateFrame.bringToFront();
        
        // Process the duplicate frame
        hideCharacters(duplicateFrame, hiddenStyle, reverseMode);
        
    } catch (e) {
        alert("Error processing frame: " + e.message);
    }
}

// Hide characters based on mode
function hideCharacters(textFrame, hiddenStyle, reverseMode) {
    try {
        // Validate frame still exists
        if (!textFrame.isValid) {
            return;
        }
        
        var allChars = textFrame.characters.everyItem().getElements();
        
        if (allChars.length === 0) {
            return;
        }
        
        var hiddenCount = 0;
        var visibleCount = 0;
        
        // Process each character with validation
        for (var i = 0; i < allChars.length; i++) {
            try {
                var currentChar = allChars[i];
                
                // Validate character still exists
                if (!currentChar.isValid) {
                    continue;
                }
                
                // Check if this character has a character style applied
                var hasStyle = false;
                var styleName = "";
                
                try {
                    styleName = currentChar.appliedCharacterStyle.name;
                    
                    // Count it as "styled" if it's NOT [None]/[Ohne] and NOT our hidden styles
                    if (styleName !== "[None]" && 
                        styleName !== "[Ohne]" && 
                        styleName !== "Hidden_Text_Styled" &&
                        styleName !== "Hidden_Text_Unstyled") {
                        hasStyle = true;
                    }
                } catch (e) {
                    hasStyle = false;
                }
                
                var shouldHide = false;
                
                if (reverseMode) {
                    // REVERSE MODE: Hide characters WITH styles, show unstyled
                    shouldHide = hasStyle;
                } else {
                    // NORMAL MODE: Hide characters WITHOUT styles, show styled
                    shouldHide = !hasStyle;
                }
                
                if (shouldHide) {
                    // Apply the hidden style with validation
                    try {
                        if (currentChar.isValid) {
                            currentChar.applyCharacterStyle(hiddenStyle, true);
                            
                            // Force the fill color at character level too
                            try {
                                var paperSwatch = app.activeDocument.swatches.itemByName("Paper");
                                currentChar.fillColor = paperSwatch;
                            } catch (e) {
                                // Ignore if this fails
                            }
                            
                            hiddenCount++;
                        }
                    } catch (e) {
                        // Skip if character manipulation fails
                    }
                } else {
                    visibleCount++;
                }
                
            } catch (e) {
                // Skip if there's an error with this character
            }
        }
        
        var modeText = reverseMode ? "unstyled" : "styled";
        alert("Visible " + modeText + " characters: " + visibleCount + "\nHidden characters: " + hiddenCount + "\nTotal: " + allChars.length);
        
    } catch (e) {
        alert("Error in hideCharacters: " + e.message);
    }
}

// Apply stroke gradient across multiple frames as one continuous gradient
function applyMultiFrameStrokeGradient(allLines, maxStroke, minStroke, targetUnstyled, targetStyleName, reverseGradient) {
    try {
        // Get black swatch
        var blackSwatch = null;
        try {
            blackSwatch = app.activeDocument.swatches.itemByName("Black");
        } catch (e) {
            // Try to get any black color
            try {
                blackSwatch = app.activeDocument.colors.itemByName("Black");
            } catch (e2) {
                alert("Could not find Black swatch!");
                return;
            }
        }
        
        var totalLines = allLines.length;
        
        if (totalLines === 0) {
            return;
        }
        
        var charsProcessed = 0;
        var charsStroked = 0;
        
        // Calculate the maximum number of characters in any line for horizontal positioning
        var maxCharsInLine = 0;
        for (var i = 0; i < totalLines; i++) {
            var lineChars = allLines[i].characters.everyItem().getElements();
            if (lineChars.length > maxCharsInLine) {
                maxCharsInLine = lineChars.length;
            }
        }
        
        // Calculate stroke weight for each line
        for (var lineIndex = 0; lineIndex < totalLines; lineIndex++) {
            // Apply stroke based on target style selection
            var line = allLines[lineIndex];
            var chars = line.characters.everyItem().getElements();
            
            for (var charIndex = 0; charIndex < chars.length; charIndex++) {
                // Calculate DIAGONAL position (top-left to bottom-right)
                var verticalPosition = lineIndex / Math.max(totalLines - 1, 1); // 0 to 1 (top to bottom)
                var horizontalPosition = charIndex / Math.max(maxCharsInLine - 1, 1); // 0 to 1 (left to right)
                
                // Diagonal position: average of vertical and horizontal positions
                // This creates a gradient from top-left (0,0) to bottom-right (1,1)
                var diagonalPosition = (verticalPosition + horizontalPosition) / 2;
                
                // Distance from middle diagonal (0.5)
                var distanceFromMiddle = Math.abs(diagonalPosition - 0.5) * 2; // 0 at middle, 1 at edges
                
                // Calculate stroke weight based on reverse setting
                var strokeWeight;
                if (reverseGradient) {
                    // REVERSED: thin-thick-thin (min at edges, max in middle)
                    strokeWeight = maxStroke - (maxStroke - minStroke) * distanceFromMiddle;
                } else {
                    // NORMAL: thick-thin-thick (max at edges, min in middle)
                    strokeWeight = minStroke + (maxStroke - minStroke) * distanceFromMiddle;
                }
                try {
                    var currentChar = chars[charIndex];
                    
                    // Validate character still exists
                    if (!currentChar.isValid) {
                        continue;
                    }
                    
                    charsProcessed++;
                    
                    // Check character style based on target selection
                    var shouldApplyStroke = false;
                    
                    try {
                        var styleName = currentChar.appliedCharacterStyle.name;
                        
                        if (targetUnstyled) {
                            // Apply to unstyled characters ([None] or [Ohne])
                            if (styleName === "[None]" || styleName === "[Ohne]") {
                                shouldApplyStroke = true;
                            }
                        } else {
                            // Apply to specific character style
                            if (styleName === targetStyleName) {
                                shouldApplyStroke = true;
                            }
                        }
                    } catch (e) {
                        // If we can't get the style name, treat as unstyled
                        if (targetUnstyled) {
                            shouldApplyStroke = true;
                        }
                    }
                    
                    // Apply stroke if this character matches our target
                    if (shouldApplyStroke && currentChar.isValid) {
                        try {
                            // Enable stroke
                            currentChar.strokeWeight = strokeWeight;
                            currentChar.strokeColor = blackSwatch;
                            currentChar.strokeTint = 100;
                            
                            // Make sure stroke is visible
                            try {
                                currentChar.strokeAlignment = StrokeAlignment.CENTER_ALIGNMENT;
                            } catch (e) {
                                // Stroke alignment might not work on text
                            }
                            
                            // Adjust text size based on stroke weight (both above and below 1pt)
                            try {
                                // Get the current font size (or default if not set)
                                var currentSize = currentChar.pointSize;
                                if (isNaN(currentSize) || currentSize === undefined) {
                                    // Try to get from parent paragraph if character size is not set
                                    try {
                                        currentSize = currentChar.parentStory.pointSize;
                                    } catch (e) {
                                        currentSize = 12; // Default fallback size
                                    }
                                }
                                
                                var newSize, baselineShift;
                                
                                if (strokeWeight < 1.0) {
                                    // For strokes below 1pt: reduce size (35% to 100%)
                                    var sizeFactor = 0.35 + (strokeWeight / 1.0) * 0.65; // Changed from 0.5 + 0.5 to 0.35 + 0.65
                                    newSize = currentSize * sizeFactor;
                                    
                                    // Calculate baseline shift to keep characters aligned
                                    var sizeReduction = currentSize - newSize;
                                    baselineShift = sizeReduction * 0.4; // Increased from 0.3 to 0.4
                                    
                                } else if (strokeWeight > 1.0) {
                                    // For strokes above 1pt: increase size (100% to 175%)
                                    // Cap the maximum stroke for size calculation to prevent extreme sizes
                                    var cappedStroke = Math.min(strokeWeight, 4.0); // Cap at 4pt for size calculation
                                    var sizeFactor = 1.0 + ((cappedStroke - 1.0) / 3.0) * 0.75; // 1.0 to 1.75 (increased from 0.5)
                                    newSize = currentSize * sizeFactor;
                                    
                                    // Calculate baseline shift to keep larger characters aligned
                                    var sizeIncrease = newSize - currentSize;
                                    baselineShift = -sizeIncrease * 0.4; // Increased from 0.3 to 0.4, negative to shift down
                                    
                                } else {
                                    // Stroke is exactly 1pt: no size change
                                    newSize = currentSize;
                                    baselineShift = 0;
                                }
                                
                                // Apply the new size and baseline shift
                                currentChar.pointSize = newSize;
                                currentChar.baselineShift = baselineShift;
                                
                            } catch (e) {
                                // Skip if text size adjustment fails
                            }
                            
                            charsStroked++;
                        } catch (e) {
                            // Skip if stroke application fails
                        }
                    }
                    
                } catch (e) {
                    // Skip character if error
                }
            }
        }
        
        var gradientType = reverseGradient ? "diagonal thin-thick-thin" : "diagonal thick-thin-thick";
        var targetDescription = targetUnstyled ? "unstyled characters" : "characters with style '" + targetStyleName + "'";
        alert("Multi-frame diagonal gradient (top-left to bottom-right, " + gradientType + ") applied to " + targetDescription + "!\nProcessed " + charsProcessed + " chars across " + totalLines + " lines\nStroked " + charsStroked + " target characters");
        
    } catch (e) {
        alert("Error applying multi-frame stroke gradient: " + e.message);
    }
}

// Run the main function
main();
