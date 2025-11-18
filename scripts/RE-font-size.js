/*
Prompt for min and max font size using a single dialog panel,
then enlarge each character in the selected text box from min to max size, left to right,
or randomly apply font sizes to characters.
Works in Adobe InDesign.
*/

if (app.documents.length === 0) {
    alert("No document open.");
    exit();
}

if (app.selection.length !== 1 || !(app.selection[0] instanceof TextFrame)) {
    alert("Please select a single text box.");
    exit();
}

// Create dialog panel for min and max font size and mode
var dlg = new Window("dialog", "Font Size Range");
dlg.orientation = "column";
dlg.alignChildren = "left";

dlg.add("statictext", undefined, "Minimum font size (pt):");
var minInput = dlg.add("edittext", undefined, "10");
minInput.characters = 6;

dlg.add("statictext", undefined, "Maximum font size (pt):");
var maxInput = dlg.add("edittext", undefined, "20");
maxInput.characters = 6;

dlg.add("statictext", undefined, "Apply to:");
var modeGroup = dlg.add("group");
var wholeRadio = modeGroup.add("radiobutton", undefined, "Whole text (gradient)");
var randomRadio = modeGroup.add("radiobutton", undefined, "Random characters");
wholeRadio.value = true;

var btnGroup = dlg.add("group");
btnGroup.alignment = "right";
var okBtn = btnGroup.add("button", undefined, "OK");
var cancelBtn = btnGroup.add("button", undefined, "Cancel");

if (dlg.show() !== 1) exit();

var minSize = parseFloat(minInput.text);
var maxSize = parseFloat(maxInput.text);

if (isNaN(minSize) || minSize <= 0) {
    alert("Invalid minimum font size.");
    exit();
}
if (isNaN(maxSize) || maxSize < minSize) {
    alert("Invalid maximum font size.");
    exit();
}

var textFrame = app.selection[0];
var story = textFrame.parentStory;
var undoName = "Enlarge Characters Left to Right or Random";

app.doScript(function () {
    for (var i = 0; i < story.paragraphs.length; i++) {
        var para = story.paragraphs[i];
        var charCount = para.characters.length;
        for (var j = 0; j < charCount; j++) {
            var size;
            if (wholeRadio.value) {
                // Linear interpolation between minSize and maxSize
                size = minSize + ((maxSize - minSize) * (j / Math.max(charCount - 1, 1)));
            } else {
                // Random size between min and max
                size = minSize + Math.random() * (maxSize - minSize);
            }
            para.characters[j].pointSize = size;
        }
    }
}, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, undoName);