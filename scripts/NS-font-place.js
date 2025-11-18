/**
 * InDesign script: randomly distribute characters, words, or linked text frames from the selected text frame across the page.
 */

(function () {
    if (app.documents.length === 0) {
        alert("Please open a document.");
        return;
    }

    app.doScript(main, ScriptLanguage.JAVASCRIPT, null, UndoModes.ENTIRE_SCRIPT, "Inhalte zufällig verteilen");

    function main() {
        var doc = app.activeDocument;
        if (app.selection.length !== 1 || !(app.selection[0] instanceof TextFrame)) {
            alert("Please select exactly one text frame.");
            return;
        }

        var textFrame = app.selection[0];
        var text = textFrame.contents;
        if (!text || text.length === 0) {
            alert("The selected text frame is empty.");
            return;
        }

        var page = textFrame.parentPage;
        if (!page) {
            alert("The selected text frame is not placed on a page.");
            return;
        }

        var preferences = showDistributionDialog();
        if (!preferences) {
            return;
        }

        var pageTop = page.bounds[0];
        var pageLeft = page.bounds[1];
        var pageBottom = page.bounds[2];
        var pageRight = page.bounds[3];
        var pageWidth = pageRight - pageLeft;
        var pageHeight = pageBottom - pageTop;

        var baseStyle = captureStyle(textFrame);
        var frameSize = determineFrameSize(baseStyle, pageWidth, pageHeight);
        var items;
        if (preferences.mode === "letters") {
            items = collectCharacters(textFrame);
        } else if (preferences.mode === "words") {
            items = collectWords(textFrame);
        } else {
            items = collectTextFrames(textFrame);
        }

        if (!items || items.length === 0) {
            var emptyMessage;
            switch (preferences.mode) {
                case "letters":
                    emptyMessage = "No usable characters were found in the text.";
                    break;
                case "words":
                    emptyMessage = "No words could be extracted from the text.";
                    break;
                default:
                    emptyMessage = "No text frames could be identified.";
            }
            alert(emptyMessage);
            return;
        }

        var tokens = items.slice();
        if (preferences.mode === "letters" || preferences.mode === "words") {
            shuffleArray(tokens);
        }

        if (preferences.mode === "textFrames") {
            placeFramesRandomly(tokens, page, pageTop, pageLeft, pageBottom, pageRight);
        } else {
            placeTokensRandomly(tokens, page, baseStyle, frameSize, pageTop, pageLeft, pageBottom, pageRight, pageWidth, pageHeight);
        }

        alert("Done! The selection has been distributed.");
    }
})();

function collectCharacters(frame) {
    var chars = [];
    try {
        var items = frame.characters.everyItem().getElements();
        for (var i = 0; i < items.length; i++) {
            var contents = items[i].contents;
            if (!contents || contents === "\r" || contents === "\n") {
                continue;
            }
            chars.push(contents);
        }
    } catch (e) {
        var source = frame.contents || "";
        for (var index = 0; index < source.length; index++) {
            var ch = source.charAt(index);
            if (ch === "\r" || ch === "\n") {
                continue;
            }
            chars.push(ch);
        }
    }
    return chars;
}

function collectWords(frame) {
    var words = [];
    try {
        var items = frame.words.everyItem().getElements();
        for (var i = 0; i < items.length; i++) {
            var contents = items[i].contents;
            if (!contents) {
                continue;
            }
            var cleaned = contents.replace(/[\r\n]/g, "").trim();
            if (cleaned.length === 0) {
                continue;
            }
            words.push(cleaned);
        }
    } catch (e) {
        var source = frame.contents || "";
        var fallback = source.match(/\S+/g);
        if (fallback) {
            for (var j = 0; j < fallback.length; j++) {
                words.push(fallback[j]);
            }
        }
    }
    return words;
}

function collectTextFrames(frame) {
    var frames = [];
    try {
        var storyFrames = frame.parentStory ? frame.parentStory.textContainers : [frame];
        for (var i = 0; i < storyFrames.length; i++) {
            var textFrame = storyFrames[i];
            if (textFrame && textFrame.hasOwnProperty("contents")) {
                frames.push(textFrame);
            }
        }
    } catch (e) {}
    return frames;
}

function shuffleArray(array) {
    for (var i = array.length - 1; i > 0; i--) {
        var j = Math.floor(Math.random() * (i + 1));
        var temp = array[i];
        array[i] = array[j];
        array[j] = temp;
    }
}

function placeTokensRandomly(tokens, page, baseStyle, frameSize, pageTop, pageLeft, pageBottom, pageRight, pageWidth, pageHeight) {
    for (var i = 0; i < tokens.length; i++) {
        var token = tokens[i];
        var tf = page.textFrames.add();
        tf.geometricBounds = [pageTop, pageLeft, pageTop + frameSize, pageLeft + frameSize];
        tf.contents = token;

        try {
            applyStyle(tf, baseStyle);
            if (typeof FitOptions !== "undefined") {
                tf.fit(FitOptions.FRAME_TO_CONTENT);
            }
            var gb = tf.geometricBounds;
            var height = gb[2] - gb[0];
            var width = gb[3] - gb[1];
            var maxXOffset = Math.max(pageWidth - width, 0);
            var maxYOffset = Math.max(pageHeight - height, 0);
            var x = pageLeft + (maxXOffset === 0 ? 0 : Math.random() * maxXOffset);
            var y = pageTop + (maxYOffset === 0 ? 0 : Math.random() * maxYOffset);
            tf.geometricBounds = [y, x, y + height, x + width];
            clampFrameToPage(tf, pageTop, pageLeft, pageBottom, pageRight);
        } catch (e) {}
    }
}

function distributeWordsTopToBottom(words, page, baseStyle, top, left, bottom, right, frameSize) {
    var usableWidth = right - left;
    var usableHeight = bottom - top;
    if (usableWidth <= 0 || usableHeight <= 0) {
        return;
    }

    var maxColumns = Math.max(1, Math.floor(usableWidth / frameSize));
    var columnWidth = maxColumns > 0 ? usableWidth / maxColumns : usableWidth;
    var columnGap = maxColumns > 1 ? Math.max((usableWidth - (columnWidth * maxColumns)) / (maxColumns - 1), 0) : 0;
    var currentColumn = 0;
    var currentY = top;

    for (var i = 0; i < words.length; i++) {
        var word = words[i];
        var tf = page.textFrames.add();
        tf.geometricBounds = [currentY, left + currentColumn * (columnWidth + columnGap), currentY + frameSize, left + currentColumn * (columnWidth + columnGap) + columnWidth];
        tf.contents = word;

        try {
            applyStyle(tf, baseStyle);
            if (typeof FitOptions !== "undefined") {
                tf.fit(FitOptions.FRAME_TO_CONTENT);
            }
            var gb = tf.geometricBounds;
            var height = gb[2] - gb[0];
            var width = gb[3] - gb[1];
            tf.geometricBounds = [currentY, left + currentColumn * (columnWidth + columnGap), currentY + height, left + currentColumn * (columnWidth + columnGap) + width];
            clampFrameToPage(tf, top, left, bottom, right);
            currentY += height;
            var remainingHeight = bottom - currentY;
            if (remainingHeight < height * 1.1 && currentColumn < maxColumns - 1) {
                currentColumn++;
                currentY = top;
            }
        } catch (e) {}
    }
}

function placeFramesRandomly(frames, page, pageTop, pageLeft, pageBottom, pageRight) {
    var pageWidth = pageRight - pageLeft;
    var pageHeight = pageBottom - pageTop;
    for (var i = 0; i < frames.length; i++) {
        var sourceFrame = frames[i];
        try {
            var clone = sourceFrame.duplicate(page);
            clone.contents = sourceFrame.contents;
            var gb = clone.geometricBounds;
            var height = gb[2] - gb[0];
            var width = gb[3] - gb[1];
            var maxXOffset = Math.max(pageWidth - width, 0);
            var maxYOffset = Math.max(pageHeight - height, 0);
            var x = pageLeft + (maxXOffset === 0 ? 0 : Math.random() * maxXOffset);
            var y = pageTop + (maxYOffset === 0 ? 0 : Math.random() * maxYOffset);
            clone.geometricBounds = [y, x, y + height, x + width];
            clampFrameToPage(clone, pageTop, pageLeft, pageBottom, pageRight);
        } catch (e) {}
    }
}

function determineFrameSize(style, pageWidth, pageHeight) {
    var fallbackSize = 20;
    var size = style.pointSize;
    var base = (!isNaN(size) && size > 0) ? size : fallbackSize;
    var maxDimension = Math.max(Math.min(pageWidth, pageHeight) * 0.3, fallbackSize);
    return Math.min(Math.max(base * 1.8, fallbackSize), maxDimension);
}

function clampFrameToPage(frame, top, left, bottom, right) {
    try {
        var gb = frame.geometricBounds;
        var height = gb[2] - gb[0];
        var width = gb[3] - gb[1];

        var newTop = Math.min(Math.max(gb[0], top), bottom - height);
        var newLeft = Math.min(Math.max(gb[1], left), right - width);
        frame.geometricBounds = [newTop, newLeft, newTop + height, newLeft + width];
    } catch (e) {}
}

function captureStyle(frame) {
    var style = {
        appliedFont: null,
        fontStyle: null,
        pointSize: 20
    };

    if (!frame) {
        return style;
    }

    try {
        if (frame.characters.length > 0) {
            var ch = frame.characters[0];
            style.appliedFont = ch.appliedFont;
            style.fontStyle = ch.fontStyle;
            style.pointSize = ch.pointSize;
        }
    } catch (e) {}

    return style;
}

function applyStyle(frame, style) {
    try {
        if (frame.texts.length === 0) {
            return;
        }
        var text = frame.texts[0];
        if (style.appliedFont) {
            text.appliedFont = style.appliedFont;
        }
        if (style.fontStyle) {
            text.fontStyle = style.fontStyle;
        }
        if (style.pointSize) {
            text.pointSize = style.pointSize;
        }
    } catch (e) {}
}

function showDistributionDialog() {
    if (typeof Window === "undefined") {
        return {
            mode: "letters"
        };
    }

    var dialog = new Window("dialog", "Configure Distribution");
    dialog.alignChildren = "fill";

    var modePanel = dialog.add("panel", undefined, "Content Source");
    modePanel.orientation = "column";
    modePanel.alignChildren = "left";

    var lettersRadio = modePanel.add("radiobutton", undefined, "Characters");
    var wordsRadio = modePanel.add("radiobutton", undefined, "Words");
    var framesRadio = modePanel.add("radiobutton", undefined, "Text Frames");
    lettersRadio.value = true;

    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "right";
    buttonGroup.add("button", undefined, "Cancel", { name: "cancel" });
    buttonGroup.add("button", undefined, "OK", { name: "ok" });

    if (dialog.show() !== 1) {
        return null;
    }

    var mode;
    if (framesRadio.value) {
        mode = "textFrames";
    } else if (wordsRadio.value) {
        mode = "words";
    } else {
        mode = "letters";
    }

    return {
        mode: mode
    };
}