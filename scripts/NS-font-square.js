/*
Arranges words in a square grid of text frames on the active page.
Uses the current text selection if available, otherwise prompts for input.
*/

(function () {
	if (app.documents.length === 0) {
		alert("Bitte öffne ein Dokument, bevor du das Skript ausführst.");
		return;
	}

	app.doScript(main, ScriptLanguage.JAVASCRIPT, null, UndoModes.ENTIRE_SCRIPT, "Text im Quadrat anordnen");

	function main() {
		var doc = app.activeDocument;
		var sourceContext = collectSourceText();
		var defaultText = String(sourceContext.text || "");
		var config = showConfigurationDialog(defaultText, !!sourceContext.bounds);
		if (!config) {
			return;
		}

		var inputText = String(config.inputText || "").replace(/^\s+|\s+$/g, "");
		if (!inputText) {
			alert("Bitte gib einen Inhalt für das Quadrat ein.");
			return;
		}

		var items = config.contentMode === "letters" ? splitIntoLetters(inputText) : tokenize(inputText);
		if (items.length === 0) {
			alert("Es wurden keine Inhalte gefunden.");
			return;
		}

		var page = resolvePage(doc);
		if (!page) {
			alert("Konnte keine Seite ermitteln.");
			return;
		}

		var styleSample = captureBaseStyle(doc);
		var cellsPerSide = Math.max(1, Math.ceil(Math.sqrt(items.length)));
		var totalCells = cellsPerSide * cellsPerSide;
		if (items.length < totalCells) {
			for (var padIndex = items.length; padIndex < totalCells; padIndex++) {
				items.push(" ");
			}
		}
		if (items.length > totalCells) {
			items = items.slice(0, totalCells);
		}
		if (config.randomize) {
			items = shuffleTokens(items);
		}

		var bounds;
		if (config.layoutMode === "selection") {
			var selectionBounds = sourceContext.bounds || findSelectionBounds();
			if (!selectionBounds) {
				alert("Kein geeigneter Textrahmen für den Bereich gefunden.");
				return;
			}
			bounds = [selectionBounds[0], selectionBounds[1], selectionBounds[2], selectionBounds[3]];
		} else {
			var pageBounds = cloneBounds(page.bounds);
			if (!pageBounds) {
				alert("Could not determine the page bounds.");
				return;
			}
			bounds = pageBounds;
		}

		var areaWidth = bounds[3] - bounds[1];
		var areaHeight = bounds[2] - bounds[0];
		var gap = Math.max(Math.min(Math.min(areaWidth, areaHeight) * 0.02, 18), 4);
		var totalGapWidth = gap * (cellsPerSide - 1);
		var totalGapHeight = gap * (cellsPerSide - 1);
		var cellWidth = (areaWidth - totalGapWidth);
		var cellHeight = (areaHeight - totalGapHeight);
		if (cellWidth <= 0 || cellHeight <= 0) {
			gap = 0;
			totalGapWidth = 0;
			totalGapHeight = 0;
			cellWidth = areaWidth;
			cellHeight = areaHeight;
		}
		cellWidth = cellWidth / cellsPerSide;
		cellHeight = cellHeight / cellsPerSide;
		var layoutWidth = (cellWidth * cellsPerSide) + totalGapWidth;
		var layoutHeight = (cellHeight * cellsPerSide) + totalGapHeight;
		var startX = bounds[1] + (areaWidth - layoutWidth) / 2;
		var startY = bounds[0] + (areaHeight - layoutHeight) / 2;

		var framesCreated = 0;
		var createdFrames = [];
		var frameStyles = [];

		for (var index = 0; index < items.length; index++) {
			var row = Math.floor(index / cellsPerSide);
			var column = index % cellsPerSide;

			var top = startY + row * (cellHeight + gap);
			var left = startX + column * (cellWidth + gap);
			var bottom = top + cellHeight;
			var right = left + cellWidth;

			try {
				var frame = page.textFrames.add();
				frame.geometricBounds = [top, left, bottom, right];
				frame.contents = items[index];
				styleFrame(frame, styleSample);
				framesCreated++;
				createdFrames.push(frame);
				frameStyles.push(captureFrameStyle(frame));
			} catch (error) {
				// Skip frames that cannot be placed, continue with next word.
			}
		}

		if (config.layoutMode === "selection" && createdFrames.length > 0) {
			var overflowDetected = false;
			for (var frameIndex = 0; frameIndex < createdFrames.length; frameIndex++) {
				if (createdFrames[frameIndex].overflows) {
					overflowDetected = true;
					break;
				}
			}
			if (overflowDetected) {
				var requiredSize = Math.max(cellWidth, cellHeight);
				for (var measureIndex = 0; measureIndex < createdFrames.length; measureIndex++) {
					try {
						var sizeInfo = measureContentSize(createdFrames[measureIndex]);
						if (!isNaN(sizeInfo.width) && sizeInfo.width > requiredSize) {
							requiredSize = sizeInfo.width;
						}
						if (!isNaN(sizeInfo.height) && sizeInfo.height > requiredSize) {
							requiredSize = sizeInfo.height;
						}
					} catch (error) {}
				}
				var resizedGap = gap;
				var resizedTotalGapWidth = totalGapWidth;
				var resizedTotalGapHeight = totalGapHeight;
				var needsGapReduction = (requiredSize * cellsPerSide) + resizedTotalGapWidth > areaWidth ||
					(requiredSize * cellsPerSide) + resizedTotalGapHeight > areaHeight;
				if (needsGapReduction) {
					resizedGap = 0;
					resizedTotalGapWidth = 0;
					resizedTotalGapHeight = 0;
				}
				var availableSquare = Math.min(
					(areaWidth - resizedTotalGapWidth) / cellsPerSide,
					(areaHeight - resizedTotalGapHeight) / cellsPerSide
				);
				if (availableSquare > 0) {
					var finalSize = Math.min(requiredSize, availableSquare);
					var adjustedLayoutWidth = (finalSize * cellsPerSide) + resizedTotalGapWidth;
					var adjustedLayoutHeight = (finalSize * cellsPerSide) + resizedTotalGapHeight;
					var adjustedStartX = bounds[1] + (areaWidth - adjustedLayoutWidth) / 2;
					var adjustedStartY = bounds[0] + (areaHeight - adjustedLayoutHeight) / 2;
					for (var resizeIndex = 0; resizeIndex < createdFrames.length; resizeIndex++) {
						var resizeFrame = createdFrames[resizeIndex];
						var resizeRow = Math.floor(resizeIndex / cellsPerSide);
						var resizeColumn = resizeIndex % cellsPerSide;
						var newTop = adjustedStartY + resizeRow * (finalSize + resizedGap);
						var newLeft = adjustedStartX + resizeColumn * (finalSize + resizedGap);
						var newBottom = newTop + finalSize;
						var newRight = newLeft + finalSize;
						resizeFrame.geometricBounds = [newTop, newLeft, newBottom, newRight];
						applyCapturedStyle(resizeFrame, frameStyles[resizeIndex]);
					}
				}
				var overflowStillPresent = false;
				for (var verifyIndex = 0; verifyIndex < createdFrames.length; verifyIndex++) {
					if (createdFrames[verifyIndex].overflows) {
						overflowStillPresent = true;
						break;
					}
				}
				if (overflowStillPresent) {
					alert("The selected text frame does not provide enough space. Please enlarge the area.");
				}
			}
		}

		alert(framesCreated + " elements were arranged in the square.");
	}

	function cloneBounds(bounds) {
		try {
			if (!bounds || bounds.length < 4) {
				return null;
			}
			return [bounds[0], bounds[1], bounds[2], bounds[3]];
		} catch (error) {
			return null;
		}
	}

	function collectSourceText() {
		var text = "";
		var bounds = null;
		if (app.selection.length > 0) {
			var item = app.selection[0];
			if (isTextTarget(item)) {
				text = item.contents;
			}
			if (!text && item.hasOwnProperty("texts") && item.texts.length > 0) {
				text = item.texts[0].contents;
			}
			if (item instanceof TextFrame) {
				var frameBounds = cloneBounds(item.geometricBounds);
				if (frameBounds) {
					bounds = frameBounds;
				}
			} else if (item.hasOwnProperty("parentTextFrames") && item.parentTextFrames.length > 0) {
				var parentBounds = cloneBounds(item.parentTextFrames[0].geometricBounds);
				if (parentBounds) {
					bounds = parentBounds;
				}
			}
		}
		return {
			text: text,
			bounds: bounds
		};
	}

	function tokenize(text) {
		var matches = text.match(/\S+/g);
		return matches ? matches : [];
	}

	function splitIntoLetters(text) {
		var cleaned = String(text || "").replace(/\s+/g, "");
		var chars = [];
		for (var i = 0; i < cleaned.length; i++) {
			var ch = cleaned.charAt(i);
			if (ch) {
				chars.push(ch);
			}
		}
		return chars;
	}

	function showConfigurationDialog(defaultText, hasSelectionBounds) {
		var dialog = new Window("dialog", "Arrange Text in a Square");
		dialog.alignChildren = "fill";

		var areaPanel = dialog.add("panel", undefined, "Target Area");
		areaPanel.alignChildren = "left";
		var entirePageRadio = areaPanel.add("radiobutton", undefined, "Use entire page bounds");
		var selectionRadio = areaPanel.add("radiobutton", undefined, "Use current text frame");
		entirePageRadio.value = true;
		selectionRadio.enabled = !!hasSelectionBounds;
		if (!selectionRadio.enabled) {
			selectionRadio.value = false;
		}

		var infoPanel = dialog.add("panel", undefined, "Layout");
		infoPanel.alignChildren = "left";
		infoPanel.add("statictext", undefined, "The square occupies the selected area without overlap.");

		var contentPanel = dialog.add("panel", undefined, "Square content");
		contentPanel.alignChildren = "left";
		var wordsRadio = contentPanel.add("radiobutton", undefined, "Words");
		var lettersRadio = contentPanel.add("radiobutton", undefined, "Letters");
		wordsRadio.value = true;

		var orderPanel = dialog.add("panel", undefined, "Order");
		orderPanel.alignChildren = "left";
		var keepOrderRadio = orderPanel.add("radiobutton", undefined, "Keep original order");
		var randomOrderRadio = orderPanel.add("radiobutton", undefined, "Shuffle order");
		keepOrderRadio.value = true;

		var placeholderWords = String(defaultText || "").replace(/^\s+|\s+$/g, "");
		if (!placeholderWords) {
			placeholderWords = "Word Word Word";
		}
		var placeholderLetters = String(defaultText || "").replace(/\s+/g, "");
		if (!placeholderLetters) {
			placeholderLetters = "ABCDE";
		}

		var wordsGroup = contentPanel.add("group");
		wordsGroup.add("statictext", undefined, "Words:");
		var wordsInput = wordsGroup.add("edittext", undefined, placeholderWords);
		wordsInput.characters = 30;

		var lettersGroup = contentPanel.add("group");
		lettersGroup.add("statictext", undefined, "Letters:");
		var lettersInput = lettersGroup.add("edittext", undefined, placeholderLetters);
		lettersInput.characters = 30;

		function updateInputs() {
			wordsInput.enabled = wordsRadio.value;
			lettersInput.enabled = lettersRadio.value;
		}

		wordsRadio.onClick = updateInputs;
		lettersRadio.onClick = updateInputs;
		updateInputs();

		var buttonGroup = dialog.add("group");
		buttonGroup.alignment = "right";
		buttonGroup.add("button", undefined, "Cancel", { name: "cancel" });
		buttonGroup.add("button", undefined, "OK", { name: "ok" });

		if (dialog.show() !== 1) {
			return null;
		}

		var layoutMode = selectionRadio.value ? "selection" : "page";
		var contentMode = wordsRadio.value ? "words" : "letters";
		var inputText = wordsRadio.value ? wordsInput.text : lettersInput.text;
		return {
			layoutMode: layoutMode,
			contentMode: contentMode,
			inputText: inputText,
			randomize: randomOrderRadio.value
		};
	}

	function resolvePage(activeDoc) {
		if (app.selection.length > 0) {
			var item = app.selection[0];
			if (item.hasOwnProperty("parentPage") && item.parentPage instanceof Page) {
				return item.parentPage;
			}
			if (item instanceof TextFrame && item.parentPage instanceof Page) {
				return item.parentPage;
			}
			if (item.parentStory && item.parentStory.parentPage instanceof Page) {
				return item.parentStory.parentPage;
			}
		}
		if (activeDoc.layoutWindows.length > 0) {
			return activeDoc.layoutWindows[0].activePage;
		}
		if (activeDoc.pages.length > 0) {
			return activeDoc.pages[0];
		}
		return null;
	}

	function isTextTarget(item) {
		return item instanceof Text || item instanceof Word ||
			item instanceof InsertionPoint || item instanceof Character ||
			item instanceof Paragraph || item instanceof Story;
	}

	function findSelectionBounds() {
		if (app.selection.length === 0) {
			return null;
		}

		for (var i = 0; i < app.selection.length; i++) {
			var item = app.selection[i];
			if (item instanceof TextFrame) {
				return item.geometricBounds;
			}
			if (item.hasOwnProperty("parentTextFrames") && item.parentTextFrames.length > 0) {
				return item.parentTextFrames[0].geometricBounds;
			}
			if (item.hasOwnProperty("parent") && item.parent instanceof TextFrame) {
				return item.parent.geometricBounds;
			}
		}

		return null;
	}

	function determineCellSize(baseStyle) {
		var minimumSize = 36;
		if (baseStyle && baseStyle.pointSize) {
			var size = parseFloat(baseStyle.pointSize);
			if (!isNaN(size) && size > 0) {
				return Math.max(size * 1.6, minimumSize);
			}
		}
		return minimumSize;
	}

	function styleFrame(textFrame, baseStyle) {
		try {
			var paragraphs = textFrame.paragraphs;
			paragraphs.everyItem().justification = Justification.centerAlign;
			if (baseStyle && textFrame.texts.length > 0) {
				var textRange = textFrame.texts[0];
				if (baseStyle.appliedFont) {
					textRange.appliedFont = baseStyle.appliedFont;
				}
				if (baseStyle.fontStyle) {
					textRange.fontStyle = baseStyle.fontStyle;
				}
				if (!isNaN(parseFloat(baseStyle.pointSize))) {
					textRange.pointSize = baseStyle.pointSize;
				}
				if (!isNaN(parseFloat(baseStyle.leading))) {
					textRange.leading = baseStyle.leading;
				}
			}
		} catch (error) {}
	}

	function captureFrameStyle(textFrame) {
		if (!textFrame || !textFrame.isValid || textFrame.texts.length === 0) {
			return null;
		}
		var textRange = textFrame.texts[0];
		if (!textRange || textRange.characters.length === 0) {
			return null;
		}
		return readCharacterStyle(textRange.characters[0]);
	}

	function applyCapturedStyle(textFrame, style) {
		if (!textFrame || !textFrame.isValid || !style || textFrame.texts.length === 0) {
			return;
		}
		var textRange = textFrame.texts[0];
		try {
			if (style.appliedFont) {
				textRange.appliedFont = style.appliedFont;
			}
			if (style.fontStyle) {
				textRange.fontStyle = style.fontStyle;
			}
			if (!isNaN(parseFloat(style.pointSize))) {
				textRange.pointSize = style.pointSize;
			}
			if (!isNaN(parseFloat(style.leading))) {
				textRange.leading = style.leading;
			}
		} catch (error) {}
	}

	function shuffleTokens(tokens) {
		var meaningful = [];
		var placeholders = [];
		for (var i = 0; i < tokens.length; i++) {
			var token = tokens[i];
			if (token && String(token).replace(/\s+/g, "") !== "") {
				meaningful.push(token);
			} else {
				placeholders.push(token);
			}
		}
		shuffleArray(meaningful);
		return meaningful.concat(placeholders);
	}

	function shuffleArray(array) {
		for (var i = array.length - 1; i > 0; i--) {
			var j = Math.floor(Math.random() * (i + 1));
			var temp = array[i];
			array[i] = array[j];
			array[j] = temp;
		}
	}

	function measureContentSize(textFrame) {
		var size = { width: 0, height: 0 };
		if (!textFrame || !textFrame.isValid) {
			return size;
		}
		var tempFrame = null;
		try {
			tempFrame = textFrame.duplicate();
			tempFrame.textFramePreferences.autoSizingReferencePoint = AutoSizingReferenceEnum.TOP_LEFT_POINT;
			tempFrame.textFramePreferences.autoSizingType = AutoSizingTypeEnum.HEIGHT_AND_WIDTH;
			tempFrame.fit(FitOptions.frameToContent);
			var gb = tempFrame.geometricBounds;
			size.width = gb[3] - gb[1];
			size.height = gb[2] - gb[0];
		} catch (error) {
			// Ignore measurement errors.
		} finally {
			if (tempFrame && tempFrame.isValid) {
				tempFrame.remove();
			}
		}
		return size;
	}

	function captureBaseStyle(activeDoc) {
		var sample = null;
		if (app.selection.length > 0) {
			var item = app.selection[0];
			var range = extractTextRange(item);
			if (range && range.characters.length > 0) {
				sample = readCharacterStyle(range.characters[0]);
			}
		}
		if (!sample) {
			sample = readDefaults(activeDoc);
		}
		return sample;
	}

	function extractTextRange(item) {
		if (!item) {
			return null;
		}
		if (item instanceof Text || item instanceof Word || item instanceof Paragraph || item instanceof Story) {
			return item;
		}
		if (item instanceof InsertionPoint) {
			return item.parent;
		}
		if (item instanceof Character) {
			return item.parent;
		}
		if (item instanceof TextFrame && item.texts.length > 0) {
			return item.texts[0];
		}
		if (item.hasOwnProperty("texts") && item.texts.length > 0) {
			return item.texts[0];
		}
		return null;
	}

	function readCharacterStyle(character) {
		var style = {};
		try {
			style.appliedFont = character.appliedFont;
			style.fontStyle = character.fontStyle;
			style.pointSize = character.pointSize;
			style.leading = character.leading;
		} catch (error) {}
		return style;
	}

	function readDefaults(activeDoc) {
		var defaults = activeDoc.textDefaults;
		var sample = {};
		try {
			sample.appliedFont = defaults.appliedFont;
			sample.fontStyle = defaults.fontStyle;
			sample.pointSize = defaults.pointSize;
			sample.leading = defaults.leading;
		} catch (error) {}
		return sample;
	}

	function clampToBounds(bounds, pageBounds) {
		var top = Math.max(bounds[0], pageBounds[0]);
		var left = Math.max(bounds[1], pageBounds[1]);
		var bottom = Math.min(bounds[2], pageBounds[2]);
		var right = Math.min(bounds[3], pageBounds[3]);

		if (bottom <= top) {
			bottom = top + 1;
			if (bottom > pageBounds[2]) {
				top = pageBounds[2] - 1;
				bottom = pageBounds[2];
			}
		}

		if (right <= left) {
			right = left + 1;
			if (right > pageBounds[3]) {
				left = pageBounds[3] - 1;
				right = pageBounds[3];
			}
		}

		return [top, left, bottom, right];
	}
})();

