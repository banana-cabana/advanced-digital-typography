// ===============================================================
// Script name: baseline_wave_skew-right.js
// Description:
//   Applies a baseline sine wave (left -> right) and a right tilt (skew)
//   whose angle is tied to the wave magnitude (|wave|).
//   Modes: Characters / Words / Lines
//   Normalization: Global (whole scope) or Line-wise (each visual line)
// Author: GPT-5 Thinking
// ===============================================================
(function () {
    // ---------- Helpers ----------
    function N(v, d){ v = Number(v); return isFinite(v) ? v : d; }
    function clamp01(x){ return x < 0 ? 0 : (x > 1 ? 1 : x); }
    function degToRad(d){ return d * Math.PI / 180; }
    function lerp(a,b,t){ return a + (b - a) * t; }

    // ---------- Selection / Scope ----------
    if (app.selection.length === 0) { alert("Please select a text frame or a text range."); return; }
    var sel = app.selection[0], scope = null;

    if (sel.constructor && sel.constructor.name === "TextFrame") {
        if (sel.characters.length === 0) { alert("The text frame is empty."); return; }
        scope = sel.texts[0];            // operate only inside the selected frame
    } else if (sel.hasOwnProperty("characters")) {
        if (sel.characters.length === 0) { alert("The text selection is empty."); return; }
        scope = sel;                      // operate only on the selection
    } else { alert("Please select a text frame or text."); return; }

    // ---------- Dialog ----------
    var dlg = new Window("dialog", "Wave + Right Skew");
    dlg.alignChildren = "fill";

    // Apply-to (only a single dropdown now)
var pUnits = dlg.add("panel", undefined, "Apply to");
pUnits.orientation = "row"; pUnits.margins = 12;
var ddMode = pUnits.add("dropdownlist", undefined, ["Characters", "Words", "Lines"]);
ddMode.selection = 0; // Characters

    // Wave parameters (baseline shift)
    var pWave = dlg.add("panel", undefined, "Wave (Baseline Shift)");
    pWave.orientation = "row"; pWave.margins = 12;
    var gAmp = pWave.add("group"); gAmp.orientation = "column";
    gAmp.add("statictext", undefined, "Amplitude (pt)");
    var etAmp = gAmp.add("edittext", undefined, "24"); etAmp.characters = 6;

    var gCyc = pWave.add("group"); gCyc.orientation = "column";
    gCyc.add("statictext", undefined, "Cycles");
    var etCycles = gCyc.add("edittext", undefined, "1"); etCycles.characters = 6; // default 1

    var gPh  = pWave.add("group"); gPh.orientation = "column";
    gPh.add("statictext", undefined, "Phase (°)");
    var etPhase = gPh.add("edittext", undefined, "90"); etPhase.characters = 6; // 90° starts "down"

    // Skew parameters (angle tied to |wave|)
    var pSkew = dlg.add("panel", undefined, "Right Skew (linked to wave)");
    pSkew.orientation = "row"; pSkew.margins = 12;
    var gMinA = pSkew.add("group"); gMinA.orientation = "column";
    gMinA.add("statictext", undefined, "Min Angle (°)");
    var etMinA = gMinA.add("edittext", undefined, "0"); etMinA.characters = 6;

    var gMaxA = pSkew.add("group"); gMaxA.orientation = "column";
    gMaxA.add("statictext", undefined, "Max Angle (°)");
    var etMaxA = gMaxA.add("edittext", undefined, "50"); etMaxA.characters = 6;

    // Options
    var pOpts = dlg.add("panel", undefined, "Options");
    pOpts.orientation = "column"; pOpts.margins = 12;
    var cbLetters = pOpts.add("checkbox", undefined, "Letters only (Characters mode)");
    cbLetters.value = false;
    var cbReset = pOpts.add("checkbox", undefined, "Reset previous baseline shifts to 0");
    cbReset.value = true;

    // Buttons
    var gBtns = dlg.add("group"); gBtns.alignment = "right";
    gBtns.add("button", undefined, "Apply", {name:"ok"});
    gBtns.add("button", undefined, "Cancel", {name:"cancel"});
    if (dlg.show() !== 1) return;

    // ---------- Read values ----------
    var mode = ddMode.selection ? ddMode.selection.text : "Characters";
var globalNorm = true; // fixed: Global only (no line-wise UI)

    var ampPt  = N(etAmp.text, 24);              // wave amplitude in points
    var cycles = Math.max(0.0001, N(etCycles.text, 1)); // number of cycles over the mapping
    var phase  = degToRad(N(etPhase.text, 90));  // phase in radians

    var minAngle = N(etMinA.text, 0);
    var maxAngle = N(etMaxA.text, 70);
    if (!(maxAngle >= minAngle)) { alert("Max angle must be ≥ min angle."); return; }

    var onlyLetters = cbLetters.value;
    var doReset     = cbReset.value;

    // ---------- Collect units ----------
    // We gather text objects on which we'll set baselineShift and skew.
    var units = [];
    if (mode === "Characters") {
        var cs = scope.characters;
        for (var i=0; i<cs.length; i++){
            var g = cs[i], s = g.contents;
            // skip control chars
            if (s === "" || s === "\r" || s === "\n" || s === "\t") continue;
            if (onlyLetters && !/^[A-Za-zÄÖÜäöüẞß]$/.test(s)) continue;
            units.push(g);
        }
    } else if (mode === "Words") {
        var ws = scope.words;
        for (var w=0; w<ws.length; w++){
            var wd = ws[w], t = "";
            try { t = wd.contents; } catch(_){}
            if (t && t.replace(/\s+/g,"").length>0) units.push(wd);
        }
    } else { // Lines
        var ls = scope.lines;
        for (var l=0; l<ls.length; l++) units.push(ls[l]);
    }
    if (units.length === 0) { alert("No suitable units found."); return; }

    // ---------- Sample X positions to map left->right into t∈[0..1] ----------
    var X = new Array(units.length);
    var minX = Number.POSITIVE_INFINITY, maxX = Number.NEGATIVE_INFINITY;
    for (var u=0; u<units.length; u++){
        var obj = units[u], x = NaN;
        try {
            // Words/Lines: first insertionPoint; Characters: insertion before glyph
            if (obj.hasOwnProperty("insertionPoints")) x = obj.insertionPoints[0].horizontalOffset;
            else x = obj.characters[0].insertionPoints[0].horizontalOffset;
        } catch(_) { x = NaN; }
        X[u] = x;
        if (isFinite(x)) { if (x < minX) minX = x; if (x > maxX) maxX = x; }
    }

    // For line-wise normalization (Characters/Words): per-line min/max X buckets
    var lineBuckets = null;
    if (!globalNorm && mode !== "Lines") {
        lineBuckets = {}; // key: rounded baseline -> {minX, maxX}
        for (var i2=0; i2<units.length; i2++){
            var y = NaN;
            try { y = (units[i2].baseline !== undefined) ? units[i2].baseline : units[i2].characters[0].baseline; } catch(_){}
            if (!isFinite(y)) continue;
            var key = Math.round(y);
            if (!lineBuckets[key]) lineBuckets[key] = {minX: Number.POSITIVE_INFINITY, maxX: Number.NEGATIVE_INFINITY};
            var xx = X[i2];
            if (isFinite(xx)) {
                if (xx < lineBuckets[key].minX) lineBuckets[key].minX = xx;
                if (xx > lineBuckets[key].maxX) lineBuckets[key].maxX = xx;
            }
        }
    }

    function tGlobal(i){
        var span = maxX - minX;
        if (!isFinite(span) || span === 0) return 0;
        return clamp01((X[i] - minX) / span);
    }
    function tLinewise(i){
        // compute t within the visual line the unit belongs to
        var y = NaN;
        try { y = (units[i].baseline !== undefined) ? units[i].baseline : units[i].characters[0].baseline; } catch(_){}
        if (!isFinite(y)) return 0;
        var key = Math.round(y), b = lineBuckets[key];
        if (!b) return 0;
        var span = b.maxX - b.minX;
        if (!isFinite(span) || span === 0) return 0;
        return clamp01((X[i] - b.minX) / span);
    }

    // ---------- Wave & angle ----------
    function waveShiftFromT(t){
        // Positive = downwards (larger baselineShift), negative = upwards
        return ampPt * Math.sin(2 * Math.PI * (t * cycles) + phase);
    }
    // Reference amplitude for mapping |wave| to angle
    var refAmp = Math.max(1, ampPt);
    function angleFromWave(w){
        var k = clamp01(Math.abs(w) / refAmp);       // 0..1
        return Math.round( lerp(minAngle, maxAngle, k) );
    }

    // ---------- Apply ----------
    app.doScript(function(){
        // 0) Reset existing baseline shifts if requested
        if (doReset) {
            if (mode === "Lines") {
                for (var L=0; L<units.length; L++)
                    try { units[L].characters.everyItem().baselineShift = 0; } catch(_){}
            } else {
                for (var R=0; R<units.length; R++)
                    try { units[R].baselineShift = 0; } catch(_){}
            }
        }

        if (mode === "Lines") {
            // For Lines: iterate characters inside each line, so t maps left->right within the line
            for (var li=0; li<units.length; li++){
                var lineObj = units[li], chs = lineObj.characters;
                if (!chs || chs.length === 0) continue;

                // min/max X inside this visual line
                var lMin = Number.POSITIVE_INFINITY, lMax = Number.NEGATIVE_INFINITY;
                var Xs = new Array(chs.length);
                for (var c=0; c<chs.length; c++){
                    var xx = NaN; try { xx = chs[c].insertionPoints[0].horizontalOffset; } catch(_){}
                    Xs[c] = xx;
                    if (isFinite(xx)) { if (xx < lMin) lMin = xx; if (xx > lMax) lMax = xx; }
                }
                var span = lMax - lMin;
                if (!isFinite(span) || span === 0) continue;

                for (var c2=0; c2<chs.length; c2++){
                    var xx2 = Xs[c2]; if (!isFinite(xx2)) continue;
                    var t = clamp01((xx2 - lMin) / span);
                    var w = waveShiftFromT(t);
                    var ang = angleFromWave(w);

                    // Apply baseline shift
                    try { chs[c2].baselineShift = w; } catch(_){}

                    // Apply skew: try .skew; fallback to shear transform
                    var ok = false;
                    try { chs[c2].skew = ang; ok = true; } catch(_){}
                    if (!ok) {
                        try {
                            var tm = app.transformationMatrices.add({
                                horizontalScaleFactor: 1,
                                verticalScaleFactor: 1,
                                clockwiseRotationAngle: 0,
                                shearAngle: ang
                            });
                            chs[c2].transform(CoordinateSpaces.pasteboardCoordinates, AnchorPoint.centerAnchor, tm);
                        } catch(_){}
                    }
                }
            }
        } else {
            // Characters / Words: map t either globally or per line bucket
            for (var i=0; i<units.length; i++){
                var t = globalNorm ? tGlobal(i) : tLinewise(i);
                var w = waveShiftFromT(t);   // baseline shift in pt
                var ang = angleFromWave(w);  // skew angle (deg)

                // Apply baseline shift
                try { units[i].baselineShift = w; } catch(_){}

                // Apply skew (right tilt)
                var ok = false;
                try { units[i].skew = ang; ok = true; } catch(_){}
                if (!ok) {
                    // Some versions don't expose .skew on text ranges; shear-transform as fallback
                    try {
                        var tm2 = app.transformationMatrices.add({
                            horizontalScaleFactor: 1,
                            verticalScaleFactor: 1,
                            clockwiseRotationAngle: 0,
                            shearAngle: ang
                        });
                        units[i].transform(CoordinateSpaces.pasteboardCoordinates, AnchorPoint.centerAnchor, tm2);
                    } catch(_){
                        // Last resort for word-mode: try per character in the word
                        if (mode === "Words") {
                            try {
                                var chInW = units[i].characters;
                                for (var k=0; k<chInW.length; k++){
                                    try { chInW[k].skew = ang; }
                                    catch(e5){
                                        try {
                                            var tm3 = app.transformationMatrices.add({
                                                horizontalScaleFactor: 1,
                                                verticalScaleFactor: 1,
                                                clockwiseRotationAngle: 0,
                                                shearAngle: ang
                                            });
                                            chInW[k].transform(CoordinateSpaces.pasteboardCoordinates, AnchorPoint.centerAnchor, tm3);
                                        } catch(e6){}
                                    }
                                }
                            } catch(e7){}
                        }
                    }
                }
            }
        }

        try { scope.parentStory.recompose(); } catch(_){}
    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Wave + Right Skew (pure)");

    alert(
        "Done! Applied wave + linked right skew.\n" +
        "Amplitude: " + ampPt + " pt | Cycles: " + cycles + " | Phase: " + (N(etPhase.text, 90)) + "°\n" +
        "Skew: " + minAngle + "° → " + maxAngle + "°\n" +""
    );
})();
