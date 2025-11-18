(function(){
    if(app.documents.length===0 || !(app.selection[0] instanceof TextFrame)){
        alert("Please select a text frame."); 
        return;
    }
    var s = app.selection[0],
        d = new Window("dialog", "Wave Effect"),
        g1 = d.add("group"),
        g2 = d.add("group");

    g1.add("statictext", undefined, "Amplitude:");
    a = g1.add("edittext", undefined, "5");
    a.characters = 5;

    g2.add("statictext", undefined, "Wavelength:");
    w = g2.add("edittext", undefined, "10");
    w.characters = 5;

    b = d.add("group");
    b.add("button", undefined, "OK", {name: "ok"});
    b.add("button", undefined, "Cancel", {name: "cancel"});

    if(d.show() != 1) return;

    var A = parseFloat(a.text),
        W = parseFloat(w.text);

    if(isNaN(A) || isNaN(W) || W <= 0){
        alert("Invalid values.");
        return;
    }

    var t = s.parentStory,
        c = t.characters.length,
        p = new Window("palette", "Progress");

    p.alignChildren = "fill";
    txt = p.add("statictext", undefined, "Applying effect...");
    bar = p.add("progressbar", undefined, 0, c);
    p.show();

    for(var i = 0; i < c; i++){
        try {
            t.characters[i].baselineShift = Math.sin((i / W) * 2 * Math.PI) * A;
        } catch(e) {}

        if(i % 20 == 0){
            bar.value = i;
            p.update();
        }
    }

    bar.value = c;
    txt.text = "Done!";
    $.sleep(200);
    p.close();
    alert("Done!");
})();