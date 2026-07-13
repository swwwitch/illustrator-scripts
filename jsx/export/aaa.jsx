#target illustrator
(function () {
    if (app.documents.length === 0) { alert("ドキュメントがありません"); return; }
    var lines = [];
    function walk(layers, depth) {
        for (var i = 0; i < layers.length; i++) {
            var L = layers[i];
            var indent = "";
            for (var d = 0; d < depth; d++) { indent += "  "; }
            lines.push(indent + "[" + L.name + "]  locked=" + L.locked + " visible=" + L.visible);
            if (L.layers && L.layers.length > 0) { walk(L.layers, depth + 1); }
        }
    }
    walk(app.activeDocument.layers, 0);
    alert(lines.join("\n"));
})();
