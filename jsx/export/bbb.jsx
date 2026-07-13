#target illustrator
(function () {
    if (app.documents.length === 0) { alert("ドキュメントがありません"); return; }
    var doc = app.activeDocument;
    var TARGET = "Guides Preview for Trim View";
    function normalize(n) { return n.replace(/^\s*\*?\s*/, "").replace(/\s+$/, ""); }

    var log = [];
    for (var i = doc.layers.length - 1; i >= 0; i--) {
        var L = doc.layers[i];
        if (normalize(L.name) !== TARGET) { continue; }
        log.push("対象発見: [" + L.name + "] locked=" + L.locked + " visible=" + L.visible);
        try {
            L.locked = false;
            L.visible = true;
            L.remove();
            log.push("  → remove() 成功");
        } catch (e) {
            log.push("  → remove() 失敗: " + e.message + " (line " + e.line + ")");
        }
    }
    if (log.length === 0) { log.push("対象が見つかりませんでした"); }
    alert(log.join("\n"));
})();
