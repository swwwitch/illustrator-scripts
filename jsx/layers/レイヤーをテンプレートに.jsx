function setLayerTemplate(doc, name, on) {
    for (var i = 0; i < doc.layers.length; i++) {
        if (doc.layers[i].name === name) {
            doc.layers[i].template = on;
            return true;
        }
    }
    return false;
}

setLayerTemplate(app.activeDocument, "下絵", true);  // テンプレート化
