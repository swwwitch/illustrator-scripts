#target illustrator

function main() {
    if (app.documents.length === 0) {
        alert("ドキュメントが開かれていません。");
        return;
    }

    var doc = app.activeDocument;
    var sel = doc.selection;

    if (sel.length === 0) {
        alert("オブジェクトを選択してから実行してください。");
        return;
    }

    var count = 0;

    // 選択された各オブジェクトを処理
    for (var i = 0; i < sel.length; i++) {
        count += processItem(sel[i], 1.0);
    }

    if (count > 0) {
        alert(count + " 個の要素を不透明度100%の色（RGB/CMYK/グレースケール対応）に変換しました。");
    } else {
        alert("処理対象（不透明度100%未満の単色オブジェクト）が見つかりませんでした。");
    }
}

// オブジェクトの種類に応じて再帰的に処理する関数
function processItem(obj, parentRatio) {
    var count = 0;
    var currentRatio = parentRatio;

    // オブジェクト自体に不透明度がある場合、割合を掛け合わせる
    if (obj.opacity !== undefined && obj.opacity < 100) {
        currentRatio = currentRatio * (obj.opacity / 100);
        obj.opacity = 100; // 不透明度を100%にリセット
    }

    // --- パスアイテムの場合 ---
    if (obj.typename === "PathItem") {
        if (currentRatio < 1.0) {
            if (obj.filled) {
                obj.fillColor = calculateNewColor(obj.fillColor, currentRatio);
            }
            if (obj.stroked) {
                obj.strokeColor = calculateNewColor(obj.strokeColor, currentRatio);
            }
            count++;
        }
    } 
    // --- 複合パスの場合 ---
    else if (obj.typename === "CompoundPathItem") {
        for (var i = 0; i < obj.pathItems.length; i++) {
            count += processItem(obj.pathItems[i], currentRatio);
        }
    } 
    // --- グループの場合 ---
    else if (obj.typename === "GroupItem") {
        for (var i = 0; i < obj.pageItems.length; i++) {
            count += processItem(obj.pageItems[i], currentRatio);
        }
    } 
    // --- テキストフレームの場合 ---
    else if (obj.typename === "TextFrame") {
        if (currentRatio < 1.0) {
            var textRange = obj.textRange;
            var charAttr = textRange.characterAttributes;
            
            if (charAttr.fillColor) {
                charAttr.fillColor = calculateNewColor(charAttr.fillColor, currentRatio);
            }
            if (charAttr.strokeColor) {
                charAttr.strokeColor = calculateNewColor(charAttr.strokeColor, currentRatio);
            }
            count++;
        }
    }

    return count;
}

// カラーモードを判別して新しい色（白背景との合成結果）を返す関数
function calculateNewColor(color, ratio) {
    if (color.typename === "CMYKColor") {
        // CMYKの場合（0が白）
        var newColor = new CMYKColor();
        newColor.cyan    = Math.min(100, color.cyan * ratio);
        newColor.magenta = Math.min(100, color.magenta * ratio);
        newColor.yellow  = Math.min(100, color.yellow * ratio);
        newColor.black   = Math.min(100, color.black * ratio);
        return newColor;
        
    } else if (color.typename === "RGBColor") {
        // RGBの場合（255が白）
        var newColor = new RGBColor();
        var invRatio = 1.0 - ratio;
        newColor.red   = Math.round(color.red * ratio + 255 * invRatio);
        newColor.green = Math.round(color.green * ratio + 255 * invRatio);
        newColor.blue  = Math.round(color.blue * ratio + 255 * invRatio);
        // 上下限のクリッピング
        newColor.red   = Math.max(0, Math.min(255, newColor.red));
        newColor.green = Math.max(0, Math.min(255, newColor.green));
        newColor.blue  = Math.max(0, Math.min(255, newColor.blue));
        return newColor;
        
    } else if (color.typename === "GrayColor") {
        // グレースケールの場合（Kと同じく0が白）
        var newColor = new GrayColor();
        newColor.gray = Math.min(100, color.gray * ratio);
        return newColor;
        
    }
    
    // グラデーションや特色(SpotColor)などは計算できないためそのまま返す
    return color;
}

main();
