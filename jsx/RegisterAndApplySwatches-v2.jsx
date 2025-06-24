#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

var swatchGroup = null;

/*
  スクリプト名：RegisterAndApplySwatches.jsx
  概要：
    選択されたオブジェクト（長方形やテキスト）の〈塗り〉カラー、
    およびパスオブジェクトの〈線〉カラーをスウォッチ（スポットカラー）として登録し、再適用します。

  処理の流れ：
    1. 閉じたPathItem または TextFrame を対象に走査
    2. RGBまたはCMYKのカラーを検出（塗り／線）
    3. 各カラーごとにスウォッチ名を生成（例：C=0 M=100 Y=100 K=0）
    4. 同名のスポットが存在すれば再利用、なければ新規作成
    5. SpotColor を生成して、該当オブジェクトに適用

  対象：
    ・PathItem（閉じたパス）
    ・TextFrame（テキスト）

  限定条件：
    ・RGB または CMYK カラーのみ対応
    ・スポット／グレースケールは対象外

  更新日：2025-06-26
*/

function createColorCopy(color) {
    var c = null;
    switch (color.typename) {
        case "RGBColor":
            c = new RGBColor();
            c.red = color.red;
            c.green = color.green;
            c.blue = color.blue;
            break;
        case "CMYKColor":
            c = new CMYKColor();
            c.cyan = color.cyan;
            c.magenta = color.magenta;
            c.yellow = color.yellow;
            c.black = color.black;
            break;
    }
    return c;
}

function registerColorToSwatch(targetItem) {
    function registerSingleColorToSwatch(item, color, isStroke) {
        if (!color || color.typename === "NoColor" || color.typename === "SpotColor" || color.typename === "GrayColor" || color.typename === "GradientColor" || color.typename === "PatternColor") {
            return false; // スルー
        }

        var swatchName = "";
        if (color.typename === "RGBColor") {
            swatchName = "R=" + Math.round(color.red) + " G=" + Math.round(color.green) + " B=" + Math.round(color.blue);
        } else if (color.typename === "CMYKColor") {
            swatchName =
                "C=" + Math.round(color.cyan) +
                " M=" + Math.round(color.magenta) +
                " Y=" + Math.round(color.yellow) +
                " K=" + Math.round(color.black);
        } else {
            // alert("対応していないカラーモデルです。RGBまたはCMYKのみ対応しています。");
            return false;
        }

        var spot = null;
        for (var i = 0; i < app.activeDocument.spots.length; i++) {
            if (app.activeDocument.spots[i].name === swatchName) {
                spot = app.activeDocument.spots[i];
                break;
            }
        }

        if (!spot) {
            var dupColor = createColorCopy(color);
            if (dupColor === null) {
                // alert("対応していないカラーモデルです。RGBまたはCMYKのみ対応しています。");
                return false;
            }

            spot = app.activeDocument.spots.add();
            if (swatchGroup) {
                swatchGroup.addSpot(spot);
            }
            spot.colorType = ColorModel.SPOT;
            spot.color = dupColor;
            spot.name = swatchName;
        }

        var spotColor = new SpotColor();
        spotColor.spot = spot;
        spotColor.tint = 100;

        if (isStroke) {
            item.strokeColor = spotColor;
        } else {
            if (item.typename === "TextFrame") {
                item.textRange.characterAttributes.fillColor = spotColor;
            } else {
                item.fillColor = spotColor;
            }
        }

        return true;
    }

    if (targetItem.typename === "PathItem") {
        var fillColor = targetItem.fillColor;
        var strokeColor = targetItem.strokeColor;

        var fillResult = false;
        var strokeResult = false;

        if (fillColor && fillColor.typename !== "NoColor") {
            fillResult = registerSingleColorToSwatch(targetItem, fillColor, false);
        }

        if (strokeColor && strokeColor.typename !== "NoColor") {
            strokeResult = registerSingleColorToSwatch(targetItem, strokeColor, true);
        }

        if (!fillResult && !strokeResult) {
            return false; // 警告なしでスキップ
        }

        return fillResult || strokeResult;

    } else if (targetItem.typename === "TextFrame") {
        var fillColor = targetItem.textRange.characterAttributes.fillColor;
        if (!fillColor || fillColor.typename === "NoColor") {
            alert("塗りが設定されていません。");
            return false;
        }
        return registerSingleColorToSwatch(targetItem, fillColor, false);
    } else {
        return false;
    }
}

function processItem(item) {
    if ((item.typename === "PathItem" && item.closed) || item.typename === "TextFrame") {
        registerColorToSwatch(item);
    } else if (item.typename === "GroupItem") {
        for (var i = 0; i < item.pageItems.length; i++) {
            processItem(item.pageItems[i]);
        }
    } else if (item.typename === "CompoundPathItem") {
        if (item.pathItems.length > 0) {
            registerColorToSwatch(item);
        }
    }
}

function main() {
    if (app.documents.length === 0 || app.activeDocument.selection.length === 0) {
        alert("オブジェクトを選択してください。");
        return;
    }

    var swatchGroupName = "登録カラー";
    var doc = app.activeDocument;
    swatchGroup = null;

    // スウォッチグループが存在しない場合は作成
    for (var g = 0; g < doc.swatchGroups.length; g++) {
        if (doc.swatchGroups[g].name === swatchGroupName) {
            swatchGroup = doc.swatchGroups[g];
            break;
        }
    }

    if (!swatchGroup) {
        swatchGroup = doc.swatchGroups.add();
        swatchGroup.name = swatchGroupName;
    }

    var sel = app.activeDocument.selection;

    try {
        for (var i = 0; i < sel.length; i++) {
            processItem(sel[i]);
        }
    } catch (e) {
        alert("カラーの適用中にエラーが発生しました:\n" + e);
    }
}

main();