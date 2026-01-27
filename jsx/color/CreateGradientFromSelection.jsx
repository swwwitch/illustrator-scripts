#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
  CreateGradientFromSelection.jsx

  選択しているオブジェクトの塗り色（フィル）を、
  配置順（左→右、上→下）で抽出してスウォッチグループに登録し、
  その順序をもとにグラデーションを自動生成します。

  ・グループ／複合パス／テキストは再帰的に処理
  ・線色（ストローク）も対象
  ・抽出した色はすべてグローバルカラー（プロセス）に変換してスウォッチ登録
  ・生成するグラデーションの各ストップには登録済みのグローバルカラーを設定
  ・作成したグラデーションはスウォッチに追加
  ・ビュー中央に長方形を作成し、生成したグラデーションを適用
  ・ドキュメントが無い／選択が無い／色が1色以下の場合やエラー発生時は無言で終了

  Version: v1.2
  更新日: 2026-01-27
*/

function getCurrentLang() {
  return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

function main() {
    // ドキュメントが開かれているか確認
    if (app.documents.length === 0) {
        return;
    }

    var doc = app.activeDocument;

    // 選択があるか確認
    if (!doc.selection || doc.selection.length === 0) {
        return;
    }

    /* =========================================
     * Color utilities / カラー取得ユーティリティ
     * ========================================= */

    function isNoColor(c) {
        // NoColor は typename が "NoColor" になる
        try {
            return (c == null) || (c.typename === "NoColor");
        } catch (e) {
            return true;
        }
    }

    function colorKey(c) {
        // 重複除去用の簡易キー
        if (!c) return "null";
        var t = c.typename;
        try {
            if (t === "RGBColor") {
                return "RGB:" + [c.red, c.green, c.blue].join(",");
            }
            if (t === "CMYKColor") {
                return "CMYK:" + [c.cyan, c.magenta, c.yellow, c.black].join(",");
            }
            if (t === "GrayColor") {
                return "Gray:" + c.gray;
            }
            if (t === "SpotColor") {
                // スポットはスポット名＋濃度
                var spotName = (c.spot && c.spot.name) ? c.spot.name : "(spot)";
                return "Spot:" + spotName + ":" + c.tint;
            }
            if (t === "PatternColor") {
                var patName = (c.pattern && c.pattern.name) ? c.pattern.name : "(pattern)";
                return "Pattern:" + patName;
            }
            if (t === "GradientColor") {
                var gName = (c.gradient && c.gradient.name) ? c.gradient.name : "(gradient)";
                return "Gradient:" + gName;
            }
        } catch (e) {}
        return "Other:" + t;
    }

    function pushUniqueColor(list, seenMap, c) {
        if (isNoColor(c)) return;
        var k = colorKey(c);
        if (seenMap[k]) return;
        seenMap[k] = true;
        list.push(c);
    }

    // 位置情報（左上）を取得 / Get top-left position
    function getItemTopLeft(item) {
        // geometricBounds: [left, top, right, bottom]
        try {
            var b = item.geometricBounds;
            return { left: b[0], top: b[1] };
        } catch (e) {
            return { left: 0, top: 0 };
        }
    }

    // 色＋位置のエントリを収集 / Collect color entries with position
    function collectFillColorEntries(item, outEntries) {
        if (!item) return;

        try {
            // グループなどは再帰
            if (item.typename === "GroupItem") {
                for (var i = 0; i < item.pageItems.length; i++) {
                    collectFillColorEntries(item.pageItems[i], outEntries);
                }
                return;
            }

            // compoundPath は pathItems を辿る
            if (item.typename === "CompoundPathItem") {
                for (var j = 0; j < item.pathItems.length; j++) {
                    collectFillColorEntries(item.pathItems[j], outEntries);
                }
                return;
            }

            // テキスト
            if (item.typename === "TextFrame") {
                var pT = getItemTopLeft(item);

                // Fill
                var tf = item.textRange.characterAttributes.fillColor;
                if (!isNoColor(tf)) {
                    outEntries.push({ left: pT.left, top: pT.top, color: tf });
                }

                // Stroke
                try {
                    var ts = item.textRange.characterAttributes.strokeColor;
                    if (!isNoColor(ts)) {
                        outEntries.push({ left: pT.left, top: pT.top, color: ts });
                    }
                } catch (eTS) {}

                return;
            }

            // PathItem など
            if (typeof item.filled !== "undefined" || typeof item.stroked !== "undefined") {
                var p = getItemTopLeft(item);

                // Fill
                if (typeof item.filled !== "undefined" && item.filled) {
                    var fc = item.fillColor;
                    if (!isNoColor(fc)) {
                        outEntries.push({ left: p.left, top: p.top, color: fc });
                    }
                }

                // Stroke
                if (typeof item.stroked !== "undefined" && item.stroked) {
                    var sc = item.strokeColor;
                    if (!isNoColor(sc)) {
                        outEntries.push({ left: p.left, top: p.top, color: sc });
                    }
                }

                return;
            }
        } catch (e) {
            // 取得できないアイテムは無視
        }
    }

    // 選択から色を「左→右、上→下」順で収集（重複は除外）
    function collectColorsFromSelection(selection) {
        var entries = [];
        for (var i = 0; i < selection.length; i++) {
            collectFillColorEntries(selection[i], entries);
        }

        // 左→右（left 昇順）、上→下（top 降順）でソート
        entries.sort(function(a, b) {
            if (a.left < b.left) return -1;
            if (a.left > b.left) return 1;
            // top は上ほど値が大きい（座標系の都合）ため降順
            if (a.top > b.top) return -1;
            if (a.top < b.top) return 1;
            return 0;
        });

        // 重複除外（同一色は最初の1つだけ）
        var colors = [];
        var seen = {};
        for (var k = 0; k < entries.length; k++) {
            var c = entries[k].color;
            if (isNoColor(c)) continue;
            var key = colorKey(c);
            if (seen[key]) continue;
            seen[key] = true;
            colors.push(c);
        }

        return colors;
    }

    function uniqueName(baseName, existsFunc) {
        var name = baseName;
        var n = 1;
        while (true) {
            try {
                if (existsFunc(name)) {
                    name = baseName + " " + n;
                    n++;
                    continue;
                }
                break;
            } catch (e) {
                break;
            }
        }
        return name;
    }

    function swatchExists(name) {
        try {
            doc.swatches.getByName(name);
            return true;
        } catch (e) {
            return false;
        }
    }

    function swatchGroupExists(name) {
        try {
            doc.swatchGroups.getByName(name);
            return true;
        } catch (e) {
            return false;
        }
    }

// カラーを「グローバルカラー（プロセス）」に変換して返す
function toGlobalProcessColor(doc, baseColor, baseName) {
    try {
        var spot = doc.spots.add();
        spot.name = baseName;
        spot.colorType = ColorModel.PROCESS; // グローバル（プロセス）
        spot.color = baseColor;

        var sc = new SpotColor();
        sc.spot = spot;
        sc.tint = 100;
        return sc;
    } catch (e) {
        // 失敗時は元のカラーを返す（無言）
        return baseColor;
    }
}

    function addSwatchForColor(colorObj, baseName) {
        var s = doc.swatches.add();
        var nm = uniqueName(baseName, swatchExists);
        s.name = nm;

        // グローバルカラー（プロセス）に変換して登録
        var globalColor = toGlobalProcessColor(doc, colorObj, nm);
        s.color = globalColor;

        try { s.selected = false; } catch (e) {}
        return s;
    }

        function getUnlockedVisibleLayer(doc) {
        try {
            if (doc.activeLayer && !doc.activeLayer.locked && doc.activeLayer.visible) return doc.activeLayer;
        } catch (e) {}
        for (var i = 0; i < doc.layers.length; i++) {
            try {
                if (!doc.layers[i].locked && doc.layers[i].visible) return doc.layers[i];
            } catch (e2) {}
        }
        return null;
    }

    // 選択オブジェクトから色を抽出
    var colors = collectColorsFromSelection(doc.selection);

    if (colors.length < 2) {
        return;
    }

    try {
        // 新規スウォッチグループを作成（重複回避）
        var baseGroupName = "AutoGradient";
        var groupName = uniqueName(baseGroupName, swatchGroupExists);
        var swGroup = doc.swatchGroups.add();
        swGroup.name = groupName;

        // 抽出色をスウォッチに登録（順番は選択の走査順）
        var createdSwatches = [];
        for (var i = 0; i < colors.length; i++) {
            var cs = addSwatchForColor(colors[i], "AutoColor");
            createdSwatches.push(cs);
            try { swGroup.addSwatch(cs); } catch (eAdd1) {}
        }

        // オブジェクトの選択解除（以降の処理は選択に依存しない）
        try {
            doc.selection = null;
        } catch (eSelClear) {}

        // 新しいグラデーションオブジェクトを作成
        var newGradient = doc.gradients.add();
        newGradient.type = GradientType.LINEAR; // 線形グラデーション（必要に応じてRADIALに変更可）

        // ストップ数を抽出色数に合わせる
        while (newGradient.gradientStops.length < colors.length) {
            newGradient.gradientStops.add();
        }
        while (newGradient.gradientStops.length > colors.length) {
            newGradient.gradientStops[newGradient.gradientStops.length - 1].remove();
        }

        // 抽出色をグラデーションストップに適用
        for (var j = 0; j < colors.length; j++) {
            var stop = newGradient.gradientStops[j];

            // 位置（RampPoint）を計算 (0 〜 100)
            var location = (j / (colors.length - 1)) * 100;
            stop.rampPoint = location;

            // 色を適用（スウォッチ登録時に作成したグローバルカラーを優先）
            try {
                if (createdSwatches && createdSwatches[j] && createdSwatches[j].color) {
                    stop.color = createdSwatches[j].color; // SpotColor（グローバル）を渡す
                } else {
                    stop.color = colors[j];
                }
            } catch (eStopColor) {
                // 無言フォールバック
                try { stop.color = colors[j]; } catch (e2) {}
            }

            // 中間点（MidPoint）をデフォルトの50に設定
            stop.midPoint = 50;

            // 不透明度
            stop.opacity = 100;
        }

        // グラデーション名（重複回避）
        var baseGradientName = "New Gradient";
        var gradientName = uniqueName(baseGradientName, function(nm) {
            try {
                // gradients.getByName は例外で判定
                doc.gradients.getByName(nm);
                return true;
            } catch (e) {
                return false;
            }
        });
        newGradient.name = gradientName;

                // 参考: ビュー中心に長方形を作成して、作成したグラデーションを適用
        try {
            var targetLayer = getUnlockedVisibleLayer(doc);
            if (targetLayer) {
                var viewCenterX = doc.activeView.centerPoint[0];
                var viewCenterY = doc.activeView.centerPoint[1];

                var RECT_SIZE = 100; // 必要なら変更
                var rectWidth = RECT_SIZE;
                var rectHeight = RECT_SIZE;
                var rectTop = viewCenterY + RECT_SIZE / 2;
                var rectLeft = viewCenterX - RECT_SIZE / 2;

                var rect = targetLayer.pathItems.rectangle(rectTop, rectLeft, rectWidth, rectHeight);

                // 作成した長方形を選択状態にする
                try {
                    doc.selection = null; // 念のためクリア
                    rect.selected = true;
                } catch (eSelRect) {}

                rect.stroked = false;
                rect.filled = true;

                var gc = new GradientColor();
                gc.gradient = newGradient;
                rect.fillColor = gc;
            }
        } catch (eRect) {}

        // 最後に追加されたスウォッチ（= 作成したグラデーション）を選択
        try {
            var idx = doc.swatches.length - 1;
            if (idx >= 0) {
                doc.swatches[idx].selected = true;
            }
        } catch (e) {}


    } catch (e) {
        // 無言（エラー通知しない）
    }
}

main();
