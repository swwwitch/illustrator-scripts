#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトの回転を水平（0°）に補正します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ResetRotation.md

### Overview

Corrects the rotation of the selected objects back to horizontal (0°).

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ResetRotation.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ResetRotation";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-15";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ResetRotation.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ResetRotation.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* ダイアログの初期値 / Dialog defaults */
    var DEFAULT_SETTINGS = {
        targetText: true,         /* テキストを対象にする / Level text */
        targetImage: true,        /* 配置画像を対象にする / Level placed images */
        targetRectangle: true,    /* 長方形を対象にする / Level rectangles */
        useClipGroupHost: true,   /* クリップグループごと回転する（常に最上位のクリップ） / Rotate the topmost clipping group */
        resetTextScale: true,     /* 文字の水平・垂直比率を100%に戻す / Reset character scale to 100% */
        epsilonDeg: 0.1           /* 水平とみなす角度（度） / Angle treated as level (degrees) */
    };

    /* 「水平とみなす範囲」の入力範囲（度） / Accepted range of the level tolerance (degrees) */
    var MIN_EPSILON_DEG = 0.01;
    var MAX_EPSILON_DEG = 10;

    // =========================================
    // レイアウト / Layout
    // =========================================
    var PANEL_MARGINS            = [15, 20, 15, 10];  /* パネル余白 [左,上,右,下] */
    var EPSILON_FIELD_CHARACTERS = 6;                 /* 「水平とみなす範囲」欄の幅 */

    /**
     * 左寄せの縦並びパネルを追加する
     * @param {Window} parent - 追加先
     * @param {string} labelPath - パネル見出しのラベルパス
     * @returns {Panel} 追加したパネル
     */
    function addPanel(parent, labelPath) {
        var panel = parent.add("panel", undefined, getLabel(labelPath));
        panel.orientation = "column";
        panel.alignChildren = "left";
        panel.margins = PANEL_MARGINS;
        panel.alignment = "left";
        return panel;
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * Illustrator の UI 言語から表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = detectUILang();

    var LABELS = {
        dialog: {
            title: { ja: "水平補正", en: "Level Objects" }
        },
        panel: {
            targets: { ja: "対象", en: "Objects to Level" },
            text: { ja: "テキスト", en: "Text" },
            options: { ja: "補正条件", en: "Correction Options" }
        },
        checkbox: {
            text: { ja: "テキスト", en: "Text" },
            image: { ja: "配置画像（埋め込み・リンク）", en: "Placed Image (Embedded/Linked)" },
            rectangle: { ja: "長方形（パス）", en: "Rectangle (Path)" },
            clipGroup: { ja: "クリップグループ", en: "Clipping Group" },
            resetTextScale: { ja: "縦横比を正す", en: "Reset Character Scale" }
        },
        fieldLabel: {
            epsilon: { ja: "水平とみなす範囲(°)", en: "Level Tolerance (°)" }
        },
        tooltip: {
            text: { ja: "テキストオブジェクトの回転を元に戻します。", en: "Clears the rotation on text objects." },
            image: { ja: "配置画像の回転を元に戻します。", en: "Clears the rotation on placed images." },
            rectangle: { ja: "長方形の回転を元に戻します。", en: "Clears the rotation on rectangles." },
            clipGroup: {
                ja: "クリップグループ内のオブジェクトは、最上位のクリップグループごと回転を戻します。",
                en: "Objects inside clipping groups are leveled by rotating the topmost clipping group."
            },
            resetTextScale: {
                ja: "テキストの文字の水平比率・垂直比率を100%に戻します。",
                en: "Resets the horizontal and vertical character scale of text to 100%."
            },
            epsilon: {
                ja: "これ以下の角度は0とみなします。わずかな傾きを無視するための値です（0.01〜10）。",
                en: "Angles below this count as zero, so tiny tilts are ignored (0.01–10)."
            }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "オブジェクトを選択してください。", en: "Please select objects." }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す
     * @param {string} labelPath - "panel.targets" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < labelPathKeys.length; i++) {
            labelNode = labelNode[labelPathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return labelNode[uiLang] || labelNode["en"] || labelPath;
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {string} labelPath - ラベルのパス
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelPath) {
        return getLabel(labelPath) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // 実行時の設定 / Runtime options
    // =========================================

    /* ダイアログで確定した補正条件（確定までは初期値） / Options confirmed in the dialog (defaults until then) */
    var levelOptions = {
        useClipGroupHost: DEFAULT_SETTINGS.useClipGroupHost,
        resetTextScale: DEFAULT_SETTINGS.resetTextScale,
        epsilonDeg: DEFAULT_SETTINGS.epsilonDeg
    };

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ↑↓キーで数値欄の値を増減する（Shift=10刻み、Option=0.1刻み）
     * @param {EditText} editText - 対象の数値欄
     * @returns {void}
     */
    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function (event) {
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var isUp = (event.keyName == "Up");
            var isDown = (event.keyName == "Down");

            if (keyboard.shiftKey) {
                /* 10の倍数にスナップ / Snap to multiples of 10 */
                if (isUp) {
                    value = Math.ceil((value + 1) / 10) * 10;
                    event.preventDefault();
                } else if (isDown) {
                    value = Math.floor((value - 1) / 10) * 10;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            } else if (keyboard.altKey) {
                if (isUp) {
                    value += 0.1;
                    event.preventDefault();
                } else if (isDown) {
                    value -= 0.1;
                    event.preventDefault();
                }
            } else {
                if (isUp) {
                    value += 1;
                    event.preventDefault();
                } else if (isDown) {
                    value -= 1;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            }

            /* Option 時は小数第1位、それ以外は整数に丸める / Round to 0.1 with Option, otherwise to integers */
            value = keyboard.altKey ? Math.round(value * 10) / 10 : Math.round(value);

            editText.text = value;
        });
    }

    /**
     * ラベルと tooltip つきのチェックボックスを追加する
     * @param {Panel} parent - 追加先
     * @param {string} labelKey - checkbox / tooltip 共通のキー
     * @param {boolean} initialValue - 初期値
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addCheckbox(parent, labelKey, initialValue) {
        var optionCheckbox = parent.add("checkbox", undefined, getLabel("checkbox." + labelKey));
        optionCheckbox.helpTip = getLabel("tooltip." + labelKey);
        optionCheckbox.value = initialValue;
        return optionCheckbox;
    }

    /**
     * 「水平とみなす範囲」の入力値を読み、範囲内に収める
     * @param {string} inputText - 入力された文字列
     * @returns {number} 角度（度）。数値でなければ初期値
     */
    function parseEpsilon(inputText) {
        var value = parseFloat(inputText);
        if (isNaN(value)) return DEFAULT_SETTINGS.epsilonDeg;
        return Math.max(MIN_EPSILON_DEG, Math.min(MAX_EPSILON_DEG, value));
    }

    /**
     * 対象と補正条件を選ぶダイアログを表示する
     * @returns {{targets: {text: boolean, image: boolean, rectangle: boolean}, useClipGroupHost: boolean, resetTextScale: boolean, epsilonDeg: number}|null} 選んだ内容（キャンセル時は null）
     */
    function showLevelDialog() {
        var dlg = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);

        var targetsPanel = addPanel(dlg, "panel.targets");
        var textCheckbox = addCheckbox(targetsPanel, "text", DEFAULT_SETTINGS.targetText);
        var imageCheckbox = addCheckbox(targetsPanel, "image", DEFAULT_SETTINGS.targetImage);
        var rectangleCheckbox = addCheckbox(targetsPanel, "rectangle", DEFAULT_SETTINGS.targetRectangle);
        var clipGroupCheckbox = addCheckbox(targetsPanel, "clipGroup", DEFAULT_SETTINGS.useClipGroupHost);

        var textPanel = addPanel(dlg, "panel.text");
        var resetTextScaleCheckbox = addCheckbox(textPanel, "resetTextScale", DEFAULT_SETTINGS.resetTextScale);

        var optionsPanel = addPanel(dlg, "panel.options");
        var epsilonRow = optionsPanel.add("group");
        epsilonRow.orientation = "row";
        epsilonRow.alignChildren = "left";
        epsilonRow.add("statictext", undefined, labelText("fieldLabel.epsilon"));
        var epsilonInput = epsilonRow.add("edittext", undefined, String(DEFAULT_SETTINGS.epsilonDeg));
        epsilonInput.helpTip = getLabel("tooltip.epsilon");
        epsilonInput.characters = EPSILON_FIELD_CHARACTERS;
        changeValueByArrowKey(epsilonInput);

        /* ボタン行（中央）/ Button row (centered) */
        var btnRowGroup = dlg.add("group");
        btnRowGroup.alignment = "center";
        btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        dlg.layout.layout(true);
        dlg.center();
        if (dlg.show() !== 1) return null;

        return {
            targets: {
                text: textCheckbox.value,
                image: imageCheckbox.value,
                rectangle: rectangleCheckbox.value
            },
            useClipGroupHost: clipGroupCheckbox.value,
            resetTextScale: resetTextScaleCheckbox.value,
            epsilonDeg: parseEpsilon(epsilonInput.text)
        };
    }

    // =========================================
    // 対象の判定 / Target detection
    // =========================================

    /**
     * 配置画像（リンク）か
     * @param {PageItem} item - 判定するオブジェクト
     * @returns {boolean} PlacedItem なら true
     */
    function isPlacedImage(item) {
        return item.typename === "PlacedItem";
    }

    /**
     * 埋め込み画像か
     * @param {PageItem} item - 判定するオブジェクト
     * @returns {boolean} RasterItem なら true
     */
    function isEmbeddedImage(item) {
        return item.typename === "RasterItem";
    }

    /**
     * 単純な長方形（閉じた4点のパス）か
     * @param {PageItem} item - 判定するオブジェクト
     * @returns {boolean} 閉じた4点の PathItem なら true
     */
    function isRectangle(item) {
        return item.typename === "PathItem" && item.closed && item.pathPoints.length === 4;
    }

    /**
     * 縦書きのテキストか
     * @param {TextFrame} textFrame - 判定するテキスト
     * @returns {boolean} 縦書きなら true
     */
    function isVertical(textFrame) {
        /* 空のテキストでは textRanges[0] が取れないことがある / textRanges[0] may fail on empty text */
        try {
            return textFrame.textRanges[0].characterAttributes.orientation === TextOrientation.VERTICAL;
        } catch (e) {
            return false;
        }
    }

    /**
     * オブジェクトごとの識別キーを返す（同じホストを二度回さないため）
     * @param {PageItem} item - 対象
     * @returns {string} 識別キー
     */
    function getItemKey(item) {
        /* uuid を持たない環境・オブジェクトに備える / Guard against items without uuid */
        try {
            if (item.uuid) return "u:" + item.uuid;
        } catch (e) {}
        try {
            if (item.index !== undefined) return item.toString() + "|" + item.index;
        } catch (e2) {}
        return "" + item;
    }

    /**
     * 親をたどり、最上位のクリップグループを返す
     * @param {PageItem} item - 起点のオブジェクト
     * @returns {PageItem} 最上位のクリップグループ（無ければ item 自身）
     */
    function getTopmostClippingGroup(item) {
        var host = null;
        /* 親が Layer・Document まで続くので、途中で読めないプロパティに備える / Parents run up to Layer/Document */
        try {
            var parentItem = item.parent;
            while (parentItem) {
                if (parentItem.typename === "GroupItem" && parentItem.clipped) host = parentItem;
                parentItem = parentItem.parent;
            }
        } catch (e) {}
        return host || item;
    }

    /* 回転の対象（ホスト）の解決結果 / Resolved rotation hosts */
    var hostCache = {};

    /**
     * 実際に回転させるオブジェクト（ホスト）を返す
     * クリップグループ設定がONならクリップグループ、OFFならオブジェクト自身
     * @param {PageItem} item - 対象のオブジェクト
     * @returns {PageItem} 回転させるオブジェクト
     */
    function resolveHost(item) {
        if (!levelOptions.useClipGroupHost) return item;
        var itemKey = getItemKey(item);
        if (!hostCache[itemKey]) {
            hostCache[itemKey] = getTopmostClippingGroup(item);
        }
        return hostCache[itemKey];
    }

    /**
     * 選択を再帰的にたどり、種類別に対象を集める
     * @param {Object} items - PageItem のコレクションまたは配列
     * @param {{text: boolean, image: boolean, rectangle: boolean}} targets - 対象の種類
     * @param {{texts: TextFrame[], images: PageItem[], rectangles: PathItem[]}} buckets - 集めた対象の格納先
     * @returns {void}
     */
    function collectTargets(items, targets, buckets) {
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (item.typename === "GroupItem") {
                collectTargets(item.pageItems, targets, buckets);
            } else if (item.typename === "CompoundPathItem") {
                collectTargets(item.pathItems, targets, buckets);
            } else if (item.pageItems) {
                /* その他の入れ子（クリップグループ等） / Other containers */
                try {
                    collectTargets(item.pageItems, targets, buckets);
                } catch (e) {}
            }

            if (targets.text && item.typename === "TextFrame") buckets.texts.push(item);
            if (targets.image && (isPlacedImage(item) || isEmbeddedImage(item))) buckets.images.push(item);
            /* クリップグループモードではマスクパスを個別に回さない / Mask paths are not leveled on their own in clip-group mode */
            if (targets.rectangle && isRectangle(item) && !(levelOptions.useClipGroupHost && item.clipping)) {
                buckets.rectangles.push(item);
            }
        }
    }

    // =========================================
    // 角度の計算 / Angle computation
    // =========================================

    /**
     * ほぼ水平（0° または 360° 付近）か
     * @param {number} deg - 角度（度）
     * @returns {boolean} しきい値未満なら true
     */
    function isNearlyLevel(deg) {
        var absDeg = Math.abs(deg);
        var eps = levelOptions.epsilonDeg;
        return (absDeg < eps) || (Math.abs(absDeg - 360) < eps);
    }

    /**
     * 角度を (-180, 180] に正規化する
     * @param {number} deg - 角度（度）
     * @returns {number} 正規化した角度
     */
    function normalizeAngle(deg) {
        var angle = deg % 360;
        if (angle <= -180) angle += 360;
        if (angle > 180) angle -= 360;
        return angle;
    }

    /**
     * グループの最初の子で条件に合うものを返す
     * @param {GroupItem} group - グループ
     * @param {function(PageItem): boolean} predicate - 条件
     * @returns {PageItem|null} 見つかった子（無ければ null）
     */
    function findChild(group, predicate) {
        for (var i = 0; i < group.pageItems.length; i++) {
            if (predicate(group.pageItems[i])) return group.pageItems[i];
        }
        return null;
    }

    /**
     * クリップグループの見かけの角度を子から推定する
     * 優先順：クリッピングパス → 配置／埋め込み画像 → 通常のパス → 入れ子のグループ
     * @param {GroupItem} group - グループ
     * @returns {number} 推定した角度（度）。推定できなければ 0
     */
    function getGroupProxyAngle(group) {
        /* 子の読み取りに失敗したら 0° とみなす / Treat unreadable children as 0° */
        try {
            var referenceItem =
                findChild(group, function (child) { return child.typename === "PathItem" && child.clipping; }) ||
                findChild(group, function (child) { return child.typename === "PlacedItem" || child.typename === "RasterItem"; }) ||
                findChild(group, function (child) { return child.typename === "PathItem" && !child.clipping; });
            if (referenceItem) return getRotationDegrees(referenceItem);

            /* 入れ子のグループは 0 以外が出た時点で返す / Return the first non-zero angle from nested groups */
            for (var i = 0; i < group.pageItems.length; i++) {
                if (group.pageItems[i].typename === "GroupItem") {
                    var nestedAngle = getGroupProxyAngle(group.pageItems[i]);
                    if (nestedAngle !== 0) return nestedAngle;
                }
            }
        } catch (e) {}
        return 0;
    }

    /**
     * オブジェクトの変換行列を返す
     * @param {PageItem} item - 対象
     * @returns {Matrix|null} 行列（読めなければ null）
     */
    function getMatrix(item) {
        /* matrix を持たない種類（PathItem・GroupItem 等）がある / Some item types have no matrix */
        try {
            var matrix = item.matrix;
            return (matrix && matrix.mValueA !== undefined) ? matrix : null;
        } catch (e) {
            return null;
        }
    }

    /**
     * オブジェクトの回転角を推定する
     * 行列 → パスの最初の辺 → グループの子、の順に試す
     * @param {PageItem} item - 対象
     * @returns {number} 回転角（度）
     */
    function getRotationDegrees(item) {
        var matrix = getMatrix(item);
        if (matrix && matrix.mValueB !== undefined) {
            var deg = Math.atan2(matrix.mValueB, matrix.mValueA) * 180 / Math.PI;
            /* グループの見かけは子の回転に依存することがある / A group's apparent angle may come from its children */
            if (item.typename === "GroupItem" && isNearlyLevel(deg)) {
                var proxyAngle = getGroupProxyAngle(item);
                if (!isNearlyLevel(proxyAngle)) return proxyAngle;
            }
            return deg;
        }

        /* パスは最初の辺（アンカー0→1）の向きから推定 / Infer a path's angle from its first segment */
        if (item.typename === "PathItem" && item.pathPoints && item.pathPoints.length >= 2) {
            try {
                var p0 = item.pathPoints[0].anchor;
                var p1 = item.pathPoints[1].anchor;
                return Math.atan2(p1[1] - p0[1], p1[0] - p0[0]) * 180 / Math.PI;
            } catch (e) {}
        }

        if (item.typename === "GroupItem") {
            var groupAngle = getGroupProxyAngle(item);
            if (!isNearlyLevel(groupAngle)) return groupAngle;
        }

        /* 推定できない種類は 0° とみなす / Assume 0° for anything else */
        return 0;
    }

    /**
     * 鏡像（行列式が負）か
     * @param {PageItem} item - 対象
     * @returns {boolean} 反転していれば true
     */
    function isMirroredTransform(item) {
        var matrix = getMatrix(item);
        if (!matrix) return false;
        return (matrix.mValueA * matrix.mValueD) - (matrix.mValueB * matrix.mValueC) < 0;
    }

    /**
     * 水平に戻すための回転量を返す
     * @param {PageItem} host - 回転させるオブジェクト
     * @returns {number|null} 回転量（度）。すでに水平なら null
     */
    function getLevelingDelta(host) {
        var normalized = normalizeAngle(getRotationDegrees(host));
        if (isNearlyLevel(normalized)) return null;
        /* 鏡像は回転方向が逆になる / Mirrored items rotate the other way */
        return isMirroredTransform(host) ? normalized : -normalized;
    }

    // =========================================
    // 補正 / Leveling
    // =========================================

    /**
     * オブジェクトを中心で回転し、バウンディングボックスをリセットする
     * @param {PageItem} item - 回転させるオブジェクト
     * @param {number} deg - 回転量（度）
     * @returns {void}
     */
    function rotateBy(item, deg) {
        var doc = app.activeDocument;
        var previousSelection = doc.selection;

        /* ロック・非表示のオブジェクトは回せない / Locked or hidden items cannot be rotated */
        try {
            item.rotate(deg, true, true, true, true, Transformation.CENTER);
        } catch (e) {
            return;
        }

        /* 対象だけを選んでバウンディングボックスをリセット / Select only the item and reset its bounding box */
        try {
            doc.selection = null;
            doc.selection = [item];
            app.executeMenuCommand("AI Reset Bounding Box");
        } catch (e2) {}

        /* 元の選択に戻す（消えたオブジェクトがあると例外） / Restore the selection (throws if an item is gone) */
        try {
            if (previousSelection) doc.selection = previousSelection;
        } catch (e3) {}
    }

    /**
     * 対象ごとにホストを解決し、同じホストは一度だけ水平に戻す
     * @param {PageItem[]} items - 対象
     * @param {function(PageItem, PageItem): boolean} [prepareItem] - 回転前の処理。false を返すとその対象を飛ばす
     * @returns {number} 回転させたホストの数
     */
    function levelItems(items, prepareItem) {
        var rotatedCount = 0;
        var seenHosts = {};
        for (var i = 0; i < items.length; i++) {
            var host = resolveHost(items[i]);
            if (prepareItem && !prepareItem(items[i], host)) continue;

            var hostKey = getItemKey(host);
            if (seenHosts[hostKey]) continue;
            seenHosts[hostKey] = true;

            var delta = getLevelingDelta(host);
            if (delta === null) continue;
            rotateBy(host, delta);
            rotatedCount++;
        }
        return rotatedCount;
    }

    /**
     * テキストの回転前処理：クリップグループ外の縦書きは飛ばし、必要なら文字比率を戻す
     * @param {TextFrame} textFrame - テキスト
     * @param {PageItem} host - 回転させるオブジェクト
     * @returns {boolean} 処理を続けるなら true
     */
    function prepareTextFrame(textFrame, host) {
        if (host === textFrame && isVertical(textFrame)) return false;
        if (levelOptions.resetTextScale) resetTextScaling(textFrame);
        return true;
    }

    /**
     * テキストの文字の水平・垂直比率を100%に戻す
     * @param {TextFrame} textFrame - テキスト
     * @returns {void}
     */
    function resetTextScaling(textFrame) {
        /* 空のテキストやロック中は書き込めないことがある / May fail on empty or locked text */
        try {
            var charAttributes = textFrame.textRange.characterAttributes;
            charAttributes.horizontalScale = 100;
            charAttributes.verticalScale = 100;
        } catch (e) {}
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログで条件を選び、選択オブジェクトの回転を水平に戻す
     * @returns {void}
     */
    function main() {
        if (!app.documents.length) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var doc = app.activeDocument;
        if (!doc.selection || doc.selection.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        var dialogResult = showLevelDialog();
        if (!dialogResult) return;
        levelOptions.useClipGroupHost = dialogResult.useClipGroupHost;
        levelOptions.resetTextScale = dialogResult.resetTextScale;
        levelOptions.epsilonDeg = dialogResult.epsilonDeg;
        hostCache = {};

        var buckets = { texts: [], images: [], rectangles: [] };
        collectTargets(doc.selection, dialogResult.targets, buckets);

        var rotatedCount = 0;
        if (buckets.texts.length) rotatedCount += levelItems(buckets.texts, prepareTextFrame);
        if (buckets.images.length) rotatedCount += levelItems(buckets.images);
        if (buckets.rectangles.length) rotatedCount += levelItems(buckets.rectangles);

        if (rotatedCount > 0) app.redraw();
    }

    main();

})();
