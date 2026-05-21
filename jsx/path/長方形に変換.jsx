#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### スクリプト名：

長方形に変換.jsx

### Readme （GitHub）：

https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/長方形に変換.md

### 概要：

- 更新日：2026-05-20
- 選択したオブジェクトと同じ大きさの長方形を作成する Illustrator スクリプト
- 個別／グループ単位の単位選択、計測基準、塗り・線プリセット、重ね順、元オブジェクトの扱いをダイアログで指定

### 主な機能：

- 個別／グループ（選択範囲全体）単位で長方形を作成
- 計測基準にプレビュー境界・テキストのアウトライン化を選択可能
- 塗り・線プリセットの選択
- 重ね順（前面／背面）の指定
- 元オブジェクトの扱い（そのまま／マスク／削除）の指定

### 更新履歴：

- v1.0.0 (2026-05-20) : 初期バージョン

*/

/*

### Script Name:

長方形に変換.jsx (ConvertToRectangle in Japanese)

### GitHub:

https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ConvertToRectangleJa.md

### Description:

- Last Updated: 2026-05-20
- Creates a rectangle matching the size of each selected object
- Provides options for per-object / whole-selection unit, bounds basis, fill/stroke preset, stacking order and original handling via dialog

### Main Features:

- Creates rectangles per object or for the whole selection
- Bounds basis: preview bounds or outlined text bounds
- Fill / stroke presets
- Stacking order (front / back)
- Original handling (keep / mask / delete)

### Changelog:

- v1.0.0 (2026-05-20) : Initial version

*/

(function () {

    // ==============================
    // スクリプト情報 / Script information
    // ==============================
    var SCRIPT_VERSION = "v1.0.0";

    // ==============================
    // 言語判定 / Language detection
    // ==============================
    function getCurrentLang() {
        return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var lang = getCurrentLang();

    // ==============================
    // ラベル定義 / Label definitions
    // ==============================
    var LABELS = {
        dialogTitle: { ja: "長方形に変換", en: "Convert to Rectangle" },
        targetPanel: { ja: "対象", en: "Target" },
        targetEach: { ja: "個別に", en: "Individually" },
        targetGroup: { ja: "グループとして", en: "As a group" },
        optionsPanel: { ja: "オプション", en: "Options" },
        previewBounds: { ja: "プレビュー境界", en: "Preview bounds" },
        boundsTip: { ja: "オンで線幅・効果を含む見た目のサイズ（visibleBounds）を基準にします", en: "When on, uses the visible size including stroke/effects (visibleBounds)" },
        outlineText: { ja: "テキストをアウトライン化", en: "Outline text" },
        outlineTextTip: { ja: "テキストを複製・アウトライン化して計測し、複製は削除します", en: "Duplicates and outlines text to measure bounds, then deletes the copy" },
        appearancePanel: { ja: "塗りと線", en: "Fill and stroke" },
        imageNotice: { ja: "リンク画像・配置画像には使用できません", en: "Not available for linked or placed images" },
        orderPanel: { ja: "重ね順", en: "Stacking order" },
        orderFront: { ja: "前面", en: "Front" },
        orderBack: { ja: "背面", en: "Back" },
        originalPanel: { ja: "元のオブジェクト", en: "Original object" },
        originalKeep: { ja: "そのまま", en: "Keep" },
        originalMask: { ja: "マスク", en: "Mask" },
        originalDelete: { ja: "削除", en: "Delete" },
        cancel: { ja: "キャンセル", en: "Cancel" },
        noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
        noSelection: { ja: "オブジェクトを選択してください。", en: "Please select one or more objects." },
        noResult: { ja: "長方形を作成できるオブジェクトがありませんでした。", en: "No rectangles could be created." }
    };

    // ラベル取得 / Get label
    function getLabel(labelKey) {
        var labelEntry = LABELS[labelKey];
        if (!labelEntry) return labelKey;
        return labelEntry[lang] || labelEntry.en || labelKey;
    }

    /* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
    function labelText(labelKey) {
        return getLabel(labelKey) + (lang === 'ja' ? '：' : ':');
    }

    // ==============================
    // 塗り・線プリセット / Fill & stroke presets
    // ==============================
    // applyAppearance: true … 生成した長方形にこのプリセットの塗り・線を適用する
    // applyAppearance: false（直前の塗り・線）… 適用せず、作成時の既定値のままにする
    var FILL_STROKE_PRESETS = [
        { ja: "塗り：なし、線：なし", en: "Fill: none, Stroke: none", applyAppearance: true, filled: false, stroked: false, strokeWidth: 0 },
        { ja: "線：黒、1pt", en: "Stroke: black, 1pt", applyAppearance: true, filled: false, stroked: true, strokeWidth: 1 },
        { ja: "線：黒、0.25pt", en: "Stroke: black, 0.25pt", applyAppearance: true, filled: false, stroked: true, strokeWidth: 0.25 },
        { ja: "直前の塗り・線の情報", en: "Current fill / stroke", applyAppearance: false }
    ];
    var DEFAULT_PRESET_INDEX = 1; // 既定：線：黒、1pt

    // 「直前の塗り・線の情報」プリセットの位置（applyAppearance: false）
    var CURRENT_PRESET_INDEX = (function () {
        for (var i = 0; i < FILL_STROKE_PRESETS.length; i++) {
            if (!FILL_STROKE_PRESETS[i].applyAppearance) return i;
        }
        return -1;
    })();

    // ドキュメントのカラースペースに合わせた黒を返す / Returns black for the document color space
    function makeBlack(isRgbDocument) {
        if (isRgbDocument) {
            var rgbColor = new RGBColor();
            rgbColor.red = 0;
            rgbColor.green = 0;
            rgbColor.blue = 0;
            return rgbColor;
        }
        var cmykColor = new CMYKColor();
        cmykColor.cyan = 0;
        cmykColor.magenta = 0;
        cmykColor.yellow = 0;
        cmykColor.black = 100;
        return cmykColor;
    }

    // ==============================
    // 選択オブジェクトの判定 / Selection inspection
    // ==============================

    // リンク画像（PlacedItem）または配置画像（RasterItem）かどうか
    function isImageItem(item) {
        return item.typename === "PlacedItem" || item.typename === "RasterItem";
    }

    // 選択オブジェクトがすべて画像かどうか / True when every selected item is an image
    function isSelectionAllImages(pageItems) {
        if (!pageItems || pageItems.length === 0) return false;
        for (var i = 0; i < pageItems.length; i++) {
            if (!isImageItem(pageItems[i])) return false;
        }
        return true;
    }

    // ==============================
    // 設定ダイアログ / Settings dialog
    // ==============================
    function showSettingsDialog(selectionAllImages) {
        var dialog = new Window("dialog", getLabel("dialogTitle") + "  " + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = "fill";
        dialog.margins = 16;
        dialog.spacing = 12;

        // パネル共通の余白・間隔 / Shared panel margins & spacing
        var PANEL_MARGINS = [15, 20, 15, 10];
        var PANEL_SPACING = 8;

        function setupPanel(panel, spacing) {
            panel.orientation = "column";
            panel.alignChildren = ['fill', 'top'];
            panel.alignment = "fill";
            panel.margins = PANEL_MARGINS;
            panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
        }

        // --- 対象 / Target ---
        var targetPanel = dialog.add("panel", undefined, getLabel("targetPanel"));
        setupPanel(targetPanel);
        targetPanel.orientation = "row";
        targetPanel.alignChildren = ["left", "center"];
        var targetEachRadio = targetPanel.add("radiobutton", undefined, getLabel("targetEach"));
        var targetGroupRadio = targetPanel.add("radiobutton", undefined, getLabel("targetGroup"));
        targetEachRadio.value = true;

        // --- オプション / Options ---
        // プレビュー境界：オフ＝geometricBounds（線幅・効果なし）/ オン＝visibleBounds（見た目のサイズ）
        var optionsPanel = dialog.add("panel", undefined, getLabel("optionsPanel"));
        setupPanel(optionsPanel);
        var previewBoundsCheck = optionsPanel.add("checkbox", undefined, getLabel("previewBounds"));
        previewBoundsCheck.helpTip = getLabel("boundsTip");
        var outlineTextCheck = optionsPanel.add("checkbox", undefined, getLabel("outlineText"));
        outlineTextCheck.helpTip = getLabel("outlineTextTip");

        // --- 塗りと線 / Fill and stroke ---
        var appearancePanel = dialog.add("panel", undefined, getLabel("appearancePanel"));
        setupPanel(appearancePanel);
        var appearanceRadios = [];
        for (var i = 0; i < FILL_STROKE_PRESETS.length; i++) {
            var presetLabel = FILL_STROKE_PRESETS[i][lang] || FILL_STROKE_PRESETS[i].en;
            appearanceRadios.push(appearancePanel.add("radiobutton", undefined, presetLabel));
        }
        appearanceRadios[DEFAULT_PRESET_INDEX].value = true;

        // リンク画像・配置画像のときは「直前の塗り・線の情報」を選べないようにする
        if (selectionAllImages && CURRENT_PRESET_INDEX >= 0) {
            appearanceRadios[CURRENT_PRESET_INDEX].enabled = false;
            appearanceRadios[CURRENT_PRESET_INDEX].helpTip = getLabel("imageNotice");
        }

        // --- 重ね順 / Stacking order ---
        var orderPanel = dialog.add("panel", undefined, getLabel("orderPanel"));
        setupPanel(orderPanel);
        orderPanel.orientation = "row";
        orderPanel.alignChildren = ["left", "center"];
        var orderFrontRadio = orderPanel.add("radiobutton", undefined, getLabel("orderFront"));
        var orderBackRadio = orderPanel.add("radiobutton", undefined, getLabel("orderBack"));
        orderFrontRadio.value = true;

        // --- 元のオブジェクト / Original object ---
        var originalPanel = dialog.add("panel", undefined, getLabel("originalPanel"));
        setupPanel(originalPanel);
        originalPanel.orientation = "row";
        originalPanel.alignChildren = ["left", "center"];
        var originalKeepRadio = originalPanel.add("radiobutton", undefined, getLabel("originalKeep"));
        var originalMaskRadio = originalPanel.add("radiobutton", undefined, getLabel("originalMask"));
        var originalDeleteRadio = originalPanel.add("radiobutton", undefined, getLabel("originalDelete"));
        originalKeepRadio.value = true;

        // 重ね順は「そのまま」のときだけ意味を持つ / Stacking order applies only when keeping the original
        function syncOrderEnabled() {
            orderFrontRadio.enabled = originalKeepRadio.value;
            orderBackRadio.enabled = originalKeepRadio.value;
        }
        originalKeepRadio.onClick = syncOrderEnabled;
        originalMaskRadio.onClick = syncOrderEnabled;
        originalDeleteRadio.onClick = syncOrderEnabled;
        syncOrderEnabled();

        // --- ボタン / Buttons ---
        var buttonRow = dialog.add("group");
        buttonRow.alignment = "right";
        var cancelButton = buttonRow.add("button", undefined, getLabel("cancel"), { name: "cancel" });
        var okButton = buttonRow.add("button", undefined, "OK", { name: "ok" });

        var dialogResult = null;
        okButton.onClick = function () {
            var appearanceIndex = DEFAULT_PRESET_INDEX;
            for (var i = 0; i < appearanceRadios.length; i++) {
                if (appearanceRadios[i].value) {
                    appearanceIndex = i;
                    break;
                }
            }

            var originalMode = "keep";
            if (originalMaskRadio.value) {
                originalMode = "mask";
            } else if (originalDeleteRadio.value) {
                originalMode = "delete";
            }

            dialogResult = {
                groupAsOne: targetGroupRadio.value,
                useVisibleBounds: previewBoundsCheck.value,
                outlineText: outlineTextCheck.value,
                appearancePreset: FILL_STROKE_PRESETS[appearanceIndex],
                placeInFront: orderFrontRadio.value,
                originalMode: originalMode
            };
            dialog.close();
        };
        cancelButton.onClick = function () {
            dialog.close();
        };

        dialog.show();
        return dialogResult;
    }

    // ==============================
    // 塗り・線の適用 / Fill & stroke
    // ==============================
    // 生成した長方形そのものにプリセットの塗り・線を適用する（元の選択オブジェクトには影響しない）
    function applyRectAppearance(rect, preset, isRgbDocument) {
        if (!preset.applyAppearance) {
            return; // 直前の塗り・線：作成時の既定値のままにする
        }
        rect.filled = preset.filled;
        rect.stroked = preset.stroked;
        if (preset.stroked) {
            rect.strokeColor = makeBlack(isRgbDocument);
            rect.strokeWidth = preset.strokeWidth;
        }
    }

    // ==============================
    // 長方形の作成 / Rectangle creation
    // ==============================

    // オブジェクトの外接矩形を取得
    // geometricBounds: 線幅・効果を含まない / visibleBounds: 線幅・効果を含む見た目のサイズ
    function getItemBounds(item, useVisibleBounds) {
        return useVisibleBounds ? item.visibleBounds : item.geometricBounds;
    }

    // 計測用の外接矩形を取得する
    // テキスト＋「テキストをアウトライン化」設定時は、複製をアウトライン化して計測し複製は破棄する
    function measureItemBounds(item, settings) {
        if (settings.outlineText && item.typename === "TextFrame") {
            var duplicatedText = null;
            var outlinedGroup = null;
            try {
                duplicatedText = item.duplicate();
                outlinedGroup = duplicatedText.createOutline(); // 成功すると duplicatedText は消費され GroupItem になる
                var outlinedBounds = getItemBounds(outlinedGroup, settings.useVisibleBounds);
                outlinedGroup.remove();
                return outlinedBounds;
            } catch (e) {
                // 計測用の複製・アウトラインがドキュメントに残らないよう後始末する
                if (outlinedGroup) {
                    try { outlinedGroup.remove(); } catch (ignoreOutlined) { }
                } else if (duplicatedText) {
                    try { duplicatedText.remove(); } catch (ignoreDuplicate) { }
                }
            }
        }
        return getItemBounds(item, settings.useVisibleBounds);
    }

    // 複数オブジェクトをまとめた外接矩形を返す / Combined bounding box of multiple objects
    function getCombinedBounds(pageItems, settings) {
        var left = null, top = null, right = null, bottom = null;
        for (var i = 0; i < pageItems.length; i++) {
            var bounds = measureItemBounds(pageItems[i], settings);
            if (left === null || bounds[0] < left) left = bounds[0];
            if (top === null || bounds[1] > top) top = bounds[1];
            if (right === null || bounds[2] > right) right = bounds[2];
            if (bottom === null || bounds[3] < bottom) bottom = bounds[3];
        }
        if (left === null) return null;
        return [left, top, right, bottom];
    }

    // 指定した外接矩形から長方形を作成し、プリセットの塗り・線を適用する
    function createRectFromBounds(doc, bounds, settings, referenceItem) {
        var left = bounds[0];
        var top = bounds[1];
        var right = bounds[2];
        var bottom = bounds[3];

        var width = right - left;
        var height = top - bottom;
        if (width <= 0 || height <= 0) {
            return null;
        }

        var rect = doc.pathItems.rectangle(top, left, width, height);
        var placement = settings.placeInFront ? ElementPlacement.PLACEBEFORE : ElementPlacement.PLACEAFTER;
        // move() の戻り値は環境により undefined になるため、元の参照をそのまま使う
        rect.move(referenceItem, placement);

        // 塗り・線は生成した長方形に直接適用する（選択中の元オブジェクトには触れない）
        var isRgbDocument = (doc.documentColorSpace === DocumentColorSpace.RGB);
        applyRectAppearance(rect, settings.appearancePreset, isRgbDocument);
        return rect;
    }

    // ==============================
    // 元のオブジェクトの処理 / Handling the original object
    // ==============================

    // 長方形をクリッピングマスクとして適用し、マスクグループを返す
    function applyClippingMask(doc, clipRect, contentItems) {
        var clipGroup = doc.groupItems.add();
        // 元オブジェクトをグループへ移動
        for (var i = 0; i < contentItems.length; i++) {
            contentItems[i].move(clipGroup, ElementPlacement.MOVETOEND);
        }
        // クリップパス（長方形）は最前面に置く
        clipRect.move(clipGroup, ElementPlacement.MOVETOBEGINNING);
        clipRect.clipping = true;
        clipGroup.clipped = true;
        return clipGroup;
    }

    // 元のオブジェクトの扱い（そのまま／マスク／削除）を適用し、選択対象の項目を返す
    function applyOriginalMode(doc, rect, originalItems, originalMode) {
        if (originalMode === "mask") {
            return applyClippingMask(doc, rect, originalItems);
        }
        if (originalMode === "delete") {
            for (var i = 0; i < originalItems.length; i++) {
                originalItems[i].remove();
            }
            return rect;
        }
        return rect; // そのまま / keep
    }

    // ==============================
    // メイン処理 / Main
    // ==============================
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("noDocument"));
            return;
        }

        var doc = app.activeDocument;

        if (!doc.selection || doc.selection.length === 0) {
            alert(getLabel("noSelection"));
            return;
        }

        // 選択状態は後で変わるため、配列にコピーしておく
        var selectedItems = [];
        for (var i = 0; i < doc.selection.length; i++) {
            selectedItems.push(doc.selection[i]);
        }

        var selectionAllImages = isSelectionAllImages(selectedItems);

        var settings = showSettingsDialog(selectionAllImages);
        if (!settings) {
            return; // キャンセル / Cancelled
        }

        // ロック・非表示オブジェクトを除いた処理対象を抽出
        var eligibleItems = [];
        for (var i = 0; i < selectedItems.length; i++) {
            if (!selectedItems[i].locked && !selectedItems[i].hidden) {
                eligibleItems.push(selectedItems[i]);
            }
        }

        var createdItems = [];
        if (settings.groupAsOne) {
            // グループとして：選択全体をまとめて1つの長方形を作成
            var combinedBounds = getCombinedBounds(eligibleItems, settings);
            if (combinedBounds) {
                var groupRect = createRectFromBounds(doc, combinedBounds, settings, eligibleItems[0]);
                if (groupRect) {
                    createdItems.push(applyOriginalMode(doc, groupRect, eligibleItems, settings.originalMode));
                }
            }
        } else {
            // 個別に：オブジェクトごとに長方形を作成
            for (var i = 0; i < eligibleItems.length; i++) {
                var itemBounds = measureItemBounds(eligibleItems[i], settings);
                var rect = createRectFromBounds(doc, itemBounds, settings, eligibleItems[i]);
                if (rect) {
                    createdItems.push(applyOriginalMode(doc, rect, [eligibleItems[i]], settings.originalMode));
                }
            }
        }

        if (createdItems.length === 0) {
            alert(getLabel("noResult"));
            return;
        }

        // 作成した長方形（マスク時はマスクグループ）のみを選択
        doc.selection = null;
        for (var i = 0; i < createdItems.length; i++) {
            createdItems[i].selected = true;
        }
    }

    main();

})();