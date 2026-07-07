
#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);
#targetengine "DialogEngine"

/*

### スクリプト名：

ResetTransform.jsx（回転／シアー／スケール／縦横比のリセット）

### 概要：

- 更新日: 2026-07-08
- 配置画像・テキスト・長方形（パス）・クリップグループ・直線パスに対して、回転／シアー（せん断）／スケール／縦横比を安全にリセットします。
- BBox リセットと左上基準の再配置により、見た目の位置を安定して維持します。

### 主な機能：

- **配置画像／ラスタ**：回転・シアー・縦横比（小さい軸を大きい軸へ揃え、％は四捨五入の整数）・スケール（指定％）の個別／同時リセット。
- **テキスト**：回転・シアー・水平/垂直比率（100%）のリセット。
- **長方形（4点パス）**：近傍角度のみスナップして回転0°/90°へ補正。
- **直線（2点パス）**：近傍軸（0°/90°）へスナップして補正。
- **クリップグループ**：子要素（配置/ラスタ **＋ マスクパス**）の回転リセット＋縦横比（配置画像から導出した等比化Δを、配置画像とマスクパスに同時適用）。
- **UI**：2カラム、対象別パネル、ホットキー、数値入力（↑↓/Shift+↑↓）、ダイアログ位置/透明度。

### 処理の流れ：

1) 選択内容を解析し、対応パネルを自動ディム表示。
2) パネルで操作項目を選択 → OK。
3) 各ハンドラが変形を適用 → BBox リセット → 左上基準で再配置。

### 謝辞：

フジタノリアキさん

### 紹介記事（note）

https://note.com/dtp_tranist/n/n52f6b645bc70

### 更新履歴：

- v1.6.0 (20260708) : ↑↓キーの小数（0.1）刻みを廃止し整数のみに。スケール％の下限を 20% に変更。長方形の回転補正を最大44°まで拡大＋BBoxリセット化、実行後に元の選択を復元、ドキュメント未オープン時のガードを追加。
- v1.5.1 (20260708) : 内部整理（IIFE化、共通ローカライズ L 関数と LABELS のカテゴリ構造化、共通UIレイアウト setupPanel、未使用コード削除、変数・関数名の明確化）。挙動は v1.5 と同等。
- v1.5 (20250818) : 反転（上下／左右）の対応を追加
- v1.4 (20250817) : クリップグループの縦横比（配置画像＋マスクパスへ同時適用）を実装。複合パス・ネストに対応し、変形の確実性を向上。
- v1.3 (20250817) : 配置/ラスタの［縦横比］オプションを追加（100%/100% 基準の等比化、整数％丸め）。
- v1.2 (20250816) : スケール％入力とホットキー、選択内容によるパネル自動ディム、安定化（atan2・EPS 集約）。
- v1.1 (20250810) : クリップグループ回転の子要素処理、2カラムUI、最大面積代表子選定。
- v1.0 (20250805) : 初期バージョン（基本機能、ダイアログ、配置/ラスタ/テキスト/長方形対応）。

*/

/*

### Script Name:

ResetTransform.jsx (Reset Rotate / Shear / Scale / Aspect Ratio)

### Overview:

- Updated: 2026-07-08
- Safely resets rotation / shear / scale / aspect ratio for Placed Images, Text, Rectangles (paths), Clip Groups, and straight Paths.
- Keeps visual position stable via BBox reset and top-left recentering.

### Key Features:

- **Placed/Raster**: Reset rotation, shear, aspect ratio (lift smaller axis to the larger; integer-rounded percent), and absolute scale, individually or together.
- **Text**: Reset rotation, shear, and horizontal/vertical scaling ratios (100%).
- **Rectangle (4-point path)**: Snap near angles to 0°/90°.
- **Line (2-point path)**: Snap to nearest axis (0°/90°).
- **Clip Group**: For children (placed/raster **and** clipping path), reset rotation and equalize aspect ratio; the uniformizing delta derived from the placed image is applied **to both** the placed image and the clipping path simultaneously. Supports nested groups and compound paths.
- **UI**: Two-column panels per target, hotkeys, numeric nudge (↑↓/Shift+↑↓), dialog position/opacity.

### Flow:

1) Analyze selection → dim unsupported panels.
2) Choose operations → OK.
3) Apply transforms per handler → reset BBox → recenter to top-left.

- Scale input supports Arrow keys: ↑↓ ±1, Shift+↑↓ snaps by 10, Alt/Option+↑↓ adjusts by ±0.1.
- Angle snap tolerance configurable via `CONFIG.rectSnapMin` / `CONFIG.rectSnapMax`.

### Changelog:

- v1.6.0 (20260708): Removed 0.1 decimal step on Up/Down keys (integer only). Scale % minimum changed to 20%. Rectangle rotation correction widened to 44° with BBox reset, original selection restored after run, and a guard for when no document is open.
- v1.5.1 (20260708): Internal cleanup (IIFE wrapper, shared localization via L() and categorized LABELS, shared panel layout via setupPanel, dead-code removal, clearer variable/function names). Behavior unchanged from v1.5.
- v1.5 (20250818): Added Flip (horizontal/vertical) support.
- v1.4 (20250817): Implemented Clip Group aspect-ratio equalization (applies same delta to placed image + clipping path). Added nested/compound support and improved robustness.
- v1.3 (20250817): Added Placed/Raster "Aspect Ratio" option (equalization with integer rounding).
- v1.2 (20250816): Added Scale % input & hotkeys, selection-aware panel dimming, stability improvements (atan2 & EPS consolidation).
- v1.1 (20250810): Clip group child-wise rotation, two-column UI, largest-area child picking.
- v1.0 (20250805): Initial release (core features, dialog, placed/raster/text/rect).

*/

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.6.0";

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================
    /* 角度スナップ許容範囲（度）と行列計算の微小値 / Angle snap tolerance (deg) and matrix epsilon */
    var CONFIG = {
        rectSnapMin: 0.5, // degrees
        rectSnapMax: 44, // degrees
        eps: 1e-8 // numerical epsilon for matrix ops
    };

    /* スケール入力の下限（%）/ Minimum scale percent allowed in the input */
    var SCALE_MIN = 20;

    // --- Dialog position memory helpers (shared across scripts via targetengine) ---
    function _getSavedLoc(key) {
        return $.global[key] && $.global[key].length === 2 ? $.global[key] : null;
    }

    function _setSavedLoc(key, loc) {
        $.global[key] = [loc[0], loc[1]];
    }

    function _clampToScreen(loc) {
        try {
            var vb = ($.screens && $.screens.length) ? $.screens[0].visibleBounds : [0, 0, 1920, 1080];
            var x = Math.max(vb[0] + 10, Math.min(loc[0], vb[2] - 10));
            var y = Math.max(vb[1] + 10, Math.min(loc[1], vb[3] - 10));
            return [x, y];
        } catch (e) {
            return loc;
        }
    }
    // A unique storage key for this script's main dialog
    var DLG_STORE_KEY = "__ResetTransform_OptionsDialog";

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /* 言語判定：日本語環境なら "ja"、それ以外は "en" / Detect UI language */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

    /* ラベル取得（{slash} を "/" に展開）/ Resolve a localized label (expands {slash} to "/") */
    function L(entry) {
        var text = (entry && entry[currentLanguage] != null) ? entry[currentLanguage] : (entry ? entry.en : "");
        return String(text).replace(/\{slash\}/g, "/");
    }

    /* UIラベル（カテゴリ別） / UI labels grouped by category */
    var LABELS = {
        dialog: {
            title: { ja: "リセット（回転・比率）", en: "Reset (Rotate / Scale)" }
        },
        panel: {
            placed: { ja: "配置画像", en: "Placed Images" },
            clip: { ja: "クリップグループ", en: "Clip Group" },
            text: { ja: "テキスト", en: "Text" },
            rect: { ja: "長方形（パス）", en: "Rectangle (Path)" },
            line: { ja: "パス（直線）", en: "Path (Line)" }
        },
        checkbox: {
            rotate: { ja: "回転", en: "Rotate" },
            shear: { ja: "シアー", en: "Shear" },
            ratio: { ja: "縦横比", en: "Aspect Ratio" },
            flip: { ja: "反転", en: "Flip" },
            scale: { ja: "スケール", en: "Scale" },
            textRatio: { ja: "垂直比率／水平比率", en: "Horizontal & Vertical Scale" }
        },
        button: {
            reset: { ja: "リセット", en: "Reset" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            selectFirst: { ja: "オブジェクトを選択してください。", en: "Please select an object." },
            noTarget: { ja: "リセットできる対象が選択されていません。", en: "No resettable objects are selected." }
        }
    };

    // =========================================
    // 単位 / Units
    // =========================================
    /* このスクリプトはルーラー単位換算を使用しません（スケールは % 指定）/ No ruler-unit conversion (scale is percent-based) */

    /* ダイアログ位置・透明度のヘルパー / Dialog position & opacity helpers */
    var DIALOG_OFFSET_X = 300; // shift in pixels (right = positive)
    var DIALOG_OFFSET_Y = 0; // shift in pixels (down = positive)
    var DIALOG_OPACITY = 0.98; // 0.0 – 1.0

    function setDialogOpacity(mainDialog, opacityValue) {
        try {
            mainDialog.opacity = opacityValue;
        } catch (e) { }
    }

    /* Analyze selection → capabilities / 選択状態から対応可否を判定 */
    function getSelectionCapabilities(sel) {
        var selectionCaps = {
            placedOrRaster: false,
            clipGroup: false,
            text: false,
            rect: false,
            line: false
        };
        if (!sel || !sel.length) return selectionCaps;
        for (var i = 0; i < sel.length; i++) {
            var it = sel[i];
            if (!it || !it.typename) continue;
            var t = it.typename;
            if (t === 'PlacedItem' || t === 'RasterItem') selectionCaps.placedOrRaster = true;
            else if (t === 'GroupItem' && it.clipped === true) selectionCaps.clipGroup = true;
            else if (t === 'TextFrame') selectionCaps.text = true;
            else if (t === 'PathItem') {
                if (it.closed && it.pathPoints && it.pathPoints.length === 4) selectionCaps.rect = true;
                if (!it.closed && it.pathPoints && it.pathPoints.length === 2) selectionCaps.line = true;
            }
        }
        return selectionCaps;
    }

    function changeValueByArrowKey(editText, minValue) {
        editText.addEventListener("keydown", function (event) {
            var value = Number(editText.text);
            if (isNaN(value)) return;

            // 修飾キーは event を優先（macOS では keyboardState が誤報する）/ Read modifiers from event first
            var keyboard = ScriptUI.environment.keyboardState;
            var shiftDown = (typeof event.shiftKey === 'boolean') ? event.shiftKey : keyboard.shiftKey;
            var delta = shiftDown ? 10 : 1;
            var floorValue = (typeof minValue === 'number') ? minValue : 0;

            if (shiftDown) {
                // Shiftキー押下時は10の倍数にスナップ / snap to multiples of 10
                if (event.keyName == "Up") {
                    value = Math.ceil((value + 1) / delta) * delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value = Math.round((value - 1) / delta) * delta;
                    event.preventDefault();
                }
            } else {
                if (event.keyName == "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value -= delta;
                    event.preventDefault();
                }
            }

            // 整数に丸めて下限でクランプ / round to integer and clamp to the minimum
            value = Math.round(value);
            if (value < floorValue) value = floorValue;
            editText.text = value;
        });
    }

    /* Hotkey → toggle checkbox / ホットキーでチェックボックスON/OFF */
    function addHotkeyToggle(dialog, keyChar, checkbox, onToggle) {
        dialog.addEventListener('keydown', function (event) {
            var k = (event.keyName || '').toUpperCase();
            if (k === String(keyChar).toUpperCase()) {
                if (!checkbox.enabled) {
                    event.preventDefault();
                    return;
                }
                checkbox.value = !checkbox.value;
                if (typeof onToggle === 'function') onToggle(checkbox.value);
                event.preventDefault();
            }
        });
    }

    /* 入力欄にフォーカスして全選択（ScriptUI差異を吸収）/ Focus an edittext and select all (across ScriptUI builds) */
    function focusAndSelectAll(editText) {
        try {
            editText.active = true;
            if (typeof editText.select === 'function') {
                editText.select();
            } else if (typeof editText.textselection !== 'undefined') {
                editText.textselection = editText.text;
            }
        } catch (e) { }
    }

    // ==============================
    // UIレイアウトの共通設定 / Shared UI layout
    // ==============================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING = 8;                  /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */

    /* ウィンドウの共通設定 / Apply shared window layout */
    function setupWindow(win, spacing) {
        win.orientation = "column";
        win.alignChildren = "fill";
        win.margins = WINDOW_MARGINS;
        win.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /* パネルの共通設定 / Apply shared panel layout */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /* 行グループの共通設定（ボタン列など） / Apply a horizontal row group */
    function setupRow(group, alignment, spacing) {
        group.orientation = "row";
        group.alignment = alignment || "left";
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    function showOptionsDialog() {
        var defaultChecks = {
            rotate: true,
            skew: true,
            scale: false
        };

        var currentSelection = null;
        try {
            currentSelection = (app && app.activeDocument) ? app.activeDocument.selection : null;
        } catch (e) { }
        var selectionCaps = getSelectionCapabilities(currentSelection);

        var mainDialog = new Window('dialog', L(LABELS.dialog.title) + ' ' + SCRIPT_VERSION);
        setupWindow(mainDialog);
        setDialogOpacity(mainDialog, DIALOG_OPACITY);

        // Override onShow to restore last position (or apply first-run offset)
        (function () {
            var saved = _getSavedLoc(DLG_STORE_KEY);
            mainDialog.onShow = function () {
                try {
                    if (saved) {
                        mainDialog.location = _clampToScreen(saved);
                    } else {
                        mainDialog.location = [mainDialog.location[0] + DIALOG_OFFSET_X, mainDialog.location[1] + DIALOG_OFFSET_Y];
                    }
                } catch (e) { }
            };
            // Save whenever the dialog is moved
            mainDialog.onMove = function () {
                try {
                    _setSavedLoc(DLG_STORE_KEY, [mainDialog.location[0], mainDialog.location[1]]);
                } catch (e) { }
            };
        })();

        /* Main group: two-column layout / 2カラムレイアウト */
        var columnsGroup = mainDialog.add('group');
        setupRow(columnsGroup, 'fill', COLUMN_SPACING);
        columnsGroup.alignChildren = 'top';

        /* Left column / 左カラム */
        var leftColumn = columnsGroup.add('group');
        leftColumn.orientation = 'column';
        leftColumn.alignChildren = 'left';
        leftColumn.alignment = 'fill';

        /* Right column / 右カラム */
        var rightColumn = columnsGroup.add('group');
        rightColumn.orientation = 'column';
        rightColumn.alignChildren = 'left';
        rightColumn.alignment = 'fill';

        /* Panel: Placed Images / 配置画像 */
        var placedImagePanel = leftColumn.add('panel', undefined, L(LABELS.panel.placed));
        setupPanel(placedImagePanel, 6);
        var placedRotateCheck = placedImagePanel.add('checkbox', undefined, L(LABELS.checkbox.rotate));
        var placedShearCheck = placedImagePanel.add('checkbox', undefined, L(LABELS.checkbox.shear));
        var placedRatioCheck = placedImagePanel.add('checkbox', undefined, L(LABELS.checkbox.ratio));
        var placedFlipCheck = placedImagePanel.add('checkbox', undefined, L(LABELS.checkbox.flip));
        var placedScaleCheck = placedImagePanel.add('checkbox', undefined, L(LABELS.checkbox.scale));
        placedRotateCheck.value = (defaultChecks.rotate !== false);
        placedShearCheck.value = (defaultChecks.skew !== false);
        placedRatioCheck.value = true;
        placedFlipCheck.value = true;
        placedScaleCheck.value = (defaultChecks.scale !== false);
        // Add scale percent field + % label after placedScaleCheck
        var scaleInputRow = placedImagePanel.add('group');
        scaleInputRow.orientation = 'row';
        scaleInputRow.alignChildren = 'center';
        scaleInputRow.alignment = 'left'; /* 入力欄はパネル幅いっぱいに広げない / keep scale input compact */

        var scalePercentInput = scaleInputRow.add('edittext', undefined, '100');
        scalePercentInput.characters = 5; // default width

        var percentLabel = scaleInputRow.add('statictext', undefined, '%');

        // Enable/disable with the checkbox
        scalePercentInput.enabled = placedScaleCheck.value;
        percentLabel.enabled = placedScaleCheck.value;
        placedScaleCheck.onClick = function () {
            var on = placedScaleCheck.value;
            scalePercentInput.enabled = on;
            percentLabel.enabled = on;
            if (on) focusAndSelectAll(scalePercentInput);
        };
        // Arrow-key increment/decrement support
        changeValueByArrowKey(scalePercentInput, SCALE_MIN);

        // Enable state based on selection capabilities
        placedImagePanel.enabled = selectionCaps.placedOrRaster;
        placedRotateCheck.enabled = selectionCaps.placedOrRaster;
        placedShearCheck.enabled = selectionCaps.placedOrRaster;
        placedRatioCheck.enabled = selectionCaps.placedOrRaster;
        placedFlipCheck.enabled = selectionCaps.placedOrRaster;
        placedScaleCheck.enabled = selectionCaps.placedOrRaster;
        scalePercentInput.enabled = selectionCaps.placedOrRaster && placedScaleCheck.value;
        percentLabel.enabled = selectionCaps.placedOrRaster && placedScaleCheck.value;

        // Hotkey: 'S' toggles Scale / SキーでスケールON/OFF
        addHotkeyToggle(mainDialog, 'S', placedScaleCheck, function (enabled) {
            scalePercentInput.enabled = enabled;
            percentLabel.enabled = enabled;
            if (enabled) focusAndSelectAll(scalePercentInput);
        });
        // Hotkey: 'F' toggles Flip / Fキーで反転ON/OFF
        addHotkeyToggle(mainDialog, 'F', placedFlipCheck);

        /* Panel: Clip Group / クリップグループ */
        var clipGroupPanel = leftColumn.add('panel', undefined, L(LABELS.panel.clip));
        setupPanel(clipGroupPanel, 6);
        var clipRotateCheck = clipGroupPanel.add('checkbox', undefined, L(LABELS.checkbox.rotate));
        var clipRatioCheck = clipGroupPanel.add('checkbox', undefined, L(LABELS.checkbox.ratio));
        var clipFlipCheck = clipGroupPanel.add('checkbox', undefined, L(LABELS.checkbox.flip));
        clipRotateCheck.value = true; // default ON
        clipRatioCheck.value = true; // default ON
        clipFlipCheck.value = true; // default ON

        clipGroupPanel.enabled = selectionCaps.clipGroup;
        clipRotateCheck.enabled = selectionCaps.clipGroup;
        clipRatioCheck.enabled = selectionCaps.clipGroup;
        clipFlipCheck.enabled = selectionCaps.clipGroup;

        /* Panel: Text / テキスト */
        var textPanel = rightColumn.add('panel', undefined, L(LABELS.panel.text));
        setupPanel(textPanel, 6);
        var textRotateCheck = textPanel.add('checkbox', undefined, L(LABELS.checkbox.rotate));
        var textShearCheck = textPanel.add('checkbox', undefined, L(LABELS.checkbox.shear));
        var textScaleRatioCheck = textPanel.add('checkbox', undefined, L(LABELS.checkbox.textRatio));
        // defaults: ON unless explicitly set to false
        textRotateCheck.value = true;
        textShearCheck.value = true;
        textScaleRatioCheck.value = true;
        textPanel.enabled = selectionCaps.text;
        textRotateCheck.enabled = selectionCaps.text;
        textShearCheck.enabled = selectionCaps.text;
        textScaleRatioCheck.enabled = selectionCaps.text;

        /* Panel: Rectangle (Path) / 長方形（パス） */
        var rectanglePanel = rightColumn.add('panel', undefined, L(LABELS.panel.rect));
        setupPanel(rectanglePanel, 6);
        var rectRotateCheck = rectanglePanel.add('checkbox', undefined, L(LABELS.checkbox.rotate));
        // default ON
        rectRotateCheck.value = true;
        rectanglePanel.enabled = selectionCaps.rect;
        rectRotateCheck.enabled = selectionCaps.rect;

        /* Panel: Path (Line) / パス（直線） */
        var linePanel = rightColumn.add('panel', undefined, L(LABELS.panel.line));
        setupPanel(linePanel, 6);
        var lineRotateCheck = linePanel.add('checkbox', undefined, L(LABELS.checkbox.rotate));
        lineRotateCheck.value = true; // default ON
        linePanel.enabled = selectionCaps.line;
        lineRotateCheck.enabled = selectionCaps.line;

        /* Dialog buttons / ダイアログボタン */
        var buttonRow = mainDialog.add('group');
        setupRow(buttonRow, 'center');
        buttonRow.add('button', undefined, L(LABELS.button.cancel), {
            name: 'cancel'
        });
        buttonRow.add('button', undefined, L(LABELS.button.reset), {
            name: 'ok'
        });

        // Persist position on button clicks (OK / Cancel)
        try {
            var cancelButton = buttonRow.children[0];
            var resetButton = buttonRow.children[1];
            resetButton.onClick = function () {
                try {
                    _setSavedLoc(DLG_STORE_KEY, [mainDialog.location[0], mainDialog.location[1]]);
                } catch (e) { }
                mainDialog.close(1);
            };
            cancelButton.onClick = function () {
                try {
                    _setSavedLoc(DLG_STORE_KEY, [mainDialog.location[0], mainDialog.location[1]]);
                } catch (e) { }
                mainDialog.close(0);
            };
        } catch (e) { }

        if (mainDialog.show() !== 1) return null; // cancelled

        var result = {
            rotate: placedRotateCheck.value,
            skew: placedShearCheck.value,
            ratio: placedRatioCheck.value,
            flip: placedFlipCheck.value,
            scale: placedScaleCheck.value,
            textRotate: textRotateCheck.value,
            textSkew: textShearCheck.value,
            textScaleRatio: textScaleRatioCheck.value,
            rectRotate: rectRotateCheck.value,
            lineRotate: lineRotateCheck.value,
            clipGroupRotate: clipRotateCheck.value,
            clipGroupRatio: clipRatioCheck.value,
            clipGroupFlip: clipFlipCheck.value,
            scalePercent: (function () {
                var n = parseFloat(scalePercentInput.text);
                if (isNaN(n)) n = 100;
                n = Math.round(n); // 整数％に統一（例：16.3 → 16）
                if (n < SCALE_MIN) n = SCALE_MIN; // 下限 20%（それ未満は 20% にクランプ）
                return n;
            })()
        };
        return result;
    }

    /* Dispatch table / ディスパッチテーブル：typename → handler */
    function makeHandlers(opts) {
        function makeCounts(p, s, pl, ra) {
            return {
                processed: p | 0,
                skipped: s | 0,
                countPlaced: pl | 0,
                countRaster: ra | 0
            };
        }

        function handleText(item) {
            var did = false;
            var needBBox = !!(opts.textRotate || opts.textSkew);
            if (needBBox) {
                resetTextOps(item, !!opts.textRotate, !!opts.textSkew, !!opts.textScaleRatio);
                did = true;
            } else if (opts.textScaleRatio) {
                // 比率のみならBBox不要で軽量処理
                resetTextScaleRatio(item);
                did = true;
            }
            return did ? makeCounts(1, 0, 0, 0) : makeCounts(0, 1, 0, 0);
        }

        function handleRect(item) {
            // PathItem（4点の長方形）：近傍角度（0.5〜44°）だけ回転補正。すでに正立なら無変更（＝正常）でも対象として処理済み扱い。
            if (opts.rectRotate && item.closed && item.pathPoints.length === 4) {
                applyRotationCorrection(item);
                return makeCounts(1, 0, 0, 0);
            }
            return makeCounts(0, 1, 0, 0);
        }

        function handleLine(item) {
            // 直線（2点）：近傍軸だけ回転補正。無変更でも対象として処理済み扱い。
            if (opts.lineRotate && !item.closed && item.pathPoints.length === 2) {
                applyRotationCorrectionLine(item);
                return makeCounts(1, 0, 0, 0);
            }
            return makeCounts(0, 1, 0, 0);
        }

        function handleClipGroup(item) {
            if (item.clipped !== true) return makeCounts(0, 1, 0, 0);

            // 1) 回転を正す / Fix rotation first
            if (opts.clipGroupRotate) processClippedGroupChildren(item, opts);
            // 1.5) 反転検出の前に BBox リセット / Reset BBox before flip detection
            if (opts.clipGroupRotate && opts.clipGroupFlip) resetBBoxOnly(item);
            // 2) 反転を調整 / Then handle flip
            if (opts.clipGroupFlip) processClippedGroupFlip(item);
            // 3) 縦横比の等比化 / Aspect ratio
            if (opts.clipGroupRatio) processClippedGroupAspect(item);

            // 対象オプションが1つでもONなら、無変更でも処理済み扱い（誤アラート防止）
            var handled = (opts.clipGroupRotate || opts.clipGroupFlip || opts.clipGroupRatio);
            return handled ? makeCounts(1, 0, 0, 0) : makeCounts(0, 1, 0, 0);
        }

        function handleImage(item, typename) {
            resetPlacedOrRasterTransforms(item, typename, opts);
            // Placed/Raster は常に処理成功としてカウント（従来仕様）
            if (typename === 'PlacedItem') return makeCounts(1, 0, 1, 0);
            if (typename === 'RasterItem') return makeCounts(1, 0, 0, 1);
            return makeCounts(1, 0, 0, 0);
        }

        return {
            TextFrame: handleText,
            PathItem: function (item) {
                if (!item.closed && item.pathPoints && item.pathPoints.length === 2) {
                    return handleLine(item); // 直線（2点）なら直線用ハンドラ
                }
                return handleRect(item); // それ以外は長方形（4点）など既存処理
            },
            GroupItem: handleClipGroup,
            PlacedItem: function (item) {
                return handleImage(item, 'PlacedItem');
            },
            RasterItem: function (item) {
                return handleImage(item, 'RasterItem');
            }
        };
    }

    function main() {
        // ドキュメント未オープンなら中断 / Abort if no document is open
        if (!app.documents.length) {
            alert(L(LABELS.alert.noDocument));
            return;
        }

        var currentDocument = app.activeDocument;
        var liveSelection = currentDocument.selection;
        if (!liveSelection || liveSelection.length === 0) {
            alert(L(LABELS.alert.selectFirst));
            return;
        }

        // 開始時の選択を控える（処理中に app.selection を張り替えるため）/ Snapshot the selection to restore later
        var originalSelection = [];
        for (var s = 0; s < liveSelection.length; s++) originalSelection.push(liveSelection[s]);

        var opts = showOptionsDialog();
        if (!opts) return;

        var processed = 0;
        var handlers = makeHandlers(opts);

        for (var i = 0; i < originalSelection.length; i++) {
            var item = originalSelection[i];
            if (!item || !item.typename) continue;
            var handler = handlers[item.typename];
            if (handler) processed += handler(item).processed || 0;
        }

        // 元の選択に戻す / Restore the original selection
        try {
            currentDocument.selection = originalSelection;
        } catch (e) {}

        if (processed === 0) {
            alert(L(LABELS.alert.noTarget));
        }
    }

    /* Helpers / ヘルパー群 */

    function withBBoxResetAndRecenter(item, opFn) {
        /* Run transform → reset BBox → recenter to keep top-left / 変形→BBoxリセット→左上基準で再配置 */
        var tl = item.position;
        var w1 = item.width,
            h1 = item.height;
        if (typeof opFn === 'function') opFn();
        try {
            app.selection = null;
            app.selection = [item];
            app.executeMenuCommand("AI Reset Bounding Box");
        } catch (e) { }
        recenterToTopLeft(item, tl, w1, h1);
    }

    /* Helper: calculate "safe" area (width * height), or -1 if not available */
    function getAreaSafe(it) {
        try {
            var w = it.width;
            var h = it.height;
            if (typeof w !== 'number' || typeof h !== 'number') return -1;
            return Math.abs(w * h);
        } catch (e) {
            return -1;
        }
    }

    /* Clip group: child-wise reset (rotate only) / クリップグループ：子要素単位の回転リセット */
    function processClippedGroupChildren(groupItem, opts) {
        // Ensure BBox is up-to-date before reading matrices
        resetBBoxOnly(groupItem);

        // Recursively find representative placed/raster and clipping path
        var found = findPlacedAndClipRecursive(groupItem);
        var placedOrRaster = found.placed;
        var clipPath = found.clip;
        var clipTarget = resolveClipTransformTarget(clipPath);

        // Compute a shared rotation delta from the representative child
        var rotDelta = 0;
        var haveDelta = false;
        if (placedOrRaster && hasMatrix(placedOrRaster)) {
            var pm = placedOrRaster.matrix;
            var psign = (placedOrRaster.typename === 'RasterItem') ? -1 : 1;
            rotDelta = getRotationAngleDeg(pm.mValueA, pm.mValueB, psign);
            haveDelta = Math.abs(rotDelta) > 0.0001;
        } else if (clipTarget && hasMatrix(clipTarget)) {
            var cm = clipTarget.matrix;
            rotDelta = getRotationAngleDeg(cm.mValueA, cm.mValueB, +1);
            haveDelta = Math.abs(rotDelta) > 0.0001;
        }

        var did = false;
        if (opts.clipGroupRotate && haveDelta) {
            // Rotate the GROUP once so children keep their relative alignment
            withBBoxResetAndRecenter(groupItem, function () {
                rotateBy(groupItem, rotDelta);
            });
            did = true;
        }
        return did;
    }

    // Recursively find the largest placed/raster and the clipping path (PathItem or CompoundPathItem) within a clipped group
    function findPlacedAndClipRecursive(container) {
        var bestPlaced = null,
            bestPlacedArea = -1;
        var bestClip = null,
            bestClipArea = -1;
        var items = (container.pageItems || []);
        for (var i = 0; i < items.length; i++) {
            var it = items[i];
            if (!it) continue;
            var t = it.typename;
            if (t === 'PlacedItem' || t === 'RasterItem') {
                var ap = getAreaSafe(it);
                if (ap > bestPlacedArea) {
                    bestPlaced = it;
                    bestPlacedArea = ap;
                }
            } else if (t === 'PathItem') {
                if (it.clipping) {
                    var ac1 = getAreaSafe(it);
                    if (ac1 > bestClipArea) {
                        bestClip = it;
                        bestClipArea = ac1;
                    }
                }
            } else if (t === 'CompoundPathItem') {
                // CompoundPathItem itself may not expose `.clipping`; check its children
                var subPaths = it.pathItems || [];
                for (var k = 0; k < subPaths.length; k++) {
                    var sp = subPaths[k];
                    try {
                        if (sp && sp.clipping) {
                            var ac2 = getAreaSafe(it); // use compound's area as proxy
                            if (ac2 > bestClipArea) {
                                bestClip = it;
                                bestClipArea = ac2;
                            }
                            break;
                        }
                    } catch (e) { }
                }
            } else if (t === 'GroupItem') {
                var found = findPlacedAndClipRecursive(it);
                if (found.placed && getAreaSafe(found.placed) > bestPlacedArea) {
                    bestPlaced = found.placed;
                    bestPlacedArea = getAreaSafe(found.placed);
                }
                if (found.clip && getAreaSafe(found.clip) > bestClipArea) {
                    bestClip = found.clip;
                    bestClipArea = getAreaSafe(found.clip);
                }
            }
        }
        return {
            placed: bestPlaced,
            clip: bestClip
        };
    }

    // Apply transform with full flags and center origin; fallback safely if enum not available
    function transformAll(item, mat) {
        try {
            item.transform(mat, true, true, true, true, true, Transformation.CENTER);
        } catch (e) {
            try {
                item.transform(mat, true, true, true, true, true);
            } catch (e2) {
                item.transform(mat);
            }
        }
    }
    // Temporarily unlock/unhide while applying a function, then restore
    function withUnlockedVisible(item, fn) {
        var wasLocked = false,
            wasHidden = false;
        try {
            wasLocked = !!item.locked;
        } catch (e) { }
        try {
            wasHidden = !!item.hidden;
        } catch (e) { }
        try {
            try {
                item.locked = false;
            } catch (e) { }
            try {
                item.hidden = false;
            } catch (e) { }
            fn();
        } finally {
            try {
                if (wasLocked) item.locked = true;
            } catch (e) { }
            try {
                if (wasHidden) item.hidden = true;
            } catch (e) { }
        }
    }

    // Reset only the bounding box for a given item (no recenter) / バウンディングボックスのみをリセット（位置は維持）
    function resetBBoxOnly(item) {
        try {
            app.selection = null;
            app.selection = [item];
            app.executeMenuCommand("AI Reset Bounding Box");
        } catch (e) { }
    }

    // Resolve actual transform target for a clipping path: if CompoundPathItem, use its clipping PathItem
    function resolveClipTransformTarget(clipCandidate) {
        if (!clipCandidate) return null;
        try {
            if (clipCandidate.typename === 'CompoundPathItem') {
                var subs = clipCandidate.pathItems || [];
                for (var i = 0; i < subs.length; i++) {
                    var sp = subs[i];
                    try {
                        if (sp && sp.clipping) return sp; // prefer the actual clipping child
                    } catch (e) { }
                }
                // fallback: transform the compound itself if no child flagged
                return clipCandidate;
            }
            return clipCandidate; // PathItem or others
        } catch (e) {
            return clipCandidate;
        }
    }

    // Clip group: aspect ratio equalization for placed/raster & clipping path
    // クリップグループ：配置画像＋マスクパス（PathItem/CompoundPathItem）に再帰的に同じ等比化を適用
    function processClippedGroupAspect(groupItem) {
        // Find placed/raster and (PathItem or CompoundPathItem) clipping path recursively
        var found = findPlacedAndClipRecursive(groupItem);
        var placedOrRaster = found.placed;
        var clipPath = found.clip;
        var clipTarget = resolveClipTransformTarget(clipPath);
        // Proceed if we have at least the placed/raster; clipPath is optional
        if (!placedOrRaster) return false;
        if (!hasMatrix(placedOrRaster)) return false; // needed to compute delta from placed image
        var haveClip = !!clipTarget; // we can still transform a clip even if it lacks `.matrix`
        var clipHasMatrix = !!(haveClip && hasMatrix(clipTarget));

        // Build target uniform scale from placed image: lift smaller axis to larger, snap to integer percent
        var pm = placedOrRaster.matrix;
        var dec = decomposeQR(pm.mValueA, pm.mValueB, pm.mValueC, pm.mValueD);
        var u = Math.max(dec.sx, dec.sy);
        var uPercent = Math.round(u * 100); // round to integer percent
        u = uPercent / 100;

        // Delta matrix from current to (u,u) with same orientation/shear
        var target = buildFromQR(dec.q1x, dec.q1y, dec.q2x, dec.q2y, u, u, dec.shear);
        var cur = {
            a: pm.mValueA,
            b: pm.mValueB,
            c: pm.mValueC,
            d: pm.mValueD
        };
        var inv = invert2x2(cur.a, cur.b, cur.c, cur.d);
        var delta = multiply2x2(inv.a, inv.b, inv.c, inv.d, target.a, target.b, target.c, target.d);
        var deltaMat = toMatrix(delta);

        // Always apply to placed/raster; also apply to clipTarget if present (even without `.matrix`)
        withUnlockedVisible(placedOrRaster, function () {
            transformAll(placedOrRaster, deltaMat);
        });
        if (haveClip) {
            withUnlockedVisible(clipTarget, function () {
                transformAll(clipTarget, deltaMat);
            });
        }

        // Verify and fallback per-item to guarantee uniform scale
        try {
            var mPl = placedOrRaster.matrix,
                dPl = decomposeQR(mPl.mValueA, mPl.mValueB, mPl.mValueC, mPl.mValueD);
            if (Math.abs(dPl.sx - dPl.sy) > 1e-6) {
                withUnlockedVisible(placedOrRaster, function () {
                    equalizeScaleToMax(placedOrRaster);
                });
            }
            if (clipHasMatrix) {
                var mCp = clipTarget.matrix,
                    dCp = decomposeQR(mCp.mValueA, mCp.mValueB, mCp.mValueC, mCp.mValueD);
                if (Math.abs(dCp.sx - dCp.sy) > 1e-6) {
                    withUnlockedVisible(clipTarget, function () {
                        equalizeScaleToMax(clipTarget);
                    });
                }
            }
        } catch (e) { }
        return true;
    }

    function getRotationAngleDeg(a, b, sign) {
        var ang = Math.atan2(b, a) * 180 / Math.PI;
        return (sign < 0) ? -ang : ang;
    }

    /* Rotation matrix helper (uses Illustrator API if available) / 回転行列ヘルパー（可能ならIllustrator標準APIを使用） */
    function getRotationMatrixSafe(deg) {
        // Prefer Illustrator's API for clarity
        try {
            if (app && typeof app.getRotationMatrix === 'function') {
                return app.getRotationMatrix(deg);
            }
        } catch (e) { }
        // Fallback: build a Matrix manually
        var rad = deg * Math.PI / 180.0;
        var cosv = Math.cos(rad),
            sinv = Math.sin(rad);
        var M = new Matrix();
        M.mValueA = cosv; // a
        M.mValueB = sinv; // b
        M.mValueC = -sinv; // c
        M.mValueD = cosv; // d
        M.mValueTX = 0;
        M.mValueTY = 0;
        return M;
    }

    /* Rotate item by degrees using Illustrator matrix API / Illustratorの回転行列APIを用いて回転 */
    function rotateBy(item, deg) {
        var mat = getRotationMatrixSafe(deg);
        item.transform(mat);
    }

    // Rotate TextFrame by deg, prefer Illustrator's native rotate with full flags and center origin
    function rotateTextBy(item, deg) {
        try {
            // Use Illustrator's native rotate for TextFrame with full flags and center origin
            item.rotate(deg, true, true, true, true, Transformation.CENTER);
        } catch (e) {
            // Fallback to matrix-based rotation
            rotateBy(item, deg);
        }
    }

    function multiply2x2(a1, b1, c1, d1, a2, b2, c2, d2) {
        return {
            a: a1 * a2 + c1 * b2,
            b: b1 * a2 + d1 * b2,
            c: a1 * c2 + c1 * d2,
            d: b1 * c2 + d1 * d2
        };
    }

    function invert2x2(a, b, c, d) {
        var det = a * d - b * c;
        if (Math.abs(det) < CONFIG.eps) det = (det < 0 ? -1 : 1) * CONFIG.eps;
        var invDet = 1.0 / det;
        return {
            a: d * invDet,
            b: -b * invDet,
            c: -c * invDet,
            d: a * invDet
        };
    }

    function toMatrix(obj) {
        var M = new Matrix();
        M.mValueA = obj.a;
        M.mValueB = obj.b;
        M.mValueC = obj.c;
        M.mValueD = obj.d;
        M.mValueTX = 0;
        M.mValueTY = 0;
        return M;
    }

    function decomposeQR(a, b, c, d) {
        var sx = Math.sqrt(a * a + b * b);
        if (sx === 0) sx = CONFIG.eps;
        var q1x = a / sx,
            q1y = b / sx;
        var r12 = q1x * c + q1y * d;
        var u2x = c - r12 * q1x;
        var u2y = d - r12 * q1y;
        var sy = Math.sqrt(u2x * u2x + u2y * u2y);
        if (sy === 0) {
            sy = CONFIG.eps;
            u2x = -q1y;
            u2y = q1x;
        }
        var q2x = u2x / sy,
            q2y = u2y / sy;
        var shear = r12 / sx;
        return {
            sx: sx,
            sy: sy,
            shear: shear,
            q1x: q1x,
            q1y: q1y,
            q2x: q2x,
            q2y: q2y
        };
    }

    function buildFromQR(q1x, q1y, q2x, q2y, sx, sy, shear) {
        var r11 = sx,
            r12 = shear * sx,
            r21 = 0,
            r22 = sy;
        return {
            a: q1x * r11 + q2x * r21,
            b: q1y * r11 + q2y * r21,
            c: q1x * r12 + q2x * r22,
            d: q1y * r12 + q2y * r22
        };
    }

    function applyDeltaToMatch(item, target2x2) {
        var m = item.matrix;
        var cur = {
            a: m.mValueA,
            b: m.mValueB,
            c: m.mValueC,
            d: m.mValueD
        };
        var inv = invert2x2(cur.a, cur.b, cur.c, cur.d);
        var delta = multiply2x2(inv.a, inv.b, inv.c, inv.d, target2x2.a, target2x2.b, target2x2.c, target2x2.d);
        item.transform(toMatrix(delta));
    }

    function removeSkewOnly(item) {
        if (!item || !hasMatrix(item)) return;
        var m = item.matrix;
        var dec = decomposeQR(m.mValueA, m.mValueB, m.mValueC, m.mValueD);
        var target = buildFromQR(dec.q1x, dec.q1y, dec.q2x, dec.q2y, dec.sx, dec.sy, 0);
        applyDeltaToMatch(item, target);
    }

    /* シアー除去＋数値誤差で残った微小シアーをもう一度除去 / Remove shear, then clear any tiny residual from numeric noise */
    function removeSkewWithSafety(item) {
        removeSkewOnly(item);
        try {
            if (hasMatrix(item)) {
                var m = item.matrix;
                var d = decomposeQR(m.mValueA, m.mValueB, m.mValueC, m.mValueD);
                if (Math.abs(d.shear) > 1e-6) removeSkewOnly(item);
            }
        } catch (e) { }
    }

    function normalizeScaleOnly(item) {
        if (!item || !hasMatrix(item)) return;
        var m = item.matrix;
        var dec = decomposeQR(m.mValueA, m.mValueB, m.mValueC, m.mValueD);
        var target = buildFromQR(dec.q1x, dec.q1y, dec.q2x, dec.q2y, 1, 1, dec.shear);
        applyDeltaToMatch(item, target);
    }

    // Make horizontal/vertical scales equal by lifting the smaller to the larger (preserving orientation/shear)
    function equalizeScaleToMax(item) {
        if (!item || !hasMatrix(item)) return;
        var m = item.matrix;
        var dec = decomposeQR(m.mValueA, m.mValueB, m.mValueC, m.mValueD);
        var u = Math.max(dec.sx, dec.sy); // choose the larger of current scales

        // ▼追加：整数％にスナップ（例：16.321%→16%、16.3%→16%）
        var uPercent = Math.round(u * 100); // 四捨五入（例：16.321%→16%、16.5%→17%）
        u = uPercent / 100;

        var target = buildFromQR(dec.q1x, dec.q1y, dec.q2x, dec.q2y, u, u, dec.shear);
        applyDeltaToMatch(item, target);
    }

    // Apply uniform scaling by percent (100 = 100%)
    function applyUniformScalePercent(item, percent) {
        var p = Number(percent);
        if (!(p > 0)) return;
        try {
            // Illustrator's resize: percent values (100 = 100%)
            item.resize(p, p, true, true, true, true, true);
        } catch (e) { }
    }

    function hasMatrix(obj) {
        try {
            return !!(obj && obj.matrix && typeof obj.matrix.mValueA !== 'undefined');
        } catch (e) {
            return false;
        }
    }

    // --- Flip detection helpers (no rotation/shear assumed) / 反転検出ヘルパー（回転・シアーなし前提） ---
    function isFlippedHorizontal(mat) {
        // 左右反転: mValueA が負
        return mat && mat.mValueA < 0;
    }
    // NOTE: Placed/Raster はデフォルトで mValueD が負になり得るため、上下反転は mValueD > 0 を基準に判定
    function isFlippedVertical(mat) {
        // 上下反転: mValueD が正（Placed/Raster の基準に合わせる）
        return mat && mat.mValueD > 0;
    }

    /* 反転解除：sx/sy に -100 を与えた軸だけ反転を打ち消す / Undo flip on the axis whose scale is -100 */
    function unflip(item, sx, sy) {
        if (!item) return;
        item.transform(
            app.getScaleMatrix(sx, sy),
            true, // changePositions
            true, // changeFillPatterns
            true, // changeFillGradients
            true, // changeStrokePattern
            true, // changeLineWidths
            Transformation.CENTER
        );
    }

    /* 反転フラグ（左右／上下）に応じて反転を打ち消す / Undo flip according to horizontal/vertical flags */
    function unflipByFlags(item, isHorizontalFlip, isVerticalFlip) {
        if (!item || (!isHorizontalFlip && !isVerticalFlip)) return;
        unflip(item, isHorizontalFlip ? -100 : 100, isVerticalFlip ? -100 : 100);
    }

    function cancelRotation(item, sign) {
        if (!hasMatrix(item)) return 0; // safety: some items may not expose matrix
        var a = item.matrix.mValueA;
        var b = item.matrix.mValueB;
        var rot = getRotationAngleDeg(a, b, sign);
        rotateBy(item, rot);
        return rot;
    }

    // Sign-agnostic cancel: always rotate by the negative of the current angle to zero out rotation
    function cancelRotationToZero(item) {
        if (!hasMatrix(item)) return 0;
        var a = item.matrix.mValueA;
        var b = item.matrix.mValueB;
        var ang = Math.atan2(b, a) * 180 / Math.PI;
        if (item.typename === 'TextFrame') {
            rotateTextBy(item, -ang);
        } else {
            rotateBy(item, -ang);
        }
        return ang;
    }

    function recenterToTopLeft(item, tl, w1, h1) {
        var w2 = item.width,
            h2 = item.height;
        item.position = [tl[0] + w1 / 2 - w2 / 2, tl[1] - h1 / 2 + h2 / 2];
    }

    function resetTextScaleRatio(item) {
        if (!item || !item.textRange) return; // safe guard
        try {
            // Pass 1: primary API
            item.textRange.scaling = [1, 1];
            // Also set legacy attributes for robustness across builds
            var ca = item.textRange.characterAttributes;
            if (ca) {
                try {
                    ca.horizontalScale = 100;
                } catch (e1) { }
                try {
                    ca.verticalScale = 100;
                } catch (e2) { }
            }
            // Pass 2: re-apply to kill residual rounding noise
            item.textRange.scaling = [1, 1];
        } catch (e) { }
    }

    function resetTextOps(item, doRot, doShear, doRatio) {
        // One-shot BBox + recenter for text ops / テキスト処理を1回のBBoxでまとめて実行
        withBBoxResetAndRecenter(item, function () {
            // 1) Rotate first (stabilize orientation for text) / まず回転を0°へ
            if (doRot) cancelRotationToZero(item); // Robust for PointText and AreaText

            // 2) Shear removal (before ratio) / シアー除去を先に
            if (doShear) {
                removeSkewWithSafety(item);
            }

            // 3) Ratio equalization LAST / 比率（水平・垂直）を最後に
            if (doRatio) {
                resetTextScaleRatio(item); // pass 1
                // Best-effort verification and second pass
                try {
                    var sc = item.textRange && item.textRange.scaling;
                    if (!sc || Math.abs(sc[0] - 1) > 1e-6 || Math.abs(sc[1] - 1) > 1e-6) {
                        resetTextScaleRatio(item); // pass 2
                    }
                } catch (e) {
                    // Fallback: second pass anyway
                    resetTextScaleRatio(item);
                }
            }
        });
    }

    function applyRotationCorrection(rect) {
        if (!rect || !rect.closed || rect.pathPoints.length !== 4) return false;

        /* Rotation from first two anchors / 最初の2点から角度を算出 */
        var ptA = rect.pathPoints[0].anchor;
        var ptB = rect.pathPoints[1].anchor;

        var dx = ptB[0] - ptA[0];
        var dy = ptB[1] - ptA[1];
        var angleRad = Math.atan2(dy, dx);
        var angleDeg = angleRad * 180 / Math.PI;
        if (angleDeg < 0) angleDeg += 360;

        var rotationAmount = angleDeg; // 水平からのズレ / offset from horizontal
        var normalized = rotationAmount % 90;
        if (normalized > 45) normalized = 90 - normalized; // 最近傍の 0/90 に正規化 / distance to nearest axis

        // 0.5〜44° の範囲でのみ補正 / only snap if near axis to avoid false positives
        if (normalized >= CONFIG.rectSnapMin && normalized <= CONFIG.rectSnapMax) {
            withBBoxResetAndRecenter(rect, function() {
                rotateBy(rect, -rotationAmount);
            });
            return true;
        }
        return false;
    }

    function applyRotationCorrectionLine(line) {
        if (!line || line.closed || line.pathPoints.length !== 2) return false;

        var ptA = line.pathPoints[0].anchor;
        var ptB = line.pathPoints[1].anchor;

        var dx = ptB[0] - ptA[0];
        var dy = ptB[1] - ptA[1];
        if (dx === 0 && dy === 0) return false;

        // Current angle from horizontal in degrees (-180..180)
        var angleRad = Math.atan2(dy, dx);
        var a = angleRad * 180 / Math.PI;

        // Compute minimal deltas to 0° (horizontal) and 90° (vertical)
        function norm180(x) {
            while (x > 180) x -= 360;
            while (x < -180) x += 360;
            return x;
        }

        function clamp90(x) {
            x = norm180(x);
            if (x > 90) x -= 180;
            if (x < -90) x += 180;
            return x;
        }

        var d0 = clamp90(-a); // rotate by this to reach 0°
        var d90 = clamp90(90 - a); // rotate by this to reach 90°

        var abs0 = Math.abs(d0);
        var abs90 = Math.abs(d90);
        var dist = Math.min(abs0, abs90);

        // Optional: only snap when reasonably close to axis, mirroring rectangle rule
        if (dist < CONFIG.rectSnapMin || dist > CONFIG.rectSnapMax) return false;

        var rot = (abs0 <= abs90) ? d0 : d90;
        withBBoxResetAndRecenter(line, function () {
            rotateBy(line, rot);
        });
        return true;
    }

    /* 反転の解除（回転・シアーなし前提で判定）/ Undo flips (assumes rotation/shear already cleared) */
    function resetFlip(item) {
        if (!hasMatrix(item)) return;
        try {
            var m = item.matrix;
            unflipByFlags(item, isFlippedHorizontal(m), isFlippedVertical(m));
        } catch (e) { }
    }

    /* 配置画像／ラスタの変形リセット（回転→反転→縦横比→スケール→シアーの順）/ Reset a placed/raster item's transforms */
    function resetPlacedOrRasterTransforms(item, objectType, opts) {
        var sign = (objectType === "RasterItem") ? -1 : 1;
        if (!opts) opts = {};
        /* 何も指定されていなければ即終了 / Fast path: nothing to do */
        if (!opts.rotate && !opts.skew && !opts.scale && !opts.ratio && !opts.flip) return;

        withBBoxResetAndRecenter(item, function () {
            /* 1) 回転（向きを安定させるため最初に）/ Rotation first */
            if (opts.rotate && hasMatrix(item)) cancelRotation(item, sign);

            /* 1.5) 反転（回転補正後に判定）/ Flip reset, after rotation cancel */
            if (opts.flip) resetFlip(item);

            /* 2) 縦横比の等比化（スケール正規化の前に）/ Equalize aspect ratio before absolute scaling */
            if (opts.ratio) equalizeScaleToMax(item);

            /* 3) 絶対スケール（100%へ正規化→指定%）/ Absolute scale: normalize to 100% then apply % */
            if (opts.scale) {
                normalizeScaleOnly(item);
                applyUniformScalePercent(item, opts.scalePercent);
            }

            /* 4) シアー除去は最後（上記で混入した微小シアーも除去）/ Shear removal last */
            if (opts.skew) removeSkewWithSafety(item);
        });
        /* 回転ONなら0°で終了、OFFなら元の回転を保持 / Exit at 0° when rotate is ON, else keep original */
    }

    // Clip group: flip reset for placed/raster & clipping path
    // クリップグループ：配置画像＋マスクパスの反転を同時に解除
    function processClippedGroupFlip(groupItem) {
        // Ensure matrices reflect post-rotation state / 回転補正後の行列で正しく判定するためにBBoxを更新
        resetBBoxOnly(groupItem);
        var found = findPlacedAndClipRecursive(groupItem);
        var placedOrRaster = found.placed;
        var clipPath = found.clip;
        var clipTarget = resolveClipTransformTarget(clipPath);
        var haveAny = !!(placedOrRaster || clipTarget);
        if (!haveAny) return false;

        // Prefer detection from placed/raster when available; otherwise fall back to clip
        var mat = null;
        if (placedOrRaster && hasMatrix(placedOrRaster)) {
            mat = placedOrRaster.matrix;
        } else if (clipTarget && hasMatrix(clipTarget)) {
            mat = clipTarget.matrix;
        }
        if (!mat) return false;

        var fh = isFlippedHorizontal(mat);
        var fv = isFlippedVertical(mat);
        if (!fh && !fv) return false;

        // Apply identical unflip to both placed and clip (if present)
        function applyUnflip(it) {
            if (!it) return;
            withUnlockedVisible(it, function () {
                unflipByFlags(it, fh, fv);
            });
        }
        applyUnflip(placedOrRaster);
        applyUnflip(clipTarget);
        return true;
    }

    main();

})();