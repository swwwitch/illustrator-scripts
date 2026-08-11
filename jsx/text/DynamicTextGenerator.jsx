#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

# DynamicTextGenerator.jsx

### 概要

- 選択したテキストからアーチ状のパスを自動生成し、パス上文字に変換します。
- 詳細は README を参照してください。

### Overview

- Generates an arc-shaped path from the selected text and converts it into type on a path.
- See the README for details.

### オリジナル / Original

- 高橋としゆき（@gautt）さん https://note.com/gautt/n/n92f6faeda048

*/

(function () {

    // =========================================
    // 基本情報 / Basic info
    // =========================================
    var SCRIPT_NAME     = "DynamicTextGenerator";         /* スクリプト名 / script name */
    var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
    var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
    var SCRIPT_RELEASED = "2026-05-18";                   /* 最初のリリース日 / first release date */
    var SCRIPT_UPDATED  = "2026-08-11";                   /* 更新日 / last updated */

    // README (Japanese)
    // https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DynamicTextGenerator.md
    // README (English)
    // https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DynamicTextGenerator.md
    var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nbabfda4b7920"; /* 紹介記事 / article URL */

    // Released under the MIT license
    // http://opensource.org/licenses/mit-license.php

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    /* 円モードで文字が円周に占める割合（カーブ0＝ゆるやかな弧／カーブ100＝ほぼ一周） */
    var CIRCLE_MIN_OCCUPANCY = 0.25;
    var CIRCLE_MAX_OCCUPANCY = 0.95;

    /* 言語判定 / Language */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var lang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialogTitle: { ja: "アーチ文字に変換", en: "Convert to Arc Text" },
        panelTitle: { ja: "アーチ文字に変換", en: "Convert to Arc Text" },
        modePanel: { ja: "モード", en: "Mode" },
        arcRoundness: { ja: "カーブ：", en: "Curve:" },
        arcDirectionUp: { ja: "上向き", en: "Up" },
        arcDirectionDown: { ja: "下向き", en: "Down" },
        arcDirectionCircle: { ja: "円", en: "Circle" },
        fit: { ja: "パス幅に合わせる：", en: "Fit to path width:" },
        fitNone: { ja: "しない", en: "None" },
        fitByFontSize: { ja: "文字サイズ", en: "Font size" },
        fitByTracking: { ja: "トラッキング", en: "Tracking" },
        effect: { ja: "効果：", en: "Effect:" },
        effectRainbow: { ja: "虹", en: "Rainbow" },
        effectDistort: { ja: "歪み", en: "Skew" },
        effectRibbon: { ja: "3D リボン", en: "3D Ribbon" },
        effectStep: { ja: "階段", en: "Stair Step" },
        effectGravity: { ja: "引力", en: "Gravity" },
        tracking: { ja: "トラッキング：", en: "Tracking:" },
        preview: { ja: "プレビュー", en: "Preview" },
        ok: { ja: "OK", en: "OK" },
        cancel: { ja: "キャンセル", en: "Cancel" },
        alertNoDoc: { ja: "ドキュメントが開かれていません", en: "No document is open." },
        alertNoText: { ja: "対象のテキストが見つかりません", en: "No target text found." },
        alertArcFail: { ja: "アーチ状のパス生成に失敗しました", en: "Failed to generate an arc path." },
        tipRoundness: { ja: "0で直線に近く、100で最も丸くなります。", en: "0 is almost straight; 100 gives the roundest curve." },
        tipDirection: { ja: "アーチの膨らむ方向を指定します。", en: "Sets the direction in which the arc bulges." },
        tipDirectionCircle: { ja: "閉じた円形のパスに変換し、円周に沿わせます。文字は円の上側中央に配置されます。", en: "Converts to a closed circular path and flows the text around it, centered at the top of the circle." },
        tipFit: { ja: "パスの端まで文字を収める方法を選びます。", en: "Chooses how the text is fitted to the path endpoints." },
        tipFitNone: { ja: "パス幅には合わせず、アーチ変換のみ行います。", en: "Does not fit to the path width; only converts to arc text." },
        tipFitMethod: { ja: "文字サイズを変えて合わせるか、文字サイズを保ったままトラッキングで合わせるかを選びます。", en: "Choose whether to fit by changing the font size, or by keeping the font size and adjusting tracking." },
        tipEffect: { ja: "Illustratorの「パス上文字オプション」の効果を適用します。", en: "Applies Illustrator's Type on a Path effect." },
        tipTracking: { ja: "既存のトラッキング値に加算します。", en: "Adds this value to the existing tracking." },
        tipTrackingToggle: { ja: "ONでトラッキングを調整できます。OFFにすると0に戻ります。", en: "Enable to adjust tracking. Turning it off resets it to 0." },
        tipPreview: { ja: "ONの間は仮の結果を表示します。OFFまたはキャンセルで元に戻ります。", en: "Shows a temporary result while enabled. Turning it off or cancelling restores the original." }
    };

    function L(key) {
        try {
            if (LABELS[key] && LABELS[key][lang]) return LABELS[key][lang];
            if (LABELS[key] && LABELS[key].ja) return LABELS[key].ja;
        } catch (_) { }
        return key;
    }

    /* ===== ユーティリティ / Utilities ===== */

    // Parse a number from a string; return fallback when not numeric
    function parseNumber(str, fallback) {
        var n = Number(str);
        if (isNaN(n)) return fallback;
        return n;
    }

    // Arrow-key increment/decrement for an edittext (Shift = 10, Option = 0.1)
    function changeValueByArrowKey(editText, allowNegative, onChanged) {
        if (!editText) return;
        try {
            editText.addEventListener('keydown', function (event) {
                try {
                    if (!event || (event.keyName !== 'Up' && event.keyName !== 'Down')) return;

                    var value = Number(editText.text);
                    if (isNaN(value)) return;

                    var keyboard = ScriptUI.environment.keyboardState;
                    var delta = 1;

                    if (keyboard.shiftKey) {
                        delta = 10;
                        // Snap to multiples of 10 for Shift
                        if (event.keyName === 'Up') {
                            value = Math.ceil((value + 1) / delta) * delta;
                        } else {
                            value = Math.floor((value - 1) / delta) * delta;
                        }
                    } else if (keyboard.altKey) {
                        delta = 0.1;
                        if (event.keyName === 'Up') value += delta;
                        else value -= delta;
                    } else {
                        delta = 1;
                        if (event.keyName === 'Up') value += delta;
                        else value -= delta;
                    }

                    // Round
                    if (keyboard.altKey) value = Math.round(value * 10) / 10;
                    else value = Math.round(value);

                    // Clamp
                    if (!allowNegative && value < 0) value = 0;

                    editText.text = String(value);

                    // Prevent default arrow key behavior (cursor move)
                    try { event.preventDefault(); } catch (_) { }

                    if (onChanged && typeof onChanged === 'function') {
                        try { onChanged(); } catch (_) { }
                    }
                } catch (_) { }
            });
        } catch (_) { }
    }

    /* ===== 選択の取得 / Selection ===== */
    if (app.documents.length === 0) {
        alert(L('alertNoDoc'));
        return;
    }
    var doc = app.activeDocument;
    var sel = doc.selection;

    // Base selection snapshot (used for stable preview while dialog is open)
    var baseSelection = [];
    try { baseSelection = sel.slice(0); } catch (_) { baseSelection = []; }

    var targetItems = getTargetTextItems(sel);
    var selectedPaths = getSelectedPathItems(sel);

    if (targetItems.length === 0) {
        alert(L('alertNoText'));
        return;
    }

    /* ===== ダイアログ / Dialog ===== */
    var dlg = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);
    dlg.orientation = 'column';
    dlg.alignChildren = ['fill', 'top'];
    dlg.margins = [15, 20, 15, 15];

    // Common width for the leading label of each row (keeps controls aligned)
    var LABEL_COLUMN_WIDTH = 120;

    /**
     * 各行の先頭ラベルの幅を揃え、右寄せにする
     * @param {StaticText} labelText - 対象のラベル
     * @returns {void}
     */
    function setupLabelColumn(labelText) {
        labelText.preferredSize.width = LABEL_COLUMN_WIDTH;
        labelText.justify = 'right';
    }

    /* モードパネル（アーチの向き・円）/ Mode panel (arc direction / circle) */
    var pnlMode = dlg.add('panel', undefined, L('modePanel'));
    pnlMode.orientation = 'row';
    pnlMode.alignChildren = ['left', 'center'];
    pnlMode.margins = [15, 20, 15, 15];

    var rbArcDirectionUp = pnlMode.add('radiobutton', undefined, L('arcDirectionUp'));
    rbArcDirectionUp.helpTip = L('tipDirection');
    var rbArcDirectionDown = pnlMode.add('radiobutton', undefined, L('arcDirectionDown'));
    rbArcDirectionDown.helpTip = L('tipDirection');
    // 円：アーチではなく閉じた円形パスに変換し、円周に沿わせる
    var rbArcDirectionCircle = pnlMode.add('radiobutton', undefined, L('arcDirectionCircle'));
    rbArcDirectionCircle.helpTip = L('tipDirectionCircle');
    rbArcDirectionUp.value = true;

    /* 設定パネル（ボタン類以外をまとめる）/ Settings panel (everything except the buttons) */
    var pnlMain = dlg.add('panel', undefined, L('panelTitle'));
    pnlMain.orientation = 'column';
    pnlMain.alignChildren = ['fill', 'top'];
    pnlMain.margins = [15, 20, 15, 15];

    /* まるみ / Roundness */
    var grpRoundness = pnlMain.add('group');
    grpRoundness.orientation = 'row';
    grpRoundness.alignChildren = ['left', 'center'];
    // grpRoundness.margins = [0, 5, 0, 10];

    var stArcRoundness = grpRoundness.add('statictext', undefined, L('arcRoundness'));
    setupLabelColumn(stArcRoundness);
    stArcRoundness.helpTip = L('tipRoundness');
    // Slider: 0 = flat, 50 = default arch, 100 = roundest
    var slArcRoundness = grpRoundness.add('slider', undefined, 50, 0, 100);
    slArcRoundness.preferredSize.width = 200;
    slArcRoundness.helpTip = L('tipRoundness');

    /* フィット / Fit */
    var grpFit = pnlMain.add('group');
    grpFit.orientation = 'row';
    grpFit.alignChildren = ['left', 'center'];

    // 先頭ラベル「パス幅に合わせる：」 / Leading label "Fit to path width:"
    var stFit = grpFit.add('statictext', undefined, L('fit'));
    setupLabelColumn(stFit);
    stFit.helpTip = L('tipFit');
    // フィット方法：しない（初期値）／文字サイズ＝サイズ変更／トラッキング＝サイズ維持で字間調整
    var rbFitNone = grpFit.add('radiobutton', undefined, L('fitNone'));
    rbFitNone.helpTip = L('tipFitNone');
    var rbFitFontSize = grpFit.add('radiobutton', undefined, L('fitByFontSize'));
    rbFitFontSize.helpTip = L('tipFitMethod');
    var rbFitTracking = grpFit.add('radiobutton', undefined, L('fitByTracking'));
    rbFitTracking.helpTip = L('tipFitMethod');
    // 既定は従来どおりパス幅に合わせない / Default = no fit (legacy default)
    rbFitNone.value = true;

    /* 効果 / Effect */
    var grpEffect = pnlMain.add('group');
    grpEffect.orientation = 'row';
    grpEffect.alignChildren = ['left', 'center'];
    // grpEffect.margins = [0, 0, 0, 10];

    var stEffect = grpEffect.add('statictext', undefined, L('effect'));
    setupLabelColumn(stEffect);
    stEffect.helpTip = L('tipEffect');
    var rbEffectRainbow = grpEffect.add('radiobutton', undefined, L('effectRainbow'));
    rbEffectRainbow.helpTip = L('tipEffect');
    var rbEffectDistort = grpEffect.add('radiobutton', undefined, L('effectDistort'));
    rbEffectDistort.helpTip = L('tipEffect');
    var rbEffectRibbon = grpEffect.add('radiobutton', undefined, L('effectRibbon'));
    rbEffectRibbon.helpTip = L('tipEffect');
    var rbEffectStep = grpEffect.add('radiobutton', undefined, L('effectStep'));
    rbEffectStep.helpTip = L('tipEffect');
    var rbEffectGravity = grpEffect.add('radiobutton', undefined, L('effectGravity'));
    rbEffectGravity.helpTip = L('tipEffect');
    // 既定はパス上文字の標準スタイルと同じ「虹」 / Default = Rainbow (Illustrator's own default)
    rbEffectRainbow.value = true;

    /* トラッキング / Tracking */
    var grpTracking = pnlMain.add('group');
    grpTracking.orientation = 'row';
    grpTracking.alignChildren = ['left', 'center'];
    // grpTracking.margins = [0, 0, 0, 10];

    var stTracking = grpTracking.add('statictext', undefined, L('tracking'));
    setupLabelColumn(stTracking);
    stTracking.helpTip = L('tipTracking');
    // チェックOFFでトラッキング加算を無効化（値は0に固定）/ Checkbox OFF disables tracking (forced to 0)
    var cbTracking = grpTracking.add('checkbox', undefined, '');
    cbTracking.helpTip = L('tipTrackingToggle');
    cbTracking.value = true;
    var etTracking = grpTracking.add('edittext', undefined, '0');
    etTracking.characters = 6;
    etTracking.helpTip = L('tipTracking');
    var slTracking = grpTracking.add('slider', undefined, 0, -100, 500);
    slTracking.preferredSize.width = 150;
    slTracking.helpTip = L('tipTracking');
    // 矢印キーで増減 / Arrow-key support for tracking
    changeValueByArrowKey(etTracking, true, function () { syncTrackingFromEdit(); refreshPreviewIfNeeded(); });

    /* フッター / Footer */
    var footer = dlg.add('group');
    footer.orientation = 'row';
    footer.alignChildren = ['fill', 'center'];
    footer.alignment = ['fill', 'top'];

    var leftFooter = footer.add('group');
    leftFooter.orientation = 'row';
    leftFooter.alignment = ['left', 'center'];
    var cbPreview = leftFooter.add('checkbox', undefined, L('preview'));
    cbPreview.helpTip = L('tipPreview');
    cbPreview.value = true;

    var rightFooter = footer.add('group');
    rightFooter.orientation = 'row';
    rightFooter.alignment = ['right', 'center'];
    var btnCancel = rightFooter.add('button', undefined, L('cancel'));
    var btnOk = rightFooter.add('button', undefined, L('ok'), { name: 'ok' });

    /* ===== プレビュー（Undoなし） / Preview (no undo) ===== */
    var previewTempItems = [];        // items created during preview
    var previewHiddenOriginals = [];  // originals hidden during preview

    function clearPreview() {
        // Remove temp items
        for (var i = previewTempItems.length - 1; i >= 0; i--) {
            try { previewTempItems[i].remove(); } catch (_) { }
        }
        previewTempItems = [];

        // Restore originals visibility
        for (var j = previewHiddenOriginals.length - 1; j >= 0; j--) {
            try { previewHiddenOriginals[j].hidden = false; } catch (_) { }
        }
        previewHiddenOriginals = [];
    }

    function hideOriginalForPreview(item) {
        try {
            if (!item) return;
            for (var k = 0; k < previewHiddenOriginals.length; k++) {
                if (previewHiddenOriginals[k] === item) return;
            }
            item.hidden = true;
            previewHiddenOriginals.push(item);
        } catch (_) { }
    }

    function applyPreview() {
        clearPreview();

        // Restore base selection so preview stays stable even after selection changes
        try { doc.selection = baseSelection; } catch (_) { }

        var currentSelection = [];
        try { currentSelection = doc.selection; } catch (_) { currentSelection = []; }
        if (!currentSelection || currentSelection.length === 0) {
            currentSelection = baseSelection;
        }
        targetItems = getTargetTextItems(currentSelection);
        selectedPaths = getSelectedPathItems(currentSelection);

        if (!targetItems || targetItems.length === 0) {
            cbPreview.value = false;
            return;
        }

        generateArcText(false, true);
        try { app.redraw(); } catch (_) { }
    }

    function refreshPreviewIfNeeded() {
        if (cbPreview.value) applyPreview();
    }

    /* ===== ハンドラ / Handlers ===== */
    slArcRoundness.onChange = refreshPreviewIfNeeded;

    /**
     * 円モードが選択されているか
     * @returns {boolean} 円モードなら true
     */
    function isCircleMode() {
        try { return !!rbArcDirectionCircle.value; } catch (_) { }
        return false;
    }

    /**
     * トラッキングでパス幅に合わせるモードか（円モードは対象外）
     * @returns {boolean} 有効なら true
     */
    function isFitByTrackingActive() {
        return !isCircleMode() && rbFitTracking.value;
    }

    /**
     * 文字サイズでパス幅に合わせるモードか（円モードは対象外）
     * @returns {boolean} 有効なら true
     */
    function isFitByFontSizeActive() {
        return !isCircleMode() && rbFitFontSize.value;
    }

    // パス幅フィットは開いたパスにしか効かないため、円モードではディム表示
    function updateFitEnabled() {
        var circleMode = isCircleMode();
        stFit.enabled = !circleMode;
        rbFitNone.enabled = !circleMode;
        rbFitFontSize.enabled = !circleMode;
        rbFitTracking.enabled = !circleMode;
    }
    // 「トラッキング」フィット選択中は手動トラッキング行をディム表示
    function updateTrackingEnabled() {
        var fitByTrackingActive = isFitByTrackingActive();
        stTracking.enabled = !fitByTrackingActive;
        cbTracking.enabled = !fitByTrackingActive;
        var manualTrackingActive = cbTracking.value && !fitByTrackingActive;
        etTracking.enabled = manualTrackingActive;
        slTracking.enabled = manualTrackingActive;
    }
    // 円モードに入るときのカーブ値（戻したときに復帰させる）
    var arcRoundnessBeforeCircle = null;

    // 円モードではカーブを最大（＝ほぼ一周）にし、アーチに戻したら元の値へ
    function syncRoundnessForMode() {
        if (isCircleMode()) {
            if (arcRoundnessBeforeCircle === null) arcRoundnessBeforeCircle = slArcRoundness.value;
            slArcRoundness.value = slArcRoundness.maxvalue;
        } else if (arcRoundnessBeforeCircle !== null) {
            slArcRoundness.value = arcRoundnessBeforeCircle;
            arcRoundnessBeforeCircle = null;
        }
    }
    function onModeChanged() {
        syncRoundnessForMode();
        updateFitEnabled();
        updateTrackingEnabled();
        refreshPreviewIfNeeded();
    }
    rbArcDirectionUp.onClick = onModeChanged;
    rbArcDirectionDown.onClick = onModeChanged;
    rbArcDirectionCircle.onClick = onModeChanged;

    function onFitMethodChanged() {
        updateTrackingEnabled();
        refreshPreviewIfNeeded();
    }
    rbFitNone.onClick = onFitMethodChanged;
    rbFitFontSize.onClick = onFitMethodChanged;
    rbFitTracking.onClick = onFitMethodChanged;
    rbEffectRainbow.onClick = refreshPreviewIfNeeded;
    rbEffectDistort.onClick = refreshPreviewIfNeeded;
    rbEffectRibbon.onClick = refreshPreviewIfNeeded;
    rbEffectStep.onClick = refreshPreviewIfNeeded;
    rbEffectGravity.onClick = refreshPreviewIfNeeded;

    /* トラッキング UI 同期 / Tracking UI sync (edittext <-> slider) */
    var trackingSyncLock = false;

    function syncTrackingFromEdit() {
        if (trackingSyncLock) return;
        trackingSyncLock = true;
        try {
            var v = Math.max(-100, Math.min(500, Math.round(parseNumber(etTracking.text, 0))));
            etTracking.text = String(v);
            try { slTracking.value = v; } catch (_) { }
        } catch (_) { }
        trackingSyncLock = false;
    }

    function syncTrackingFromSlider() {
        if (trackingSyncLock) return;
        trackingSyncLock = true;
        try {
            var v = Math.round(slTracking.value);
            etTracking.text = String(v);
        } catch (_) { }
        trackingSyncLock = false;
    }

    etTracking.onChanging = function () { syncTrackingFromEdit(); refreshPreviewIfNeeded(); };
    slTracking.onChanging = function () { syncTrackingFromSlider(); };
    slTracking.onChange = function () { syncTrackingFromSlider(); refreshPreviewIfNeeded(); };
    cbTracking.onClick = function () {
        // OFFにしたらトラッキングを0に戻す / Reset tracking to 0 when turned off
        if (!cbTracking.value) {
            etTracking.text = '0';
            syncTrackingFromEdit();
        }
        updateTrackingEnabled();
        refreshPreviewIfNeeded();
    };
    syncTrackingFromEdit();
    updateFitEnabled();
    updateTrackingEnabled();

    cbPreview.onClick = function () {
        if (cbPreview.value) {
            applyPreview();
        } else {
            clearPreview();
            try { app.redraw(); } catch (_) { }
        }
    };

    btnCancel.onClick = function () {
        clearPreview();
        dlg.close(0);
    };

    btnOk.onClick = function () {
        // If preview is ON, clear it first so we don't stack temporary objects.
        if (cbPreview.value) clearPreview();
        generateArcText(true, false);
        dlg.close(1);
    };

    /* 起動時に一度プレビュー / Auto-apply preview once on open */
    try {
        if (cbPreview.value) {
            applyPreview();
            try { app.redraw(); } catch (_) { }
        }
    } catch (_) { }

    var dialogResult = dlg.show();
    if (dialogResult !== 1) return;

    /* ===== テキスト・パス収集 / Collect text & paths ===== */

    // Get target text items (point text / path text), recursing into groups
    function getTargetTextItems(items) {
        var found = [];
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (item.typename === 'TextFrame') {
                try {
                    if (item.kind === TextType.POINTTEXT || item.kind === TextType.PATHTEXT) {
                        found.push(item);
                    }
                } catch (_) { }
            } else if (item.typename === 'GroupItem') {
                found = found.concat(getTargetTextItems(item.pageItems));
            }
        }
        return found;
    }

    // Get selected path items, recursing into groups
    function getSelectedPathItems(items) {
        var found = [];
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (item.typename === 'PathItem' || item.typename === 'CompoundPathItem') {
                found.push(item);
            } else if (item.typename === 'GroupItem') {
                found = found.concat(getSelectedPathItems(item.pageItems));
            }
        }
        return found;
    }

    /* ===== パススタイル / Path style ===== */

    // Run styleFn for each underlying PathItem (handles CompoundPathItem too).
    function forEachPathItem(pathItem, styleFn) {
        try {
            if (!pathItem || !styleFn) return;
            if (pathItem.typename === 'CompoundPathItem') {
                for (var cp = 0; cp < pathItem.pathItems.length; cp++) {
                    try { styleFn(pathItem.pathItems[cp]); } catch (_) { }
                }
            } else {
                try { styleFn(pathItem); } catch (_) { }
            }
        } catch (_) { }
    }

    // Per-PathItem style: invisible (no fill, stroke color/weight 0).
    function styleInvisiblePath(pi) {
        pi.filled = false;
        pi.stroked = false;
        pi.strokeWidth = 0;
    }

    // Generated arc path: stroke color/weight 0 (invisible guide).
    function applyInvisiblePathStyle(pathItem) {
        forEachPathItem(pathItem, styleInvisiblePath);
    }

    /* ===== アーチ生成 / Arc generation ===== */

    // If a path is selected together with text, it should not remain.
    // Preview: hide it (restored by clearPreview). Execute: delete it.
    function handleSelectedPaths(previewMode) {
        try {
            if (!selectedPaths || selectedPaths.length === 0) return;
            for (var i = selectedPaths.length - 1; i >= 0; i--) {
                var path = selectedPaths[i];
                if (!path) continue;
                if (previewMode) {
                    hideOriginalForPreview(path);
                } else {
                    try { path.remove(); } catch (_) { }
                }
            }
        } catch (_) { }
    }

    // Apply center justification to all paragraphs of a text frame
    function applyCenterJustification(textFrame) {
        try {
            if (!textFrame) return;
            try {
                if (textFrame.paragraphs && textFrame.paragraphs.length > 0) {
                    for (var i = 0; i < textFrame.paragraphs.length; i++) {
                        try {
                            textFrame.paragraphs[i].paragraphAttributes.justification = Justification.CENTER;
                        } catch (_) { }
                    }
                }
            } catch (_) { }
            try { textFrame.textRange.paragraphAttributes.justification = Justification.CENTER; } catch (_) { }
        } catch (_) { }
    }

    /* ===== 効果・トラッキング / Effect & tracking ===== */

    // Resolve the menu command for the selected effect (null = none)
    function getSelectedEffectCommand() {
        if (rbEffectRainbow.value) return 'Rainbow';
        if (rbEffectDistort.value) return 'Skew';
        if (rbEffectRibbon.value) return '3D ribbon';
        if (rbEffectStep.value) return 'Stair Step';
        if (rbEffectGravity.value) return 'Gravity';
        return null;
    }

    // Apply the selected path-text effect via menu command (requires selection)
    function applyPathTextEffect(textFrame) {
        var cmd = getSelectedEffectCommand();
        if (!cmd) return;

        var prevSelection = null;
        try { prevSelection = doc.selection; } catch (_) { prevSelection = null; }

        try {
            try { doc.selection = []; } catch (_) { }
            try { textFrame.selected = true; } catch (_) { }
            app.executeMenuCommand(cmd);
        } catch (_) { }

        // Always restore the previous selection (safer)
        try { doc.selection = prevSelection; } catch (_) { }
    }

    // Collect every textRange of a frame (falls back to its single textRange)
    function collectTextRanges(textFrame) {
        var ranges = [];
        try {
            if (textFrame.textRanges && textFrame.textRanges.length > 0) {
                for (var i = 0; i < textFrame.textRanges.length; i++) ranges.push(textFrame.textRanges[i]);
            }
        } catch (_) { }
        if (ranges.length === 0) {
            try { if (textFrame.textRange) ranges = [textFrame.textRange]; } catch (_) { ranges = []; }
        }
        return ranges;
    }

    // Add the tracking delta to the existing tracking of each textRange
    function applyTrackingDelta(textFrame) {
        try {
            if (!textFrame) return;
            var delta = Math.round(parseNumber(etTracking.text, 0));
            if (!delta) return; // 0 -> do nothing

            var ranges = collectTextRanges(textFrame);
            for (var r = 0; r < ranges.length; r++) {
                var textRange = ranges[r];
                try {
                    var currentTracking = textRange.characterAttributes.tracking;
                    textRange.characterAttributes.tracking = currentTracking + delta;
                } catch (_) { }
            }
        } catch (_) { }
    }

    // Measure rendered text bounds via temporary outlines: [L, T, R, B] or null
    function measureTextBounds(originalText) {
        try {
            // [0] holds only the first line (used for the baseline),
            // [1] holds everything (used for the left/top/right extents).
            var measureTexts = [originalText.duplicate(), originalText.duplicate()];
            measureTexts[0].contents = '';
            for (var i = 0; i < originalText.lines[0].length; i++) {
                originalText.textRanges[i].duplicate(measureTexts[0]);
            }

            for (var k = 0; k < measureTexts.length; k++) {
                measureTexts[k] = measureTexts[k].createOutline();
            }

            var bounds = measureTexts[1].geometricBounds; // [L, T, R, B]
            bounds[3] = measureTexts[0].geometricBounds[3]; // baseline from the first line

            for (var r = 0; r < measureTexts.length; r++) {
                try { measureTexts[r].remove(); } catch (_) { }
            }
            return bounds;
        } catch (e) {
            return null;
        }
    }

    // Roundness from the slider, clamped to 0-100 (50 = default arch)
    function readRoundnessPercent() {
        var percent = 50;
        try { percent = Number(slArcRoundness.value); } catch (_) { percent = 50; }
        if (isNaN(percent)) return 50;
        return Math.max(0, Math.min(100, percent));
    }

    // Bend a 2-point straight path into an arc by moving its Bézier handles.
    // directionSign: +1 = bulge upward, -1 = bulge downward.
    // 水平ハンドルは固定、垂直ハンドル＝カーブの深さ（50% で pathLength/4）。
    function applyArcHandles(arcPath, roundnessPercent, directionSign) {
        var pathLength = arcPath.length;
        var horizontalHandle = pathLength / 3.5;
        var verticalHandle = pathLength * (roundnessPercent / 100) * 0.5 * directionSign;

        var pathPoints = arcPath.pathPoints;
        pathPoints[0].rightDirection = [
            pathPoints[0].rightDirection[0] + horizontalHandle,
            pathPoints[0].rightDirection[1] + verticalHandle
        ];
        pathPoints[1].leftDirection = [
            pathPoints[1].leftDirection[0] - horizontalHandle,
            pathPoints[1].leftDirection[1] + verticalHandle
        ];
    }

    /**
     * 円モードで文字が占める円周の割合をカーブスライダーから求める
     * カーブ0＝大きい円（ゆるやかな弧）／カーブ100＝小さい円（ほぼ一周）
     * @returns {number} 円周に対する文字の割合（0〜1）
     */
    function readCircleOccupancy() {
        var ratio = readRoundnessPercent() / 100;
        return CIRCLE_MIN_OCCUPANCY + (CIRCLE_MAX_OCCUPANCY - CIRCLE_MIN_OCCUPANCY) * ratio;
    }

    /**
     * 4つのアンカーポイントで閉じた正円のパスを作成する
     * 始点を下（6時）に置き、下→左→上→右の時計回りにする。
     * 中央揃えの文字がパス長の中間＝円の上側中央に来るため、上向きの円弧文字になる。
     * @param {Layer} layer - パスを追加するレイヤー
     * @param {number} centerX - 円の中心X座標
     * @param {number} centerY - 円の中心Y座標
     * @param {number} radius - 円の半径
     * @returns {PathItem} 作成した閉じた円形パス
     */
    function createCirclePath(layer, centerX, centerY, radius) {
        var HANDLE_RATIO = 0.5522847498; // ベジェ4分割で正円に近似する定数
        var handle = radius * HANDLE_RATIO;

        var circlePath = layer.pathItems.add();
        circlePath.setEntirePath([
            [centerX, centerY - radius],  // 下 / bottom
            [centerX - radius, centerY],  // 左 / left
            [centerX, centerY + radius],  // 上 / top
            [centerX + radius, centerY]   // 右 / right
        ]);
        circlePath.closed = true;

        // 各アンカーの方向線を接線方向へ倒して直線を円弧にする
        var pathPoints = circlePath.pathPoints;
        pathPoints[0].leftDirection = [centerX + handle, centerY - radius];
        pathPoints[0].rightDirection = [centerX - handle, centerY - radius];
        pathPoints[1].leftDirection = [centerX - radius, centerY - handle];
        pathPoints[1].rightDirection = [centerX - radius, centerY + handle];
        pathPoints[2].leftDirection = [centerX - handle, centerY + radius];
        pathPoints[2].rightDirection = [centerX + handle, centerY + radius];
        pathPoints[3].leftDirection = [centerX + radius, centerY + handle];
        pathPoints[3].rightDirection = [centerX + radius, centerY - handle];

        try {
            circlePath.stroked = false;
            circlePath.filled = false;
        } catch (_) { }

        return circlePath;
    }

    /**
     * テキストの外接範囲から円形のパスを作成する
     * 文字幅が円周の一定割合になる半径を求め、円の頂点を元のベースライン位置に合わせる
     * @param {number[]} textBounds - 測定済みのテキスト範囲 [L, T, R, B]
     * @param {number} baselineY - 元テキストのベースラインY座標
     * @param {Layer} layer - パスを追加するレイヤー
     * @returns {PathItem|null} 作成した閉じた円形パス（幅が0のときは null）
     */
    function createCirclePathFromText(textBounds, baselineY, layer) {
        var textWidth = textBounds[2] - textBounds[0];
        if (!(textWidth > 0)) return null;

        var circumference = textWidth / readCircleOccupancy();
        var radius = circumference / (Math.PI * 2);
        var centerX = (textBounds[0] + textBounds[2]) / 2;

        return createCirclePath(layer, centerX, baselineY - radius, radius);
    }

    // Create an arc-shaped path sized to the given point text
    function createArcPathFromText(originalText, layer) {
        var baselineYMultiplier = 1.02;

        // Guard: empty / invalid text
        try {
            if (!originalText || originalText.typename !== 'TextFrame') return null;
            if (!originalText.lines || originalText.lines.length === 0) return null;
            if (!originalText.textRanges || originalText.textRanges.length === 0) return null;
        } catch (_) {
            return null;
        }

        try {
            var textBounds = measureTextBounds(originalText);
            if (!textBounds) return null;

            var baselineY = textBounds[3] * baselineYMultiplier;

            // 円モードはアーチではなく閉じた円形パスを作る
            if (isCircleMode()) {
                return createCirclePathFromText(textBounds, baselineY, layer);
            }

            // Base straight path along the baseline
            var arcPath = layer.pathItems.add();
            arcPath.setEntirePath([
                [textBounds[0], baselineY],
                [textBounds[2], baselineY]
            ]);
            try {
                arcPath.stroked = false;
                arcPath.filled = false;
            } catch (_) { }

            // Bend the straight path into an arc（上＝＋ / 下＝−）
            var directionSign = (rbArcDirectionDown && rbArcDirectionDown.value) ? -1 : 1;
            applyArcHandles(arcPath, readRoundnessPercent(), directionSign);

            return arcPath;
        } catch (e) {
            return null;
        }
    }

    // Main process: generate an arc path and place the text on it
    function generateArcText(showAlerts, previewMode) {
        if (typeof previewMode === 'undefined') previewMode = false;
        if (typeof showAlerts === 'undefined') showAlerts = true;
        var createdTexts = [];

        // A path selected together with text should not remain in arc mode
        handleSelectedPaths(previewMode);

        for (var j = 0; j < targetItems.length; j++) {
            var originalText = targetItems[j];
            var currentLayer = originalText.layer;

            // Create an arc-like path from the text bounds
            var arcPath = createArcPathFromText(originalText, currentLayer);
            if (!arcPath) {
                if (showAlerts) alert(L('alertArcFail'));
                continue;
            }

            // Generated arc path: stroke color/weight 0 (invisible guide)
            applyInvisiblePathStyle(arcPath);
            if (previewMode) previewTempItems.push(arcPath);

            // Create text on the path
            var textOnAPath = currentLayer.textFrames.pathText(arcPath);
            // Keep stacking position (avoid appearing to disappear behind other objects)
            try { textOnAPath.move(originalText, ElementPlacement.PLACEBEFORE); } catch (_) { }
            if (previewMode) previewTempItems.push(textOnAPath);

            // Keep the path used by the PathText invisible (AI may override style on conversion)
            try {
                var pathTextPath = textOnAPath.textPath;
                if (pathTextPath) applyInvisiblePathStyle(pathTextPath);
            } catch (_) { }

            // Duplicate textRanges from the original text frame
            for (var i = 0; i < originalText.textRanges.length; i++) {
                originalText.textRanges[i].duplicate(textOnAPath);
            }

            // 行揃え：常に中央 ※ duplicate 後に適用しないと上書きされる
            applyCenterJustification(textOnAPath);

            // トラッキング（既存値 + 指定値）※ フィット前に適用してオーバーセット判定へ反映
            // 「トラッキング」フィット時は手動値を加算しない（ディム表示と挙動を一致させる）
            if (!isFitByTrackingActive()) applyTrackingDelta(textOnAPath);

            // 効果（パス上文字の効果をメニューコマンドで適用）
            applyPathTextEffect(textOnAPath);

            // Collect created PathText for fit (even in preview)
            createdTexts.push(textOnAPath);

            // Remove or hide the original text frame
            if (previewMode) {
                hideOriginalForPreview(originalText);
            } else {
                originalText.remove();
            }

            // Select the created text on a path
            if (!previewMode) {
                try { textOnAPath.selected = true; } catch (_) { }
            }
        }

        // フィット：「しない」以外を選んだとき（ループ後にまとめて適用）
        if (isFitByTrackingActive()) {
            // 文字サイズを保ったまま、トラッキングでパス幅に合わせる
            try { fitTextToOpenPathByTracking(createdTexts); } catch (_) { }
        } else if (isFitByFontSizeActive()) {
            // 文字サイズを変更してパス幅に合わせる（従来）
            try { fitTextToOpenPath(createdTexts); } catch (_) { }
        }
    }

    /* ===== フィット / Fit ===== */

    // True if the path (or first sub-path of a compound path) is closed
    function isClosedPathItem(item) {
        try {
            if (!item) return false;
            if (item.typename === 'CompoundPathItem') {
                if (item.pathItems && item.pathItems.length > 0) return !!item.pathItems[0].closed;
                return false;
            }
            if (item.typename === 'PathItem') return !!item.closed;
        } catch (_) { }
        return false;
    }

    // True if the frame is editable PathText placed on an OPEN path (a fit target)
    function isTargetPathText(textFrame) {
        try {
            if (!textFrame || textFrame.typename !== 'TextFrame') return false;
            if (textFrame.kind !== TextType.PATHTEXT) return false;
            if (!textFrame.editable || textFrame.locked || textFrame.hidden) return false;
            var path = textFrame.textPath;
            if (!path) return false;
            // Only OPEN paths
            if (isClosedPathItem(path)) return false;
            return true;
        } catch (_) { }
        return false;
    }

    // Overset detection: some characters are pushed past the visible lines
    function isOverset(textFrame, lineAmount) {
        try {
            if (!textFrame) return false;

            if (textFrame.lines.length > 0) {
                var charactersOnVisibleLines = 0;

                if (typeof (lineAmount) === 'undefined' || lineAmount === null) {
                    lineAmount = 1;
                } else {
                    lineAmount = Math.floor(lineAmount);
                    if (lineAmount < 1) lineAmount = 1;
                    if (lineAmount > textFrame.lines.length) lineAmount = textFrame.lines.length;
                }

                for (var i = 0; i < lineAmount; i++) {
                    charactersOnVisibleLines += textFrame.lines[i].characters.length;
                }
                return (charactersOnVisibleLines < textFrame.characters.length);
            } else if (textFrame.characters.length > 0) {
                return true;
            }
        } catch (_) { }
        return false;
    }

    // Visible line count of a frame (always >= 1)
    function getLineAmount(textFrame) {
        try {
            if (textFrame.lines && textFrame.lines.length > 0) return textFrame.lines.length;
        } catch (_) { }
        return 1;
    }

    // Fit PathText to OPEN path endpoints by adjusting font size only.
    // - If overset: shrink by small steps until it fits.
    // - If NOT overset: grow to intentionally create overset, then shrink to fit.
    function fitTextToOpenPath(frames) {
        if (!frames || frames.length === 0) return false;

        var opt = {
            increment: 0.1,
            minFontSize: 0.1,
            maxShrinkIter: 2000,
            maxGrowIter: 10,
            maxFontSize: 2000
        };

        function shrinkFont(textFrame) {
            try {
                if (!textFrame || textFrame.characters.length <= 0) return;

                var lineAmount = getLineAmount(textFrame);

                // If it is NOT overset, grow first (doubling) until it becomes overset
                if (!isOverset(textFrame, lineAmount)) {
                    var growIter = 0;
                    while (!isOverset(textFrame, lineAmount) && growIter < opt.maxGrowIter) {
                        var growSize = textFrame.textRange.characterAttributes.size;
                        if (growSize >= opt.maxFontSize) break;
                        textFrame.textRange.characterAttributes.size = Math.min(opt.maxFontSize, growSize * 2);
                        growIter++;
                    }
                }

                // Then shrink in small steps until it fits
                var shrinkIter = 0;
                while (isOverset(textFrame, lineAmount)) {
                    var currentSize = textFrame.textRange.characterAttributes.size;
                    if (currentSize <= opt.minFontSize) break;

                    textFrame.textRange.characterAttributes.size = Math.max(opt.minFontSize, currentSize - opt.increment);

                    shrinkIter++;
                    if (shrinkIter >= opt.maxShrinkIter) break;
                }
            } catch (_) { }
        }

        for (var i = 0; i < frames.length; i++) {
            var textFrame = frames[i];
            if (!isTargetPathText(textFrame)) continue;
            shrinkFont(textFrame);
        }

        return true;
    }

    // Fit PathText to OPEN path endpoints by adjusting tracking only (font size kept).
    // - A coarse pass drives the text across the overset boundary,
    // - then a fine pass settles on the widest tracking that still fits.
    function fitTextToOpenPathByTracking(frames) {
        if (!frames || frames.length === 0) return false;

        var opt = {
            coarseStep: 50,     // tracking units per coarse step
            fineStep: 1,        // tracking units per fine step
            minTracking: -1000, // tightest allowed cumulative delta
            maxTracking: 20000, // loosest allowed cumulative delta
            maxIter: 4000
        };

        // Add a tracking delta to every textRange of the frame
        function addTrackingToFrame(textFrame, delta) {
            if (!delta) return;
            var ranges = collectTextRanges(textFrame);
            for (var r = 0; r < ranges.length; r++) {
                try {
                    var attributes = ranges[r].characterAttributes;
                    attributes.tracking = attributes.tracking + delta;
                } catch (_) { }
            }
        }

        function fitByTracking(textFrame) {
            try {
                if (!textFrame || textFrame.characters.length <= 0) return;

                var lineAmount = getLineAmount(textFrame);
                var applied = 0; // cumulative tracking delta applied so far
                var iterations;

                if (isOverset(textFrame, lineAmount)) {
                    // Too wide: tighten (coarse) until it fits
                    iterations = 0;
                    while (isOverset(textFrame, lineAmount) && iterations < opt.maxIter) {
                        if (applied - opt.coarseStep < opt.minTracking) break;
                        addTrackingToFrame(textFrame, -opt.coarseStep);
                        applied -= opt.coarseStep;
                        iterations++;
                    }
                    // Loosen back (fine) until it overflows again
                    iterations = 0;
                    while (!isOverset(textFrame, lineAmount) && iterations < opt.maxIter) {
                        if (applied + opt.fineStep > opt.maxTracking) break;
                        addTrackingToFrame(textFrame, opt.fineStep);
                        applied += opt.fineStep;
                        iterations++;
                    }
                    // Stepped one fineStep too far: pull back once so it fits
                    if (isOverset(textFrame, lineAmount)) {
                        addTrackingToFrame(textFrame, -opt.fineStep);
                        applied -= opt.fineStep;
                    }
                } else {
                    // Fits with room: loosen (coarse) until it overflows
                    iterations = 0;
                    while (!isOverset(textFrame, lineAmount) && iterations < opt.maxIter) {
                        if (applied + opt.coarseStep > opt.maxTracking) break;
                        addTrackingToFrame(textFrame, opt.coarseStep);
                        applied += opt.coarseStep;
                        iterations++;
                    }
                    // Tighten back (fine) until it fits
                    iterations = 0;
                    while (isOverset(textFrame, lineAmount) && iterations < opt.maxIter) {
                        if (applied - opt.fineStep < opt.minTracking) break;
                        addTrackingToFrame(textFrame, -opt.fineStep);
                        applied -= opt.fineStep;
                        iterations++;
                    }
                }
            } catch (_) { }
        }

        for (var i = 0; i < frames.length; i++) {
            var textFrame = frames[i];
            if (!isTargetPathText(textFrame)) continue;
            fitByTracking(textFrame);
        }

        return true;
    }
}());
