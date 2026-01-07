#target illustrator

var SCRIPT_VERSION = '1.2';
var SCRIPT_TITLE = {
    ja: 'ブレンドSpecial',
    en: 'Blend Special'
};
var SCRIPT_UPDATED = '2026-01-01';

/*
 * 概要
 * - 選択にブレンド（BlendItem/PluginItem）が含まれる場合：ブレンドは作成せず、設定／調整のみを行います（ブレンド＋パス等の混在選択も含む）。
 * - 選択がブレンド内のパス等（PathItem など）の場合でも、可能な限りブレンド本体を見つけて同様に扱います（※環境差あり）。
 * - 選択にブレンドが含まれない場合：ブレンドを作成してから、設定／調整を行います。
 * - ステップ数／方向はダイアログ中にライブプレビュー（Undo で戻せる方式）します。
 * - ステップ数は整数のみ（小数不可）、0〜1000 の範囲で扱います（0 可）。
 * - ダイアログ表示時、ブレンドを選択している場合はそのステップ数を取得してデフォルト表示します（取得できない場合は 8）。
 * - ブレンドを選択していない場合、［ステップ数］のデフォルトは 8 です。
 * - ステップ数の下にスライダーを表示：
 *    - 通常：0〜32
 *    - option：0〜128
 *    - shift：0〜1000（shift > option の優先）
 * - 「解除／拡張／ブレンド軸を置き換え」は安全のためプレビューでは実行せず、OK 押下時にのみ実行します。
 * - 「反転」はチェック操作で即時プレビューします（トグル二重実行で元に戻らないよう制御）。
 * - ローカライズ文言は 1 箇所（LABELS）にまとめています（日本語／英語）。
 * - ダイアログは 2 カラム構成：左（方向／反転）、右（その他）。ステップ数は上部に全幅で表示します。
 * - 「その他」が「なし」以外のとき、方向／反転パネルを無効化します（状態は保持）。
 * - OK／キャンセルボタンはダイアログ中央に配置します。
 */

$.localize = true;

// --- UI Localizations (one place) ---
var LABELS = {
    title: {
        ja: SCRIPT_TITLE.ja + ' v' + SCRIPT_VERSION,
        en: SCRIPT_TITLE.en + ' v' + SCRIPT_VERSION
    },
    stepLabel: {
        ja: 'ステップ数',
        en: 'Steps'
    },
    alignToPage: {
        ja: '垂直方向',
        en: 'Align to Page'
    },
    alignToPath: {
        ja: 'パスに沿う',
        en: 'Align to Path'
    },
    ok: {
        ja: 'OK',
        en: 'OK'
    },
    cancel: {
        ja: 'キャンセル',
        en: 'Cancel'
    },
    invalidStep: {
        ja: 'ステップ数は 0〜1000 の整数で入力してください。',
        en: 'Please enter an integer from 0 to 1000.'
    },
    adjustPanel: {
        ja: 'その他',
        en: 'Misc'
    },
    reversePanel: {
        ja: '反転',
        en: 'Reverse'
    },
    orientationPanel: {
        ja: '方向',
        en: 'Orientation'
    },
    adjustNone: {
        ja: 'なし',
        en: 'None'
    },
    adjustRelease: {
        ja: '解除',
        en: 'Release'
    },
    adjustExpand: {
        ja: '拡張',
        en: 'Expand'
    },
    adjustReplace: {
        ja: 'ブレンド軸を置き換え',
        en: 'Replace Spine'
    },
    reverseSpine: {
        ja: 'ブレンド軸を反転',
        en: 'Reverse Spine'
    },
    reverseStack: {
        ja: '前後を反転',
        en: 'Reverse Front to Back'
    }
};

// ブレンドオプション > 方向
// 0: Align to Page（ページに揃える） / 1: Align to Path（パスに沿う）
var BlendOrientation = {
    'Align to Page': 0,
    'Align to Path': 1
};

// エントリーポイント（UI専用スクリプト）
// 選択内容を元にダイアログを表示し、ブレンドの設定／調整のみを行う
main();

function main() {
    if (app.documents.length <= 0) {
        return;
    }
    var doc = app.activeDocument;
    var selection = doc.selection;

    // ダイアログを開いた時点で「ブレンドオブジェクトを選択していたか」を記録
    // （この後 Path Blend Make を実行して選択がブレンドに変わることがあるため）
    var wasBlendSelectedAtOpen = false;
    var blendInSelectionAtOpen = null;
    try {
        blendInSelectionAtOpen = findFirstPluginItem(selection);
        wasBlendSelectedAtOpen = !!blendInSelectionAtOpen;
    } catch (e) {
        blendInSelectionAtOpen = null;
        wasBlendSelectedAtOpen = false;
    }

    // 選択にブレンド（PluginItem/BlendItem）が含まれない場合のみ、ブレンドを作成
    // （既に PluginItem/BlendItem が含まれる場合は、作成せずオプション/調整のみ行う）
    if (!blendInSelectionAtOpen) {
        try {
            app.executeMenuCommand('Path Blend Make');
        } catch (e) {}

        // コマンド実行後に選択が変わる可能性があるため、再取得
        selection = doc.selection;
        if (!selection || selection.length <= 0) {
            return;
        }

        // 生成された PluginItem（ブレンド）だけを選択状態にする
        var madeBlend = findFirstPluginItem(selection);
        if (madeBlend) {
            selectOnlyItem(doc, madeBlend);
            selection = doc.selection;
        }
    } else {
        // Selection already contains a blend object; keep current selection (even if paths are also selected)
        selection = doc.selection;
        if (!selection || selection.length <= 0) {
            return;
        }
    }

    // ダイアログを表示してユーザー入力を取得
    // ※ 引数実行は行わない（UI操作専用）
    var input = getInput(selection, wasBlendSelectedAtOpen);
    if (!input) {
        return;
    } // cancel

    var blendStep = input.step;
    var orientationValue = input.orientation;
    var adjustMode = input.adjustMode || 'none';
    var reverseSpine = !!input.reverseSpine;
    var reverseStack = !!input.reverseStack;

    if (adjustMode === 'release') {
        app.executeMenuCommand('Path Blend Release');
        // ※ コマンド実行後に選択が変わる可能性があるため、doc.selection を参照する
        try {
            deleteFrontmostPath(doc.selection);
        } catch (e) {}
        return;
    }
    if (adjustMode === 'expand') {
        app.executeMenuCommand('Path Blend Expand');
        return;
    }
    if (adjustMode === 'replaceSpine') {
        app.executeMenuCommand('Path Blend Replace Spine');
        // ※ コマンド実行後に選択が変わる可能性があるため、doc.selection を参照する
        try {
            deleteFrontmostPath(doc.selection);
        } catch (e) {}
        return;
    }

    if (reverseSpine) {
        app.executeMenuCommand('Path Blend Reverse Spine');
        // ※ コマンド実行後に選択が変わる可能性があるため、doc.selection を参照する
        try {
            deleteFrontmostPath(doc.selection);
        } catch (e) {}
    }
    if (reverseStack) {
        app.executeMenuCommand('Path Blend Reverse Stack');
        // ※ コマンド実行後に選択が変わる可能性があるため、doc.selection を参照する
        try {
            deleteFrontmostPath(doc.selection);
        } catch (e) {}
    }
    if (reverseSpine || reverseStack) {
        return;
    }

    // Apply option only (no new blend creation)
    setBlendOption(blendStep, orientationValue);
}

// 選択オブジェクトがすべて PluginItem または BlendItem かどうかを判定する
// すべて PluginItem/BlendItem なら「既にブレンド（などのプラグインオブジェクト）」扱い
function isAllPluginItems(selection) {
    // Backward-compatible name: treat PluginItem and BlendItem as "already blend-like"
    try {
        if (!selection || selection.length <= 0) {
            return false;
        }
        for (var i = 0; i < selection.length; i++) {
            var it = selection[i];
            if (!it) {
                return false;
            }
            var t = it.typename;
            if (t !== 'PluginItem' && t !== 'BlendItem') {
                return false;
            }
        }
        return true;
    } catch (e) {
        return false;
    }
}

// selection から最初の BlendItem または PluginItem を返す（BlendItem優先、なければPluginItem、なければ null）
function findFirstPluginItem(selection) {
    // Backward-compatible name: return first BlendItem/PluginItem (BlendItem preferred)
    try {
        if (!selection || selection.length <= 0) {
            return null;
        }
        // Prefer BlendItem if present (it usually exposes blendOptions.steps)
        for (var i = 0; i < selection.length; i++) {
            if (selection[i] && selection[i].typename === 'BlendItem') {
                return selection[i];
            }
        }
        for (var j = 0; j < selection.length; j++) {
            if (selection[j] && selection[j].typename === 'PluginItem') {
                return selection[j];
            }
        }
    } catch (e) {}
    return null;
}

// 指定アイテムのみを選択状態にする（失敗しても例外は投げない）
function selectOnlyItem(doc, item) {
    try {
        if (!doc || !item) {
            return;
        }
        doc.selection = null;
        item.selected = true;
    } catch (e) {}
}

// 選択オブジェクトが「ブレンドオブジェクトのみ」かどうかを判定する
// （ブレンド＋パス 等の混在選択なら false）
// 「ブレンド軸を置き換え」オプションの有効／無効制御に使用
function isOnlyBlendObjects(selection) {
    try {
        if (!selection || selection.length <= 0) {
            return false;
        }
        // 1つ以上あり、全てが BlendItem / PluginItem のとき「ブレンドのみ」
        for (var i = 0; i < selection.length; i++) {
            var it = selection[i];
            if (!it) {
                return false;
            }
            var t = it.typename;
            if (t !== 'PluginItem' && t !== 'BlendItem') {
                return false;
            }
        }
        return true;
    } catch (e) {
        return false;
    }
}

// PluginItem のブレンドからステップ数（中間ステップ数）を取得（取得失敗時は -1）
// ※元データを壊さないため、複製→Expand→カウント→後片付け で算出する
function getBlendStepsFromPluginItem(blendItem) {
    // Guard: PluginItem 以外は対象外
    if (!blendItem || blendItem.typename !== 'PluginItem') {
        return -1;
    }

    var doc = app.activeDocument;
    var stepCount = -1;
    var tempItem = null;

    // 選択状態を退避（この関数内で選択を変更するため）
    var originalSelection = [];
    try {
        var sel = doc.selection;
        if (sel && sel.length) {
            for (var i = 0; i < sel.length; i++) {
                originalSelection.push(sel[i]);
            }
        }
    } catch (e0) {}

    try {
        // 1) 複製（元データ保護）
        tempItem = blendItem.duplicate();

        // 2) 複製したものだけを選択
        doc.selection = null;
        tempItem.selected = true;

        // 3) Expand
        app.executeMenuCommand('Path Blend Expand');

        // 4) Expand 後の選択（展開結果）を確認
        if (doc.selection && doc.selection.length > 0) {
            // 5) トップレベルのグループを解除（ネストまでは無理に追わない）
            try { app.executeMenuCommand('ungroup'); } catch (e1) {}

            // 6) 数をカウント
            var totalItems = doc.selection.length;

            // 7) 始点・終点を除外（全要素数 - 2）
            stepCount = (totalItems >= 2) ? (totalItems - 2) : 0;
        }
    } catch (e2) {
        stepCount = -1;
    } finally {
        // 展開残骸を削除
        try {
            var junk = doc.selection;
            if (junk && junk.length) {
                for (var x = 0; x < junk.length; x++) {
                    try { junk[x].remove(); } catch (e3) {}
                }
            }
        } catch (e4) {}

        // 複製が残っていた場合の保険
        try {
            if (tempItem) { tempItem.remove(); }
        } catch (e5) {}

        // 選択状態を復元
        try {
            doc.selection = null;
            for (var j = 0; j < originalSelection.length; j++) {
                try { originalSelection[j].selected = true; } catch (e6) {}
            }
        } catch (e7) {}
    }

    return stepCount;
}

function getInput(selection, wasBlendSelectedAtOpen) {

    var doc = app.activeDocument;

    // ダイアログ開始時の選択を保存（Replace Spine などで「ブレンド＋パス」の混在選択を維持するため）
    var originalSelectionForDialog = [];
    try {
        var sel0 = doc.selection;
        if (sel0 && sel0.length) {
            for (var si = 0; si < sel0.length; si++) {
                originalSelectionForDialog.push(sel0[si]);
            }
        }
    } catch (eSel0) {}

    function restoreOriginalSelectionForDialog() {
        try {
            doc.selection = null;
            for (var k = 0; k < originalSelectionForDialog.length; k++) {
                try { originalSelectionForDialog[k].selected = true; } catch (eSel1) {}
            }
        } catch (eSel2) {}
    }

    // 選択内容を元にダイアログを構築・表示する
    var dlg = new Window('dialog', localize(LABELS.title));
    dlg.orientation = 'column';
    dlg.alignChildren = ['fill', 'top'];
    dlg.spacing = 10;
    dlg.margins = 16;

    // --- Step (one-column, full width) ---
    var stepArea = dlg.add('group');
    stepArea.orientation = 'column';
    stepArea.alignChildren = ['fill', 'top'];
    stepArea.spacing = 6;

    var stepRow = stepArea.add('group');
    stepRow.orientation = 'row';
    stepRow.alignChildren = ['left', 'center'];
    stepRow.add('statictext', undefined, localize(LABELS.stepLabel));

    // Default step:
    // - If a blend object was selected when opening the dialog: get its steps when possible (fallback 8)
    // - If not: default is 8
    var _defaultStep = 8;
    if (wasBlendSelectedAtOpen) {
        try {
            var _bi = findFirstPluginItem(selection);
            // BlendItem typically exposes blendOptions.steps; PluginItem may not
            if (_bi && _bi.blendOptions && typeof _bi.blendOptions.steps === 'number') {
                _defaultStep = _bi.blendOptions.steps;
            } else if (_bi && _bi.typename === 'PluginItem') {
                var _calc = getBlendStepsFromPluginItem(_bi);
                if (typeof _calc === 'number' && _calc >= 0) {
                    _defaultStep = _calc;
                }
            }
        } catch (e) {}
    }

    var edtStep = stepRow.add('edittext', undefined, String(_defaultStep));
    edtStep.characters = 4;

    // --- Slider (under Steps) ---
    var sldStep = stepArea.add('slider', undefined, _defaultStep, 0, 32);
    sldStep.alignment = ['fill', 'center'];

    // Keep the slider thumb position stable when switching range (32 <-> 128 <-> 1000)
    var _stepSliderLastMax = 32;
    function updateStepSliderRangeFromKeyboard() {
        try {
            var kb = ScriptUI.environment.keyboardState;

            // Priority: Shift (0-1000) > Option (0-128) > default (0-32)
            var nextMax = 32;
            if (kb && kb.shiftKey) {
                nextMax = 1000;
            } else if (kb && kb.altKey) {
                nextMax = 128;
            }

            var prevMax = _stepSliderLastMax;
            if (prevMax !== nextMax) {
                // Preserve relative position of the thumb (ratio) when changing range
                var ratio = 0;
                try {
                    ratio = (prevMax > 0) ? (sldStep.value / prevMax) : 0;
                    if (!isFinite(ratio) || isNaN(ratio)) ratio = 0;
                } catch (e0) {
                    ratio = 0;
                }

                sldStep.maxvalue = nextMax;

                var nextVal = Math.round(ratio * nextMax);
                if (nextVal < 0) nextVal = 0;
                if (nextVal > nextMax) nextVal = nextMax;

                try { sldStep.value = nextVal; } catch (e1) {}
                // Keep edit text consistent with the new value
                edtStep.text = String(nextVal);

                _stepSliderLastMax = nextMax;
            } else {
                // Ensure range is correct even if maxvalue was modified elsewhere
                if (sldStep.maxvalue !== nextMax) {
                    sldStep.maxvalue = nextMax;
                }
            }
        } catch (e) {}
    }

    function clampStepValue(v) {
        v = Math.round(v);
        if (v < 0) v = 0;
        var _max = 32;
        try { _max = sldStep.maxvalue; } catch (e) { _max = 32; }
        if (v > _max) v = _max;
        return v;
    }

    function syncSliderFromEditText() {
        var v = parseInt(edtStep.text, 10);
        if (isNaN(v)) return;
        v = clampStepValue(v);
        try { sldStep.value = v; } catch (e) {}
    }

    function syncEditTextFromSlider() {
        var v = clampStepValue(sldStep.value);
        edtStep.text = String(v);
    }

    // Arrow keys: update slider + preview
    changeValueByArrowKey(edtStep, function() {
        syncSliderFromEditText();
        applyPreviewFromUI();
    });

    // Slider: update edit text + preview
    sldStep.onChanging = function() {
        updateStepSliderRangeFromKeyboard();
        syncEditTextFromSlider();
        applyPreviewFromUI();
    };
    sldStep.onChange = function() {
        updateStepSliderRangeFromKeyboard();
        syncEditTextFromSlider();
        applyPreviewFromUI();
    };

    // --- 2 columns ---
    var cols = dlg.add('group');
    cols.orientation = 'row';
    cols.alignChildren = ['fill', 'top'];
    cols.spacing = 12;

    var colLeft = cols.add('group');
    colLeft.orientation = 'column';
    colLeft.alignChildren = ['fill', 'top'];
    colLeft.spacing = 10;

    var colRight = cols.add('group');
    colRight.orientation = 'column';
    colRight.alignChildren = ['fill', 'top'];
    colRight.spacing = 10;

    var panelBlend = colLeft.add('panel', undefined, localize(LABELS.orientationPanel));
    panelBlend.orientation = 'column';
    panelBlend.alignChildren = ['fill', 'top'];
    panelBlend.spacing = 8;
    panelBlend.margins = [12, 18, 12, 12];

    var row2 = panelBlend.add('group');
    row2.orientation = 'row';
    row2.alignChildren = ['left', 'top'];

    var orientGroup = row2.add('group');
    orientGroup.orientation = 'column';
    orientGroup.alignChildren = ['left', 'top'];
    orientGroup.spacing = 4;

    var rbVertical = orientGroup.add('radiobutton', undefined, localize(LABELS.alignToPage));
    var rbAlign = orientGroup.add('radiobutton', undefined, localize(LABELS.alignToPath));
    rbVertical.value = true;

    // --- Live Preview (best-effort) ---
    // Count how many preview operations we applied so Cancel can rollback.
    var previewUndoDepth = 0;
    var lastBlendPreviewApplied = false;

    // Target blend item for preview operations (ensure it's selected when running commands)
    var targetBlendItem = null;
    try {
        targetBlendItem = findFirstPluginItem(app.activeDocument.selection);
    } catch (e) {}

    function ensureTargetBlendSelected() {
        try {
            if (!targetBlendItem) {
                targetBlendItem = findFirstPluginItem(app.activeDocument.selection);
            }
            if (targetBlendItem) {
                selectOnlyItem(app.activeDocument, targetBlendItem);
            }
        } catch (e) {}
    }

    function getUIValuesForPreview() {
        var step = parseInt(edtStep.text, 10);
        if (isNaN(step)) {
            return null;
        }
        if (step < 0) {
            step = 0;
        }
        if (step > 1000) {
            step = 1000;
        }

        var orientation = (rbAlign.value) ? BlendOrientation['Align to Path'] : BlendOrientation['Align to Page'];
        return {
            step: step,
            orientation: orientation
        };
    }

    function applyPreview(values) {
        if (!values) {
            return;
        }
        try {
            ensureTargetBlendSelected();
            setBlendOption(values.step, values.orientation);
            previewUndoDepth++;
            lastBlendPreviewApplied = true; // informational only
            try {
                app.redraw();
            } catch (e2) {}
        } catch (e3) {}
    }

    function applyPreviewFromUI() {
        var v = getUIValuesForPreview();
        applyPreview(v);
    }

    function applyAdjustPreviewFromUI() {
        // NOTE: 安全重視のため、解除/拡張/ブレンド軸置き換えはプレビューでは実行しない。
        // 実行は OK 押下時（main側）で初めて行う。
        return;
    }

    // Step: preview on manual edits as well
    edtStep.onChanging = function() {
        // Avoid spamming undo/apply too aggressively by only previewing when the value parses
        // (Arrow keys already preview via changeValueByArrowKey callback)
        var v = parseInt(edtStep.text, 10);
        if (isNaN(v)) {
            return;
        }
        // integer only + clamp (do not force-write on every keystroke)
        if (v < 0) v = 0;
        if (v > 1000) v = 1000;
        syncSliderFromEditText();
        applyPreviewFromUI();
    };

    // Sanitize on commit (enter / focus-out)
    edtStep.onChange = function() {
        var v = parseInt(edtStep.text, 10);
        if (isNaN(v)) {
            v = 0;
        }
        if (v < 0) v = 0;
        if (v > 1000) v = 1000;
        edtStep.text = String(Math.round(v));
        syncSliderFromEditText();
        try { applyPreviewFromUI(); } catch (e) {}
    };

    // Orientation: preview on change
    rbVertical.onClick = applyPreviewFromUI;
    rbAlign.onClick = applyPreviewFromUI;

    var panelAdjust = colRight.add('panel', undefined, localize(LABELS.adjustPanel));
    panelAdjust.orientation = 'column';
    panelAdjust.alignChildren = ['left', 'top'];
    panelAdjust.spacing = 6;
    panelAdjust.margins = [12, 18, 12, 12];

    var adjustGroup = panelAdjust.add('group');
    adjustGroup.orientation = 'column';
    adjustGroup.alignChildren = ['left', 'top'];
    adjustGroup.spacing = 4;

    var rbAdjustNone = adjustGroup.add('radiobutton', undefined, localize(LABELS.adjustNone));
    var rbAdjustRelease = adjustGroup.add('radiobutton', undefined, localize(LABELS.adjustRelease));
    var rbAdjustExpand = adjustGroup.add('radiobutton', undefined, localize(LABELS.adjustExpand));
    var rbAdjustReplace = adjustGroup.add('radiobutton', undefined, localize(LABELS.adjustReplace));
    rbAdjustNone.value = true;

    // --- Reverse panel ---
    var panelReverse = colLeft.add('panel', undefined, localize(LABELS.reversePanel));
    panelReverse.orientation = 'column';
    panelReverse.alignChildren = ['left', 'top'];
    panelReverse.spacing = 6;
    panelReverse.margins = [12, 18, 12, 12];

    var reverseGroup = panelReverse.add('group');
    reverseGroup.orientation = 'column';
    reverseGroup.alignChildren = ['left', 'top'];
    reverseGroup.spacing = 4;

    var cbReverseSpine = reverseGroup.add('checkbox', undefined, localize(LABELS.reverseSpine));
    var cbReverseStack = reverseGroup.add('checkbox', undefined, localize(LABELS.reverseStack));
    cbReverseSpine.value = false;
    cbReverseStack.value = false;


    // Cross-panel radio exclusivity (ScriptUI only auto-excludes within same parent)
    function selectAdjustMode(which) {
        rbAdjustNone.value = (which === 'none');
        rbAdjustRelease.value = (which === 'release');
        rbAdjustExpand.value = (which === 'expand');
        rbAdjustReplace.value = (which === 'replaceSpine');
    }

    // 「その他」が「なし」のとき、他の選択肢を視覚的にディム表示（ただし選択は可能）
    // ※ enabled=false にすると選択できなくなるため、文字色のみ変更する。
    var _dimPen = null;
    var _normalPen = null;
    function syncAdjustOptionDimming() {
        try {
            if (!_dimPen) {
                _dimPen = dlg.graphics.newPen(PenType.SOLID_COLOR, [0.5, 0.5, 0.5, 1], 1);
            }
            if (!_normalPen) {
                _normalPen = dlg.graphics.newPen(PenType.SOLID_COLOR, [0, 0, 0, 1], 1);
            }

            var useDim = !!rbAdjustNone.value;

            // Keep existing enable/disable behavior (e.g., GroupItem-only restriction) intact.
            // Only change the label color for the radios that are currently enabled.
            var targets = [rbAdjustRelease, rbAdjustExpand, rbAdjustReplace];
            for (var i = 0; i < targets.length; i++) {
                var rb = targets[i];
                if (!rb) continue;

                // If disabled (e.g. replace on GroupItem-only), keep it dim.
                if (rb.enabled === false) {
                    rb.graphics.foregroundColor = _dimPen;
                    continue;
                }

                rb.graphics.foregroundColor = useDim ? _dimPen : _normalPen;
            }
        } catch (e) {}
    }

    // 選択が「ブレンドオブジェクトのみ」の場合は
    // 「ブレンド軸を置き換え」をディム表示にする
    // （ブレンド＋パス等の混在選択なら有効）
    function updateAdjustAvailability() {
        try {
            var onlyBlend = isOnlyBlendObjects(selection);
            rbAdjustReplace.enabled = !onlyBlend;
            if (!rbAdjustReplace.enabled) {
                // If it was selected somehow, fall back to None
                if (rbAdjustReplace.value) {
                    selectAdjustMode('none');
                }
            }
            syncAdjustOptionDimming();
        } catch (e) {}
    }

    // 「調整」が「なし」以外のときのみ
    // ブレンド設定パネルを無効化する（反転操作ではディムにしない）
    function syncBlendPanelEnabled() {
        var enable = !!rbAdjustNone.value;
        panelBlend.enabled = enable;

        // Keep focus sensible
        try {
            if (enable) {
                edtStep.active = true;
                edtStep.setSelection(0, edtStep.text.length);
            } else {
                // Prefer focusing the currently-selected option if possible
                if (rbAdjustRelease.value) rbAdjustRelease.active = true;
                else if (rbAdjustExpand.value) rbAdjustExpand.active = true;
                else if (rbAdjustReplace.value) rbAdjustReplace.active = true;
                else rbAdjustNone.active = true;
            }
        } catch (e) {}
    }

    // 「調整」が「なし」以外のときは
    // 「反転」パネルを無効化する
    function syncReversePanelEnabled() {
        var enable = !!rbAdjustNone.value;
        panelReverse.enabled = enable;

        // Disable without mutating checkbox state (do not clear values)
        // Preview changes are not applied here to avoid losing the user's intent on OK.
        // Any preview already applied remains as-is until Cancel/OK handling.
        if (!enable) {
            // no-op
        }
    }

    // When choosing any adjust option (other than None), dim blend options
    rbAdjustNone.onClick = function() {
        selectAdjustMode('none');
        updateAdjustAvailability();
        syncBlendPanelEnabled();
        syncReversePanelEnabled();
        syncAdjustOptionDimming();
    };
    rbAdjustRelease.onClick = function() {
        selectAdjustMode('release');
        updateAdjustAvailability();
        syncBlendPanelEnabled();
        syncReversePanelEnabled();
        syncAdjustOptionDimming();
    };
    rbAdjustExpand.onClick = function x() {
        selectAdjustMode('expand');
        updateAdjustAvailability();
        syncBlendPanelEnabled();
        syncReversePanelEnabled();
        syncAdjustOptionDimming();
    };
    rbAdjustReplace.onClick = function() {
        selectAdjustMode('replaceSpine');
        updateAdjustAvailability();
        syncBlendPanelEnabled();
        syncReversePanelEnabled();
        syncAdjustOptionDimming();
    };
    cbReverseSpine.onClick = function() {
        // Reverse actions are independent; keep Adjust in "None" to avoid mixed-mode confusion
        selectAdjustMode('none');
        updateAdjustAvailability();
        syncBlendPanelEnabled();
        syncReversePanelEnabled();
        applyReversePreviewFromUI();
    };
    cbReverseStack.onClick = function() {
        selectAdjustMode('none');
        updateAdjustAvailability();
        syncBlendPanelEnabled();
        syncReversePanelEnabled();
        applyReversePreviewFromUI();
    };
    // --- Immediate Reverse Preview ---
    var lastReverseSpine = cbReverseSpine.value;
    var lastReverseStack = cbReverseStack.value;
    // Reverse commands are toggles; if we executed them during preview, do NOT execute again on OK.
    var reversePreviewTouched = false;

    function applyReversePreviewFromUI() {
        // Immediate reverse preview (no delay)
        // Reverse commands are toggles; apply only when state changed.
        var wantSpine = !!cbReverseSpine.value;
        var wantStack = !!cbReverseStack.value;

        ensureTargetBlendSelected();

        try {
            if (wantSpine !== lastReverseSpine) {
                app.executeMenuCommand('Path Blend Reverse Spine');
                previewUndoDepth++;
                lastReverseSpine = wantSpine;
                // lastBlendPreviewApplied = false; // removed, no longer needed
                reversePreviewTouched = true;
            }
        } catch (e1) {}

        try {
            if (wantStack !== lastReverseStack) {
                app.executeMenuCommand('Path Blend Reverse Stack');
                previewUndoDepth++;
                lastReverseStack = wantStack;
                // lastBlendPreviewApplied = false; // removed, no longer needed
                reversePreviewTouched = true;
            }
        } catch (e2) {}

        try {
            app.redraw();
        } catch (e3) {}
    }

    // Initial state
    updateAdjustAvailability();
    syncBlendPanelEnabled();
    syncReversePanelEnabled();
    syncAdjustOptionDimming();

    var btns = dlg.add('group');
    btns.orientation = 'row';
    btns.alignment = 'center';
    var btnCancel = btns.add('button', undefined, localize(LABELS.cancel), {
        name: 'cancel'
    });
    var btnOk = btns.add('button', undefined, localize(LABELS.ok), {
        name: 'ok'
    });

    dlg.onShow = function() {
        try {
            applyPreviewFromUI();
            updateAdjustAvailability();
            syncBlendPanelEnabled();
            syncReversePanelEnabled();
            syncAdjustOptionDimming();
        } catch (e) {}
    };

    btnOk.onClick = function() {
        // Snapshot reverse states to avoid any UI sync side-effects
        var _keepReverseSpine = !!cbReverseSpine.value;
        var _keepReverseStack = !!cbReverseStack.value;

        var step = parseInt(edtStep.text, 10);
        if (isNaN(step) || step < 0 || step > 1000) {
            alert(localize(LABELS.invalidStep));
            return;
        }
        var adjustMode = 'none';
        if (rbAdjustRelease.value) {
            adjustMode = 'release';
        } else if (rbAdjustExpand.value) {
            adjustMode = 'expand';
        } else if (rbAdjustReplace.value) {
            adjustMode = 'replaceSpine';
        }

        // Replace Spine は「ブレンド＋置き換え用パス」の同時選択が必要。
        // プレビュー中にブレンド単体選択へ切り替わっていることがあるため、ここで元の選択を復元する。
        if (adjustMode === 'replaceSpine') {
            restoreOriginalSelectionForDialog();
        }

        // Restore (defensive) before closing
        cbReverseSpine.value = _keepReverseSpine;
        cbReverseStack.value = _keepReverseStack;

        dlg._result = {
            step: step,
            orientation: (rbAlign.value) ? BlendOrientation['Align to Path'] : BlendOrientation['Align to Page'],
            adjustMode: adjustMode,
            // If reverse preview ran, the document is already in the desired state.
            // Passing true here would toggle again in main() and cancel the effect.
            reverseSpine: reversePreviewTouched ? false : _keepReverseSpine,
            reverseStack: reversePreviewTouched ? false : _keepReverseStack
        };
        // Keep preview result; prevent cancel rollback
        previewUndoDepth = 0;
        dlg.close(1);
    };

    btnCancel.onClick = function() {
        // Revert all preview changes applied during this dialog session
        try {
            while (previewUndoDepth > 0) {
                app.executeMenuCommand('undo');
                previewUndoDepth--;
            }
            try {
                app.redraw();
            } catch (e2) {}
        } catch (e) {}
        // 取り消し時も、ダイアログ開始時の選択に戻す
        restoreOriginalSelectionForDialog();
        dlg.close(0);
    };

    var r = dlg.show();
    if (r !== 1) {
        return null;
    }
    return dlg._result || null;
}

function changeValueByArrowKey(editText, onValueChanged) {
    editText.addEventListener("keydown", function(event) {
        // Only handle up/down keys
        if (event.keyName !== "Up" && event.keyName !== "Down") return;

        var value = parseInt(editText.text, 10);
        if (isNaN(value)) value = 0;

        var keyboard = ScriptUI.environment.keyboardState;
        var delta = 1;

        // Shiftキー押下時は10刻み（整数のみ）
        if (keyboard.shiftKey) {
            delta = 10;
        }

        if (event.keyName === "Up") {
            value += delta;
            event.preventDefault();
        } else if (event.keyName === "Down") {
            value -= delta;
            event.preventDefault();
        }

        // Clamp: integer only, min 0, max 1000
        value = Math.round(value);
        if (value < 0) value = 0;
        if (value > 1000) value = 1000;

        editText.text = String(value);

        try {
            if (typeof onValueChanged === 'function') {
                onValueChanged();
            }
        } catch (e) {}
    });
}

function deleteFrontmostPath(selection) {
    try {
        if (!selection || selection.length <= 0) {
            return false;
        }

        var target = null;
        var bestZ = null;

        for (var i = 0; i < selection.length; i++) {
            var it = selection[i];
            if (!it) {
                continue;
            }

            // PathItem only
            if (it.typename === 'PathItem') {
                var z = null;
                try {
                    z = it.zOrderPosition;
                } catch (e1) {
                    z = null;
                }

                if (bestZ === null || bestZ === undefined) {
                    target = it;
                    bestZ = z;
                } else {
                    // Prefer higher zOrderPosition when available; otherwise keep the first one.
                    if (z !== null && z !== undefined && (bestZ === null || bestZ === undefined || z > bestZ)) {
                        target = it;
                        bestZ = z;
                    }
                }
            }
        }

        if (target) {
            target.remove();
            return true;
        }
    } catch (e2) {}

    return false;
}

/**
 * アクションを文字列から生成し実行するブロック構文。終了時・エラー発生時の後片付けは自動
 * @param {String} actionCode アクションのソースコード
 * @param {Function} func ブロック内処理をここに記述する
 * @return なし
 */
function tempAction(actionCode, func) {
    // utf8の16進数文字コードをJavaScript内部で扱える文字列に変換する
    var hexToString = function(hex) {
        var res = decodeURIComponent(hex.replace(/(.{2})/g, '%$1'));
        return res;
    };

    // ActionItemのconstructor。ActionItem.exec()を使えばわざわざ名前を直接指定しなくても実行できる
    var ActionItem = function ActionItem(index, name, parent) {
        this.index = index;
        this.name = name; // actionName
        this.parent = parent; // setName
    };
    ActionItem.prototype.exec = function(showDialog) {
        doScript(this.name, this.parent, showDialog);
    };

    // ActionItemsのconstructor。
    // ActionItems['actionName'],  ActionItems.getByName('actionName'),  
    // ActionItems[0],  ActionItems.index(-1)
    // などの形式で中身のアクションを取得できる
    var ActionItems = function ActionItems() {
        this.length = 0;
    };
    ActionItems.prototype.getByName = function(nameStr) {
        for (var i = 0, len = this.length; i < len; i++) {
            if (this[i].name == nameStr) {
                return this[i];
            }
        }
    };
    ActionItems.prototype.index = function(keyNumber) {
        var res;
        if (keyNumber >= 0) {
            res = this[keyNumber];
        } else {
            res = this[this.length + keyNumber];
        }
        return res;
    };

    // アクションセット名を取得
    var regExpSetName = /^\/name\s+\[\s+\d+\s+([^\]]+?)\s+\]/m;
    var setName = hexToString(actionCode.match(regExpSetName)[1].replace(/\s+/g, ''));

    // セット内のアクションを取得
    var regExpActionNames = /^\/action-\d+\s+\{\s+\/name\s+\[\s+\d+\s+([^\]]+?)\s+\]/mg;
    var actionItemsObj = new ActionItems();
    var i = 0;
    var matchObj;
    while (matchObj = regExpActionNames.exec(actionCode)) {
        var actionName = hexToString(matchObj[1].replace(/\s+/g, ''));
        var actionObj = new ActionItem(i, actionName, setName);
        actionItemsObj[actionName] = actionObj;
        actionItemsObj[i] = actionObj;
        i++;
        if (i > 1000) {
            break;
        } // limiter
    }
    actionItemsObj.length = i;

    // aiaファイルとして書き出し
    var failed = false;
    var aiaFileObj = new File(Folder.temp + '/tempActionSet.aia');
    try {
        aiaFileObj.open('w');
        aiaFileObj.write(actionCode);
    } catch (e) {
        failed = true;
        alert(e);
        return;
    } finally {
        aiaFileObj.close();
        if (failed) {
            try {
                aiaFileObj.remove();
            } catch (e) {}
        }
    }

    // 同名アクションセットがあったらunloadする。これは余計なお世話かもしれない
    try {
        app.unloadAction(setName, '');
    } catch (e) {}

    // アクションを読み込み実行する
    var actionLoaded = false;
    try {
        app.loadAction(aiaFileObj);
        actionLoaded = true;
        func.call(func, actionItemsObj);
    } catch (e) {
        alert(e);
    } finally {
        // 読み込んだアクションと，そのaiaファイルを削除
        if (actionLoaded) {
            app.unloadAction(setName, '');
        }
        aiaFileObj.remove();
    }
}

function setBlendOption(step, orientationValue) {
  var actionCode = '''
/version 3
/name [ 5
	426c656e64
]
/isOpen 1
/actionCount 1
/action-1 {
	/name [ 7
		73657453746570
	]
	/keyIndex 0
	/colorIndex 0
	/isOpen 1
	/eventCount 1
	/event-1 {
		/useRulersIn1stQuadrant 0
		/internalName (ai_plugin_liveblend)
		/localizedName [ 12
			e38396e383ace383b3e38389
		]
		/isOpen 0
		/isOn 1
		/hasDialog 1
		/showDialog 0
		/parameterCount 3
		/parameter-1 {
			/key 1835363957
			/showInPalette 4294967295
			/type (enumerated)
			/name [ 15
				e382aae38397e382b7e383a7e383b3
			]
			/value 5
		}
		/parameter-2 {
			/key 1937007984
			/showInPalette 4294967295
			/type (integer)
			/value ''' + String(step) + '''
		}
		/parameter-3 {
			/key 1919906913
			/showInPalette 4294967295
			/type (enumerated)
			/name [ 12
				e59e82e79bb4e696b9e59091
			]
			/value ''' + String(orientationValue) + '''
		}
	}
}
''' ;
  
  tempAction(actionCode, function(actionItems) {
    actionItems[0].exec(false);
  });
}