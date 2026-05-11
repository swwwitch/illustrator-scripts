#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

(function () {

    /*
    ### スクリプト名：

    SmartRenamer.jsx

    ### 概要

    - Illustrator のアートボード／シンボル／レイヤー名を、接頭辞・接尾辞・名前ソース・検索置換を組み合わせて一括リネーム／並び替えするスクリプトです。
    - ダイアログ上で設定を変更すると、結果を自動でプレビュー更新します。

    #### 対象タイプ（ダイアログ最上段・group）

    - 「アートボード」「シンボル」「レイヤー」から切り替え可能（中央配置の group UI）
    - 切替時は右カラムの一覧と「フィルター」「固定」パネルの状態がそのタイプ向けに再構築される
    - シンボル／レイヤーでも「固定」パネルを表示し、「元の名称」または指定テキストをベースに接頭辞・接尾辞・検索置換を適用

    #### リネーム条件（左カラム）

    - 接頭辞・接尾辞：連番トークン（`{#1}` → 1,2,3 ／ `{#01}` → 01,02,03 のゼロパディング対応）、`#FN`（ファイル名）、`#DT`（日付）、区切り（`-` `_`）の挿入ボタン
    - 固定：「元の名称」「最前面のテキスト」（アートボード時のみ。レイヤー／グループを再帰スキャンし、アートボード内で最初に見つかった可視・非ロック TextFrame）または「指定」（任意テキスト）
    - 検索・置換：検索文字列と置換文字列。`正規表現` チェックで RegExp 動作（既定 ON）、OFF はリテラル（特殊文字エスケープ）。置換側にもトークンボタン（`#`→`{#1}` ／ `##`→`{#01}` ／ `-` ／ `_` ／ `x` でクリア）。検索側にも `\d` / `\d+` / `.+`（全体マッチ）を入力するショートカットボタンを追加
    - 同名が重複する場合は "_1", "_2" などを自動付加（接頭辞・接尾辞・置換テキストに連番トークンがある場合はユニーク前提でスキップ）

    #### フィルター（右カラム上段パネル）

    - 1 行目：「すべて」または「指定範囲」（`1-3,5` 形式）。並び替え／リネームパネルのチェックと双方向連動
    - 2 行目：「検索でフィルター」チェック＋テキスト。ON のときは対象選択（すべて／指定範囲）を無視し、現在名に該当文字列を含むアイテムだけを対象化
    - フィルター UI は専用パネル化され、右カラム上部へ移動

    #### 並び替え／リネーム（右カラム下段）

    - 各アイテムの「元の名前 → 新しい名前」を一覧表示。［更新］後は確定後の現在名を右カラムの基準名として再取得
    - チェックを付けた行のみ「新しい名前」を手動上書き可能（手動編集後はプレビューで上書きされない）
    - チェック行を「↑先頭へ／↑上へ／↓下へ／↓末尾へ」で並び替え。アートボードは rect 入れ替え、レイヤーは `.move()`、シンボルは並び替え不可
    - チェックボックスで Option+クリック：
      - 全行 OFF 状態：全行 ON
      - 一部 ON 状態：クリック値に合わせて一括切替
      - 全行 ON 状態：クリック行のみ ON で他を OFF（孤立化）

    #### ボタン

    - 「更新」で現在の設定・並び替え・手動編集名を canvas に即時反映。シンボル／レイヤーでも OK 前に即時反映される
    - キャンセル時はダイアログ表示前の名前を全タイプ復元

    ### 更新履歴

    - v1.0 (20250509) : 初期バージョン
    - v1.5.2 (20260508) : ［更新］ボタンで右カラムの手動編集名と並び替えも確定するように変更
    - v1.5.3 (20260508) : 同名重複時、最初のアイテムから "_1" を付加するように変更（重複しない名前はそのまま）
    - v1.6.0 (20260511) : 対象タイプ（アートボード／シンボル／レイヤー）対応、フィルターパネル追加、「元の名称」モード追加、検索でフィルター、検索・置換（正規表現対応）、検索正規表現ショートカット、`{#N}` 連番トークン、Option+クリック孤立化、［更新］後の再ベースライン化、シンボル／レイヤーの即時更新対応を追加

    ---

    ### Script Name:

    SmartRenamer.jsx

    ### Overview

    - Batch rename and reorder Illustrator artboards / symbols / layers by combining prefix, suffix, name source, and find/replace.
    - As settings change in the dialog, results are updated automatically for preview.

    #### Target Type (top dialog group)

    - Switch between Artboard / Symbol / Layer (center-aligned group UI)
    - On switch, the right-side list and Filter / Fixed panels are rebuilt for that type
    - In Symbol / Layer mode the Fixed panel remains visible, and renaming can use "Original Name" or custom text plus prefix / suffix / find-replace

    #### Rename conditions (left column)

    - Prefix / Suffix: sequence tokens (`{#1}` → 1,2,3 ; `{#01}` → 01,02,03 with zero padding), `#FN` (file name), `#DT` (date), separators (`-` `_`)
    - Fixed: "Original Name", "Frontmost Text" (Artboard mode only; recursively scans layers/groups and uses the first visible unlocked TextFrame found inside the artboard), or "Custom" (free text)
    - Find / Replace: find and replace strings. The Regex checkbox switches between RegExp mode (default ON) and literal (escaped) mode. The replace field also has token buttons (`#`→`{#1}`, `##`→`{#01}`, `-`, `_`, `x` to clear). The find field also includes shortcut buttons for `\d`, `\d+`, and `.+` (match-all)
    - Automatically appends "_1", "_2", etc. only on name collision (skipped when prefix / suffix / replace contains a sequence token, since results are assumed unique)

    #### Filter (top-right panel)

    - Row 1: "All" or "Range" (`1-3,5` form). Two-way linked with the Reorder / Rename checkboxes
    - Row 2: "Filter by search" checkbox + text. When ON, ignores the All/Range selection and only checks items whose current name contains the text
    - The filter UI is grouped into its own dedicated panel at the top of the right column

    #### Reorder / Rename (right column, bottom)

    - Lists each item as "Original Name -> New Name". After Refresh, the committed current names are reloaded as the right-column baseline names
    - Checked rows allow manual override of the new name (manual edits stay sticky)
    - Move checked rows with Top / Up / Down / Bottom. Artboards swap rects, layers use `.move()`, symbols are not reorderable
    - Option-click a checkbox:
      - When no rows are checked: turn all rows ON
      - When some rows are checked: bulk toggle to the new value
      - When all rows are checked: isolate (only the clicked row stays ON)

    #### Buttons

    - "Refresh" immediately commits current settings, reorder state, and manual-edited names. Symbols and layers are also updated immediately before pressing OK
    - Cancelling restores pre-dialog names for all three target types

    ### Update History

    - v1.0 (20250509): Initial version
    - v1.5.2 (20260508): Refresh now commits right-column manual name edits and reorder changes
    - v1.5.3 (20260508): On name collisions, the first occurrence now starts at "_1" (non-duplicates remain untouched)
    - v1.6.0 (20260511): Added target type (artboard / symbol / layer), dedicated filter panel, "Original Name" mode, search filter, regex shortcut buttons, find / replace (regex-capable), `{#N}` sequence tokens, Option-click isolate, post-Refresh re-baselining, and immediate symbol/layer Refresh updates
    */

    // =========================================
    // バージョンとローカライズ
    // =========================================

    var SCRIPT_VERSION = "v1.6.0";

    var lang = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

    /* ローカライズキーから現在言語のラベル文字列を取得 / Look up a localized label by key */
    function L(key) {
        if (!LABELS[key]) return key;
        return LABELS[key][lang] || LABELS[key].en || key;
    }

    // =========================================
    // ラベル定義 / Japanese-English label definitions
    // =========================================

    var LABELS = {
        dialogTitle: {
            ja: "スマートリネーム",
            en: "Smart Renamer"
        },
        prefixLabel: {
            ja: "接頭辞",
            en: "Prefix"
        },
        suffixLabel: {
            ja: "接尾辞",
            en: "Suffix"
        },
        sourceLabel: {
            ja: "固定",
            en: "Fixed"
        },
        basicLabel: {
            ja: "リネーム",
            en: "Rename"
        },
        originalNameOption: {
            ja: "元の名称",
            en: "Original Name"
        },
        frontmostOption: {
            ja: "最前面のテキスト",
            en: "Frontmost Text"
        },
        customOption: {
            ja: "指定",
            en: "Custom"
        },
        targetLabel: {
            ja: "対象",
            en: "Target"
        },
        filterPanelLabel: {
            ja: "フィルター",
            en: "Filter"
        },
        targetTypeLabel: {
            ja: "対象",
            en: "Target"
        },
        targetTypeArtboard: {
            ja: "アートボード",
            en: "Artboard"
        },
        targetTypeSymbol: {
            ja: "シンボル",
            en: "Symbol"
        },
        targetTypeLayer: {
            ja: "レイヤー",
            en: "Layer"
        },
        allBoards: {
            ja: "すべて",
            en: "All"
        },
        specificBoards: {
            ja: "指定範囲",
            en: "Range"
        },
        searchLabel: {
            ja: "検索でフィルター",
            en: "Filter by search"
        },
        searchHelpTip: {
            ja: "テキストに照合するものだけに絞り込み",
            en: "Narrow down to only those matching the text"
        },
        findReplaceLabel: {
            ja: "検索・置換",
            en: "Find / Replace"
        },
        findLabel: {
            ja: "検索",
            en: "Find"
        },
        replaceLabel: {
            ja: "置換",
            en: "Replace"
        },
        regexOption: {
            ja: "正規表現",
            en: "Regex"
        },
        cancelButton: {
            ja: "キャンセル",
            en: "Cancel"
        },
        okButton: {
            ja: "OK",
            en: "OK"
        },
        alertNeedSettings: {
            ja: "接頭辞・接尾辞のいずれかを入力するか、アートボード名を指定してください。",
            en: "Enter a prefix or suffix, or choose an artboard name source."
        },
        reorderLabel: {
            ja: "並び替え／リネーム",
            en: "Reorder / Rename"
        },
        orderHeader: {
            ja: "順",
            en: "#"
        },
        targetHeader: {
            ja: "対象",
            en: "Sel"
        },
        originalNameHeader: {
            ja: "元の名前",
            en: "Original Name"
        },
        newNameHeader: {
            ja: "新しい名前",
            en: "New Name"
        },
        moveTopBtn: {
            ja: "↑ 先頭へ",
            en: "↑ Top"
        },
        moveUpBtn: {
            ja: "↑ 上へ",
            en: "↑ Up"
        },
        moveDownBtn: {
            ja: "↓ 下へ",
            en: "↓ Down"
        },
        moveBottomBtn: {
            ja: "↓ 末尾へ",
            en: "↓ Bottom"
        },
        emptyNameAlert: {
            ja: "{n} 番目の新しい名前が空です。名前を入力してください。",
            en: "Artboard {n}: new name is empty. Please enter a name."
        },
        refreshBtn: {
            ja: "更新",
            en: "Refresh"
        }
    };

    // =========================================
    // UI ヘルパー / UI helpers
    // =========================================

    var PANEL_MARGINS = [15, 20, 15, 15];

    /* パネルの共通マージン・スペーシングを設定 / Apply shared panel margin and spacing */
    function setupPanel(panel, spacing) {
        panel.margins = PANEL_MARGINS;
        if (typeof spacing !== "undefined") panel.spacing = spacing;
    }

    // =========================================
    // 設定の読み書きとダイアログイベント / Settings I/O and dialog events
    // =========================================

    /* ダイアログ表示前の状態（名前・アートボードrect・レイヤー順）をキャプチャ / Capture pre-dialog state (names, artboard rects, layer order) */
    function captureOriginalState(doc) {
        var state = {
            artboard: { names: [], rects: [] },
            symbol: { names: [] },
            layer: { names: [], refs: [] }
        };
        for (var ai = 0; ai < doc.artboards.length; ai++) {
            state.artboard.names.push(doc.artboards[ai].name);
            state.artboard.rects.push(doc.artboards[ai].artboardRect);
        }
        for (var si = 0; si < doc.symbols.length; si++) {
            state.symbol.names.push(doc.symbols[si].name);
        }
        for (var li = 0; li < doc.layers.length; li++) {
            state.layer.names.push(doc.layers[li].name);
            state.layer.refs.push(doc.layers[li]);
        }
        return state;
    }

    /* キャンセル時にダイアログ表示前の状態を復元（名前・アートボードrect・レイヤー順） / Restore pre-dialog state on cancel */
    function restoreOriginalState(doc, originalState) {
        // アートボード：rect と名前を復元（一時名で衝突回避）
        var artboards = doc.artboards;
        var artboardCount = Math.min(artboards.length, originalState.artboard.names.length);
        for (var tmpIdx = 0; tmpIdx < artboardCount; tmpIdx++) {
            try { artboards[tmpIdx].name = "__tmp_ab_restore_" + tmpIdx + "__"; } catch (tmpError) { $.writeln("[SmartRenamer] tmp name failed: " + tmpError); }
        }
        for (var abIdx = 0; abIdx < artboardCount; abIdx++) {
            try {
                artboards[abIdx].artboardRect = originalState.artboard.rects[abIdx];
                artboards[abIdx].name = originalState.artboard.names[abIdx];
            } catch (abError) { $.writeln("[SmartRenamer] artboard restore failed at " + abIdx + ": " + abError); }
        }

        // シンボル：インデックス順で復元（並び替え非対応）
        var symbols = doc.symbols;
        var symbolCount = Math.min(symbols.length, originalState.symbol.names.length);
        for (var symIdx = 0; symIdx < symbolCount; symIdx++) {
            try { symbols[symIdx].name = originalState.symbol.names[symIdx]; } catch (symError) { $.writeln("[SmartRenamer] symbol restore failed at " + symIdx + ": " + symError); }
        }

        // レイヤー：安定参照を使って元の順序へ戻し、各参照に元の名前を当てる
        var layerRefs = originalState.layer.refs;
        for (var reverseIdx = layerRefs.length - 1; reverseIdx >= 0; reverseIdx--) {
            try { layerRefs[reverseIdx].move(doc, ElementPlacement.PLACEATBEGINNING); } catch (layerOrderError) { $.writeln("[SmartRenamer] layer order restore failed at " + reverseIdx + ": " + layerOrderError); }
        }
        for (var nameIdx = 0; nameIdx < layerRefs.length; nameIdx++) {
            try { layerRefs[nameIdx].name = originalState.layer.names[nameIdx]; } catch (layerNameError) { $.writeln("[SmartRenamer] layer name restore failed at " + nameIdx + ": " + layerNameError); }
        }
    }

    /* ダイアログ各コントロールから設定オブジェクトを構築 / Read settings object from dialog controls */
    /* rangeMode / rangeText は itemEntries の checked 状態から実効値を導出する（検索フィルタ・手動チェック後の選択を反映） */
    function readDialogSettings(dialogUI) {
        var mode = dialogUI.frontmostRadio.value ? "frontmost"
            : (dialogUI.originalNameRadio.value ? "original" : "custom");
        var targetType = dialogUI.targetTypeSymbolRadio.value ? "symbol"
            : (dialogUI.targetTypeLayerRadio.value ? "layer" : "artboard");

        var settings = {
            mode: mode,
            targetType: targetType,
            prefix: dialogUI.prefixInput.text,
            suffix: dialogUI.suffixInput.text,
            customText: dialogUI.customInput.text,
            rangeMode: dialogUI.targetAllRadio.value ? "all" : "numbered",
            rangeText: dialogUI.rangeInput.text,
            findText: dialogUI.findInput.text,
            replaceText: dialogUI.replaceInput.text,
            useRegex: dialogUI.regexCheckbox.value,
        };

        if (dialogUI.getItemEntries) {
            var entries = dialogUI.getItemEntries();
            var checkedIndices = [];
            for (var i = 0; i < entries.length; i++) {
                if (entries[i].checked) checkedIndices.push(entries[i].originalIndex);
            }
            checkedIndices.sort(function (a, b) { return a - b; });
            if (entries.length > 0 && checkedIndices.length === entries.length) {
                settings.rangeMode = "all";
                settings.rangeText = "";
            } else {
                settings.rangeMode = "numbered";
                settings.rangeText = buildRangeString(checkedIndices);
            }
        }

        return settings;
    }

    /* 検索・置換を適用（正規表現対応） / Apply find/replace (regex-capable) */
    function applyFindReplace(name, findPattern, replaceText, useRegex) {
        if (!findPattern) return name;
        try {
            var regex;
            var effectiveReplace = replaceText;
            if (useRegex) {
                regex = new RegExp(findPattern, "g");
            } else {
                var escapedPattern = findPattern.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
                regex = new RegExp(escapedPattern, "g");
                // リテラル置換では replaceText 中の $ も無効化する（$1 や $& を特殊解釈させない）
                effectiveReplace = replaceText.replace(/\$/g, "$$$$");
            }
            return name.replace(regex, effectiveReplace);
        } catch (regexError) {
            return name;
        }
    }

    /* ダイアログ各コントロールにイベントハンドラを設定 / Wire dialog event handlers */
    function bindDialogEvents(dialogUI, doc) {
        function updatePreview() {
            // canvas は触らず、未確定プレビュー名を計算して右カラムを更新
            var previewNames = computePreviewNames(doc, readDialogSettings(dialogUI));
            if (dialogUI.syncPreviewToReorderRows) {
                dialogUI.syncPreviewToReorderRows(previewNames);
            }
        }

        function commitCurrentSettings() {
            /* 現在の設定と右カラムの手動編集を canvas に確定（[更新]） / Commit current settings and right-column manual edits to the canvas (Refresh) */
            if (dialogUI.syncEditingValues) {
                dialogUI.syncEditingValues();
            }

            var settings = readDialogSettings(dialogUI);
            if (dialogUI.getItemEntries) {
                settings.itemEntries = dialogUI.getItemEntries();
            }

            var hasManualOrReorder = settings.itemEntries && hasReorderOrRename(settings.itemEntries);
            var committed = false;

            if (hasManualOrReorder) {
                applyReorderAndRename(doc, settings.itemEntries, settings);
                committed = true;
            } else if (executeRename(doc, settings, { silent: true })) {
                committed = true;
            }

            if (committed) {
                app.redraw();
                if (dialogUI.rebaselineEntriesAfterCommit) {
                    dialogUI.rebaselineEntriesAfterCommit();
                }
            }
            updatePreview();
        }

        function syncTargetInput() {
            dialogUI.rangeInput.enabled = dialogUI.targetRangeRadio.value;
            if (dialogUI.applyTargetToCheckboxes) {
                if (dialogUI.targetAllRadio.value) {
                    dialogUI.applyTargetToCheckboxes("all", "");
                } else {
                    dialogUI.applyTargetToCheckboxes("numbered", dialogUI.rangeInput.text);
                }
            }
            updatePreview();
        }
        dialogUI.targetAllRadio.onClick = function () {
            dialogUI.targetRangeRadio.value = false;
            syncTargetInput();
        };
        dialogUI.targetRangeRadio.onClick = function () {
            dialogUI.targetAllRadio.value = false;
            syncTargetInput();
        };

        /* 対象タイプ切替後に呼ばれる：「すべて」適用とプレビュー更新 / Post type-switch: apply "all" target and refresh preview */
        if (dialogUI.setTypeChangeCallback) {
            dialogUI.setTypeChangeCallback(function () {
                if (dialogUI.applyTargetToCheckboxes) {
                    dialogUI.applyTargetToCheckboxes("all", "");
                }
                updatePreview();
            });
        }

        function selectTargetTypeRadio(type) {
            dialogUI.targetTypeArtboardRadio.value = (type === "artboard");
            dialogUI.targetTypeSymbolRadio.value = (type === "symbol");
            dialogUI.targetTypeLayerRadio.value = (type === "layer");
            dialogUI.setTargetType(type);
        }
        dialogUI.targetTypeArtboardRadio.onClick = function () { selectTargetTypeRadio("artboard"); };
        dialogUI.targetTypeSymbolRadio.onClick = function () { selectTargetTypeRadio("symbol"); };
        dialogUI.targetTypeLayerRadio.onClick = function () { selectTargetTypeRadio("layer"); };

        var sourceRadios = [dialogUI.originalNameRadio, dialogUI.customRadio, dialogUI.frontmostRadio];

        dialogUI.originalNameRadio.onClick = function () {
            selectSourceRadio(dialogUI.originalNameRadio);
        };

        function selectSourceRadio(selected) {
            for (var i = 0; i < sourceRadios.length; i++) {
                sourceRadios[i].value = (sourceRadios[i] === selected);
            }
            dialogUI.customInput.enabled = dialogUI.customRadio.value;
            if (dialogUI.customRadio.value) {
                try { dialogUI.customInput.active = true; } catch (focusError) { }
            }
            updatePreview();
        }
        dialogUI.customRadio.onClick = function () { selectSourceRadio(dialogUI.customRadio); };
        dialogUI.frontmostRadio.onClick = function () { selectSourceRadio(dialogUI.frontmostRadio); };
        dialogUI.customInput.onChange = function () {
            if (dialogUI.customRadio.value) updatePreview();
        };

        dialogUI.prefixInput.onChange = updatePreview;
        dialogUI.suffixInput.onChange = updatePreview;
        dialogUI.rangeInput.onChange = function () {
            if (dialogUI.targetRangeRadio.value) {
                if (dialogUI.applyTargetToCheckboxes) {
                    dialogUI.applyTargetToCheckboxes("numbered", dialogUI.rangeInput.text);
                }
                updatePreview();
            }
        };

        /* 現在の対象モード（すべて／指定範囲）を再適用（検索フィルタも反映） / Re-apply current target mode (with search filter) */
        function reapplyCurrentTarget() {
            if (!dialogUI.applyTargetToCheckboxes) return;
            if (dialogUI.targetAllRadio.value) {
                dialogUI.applyTargetToCheckboxes("all", "");
            } else {
                dialogUI.applyTargetToCheckboxes("numbered", dialogUI.rangeInput.text);
            }
        }

        dialogUI.searchFilterCheckbox.onClick = function () {
            var filterOn = dialogUI.searchFilterCheckbox.value;
            dialogUI.searchInput.enabled = filterOn;
            // フィルタ ON 中は「すべて／指定範囲」を無効化し、状態が紛らわしくならないようにする
            dialogUI.targetAllRadio.enabled = !filterOn;
            dialogUI.targetRangeRadio.enabled = !filterOn;
            dialogUI.rangeInput.enabled = !filterOn && dialogUI.targetRangeRadio.value;
            if (filterOn) {
                try { dialogUI.searchInput.active = true; } catch (focusError) { }
            }
            reapplyCurrentTarget();
            updatePreview();
        };
        dialogUI.searchInput.onChange = function () {
            if (dialogUI.searchFilterCheckbox.value) {
                reapplyCurrentTarget();
                updatePreview();
            }
        };

        dialogUI.findInput.onChange = updatePreview;
        dialogUI.replaceInput.onChange = updatePreview;
        dialogUI.regexCheckbox.onClick = updatePreview;

        if (dialogUI.refreshBtn) {
            dialogUI.refreshBtn.onClick = function () {
                /* edittext のフォーカス未解除でも値を確実に反映させる / Force onChange to commit pending edittext value */
                try { dialogUI.rangeInput.notify("onChange"); } catch (notifyError) { }
                commitCurrentSettings();
            };
        }

        // チェックボックス操作 → 対象設定 への逆同期に updatePreview を注入
        if (dialogUI.setRequestPreviewUpdate) {
            dialogUI.setRequestPreviewUpdate(updatePreview);
        }

        // 初期同期：対象設定 → チェックボックス
        if (dialogUI.applyTargetToCheckboxes) {
            if (dialogUI.targetAllRadio.value) {
                dialogUI.applyTargetToCheckboxes("all", "");
            } else {
                dialogUI.applyTargetToCheckboxes("numbered", dialogUI.rangeInput.text);
            }
        }

        // 初期プレビュー
        updatePreview();
    }

    /* 接頭辞・接尾辞用のトークン挿入ボタンを追加（最後に対象フィールドのクリアボタン x を配置） / Add prefix/suffix token-insert buttons (with a trailing clear "x" button) */
    function addTokenButtons(panel, targetInput) {
        var tokens = [
            { label: "1", value: "{#1}", width: 22 },
            { label: "01", value: "{#01}" },
            { label: "-", value: "-", width: 22 },
            { label: "_", value: "_", width: 22 },
            { label: "#FN", value: "#FN", width: 40 },
            { label: "#DT", value: "#DT", width: 40 }
        ];
        var tokenButtonRow = panel.add("group");
        tokenButtonRow.orientation = "row";
        tokenButtonRow.spacing = 4;
        tokenButtonRow.margins = 0;
        for (var tokenIdx = 0; tokenIdx < tokens.length; tokenIdx++) {
            (function (token) {
                var tokenBtn = tokenButtonRow.add("button", undefined, token.label);
                tokenBtn.preferredSize = [token.width || 28, 20];
                tokenBtn.onClick = function () {
                    targetInput.text = targetInput.text + token.value;
                    targetInput.notify("onChange");
                };
            })(tokens[tokenIdx]);
        }
        var clearFieldBtn = tokenButtonRow.add("button", undefined, "x");
        clearFieldBtn.preferredSize = [22, 20];
        clearFieldBtn.onClick = function () {
            targetInput.text = "";
            targetInput.notify("onChange");
        };
    }

    // =========================================
    // ダイアログ構築と表示 / Dialog build and show
    // =========================================

    /* リネームダイアログのUIを構築 / Build the rename dialog UI */
    function createRenameDialog(doc) {
        var dialog = new Window("dialog", L('dialogTitle') + ' ' + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = "fill";

        // 対象タイプ（ダイアログ最上段：アートボード／シンボル／レイヤー）
        var targetTypePanel = dialog.add("group");
        targetTypePanel.orientation = "row";
        targetTypePanel.alignChildren = ["center", "center"];
        targetTypePanel.alignment = ["fill", "top"];
        targetTypePanel.margins = [0, 0, 0, 0];

        var targetTypeArtboardRadio = targetTypePanel.add("radiobutton", undefined, L('targetTypeArtboard'));
        var targetTypeSymbolRadio = targetTypePanel.add("radiobutton", undefined, L('targetTypeSymbol'));
        var targetTypeLayerRadio = targetTypePanel.add("radiobutton", undefined, L('targetTypeLayer'));
        targetTypeArtboardRadio.value = true;

        // コンテンツ行（左：リネーム条件、右：対象＋並び替え／リネーム）
        var contentRow = dialog.add("group");
        contentRow.orientation = "row";
        contentRow.alignChildren = "top";

        var leftColumn = contentRow.add("group");
        leftColumn.orientation = "column";
        leftColumn.alignChildren = "fill";

        var basicPanel = leftColumn.add("panel", undefined, L('basicLabel'));
        basicPanel.orientation = "column";
        basicPanel.alignChildren = "fill";
        setupPanel(basicPanel);

        /* 接頭辞入力欄 / Prefix input */
        var prefixPanel = basicPanel.add("panel", undefined, L('prefixLabel'));
        prefixPanel.orientation = "column";
        prefixPanel.alignChildren = "left";
        setupPanel(prefixPanel);
        var prefixInput = prefixPanel.add("edittext", undefined, "");
        prefixInput.characters = 14;
        prefixInput.active = true;
        addTokenButtons(prefixPanel, prefixInput);

        // アートボード名
        var sourcePanel = basicPanel.add("panel", undefined, L('sourceLabel'));
        sourcePanel.orientation = "column";
        sourcePanel.alignChildren = "left";
        setupPanel(sourcePanel);

        var originalNameRadio = sourcePanel.add("radiobutton", undefined, L('originalNameOption'));

        var customRow = sourcePanel.add("group");
        customRow.orientation = "row";
        customRow.alignChildren = ["left", "center"];
        var customRadio = customRow.add("radiobutton", undefined, L('customOption'));
        var customInput = customRow.add("edittext", undefined, "");
        customInput.characters = 12;
        customInput.enabled = false;

        var frontmostRadio = sourcePanel.add("radiobutton", undefined, L('frontmostOption'));

        // 全ラジオを追加した後に初期値を設定
        originalNameRadio.value = false;
        customRadio.value = false;
        frontmostRadio.value = true;

        // 検索・置換
        var findReplacePanel = basicPanel.add("panel", undefined, L('findReplaceLabel'));
        findReplacePanel.orientation = "column";
        findReplacePanel.alignChildren = "fill";
        setupPanel(findReplacePanel);

        var findRow = findReplacePanel.add("group");
        findRow.orientation = "row";
        findRow.alignChildren = ["left", "center"];
        var findLabelStatic = findRow.add("statictext", undefined, L('findLabel'));
        findLabelStatic.preferredSize.width = 36;
        var findInput = findRow.add("edittext", undefined, "");
        findInput.characters = 14;

        // --- Find pattern shortcut buttons ---
        var findPatternRow = findReplacePanel.add("group");
        findPatternRow.orientation = "row";
        findPatternRow.spacing = 4;
        var findPatternTokens = [
            { label: "#", value: "\\d", width: 22 },
            { label: "##", value: "\\d+", width: 28 },
            { label: "*", value: ".+", width: 22 }
        ];
        for (var findPatternIdx = 0; findPatternIdx < findPatternTokens.length; findPatternIdx++) {
            (function (token) {
                var patternBtn = findPatternRow.add("button", undefined, token.label);
                patternBtn.preferredSize = [token.width || 28, 20];
                patternBtn.onClick = function () {
                    findInput.text = findInput.text + token.value;
                    findInput.notify("onChange");
                };
            })(findPatternTokens[findPatternIdx]);
        }

        var replaceRow = findReplacePanel.add("group");
        replaceRow.orientation = "row";
        replaceRow.alignChildren = ["left", "center"];
        var replaceLabelStatic = replaceRow.add("statictext", undefined, L('replaceLabel'));
        replaceLabelStatic.preferredSize.width = 36;
        var replaceInput = replaceRow.add("edittext", undefined, "");
        replaceInput.characters = 14;

        var replaceTokenRow = findReplacePanel.add("group");
        replaceTokenRow.orientation = "row";
        replaceTokenRow.spacing = 4;
        var replaceTokens = [
            { label: "#", value: "{#1}", width: 22 },
            { label: "##", value: "{#01}", width: 28 },
            { label: "-", value: "-", width: 22 },
            { label: "_", value: "_", width: 22 }
        ];
        for (var replaceTokenIdx = 0; replaceTokenIdx < replaceTokens.length; replaceTokenIdx++) {
            (function (token) {
                var btn = replaceTokenRow.add("button", undefined, token.label);
                btn.preferredSize = [token.width || 28, 20];
                btn.onClick = function () {
                    replaceInput.text = replaceInput.text + token.value;
                    replaceInput.notify("onChange");
                };
            })(replaceTokens[replaceTokenIdx]);
        }
        var replaceClearBtn = replaceTokenRow.add("button", undefined, "x");
        replaceClearBtn.preferredSize = [22, 20];
        replaceClearBtn.onClick = function () {
            replaceInput.text = "";
            replaceInput.notify("onChange");
        };
        var regexSpacer = replaceTokenRow.add("group");
        regexSpacer.preferredSize = [20, 1];
        var regexCheckbox = replaceTokenRow.add("checkbox", undefined, L('regexOption'));
        regexCheckbox.value = true;

        // 接尾辞
        var suffixPanel = basicPanel.add("panel", undefined, L('suffixLabel'));
        suffixPanel.orientation = "column";
        suffixPanel.alignChildren = "left";
        setupPanel(suffixPanel);
        var suffixInput = suffixPanel.add("edittext", undefined, "");
        suffixInput.characters = 16;
        addTokenButtons(suffixPanel, suffixInput);

        // 右カラム：対象＋並び替え／リネーム
        var rightColumn = contentRow.add("group");
        rightColumn.orientation = "column";
        rightColumn.alignChildren = "fill";

        // フィルター（右カラム上段）
        var targetPanel = rightColumn.add("panel", undefined, L('filterPanelLabel'));
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["left", "center"];
        setupPanel(targetPanel, 6);

        var targetRangeRow = targetPanel.add("group");
        targetRangeRow.orientation = "row";
        targetRangeRow.alignChildren = ["left", "center"];
        var targetAllRadio = targetRangeRow.add("radiobutton", undefined, L('allBoards'));
        var targetRangeRadio = targetRangeRow.add("radiobutton", undefined, L('specificBoards'));
        var rangeInput = targetRangeRow.add("edittext", undefined, "");
        rangeInput.characters = 10;
        rangeInput.enabled = false;
        targetAllRadio.value = true;
        targetRangeRadio.value = false;

        var targetSearchRow = targetPanel.add("group");
        targetSearchRow.orientation = "row";
        targetSearchRow.alignChildren = ["left", "center"];
        var searchFilterCheckbox = targetSearchRow.add("checkbox", undefined, L('searchLabel'));
        searchFilterCheckbox.value = false;
        var searchInput = targetSearchRow.add("edittext", undefined, "");
        searchInput.characters = 10;
        searchInput.helpTip = L('searchHelpTip');
        searchInput.enabled = false;

        var reorderPanel = rightColumn.add("panel", undefined, L('reorderLabel'));
        reorderPanel.orientation = "column";
        reorderPanel.alignChildren = "fill";
        setupPanel(reorderPanel, 6);

        var currentTargetType = "artboard";

        function buildEntriesForType(type) {
            var items = getDocumentItems(doc, type);
            var entries = [];
            for (var i = 0; i < items.length; i++) {
                var entry = {
                    originalIndex: i,
                    name: items[i].name,
                    newName: items[i].name,
                    checked: false,
                    userEdited: false
                };
                if (type === "artboard") {
                    entry.rect = items[i].artboardRect;
                }
                entries.push(entry);
            }
            return entries;
        }

        var itemEntries = buildEntriesForType(currentTargetType);

        // bindDialogEvents から updatePreview / 対象タイプ切替後処理を注入してもらうコールバック
        var requestPreviewUpdate = null;
        var typeChangeCallback = null;

        /* ［更新］確定後の再ベースライン化：entries の name / originalIndex を canvas 現状に合わせ、userEdited を解除 / Re-baseline entries to current canvas after Refresh; clear userEdited so future previews pick up new prefix/suffix changes */
        function rebaselineEntriesAfterCommit() {
            var items = getDocumentItems(doc, currentTargetType);
            for (var i = 0; i < itemEntries.length && i < items.length; i++) {
                itemEntries[i].originalIndex = i;
                itemEntries[i].name = items[i].name;
                itemEntries[i].newName = items[i].name;
                itemEntries[i].userEdited = false;
                if (currentTargetType === "artboard") {
                    itemEntries[i].rect = items[i].artboardRect;
                }
            }
        }

        /* 対象タイプ切替：エントリを再構築し UI 状態をリセット / Switch target type: rebuild entries and reset UI state */
        function setTargetType(type) {
            currentTargetType = type;
            itemEntries = buildEntriesForType(type);

            // 対象パネルを「すべて」に初期化（検索フィルタが ON なら依然として無効）
            var filterOn = searchFilterCheckbox.value;
            targetAllRadio.value = true;
            targetRangeRadio.value = false;
            rangeInput.text = "";
            targetAllRadio.enabled = !filterOn;
            targetRangeRadio.enabled = !filterOn;
            rangeInput.enabled = false;

            // 「固定」パネルはすべての対象タイプで表示
            sourcePanel.visible = true;
            sourcePanel.maximumSize.height = 1000;
            sourcePanel.minimumSize.height = 0;

            // 「最前面のテキスト」はアートボード時のみ有効。シンボル／レイヤーでは「指定」を使う
            frontmostRadio.enabled = (type === "artboard");
            if (!frontmostRadio.enabled && frontmostRadio.value) {
                frontmostRadio.value = false;
                customRadio.value = true;
            }
            customInput.enabled = customRadio.value;

            // パネル高さ変更後にレイアウトを再計算
            leftColumn.layout.layout(true);
            contentRow.layout.layout(true);
            dialog.layout.layout(true);
            refreshReorderRows();
            if (typeChangeCallback) typeChangeCallback();
        }

        /* チェック状態 → 対象設定（ラジオ＋指定範囲テキスト）への逆同期 / Sync target settings from checkbox state */
        function syncTargetFromCheckboxes() {
            var checkedPositions = [];
            for (var i = 0; i < itemEntries.length; i++) {
                if (itemEntries[i].checked) checkedPositions.push(i);
            }
            if (checkedPositions.length === itemEntries.length) {
                targetAllRadio.value = true;
                targetRangeRadio.value = false;
                rangeInput.enabled = false;
            } else {
                targetAllRadio.value = false;
                targetRangeRadio.value = true;
                rangeInput.enabled = true;
                /* 並び替え後の表示位置を基準に指定範囲へ反映（チェック0件なら空文字） / Reflect reordered display positions in the range field (empty when no rows checked) */
                rangeInput.text = buildRangeString(checkedPositions);
            }
        }

        var entryRows = [];

        var entryRowsHost = reorderPanel.add("group");
        entryRowsHost.orientation = "column";
        entryRowsHost.alignChildren = ["fill", "top"];
        entryRowsHost.spacing = 4;
        entryRowsHost.maximumSize.height = 320;

        function syncEditingValues() {
            for (var rowIdx = 0; rowIdx < entryRows.length; rowIdx++) {
                var row = entryRows[rowIdx];
                itemEntries[row.dataIndex].checked = row.checkbox.value;
                if (row.checkbox.value) {
                    itemEntries[row.dataIndex].newName = row.newNameField.text;
                }
            }
        }

        function swapEntries(indexA, indexB) {
            var temp = itemEntries[indexA];
            itemEntries[indexA] = itemEntries[indexB];
            itemEntries[indexB] = temp;
        }

        function moveCheckedUp() {
            syncEditingValues();
            for (var i = 1; i < itemEntries.length; i++) {
                if (itemEntries[i].checked && !itemEntries[i - 1].checked) swapEntries(i, i - 1);
            }
            refreshReorderRows();
        }

        function moveCheckedDown() {
            syncEditingValues();
            for (var i = itemEntries.length - 2; i >= 0; i--) {
                if (itemEntries[i].checked && !itemEntries[i + 1].checked) swapEntries(i, i + 1);
            }
            refreshReorderRows();
        }

        function moveCheckedToTop() {
            syncEditingValues();
            var checkedEntries = [], uncheckedEntries = [];
            for (var i = 0; i < itemEntries.length; i++) {
                if (itemEntries[i].checked) checkedEntries.push(itemEntries[i]);
                else uncheckedEntries.push(itemEntries[i]);
            }
            itemEntries = checkedEntries.concat(uncheckedEntries);
            refreshReorderRows();
        }

        function moveCheckedToBottom() {
            syncEditingValues();
            var checkedEntries = [], uncheckedEntries = [];
            for (var i = 0; i < itemEntries.length; i++) {
                if (itemEntries[i].checked) checkedEntries.push(itemEntries[i]);
                else uncheckedEntries.push(itemEntries[i]);
            }
            itemEntries = uncheckedEntries.concat(checkedEntries);
            refreshReorderRows();
        }

        function canMoveUp() {
            for (var i = 1; i < itemEntries.length; i++) {
                if (itemEntries[i].checked && !itemEntries[i - 1].checked) return true;
            }
            return false;
        }
        function canMoveDown() {
            for (var i = 0; i < itemEntries.length - 1; i++) {
                if (itemEntries[i].checked && !itemEntries[i + 1].checked) return true;
            }
            return false;
        }

        function updateMoveButtonsState() {
            var canMoveUpwards = canMoveUp();
            var canMoveDownwards = canMoveDown();
            moveToTopBtn.enabled = canMoveUpwards;
            moveUpBtn.enabled = canMoveUpwards;
            moveDownBtn.enabled = canMoveDownwards;
            moveToBottomBtn.enabled = canMoveDownwards;
        }

        function refreshReorderRows() {
            while (entryRowsHost.children.length > 0) {
                entryRowsHost.remove(entryRowsHost.children[0]);
            }
            entryRows = [];

            var headerRow = entryRowsHost.add("group");
            headerRow.orientation = "row";
            headerRow.spacing = 6;
            var orderHeaderLbl = headerRow.add("statictext", undefined, L('orderHeader'));
            var targetHeaderLbl = headerRow.add("statictext", undefined, L('targetHeader'));
            var currentNameHeaderLbl = headerRow.add("statictext", undefined, L('originalNameHeader'));
            var arrowHeaderLbl = headerRow.add("statictext", undefined, "→");
            var newNameHeaderLbl = headerRow.add("statictext", undefined, L('newNameHeader'));
            orderHeaderLbl.preferredSize.width = 24;
            targetHeaderLbl.preferredSize.width = 28;
            currentNameHeaderLbl.preferredSize.width = 140;
            arrowHeaderLbl.preferredSize.width = 14;
            newNameHeaderLbl.preferredSize.width = 160;

            for (var entryIndex = 0; entryIndex < itemEntries.length; entryIndex++) {
                (function (idx) {
                    var row = entryRowsHost.add("group");
                    row.orientation = "row";
                    row.alignChildren = ["left", "center"];
                    row.spacing = 6;

                    var orderLabel = row.add("statictext", undefined, (idx + 1) + "");
                    orderLabel.preferredSize.width = 24;

                    var rowCheckbox = row.add("checkbox", undefined, "");
                    rowCheckbox.value = itemEntries[idx].checked;
                    rowCheckbox.preferredSize.width = 28;

                    var currentNameLabel = row.add("statictext", undefined, itemEntries[idx].name);
                    currentNameLabel.preferredSize.width = 140;
                    currentNameLabel.helpTip = itemEntries[idx].name;

                    row.add("statictext", undefined, "→").preferredSize.width = 14;

                    var newNameField = row.add("edittext", undefined, itemEntries[idx].newName);
                    newNameField.preferredSize.width = 160;
                    newNameField.enabled = itemEntries[idx].checked;

                    rowCheckbox.onClick = function () {
                        var newCheckedValue = rowCheckbox.value;
                        var optionKeyHeld = ScriptUI.environment && ScriptUI.environment.keyboardState && ScriptUI.environment.keyboardState.altKey;
                        if (optionKeyHeld) {
                            // 全選択中に Option+クリックされたら、クリック行のみ ON で他はすべて OFF（孤立化）
                            var allWereOn = itemEntries.length > 0;
                            for (var preIdx = 0; preIdx < itemEntries.length; preIdx++) {
                                if (!itemEntries[preIdx].checked) { allWereOn = false; break; }
                            }
                            if (allWereOn) {
                                for (var isoIdx = 0; isoIdx < itemEntries.length; isoIdx++) {
                                    itemEntries[isoIdx].checked = (isoIdx === idx);
                                    if (isoIdx !== idx) {
                                        itemEntries[isoIdx].newName = itemEntries[isoIdx].name;
                                        itemEntries[isoIdx].userEdited = false;
                                    }
                                }
                                for (var isoRowIdx = 0; isoRowIdx < entryRows.length; isoRowIdx++) {
                                    var isoOtherRow = entryRows[isoRowIdx];
                                    var isoOtherEntry = itemEntries[isoOtherRow.dataIndex];
                                    isoOtherRow.checkbox.value = isoOtherEntry.checked;
                                    isoOtherRow.newNameField.enabled = isoOtherEntry.checked;
                                    if (!isoOtherEntry.checked) {
                                        isoOtherRow.newNameField.text = isoOtherEntry.name;
                                    }
                                }
                            } else {
                                for (var entryIdx = 0; entryIdx < itemEntries.length; entryIdx++) {
                                    itemEntries[entryIdx].checked = newCheckedValue;
                                    if (!newCheckedValue) {
                                        itemEntries[entryIdx].newName = itemEntries[entryIdx].name;
                                        itemEntries[entryIdx].userEdited = false;
                                    }
                                }
                                for (var rowIdx = 0; rowIdx < entryRows.length; rowIdx++) {
                                    var otherRow = entryRows[rowIdx];
                                    var otherEntry = itemEntries[otherRow.dataIndex];
                                    otherRow.checkbox.value = newCheckedValue;
                                    otherRow.newNameField.enabled = newCheckedValue;
                                    if (!newCheckedValue) {
                                        otherRow.newNameField.text = otherEntry.name;
                                    }
                                }
                            }
                        } else {
                            itemEntries[idx].checked = newCheckedValue;
                            newNameField.enabled = newCheckedValue;
                            if (!newCheckedValue) {
                                newNameField.text = itemEntries[idx].name;
                                itemEntries[idx].newName = itemEntries[idx].name;
                                itemEntries[idx].userEdited = false;
                            }
                        }
                        syncTargetFromCheckboxes();
                        if (requestPreviewUpdate) requestPreviewUpdate();
                        updateMoveButtonsState();
                    };

                    newNameField.onChange = function () {
                        itemEntries[idx].newName = newNameField.text;
                        itemEntries[idx].userEdited = true;
                    };

                    entryRows.push({
                        group: row,
                        checkbox: rowCheckbox,
                        currentNameLabel: currentNameLabel,
                        newNameField: newNameField,
                        dataIndex: idx
                    });
                })(entryIndex);
            }

            updateMoveButtonsState();
            entryRowsHost.layout.layout(true);
            dialog.layout.layout(true);
        }

        // 並び替えボタン
        var moveButtonRow = reorderPanel.add("group");
        moveButtonRow.orientation = "row";
        moveButtonRow.alignment = "center";
        moveButtonRow.spacing = 4;
        moveButtonRow.margins = [0, 10, 0, 0];
        var moveToTopBtn = moveButtonRow.add("button", undefined, L('moveTopBtn'));
        var moveUpBtn = moveButtonRow.add("button", undefined, L('moveUpBtn'));
        var moveDownBtn = moveButtonRow.add("button", undefined, L('moveDownBtn'));
        var moveToBottomBtn = moveButtonRow.add("button", undefined, L('moveBottomBtn'));
        var MOVE_BUTTON_WIDTH = 56;
        var MOVE_BUTTON_HEIGHT = 22;
        moveToTopBtn.preferredSize = [MOVE_BUTTON_WIDTH + 4, MOVE_BUTTON_HEIGHT];
        moveUpBtn.preferredSize = [MOVE_BUTTON_WIDTH, MOVE_BUTTON_HEIGHT];
        moveDownBtn.preferredSize = [MOVE_BUTTON_WIDTH, MOVE_BUTTON_HEIGHT];
        moveToBottomBtn.preferredSize = [MOVE_BUTTON_WIDTH + 4, MOVE_BUTTON_HEIGHT];
        moveToTopBtn.onClick = function () { moveCheckedToTop(); };
        moveUpBtn.onClick = function () { moveCheckedUp(); };
        moveDownBtn.onClick = function () { moveCheckedDown(); };
        moveToBottomBtn.onClick = function () { moveCheckedToBottom(); };

        // ボタンエリア / Button area（左：更新｜中央：spacer｜右：キャンセル＋OK）
        var buttonArea = dialog.add("group");
        buttonArea.orientation = "row";
        buttonArea.alignChildren = ["fill", "center"];
        buttonArea.alignment = "fill";

        var refreshBtn = buttonArea.add("button", undefined, L('refreshBtn'));
        refreshBtn.alignment = ["left", "center"];

        var spacer = buttonArea.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var cancelBtn = buttonArea.add("button", undefined, L('cancelButton'), { name: "cancel" });
        cancelBtn.alignment = ["right", "center"];

        var okBtn = buttonArea.add("button", undefined, L('okButton'), { name: "ok" });
        okBtn.alignment = ["right", "center"];

        okBtn.onClick = function () {
            syncEditingValues();
            for (var entryIdx = 0; entryIdx < itemEntries.length; entryIdx++) {
                if (itemEntries[entryIdx].checked && itemEntries[entryIdx].newName === "") {
                    alert(L('emptyNameAlert').replace("{n}", (entryIdx + 1)));
                    return;
                }
            }
            dialog.close(1);
        };

        // 全UI構築後に初回レンダリング
        refreshReorderRows();

        return {
            dialog: dialog,
            prefixInput: prefixInput,
            suffixInput: suffixInput,
            frontmostRadio: frontmostRadio,
            originalNameRadio: originalNameRadio,
            customRadio: customRadio,
            customInput: customInput,
            targetAllRadio: targetAllRadio,
            targetRangeRadio: targetRangeRadio,
            rangeInput: rangeInput,
            searchFilterCheckbox: searchFilterCheckbox,
            searchInput: searchInput,
            findInput: findInput,
            replaceInput: replaceInput,
            regexCheckbox: regexCheckbox,
            refreshBtn: refreshBtn,
            targetTypeArtboardRadio: targetTypeArtboardRadio,
            targetTypeSymbolRadio: targetTypeSymbolRadio,
            targetTypeLayerRadio: targetTypeLayerRadio,
            setTargetType: setTargetType,
            getCurrentTargetType: function () { return currentTargetType; },
            setTypeChangeCallback: function (callback) { typeChangeCallback = callback; },
            getItemEntries: function () { return itemEntries; },
            rebaselineEntriesAfterCommit: rebaselineEntriesAfterCommit,
            syncEditingValues: syncEditingValues,
            setRequestPreviewUpdate: function (callback) { requestPreviewUpdate = callback; },
            syncPreviewToReorderRows: function (previewNames) {
                var items = getDocumentItems(doc, currentTargetType);
                for (var rowIdx = 0; rowIdx < entryRows.length; rowIdx++) {
                    var row = entryRows[rowIdx];
                    var entry = itemEntries[row.dataIndex];
                    var originalIdx = entry.originalIndex;
                    var currentName = items[originalIdx].name;

                    // 元の名前 列：現在の canvas 名（[更新] 後は確定後の名前を表示）
                    row.currentNameLabel.text = currentName;
                    row.currentNameLabel.helpTip = currentName;

                    // 新しい名前 列：未確定プレビュー（手動編集行はスキップ）
                    if (entry.userEdited) continue;
                    var previewName = (previewNames && previewNames[originalIdx] != null)
                        ? previewNames[originalIdx]
                        : currentName;
                    row.newNameField.text = previewName;
                    entry.newName = previewName;
                }
            },
            applyTargetToCheckboxes: function (rangeMode, rangeText) {
                // 検索フィルタ：ON かつテキスト非空のときは、検索条件に見合うものだけにチェック（対象選択は無視）
                var filterActive = searchFilterCheckbox.value && searchInput.text !== "";
                var lowerQuery = filterActive ? searchInput.text.toLowerCase() : null;

                var isTargetIndex = {};
                if (!filterActive) {
                    var targetIndices = getTargetItemIndices(getDocumentItems(doc, currentTargetType).length, rangeMode, rangeText);
                    for (var targetIdx = 0; targetIdx < targetIndices.length; targetIdx++) {
                        isTargetIndex[targetIndices[targetIdx]] = true;
                    }
                }

                for (var entryIdx = 0; entryIdx < itemEntries.length; entryIdx++) {
                    var entry = itemEntries[entryIdx];
                    if (filterActive) {
                        entry.checked = entry.name.toLowerCase().indexOf(lowerQuery) !== -1;
                    } else {
                        entry.checked = !!isTargetIndex[entry.originalIndex];
                    }
                }

                for (var rowIdx = 0; rowIdx < entryRows.length; rowIdx++) {
                    var row = entryRows[rowIdx];
                    var rowEntry = itemEntries[row.dataIndex];
                    row.checkbox.value = rowEntry.checked;
                    row.newNameField.enabled = rowEntry.checked;
                    if (!rowEntry.checked) {
                        row.newNameField.text = rowEntry.name;
                        rowEntry.newName = rowEntry.name;
                        rowEntry.userEdited = false;
                    }
                }
                updateMoveButtonsState();
            }
        };
    }

    /* ダイアログを表示し、確定された設定を返す（キャンセル時は null） / Show dialog and return resulting settings (null on cancel) */
    function showRenameDialog(doc) {
        var dialogUI = createRenameDialog(doc);
        bindDialogEvents(dialogUI, doc);

        if (dialogUI.dialog.show() !== 1) {
            return null;
        }

        var settings = readDialogSettings(dialogUI);
        if (dialogUI.getItemEntries) {
            settings.itemEntries = dialogUI.getItemEntries();
        }
        return settings;
    }

    // =========================================
    // 並び替えと適用 / Reorder and apply
    // =========================================

    /* 並び替えまたはユーザー手動上書きが存在するかを判定 / Determine if reorder or user-edited rename is pending */
    function hasReorderOrRename(itemEntries) {
        for (var entryIdx = 0; entryIdx < itemEntries.length; entryIdx++) {
            if (itemEntries[entryIdx].originalIndex !== entryIdx) return true;
            if (itemEntries[entryIdx].userEdited) return true;
        }
        return false;
    }

    /* 並び替え結果と手動リネームをアイテムに適用 / Apply reorder result and manual renames to items */
    function applyReorderAndRename(doc, itemEntries, settings) {
        if (!itemEntries || !hasReorderOrRename(itemEntries)) return;

        var targetType = (settings && settings.targetType) || "artboard";
        var items = getDocumentItems(doc, targetType);
        var itemCount = items.length;

        // ユーザーが手動上書きした行を新しい位置で記録
        var userOverridesByNewPosition = {};
        for (var entryIdx = 0; entryIdx < itemEntries.length; entryIdx++) {
            if (itemEntries[entryIdx].userEdited && itemEntries[entryIdx].checked) {
                userOverridesByNewPosition[entryIdx] = itemEntries[entryIdx].newName;
            }
        }

        // 並べ替え前に現在の canvas 名を originalIndex で記録（[更新] 済みの名前を保持）
        var currentNamesByOriginalIndex = [];
        for (var origIdx = 0; origIdx < itemCount; origIdx++) {
            currentNamesByOriginalIndex.push(items[origIdx].name);
        }

        if (targetType === "artboard") {
            // rect と現在の canvas 名を新しい位置に並べ替え（一時名で衝突回避）
            for (var tempIdx = 0; tempIdx < itemCount; tempIdx++) {
                items[tempIdx].name = "__tmp_ab_" + tempIdx + "__";
            }
            for (var newPos = 0; newPos < itemEntries.length; newPos++) {
                items[newPos].artboardRect = itemEntries[newPos].rect;
                items[newPos].name = currentNamesByOriginalIndex[itemEntries[newPos].originalIndex];
            }
        } else if (targetType === "layer") {
            // レイヤー：安定参照を取得し、末尾から先頭へ PLACEATBEGINNING で並び替え
            var layerRefs = [];
            for (var layerInitIdx = 0; layerInitIdx < itemCount; layerInitIdx++) layerRefs.push(items[layerInitIdx]);
            for (var entryReverseIdx = itemEntries.length - 1; entryReverseIdx >= 0; entryReverseIdx--) {
                try {
                    layerRefs[itemEntries[entryReverseIdx].originalIndex].move(doc, ElementPlacement.PLACEATBEGINNING);
                } catch (moveError) {
                    $.writeln("[SmartRenamer] Layer move failed at entry " + entryReverseIdx + ": " + moveError);
                }
            }
            items = getDocumentItems(doc, "layer");
        }
        // symbol: 並び替え不可、リネームのみ

        /* 並び替え後の位置を基準に対象範囲を作り直して再リネーム / Rebuild the target range based on reordered positions before renaming */
        if (settings) {
            var reorderedTargetPositions = [];
            for (var targetPos = 0; targetPos < itemEntries.length; targetPos++) {
                if (itemEntries[targetPos].checked) {
                    reorderedTargetPositions.push(targetPos);
                }
            }

            var reorderedSettings = {};
            for (var settingsKey in settings) {
                if (settings.hasOwnProperty(settingsKey)) {
                    reorderedSettings[settingsKey] = settings[settingsKey];
                }
            }

            if (reorderedTargetPositions.length === itemCount) {
                reorderedSettings.rangeMode = "all";
                reorderedSettings.rangeText = "";
            } else {
                reorderedSettings.rangeMode = "numbered";
                reorderedSettings.rangeText = buildRangeString(reorderedTargetPositions);
            }

            executeRename(doc, reorderedSettings, { silent: true });
        }

        // 手動上書きを適用（items を再取得：レイヤー移動後の最新参照に追従）
        items = getDocumentItems(doc, targetType);
        for (var posKey in userOverridesByNewPosition) {
            if (!userOverridesByNewPosition.hasOwnProperty(posKey)) continue;
            var positionIndex = parseInt(posKey, 10);
            if (positionIndex >= 0 && positionIndex < items.length) {
                items[positionIndex].name = userOverridesByNewPosition[posKey];
            }
        }
    }

    // =========================================
    // メイン処理 / Main entry
    // =========================================

    /* スクリプトのエントリポイント / Script entry point */
    function main() {
        if (app.documents.length === 0) {
            return;
        }

        var doc = app.activeDocument;
        var originalState = captureOriginalState(doc);

        var dialogResult = showRenameDialog(doc);

        if (dialogResult) {
            // OK：途中の [更新] コミットを残したまま、最後の設定をもう一度適用
            if (dialogResult.itemEntries && hasReorderOrRename(dialogResult.itemEntries)) {
                applyReorderAndRename(doc, dialogResult.itemEntries, dialogResult);
            } else {
                executeRename(doc, dialogResult);
            }
        } else {
            // キャンセル：ダイアログを開く前の状態まで戻す（名前・アートボードrect・レイヤー順）
            restoreOriginalState(doc, originalState);
        }
    }

    // =========================================
    // リネーム実行とプレビュー / Rename execution and preview
    // =========================================

    /* 設定に従いアイテムをリネームして canvas を更新 / Execute rename on canvas based on settings */
    function executeRename(doc, settings, options) {
        var mode = settings.mode;
        var prefix = settings.prefix;
        var suffix = settings.suffix;
        var customText = settings.customText;
        var rangeMode = settings.rangeMode;
        var rangeText = settings.rangeText;
        var targetType = settings.targetType || "artboard";
        var findReplace = {
            findText: settings.findText || "",
            replaceText: settings.replaceText || "",
            useRegex: !!settings.useRegex
        };

        var silent = options && options.silent;

        if (mode === "custom" && customText === "" && prefix === "" && suffix === "" && !findReplace.findText) {
            if (!silent) {
                alert(L('alertNeedSettings'));
            }
            return false;
        }

        var items = getDocumentItems(doc, targetType);
        var itemTextMap = {};
        if (mode === "frontmost" && targetType === "artboard") {
            itemTextMap = mapTextFramesToArtboards(getFrontmostTextFramesPerArtboard(doc), items);
        } else if (mode === "original") {
            for (var originalModeIdx = 0; originalModeIdx < items.length; originalModeIdx++) {
                itemTextMap[originalModeIdx] = [items[originalModeIdx].name];
            }
        } else if (mode === "custom") {
            for (var customModeIdx = 0; customModeIdx < items.length; customModeIdx++) {
                itemTextMap[customModeIdx] = [customText];
            }
        }

        var targetIndices = getTargetItemIndices(items.length, rangeMode, rangeText);
        renameItems(items, itemTextMap, prefix, suffix, targetIndices, findReplace);
        return true;
    }

    /* canvas を変更せずに「いま [更新] したらこうなる」名前を計算する / Compute preview names without modifying the canvas */
    function computePreviewNames(doc, settings) {
        var targetType = settings.targetType || "artboard";
        var items = getDocumentItems(doc, targetType);
        var itemCount = items.length;

        // 既存の名前で初期化（対象外は現在の名前を維持）
        var previewNames = [];
        for (var initIdx = 0; initIdx < itemCount; initIdx++) {
            previewNames.push(items[initIdx].name);
        }

        var mode = settings.mode;
        var prefix = settings.prefix;
        var suffix = settings.suffix;
        var customText = settings.customText;
        var rangeMode = settings.rangeMode;
        var rangeText = settings.rangeText;
        var findText = settings.findText || "";
        var replaceText = settings.replaceText || "";
        var useRegex = !!settings.useRegex;

        if (mode === "custom" && customText === "" && prefix === "" && suffix === "" && findText === "") {
            return previewNames;
        }

        var itemTextMap = {};
        if (mode === "frontmost" && targetType === "artboard") {
            itemTextMap = mapTextFramesToArtboards(getFrontmostTextFramesPerArtboard(doc), items);
        } else if (mode === "original") {
            for (var originalModeIdx = 0; originalModeIdx < itemCount; originalModeIdx++) {
                itemTextMap[originalModeIdx] = [items[originalModeIdx].name];
            }
        } else if (mode === "custom") {
            for (var customModeIdx = 0; customModeIdx < itemCount; customModeIdx++) {
                itemTextMap[customModeIdx] = [customText];
            }
        }

        var targetIndices = getTargetItemIndices(itemCount, rangeMode, rangeText);
        var renamePlan = buildRenamePlan(items, itemTextMap, prefix, suffix, targetIndices, {
            findText: findText,
            replaceText: replaceText,
            useRegex: useRegex
        });
        for (var finalIdx = 0; finalIdx < renamePlan.indices.length; finalIdx++) {
            previewNames[renamePlan.indices[finalIdx]] = renamePlan.names[finalIdx];
        }

        return previewNames;
    }
    /* リネーム対象の最終名プランを作成（プレビュー／本番共通） / Build the final rename plan shared by preview and execution */
    function buildRenamePlan(items, itemTextMap, prefixTemplate, suffixTemplate, targetIndices, findReplace) {
        var findText = findReplace ? findReplace.findText : "";
        var replaceText = findReplace ? findReplace.replaceText : "";
        var useRegex = findReplace ? findReplace.useRegex : false;
        var skipUniquification = hasSequenceToken(prefixTemplate) || hasSequenceToken(suffixTemplate) || (findText && hasSequenceToken(replaceText));
        var reservedNames = getReservedItemNames(items, targetIndices);
        var targetIndexSet = makeIndexSet(targetIndices);
        var tokenContext = createTokenContext();
        var plannedBaseNames = [];
        var plannedIndices = [];
        var sequenceIndex = 1;

        for (var itemIdx = 0; itemIdx < items.length; itemIdx++) {
            if (!targetIndexSet[itemIdx]) continue;
            var expandedPrefix = expandTemplateTokens(prefixTemplate, sequenceIndex, tokenContext);
            var expandedSuffix = expandTemplateTokens(suffixTemplate, sequenceIndex, tokenContext);
            var textPart = itemTextMap[itemIdx] ? itemTextMap[itemIdx].join(" ") : "";
            var baseName = expandedPrefix + textPart + expandedSuffix;

            // 接頭辞・参照・接尾辞のいずれも未指定なら、検索・置換のみで現在名を加工
            if (!baseName && findText) {
                baseName = items[itemIdx].name;
            }

            if (findText) {
                var expandedReplace = expandTemplateTokens(replaceText, sequenceIndex, tokenContext);
                baseName = applyFindReplace(baseName, findText, expandedReplace, useRegex);
            }

            if (baseName) {
                plannedBaseNames.push(baseName);
                plannedIndices.push(itemIdx);
                sequenceIndex++;
            } else {
                // 対象だが結果が空のためスキップ → 現在の名前を予約して衝突を防ぐ
                reservedNames.push(items[itemIdx].name);
            }
        }

        return {
            indices: plannedIndices,
            names: resolveUniqueNames(plannedBaseNames, reservedNames, skipUniquification)
        };
    }

    // =========================================
    // ユーティリティ / Utilities
    // =========================================

    /* 対象タイプ（artboard / symbol / layer）に応じたドキュメントコレクションを返す / Return the document collection for the given target type */
    function getDocumentItems(doc, targetType) {
        if (targetType === "symbol") return doc.symbols;
        if (targetType === "layer") return doc.layers;
        return doc.artboards;
    }

    /* 対象外アイテムの名前を予約として返す（衝突回避用） / Return non-target item names as reserved set */
    function getReservedItemNames(items, targetIndices) {
        var targetIndexSet = makeIndexSet(targetIndices);
        var reserved = [];
        for (var i = 0; i < items.length; i++) {
            if (!targetIndexSet[i]) {
                reserved.push(items[i].name);
            }
        }
        return reserved;
    }

    /* テキストフレームの可視バウンズの中心座標を取得 / Get the center point of a text frame's visible bounds */
    function getTextCenter(textFrame) {
        var bounds = textFrame.visibleBounds;
        return [(bounds[0] + bounds[2]) / 2, (bounds[1] + bounds[3]) / 2];
    }

    /* "1-3,5" 形式の範囲文字列を 0-based のインデックス配列にパース / Parse "1-3,5" range string into 0-based indices */
    function parseRangeString(rangeText) {
        var result = [];
        var parts = rangeText.split(",");
        for (var i = 0; i < parts.length; i++) {
            var part = parts[i].replace(/\s+/g, "");
            if (/^\d+$/.test(part)) {
                result.push(parseInt(part, 10) - 1);
            } else if (/^\d+-\d+$/.test(part)) {
                var range = part.split("-");
                var rangeStart = parseInt(range[0], 10);
                var rangeEnd = parseInt(range[1], 10);
                for (var j = rangeStart; j <= rangeEnd; j++) result.push(j - 1);
            }
        }
        return result;
    }

    /* "all" / "numbered" モードに応じて対象アイテムのインデックス配列を返す / Get target item indices based on range mode */
    function getTargetItemIndices(itemCount, rangeMode, rangeText) {
        if (rangeMode === "all") {
            var allIndices = [];
            for (var i = 0; i < itemCount; i++) allIndices.push(i);
            return allIndices;
        }
        return parseRangeString(rangeText);
    }

    /* 0-based のインデックス配列から "1-3,5" 形式の 1-based 範囲文字列を生成 / Build "1-3,5" string from 0-based indices */
    function buildRangeString(zeroBasedIndices) {
        if (!zeroBasedIndices || zeroBasedIndices.length === 0) return "";
        var sorted = [];
        for (var i = 0; i < zeroBasedIndices.length; i++) sorted.push(zeroBasedIndices[i]);
        sorted.sort(function (a, b) { return a - b; });

        var parts = [];
        var rangeStart = sorted[0];
        var rangeEnd = sorted[0];
        for (var j = 1; j < sorted.length; j++) {
            if (sorted[j] === rangeEnd + 1) {
                rangeEnd = sorted[j];
            } else {
                parts.push(rangeStart === rangeEnd
                    ? (rangeStart + 1) + ""
                    : (rangeStart + 1) + "-" + (rangeEnd + 1));
                rangeStart = sorted[j];
                rangeEnd = sorted[j];
            }
        }
        parts.push(rangeStart === rangeEnd
            ? (rangeStart + 1) + ""
            : (rangeStart + 1) + "-" + (rangeEnd + 1));
        return parts.join(",");
    }

    /* テキストフレームを所属アートボードに紐付けてマッピング / Map each text frame to its containing artboard */
    function mapTextFramesToArtboards(textFrames, artboards) {
        var map = {};
        for (var i = 0; i < textFrames.length; i++) {
            var textFrame = textFrames[i];
            var center = getTextCenter(textFrame);
            for (var j = 0; j < artboards.length; j++) {
                var artboardBounds = artboards[j].artboardRect;
                if (center[0] >= artboardBounds[0] && center[0] <= artboardBounds[2] &&
                    center[1] <= artboardBounds[1] && center[1] >= artboardBounds[3]) {
                    if (!map[j]) map[j] = [];
                    var cleanedText = textFrame.contents.replace(/[\r\n\t]/g, "");
                    map[j].push(cleanedText);
                    break;
                }
            }
        }
        return map;
    }

    /* 計画名のリストから、重複しているものだけ "_1", "_2" を付けて最終名を返す（重複しない名前はそのまま） / Resolve unique names by appending "_1", "_2" only when collisions exist (in plan or reserved) */
    /* ハッシュ集合を使い O(N) で衝突判定する */
    function resolveUniqueNames(plannedBaseNames, reservedNames, skipUniquification) {
        var resolvedNames = [];
        if (skipUniquification) {
            for (var i = 0; i < plannedBaseNames.length; i++) {
                resolvedNames.push(plannedBaseNames[i]);
            }
            return resolvedNames;
        }

        // plan 内での baseName 出現回数（>1 なら重複）
        var baseNameCountInPlan = {};
        for (var countIdx = 0; countIdx < plannedBaseNames.length; countIdx++) {
            var baseNameKey = plannedBaseNames[countIdx];
            baseNameCountInPlan[baseNameKey] = (baseNameCountInPlan[baseNameKey] || 0) + 1;
        }

        var reservedSet = {};
        var usedNameSet = {};
        for (var reservedIdx = 0; reservedIdx < reservedNames.length; reservedIdx++) {
            reservedSet[reservedNames[reservedIdx]] = true;
            usedNameSet[reservedNames[reservedIdx]] = true;
        }

        var collisionSeqByBase = {};
        for (var planIdx = 0; planIdx < plannedBaseNames.length; planIdx++) {
            var baseName = plannedBaseNames[planIdx];
            var needsSuffix = baseNameCountInPlan[baseName] > 1 || reservedSet[baseName] === true;
            var finalName;
            if (!needsSuffix) {
                finalName = baseName;
            } else {
                collisionSeqByBase[baseName] = (collisionSeqByBase[baseName] || 0) + 1;
                finalName = baseName + "_" + collisionSeqByBase[baseName];
                while (usedNameSet[finalName] === true) {
                    collisionSeqByBase[baseName]++;
                    finalName = baseName + "_" + collisionSeqByBase[baseName];
                }
            }
            usedNameSet[finalName] = true;
            resolvedNames.push(finalName);
        }
        return resolvedNames;
    }

    /* インデックス配列をハッシュ集合へ変換 / Convert an index array into a hash set for O(1) membership checks */
    function makeIndexSet(indices) {
        var set = {};
        for (var i = 0; i < indices.length; i++) set[indices[i]] = true;
        return set;
    }

    /* expandTemplateTokens 用のキャッシュ可能な実行コンテキスト（fileName / dateString は呼び出しごとに同一） / Build a reusable token-expansion context */
    function createTokenContext() {
        var now = new Date();
        return {
            fileName: app.activeDocument.name.replace(/\.[^.]+$/, ""),
            dateString: now.getFullYear().toString() +
                ("0" + (now.getMonth() + 1)).slice(-2) +
                ("0" + now.getDate()).slice(-2)
        };
    }

    /* テンプレート文字列の連番／#FN／#DT トークンを展開 / Expand sequence/#FN/#DT tokens in a template */
    function expandTemplateTokens(template, index, context) {
        var ctx = context || createTokenContext();
        var fileName = ctx.fileName;
        var dateString = ctx.dateString;

        var result = template;

        // 連番トークン {#N} を展開（ゼロパディング対応：{#01} → 01, 02, ...）
        result = result.replace(/\{#(\d+)\}/g, function (match, token) {
            var base = parseInt(token, 10);
            var value = base + index - 1;
            if (token.charAt(0) === "0" && token.length > 1) {
                return ("0000000000" + value).slice(-token.length);
            }
            return value.toString();
        });

        result = result.replace(/#FN/g, fileName);
        result = result.replace(/#DT/g, dateString);

        return result;
    }

    /* テンプレートに連番トークン {#N} が含まれるかを判定 / Check whether the template contains a sequence token {#N} */
    function hasSequenceToken(template) {
        return /\{#\d+\}/.test(template);
    }

    /* 対象アイテムを実際にリネーム実行（リネーム本体） / Perform the actual rename pass over target items */
    function renameItems(items, itemTextMap, prefixTemplate, suffixTemplate, targetIndices, findReplace) {
        var renamePlan = buildRenamePlan(items, itemTextMap, prefixTemplate, suffixTemplate, targetIndices, findReplace);
        for (var finalIdx = 0; finalIdx < renamePlan.indices.length; finalIdx++) {
            try {
                items[renamePlan.indices[finalIdx]].name = renamePlan.names[finalIdx];
            } catch (renameError) {
                $.writeln("[SmartRenamer] Failed to rename item index " + renamePlan.indices[finalIdx] + " to '" + renamePlan.names[finalIdx] + "': " + renameError);
            }
        }
    }

    /* 各アートボードの最前面 TextFrame をレイヤー／グループ階層を再帰して取得 / Get the frontmost text frame per artboard by recursively scanning layers and groups */
    /* 判定順はレイヤー順・pageItems順に依存（Illustrator の厳密な描画Z順ではない） / Detection order depends on layer order and pageItems order (not Illustrator's exact drawing Z-order) */
    function getFrontmostTextFramesPerArtboard(doc) {
        var result = [];
        var artboardCount = doc.artboards.length;

        /* アイテムが指定アートボード内にあるかを中心座標で判定 / Check whether an item is inside the target artboard by center point */
        function isTextFrameOnArtboard(textFrame, artboardBounds) {
            var center = getTextCenter(textFrame);
            return center[0] >= artboardBounds[0] && center[0] <= artboardBounds[2] &&
                center[1] <= artboardBounds[1] && center[1] >= artboardBounds[3];
        }

        /* コンテナ内を再帰的に走査して最初に見つかった TextFrame を返す / Recursively scan a container and return the first matching text frame */
        function findFrontmostTextFrameInContainer(container, artboardBounds) {
            var items = container.pageItems;
            for (var itemIndex = 0; itemIndex < items.length; itemIndex++) {
                var item = items[itemIndex];
                if (item.hidden || item.locked) continue;

                if (item.typename === "TextFrame") {
                    if (isTextFrameOnArtboard(item, artboardBounds)) {
                        return item;
                    }
                } else if (item.typename === "GroupItem") {
                    var nestedTextFrame = findFrontmostTextFrameInContainer(item, artboardBounds);
                    if (nestedTextFrame) return nestedTextFrame;
                }
            }
            return null;
        }

        /* レイヤーとサブレイヤーを再帰的に走査 / Recursively scan layers and sublayers */
        function findFrontmostTextFrameInLayer(layer, artboardBounds) {
            if (!layer.visible || layer.locked) return null;

            var textFrame = findFrontmostTextFrameInContainer(layer, artboardBounds);
            if (textFrame) return textFrame;

            for (var layerIndex = 0; layerIndex < layer.layers.length; layerIndex++) {
                var nestedTextFrame = findFrontmostTextFrameInLayer(layer.layers[layerIndex], artboardBounds);
                if (nestedTextFrame) return nestedTextFrame;
            }
            return null;
        }

        for (var artboardIndex = 0; artboardIndex < artboardCount; artboardIndex++) {
            var artboardBounds = doc.artboards[artboardIndex].artboardRect;
            var frontmostFrame = null;
            var layers = doc.layers;
            for (var layerIndex = 0; layerIndex < layers.length; layerIndex++) {
                frontmostFrame = findFrontmostTextFrameInLayer(layers[layerIndex], artboardBounds);
                if (frontmostFrame) break;
            }
            if (frontmostFrame) result.push(frontmostFrame);
        }

        return result;
    }

    main();

})();