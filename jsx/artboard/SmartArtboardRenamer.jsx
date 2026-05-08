#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

(function () {

    /*
    ### スクリプト名：

    SmartArtboardRenamer.jsx

    ### 概要

    - Illustratorのアートボード名を、接頭辞・接尾辞・参照テキストを組み合わせて一括リネームするスクリプトです。
    - ダイアログ上で設定を変更すると、アートボード名を自動でプレビュー更新します。

    #### リネーム条件（左カラム）

    - 接頭辞・接尾辞に連番（1, 01）、ファイル名（#FN）、日付（#DT）、区切り（- _）を挿入できるトークンボタン
    - 「最前面のテキスト」モードで、レイヤー／グループ階層を再帰して各アートボードの最前面テキストを参照
    - 「レイヤー」モードで、指定レイヤー配下（サブレイヤー・グループ含む）のテキストを結合して参照
    - 「元のアートボード名」モードでダイアログを開いた時点の名前を参照（[更新] 後の名前ではない）
    - 「指定」モードで入力した任意のテキストを参照（空にすれば参照テキストなし）
    - 同名が重複する場合は "_2", "_3" などを自動付加（接頭辞・接尾辞に連番トークンがある場合はスキップ）
    - 非表示レイヤーは無視

    #### 対象アートボード（右カラム上段）

    - 「すべて」または「指定範囲」から選択可能（指定範囲は "1-3,5" 形式）
    - 並び替え／リネームパネルのチェックボックスと双方向に連動（チェック操作で「指定範囲」テキストが自動更新、全選択時は「すべて」へ自動切替）

    #### 並び替え／リネーム（右カラム下段）

    - 各アートボードの「元の名前 → 新しい名前」を一覧表示し、リネーム条件のプレビューがそのまま反映
    - チェックを付けた行のみ「新しい名前」を手動で上書き可能（手動編集後はプレビューに上書きされない）
    - 行のチェックを付けた状態で「↑先頭へ／↑上へ／↓下へ／↓末尾へ」で並び替え（並び替え後の位置基準でリネーム対象も再構築）
    - チェックボックス上で Option+クリックすると全行 ON/OFF を一括切替

    #### ボタン

    - 「更新」で現在の設定をカンバスのアートボード名に即時反映（プレビュー結果を確定）
    - キャンセル時はダイアログ表示前のアートボード名に戻す

    ### 更新履歴

    - v1.0 (20250509) : 初期バージョン
    - v1.5.1 (20260508) : 接頭辞／接尾辞のクリア x ボタンを各行から最終行 1 個のみに整理

    ---

    ### Script Name:

    SmartArtboardRenamer.jsx

    ### Overview

    - A script to batch rename Illustrator artboards by combining prefix, suffix, and reference text.
    - As settings change in the dialog, the artboard names are updated automatically for preview.

    #### Rename conditions (left column)

    - Token buttons can insert sequential tokens (1, 01), file name (#FN), date (#DT), and separators (- _) into the prefix/suffix
    - "Frontmost Text" mode recursively scans layers/groups and uses the frontmost text frame on each artboard
    - "Layer" mode combines text from the selected layer hierarchy (including sublayers and groups)
    - "Original Artboard Name" mode uses the name as it was when the dialog opened (not the post-[Refresh] name)
    - "Custom" mode uses the entered text as the reference text (leave empty for no reference text)
    - Automatically appends "_2", "_3", etc. on name collision (skipped when prefix/suffix contains a sequence token)
    - Ignores hidden layers

    #### Target artboards (right column, top)

    - Choose "All" or "Range" (e.g. "1-3,5")
    - Two-way linked with the Reorder / Rename checkboxes (toggling checkboxes auto-updates the Range text and switches to "All" when every row is on)

    #### Reorder / Rename (right column, bottom)

    - Lists each artboard as "Original Name -> New Name" reflecting the live preview
    - Checked rows allow manual override of the new name (manual edits stay sticky)
    - Move checked rows with Top / Up / Down / Bottom buttons (rename targets are rebuilt from reordered positions)
    - Option-click a checkbox to toggle all rows on/off

    #### Buttons

    - "Refresh" commits the current settings to the canvas immediately (applies preview names)
    - Cancelling restores the original artboard names from before the dialog opened

    ### Update History

    - v1.0 (20250509): Initial version
    - v1.5.1 (20260508): Consolidated the prefix/suffix clear "x" button to a single one on the last row
    */

    // =========================================
    // バージョンとローカライズ
    // =========================================

    var SCRIPT_VERSION = "v1.5.1";

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
            ja: "アートボードのリネームとソート",
            en: "Smart Artboard Renamer"
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
            ja: "参照テキスト",
            en: "Text Source"
        },
        basicLabel: {
            ja: "リネーム",
            en: "Rename"
        },
        layerOption: {
            ja: "レイヤー",
            en: "Layer"
        },
        frontmostOption: {
            ja: "最前面のテキスト",
            en: "Frontmost Text"
        },
        originalNameOption: {
            ja: "元のアートボード名",
            en: "Original Artboard Name"
        },
        customOption: {
            ja: "指定",
            en: "Custom"
        },
        targetLabel: {
            ja: "対象アートボード",
            en: "Target Artboards"
        },
        allBoards: {
            ja: "すべて",
            en: "All"
        },
        specificBoards: {
            ja: "指定範囲",
            en: "Range"
        },
        cancelButton: {
            ja: "キャンセル",
            en: "Cancel"
        },
        okButton: {
            ja: "OK",
            en: "OK"
        },
        alertNoLayer: {
            ja: "指定されたレイヤーが見つからないか、非表示です。",
            en: "The specified layer was not found or is hidden."
        },
        alertNeedSettings: {
            ja: "接頭辞・接尾辞のいずれかを入力するか、参照テキストを指定してください。",
            en: "Enter a prefix or suffix, or choose a text source."
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

    /* キャンセル時にアートボード名を元に戻す / Restore artboard names on cancel */
    function restoreOriginalArtboardNames(doc, names) {
        for (var i = 0; i < doc.artboards.length; i++) {
            doc.artboards[i].name = names[i];
        }
    }

    /* ダイアログ各コントロールから設定オブジェクトを構築 / Read settings object from dialog controls */
    function readDialogSettings(dialogUI) {
        var mode = dialogUI.frontmostRadio.value ? "frontmost"
            : (dialogUI.layerRadio.value ? "layer"
                : (dialogUI.originalNameRadio.value ? "originalName"
                    : "custom"));
        return {
            mode: mode,
            prefix: dialogUI.prefixInput.text,
            suffix: dialogUI.suffixInput.text,
            customText: dialogUI.customInput.text,
            artboardTarget: dialogUI.allArtboardsRadio.value ? "all" : "numbered",
            numberInput: dialogUI.rangeInput.text,
            selectedLayerName: dialogUI.layerRadio.value && dialogUI.layerDropdown.selection ? dialogUI.layerDropdown.selection.text : null
        };
    }

    /* ダイアログ各コントロールにイベントハンドラを設定 / Wire dialog event handlers */
    function bindDialogEvents(dialogUI, doc, originalNames) {
        function buildSettings() {
            var settings = readDialogSettings(dialogUI);
            settings.originalNames = originalNames;
            return settings;
        }

        function updatePreview() {
            // canvas は触らず、未確定プレビュー名を計算して右カラムを更新
            var previewNames = computePreviewNames(doc, buildSettings());
            if (dialogUI.syncPreviewToReorderRows) {
                dialogUI.syncPreviewToReorderRows(previewNames);
            }
        }

        function commitCurrentSettings() {
            // 現在の設定を canvas に確定（[更新] / OK）
            if (executeRename(doc, buildSettings(), { silent: true })) {
                app.redraw();
            }
            updatePreview();
        }

        function syncTargetInput() {
            dialogUI.rangeInput.enabled = dialogUI.rangeArtboardsRadio.value;
            if (dialogUI.applyTargetToCheckboxes) {
                if (dialogUI.allArtboardsRadio.value) {
                    dialogUI.applyTargetToCheckboxes("all", "");
                } else {
                    dialogUI.applyTargetToCheckboxes("numbered", dialogUI.rangeInput.text);
                }
            }
            updatePreview();
        }
        dialogUI.allArtboardsRadio.onClick = function () {
            dialogUI.rangeArtboardsRadio.value = false;
            syncTargetInput();
        };
        dialogUI.rangeArtboardsRadio.onClick = function () {
            dialogUI.allArtboardsRadio.value = false;
            syncTargetInput();
        };

        var sourceRadios = [dialogUI.originalNameRadio, dialogUI.customRadio, dialogUI.frontmostRadio, dialogUI.layerRadio];

        function selectSourceRadio(selected) {
            for (var i = 0; i < sourceRadios.length; i++) {
                sourceRadios[i].value = (sourceRadios[i] === selected);
            }
            dialogUI.layerDropdown.enabled = dialogUI.layerRadio.value;
            dialogUI.customInput.enabled = dialogUI.customRadio.value;
            if (dialogUI.customRadio.value) {
                try { dialogUI.customInput.active = true; } catch (focusError) { }
            }
            updatePreview();
        }
        dialogUI.originalNameRadio.onClick = function () { selectSourceRadio(dialogUI.originalNameRadio); };
        dialogUI.customRadio.onClick = function () { selectSourceRadio(dialogUI.customRadio); };
        dialogUI.frontmostRadio.onClick = function () { selectSourceRadio(dialogUI.frontmostRadio); };
        dialogUI.layerRadio.onClick = function () { selectSourceRadio(dialogUI.layerRadio); };
        dialogUI.layerDropdown.onChange = function () {
            if (dialogUI.layerRadio.value) updatePreview();
        };
        dialogUI.customInput.onChange = function () {
            if (dialogUI.customRadio.value) updatePreview();
        };

        dialogUI.prefixInput.onChange = updatePreview;
        dialogUI.suffixInput.onChange = updatePreview;
        dialogUI.rangeInput.onChange = function () {
            if (dialogUI.rangeArtboardsRadio.value) {
                if (dialogUI.applyTargetToCheckboxes) {
                    dialogUI.applyTargetToCheckboxes("numbered", dialogUI.rangeInput.text);
                }
                updatePreview();
            }
        };

        if (dialogUI.refreshBtn) {
            dialogUI.refreshBtn.onClick = function () {
                /* edittext のフォーカス未解除でも値を確実に反映させる / Force onChange to commit pending edittext value */
                try { dialogUI.rangeInput.notify("onChange"); } catch (notifyError) { }
                commitCurrentSettings();
            };
        }

        // チェックボックス操作 → 対象アートボード設定 への逆同期に updatePreview を注入
        if (dialogUI.setRequestPreviewUpdate) {
            dialogUI.setRequestPreviewUpdate(updatePreview);
        }

        // 初期同期：対象アートボード設定 → チェックボックス
        if (dialogUI.applyTargetToCheckboxes) {
            if (dialogUI.allArtboardsRadio.value) {
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
        var tokenRows = [
            [
                { label: "1", value: "1", width: 22 },
                { label: "01", value: "01" },
                { label: "-", value: "-", width: 22 },
                { label: "_", value: "_", width: 22 }
            ],
            [
                { label: "#FN", value: "#FN", width: 40 },
                { label: "#DT", value: "#DT", width: 40 }
            ]
        ];
        for (var rowIdx = 0; rowIdx < tokenRows.length; rowIdx++) {
            var tokenButtonRow = panel.add("group");
            tokenButtonRow.orientation = "row";
            tokenButtonRow.spacing = 4;
            tokenButtonRow.margins = 0;
            var tokens = tokenRows[rowIdx];
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
            /* 最終行末にのみ、対象フィールドをクリアする x ボタンを配置 / Add the clear button only on the last row */
            if (rowIdx === tokenRows.length - 1) {
                var clearFieldBtn = tokenButtonRow.add("button", undefined, "x");
                clearFieldBtn.preferredSize = [22, 20];
                clearFieldBtn.onClick = function () {
                    targetInput.text = "";
                    targetInput.notify("onChange");
                };
            }
        }
    }

    // =========================================
    // ダイアログ構築と表示 / Dialog build and show
    // =========================================

    /* リネームダイアログのUIを構築 / Build the rename dialog UI */
    function createRenameDialog(doc) {
        var dialog = new Window("dialog", L('dialogTitle') + ' ' + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = "fill";

        // コンテンツ行（左：リネーム条件、右：対象アートボード＋並び替え／リネーム）
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

        // 参照テキスト
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

        // 「レイヤー」ラジオとレイヤードロップダウンを横並びに
        var layerRow = sourcePanel.add("group");
        layerRow.orientation = "row";
        layerRow.alignChildren = ["left", "center"];
        var layerRadio = layerRow.add("radiobutton", undefined, L('layerOption'));
        var layerDropdown = layerRow.add("dropdownlist", undefined, []);
        layerDropdown.minimumSize.width = 100;
        var allLayers = doc.layers;
        for (var i = 0; i < allLayers.length; i++) {
            if (!allLayers[i].visible) continue;
            layerDropdown.add("item", allLayers[i].name);
        }
        if (layerDropdown.items.length > 0) layerDropdown.selection = 0;
        layerDropdown.enabled = false;

        // 全ラジオを追加した後に初期値を設定
        originalNameRadio.value = false;
        customRadio.value = false;
        frontmostRadio.value = true;
        layerRadio.value = false;

        // 接尾辞
        var suffixPanel = basicPanel.add("panel", undefined, L('suffixLabel'));
        suffixPanel.orientation = "column";
        suffixPanel.alignChildren = "left";
        setupPanel(suffixPanel);
        var suffixInput = suffixPanel.add("edittext", undefined, "");
        suffixInput.characters = 16;
        addTokenButtons(suffixPanel, suffixInput);

        // 右カラム：対象アートボード＋並び替え／リネーム
        var rightColumn = contentRow.add("group");
        rightColumn.orientation = "column";
        rightColumn.alignChildren = "fill";

        // 対象アートボード（右カラム上段）
        var targetPanel = rightColumn.add("panel", undefined, L('targetLabel'));
        targetPanel.orientation = "row";
        targetPanel.alignChildren = ["left", "center"];
        setupPanel(targetPanel);

        var allArtboardsRadio = targetPanel.add("radiobutton", undefined, L('allBoards'));
        var rangeArtboardsRadio = targetPanel.add("radiobutton", undefined, L('specificBoards'));
        var rangeInput = targetPanel.add("edittext", undefined, "");
        rangeInput.characters = 10;
        rangeInput.enabled = false;
        allArtboardsRadio.value = true;
        rangeArtboardsRadio.value = false;

        var reorderPanel = rightColumn.add("panel", undefined, L('reorderLabel'));
        reorderPanel.orientation = "column";
        reorderPanel.alignChildren = "fill";
        setupPanel(reorderPanel, 6);

        var artboardEntries = [];
        for (var ai = 0; ai < doc.artboards.length; ai++) {
            artboardEntries.push({
                originalIndex: ai,
                name: doc.artboards[ai].name,
                newName: doc.artboards[ai].name,
                rect: doc.artboards[ai].artboardRect,
                checked: false,
                userEdited: false
            });
        }

        // bindDialogEvents から updatePreview を注入してもらうコールバック
        var requestPreviewUpdate = null;

        /* チェック状態 → 対象アートボード設定（ラジオ＋指定範囲テキスト）への逆同期 / Sync target settings from checkbox state */
        function syncTargetFromCheckboxes() {
            var checkedPositions = [];
            for (var i = 0; i < artboardEntries.length; i++) {
                if (artboardEntries[i].checked) checkedPositions.push(i);
            }
            if (checkedPositions.length === artboardEntries.length) {
                allArtboardsRadio.value = true;
                rangeArtboardsRadio.value = false;
                rangeInput.enabled = false;
            } else {
                allArtboardsRadio.value = false;
                rangeArtboardsRadio.value = true;
                rangeInput.enabled = true;
                /* 並び替え後の表示位置を基準に指定範囲へ反映 / Reflect reordered display positions in the range field */
                if (checkedPositions.length > 0) {
                    rangeInput.text = buildArtboardRangeString(checkedPositions);
                }
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
                artboardEntries[row.dataIndex].checked = row.checkbox.value;
                if (row.checkbox.value) {
                    artboardEntries[row.dataIndex].newName = row.newNameField.text;
                }
            }
        }

        function swapEntries(indexA, indexB) {
            var temp = artboardEntries[indexA];
            artboardEntries[indexA] = artboardEntries[indexB];
            artboardEntries[indexB] = temp;
        }

        function moveCheckedUp() {
            syncEditingValues();
            for (var i = 1; i < artboardEntries.length; i++) {
                if (artboardEntries[i].checked && !artboardEntries[i - 1].checked) swapEntries(i, i - 1);
            }
            refreshReorderRows();
        }

        function moveCheckedDown() {
            syncEditingValues();
            for (var i = artboardEntries.length - 2; i >= 0; i--) {
                if (artboardEntries[i].checked && !artboardEntries[i + 1].checked) swapEntries(i, i + 1);
            }
            refreshReorderRows();
        }

        function moveCheckedToTop() {
            syncEditingValues();
            var checkedEntries = [], uncheckedEntries = [];
            for (var i = 0; i < artboardEntries.length; i++) {
                if (artboardEntries[i].checked) checkedEntries.push(artboardEntries[i]);
                else uncheckedEntries.push(artboardEntries[i]);
            }
            artboardEntries = checkedEntries.concat(uncheckedEntries);
            refreshReorderRows();
        }

        function moveCheckedToBottom() {
            syncEditingValues();
            var checkedEntries = [], uncheckedEntries = [];
            for (var i = 0; i < artboardEntries.length; i++) {
                if (artboardEntries[i].checked) checkedEntries.push(artboardEntries[i]);
                else uncheckedEntries.push(artboardEntries[i]);
            }
            artboardEntries = uncheckedEntries.concat(checkedEntries);
            refreshReorderRows();
        }

        function canMoveUp() {
            for (var i = 1; i < artboardEntries.length; i++) {
                if (artboardEntries[i].checked && !artboardEntries[i - 1].checked) return true;
            }
            return false;
        }
        function canMoveDown() {
            for (var i = 0; i < artboardEntries.length - 1; i++) {
                if (artboardEntries[i].checked && !artboardEntries[i + 1].checked) return true;
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

            for (var entryIndex = 0; entryIndex < artboardEntries.length; entryIndex++) {
                (function (idx) {
                    var row = entryRowsHost.add("group");
                    row.orientation = "row";
                    row.alignChildren = ["left", "center"];
                    row.spacing = 6;

                    var orderLabel = row.add("statictext", undefined, (idx + 1) + "");
                    orderLabel.preferredSize.width = 24;

                    var rowCheckbox = row.add("checkbox", undefined, "");
                    rowCheckbox.value = artboardEntries[idx].checked;
                    rowCheckbox.preferredSize.width = 28;

                    var currentNameLabel = row.add("statictext", undefined, artboardEntries[idx].name);
                    currentNameLabel.preferredSize.width = 140;
                    currentNameLabel.helpTip = artboardEntries[idx].name;

                    row.add("statictext", undefined, "→").preferredSize.width = 14;

                    var newNameField = row.add("edittext", undefined, artboardEntries[idx].newName);
                    newNameField.preferredSize.width = 160;
                    newNameField.enabled = artboardEntries[idx].checked;

                    rowCheckbox.onClick = function () {
                        var newCheckedValue = rowCheckbox.value;
                        var optionKeyHeld = ScriptUI.environment && ScriptUI.environment.keyboardState && ScriptUI.environment.keyboardState.altKey;
                        if (optionKeyHeld) {
                            for (var entryIdx = 0; entryIdx < artboardEntries.length; entryIdx++) {
                                artboardEntries[entryIdx].checked = newCheckedValue;
                                if (!newCheckedValue) {
                                    artboardEntries[entryIdx].newName = artboardEntries[entryIdx].name;
                                    artboardEntries[entryIdx].userEdited = false;
                                }
                            }
                            for (var rowIdx = 0; rowIdx < entryRows.length; rowIdx++) {
                                var otherRow = entryRows[rowIdx];
                                var otherEntry = artboardEntries[otherRow.dataIndex];
                                otherRow.checkbox.value = newCheckedValue;
                                otherRow.newNameField.enabled = newCheckedValue;
                                if (!newCheckedValue) {
                                    otherRow.newNameField.text = otherEntry.name;
                                }
                            }
                        } else {
                            artboardEntries[idx].checked = newCheckedValue;
                            newNameField.enabled = newCheckedValue;
                            if (!newCheckedValue) {
                                newNameField.text = artboardEntries[idx].name;
                                artboardEntries[idx].newName = artboardEntries[idx].name;
                                artboardEntries[idx].userEdited = false;
                            }
                        }
                        syncTargetFromCheckboxes();
                        if (requestPreviewUpdate) requestPreviewUpdate();
                        updateMoveButtonsState();
                    };

                    newNameField.onChange = function () {
                        artboardEntries[idx].newName = newNameField.text;
                        artboardEntries[idx].userEdited = true;
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
            for (var r = 0; r < artboardEntries.length; r++) {
                if (artboardEntries[r].checked && artboardEntries[r].newName === "") {
                    alert(L('emptyNameAlert').replace("{n}", (r + 1)));
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
            layerRadio: layerRadio,
            originalNameRadio: originalNameRadio,
            customRadio: customRadio,
            customInput: customInput,
            layerDropdown: layerDropdown,
            allArtboardsRadio: allArtboardsRadio,
            rangeArtboardsRadio: rangeArtboardsRadio,
            rangeInput: rangeInput,
            refreshBtn: refreshBtn,
            getArtboardEntries: function () { return artboardEntries; },
            syncEditingValues: syncEditingValues,
            setRequestPreviewUpdate: function (cb) { requestPreviewUpdate = cb; },
            syncPreviewToReorderRows: function (previewNames) {
                for (var rowIdx = 0; rowIdx < entryRows.length; rowIdx++) {
                    var row = entryRows[rowIdx];
                    var entry = artboardEntries[row.dataIndex];
                    var originalIdx = entry.originalIndex;
                    var currentName = doc.artboards[originalIdx].name;

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
            applyTargetToCheckboxes: function (targetType, numberInput) {
                var targetIndices = getTargetArtboardIndices(doc.artboards.length, targetType, numberInput);
                var isTargetIndex = {};
                for (var targetIdx = 0; targetIdx < targetIndices.length; targetIdx++) {
                    isTargetIndex[targetIndices[targetIdx]] = true;
                }

                for (var entryIdx = 0; entryIdx < artboardEntries.length; entryIdx++) {
                    artboardEntries[entryIdx].checked = !!isTargetIndex[artboardEntries[entryIdx].originalIndex];
                }

                for (var rowIdx = 0; rowIdx < entryRows.length; rowIdx++) {
                    var row = entryRows[rowIdx];
                    var entry = artboardEntries[row.dataIndex];
                    row.checkbox.value = entry.checked;
                    row.newNameField.enabled = entry.checked;
                    if (!entry.checked) {
                        row.newNameField.text = entry.name;
                        entry.newName = entry.name;
                        entry.userEdited = false;
                    }
                }
                updateMoveButtonsState();
            }
        };
    }

    /* ダイアログを表示し、確定された設定を返す（キャンセル時は null） / Show dialog and return resulting settings (null on cancel) */
    function showRenameDialog(doc, originalNames) {
        var dialogUI = createRenameDialog(doc);
        bindDialogEvents(dialogUI, doc, originalNames);

        if (dialogUI.dialog.show() !== 1) {
            return null;
        }

        var settings = readDialogSettings(dialogUI);
        settings.originalNames = originalNames;
        if (dialogUI.getArtboardEntries) {
            settings.artboardEntries = dialogUI.getArtboardEntries();
        }
        return settings;
    }

    // =========================================
    // 並び替えと適用 / Reorder and apply
    // =========================================

    /* 並び替えまたはユーザー手動上書きが存在するかを判定 / Determine if reorder or user-edited rename is pending */
    function hasReorderOrRename(artboardEntries) {
        for (var r = 0; r < artboardEntries.length; r++) {
            if (artboardEntries[r].originalIndex !== r) return true;
            if (artboardEntries[r].userEdited) return true;
        }
        return false;
    }

    /* 並び替え結果と手動リネームをアートボードに適用 / Apply reorder result and manual renames to artboards */
    function applyReorderAndRename(doc, artboardEntries, settings) {
        if (!artboardEntries || !hasReorderOrRename(artboardEntries)) return;

        var artboards = doc.artboards;
        var artboardCount = artboards.length;

        // ユーザーが手動上書きした行を新しい位置で記録
        var userOverridesByNewPosition = {};
        for (var entryIdx = 0; entryIdx < artboardEntries.length; entryIdx++) {
            if (artboardEntries[entryIdx].userEdited && artboardEntries[entryIdx].checked) {
                userOverridesByNewPosition[entryIdx] = artboardEntries[entryIdx].newName;
            }
        }

        // 並べ替え前に現在の canvas 名を originalIndex で記録（[更新] 済みの名前を保持）
        var currentNamesByOriginalIndex = [];
        for (var origIdx = 0; origIdx < artboardCount; origIdx++) {
            currentNamesByOriginalIndex.push(artboards[origIdx].name);
        }

        // rect と現在の canvas 名を新しい位置に並べ替え（一時名で衝突回避）
        for (var tempIdx = 0; tempIdx < artboardCount; tempIdx++) {
            artboards[tempIdx].name = "__tmp_ab_" + tempIdx + "__";
        }
        for (var newPos = 0; newPos < artboardEntries.length; newPos++) {
            artboards[newPos].artboardRect = artboardEntries[newPos].rect;
            artboards[newPos].name = currentNamesByOriginalIndex[artboardEntries[newPos].originalIndex];
        }

        /* 並び替え後の位置を基準に対象範囲を作り直して再リネーム / Rebuild the target range based on reordered positions before renaming */
        if (settings) {
            var reorderedTargetPositions = [];
            for (var targetPos = 0; targetPos < artboardEntries.length; targetPos++) {
                if (artboardEntries[targetPos].checked) {
                    reorderedTargetPositions.push(targetPos);
                }
            }

            var reorderedSettings = {};
            for (var settingsKey in settings) {
                if (settings.hasOwnProperty(settingsKey)) {
                    reorderedSettings[settingsKey] = settings[settingsKey];
                }
            }

            if (reorderedTargetPositions.length === artboardCount) {
                reorderedSettings.artboardTarget = "all";
                reorderedSettings.numberInput = "";
            } else {
                reorderedSettings.artboardTarget = "numbered";
                reorderedSettings.numberInput = buildArtboardRangeString(reorderedTargetPositions);
            }

            /* originalName モード用に、元の名前参照順も並び替え後へ補正 / Rebuild original-name references based on reordered positions */
            if (reorderedSettings.originalNames && reorderedSettings.originalNames.length) {
                var reorderedOriginalNames = [];
                for (var reorderedIndex = 0; reorderedIndex < artboardEntries.length; reorderedIndex++) {
                    reorderedOriginalNames.push(
                        reorderedSettings.originalNames[artboardEntries[reorderedIndex].originalIndex]
                    );
                }
                reorderedSettings.originalNames = reorderedOriginalNames;
            }

            executeRename(doc, reorderedSettings, { silent: true });
        }

        // 手動上書きを適用
        for (var posKey in userOverridesByNewPosition) {
            if (!userOverridesByNewPosition.hasOwnProperty(posKey)) continue;
            var positionIndex = parseInt(posKey, 10);
            if (positionIndex >= 0 && positionIndex < artboardCount) {
                artboards[positionIndex].name = userOverridesByNewPosition[posKey];
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
        var originalNames = [];
        for (var i = 0; i < doc.artboards.length; i++) {
            originalNames.push(doc.artboards[i].name);
        }

        var dialogResult = showRenameDialog(doc, originalNames);

        if (dialogResult) {
            // OK：途中の [更新] コミットを残したまま、最後の設定をもう一度適用
            executeRename(doc, dialogResult);
            applyReorderAndRename(doc, dialogResult.artboardEntries, dialogResult);
        } else {
            // キャンセル：ダイアログを開く前の名前まで戻す
            restoreOriginalArtboardNames(doc, originalNames);
        }
    }

    // =========================================
    // リネーム実行とプレビュー / Rename execution and preview
    // =========================================

    /* 設定に従いアートボードをリネームして canvas を更新 / Execute rename on canvas based on settings */
    function executeRename(doc, settings, options) {
        var mode = settings.mode;
        var prefix = settings.prefix;
        var suffix = settings.suffix;
        var customText = settings.customText;
        var targetType = settings.artboardTarget;
        var numberInput = settings.numberInput;
        var layerName = settings.selectedLayerName;

        var silent = options && options.silent;

        if (mode === "custom" && customText === "" && prefix === "" && suffix === "") {
            if (!silent) {
                alert(L('alertNeedSettings'));
            }
            return false;
        }

        var artboardTextMap = {};
        if (mode === "layer") {
            var sourceLayer = findLayerByName(doc, layerName);
            if (!sourceLayer || !sourceLayer.visible) {
                if (!silent) {
                    alert(L('alertNoLayer'));
                }
                return false;
            }
            artboardTextMap = mapTextFramesToArtboards(getTextFramesInLayer(sourceLayer), doc.artboards);
        } else if (mode === "frontmost") {
            artboardTextMap = mapTextFramesToArtboards(getFrontmostTextFramesPerArtboard(doc), doc.artboards);
        } else if (mode === "originalName") {
            // ダイアログ表示前の名前を参照（更新後の名前ではなく、開いた時点の元の名前）
            var originalNamesSrc = settings.originalNames;
            for (var origModeIdx = 0; origModeIdx < doc.artboards.length; origModeIdx++) {
                var srcName = (originalNamesSrc && originalNamesSrc[origModeIdx] != null)
                    ? originalNamesSrc[origModeIdx]
                    : doc.artboards[origModeIdx].name;
                artboardTextMap[origModeIdx] = [srcName];
            }
        } else if (mode === "custom") {
            for (var customModeIdx = 0; customModeIdx < doc.artboards.length; customModeIdx++) {
                artboardTextMap[customModeIdx] = [customText];
            }
        }

        var targetIndices = getTargetArtboardIndices(doc.artboards.length, targetType, numberInput);
        renameArtboards(doc.artboards, artboardTextMap, prefix, suffix, targetIndices);
        return true;
    }

    /* canvas を変更せずに「いま [更新] したらこうなる」名前を計算する / Compute preview names without modifying the canvas */
    function computePreviewNames(doc, settings) {
        var artboards = doc.artboards;
        var artboardCount = artboards.length;

        // 既存の名前で初期化（対象外は現在の名前を維持）
        var previewNames = [];
        for (var initIdx = 0; initIdx < artboardCount; initIdx++) {
            previewNames.push(artboards[initIdx].name);
        }

        var mode = settings.mode;
        var prefix = settings.prefix;
        var suffix = settings.suffix;
        var customText = settings.customText;
        var targetType = settings.artboardTarget;
        var numberInput = settings.numberInput;
        var layerName = settings.selectedLayerName;

        if (mode === "custom" && customText === "" && prefix === "" && suffix === "") {
            return previewNames;
        }

        var artboardTextMap = {};
        if (mode === "layer") {
            var sourceLayer = findLayerByName(doc, layerName);
            if (!sourceLayer || !sourceLayer.visible) return previewNames;
            artboardTextMap = mapTextFramesToArtboards(getTextFramesInLayer(sourceLayer), artboards);
        } else if (mode === "frontmost") {
            artboardTextMap = mapTextFramesToArtboards(getFrontmostTextFramesPerArtboard(doc), artboards);
        } else if (mode === "originalName") {
            // ダイアログ表示前の名前を参照（[更新] 後の現在名ではなく、開いた時点の元の名前）
            var originalNamesSrc = settings.originalNames;
            for (var origModeIdx = 0; origModeIdx < artboardCount; origModeIdx++) {
                var srcName = (originalNamesSrc && originalNamesSrc[origModeIdx] != null)
                    ? originalNamesSrc[origModeIdx]
                    : artboards[origModeIdx].name;
                artboardTextMap[origModeIdx] = [srcName];
            }
        } else if (mode === "custom") {
            for (var customModeIdx = 0; customModeIdx < artboardCount; customModeIdx++) {
                artboardTextMap[customModeIdx] = [customText];
            }
        }

        var targetIndices = getTargetArtboardIndices(artboardCount, targetType, numberInput);

        // 対象外アートボードは現在の名前を維持するため、それらを衝突回避リストに含める
        var usedNames = getReservedArtboardNames(artboards, targetIndices);

        // 接頭辞・接尾辞に連番トークンがあれば結果がユニークになる前提で _2/_3 自動付加をスキップ
        var skipUniquification = hasSequenceToken(prefix) || hasSequenceToken(suffix);
        var sequenceIndex = 1;
        for (var artboardIdx = 0; artboardIdx < artboardCount; artboardIdx++) {
            if (!contains(targetIndices, artboardIdx)) continue;
            var expandedPrefix = expandTemplateTokens(prefix, sequenceIndex);
            var expandedSuffix = expandTemplateTokens(suffix, sequenceIndex);
            var textPart = artboardTextMap[artboardIdx] ? artboardTextMap[artboardIdx].join(" ") : "";
            if (expandedPrefix || expandedSuffix || textPart) {
                var baseName = expandedPrefix + textPart + expandedSuffix;
                var uniqueName = skipUniquification ? baseName : generateUniqueName(baseName, usedNames);
                usedNames.push(uniqueName);
                previewNames[artboardIdx] = uniqueName;
                sequenceIndex++;
            } else {
                // 対象だが結果が空のためスキップ → 現在の名前を予約して衝突を防ぐ
                usedNames.push(previewNames[artboardIdx]);
            }
        }

        return previewNames;
    }

    // =========================================
    // ユーティリティ / Utilities
    // =========================================

    /* 対象外アートボードの名前を予約として返す（衝突回避用） / Return non-target artboard names as reserved set */
    function getReservedArtboardNames(artboards, targetIndices) {
        var reserved = [];
        for (var i = 0; i < artboards.length; i++) {
            if (!contains(targetIndices, i)) {
                reserved.push(artboards[i].name);
            }
        }
        return reserved;
    }

    /* 名前でトップレベルレイヤーを検索 / Find a top-level layer by name */
    function findLayerByName(doc, name) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === name) return doc.layers[i];
        }
        return null;
    }

    function getTextFramesInLayer(layer) {
        var results = [];

        /* コンテナ内の TextFrame をグループ階層まで再帰的に収集 / Recursively collect text frames inside containers and groups */
        function collectTextFramesInContainer(container) {
            var items = container.pageItems;
            for (var itemIndex = 0; itemIndex < items.length; itemIndex++) {
                var item = items[itemIndex];
                if (item.hidden || item.locked) continue;

                if (item.typename === "TextFrame") {
                    results.push(item);
                } else if (item.typename === "GroupItem") {
                    collectTextFramesInContainer(item);
                }
            }
        }

        /* レイヤーとサブレイヤー内の TextFrame を再帰的に収集 / Recursively collect text frames inside layers and sublayers */
        function collectTextFramesInLayer(targetLayer) {
            if (!targetLayer.visible || targetLayer.locked) return;

            collectTextFramesInContainer(targetLayer);

            for (var layerIndex = 0; layerIndex < targetLayer.layers.length; layerIndex++) {
                collectTextFramesInLayer(targetLayer.layers[layerIndex]);
            }
        }

        collectTextFramesInLayer(layer);
        return results;
    }

    /* テキストフレームの可視バウンズの中心座標を取得 / Get the center point of a text frame's visible bounds */
    function getTextCenter(textFrame) {
        var bounds = textFrame.visibleBounds;
        return [(bounds[0] + bounds[2]) / 2, (bounds[1] + bounds[3]) / 2];
    }

    /* "1-3,5" 形式の範囲文字列を 0-based のインデックス配列にパース / Parse "1-3,5" range string into 0-based indices */
    function parseArtboardRange(input) {
        var result = [];
        var parts = input.split(",");
        for (var i = 0; i < parts.length; i++) {
            var part = parts[i].replace(/\s+/g, "");
            if (/^\d+$/.test(part)) {
                result.push(parseInt(part, 10) - 1);
            } else if (/^\d+-\d+$/.test(part)) {
                var range = part.split("-");
                var start = parseInt(range[0], 10);
                var end = parseInt(range[1], 10);
                for (var j = start; j <= end; j++) result.push(j - 1);
            }
        }
        return result;
    }

    /* "all" / "numbered" モードに応じて対象インデックス配列を返す / Get target artboard indices based on mode */
    function getTargetArtboardIndices(total, mode, input) {
        if (mode === "all") {
            var list = [];
            for (var i = 0; i < total; i++) list.push(i);
            return list;
        }
        return parseArtboardRange(input);
    }

    /* 0-based のインデックス配列から "1-3,5" 形式の 1-based 範囲文字列を生成 / Build "1-3,5" string from 0-based indices */
    function buildArtboardRangeString(zeroBasedIndices) {
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

    /* 配列に値が含まれるかを判定するヘルパー / Simple Array.contains helper */
    function contains(array, value) {
        for (var i = 0; i < array.length; i++) {
            if (array[i] === value) return true;
        }
        return false;
    }

    /* 重複を避けるため "_2", "_3" 付きの一意な名前を生成 / Generate a unique name by appending "_2", "_3" on conflict */
    function generateUniqueName(baseName, usedNames) {
        var name = baseName;
        var count = 2;
        while (contains(usedNames, name)) {
            name = baseName + "_" + count;
            count++;
        }
        return name;
    }

    /* テンプレート文字列の連番／#FN／#DT トークンを展開 / Expand sequence/#FN/#DT tokens in a template */
    function expandTemplateTokens(template, index) {
        var fileName = app.activeDocument.name.replace(/\.[^.]+$/, "");
        var now = new Date();
        var dateString = now.getFullYear().toString() +
            ("0" + (now.getMonth() + 1)).slice(-2) +
            ("0" + now.getDate()).slice(-2);

        var result = template;

        // 連番置換を先に行う（#DT/#FN を先に展開すると、日付やファイル名内の数字を誤って捕捉してしまうため）
        var match = result.match(/\d+/);
        if (match) {
            var token = match[0];
            var base = parseInt(token, 10);
            var value = base + index - 1;
            var replacement = (token.charAt(0) === "0" && token.length > 1)
                ? ("0000000000" + value).slice(-token.length)
                : value.toString();
            result = result.replace(token, replacement);
        }

        result = result.replace(/#FN/g, fileName);
        result = result.replace(/#DT/g, dateString);

        return result;
    }

    /* テンプレートに連番トークン（数字）が含まれるかを判定 / Check whether the template contains a numeric sequence token */
    function hasSequenceToken(template) {
        return /\d+/.test(template);
    }

    /* 対象アートボードを実際にリネーム実行（リネーム本体） / Perform the actual rename pass over target artboards */
    function renameArtboards(artboards, artboardTextMap, prefixTemplate, suffixTemplate, targetIndices) {
        // 接頭辞・接尾辞に連番トークンがあれば結果がユニークになる前提で _2/_3 自動付加をスキップ
        var skipUniquification = hasSequenceToken(prefixTemplate) || hasSequenceToken(suffixTemplate);
        var usedNames = getReservedArtboardNames(artboards, targetIndices);
        var sequenceIndex = 1;
        for (var artboardIdx = 0; artboardIdx < artboards.length; artboardIdx++) {
            if (!contains(targetIndices, artboardIdx)) continue;
            var expandedPrefix = expandTemplateTokens(prefixTemplate, sequenceIndex);
            var expandedSuffix = expandTemplateTokens(suffixTemplate, sequenceIndex);
            var textPart = artboardTextMap[artboardIdx] ? artboardTextMap[artboardIdx].join(" ") : "";
            if (expandedPrefix || expandedSuffix || textPart) {
                var baseName = expandedPrefix + textPart + expandedSuffix;
                var uniqueName = skipUniquification ? baseName : generateUniqueName(baseName, usedNames);
                usedNames.push(uniqueName);
                try {
                    artboards[artboardIdx].name = uniqueName;
                } catch (renameError) {
                    $.writeln("[SmartArtboardRenamer] Failed to rename artboard index " + artboardIdx + " to '" + uniqueName + "': " + renameError);
                }
                sequenceIndex++;
            } else {
                // リネーム対象だが結果が空のためスキップ → 現在の名前を予約して衝突を防ぐ
                usedNames.push(artboards[artboardIdx].name);
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