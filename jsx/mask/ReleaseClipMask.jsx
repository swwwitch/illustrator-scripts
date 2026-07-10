#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

- 選択したクリッピングマスクをモード別に解除するIllustrator用スクリプト。
- 単純解除、マスクパス削除、マスク内容削除の3つの方法に対応。

### 主な機能

- 単純に解除（マスクパスと内容を両方残す）
- マスク内容を残して、マスクパスを削除
- マスクパスを残して、マスク内容を削除
- マスクパスにK100の塗り（不透明度15%）を適用するオプション付き
- 複合パス（CompoundPathItem）のマスクに対応
- 日本語／英語UI対応

### 対象範囲

- 直接選択したクリッピンググループのみが対象。
- 通常グループの内側にネストしたクリッピンググループは対象外。

### 処理の流れ

1. ダイアログでモードとオプションを選択
2. 選択されたモードに従ってマスクを解除
3. 必要に応じてマスクパスに塗りを適用

### 紹介記事（note）

https://note.com/dtp_tranist/n/nebc832e574f7
https://note.com/dtp_tranist/n/ne64abe5a2e8c

### 更新履歴

- v1.0 (20250606) : 初期バージョン
- v1.2.1 (20260711) : IIFE化・命名整理・ローカライズ構造化、重なり順保持・マスク内容削除・複合パス対応・フォールバック追加

*/

/*

### Overview

- Illustrator script to release selected clipping masks by mode.
- Supports simple release, remove mask path, or remove masked content.

### Main Features

- Simple release (keep both mask path and content)
- Keep masked content, remove the mask path
- Keep the mask path, remove masked content
- Optional K100 fill (15% opacity) on the mask path
- Supports compound-path (CompoundPathItem) masks
- Japanese/English UI support

### Scope

- Only directly selected clipping groups are processed.
- Clipping groups nested inside normal groups are not processed.

### Process Flow

1. Choose mode and option in dialog
2. Release clipping mask accordingly
3. Optionally apply fill color to the mask path

### Update History

- v1.0 (20250606) : Initial release
- v1.2.1 (20260711) : IIFE, naming, localization; stacking order, masked-content removal, compound-path support, fallbacks

*/

// =========================================
// バージョン / Version
// =========================================

var SCRIPT_VERSION = "v1.2.1";

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /**
     * 起動時のデフォルト値。/ Default values on launch.
     *
     * @type {{releaseMode: string, applyFill: boolean}}
     */
    var USER_DEFAULTS = {
        releaseMode: "removePath", /* simple | removePath | removeContent */
        applyFill: true
    };

    /** パスへ適用する塗りの不透明度（%）/ Fill opacity applied to path (%) @type {number} */
    var PATH_FILL_OPACITY = 15;

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 現在のロケールに基づく言語コードを返します。/ Return current language code.
     *
     * @returns {string} 日本語環境なら "ja"、それ以外は "en"。
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    /** 現在の言語コード / Current language code @type {string} */
    var currentLanguage = getCurrentLang();

    /** 日英ラベル定義（カテゴリ分け）/ Japanese-English labels (categorized) */
    var LABELS = {
        dialog: {
            title: { ja: "クリッピングマスクの解除", en: "Release Clipping Mask" },
            releaseMode: { ja: "解除方法", en: "Release Mode" }
        },
        radio: {
            simpleRelease: { ja: "単純に解除", en: "Simply release" },
            removeMaskPath: { ja: "マスク内容を残して、パスを削除", en: "Keep content, remove path" },
            removeContent: { ja: "パスを残して、マスク内容を削除", en: "Keep path, remove content" }
        },
        checkbox: {
            applyFill: {
                ja: "解除時、マスクパスにカラーを設定",
                en: "Set fill color to mask path when releasing"
            }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        tooltip: {
            simpleRelease: {
                ja: "マスクパスと内容の両方を残し、クリッピングマスクだけを解除します。",
                en: "Release the clipping mask while keeping both the mask path and the content."
            },
            removeMaskPath: {
                ja: "マスク用のパスを削除し、マスクされていた内容だけを残します。",
                en: "Delete the mask path and keep only the masked content."
            },
            removeContent: {
                ja: "マスクパス以外（配置画像・埋め込み画像・グループ・テキストなど）をすべて削除し、マスクパスだけを残します。",
                en: "Delete everything except the mask path (images, groups, text, ...) and keep only the mask path."
            },
            applyFill: {
                ja: "残ったマスクパスにK100・不透明度15%の塗りを設定します。",
                en: "Apply a K100 fill at 15% opacity to the remaining mask path."
            }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "オブジェクトが選択されていません。", en: "No objects selected." },
            noClippingMask: { ja: "選択範囲にクリッピングマスクがありません。", en: "No clipping mask in the selection." },
            releaseFailed: {
                ja: "{count}個のクリッピングマスクの解除に失敗しました。",
                en: "Failed to release {count} clipping mask(s)."
            }
        }
    };

    /**
     * ドットパスで指定したラベルを現在の言語で返します。/ Return the localized label for a dot-path.
     *
     * キーが見つからない場合は keyPath をそのまま返し、現在言語のラベルが無い場合は
     * 日本語→英語の順にフォールバックします。
     *
     * @param {string} keyPath - LABELS 内のキーをドットで連結したパス（例 "dialog.title"）。
     * @returns {string} ローカライズ済み文字列。未定義時は keyPath を返す。
     */
    function getLocalizedText(keyPath) {
        var keyParts = keyPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < keyParts.length; i++) {
            if (labelNode === undefined || labelNode === null) {
                return keyPath;
            }
            labelNode = labelNode[keyParts[i]];
        }
        if (labelNode === undefined || labelNode === null) {
            return keyPath;
        }
        var localizedText = labelNode[currentLanguage];
        if (localizedText === undefined || localizedText === null) localizedText = labelNode.ja;
        if (localizedText === undefined || localizedText === null) localizedText = labelNode.en;
        return (localizedText === undefined || localizedText === null) ? keyPath : localizedText;
    }

    // =========================================
    // UIレイアウトの共通設定 / Shared UI layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 8;                  /* パネル内の要素間隔 / panel spacing */

    /**
     * ウィンドウ（ダイアログ）へ共通のレイアウト設定を適用します。/ Apply shared window layout.
     *
     * @param {Window} targetWindow - 設定対象のウィンドウ。
     * @param {number} [spacing] - 要素間隔。省略時は WINDOW_SPACING。
     * @returns {void}
     */
    function setupWindow(targetWindow, spacing) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = "fill";
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルへ共通のレイアウト設定を適用します。/ Apply shared panel layout.
     *
     * @param {Panel} panel - 設定対象のパネル。
     * @param {number} [spacing] - 要素間隔。省略時は PANEL_SPACING。
     * @returns {void}
     */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    // =========================================
    // 選択・マスク判定ヘルパー / Selection & mask helpers
    // =========================================

    /**
     * ドキュメントの選択を安全に配列で取得します。/ Safely get the selection as an array.
     *
     * テキスト選択（TextRange）など配列でない場合は空配列を返します。
     *
     * @param {Document} doc - 対象ドキュメント。
     * @returns {Array<PageItem>} 選択アイテムの配列。取得できない場合は空配列。
     */
    function getSelectionItems(doc) {
        var selection = doc.selection;
        if (!selection || typeof selection.length !== "number" || selection.typename) {
            return [];
        }
        /* ライブ選択を実配列へコピー（以降のDOM変更の影響を受けない）
           / Copy the live selection into a real array so later DOM changes do not affect it */
        var selectionItems = [];
        for (var i = 0; i < selection.length; i++) {
            selectionItems.push(selection[i]);
        }
        return selectionItems;
    }

    /**
     * クリッピングマスクを持つグループかどうかを判定します。/ Check whether the item is a clipping-mask group.
     *
     * @param {PageItem} pageItem - 判定対象のページアイテム。
     * @returns {boolean} クリップグループなら true。
     */
    function isClippingGroup(pageItem) {
        return pageItem.typename === "GroupItem" && pageItem.clipped === true;
    }

    /**
     * 選択からクリップグループだけを抽出して配列で返します。/ Collect only clipping groups from the selection.
     *
     * @param {Array<PageItem>} selectionItems - 選択アイテムの配列。
     * @returns {Array<GroupItem>} クリップグループの配列。
     */
    function collectClippingGroups(selectionItems) {
        var clippingGroups = [];
        for (var i = 0; i < selectionItems.length; i++) {
            if (isClippingGroup(selectionItems[i])) {
                clippingGroups.push(selectionItems[i]);
            }
        }
        return clippingGroups;
    }

    /**
     * 指定アイテムがクリッピングマスクとして機能しているか判定します。/ Check whether the item acts as a clipping mask.
     *
     * PathItem は clipping、CompoundPathItem は内部 pathItems の clipping を確認します。
     *
     * @param {PageItem} pageItem - 判定対象のページアイテム。
     * @returns {boolean} マスクとして機能していれば true。
     */
    function isMaskItem(pageItem) {
        if (pageItem.typename === "PathItem") {
            return pageItem.clipping === true;
        }
        if (pageItem.typename === "CompoundPathItem") {
            var subPaths = pageItem.pathItems;
            for (var i = 0; i < subPaths.length; i++) {
                if (subPaths[i].clipping === true) {
                    return true;
                }
            }
        }
        return false;
    }

    /**
     * クリップグループ内で最初にマスクとして機能しているアイテムを返します。/ Return the first mask item in the group.
     *
     * @param {GroupItem} clippingGroup - 対象のクリップグループ。
     * @returns {PageItem|null} マスクアイテム。見つからなければ null。
     */
    function findMaskItem(clippingGroup) {
        for (var i = 0; i < clippingGroup.pageItems.length; i++) {
            var pageItem = clippingGroup.pageItems[i];
            if (isMaskItem(pageItem)) {
                return pageItem;
            }
        }
        return null;
    }

    /**
     * @typedef {Object} MaskRemovalPlan
     * @property {PageItem|null} maskItem - 最初に見つかったマスクアイテム。無ければ null。
     * @property {Array<PageItem>} itemsToRemove - マスク以外の子（削除対象）。
     */

    /**
     * マスクと削除対象を1回の走査で確定します。/ Determine the mask and removal targets in a single pass.
     *
     * 最初に見つかったマスクを残し、それ以外を削除対象とすることで参照比較（===）を避けます。
     *
     * @param {GroupItem} clippingGroup - 対象のクリップグループ。
     * @returns {MaskRemovalPlan} マスク参照と削除対象アイテムの一覧。
     */
    function collectContentToRemove(clippingGroup) {
        var maskItem = null;
        var itemsToRemove = [];
        for (var i = 0; i < clippingGroup.pageItems.length; i++) {
            var pageItem = clippingGroup.pageItems[i];
            if (!maskItem && isMaskItem(pageItem)) {
                maskItem = pageItem;
            } else {
                itemsToRemove.push(pageItem);
            }
        }
        return { maskItem: maskItem, itemsToRemove: itemsToRemove };
    }

    /**
     * クリップグループ内にマスク以外の削除対象（マスク内容）が存在するか判定します。/ Check for maskable content.
     *
     * @param {Array<GroupItem>} clippingGroups - 対象のクリップグループ配列。
     * @returns {boolean} 1つでもマスク以外の子があれば true。
     */
    function hasMaskedContent(clippingGroups) {
        for (var i = 0; i < clippingGroups.length; i++) {
            var plan = collectContentToRemove(clippingGroups[i]);
            if (plan.maskItem && plan.itemsToRemove.length > 0) {
                return true;
            }
        }
        return false;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 処理の起点。ドキュメント・選択・クリップグループを確認し、ダイアログ確定時に解除を実行します。
     * / Entry logic: validate, show dialog, and run release on confirm.
     *
     * @returns {void}
     */
    function main() {
        if (!app.documents.length) {
            alert(getLocalizedText("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;
        var selectionItems = getSelectionItems(doc);
        if (!selectionItems.length) {
            alert(getLocalizedText("alert.noSelection"));
            return;
        }

        var clippingGroups = collectClippingGroups(selectionItems);
        if (!clippingGroups.length) {
            alert(getLocalizedText("alert.noClippingMask"));
            return;
        }

        var settings = showReleaseDialog(clippingGroups);
        if (!settings) return; /* キャンセル / Cancelled */

        executeRelease(settings.releaseMode, settings.applyFill, clippingGroups);
    }

    /**
     * @typedef {Object} ReleaseSettings
     * @property {string} releaseMode - 解除モード（"simple" | "removePath" | "removeContent"）。
     * @property {boolean} applyFill - 解除時に残ったマスクパスへ塗りを適用する場合は true。
     */

    /**
     * 解除方法・オプションを選ぶダイアログを構築し、選択結果を返します。/ Build the options dialog and return the choice.
     *
     * @param {Array<GroupItem>} clippingGroups - 対象のクリップグループ配列（有効/無効判定に使用）。
     * @returns {ReleaseSettings|null} OK 時は設定オブジェクト、キャンセル・クローズ時は null。
     */
    function showReleaseDialog(clippingGroups) {
        var dialog = new Window("dialog", getLocalizedText("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(dialog);

        /* 解除方法パネル / Release mode panel */
        var releaseModePanel = dialog.add("panel", undefined, getLocalizedText("dialog.releaseMode"));
        setupPanel(releaseModePanel, 6);

        var radioSimpleRelease = releaseModePanel.add("radiobutton", undefined, getLocalizedText("radio.simpleRelease"));
        var radioRemoveMaskPath = releaseModePanel.add("radiobutton", undefined, getLocalizedText("radio.removeMaskPath"));
        var radioRemoveContent = releaseModePanel.add("radiobutton", undefined, getLocalizedText("radio.removeContent"));

        radioSimpleRelease.helpTip = getLocalizedText("tooltip.simpleRelease");
        radioRemoveMaskPath.helpTip = getLocalizedText("tooltip.removeMaskPath");
        radioRemoveContent.helpTip = getLocalizedText("tooltip.removeContent");

        /* デフォルトの解除方法を選択 / Select default release mode */
        if (USER_DEFAULTS.releaseMode === "simple") {
            radioSimpleRelease.value = true;
        } else if (USER_DEFAULTS.releaseMode === "removeContent") {
            radioRemoveContent.value = true;
        } else {
            radioRemoveMaskPath.value = true;
        }

        /* マスク以外の内容が無ければ「マスク内容を削除」を無効化 / Disable content removal when there is nothing to remove */
        if (!hasMaskedContent(clippingGroups)) {
            radioRemoveContent.enabled = false;
            if (radioRemoveContent.value) {
                radioRemoveMaskPath.value = true;
            }
        }

        /* 解除時に塗りカラーを適用するチェックボックス / Checkbox for fill color on release */
        var fillOptionGroup = dialog.add("group");
        fillOptionGroup.orientation = "column";
        fillOptionGroup.alignChildren = "left";
        fillOptionGroup.margins = [20, 5, 0, 10];
        var applyFillCheckbox = fillOptionGroup.add("checkbox", undefined, getLocalizedText("checkbox.applyFill"));
        applyFillCheckbox.helpTip = getLocalizedText("tooltip.applyFill");
        applyFillCheckbox.value = USER_DEFAULTS.applyFill;

        /* パス削除モードでは塗りチェックボックスを無効化 / Disable fill checkbox in remove-path mode */
        function updateFillCheckboxState() {
            applyFillCheckbox.enabled = !radioRemoveMaskPath.value;
        }
        radioSimpleRelease.onClick = updateFillCheckboxState;
        radioRemoveContent.onClick = updateFillCheckboxState;
        radioRemoveMaskPath.onClick = updateFillCheckboxState;
        updateFillCheckboxState();

        /* ボタン行 / Button row */
        var buttonGroup = dialog.add("group");
        buttonGroup.orientation = "row";
        buttonGroup.alignment = "center";

        buttonGroup.add("button", undefined, getLocalizedText("button.cancel"), { name: "cancel" });
        buttonGroup.add("button", undefined, "OK", { name: "ok" });

        /* OK以外（キャンセル・クローズ）は null / Return null unless OK was pressed */
        if (dialog.show() !== 1) return null;

        var selectedMode = "simple";
        if (radioRemoveMaskPath.value) {
            selectedMode = "removePath";
        } else if (radioRemoveContent.value) {
            selectedMode = "removeContent";
        }

        /* 無効時は値が true でも false を返す / Return false when the checkbox is disabled */
        return { releaseMode: selectedMode, applyFill: applyFillCheckbox.enabled && applyFillCheckbox.value };
    }

    // =========================================
    // 解除処理 / Release logic
    // =========================================

    /**
     * 選択モードに従って各クリップグループのマスクを解除します。/ Release masks per mode.
     *
     * @param {string} releaseMode - 解除モード（"simple" | "removePath" | "removeContent"）。
     * @param {boolean} shouldApplyFill - 残ったマスクパスへ塗りを適用する場合は true。
     * @param {Array<GroupItem>} clippingGroups - 対象のクリップグループ配列。
     * @returns {void}
     */
    function executeRelease(releaseMode, shouldApplyFill, clippingGroups) {
        /* グループ単位で try/catch し、途中失敗しても残りを処理して件数を通知
           / Isolate each group so one failure does not abort the rest */
        var failureReasons = [];
        for (var i = 0; i < clippingGroups.length; i++) {
            var clippingGroup = clippingGroups[i];
            try {
                if (releaseMode === "removePath") {
                    releaseKeepContent(clippingGroup);
                } else if (releaseMode === "removeContent") {
                    releaseKeepMaskPath(clippingGroup, shouldApplyFill);
                } else {
                    releaseKeepBoth(clippingGroup, shouldApplyFill);
                }
            } catch (err) {
                /* 失敗理由を保持して後でまとめて通知（ExtendScript は description のことがある）
                   / Keep the reason and report later (ExtendScript may use description) */
                failureReasons.push(err.message || err.description || String(err));
            }
        }
        if (failureReasons.length > 0) {
            var failureMessage = getLocalizedText("alert.releaseFailed").replace("{count}", failureReasons.length);
            alert(failureMessage + "\n" + failureReasons.join("\n"));
        }
    }

    /**
     * マスクパスと内容を両方残し、マスクだけを解除します。/ Release the mask keeping both mask path and content.
     *
     * マスクを特定できない場合でも解除は行い、塗り適用のみスキップします（非破壊）。
     *
     * @param {GroupItem} clippingGroup - 対象のクリップグループ。
     * @param {boolean} shouldApplyFill - マスクパスへ塗りを適用する場合は true。
     * @returns {void}
     */
    function releaseKeepBoth(clippingGroup, shouldApplyFill) {
        var maskItem = findMaskItem(clippingGroup);
        clippingGroup.clipped = false;
        if (shouldApplyFill && maskItem) {
            applyMaskFill(maskItem);
        }
        ungroup(clippingGroup);
    }

    /**
     * マスク内容を残し、マスクパスを削除します。/ Keep masked content, remove the mask path.
     *
     * @param {GroupItem} clippingGroup - 対象のクリップグループ。
     * @returns {void}
     * @throws {Error} マスクを特定できない場合（このグループの処理を中止）。
     */
    function releaseKeepContent(clippingGroup) {
        /* 破壊操作の前にマスクを確定（clipping フラグ喪失を避ける）
           / Identify the mask before any destructive op */
        var maskItem = findMaskItem(clippingGroup);
        if (!maskItem) {
            throw new Error("Clipping mask item not found.");
        }
        clippingGroup.clipped = false;
        maskItem.remove();
        ungroup(clippingGroup);
    }

    /**
     * マスクパスを残し、マスク内容（画像・グループ・テキスト等）を全て削除します。/ Keep the mask path, remove all masked content.
     *
     * @param {GroupItem} clippingGroup - 対象のクリップグループ。
     * @param {boolean} shouldApplyFill - 残ったマスクパスへ塗りを適用する場合は true。
     * @returns {void}
     * @throws {Error} マスクを特定できない場合（全削除の誤動作を防ぐため処理を中止）。
     */
    function releaseKeepMaskPath(clippingGroup, shouldApplyFill) {
        /* 破壊操作の前にマスクと削除対象を確定 / Determine mask & targets before any destructive op */
        var plan = collectContentToRemove(clippingGroup);
        if (!plan.maskItem) {
            throw new Error("Clipping mask item not found.");
        }
        clippingGroup.clipped = false;
        removeItems(plan.itemsToRemove);
        if (shouldApplyFill) {
            applyMaskFill(plan.maskItem);
        }
        ungroup(clippingGroup);
    }

    /**
     * 収集済みのアイテムを順に削除します。/ Remove the collected items in order.
     *
     * @param {Array<PageItem>} items - 削除対象アイテムの配列。
     * @returns {void}
     */
    function removeItems(items) {
        for (var i = 0; i < items.length; i++) {
            items[i].remove();
        }
    }

    /**
     * マスクパスにK100・指定不透明度の塗りを適用します。/ Apply K100 fill at the given opacity to the mask path.
     *
     * PathItem はそのまま、CompoundPathItem は内部 pathItems へ塗りを設定します。
     * それ以外の型（テキストマスク等）は対象外です。
     *
     * @param {PageItem} maskItem - 塗りを適用するマスクアイテム。
     * @returns {void}
     */
    function applyMaskFill(maskItem) {
        if (!maskItem) return;
        if (maskItem.typename === "PathItem") {
            maskItem.filled = true;
            maskItem.fillColor = getK100Black();
            maskItem.opacity = PATH_FILL_OPACITY;
        } else if (maskItem.typename === "CompoundPathItem") {
            maskItem.opacity = PATH_FILL_OPACITY;
            var subPaths = maskItem.pathItems;
            for (var i = 0; i < subPaths.length; i++) {
                subPaths[i].filled = true;
                subPaths[i].fillColor = getK100Black();
            }
        }
    }

    // =========================================
    // ユーティリティ / Utilities
    // =========================================

    /**
     * K100（スミベタ）の CMYK カラーを生成して返します。/ Return a K100 CMYK color.
     *
     * @returns {CMYKColor} K=100、他0の CMYK カラー。
     */
    function getK100Black() {
        var cmykColor = new CMYKColor();
        cmykColor.cyan = 0;
        cmykColor.magenta = 0;
        cmykColor.yellow = 0;
        cmykColor.black = 100;
        return cmykColor;
    }

    /**
     * グループを解除します。子要素をグループの直前へ移動して重なり順を保ちます。/ Ungroup, preserving stacking order.
     *
     * 子を親末尾（PLACEATEND）へ送ると周囲との前後関係が崩れるため、グループ自身を基準に
     * PLACEBEFORE で移動し、元のグループ位置・内部順序を維持します。
     *
     * @param {GroupItem} clippingGroup - 解除対象のグループ。
     * @returns {void}
     */
    function ungroup(clippingGroup) {
        while (clippingGroup.pageItems.length > 0) {
            clippingGroup.pageItems[0].move(clippingGroup, ElementPlacement.PLACEBEFORE);
        }
        clippingGroup.remove();
    }

    // =========================================
    // エントリーポイント / Entry point
    // =========================================

    main();

})();
