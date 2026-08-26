#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトのグラデーションを、指定した数の単色オブジェクトに分割します。
分割後の後処理として、重なりを整理してひとまとめにする／両端からブレンドを作成する、を選べます。

詳細は README を参照してください。

### Overview

Splits the gradient on the selected objects into a given number of solid-color objects.
Post-processing can tidy the overlaps into a single set or build a blend from the two end objects.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ExpandGradient";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-05-25";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-27";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ExpandGradient.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ExpandGradient.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nbe084e691ba5"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    /* 分割数（既定値、2 以上の整数） / Default gradient step count (integer ≥ 2) */
    var DEFAULT_GRADIENT_STEPS = 5;

    /* 分割・拡張の補正値（Illustrator は指定値より1つ少ないオブジェクトを生成する） / Offset for Expand (Illustrator yields one object fewer than specified) */
    var EXPAND_STEP_OFFSET = 1;

    /* 「ブレンドに変換」時のステップ数（固定・ダイアログではディム表示） / Fixed step count for "Convert to blend" (dimmed in the dialog) */
    var BLEND_FIXED_STEPS = 2;

    /* 後処理モードの既定値（"none" / "simple" / "blend"） / Default post-process mode ("none" / "simple" / "blend") */
    var DEFAULT_POST_PROCESS_MODE = "simple";

    /* ダイアログ表示の有無（false で既定値のまま即実行） / Show dialog (false: run silently with default) */
    var SHOW_DIALOG = true;

    /* 一時アクションのセット名（衝突回避のためユニーク名） / Temporary action set name (unique to avoid collisions) */
    var ACTION_SET_NAME = "ExpandGradient_tmp";

    /* 一時アクションのアクション名 / Temporary action name */
    var ACTION_NAME = "Expand-gradient";


    (function () {

        // =========================================
        // ローカライズ / Localization
        // =========================================

        function getCurrentLanguage() {
            return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
        }
        var currentLanguage = getCurrentLanguage();

        /* 日英ラベル定義 / Japanese-English label definitions */
        var LABELS = {

            /* === 共通 / Common === */
            cancel: {
                ja: "キャンセル",
                en: "Cancel"
            },
            /* === ダイアログ / Dialog === */
            dialogTitle: {
                ja: "グラデーションを分割・拡張",
                en: "Expand Gradient"
            },
            steps: {
                ja: "ステップ数",
                en: "Steps"
            },
            postProcessPanel: {
                ja: "実行後の処理",
                en: "Post-Processing"
            },
            postProcessNone: {
                ja: "なし",
                en: "None"
            },
            postProcessSimple: {
                ja: "単純に拡張",
                en: "Simple expand"
            },
            postProcessBlend: {
                ja: "ブレンドに変換",
                en: "Convert to blend"
            },

            /* === アラート / Alerts === */
            alertNoDocument: {
                ja: "ドキュメントが開かれていません。",
                en: "No document is open."
            },
            alertNoSelection: {
                ja: "オブジェクトを選択してください。",
                en: "Please select an object."
            },
            alertInvalidSteps: {
                ja: "ステップ数は 2 以上の整数で指定してください。",
                en: "Steps must be an integer of 2 or more."
            },
            alertMergeTargetNotFound: {
                ja: "Pathfinder Merge を適用できる対象を取得できませんでした。処理を中断します。",
                en: "Could not find a valid target for applying Pathfinder Merge. The process will stop."
            },
            alertBlendTargetNotFound: {
                ja: "ブレンドに必要なオブジェクトを取得できませんでした。処理を中断します。",
                en: "Could not get the objects needed for the blend. The process will stop."
            }
        };

        function L(key) {
            return (LABELS[key] && LABELS[key][currentLanguage]) ? LABELS[key][currentLanguage] : key;
        }

        function labelText(key) {
            return L(key) + (currentLanguage === "ja" ? "：" : ":");
        }

        // =========================================
        // 入口チェック / Entry checks
        // =========================================

        if (app.documents.length === 0) {
            alert(L("alertNoDocument"));
            return;
        }

        var activeDoc = app.activeDocument;
        if (activeDoc.selection.length === 0) {
            alert(L("alertNoSelection"));
            return;
        }

        // =========================================
        // ダイアログ / Dialog
        // =========================================

        var gradientSteps = DEFAULT_GRADIENT_STEPS;
        var postProcessMode = DEFAULT_POST_PROCESS_MODE;

        if (SHOW_DIALOG) {
            var dialogResult = showStepsDialog(SCRIPT_VERSION, DEFAULT_GRADIENT_STEPS, DEFAULT_POST_PROCESS_MODE, L, labelText);
            if (dialogResult === null) return;
            gradientSteps = dialogResult.steps;
            postProcessMode = dialogResult.postProcessMode;
        }

        playEmbeddedAction(buildExpandActionSource(gradientSteps + EXPAND_STEP_OFFSET, ACTION_SET_NAME, ACTION_NAME), ACTION_SET_NAME, ACTION_NAME);

        if (postProcessMode === "simple") {
            if (!furtherExpandSelection(activeDoc)) {
                alert(L("alertMergeTargetNotFound"));
                return;
            }
        } else if (postProcessMode === "blend") {
            if (!convertToBlend(activeDoc)) {
                alert(L("alertBlendTargetNotFound"));
                return;
            }
        }

    })();

    // =========================================
    // 後処理 / Post-processing
    // =========================================

    /**
     * 分割・拡張の結果をクロップし、Pathfinder Merge で同色の重なりを整理する。
     * @param {Document} targetDoc - 対象ドキュメント
     * @returns {boolean} 整理できたら true、対象が見つからなければ false
     */
    function furtherExpandSelection(targetDoc) {
        targetDoc.activate();
        app.executeMenuCommand('Live Pathfinder Crop');
        app.executeMenuCommand('expandStyle');

        /* Pathfinder Merge ライブエフェクト（command 8）を適用し、アピアランスを分割
           Apply the Pathfinder Merge live effect (command 8) and expand appearance */
        var pathfinderMergeXml = '<LiveEffect name="Adobe Pathfinder" isPre="1">'
            + '<Dict data="I Command 8 B ConvertCustom 1 B ExtractUnpainted 1 R Mix 0.5 R Precision 10 B RemovePoints 1 R TrapAspect 1 B TrapConvertCustom 1 R TrapMaxTint 1 B TrapReverse 0 R TrapThickness 0.25 R TrapTint 0.4 R TrapTintTolerance 0.05">'
            + '<Entry name="DisplayString" value="Merge" valueType="S"/>'
            + '</Dict></LiveEffect>';

        var mergeTarget = getFirstEffectApplicableSelectionItem(targetDoc);
        if (mergeTarget === null) return false;

        mergeTarget.applyEffect(pathfinderMergeXml);
        app.redraw();
        app.executeMenuCommand("deselectall");
        mergeTarget.selected = true;
        app.executeMenuCommand('expandStyle');

        return true;
    }

    /**
     * 分割結果の両端だけを残し、ブレンドに置き換える。
     * @param {Document} targetDoc - 対象ドキュメント
     * @returns {boolean} 変換できたら true、対象が足りなければ false
     */
    function convertToBlend(targetDoc) {
        /* furtherExpandSelection と同じ前処理（Crop → expandStyle → Merge → expandStyle）
           Same pre-processing as furtherExpandSelection */
        if (!furtherExpandSelection(targetDoc)) return false;

        var endPaths = reduceToEndPaths(targetDoc);
        if (endPaths === null) return false;

        /* 2点を選択してブレンド作成 / Select the two items and run Blend Make */
        app.executeMenuCommand("deselectall");
        endPaths.backItem.selected = true;
        endPaths.frontItem.selected = true;
        app.executeMenuCommand('Path Blend Make');

        return true;
    }

    /**
     * 選択配下の塗りパスから最前面・最背面だけを残し、レイヤー直下へ移動する。
     * @param {Document} targetDoc - 対象ドキュメント
     * @returns {object} frontItem / backItem を持つオブジェクト、2つ揃わなければ null
     */
    function reduceToEndPaths(targetDoc) {
        /* 包んでいるグループ／コンパウンドは後で片付けるため記録
           Remember the wrapping containers so we can drop them later */
        var originalContainers = [];
        var paintedPaths = [];
        for (var i = 0; i < targetDoc.selection.length; i++) {
            originalContainers.push(targetDoc.selection[i]);
            collectPaintedDescendantPaths(targetDoc.selection[i], paintedPaths);
        }
        if (paintedPaths.length < 2) return null;

        /* 親共通なら親内 z 順（前面=0）で並べ替え / Sort by parent's z-order when parents are shared */
        paintedPaths = sortByZOrderWhenShared(paintedPaths);

        var frontItem = paintedPaths[0];
        var backItem = paintedPaths[paintedPaths.length - 1];
        for (var i = 1; i < paintedPaths.length - 1; i++) {
            paintedPaths[i].remove();
        }

        /* front/back をレイヤー直下へ移動し、空になった包みを除去
           Lift front/back to the layer and drop the now-empty wrappers */
        var hostLayer = targetDoc.activeLayer;
        frontItem.move(hostLayer, ElementPlacement.PLACEATEND);
        backItem.move(hostLayer, ElementPlacement.PLACEATEND);
        for (var i = 0; i < originalContainers.length; i++) {
            /* 選択項目自体が両端のパスだった場合は消さない / Keep the end paths when they were selected directly */
            if (originalContainers[i] === frontItem || originalContainers[i] === backItem) continue;
            /* 移動で空になった包みは既に消えている場合がある / The wrapper may already be gone */
            try { originalContainers[i].remove(); } catch (e) { }
        }

        return { frontItem: frontItem, backItem: backItem };
    }

    /**
     * 塗りまたは線を持つ葉パスを再帰的に収集する。
     * @param {object} item - 走査対象のページアイテム
     * @param {object[]} result - 収集先の配列
     * @returns {void}
     */
    function collectPaintedDescendantPaths(item, result) {
        if (!item) return;
        if (item.typename === "GroupItem") {
            for (var i = 0; i < item.pageItems.length; i++) {
                collectPaintedDescendantPaths(item.pageItems[i], result);
            }
        } else if (item.typename === "CompoundPathItem") {
            for (var i = 0; i < item.pathItems.length; i++) {
                collectPaintedDescendantPaths(item.pathItems[i], result);
            }
        } else if (item.typename === "PathItem") {
            if (item.clipping) return;
            if (!item.filled && !item.stroked) return;
            if (item.pathPoints && item.pathPoints.length === 2) return;
            result.push(item);
        }
    }

    /**
     * 親が共通のときだけ、親内の z 順（前面が先頭）に並べ替える。
     * @param {object[]} paths - 並べ替えるパスの配列
     * @returns {object[]} 並べ替えた配列、親がばらけていれば元の配列
     */
    function sortByZOrderWhenShared(paths) {
        if (paths.length < 2) return paths;

        var sharedParent = paths[0].parent;
        for (var i = 1; i < paths.length; i++) {
            if (paths[i].parent !== sharedParent) return paths;
        }

        /* 親の pageItems は前面から後面の順なので、その順に拾い直す
           The parent's pageItems run front to back, so collect in that order */
        var sorted = [];
        for (var i = 0; i < sharedParent.pageItems.length; i++) {
            for (var j = 0; j < paths.length; j++) {
                if (paths[j] === sharedParent.pageItems[i]) {
                    sorted.push(paths[j]);
                    break;
                }
            }
        }

        return (sorted.length === paths.length) ? sorted : paths;
    }

    /**
     * 選択中でライブ効果を適用できる最初のアイテムを返す。
     * @param {Document} targetDoc - 対象ドキュメント
     * @returns {object} 適用できるアイテム、見つからなければ null
     */
    function getFirstEffectApplicableSelectionItem(targetDoc) {
        if (!targetDoc || !targetDoc.selection || targetDoc.selection.length === 0) return null;

        for (var i = 0; i < targetDoc.selection.length; i++) {
            var selectedItem = targetDoc.selection[i];
            if (selectedItem && typeof selectedItem.applyEffect === "function") {
                return selectedItem;
            }
        }

        return null;
    }

    // =========================================
    // ダイアログUI / Dialog UI
    // =========================================

    /**
     * ステップ数と実行後の処理を指定するダイアログを表示する。
     * @param {string} scriptVersion - タイトルに表示するバージョン
     * @param {number} defaultSteps - ステップ数の初期値
     * @param {string} defaultPostProcessMode - 実行後の処理の初期値（"none" / "simple" / "blend"）
     * @param {function} L - ラベル取得関数
     * @param {function} labelText - コロン付きラベル取得関数
     * @returns {object} steps / postProcessMode を持つオブジェクト、キャンセル時は null
     */
    function showStepsDialog(scriptVersion, defaultSteps, defaultPostProcessMode, L, labelText) {

        var stepsDialog = new Window("dialog", L("dialogTitle") + " " + scriptVersion);
        stepsDialog.orientation = "column";
        stepsDialog.alignChildren = ["fill", "top"];
        stepsDialog.margins = 16;

        var stepsRow = stepsDialog.add("group");
        stepsRow.orientation = "row";
        stepsRow.alignChildren = ["left", "center"];
        stepsRow.add("statictext", undefined, labelText("steps"));
        var stepsInput = stepsRow.add("edittext", undefined, String(defaultSteps));
        stepsInput.characters = 5;
        stepsInput.active = true;
        changeValueByArrowKey(stepsInput);

        var postProcessPanel = stepsDialog.add("panel", undefined, L("postProcessPanel"));
        postProcessPanel.orientation = "column";
        postProcessPanel.alignChildren = ["left", "top"];
        postProcessPanel.margins = [15, 20, 15, 10];

        var postProcessNoneRb = postProcessPanel.add("radiobutton", undefined, L("postProcessNone"));
        var postProcessSimpleRb = postProcessPanel.add("radiobutton", undefined, L("postProcessSimple"));
        var postProcessBlendRb = postProcessPanel.add("radiobutton", undefined, L("postProcessBlend"));

        postProcessSimpleRb.value = (defaultPostProcessMode === "simple");
        postProcessBlendRb.value = (defaultPostProcessMode === "blend");
        postProcessNoneRb.value = (!postProcessSimpleRb.value && !postProcessBlendRb.value);

        /* 「ブレンドに変換」はステップ数を固定し、入力欄をディムにする（戻したときは元の値へ）
           Convert to blend fixes the step count and dims the field; the previous value returns on switch back */
        var keptStepsText = String(defaultSteps);
        function updateStepsAvailability() {
            if (postProcessBlendRb.value) {
                if (stepsInput.enabled) keptStepsText = stepsInput.text;
                stepsInput.text = String(BLEND_FIXED_STEPS);
                stepsInput.enabled = false;
                return;
            }
            if (!stepsInput.enabled) stepsInput.text = keptStepsText;
            stepsInput.enabled = true;
        }
        postProcessNoneRb.onClick = updateStepsAvailability;
        postProcessSimpleRb.onClick = updateStepsAvailability;
        postProcessBlendRb.onClick = updateStepsAvailability;
        updateStepsAvailability();

        var okCancelGroup = stepsDialog.add("group");
        okCancelGroup.alignment = ["right", "center"];
        okCancelGroup.add("button", undefined, L("cancel"), { name: "cancel" });
        okCancelGroup.add("button", undefined, "OK", { name: "ok" });

        if (stepsDialog.show() !== 1) return null;

        var parsedSteps = parseInt(stepsInput.text, 10);
        if (isNaN(parsedSteps) || parsedSteps < 2) {
            alert(L("alertInvalidSteps"));
            return null;
        }

        var selectedMode = "none";
        if (postProcessSimpleRb.value) selectedMode = "simple";
        else if (postProcessBlendRb.value) selectedMode = "blend";

        return { steps: parsedSteps, postProcessMode: selectedMode };

    }

    /**
     * 数値入力欄を↑↓キーで増減できるようにする（Shiftで10の倍数にスナップ）。
     * @param {object} editText - 対象の edittext
     * @returns {void}
     */
    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;

            var value = parseInt(editText.text, 10);
            if (isNaN(value)) return;

            var sign = (event.keyName === "Up") ? 1 : -1;
            if (ScriptUI.environment.keyboardState.shiftKey) {
                /* Shiftキー押下時は10の倍数にスナップ / Snap to multiples of 10 when Shift is held */
                value = (sign > 0) ? Math.ceil((value + 1) / 10) * 10 : Math.floor((value - 1) / 10) * 10;
            } else {
                value += sign;
            }

            if (value < 2) value = 2;
            editText.text = String(Math.round(value));
            event.preventDefault();
        });
    }

    // =========================================
    // 一時アクション生成 / Temporary action generation
    // =========================================

    /**
     * 分割・拡張（ai_plugin_expand）のアクションソースを組み立てる。
     * @param {number} gradientSteps - グラデーションの分割数
     * @param {string} setName - アクションセット名
     * @param {string} actionName - アクション名
     * @returns {string} .aia のソース
     */
    function buildExpandActionSource(gradientSteps, setName, actionName) {
        var parameters = [
            buildParameterLine(1, 1868720756, "boolean", 0, ""),
            buildParameterLine(2, 1718185068, "boolean", 1, ""),
            buildParameterLine(3, 1937011307, "boolean", 0, ""),
            buildParameterLine(4, 1936553064, "boolean", 0, ""),
            buildParameterLine(5, 1937007984, "integer", gradientSteps, "")
        ];
        /* 表示名「分割・拡張」 / Localized name */
        return buildActionSource(setName, actionName, "ai_plugin_expand", "e58886e589b2e383bbe68ba1e5bcb5", parameters);
    }

    /**
     * 1イベントだけのアクションセットを組み立てる。
     * @param {string} setName - アクションセット名
     * @param {string} actionName - アクション名
     * @param {string} internalName - プラグインの内部名
     * @param {string} localizedNameHex - パネル表示名のUTF-8 16進表現
     * @param {string[]} parameters - パラメーター行の配列
     * @returns {string} .aia のソース
     */
    function buildActionSource(setName, actionName, internalName, localizedNameHex, parameters) {
        return ''
            + '/version 3'
            + buildActionNameLine(setName)
            + '/isOpen 1'
            + '/actionCount 1'
            + '/action-1 {'
            + ' ' + buildActionNameLine(actionName)
            + ' /keyIndex 0'
            + ' /colorIndex 0'
            + ' /isOpen 1'
            + ' /eventCount 1'
            + ' /event-1 {'
            + ' /useRulersIn1stQuadrant 0'
            + ' /internalName (' + internalName + ')'
            + ' /localizedName ' + buildHexTextBlock(localizedNameHex)
            + ' /isOpen 1'
            + ' /isOn 1'
            + ' /hasDialog 1'
            + ' /showDialog 0'
            + ' /parameterCount ' + parameters.length
            + parameters.join('')
            + ' }'
            + '}';
    }

    /**
     * アクションのパラメーター1行を組み立てる。
     * @param {number} index - パラメーター番号（1始まり）
     * @param {number} key - パラメーターキー
     * @param {string} type - 値の型（boolean / integer / enumerated）
     * @param {number} value - 値
     * @param {string} localizedNameHex - 表示名のUTF-8 16進表現（不要なら空文字）
     * @returns {string} パラメーター行
     */
    function buildParameterLine(index, key, type, value, localizedNameHex) {
        return ' /parameter-' + index
            + ' { /key ' + key
            + ' /showInPalette 4294967295'
            + ' /type (' + type + ')'
            + (localizedNameHex ? ' /name ' + buildHexTextBlock(localizedNameHex) : '')
            + ' /value ' + value + ' }';
    }

    /**
     * アクション名の行を組み立てる。
     * @param {string} actionName - アクション名（ASCII）
     * @returns {string} /name の行
     */
    function buildActionNameLine(actionName) {
        return '/name ' + buildHexTextBlock(stringToHex(actionName)) + '\n';
    }

    /**
     * .aia の文字列表記 [ バイト数 16進 ] を組み立てる。
     * @param {string} hexText - 16進表現
     * @returns {string} [ バイト数 16進 ] の形式
     */
    function buildHexTextBlock(hexText) {
        return '[ ' + (hexText.length / 2) + ' ' + hexText + ' ]';
    }

    /**
     * 文字列を16進表現に変換する。
     * @param {string} sourceText - 変換元の文字列（ASCII）
     * @returns {string} 16進表現
     */
    function stringToHex(sourceText) {
        var hexText = "";
        for (var i = 0; i < sourceText.length; i++) {
            var hexValue = sourceText.charCodeAt(i).toString(16);
            if (hexValue.length < 2) hexValue = "0" + hexValue;
            hexText += hexValue;
        }
        return hexText;
    }

    // =========================================
    // 一時アクション実行 / Temporary action playback
    // =========================================

    /**
     * 一時的な .aia を書き出してロード・再生し、後始末まで行う。
     * @param {string} actionSource - .aia のソース
     * @param {string} setName - アクションセット名
     * @param {string} actionName - アクション名
     * @returns {void}
     */
    function playEmbeddedAction(actionSource, setName, actionName) {

        var actionFile = new File('~/ExpandGradientAction.aia');

        /* 既存の同名セットが残っているとロードが効かないので、先に外す / Unload first in case the set is still loaded */
        try { app.unloadAction(setName, ""); } catch (e) { }

        try {
            if (!actionFile.open('w')) {
                throw new Error('Failed to open temporary action file for writing.');
            }
            actionFile.write(actionSource);
            actionFile.close();

            app.loadAction(actionFile);
            app.doScript(actionName, setName, false);
        } finally {
            /* close / remove は失敗しても false を返すだけ / close and remove just return false on failure */
            actionFile.close();
            actionFile.remove();
            try { app.unloadAction(setName, ""); } catch (e) { }
        }

    }

})();
