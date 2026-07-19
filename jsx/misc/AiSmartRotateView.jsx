#target illustrator
#targetengine "AiSmartRotateView"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

- 開いているドキュメントのアクティブビューの回転角度（表示の回転）を取得
- パレットに回転角度を表示
- 「適用」を押すと、環境設定の「角度の制限」（Shiftキーを押したときの角度）に値を設定

たとえば表示を30度回転して作業しているとき、Shiftキーで回転させる角度も30度にそろい、
表示と操作の角度が一致して作業しやすくなります。

パレット（常駐ウィンドウ）化のため、DOMを参照する処理（ビュー回転角度の取得・環境設定の適用）は
ボタンを押すたびにメインエンジンへ BridgeTalk で委譲します。

### Overview

- Get the active view rotation angle (view rotation) of the current document
- Show the rotation angle in a palette
- Click "Apply" to set the "constrain angle" preference (Shift-key angle)

For example, when working with the view rotated 30 degrees, the Shift-key rotation angle
also becomes 30 degrees, so the view and operation angles match for easier work.

Because this is a palette (persistent window), DOM access (reading the view rotation and
applying the preference) is delegated to the main engine via BridgeTalk on each button press.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiSmartRotateView";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/* 同一リポジトリ内の transform フォルダのパス（外部スクリプト実行用）/ Path to the sibling transform folder (for running external scripts) */
var TRANSFORM_FOLDER = (function () {
    try {
        return File($.fileName).parent.parent.fsName + "/transform";
    } catch (e) {
        return "";
    }
})();

// =========================================
// ローカライズ / Localization
// =========================================

/* 言語判定 / Detect language */
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var currentLanguage = getCurrentLang();

/* ラベル定義 / Label definitions */
var LABELS = {
    dialog: {
        title: { ja: "ビューの回転角度", en: "View Rotation Angle" }
    },
    panel: {
        info: { ja: "アクティブビューの回転角度", en: "Active View Rotation Angle" },
        selection: { ja: "選択したオブジェクト", en: "Selected Object" },
        reset: { ja: "リセット", en: "Reset" },
        constrain: { ja: "角度の制限", en: "Constrain Angle" }
    },
    label: {
        rotation: { ja: "アクティブビューの回転角度", en: "Active view rotation angle" },
        selection: { ja: "選択したオブジェクトの角度", en: "Selected object angle" },
        constrain: { ja: "角度の制限", en: "Constrain angle" }
    },
    button: {
        apply: { ja: "「角度の制限」の値を変更", en: "Change constrain angle" },
        rotateToSelection: { ja: "選択したオブジェクトに合わせてビューを回転", en: "Rotate view to match selection" },
        rotateSelectionToView: { ja: "選択したオブジェクトをビューの回転に合わせて回転", en: "Rotate selection to match view" },
        resetSelectionRotation: { ja: "選択したオブジェクトの回転をリセット", en: "Reset selection rotation" },
        resetTextTilt: { ja: "テキストの傾きをリセット", en: "Reset text tilt" },
        resetImageTilt: { ja: "画像の傾きをリセット", en: "Reset image tilt" },
        resetRotation: { ja: "リセット", en: "Reset" },
        resetConstrain: { ja: "リセット", en: "Reset" },
        close: { ja: "閉じる", en: "Close" }
    },
    status: {
        applied: { ja: "制限角度に適用しました。", en: "Applied to the constrain angle." },
        resetRotation: { ja: "ビューの回転を0°にリセットしました。", en: "Reset the view rotation to 0°." },
        resetConstrain: { ja: "制限角度を0°にリセットしました。", en: "Reset the constrain angle to 0°." },
        rotatedToSelection: { ja: "選択したオブジェクトに合わせてビューを回転しました。", en: "Rotated the view to match the selection." },
        rotatedSelectionToView: { ja: "選択したオブジェクトをビューの回転に合わせて回転しました。", en: "Rotated the selection to match the view." },
        resetSelectionRotation: { ja: "選択したオブジェクトの回転をリセットしました。", en: "Reset the selection's rotation." }
    },
    alert: {
        noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
        noSelection: { ja: "オブジェクトが選択されていません。", en: "No object is selected." },
        noAngle: { ja: "角度を取得できるパスを選択してください。", en: "Select a path whose angle can be measured." },
        noTag: { ja: "回転情報（BBAccumRotation）が見つかりません。", en: "Rotation data (BBAccumRotation) not found." },
        noFile: { ja: "スクリプトファイルが見つかりません。", en: "Script file not found." },
        invalidAngle: { ja: "角度には数値を入力してください。", en: "Please enter a numeric angle." },
        error: { ja: "エラーが発生しました：", en: "An error occurred:" }
    }
};

/* ネストしたラベルをドット区切りパスで取得 / Get a nested label by dot-separated path */
function getLabel(path) {
    var parts = path.split(".");
    var node = LABELS;
    for (var i = 0; i < parts.length; i++) {
        node = node[parts[i]];
    }
    return node[currentLanguage];
}
function L(path) {
    return getLabel(path);
}

/* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
function labelText(path) {
    return getLabel(path) + (currentLanguage === "ja" ? "：" : ":");
}

// =========================================
// 結果ハンドリング / Result handling
// =========================================

/* ワーカーの 'ERR:XXX' コードを対応するアラート文へ変換 / Map a worker 'ERR:XXX' code to its alert text */
var RESULT_ERROR_LABELS = {
    NODOC: "alert.noDocument",
    NOSEL: "alert.noSelection",
    NOANGLE: "alert.noAngle",
    NOTAG: "alert.noTag",
    NOFILE: "alert.noFile"
};

/* 結果文字列を表示用ステータスへ変換（既知コードは専用文、未知は汎用エラー）
   / Convert a result string to status text (known codes get a specific message, unknown ones the generic error) */
function statusFromResult(result) {
    for (var code in RESULT_ERROR_LABELS) {
        if (RESULT_ERROR_LABELS.hasOwnProperty(code) && result.indexOf(code) !== -1) {
            return L(RESULT_ERROR_LABELS[code]);
        }
    }
    return L("alert.error") + " " + result;
}

// =========================================
// UI共通設定 / UI helpers
// =========================================

/* パネルの余白と間隔 / Panel margins and spacing */
var PANEL_MARGINS = [16, 20, 16, 12];
var PANEL_SPACING = 8;

/* パネルの共通設定 / Apply shared panel layout */
function setupPanel(panel, spacing) {
    panel.orientation = "column";
    panel.alignChildren = ["fill", "top"];
    panel.alignment = "fill";
    panel.margins = PANEL_MARGINS;
    panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/* グループの共通設定（row/column で整列を切り替え）/ Apply shared group layout (alignChildren switches by orientation) */
function setupGroup(group, orientation, spacing) {
    var groupOrientation = orientation || "column";
    group.orientation = groupOrientation;
    /* row は横並びなので縦中央、column は縦並びなので左揃え / row: vertically centered, column: left-aligned */
    group.alignChildren = (groupOrientation === "row") ? ["left", "center"] : ["left", "top"];
    group.alignment = "fill";
    group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

// =========================================
// 角度の計算 / Angle helpers
// =========================================

/* 角度を -180〜180 に正規化 / Normalize an angle into the -180..180 range */
function normalizeAngle(angle) {
    var a = angle % 360;
    if (a > 180) { a -= 360; }
    if (a < -180) { a += 360; }
    return a;
}

/* スライダー値を整数に丸める。Shift併用時は15°単位にクランプ / Round the slider value to an integer; clamp to 15° steps when Shift is held */
function snapAngle(angle, useShift) {
    if (useShift) {
        return Math.round(angle / 15) * 15;
    }
    return Math.round(angle);
}

/* 表示用に小数2桁へ丸める（atan2 由来の 30.00000001 のような桁あふれを抑える）
   / Round to 2 decimals for display (suppresses float noise like 30.00000001 from atan2) */
function roundAngle(angle) {
    return Math.round(angle * 100) / 100;
}

/* 角度を表示用文字列にする（丸め＋°付き）/ Format an angle for display (rounded, with the degree sign) */
function formatAngle(angle) {
    return roundAngle(angle) + "°";
}

// =========================================
// メインエンジンへの委譲 / Delegation to the main engine
// =========================================

/* メインエンジンでコードを実行する（常駐パレットの app は DOM 接続を失うため）。
   本文は encodeURIComponent + eval で送り、バックスラッシュ・多バイト文字を無傷で渡す。
   / Run code in the main engine (the palette's app loses DOM access).
   The body is sent via encodeURIComponent + eval so backslashes and multibyte chars survive intact. */
function runInMainEngine(code, onResult) {
    var bridge = new BridgeTalk();
    bridge.target = "illustrator";
    bridge.body = 'eval(decodeURIComponent("' + encodeURIComponent(code) + '"));';
    bridge.onResult = function (response) {
        onResult(String(response.body));
    };
    bridge.onError = function (response) {
        onResult("ERR:" + String(response.body));
    };
    bridge.send();
}

/* worker 本文を IIFE で包む / Wrap a worker body in an IIFE */
function workerBody(body) {
    return "(function(){" + body + "})()";
}

/* worker 断片：ドキュメント有無を確認し doc を確保 / Worker fragment: guard the open document, bind `doc` */
var W_DOC = "if(app.documents.length===0){return 'ERR:NODOC';}var doc=app.activeDocument;";

/* worker 断片：選択を確認し sel を確保（W_DOC の後に置く）/ Worker fragment: guard the selection, bind `sel` (place after W_DOC) */
var W_SEL = "var sel=doc.selection;if(!sel||sel.length===0){return 'ERR:NOSEL';}";

/* worker 断片：sel[0] が2点以上のパスか確認し、最初の2点を結ぶ線の傾き deg を算出（W_SEL の後）
   / Worker fragment: verify sel[0] is a path with >=2 points and compute the slope `deg` of its first two anchors (place after W_SEL) */
var W_PATH_DEG =
    "var item=sel[0];" +
    "if(item.typename!=='PathItem'||item.pathPoints.length<2){return 'ERR:NOANGLE';}" +
    "var a=item.pathPoints[0].anchor,b=item.pathPoints[1].anchor;" +
    "var deg=Math.atan2(b[1]-a[1],b[0]-a[0])*180/Math.PI;";

/* ビュー回転角度・制限角度・選択角度を1回の委譲でまとめて取得（"OK:回転,制限,選択"。選択は無効なら空）
   状態取得を1本化することで、別々に投げていた頃の表示ズレを防ぐ。
   / Fetch the view rotation, constrain angle, and selection angle in a single delegation ("OK:rotation,constrain,selection"; selection is empty when invalid).
   Consolidating the state read avoids the display drift from issuing two separate calls. */
function fetchState(onResult) {
    runInMainEngine(workerBody(
        W_DOC +
        "var r=doc.activeView.rotateAngle;" +
        "var c=app.preferences.getRealPreference('constrain/angle');" +
        "var s='';" +
        "var sel=doc.selection;" +
        "if(sel&&sel.length>0){var item=sel[0];" +
        "if(item.typename==='PathItem'&&item.pathPoints.length>=2){" +
        "var a=item.pathPoints[0].anchor,b=item.pathPoints[1].anchor;" +
        "s=Math.atan2(b[1]-a[1],b[0]-a[0])*180/Math.PI;}}" +
        "return 'OK:'+r+','+c+','+s;"
    ), onResult);
}

/* ビューの回転角度を設定（制限角度は自動では変えず、ユーザーが適用ボタンを押したときだけ反映）
   / Set the view rotation angle only (the constrain angle is not changed automatically; it is applied only when the user presses the Apply button) */
function applyViewRotation(angle, onResult) {
    runInMainEngine(workerBody(
        W_DOC +
        "doc.activeView.rotateAngle=" + angle + ";" +
        "return 'OK';"
    ), onResult);
}

/* 制限角度をメインエンジンで環境設定に適用 / Apply the constrain angle to the preference in the main engine */
function applyConstrainAngle(angle, onResult) {
    runInMainEngine(workerBody(
        "app.preferences.setRealPreference('constrain/angle'," + angle + ");" +
        "return 'OK';"
    ), onResult);
}

/* ビューの回転角度だけ0°に戻す / Reset only the view rotation to 0° */
function resetViewRotation(onResult) {
    runInMainEngine(workerBody(
        W_DOC +
        "doc.activeView.rotateAngle=0;" +
        "return 'OK';"
    ), onResult);
}

/* 「角度の制限」だけ0°に戻す / Reset only the constrain angle to 0° */
function resetConstrain(onResult) {
    runInMainEngine(workerBody(
        "app.preferences.setRealPreference('constrain/angle',0);" +
        "return 'OK';"
    ), onResult);
}

/* 選択パスの角度（最初の2アンカー点を結ぶ線の傾き）を取得してビューを回転（"OK:角度"）
   / Measure the selected path's angle (slope of the first two anchors), rotate the view to it ("OK:angle") */
function rotateViewToSelection(onResult) {
    runInMainEngine(workerBody(
        W_DOC + W_SEL + W_PATH_DEG +
        "doc.activeView.rotateAngle=deg;" +
        "return 'OK:'+deg;"
    ), onResult);
}

/* 選択オブジェクトをビューの回転角度だけ回転し、バウンディングボックスをリセット
   / Rotate the selected objects by the view rotation angle, then reset the bounding box */
function rotateSelectionToView(onResult) {
    runInMainEngine(workerBody(
        W_DOC + W_SEL +
        "var ang=doc.activeView.rotateAngle;" +
        "for(var i=0;i<sel.length;i++){sel[i].rotate(ang);}" +
        "app.executeMenuCommand('AI Reset Bounding Box');" +
        "return 'OK';"
    ), onResult);
}

/* 指定したスクリプトファイルをメインエンジンで実行（各スクリプトが自前で選択チェック・アラートを行う）
   / Run a script file in the main engine (each script handles its own selection checks and alerts) */
function runScriptFile(path, onResult) {
    runInMainEngine(workerBody(
        "var f=new File('" + path + "');" +
        "if(!f.exists){return 'ERR:NOFILE';}" +
        "$.evalFile(f);" +
        "return 'OK';"
    ), onResult);
}

/* 選択オブジェクトの回転をリセット（BBAccumRotation タグの蓄積回転を戻す）
   / Reset the selected object's rotation (undo the accumulated rotation in the BBAccumRotation tag) */
function resetSelectionRotation(onResult) {
    runInMainEngine(workerBody(
        W_DOC + W_SEL +
        "var item=sel[0];" +
        "if(item.tags.length>0&&item.tags[0].name==='BBAccumRotation'){" +
        "var deg=180*parseFloat(item.tags[0].value)/Math.PI;" +
        "item.rotate(deg);" +
        "return 'OK';" +
        "}" +
        "return 'ERR:NOTAG';"
    ), onResult);
}

// =========================================
// パレット / Palette
// =========================================

/* パレットを作成して表示する（IIFEで即時実行）/ Build and show the palette (run immediately as an IIFE) */
(function () {
    var palette = new Window("palette", L("dialog.title") + " " + SCRIPT_VERSION);
    palette.orientation = "column";
    palette.alignChildren = "fill";
    palette.margins = 16;
    palette.spacing = 12;

    /* ビューの回転角度を表示するパネル / Panel showing the view rotation angle */
    var infoPanel = palette.add("panel", undefined, L("panel.info"));
    setupPanel(infoPanel, 6);

    /* アクティブビューの回転角度（パネルタイトルと重複するため内側ラベルは省略）/ Active view rotation angle (inner label omitted; the panel title already states it) */
    var rotationGroup = infoPanel.add("group");
    setupGroup(rotationGroup, "row");
    var rotationValue = rotationGroup.add("statictext", undefined, "—°");
    rotationValue.preferredSize.width = 50;

    /* 回転角度スライダー（-180〜180、Shiftで15°単位にクランプ）/ Rotation slider (-180..180, snaps to 15° with Shift) */
    var rotationSlider = infoPanel.add("slider", undefined, 0, -180, 180);
    rotationSlider.alignment = "fill";

    /* ビューの回転だけ0°に戻す / Reset only the view rotation to 0° */
    var resetRotationButton = infoPanel.add("button", undefined, L("button.resetRotation"));
    resetRotationButton.alignment = "right";

    /* 角度の制限を変更するパネル / Panel for changing the constrain angle */
    var constrainPanel = palette.add("panel", undefined, L("panel.constrain"));
    setupPanel(constrainPanel, 6);

    /* 角度の制限（編集可。回転角度に追従して候補値が入るが、反映は適用ボタンを押したときだけ）
       / Constrain angle (editable; the field tracks the rotation as a suggested value, but is committed only on the Apply button) */
    var constrainGroup = constrainPanel.add("group");
    setupGroup(constrainGroup, "row");
    constrainGroup.add("statictext", undefined, labelText("label.constrain"));
    var constrainInput = constrainGroup.add("edittext", undefined, "");
    constrainInput.characters = 6;
    constrainGroup.add("statictext", undefined, "°");

    /* ボタン行：適用と、制限角度だけのリセット / Button row: Apply, and reset only the constrain angle */
    var constrainButtonGroup = constrainPanel.add("group");
    setupGroup(constrainButtonGroup, "row");
    constrainButtonGroup.alignment = "right";

    /* 角度の制限だけ0°に戻す / Reset only the constrain angle to 0° */
    var resetConstrainButton = constrainButtonGroup.add("button", undefined, L("button.resetConstrain"));

    /* 「角度の制限」の値を変更ボタン / Change-constrain-angle button */
    var applyButton = constrainButtonGroup.add("button", undefined, L("button.apply"));

    /* 選択したオブジェクトのパネル / Selected-object panel */
    var selectionPanel = palette.add("panel", undefined, L("panel.selection"));
    setupPanel(selectionPanel, 6);

    /* 選択したオブジェクトの角度（表示）/ Selected object angle (display) */
    var selectionGroup = selectionPanel.add("group");
    setupGroup(selectionGroup, "row");
    selectionGroup.add("statictext", undefined, labelText("label.selection"));
    var selectionValue = selectionGroup.add("statictext", undefined, "—°");
    selectionValue.preferredSize.width = 50;

    /* 選択に合わせてビューを回転 / Rotate the view to match the selection */
    var rotateViewButton = selectionPanel.add("button", undefined, L("button.rotateToSelection"));
    rotateViewButton.alignment = "left";

    /* 選択をビューの回転に合わせて回転 / Rotate the selection to match the view rotation */
    var rotateSelectionButton = selectionPanel.add("button", undefined, L("button.rotateSelectionToView"));
    rotateSelectionButton.alignment = "left";

    /* 選択の回転をリセット / Reset the selection's rotation */
    var resetSelectionButton = selectionPanel.add("button", undefined, L("button.resetSelectionRotation"));
    resetSelectionButton.alignment = "left";

    /* リセットパネル（外部スクリプトを実行）/ Reset panel (runs external scripts) */
    var resetPanel = palette.add("panel", undefined, L("panel.reset"));
    setupPanel(resetPanel, 6);

    /* テキストの傾き（ResetText.jsx）/ Text tilt (ResetText.jsx) */
    var resetTextButton = resetPanel.add("button", undefined, L("button.resetTextTilt"));
    resetTextButton.alignment = "left";

    /* 画像の傾き（ResetRotation.jsx）/ Image tilt (ResetRotation.jsx) */
    var resetImageButton = resetPanel.add("button", undefined, L("button.resetImageTilt"));
    resetImageButton.alignment = "left";

    /* ステータス表示 / Status line */
    var statusText = palette.add("statictext", undefined, "");
    statusText.alignment = "fill";

    /* 現在の状態（各リセットボタンのディム判定に使用）/ Current state (used to dim each Reset button) */
    var currentRotation = 0;
    var currentConstrain = 0;

    /* 0°のリセットボタンをそれぞれディム（回転と制限は独立）/ Dim each Reset button when its value is 0° (rotation and constrain are independent) */
    function updateResetState() {
        resetRotationButton.enabled = (currentRotation !== 0);
        resetConstrainButton.enabled = (currentConstrain !== 0);
    }

    /* 回転表示・スライダー・状態を更新し、制限入力欄に回転角度を候補値として入れる（適用はボタン押下時のみ）
       / Update the rotation display, slider, and state; seed the constrain field with the rotation as a suggestion (applied only on the button) */
    function setRotation(angle) {
        currentRotation = normalizeAngle(angle);
        rotationValue.text = formatAngle(currentRotation);
        rotationSlider.value = currentRotation;
        constrainInput.text = roundAngle(currentRotation);
        updateResetState();
    }

    /* 適用済みの制限角度を状態に記録（入力欄は触らない。回転追従とユーザー入力を尊重）
       / Record the applied constrain angle in state (does not touch the field, preserving rotation-follow and user input) */
    function setConstrain(angle) {
        currentConstrain = angle;
        updateResetState();
    }

    /* ビュー回転角度・制限角度・選択角度を1回の委譲で取得して表示へ反映 / Fetch the view rotation, constrain angle, and selection angle in a single delegation and reflect them in the display */
    function refresh() {
        fetchState(function (result) {
            if (result.indexOf("OK:") === 0) {
                var parts = result.substring(3).split(",");
                setRotation(parseFloat(parts[0]));
                setConstrain(parseFloat(parts[1]));
                selectionValue.text = (parts[2] === "") ? "—°" : formatAngle(normalizeAngle(parseFloat(parts[2])));
                statusText.text = "";
            } else {
                if (result.indexOf("NODOC") !== -1) {
                    rotationValue.text = "—°";
                    selectionValue.text = "—°";
                }
                statusText.text = statusFromResult(result);
            }
        });
    }

    /* スライダー操作中：表示だけ更新（Shiftで15°クランプ）/ While dragging: update the display only (snap to 15° with Shift) */
    rotationSlider.onChanging = function () {
        var useShift = ScriptUI.environment.keyboardState.shiftKey;
        var angle = snapAngle(rotationSlider.value, useShift);
        if (useShift) { rotationSlider.value = angle; }
        rotationValue.text = angle + "°";
    };

    /* スライダー確定：ビューの回転を適用 / On release: apply the view rotation */
    rotationSlider.onChange = function () {
        var useShift = ScriptUI.environment.keyboardState.shiftKey;
        var angle = snapAngle(rotationSlider.value, useShift);
        setRotation(angle);
        applyViewRotation(angle, function (result) {
            if (result.indexOf("OK") === 0) {
                statusText.text = "";
            } else {
                statusText.text = statusFromResult(result);
            }
        });
    };

    /* 選択に合わせてビューを回転：選択パスの角度を取得し、表示とビュー回転へ反映 / Rotate view to match selection: measure the path angle and reflect it in the display and view */
    rotateViewButton.onClick = function () {
        rotateViewToSelection(function (result) {
            if (result.indexOf("OK:") === 0) {
                var angle = parseFloat(result.substring(3));
                selectionValue.text = formatAngle(normalizeAngle(angle));
                setRotation(angle);
                statusText.text = L("status.rotatedToSelection");
            } else {
                statusText.text = statusFromResult(result);
            }
        });
    };

    /* 選択をビューの回転に合わせて回転 / Rotate the selection to match the view rotation */
    rotateSelectionButton.onClick = function () {
        rotateSelectionToView(function (result) {
            if (result.indexOf("OK") === 0) {
                statusText.text = L("status.rotatedSelectionToView");
            } else {
                statusText.text = statusFromResult(result);
            }
        });
    };

    /* 選択の回転をリセット / Reset the selection's rotation */
    resetSelectionButton.onClick = function () {
        resetSelectionRotation(function (result) {
            if (result.indexOf("OK") === 0) {
                statusText.text = L("status.resetSelectionRotation");
            } else {
                statusText.text = statusFromResult(result);
            }
        });
    };

    /* 外部スクリプトを実行する共通ハンドラ / Shared handler that runs an external script */
    function runExternalScript(fileName) {
        runScriptFile(TRANSFORM_FOLDER + "/" + fileName, function (result) {
            if (result.indexOf("OK") === 0) {
                statusText.text = "";
            } else {
                statusText.text = statusFromResult(result);
            }
        });
    }

    /* テキストの傾きをリセット（ResetText.jsx）/ Reset text tilt (ResetText.jsx) */
    resetTextButton.onClick = function () {
        runExternalScript("ResetText.jsx");
    };

    /* 画像の傾きをリセット（ResetRotation.jsx）/ Reset image tilt (ResetRotation.jsx) */
    resetImageButton.onClick = function () {
        runExternalScript("ResetRotation.jsx");
    };

    /* リセット（ビューの回転だけ）：0°に戻して表示を更新 / Reset (view rotation only): set it to 0° and refresh the display */
    resetRotationButton.onClick = function () {
        resetViewRotation(function (result) {
            if (result.indexOf("OK") === 0) {
                setRotation(0);
                statusText.text = L("status.resetRotation");
            } else {
                statusText.text = statusFromResult(result);
            }
        });
    };

    /* リセット（制限角度だけ）：0°に戻して入力欄と状態を更新 / Reset (constrain angle only): set it to 0° and refresh the field and state */
    resetConstrainButton.onClick = function () {
        resetConstrain(function (result) {
            if (result.indexOf("OK") === 0) {
                setConstrain(0);
                constrainInput.text = "0";
                statusText.text = L("status.resetConstrain");
            } else {
                statusText.text = statusFromResult(result);
            }
        });
    };

    /* 適用ボタン：入力値を検証してメインエンジンで適用 / Apply button: validate the input and apply in the main engine */
    applyButton.onClick = function () {
        var constrainAngle = parseFloat(constrainInput.text);
        if (isNaN(constrainAngle)) {
            statusText.text = L("alert.invalidAngle");
            return;
        }
        applyConstrainAngle(constrainAngle, function (result) {
            if (result.indexOf("OK") === 0) {
                setConstrain(constrainAngle);
                statusText.text = L("status.applied");
            } else {
                statusText.text = statusFromResult(result);
            }
        });
    };

    /* 初期表示時、およびパレットがアクティブになるたびに最新のビュー回転角度を取得
       / Fetch the latest view rotation on first show and whenever the palette becomes active */
    palette.onShow = refresh;
    palette.onActivate = refresh;

    palette.show();
})();
