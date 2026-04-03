#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
  Long Shadow Maker (Extrude and Merge with Preview)
  ダイアログで距離・角度・スケールを指定してロングシャドウを生成。プレビュー機能付き。
  v1.0: 初版。ロングシャドウ生成（簡易プレビュー付き）。
  更新日: 2026-02-09
*/

var SCRIPT_VERSION = "v1.0";

function getCurrentLang() {
  return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
  dialogTitle: {
    ja: "ロングシャドーメーカー",
    en: "Long Shadow Maker"
  },
  openDocument: {
    ja: "ドキュメントを開いてください。",
    en: "Please open a document."
  },
  selectOneClosed: {
    ja: "閉パス（単一パス／複合パス／グループ）を1つだけ選択してください。",
    en: "Please select exactly one closed shape (path / compound path / group)."
  },
  selectClosed: {
    ja: "閉パスを選択してください。",
    en: "Please select a closed path."
  },
  selectClosedGroup: {
    ja: "閉パスのグループを選択してください。",
    en: "Please select a group that consists of closed paths."
  },
  cannotBuildFromGroup: {
    ja: "グループから形状を作成できませんでした。閉パスのグループを選択してください。",
    en: "Could not build a shape from the group. Please select a group of closed paths."
  },
  preset: {
    ja: "プリセット",
    en: "Preset"
  },
  settings: {
    ja: "設定",
    en: "Settings"
  },
  offset: {
    ja: "オフセット",
    en: "Offset"
  },
  shape: {
    ja: "形状",
    en: "Corner"
  },
  distance: {
    ja: "距離",
    en: "Distance"
  },
  angle: {
    ja: "角度",
    en: "Angle"
  },
  scale: {
    ja: "スケール",
    en: "Scale"
  },
  preview: {
    ja: "プレビュー",
    en: "Preview"
  },
  cancel: {
    ja: "キャンセル",
    en: "Cancel"
  },
  ok: {
    ja: "OK",
    en: "OK"
  },
  joinMiter: {
    ja: "マイター",
    en: "Miter"
  },
  joinRound: {
    ja: "ラウンド",
    en: "Round"
  },
  joinBevel: {
    ja: "ベベル",
    en: "Bevel"
  }
};

function L(key) {
  try {
    if (LABELS[key] && LABELS[key][lang]) return LABELS[key][lang];
    if (LABELS[key] && LABELS[key].en) return LABELS[key].en;
  } catch (_) { }
  return key;
}

(function () {
    if (app.documents.length === 0) {
        alert(L('openDocument'));
        return;
    }

    var doc = app.activeDocument;
    var sel = doc.selection;

    function isSupportedItem(it) {
        return it && (
            it.typename === "PathItem" ||
            it.typename === "CompoundPathItem" ||
            it.typename === "GroupItem" ||
            it.typename === "TextFrame"
        );
    }

    function getSubPaths(it) {
        if (!it) return [];

        if (it.typename === "PathItem") return [it];

        if (it.typename === "CompoundPathItem") {
            var arr = [];
            try {
                for (var i = 0; i < it.pathItems.length; i++) arr.push(it.pathItems[i]);
            } catch (_) { }
            return arr;
        }

        // GroupItem: 中にある PathItem / CompoundPathItem を再帰的に集める
        if (it.typename === "GroupItem") {
            var out = [];

            function collectFrom(container) {
                if (!container) return;
                var items = null;
                try { items = container.pageItems; } catch (_) { items = null; }
                if (!items) return;

                for (var j = 0; j < items.length; j++) {
                    var child = items[j];
                    if (!child) continue;

                    if (child.typename === "PathItem") {
                        out.push(child);
                    } else if (child.typename === "CompoundPathItem") {
                        try {
                            for (var k = 0; k < child.pathItems.length; k++) out.push(child.pathItems[k]);
                        } catch (_) { }
                    } else if (child.typename === "GroupItem") {
                        collectFrom(child);
                    }
                }
            }

            collectFrom(it);
            return out;
        }

        return [];
    }

    function isAllClosed(subPaths) {
        if (!subPaths || subPaths.length === 0) return false;
        for (var i = 0; i < subPaths.length; i++) {
            try {
                if (!subPaths[i].closed) return false;
            } catch (_) {
                return false;
            }
        }
        return true;
    }

    if (sel.length !== 1 || !isSupportedItem(sel[0])) {
        alert(L('selectOneClosed'));
        return;
    }

    var originalPath = sel[0];
    var originalSubPaths = getSubPaths(originalPath);

    // Path/Compound はここで閉パス検証。Group/Text は実行時に一時パスへ変換して検証する。
    if (originalPath.typename !== "GroupItem" && originalPath.typename !== "TextFrame") {
        if (!isAllClosed(originalSubPaths)) {
            alert(L('selectClosed'));
            return;
        }
    }
    /* グループの一時合体 / Build temporary merged item from GroupItem */
    function executePathfinderAddAndExpand() {
        app.executeMenuCommand('Live Pathfinder Add');
        app.executeMenuCommand('expandStyle');
    }

    function executeNoCompoundAndAddExpand() {
        app.executeMenuCommand('noCompoundPath');
        app.executeMenuCommand('Live Pathfinder Add');
        app.executeMenuCommand('expandStyle');
    }

    // GroupItem を「一時的に」単一パス/複合パスへ変換して返す（元グループは残す）
    // 戻り値: { item: PageItem, cleanup: Function, ok: Boolean, message: String }
    function buildMergedItemFromGroup(groupItem) {
        var res = { item: null, cleanup: function () { }, ok: false, message: "" };

        if (!groupItem || groupItem.typename !== "GroupItem") {
            res.message = "GroupItem ではありません。";
            return res;
        }

        var tempDup = null;
        var merged = null;
        var trash = []; // 一時生成物を全部ここに入れて最後に削除する

        res.cleanup = function () {
            try { doc.selection = null; } catch (_) { }
            try {
                for (var ti = trash.length - 1; ti >= 0; ti--) {
                    try {
                        var it = trash[ti];
                        if (it && it.isValid) it.remove();
                    } catch (_) { }
                }
            } catch (_) { }

            // 念のため
            try { if (merged && merged.isValid) merged.remove(); } catch (_) { }
            try { if (tempDup && tempDup.isValid) tempDup.remove(); } catch (_) { }
        };

        try {
            tempDup = groupItem.duplicate();
            try { trash.push(tempDup); } catch (_) { }

            // duplicate を選択して PathFinder → Expand
            doc.selection = null;
            tempDup.selected = true;
            executePathfinderAddAndExpand();

            var items = doc.selection;
            doc.selection = null;
            try {
                if (items && items.length) {
                    for (var t0 = 0; t0 < items.length; t0++) trash.push(items[t0]);
                }
            } catch (_) { }

            if (!items || items.length === 0) {
                res.message = "グループの合体結果を取得できませんでした。";
                return res;
            }

            if (items.length === 1) {
                merged = items[0];
            } else {
                // 複数残った場合は一度グループ化して再合体
                var g = doc.groupItems.add();
                try { trash.push(g); } catch (_) { }
                for (var i = 0; i < items.length; i++) {
                    try { items[i].move(g, ElementPlacement.PLACEATEND); } catch (_) { }
                }
                doc.selection = null;
                g.selected = true;
                executePathfinderAddAndExpand();

                var items2 = doc.selection;
                doc.selection = null;
                try {
                    if (items2 && items2.length) {
                        for (var t1 = 0; t1 < items2.length; t1++) trash.push(items2[t1]);
                    }
                } catch (_) { }

                if (items2 && items2.length === 1) {
                    merged = items2[0];
                } else {
                    merged = (items2 && items2.length > 0) ? items2[0] : items[0];
                }
            }

            try { if (merged) trash.push(merged); } catch (_) { }
            res.item = merged;
            res.ok = true;
            return res;

        } catch (e) {
            res.message = "グループの合体中にエラーが発生しました: " + e;
            return res;
        }
    }
    // TextFrame を「一時的に」アウトライン化→合体して単一パス/複合パスへ変換して返す（元テキストは残す）
    // 戻り値: { item: PageItem, cleanup: Function, ok: Boolean, message: String }
    function buildMergedItemFromText(textFrame) {
        var res = { item: null, cleanup: function () { }, ok: false, message: "" };

        if (!textFrame || textFrame.typename !== "TextFrame") {
            res.message = "TextFrame ではありません。";
            return res;
        }

        var tempDup = null;
        var merged = null;
        var trash = [];

        res.cleanup = function () {
            try { doc.selection = null; } catch (_) { }
            try {
                for (var ti = trash.length - 1; ti >= 0; ti--) {
                    try {
                        var it = trash[ti];
                        if (it && it.isValid) it.remove();
                    } catch (_) { }
                }
            } catch (_) { }
            try { if (merged && merged.isValid) merged.remove(); } catch (_) { }
            try { if (tempDup && tempDup.isValid) tempDup.remove(); } catch (_) { }
        };

        try {
            // ① テキストを複製
            tempDup = textFrame.duplicate();
            try { trash.push(tempDup); } catch (_) { }

            // ② アウトライン化（v15方式：TextFrame#createOutline を使用）
            // createOutline() は元の TextFrame を削除して GroupItem/CompoundPathItem を返す
            // ここでは複製に対して実行するので、オリジナルは残る
            var outlined = null;
            try {
                if (tempDup && typeof tempDup.createOutline === 'function') {
                    outlined = tempDup.createOutline();
                }
            } catch (eOutline) {
                outlined = null;
            }

            // createOutline() 成功時は tempDup が消えるので trash から外す必要はない（cleanup側でisValidを見て消す）

            if (!outlined) {
                res.message = "アウトライン化に失敗しました。";
                return res;
            }

            // outlined は GroupItem / CompoundPathItem / PathItem のいずれか
            // 後で確実に消すため、子要素も含めて trash に登録しておく
            function pushAllDescendants(it) {
                if (!it) return;
                try { trash.push(it); } catch (_) { }
                try {
                    if (it.typename === 'GroupItem') {
                        for (var i = 0; i < it.pageItems.length; i++) {
                            pushAllDescendants(it.pageItems[i]);
                        }
                    } else if (it.typename === 'CompoundPathItem') {
                        // compound 自体も + 内部 pathItems も
                        for (var j = 0; j < it.pathItems.length; j++) {
                            try { trash.push(it.pathItems[j]); } catch (_) { }
                        }
                    }
                } catch (_) { }
            }

            pushAllDescendants(outlined);

            // 選択状態が残ると後段の処理に巻き込まれるので解除
            try { doc.selection = null; } catch (_) { }
            try { outlined.selected = false; } catch (_) { }

            merged = outlined;
            res.item = merged;
            res.ok = true;
            return res;

        } catch (e) {
            res.message = "テキストの合体中にエラーが発生しました: " + e;
            return res;
        }
    }

    var isPreviewing = false;
    var __applyOffsetNow = true; // プレビュー時は false にしてオフセットを無視

    /* ダイアログ作成 / Build dialog */
    var dlg = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);
    dlg.orientation = "column";
    dlg.alignChildren = "fill";

    // プリセット（スケール＋角度） ※距離は含めない（ダイアログ最上部）
    // 外側グループで左右中央に配置し、内側グループにラベル＋プルダウンをまとめる
    var grpPresetWrap = dlg.add("group");
    grpPresetWrap.orientation = "row";
    grpPresetWrap.alignChildren = ["center", "center"];
    grpPresetWrap.alignment = "center";

    var grpPreset = grpPresetWrap.add("group");
    grpPreset.orientation = "row";
    grpPreset.alignChildren = ["left", "center"];

    grpPreset.add("statictext", undefined, L('preset'));

    var ddPreset = grpPreset.add("dropdownlist", undefined, [
        "100% /  45°",
        "100% /  30°",
        "100% /  60°",
        " 50% /  90°",
        "  1% /  90°",
        "100% / 135°",
        "100% / 120°",
        "100% / 150°"
    ]);
    ddPreset.selection = 0; // デフォルト：45° / 100%

    // --- 1カラム（プリセットは貫通） ---
    var cols = dlg.add("group");
    cols.orientation = "column";
    cols.alignChildren = ["fill", "top"];
    cols.alignment = "fill";

    // 左：現状（設定）
    var mainGroup = cols.add("panel", undefined, L('settings'));
    mainGroup.orientation = "column";
    mainGroup.alignChildren = "left";
    mainGroup.margins = [15, 20, 15, 10];

    // 右：オフセット（中身は後で実装）
    var offsetPanel = cols.add("panel", undefined, L('offset'));
    offsetPanel.orientation = "column";
    offsetPanel.alignChildren = "left";
    offsetPanel.margins = [15, 20, 15, 10];

    // --- オフセットpanel内 2カラム ---
    var offCols = offsetPanel.add("group");
    offCols.orientation = "row";
    offCols.alignChildren = ["fill", "top"];
    // offCols.alignment = "fill";

    // 左：□［　］pt
    var offLeft = offCols.add("group");
    offLeft.orientation = "column";
    offLeft.alignChildren = ["left", "top"];

    // オフセット（チェック＋値を1行に）
    var gOff = offLeft.add("group");
    gOff.orientation = "row";
    gOff.alignChildren = ["left", "center"];

    var cbOffsetRow = gOff.add("checkbox", undefined, "");
    cbOffsetRow.value = false;

    // オフセット初期値を選択オブジェクト（originalPath）のサイズから計算（pt）
    var initialOffset = 0;
    try {
        var bOff = null;
        try { bOff = originalPath.geometricBounds; } catch (_) { bOff = null; }
        if (!bOff) {
            try { bOff = originalPath.visibleBounds; } catch (_) { bOff = null; }
        }
        if (bOff && bOff.length === 4) {
            var wPtOff = Math.abs(bOff[2] - bOff[0]);
            var hPtOff = Math.abs(bOff[1] - bOff[3]);

            // 幅と高さの合計（pt）
            var sizeSumPt = wPtOff + hPtOff;
            // 平均サイズ（pt）
            var halfSizePt = sizeSumPt / 2;

            // PathItem 想定：divisor は 20
            var divisor = 20;

            // オフセット基準値（pt）
            var offsetBasePt = halfSizePt / divisor;

            if (!isNaN(offsetBasePt)) {
                initialOffset = Math.round(offsetBasePt);
            }
        }
    } catch (_) { }

    var etOff = gOff.add("edittext", undefined, String(initialOffset));
    etOff.characters = 4;
    changeValueByArrowKey(etOff, false, false);
    var stOffUnit = gOff.add("statictext", undefined, "pt");

    // 右：形状panel
    var offRight = offCols.add("group");
    offRight.orientation = "column";
    offRight.alignChildren = ["fill", "top"];

    // 形状（角の処理） - オフセットpanel内（2カラム：左=見出し / 右=ラジオ）
    var pJoin = offRight.add("group");
    pJoin.orientation = "row";
    pJoin.alignChildren = ["left", "top"];
    pJoin.margins = [0, 0, 0, 0];

    // 左：見出し
    var gJoinLabel = pJoin.add("group");
    gJoinLabel.orientation = "column";
    gJoinLabel.alignChildren = ["left", "top"];
    var stJoin = gJoinLabel.add("statictext", undefined, L('shape'));
    stJoin.preferredSize.width = 60;
    stJoin.justify = "right";

    // 右：ラジオ（垂直並び）
    var gJoinCol = pJoin.add("group");
    gJoinCol.orientation = "column";
    gJoinCol.alignChildren = ["left", "center"];

    var rbMiter = gJoinCol.add("radiobutton", undefined, L('joinMiter'));
    var rbRound = gJoinCol.add("radiobutton", undefined, L('joinRound'));
    var rbBevel = gJoinCol.add("radiobutton", undefined, L('joinBevel'));
    rbMiter.value = false;
    rbRound.value = true;  // default = Round
    rbBevel.value = false;

    // オフセットUIの有効/無効（ディム表示）
    function updateOffsetEnabled() {
        var on = !!cbOffsetRow.value;
        try { etOff.enabled = on; } catch (_) { }
        try { rbMiter.enabled = on; } catch (_) { }
        try { rbRound.enabled = on; } catch (_) { }
        try { rbBevel.enabled = on; } catch (_) { }
        try { stOffUnit.enabled = on; } catch (_) { }
        try { pJoin.enabled = on; } catch (_) { }
    }

    cbOffsetRow.onClick = function () {
        updateOffsetEnabled();
    };

    // 初期状態：OFF
    updateOffsetEnabled();

    // 距離入力
    var grpDistance = mainGroup.add("group");
    var lblDistance = grpDistance.add("statictext", undefined, L('distance'));
    lblDistance.preferredSize.width = 60;
    lblDistance.justify = "right";

    // オリジナルの幅＋高さを足した値をデフォルト距離にする（pt）
    var gb = originalPath.geometricBounds; // [left, top, right, bottom]
    var w = gb[2] - gb[0];
    var h = gb[1] - gb[3];
    var defaultDistance = (w + h);

    var inputDistance = grpDistance.add("edittext", undefined, Math.round(defaultDistance).toString());
    changeValueByArrowKey(inputDistance, false, true);
    inputDistance.characters = 4;
    grpDistance.add("statictext", undefined, "pt");

    // スライダー（距離）※右側に配置
    var maxDist = Math.max(500, Math.round(defaultDistance * 3));
    var slDistance = grpDistance.add("slider", undefined, Math.round(defaultDistance), 0, maxDist);
    slDistance.preferredSize.width = 170;

    // 角度入力
    var grpAngle = mainGroup.add("group");
    var lblAngle = grpAngle.add("statictext", undefined, L('angle'));
    lblAngle.preferredSize.width = 60;
    lblAngle.justify = "right";
    var inputAngle = grpAngle.add("edittext", undefined, "45");
    changeValueByArrowKey(inputAngle, true, true);
    inputAngle.characters = 4;
    grpAngle.add("statictext", undefined, "°");

    // スライダー（角度）※右側に配置
    var slAngle = grpAngle.add("slider", undefined, 45, -180, 180);
    slAngle.preferredSize.width = 170;

    // スケール入力
    var grpScale = mainGroup.add("group");
    var lblScale = grpScale.add("statictext", undefined, L('scale'));
    lblScale.preferredSize.width = 60;
    lblScale.justify = "right";
    var inputScale = grpScale.add("edittext", undefined, "100"); // デフォルト 100%
    changeValueByArrowKey(inputScale, false, true);
    inputScale.characters = 4;
    grpScale.add("statictext", undefined, "%");

    // スライダー（スケール）※右側に配置
    var slScale = grpScale.add("slider", undefined, 100, 1, 300);
    slScale.preferredSize.width = 170;

    // スライダー連動（edittext <-> slider）
    bindSliderNumber(inputDistance, slDistance, 0, maxDist, false);
    bindSliderNumber(inputAngle, slAngle, -180, 180, true);
    bindSliderNumber(inputScale, slScale, 1, 300, false);

    // プリセット選択時：スケールと角度に反映（距離は変更しない）
    ddPreset.onChange = function () {
        if (!ddPreset.selection) return;

        // 例: "100% / 45°"
        var s = ddPreset.selection.text;
        var m = s.match(/^\s*([\-]?\d+(?:\.\d+)?)\s*%\s*\/\s*([\-]?\d+(?:\.\d+)?)\s*°\s*$/);
        if (!m) return;

        var scl = m[1];
        var ang = m[2];

        inputAngle.text = ang;
        inputScale.text = scl;

        if (chkPreview && chkPreview.value) updatePreview();
    };

    // Buttons（左：プレビュー（無効） / 右：キャンセル・OK）
    var gBtn = dlg.add("group");
    gBtn.orientation = "row";
    gBtn.alignment = ["fill", "top"];
    gBtn.alignChildren = ["right", "center"];

    // 左端：プレビュー
    var chkPreview = gBtn.add("checkbox", undefined, L('preview'));
    chkPreview.value = true;   // デフォルトON
    chkPreview.enabled = true; // 有効

    // スペーサー（左右を分離）
    var spacer = gBtn.add("group");
    spacer.alignment = ["fill", "fill"];

    // 右端：キャンセル / OK
    var btnCancel = gBtn.add("button", undefined, L('cancel'), { name: "cancel" });
    var btnOk = gBtn.add("button", undefined, L('ok'), { name: "ok" });
    btnOk.active = true;

    // --- ↑↓キーで数値を増減（Shift=±10, Option=±0.1） ---
    function changeValueByArrowKey(editText, allowNegative, enablePreviewUpdate) {
        editText.addEventListener("keydown", function (event) {
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;

            if (keyboard.shiftKey) {
                delta = 10;
                // Shiftキー押下時は10の倍数にスナップ
                if (event.keyName == "Up") {
                    value = Math.ceil((value + 1) / delta) * delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value = Math.floor((value - 1) / delta) * delta;
                    event.preventDefault();
                }
            } else if (keyboard.altKey) {
                delta = 0.1;
                // Optionキー押下時は0.1単位で増減
                if (event.keyName == "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value -= delta;
                    event.preventDefault();
                }
            } else {
                delta = 1;
                if (event.keyName == "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value -= delta;
                    event.preventDefault();
                }
            }

            if (keyboard.altKey) {
                // 小数第1位までに丸め
                value = Math.round(value * 10) / 10;
            } else {
                // 整数に丸め
                value = Math.round(value);
            }

            if (!allowNegative && value < 0) value = 0;

            // 反映
            editText.text = value;

            // プレビュー更新（必要なフィールドのみ／プレビューONの時のみ）
            if (enablePreviewUpdate !== false && typeof updatePreview === "function" && chkPreview && chkPreview.value) {
                updatePreview();
            }
        });
    }

    // --- スライダー連動（距離/スケール/角度） ---
    function clamp(v, min, max) {
        return Math.max(min, Math.min(max, v));
    }

    // slider は整数のみなので、必要に応じて丸める
    function bindSliderNumber(editText, slider, min, max, isAngle) {
        var syncing = false;

        function setEditFromSlider() {
            if (syncing) return;
            syncing = true;
            try {
                var v = slider.value;
                if (!isAngle) {
                    // 距離/スケールは整数
                    v = Math.round(v);
                } else {
                    // 角度も整数扱い
                    v = Math.round(v);
                }
                editText.text = String(v);
            } catch (_) { }
            syncing = false;

            // プレビュー更新（ONの時のみ）
            if (typeof updatePreview === "function" && chkPreview && chkPreview.value) {
                updatePreview();
            }
        }

        function setSliderFromEdit() {
            if (syncing) return;
            syncing = true;
            try {
                var v = Number(editText.text);
                if (isNaN(v)) v = 0;
                v = clamp(v, min, max);
                slider.value = Math.round(v);
            } catch (_) { }
            syncing = false;
        }

        // 初期同期
        setSliderFromEdit();

        slider.onChanging = setEditFromSlider;
        slider.onChange = setEditFromSlider; // マウス離したときも反映

        // editText側の変化に追随（既存のonChangingは残しつつ、同期だけ追加）
        var prevOnChanging = editText.onChanging;
        editText.onChanging = function () {
            try { setSliderFromEdit(); } catch (_) { }
            if (typeof prevOnChanging === "function") prevOnChanging();
        };
    }

    /* 色処理 / Color utilities */
    // A（オリジナル）の塗り色を元に、彩度を下げた色を作る
    // factor: 0〜1（1=元の彩度、0=完全に無彩色）
    function desaturateColorFromFill(fillColor, factor) {
        if (!fillColor) return null;
        if (factor === undefined || factor === null) factor = 0.7;

        // Grayはそのまま
        if (fillColor.typename === "GrayColor") {
            var g = new GrayColor();
            g.gray = fillColor.gray;
            return g;
        }

        // CMYKは一旦RGBに近似変換して処理（戻しはRGBで適用）
        var rgb = null;
        if (fillColor.typename === "RGBColor") {
            rgb = { r: fillColor.red, g: fillColor.green, b: fillColor.blue };
        } else if (fillColor.typename === "CMYKColor") {
            // 0-100 -> 0-1
            var c = fillColor.cyan / 100.0;
            var m = fillColor.magenta / 100.0;
            var y = fillColor.yellow / 100.0;
            var k = fillColor.black / 100.0;
            // 簡易 CMYK -> RGB
            var rr = 255 * (1 - c) * (1 - k);
            var gg = 255 * (1 - m) * (1 - k);
            var bb = 255 * (1 - y) * (1 - k);
            rgb = { r: rr, g: gg, b: bb };
        } else {
            // 対応外はそのまま返す（NoColor等）
            return fillColor;
        }

        // RGB -> HSL で彩度調整（HSLのSをfactor倍）
        function clamp01(v) { return Math.max(0, Math.min(1, v)); }
        var r1 = clamp01(rgb.r / 255.0);
        var g1 = clamp01(rgb.g / 255.0);
        var b1 = clamp01(rgb.b / 255.0);

        var max = Math.max(r1, g1, b1);
        var min = Math.min(r1, g1, b1);
        var h = 0, s = 0, l = (max + min) / 2;
        var d = max - min;

        if (d !== 0) {
            s = d / (1 - Math.abs(2 * l - 1));
            switch (max) {
                case r1: h = ((g1 - b1) / d) % 6; break;
                case g1: h = ((b1 - r1) / d) + 2; break;
                case b1: h = ((r1 - g1) / d) + 4; break;
            }
            h = h * 60;
            if (h < 0) h += 360;
        }

        s = clamp01(s * factor);

        // HSL -> RGB
        function hue2rgb(p, q, t) {
            if (t < 0) t += 1;
            if (t > 1) t -= 1;
            if (t < 1 / 6) return p + (q - p) * 6 * t;
            if (t < 1 / 2) return q;
            if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
            return p;
        }

        var r2, g2, b2;
        if (s === 0) {
            r2 = g2 = b2 = l;
        } else {
            var q = l < 0.5 ? l * (1 + s) : (l + s - l * s);
            var p = 2 * l - q;
            var hk = h / 360;
            r2 = hue2rgb(p, q, hk + 1 / 3);
            g2 = hue2rgb(p, q, hk);
            b2 = hue2rgb(p, q, hk - 1 / 3);
        }

        var out = new RGBColor();
        out.red = Math.round(r2 * 255);
        out.green = Math.round(g2 * 255);
        out.blue = Math.round(b2 * 255);
        return out;
    }

    // PathItem / CompoundPathItem / GroupItem に塗り色を適用
    function applyFillColorToItem(item, fillColor) {
        if (!item || !fillColor) return;

        try {
            if (item.typename === "PathItem") {
                item.filled = true;
                item.fillColor = fillColor;
                return;
            }
            if (item.typename === "CompoundPathItem") {
                for (var i = 0; i < item.pathItems.length; i++) {
                    try {
                        item.pathItems[i].filled = true;
                        item.pathItems[i].fillColor = fillColor;
                    } catch (_) { }
                }
                return;
            }
            if (item.typename === "GroupItem") {
                for (var j = 0; j < item.pageItems.length; j++) {
                    applyFillColorToItem(item.pageItems[j], fillColor);
                }
                return;
            }
        } catch (_) { }
    }

    /* 形状サンプリング & 凸包 / Sampling & convex hull (visual approximation) */
    function isCurvedPath(item) {
        var subs = getSubPaths(item);
        if (!subs || subs.length === 0) return false;
        try {
            for (var si = 0; si < subs.length; si++) {
                var pts = subs[si].pathPoints;
                for (var i = 0; i < pts.length; i++) {
                    var p = pts[i];
                    // ハンドルがアンカーと一致しない＝曲線セグメントがある可能性が高い
                    if (p.leftDirection[0] !== p.anchor[0] || p.leftDirection[1] !== p.anchor[1]) return true;
                    if (p.rightDirection[0] !== p.anchor[0] || p.rightDirection[1] !== p.anchor[1]) return true;
                }
            }
        } catch (_) { }
        return false;
    }

    function bezierPoint(p0, p1, p2, p3, t) {
        var u = 1 - t;
        var tt = t * t;
        var uu = u * u;
        var uuu = uu * u;
        var ttt = tt * t;
        return [
            uuu * p0[0] + 3 * uu * t * p1[0] + 3 * u * tt * p2[0] + ttt * p3[0],
            uuu * p0[1] + 3 * uu * t * p1[1] + 3 * u * tt * p2[1] + ttt * p3[1]
        ];
    }

    // PathItem をサンプリングして点群を返す（閉パス前提）
    function sampleClosedPath(pathItem, samplesPerSeg) {
        if (!samplesPerSeg) samplesPerSeg = 16;
        var pts = [];
        var pp = pathItem.pathPoints;
        var n = pp.length;

        for (var i = 0; i < n; i++) {
            var a = pp[i];
            var b = pp[(i + 1) % n];

            var p0 = a.anchor;
            var p1 = a.rightDirection;
            var p2 = b.leftDirection;
            var p3 = b.anchor;

            // t=0 は次セグメントと重複しやすいので 0 は入れるが、1 は入れない（最後に閉じる）
            for (var s = 0; s < samplesPerSeg; s++) {
                var t = s / samplesPerSeg;
                // 直線の場合もこれでOK（p1/p2 が p0/p3 と一致する）
                pts.push(bezierPoint(p0, p1, p2, p3, t));
            }
        }
        return pts;
    }

    // PathItem または CompoundPathItem をサンプリングして点群を返す
    function sampleClosedItem(item, samplesPerSeg) {
        var subs = getSubPaths(item);
        var pts = [];
        for (var i = 0; i < subs.length; i++) {
            try {
                var p = sampleClosedPath(subs[i], samplesPerSeg);
                for (var k = 0; k < p.length; k++) pts.push(p[k]);
            } catch (_) { }
        }
        return pts;
    }

    /* 実パス生成（サンプリング面生成） / Build shadow by sampling & bridging */
    var CURVE_SAMPLE_COUNT = 8; // v15 相当（必要なら調整）

    function duplicateMoveScale(item, dx, dy, scalePercent) {
        if (!item) return null;
        var dup = null;
        try { dup = item.duplicate(); } catch (_) { dup = null; }
        if (!dup) return null;

        // scale: percent (100 = 等倍)
        try {
            var sp = (isFinite(scalePercent) ? Number(scalePercent) : 100);
            if (sp !== 100 && dup && typeof dup.resize === 'function') {
                try { dup.resize(sp, sp, true, true, true, true, true, Transformation.CENTER); }
                catch (eR1) {
                    try { dup.resize(sp, sp, true, true, true, true, true); } catch (_) { }
                }
            }
        } catch (_) { }

        try { dup.translate(dx, dy); } catch (_) { }
        return dup;
    }

    // PathItem を「曲線も含めて」点列にする（閉パス前提）
    function samplePathToPoints(pathItem, curveSamplesPerSeg) {
        if (!curveSamplesPerSeg) curveSamplesPerSeg = CURVE_SAMPLE_COUNT;
        var pts = [];
        if (!pathItem) return pts;

        try {
            var pp = pathItem.pathPoints;
            var n = pp.length;
            if (!pp || n < 2) return pts;

            for (var i = 0; i < n; i++) {
                var a = pp[i];
                var b = pp[(i + 1) % n];

                var p0 = a.anchor;
                var p1 = a.rightDirection;
                var p2 = b.leftDirection;
                var p3 = b.anchor;

                for (var s = 0; s < curveSamplesPerSeg; s++) {
                    var t = s / curveSamplesPerSeg;
                    pts.push(bezierPoint(p0, p1, p2, p3, t));
                }
            }
        } catch (_) { }

        return pts;
    }

    // PathItem/Compound/Group を「対応順序を保って」PathItem 配列として抽出（閉パスのみ）
    function collectClosedPathsForBridge(item) {
        var out = [];
        if (!item) return out;

        function pushIfClosed(p) {
            try { if (p && p.typename === 'PathItem' && p.closed) out.push(p); } catch (_) { }
        }

        try {
            if (item.typename === 'PathItem') {
                pushIfClosed(item);
                return out;
            }
            if (item.typename === 'CompoundPathItem') {
                for (var i = 0; i < item.pathItems.length; i++) pushIfClosed(item.pathItems[i]);
                return out;
            }
            if (item.typename === 'GroupItem') {
                // pageItems 順で再帰（見た目の重なり順に近い）
                for (var j = 0; j < item.pageItems.length; j++) {
                    var ch = item.pageItems[j];
                    if (!ch) continue;
                    if (ch.typename === 'PathItem') {
                        pushIfClosed(ch);
                    } else if (ch.typename === 'CompoundPathItem') {
                        for (var k = 0; k < ch.pathItems.length; k++) pushIfClosed(ch.pathItems[k]);
                    } else if (ch.typename === 'GroupItem') {
                        var sub = collectClosedPathsForBridge(ch);
                        for (var m = 0; m < sub.length; m++) out.push(sub[m]);
                    }
                }
                return out;
            }
        } catch (_) { }

        return out;
    }

    // 点列同士を結ぶ側面（四角形）を作る
    function buildSideFacesFromPointLists(parentGroup, ptsA, ptsB, baseFill) {
        var created = [];
        if (!parentGroup || !ptsA || !ptsB) return created;
        var n = Math.min(ptsA.length, ptsB.length);
        if (n < 2) return created;

        for (var i = 0; i < n; i++) {
            var next = (i + 1) % n;
            var p = null;
            try {
                p = parentGroup.pathItems.add();
                p.setEntirePath([ptsA[i], ptsB[i], ptsB[next], ptsA[next]]);
                p.closed = true;
                p.stroked = false;
                p.filled = true;
                if (baseFill) p.fillColor = baseFill;
                created.push(p);
            } catch (_) {
                try { if (p && p.isValid) p.remove(); } catch (_) { }
            }
        }
        return created;
    }

    // v15系：サンプリングで面を作る実生成
    // 戻り値: 生成した影グループ（GroupItem）
    function generateShadowBySampling(baseItem, dx, dy, scalePercent) {
        if (!baseItem) return null;

        // ベースの閉パス群を抽出
        var basePaths = collectClosedPathsForBridge(baseItem);
        if (!basePaths || basePaths.length === 0) return null;

        // 複製して移動/スケール（複製側）
        var dup = duplicateMoveScale(baseItem, dx, dy, scalePercent);
        if (!dup) return null;

        // 複製側の閉パス群を抽出
        var dupPaths = collectClosedPathsForBridge(dup);
        if (!dupPaths || dupPaths.length === 0) {
            try { dup.remove(); } catch (_) { }
            return null;
        }

        // ペアリング：数が違う場合は最小数で処理（落とさない優先）
        var pairCount = Math.min(basePaths.length, dupPaths.length);
        if (pairCount === 0) {
            try { dup.remove(); } catch (_) { }
            return null;
        }

        // 影用グループ
        var g = doc.groupItems.add();

        // ベースの塗りを基準に彩度を落とす
        var baseFill = null;
        try { baseFill = desaturateColorFromFill(basePaths[0].fillColor, 0.7); } catch (_) { baseFill = null; }

        for (var i = 0; i < pairCount; i++) {
            var a = basePaths[i];
            var b = dupPaths[i];

            var ptsA = samplePathToPoints(a, CURVE_SAMPLE_COUNT);
            var ptsB = samplePathToPoints(b, CURVE_SAMPLE_COUNT);

            // 点数がずれた場合に備えて、短い方に合わせる
            if (ptsA.length < 2 || ptsB.length < 2) continue;

            buildSideFacesFromPointLists(g, ptsA, ptsB, baseFill);
        }

        // 複製側（dup）は最終結果には不要
        try { dup.remove(); } catch (_) { }

        return g;
    }

    // 凸包（Monotone chain）
    function convexHull(points) {
        if (!points || points.length < 3) return points;

        // 重複をある程度削る（丸めてキー化）
        var map = {};
        var uniq = [];
        for (var i = 0; i < points.length; i++) {
            var x = Math.round(points[i][0] * 1000) / 1000;
            var y = Math.round(points[i][1] * 1000) / 1000;
            var key = x + "," + y;
            if (!map[key]) {
                map[key] = true;
                uniq.push([x, y]);
            }
        }
        if (uniq.length < 3) return uniq;

        uniq.sort(function (a, b) {
            if (a[0] === b[0]) return a[1] - b[1];
            return a[0] - b[0];
        });

        function cross(o, a, b) {
            return (a[0] - o[0]) * (b[1] - o[1]) - (a[1] - o[1]) * (b[0] - o[0]);
        }

        var lower = [];
        for (var i = 0; i < uniq.length; i++) {
            while (lower.length >= 2 && cross(lower[lower.length - 2], lower[lower.length - 1], uniq[i]) <= 0) {
                lower.pop();
            }
            lower.push(uniq[i]);
        }

        var upper = [];
        for (var j = uniq.length - 1; j >= 0; j--) {
            while (upper.length >= 2 && cross(upper[upper.length - 2], upper[upper.length - 1], uniq[j]) <= 0) {
                upper.pop();
            }
            upper.push(uniq[j]);
        }

        upper.pop();
        lower.pop();
        return lower.concat(upper);
    }

    // 点群から C（凸包外形）を作成して返す
    function createHullPath(doc, hullPoints) {
        var p = doc.pathItems.add();
        p.stroked = false;
        p.filled = true;
        p.closed = true;
        p.setEntirePath(hullPoints);
        return p;
    }

    /* 曲線っぽい角をスムーズ化 / Smooth polyline segments that approximate curves */
    // handleFactor: 0〜0.5（小さいほど控えめ）
    // angleThresholdDeg: この角度以上（180に近い）ならスムーズ候補。例: 150
    function smoothPolylineAsCurve(pathItem, handleFactor, angleThresholdDeg) {
        if (!pathItem) return;
        if (handleFactor === undefined || handleFactor === null) handleFactor = 0.25;
        if (angleThresholdDeg === undefined || angleThresholdDeg === null) angleThresholdDeg = 150;

        try {
            if (!pathItem.closed) return;
        } catch (_) { return; }

        try {
            var pp = pathItem.pathPoints;
            var n = pp.length;
            if (!pp || n < 3) return;

            function len(v) { return Math.sqrt(v[0] * v[0] + v[1] * v[1]); }
            function norm(v) {
                var l = len(v);
                if (l === 0) return [0, 0];
                return [v[0] / l, v[1] / l];
            }
            function sub(a, b) { return [a[0] - b[0], a[1] - b[1]]; }
            function add(a, b) { return [a[0] + b[0], a[1] + b[1]]; }
            function mul(a, s) { return [a[0] * s, a[1] * s]; }
            function dot(a, b) { return a[0] * b[0] + a[1] * b[1]; }

            // 角度計算（Bでの角）
            function angleDeg(A, B, C) {
                var v1 = sub(A, B); // BA
                var v2 = sub(C, B); // BC
                var l1 = len(v1);
                var l2 = len(v2);
                if (l1 < 1e-6 || l2 < 1e-6) return 0;
                var c = dot(v1, v2) / (l1 * l2);
                if (c > 1) c = 1;
                if (c < -1) c = -1;
                return Math.acos(c) * 180 / Math.PI;
            }

            for (var i = 0; i < n; i++) {
                var prev = pp[(i - 1 + n) % n];
                var cur = pp[i];
                var next = pp[(i + 1) % n];

                var A = prev.anchor;
                var B = cur.anchor;
                var C = next.anchor;

                var ang = angleDeg(A, B, C);

                // 角度が大きい（180°に近い）＝曲線近似の折れ線とみなしてスムーズ化
                if (ang >= angleThresholdDeg) {
                    var v1 = sub(B, A);
                    var v2 = sub(C, B);
                    var l1 = len(v1);
                    var l2 = len(v2);

                    if (l1 < 1e-3 || l2 < 1e-3) {
                        try {
                            cur.pointType = PointType.CORNER;
                            cur.leftDirection = B;
                            cur.rightDirection = B;
                        } catch (_) { }
                        continue;
                    }

                    // 接線方向：入射方向と出射方向の単位ベクトルを合成
                    var t = add(norm(v1), norm(v2));
                    var tl = len(t);
                    if (tl < 1e-6) {
                        t = norm(v2);
                    } else {
                        t = norm(t);
                    }

                    var hLen = Math.min(l1, l2) * handleFactor;
                    var left = sub(B, mul(t, hLen));
                    var right = add(B, mul(t, hLen));

                    try {
                        cur.pointType = PointType.SMOOTH;
                        cur.leftDirection = left;
                        cur.rightDirection = right;
                    } catch (_) { }
                } else {
                    // 鋭角はコーナーを維持
                    try {
                        cur.pointType = PointType.CORNER;
                        cur.leftDirection = B;
                        cur.rightDirection = B;
                    } catch (_) { }
                }
            }
        } catch (_) { }
    }

    function smoothPolylineInItem(item, handleFactor, angleThresholdDeg) {
        if (!item) return;
        try {
            if (item.typename === 'PathItem') {
                smoothPolylineAsCurve(item, handleFactor, angleThresholdDeg);
                return;
            }
            if (item.typename === 'CompoundPathItem') {
                for (var i = 0; i < item.pathItems.length; i++) {
                    smoothPolylineAsCurve(item.pathItems[i], handleFactor, angleThresholdDeg);
                }
                return;
            }
            if (item.typename === 'GroupItem') {
                for (var j = 0; j < item.pageItems.length; j++) {
                    smoothPolylineInItem(item.pageItems[j], handleFactor, angleThresholdDeg);
                }
                return;
            }
        } catch (_) { }
    }

    /* オフセット（ライブ効果） / Offset Path live effect */
    // joinTypes: 0 = Round, 1 = Bevel, 2 = Miter
    function getJoinCode() {
        try {
            if (rbRound && rbRound.value) return 0;
            if (rbBevel && rbBevel.value) return 1;
        } catch (_) { }
        return 2; // マイター
    }

    function buildOffsetEffectXML(offsetPt, joinCode) {
        // mlim はマイター制限（既定 4）、ofst は pt
        return '<LiveEffect name="Adobe Offset Path"><Dict data="R mlim 4 R ofst ' + offsetPt + ' I jntp ' + joinCode + ' "/></LiveEffect>';
    }

    function groupItems(doc, items) {
        var g = doc.groupItems.add();
        // 受け取った順序で入れる
        for (var i = 0; i < items.length; i++) {
            try {
                items[i].move(g, ElementPlacement.PLACEATEND);
            } catch (_) { }
        }
        return g;
    }

    function applyOffsetEffectToCGroup(cGroup) {
        if (!cGroup) return;
        if (!__applyOffsetNow) return;
        if (!cbOffsetRow || !cbOffsetRow.value) return;

        var off = Number(etOff.text);
        if (isNaN(off)) off = 0;

        // 0でもONなら適用する（結果が変わらないだけ）
        var joinCode = getJoinCode();
        var xml = buildOffsetEffectXML(off, joinCode);

        try {
            cGroup.applyEffect(xml);
        } catch (_) { }
    }

    /* 実行関数 / Main execution */
    function doExtrude(applyOffset) {
        __applyOffsetNow = (applyOffset !== false);
        var dist = parseFloat(inputDistance.text) || 0;
        var scale = parseFloat(inputScale.text);
        if (isNaN(scale) || scale <= 0) scale = 100; // デフォルト 100%
        var ang = -(parseFloat(inputAngle.text) || 0); // 入力角度を反転

        var rad = ang * Math.PI / 180;
        var dx = dist * Math.cos(rad);
        var dy = dist * Math.sin(rad);

        // 実行対象（ベース形状）の確定
        // GroupItem の場合は「複製→合体→expand」した一時パス/複合パスを使う（元は残す）
        var baseItem = originalPath;
        var baseCleanup = function () { };
        var baseSubPaths = originalSubPaths;
        var baseIsTemp = false; // Group/Text 由来の一時ベースなら true

        if (originalPath.typename === "GroupItem") {
            var built = buildMergedItemFromGroup(originalPath);
            if (!built.ok || !built.item) {
                alert(built.message || L('cannotBuildFromGroup'));
                try { built.cleanup(); } catch (_) { }
                return;
            }
            baseItem = built.item;
            baseCleanup = built.cleanup;
            baseIsTemp = true;
            baseSubPaths = getSubPaths(baseItem);
            if (!isAllClosed(baseSubPaths)) {
                try { baseCleanup(); } catch (_) { }
                alert(L('selectClosedGroup'));
                return;
            }
        }

        if (originalPath.typename === "TextFrame") {
            var builtT = buildMergedItemFromText(originalPath);
            if (!builtT.ok || !builtT.item) {
                alert(builtT.message || L('selectClosed'));
                try { builtT.cleanup(); } catch (_) { }
                return;
            }
            baseItem = builtT.item;
            baseCleanup = builtT.cleanup;
            baseIsTemp = true;
            baseSubPaths = getSubPaths(baseItem);
            if (!isAllClosed(baseSubPaths)) {
                try { baseCleanup(); } catch (_) { }
                alert(L('selectClosed'));
                return;
            }
        }

        try {

        // --- プレビュー（簡易） ---
        if (!__applyOffsetNow) {
            // TextFrame はアウトライン化せず、テキストそのものを複製して表示
            if (originalPath.typename === "TextFrame") {
                function makeTextPreviewDup(f, op) {
                    var d = null;
                    try { d = originalPath.duplicate(); } catch (_) { d = null; }
                    if (!d) return null;

                    var s = 100 + (scale - 100) * f;
                    try { d.resize(s, s, true, true, true, true, true, Transformation.CENTER); } catch (_) { }
                    try { d.translate(dx * f, dy * f); } catch (_) { }
                    try { d.move(originalPath, ElementPlacement.PLACEBEFORE); } catch (_) { }
                    try { d.opacity = op; } catch (_) { }
                    return d;
                }

                makeTextPreviewDup(1, 40);
                makeTextPreviewDup(0.75, 60);
                makeTextPreviewDup(0.5, 80);
                makeTextPreviewDup(0.25, 100);

                try { doc.selection = null; } catch (_) { }
                app.redraw();
                return;
            }

            function makePreviewDup(f, op) {
                var d = null;
                try { d = baseItem.duplicate(); } catch (_) { d = null; }
                if (!d) return null;

                var s = 100 + (scale - 100) * f;
                try { d.resize(s, s, true, true, true, true, true, Transformation.CENTER); } catch (_) { }
                try { d.translate(dx * f, dy * f); } catch (_) { }
                try { d.move(originalPath, ElementPlacement.PLACEBEFORE); } catch (_) { }
                try { d.opacity = op; } catch (_) { }
                return d;
            }

            makePreviewDup(1, 40);
            makePreviewDup(0.75, 60);
            makePreviewDup(0.5, 80);
            makePreviewDup(0.25, 100);

            try { doc.selection = null; } catch (_) { }
            app.redraw();
            return;
        }

        // --- 確定実行時のみ複製を作成 ---
        var pathB = baseItem.duplicate();
        // scale: percent (100 = 等倍)
        pathB.resize(scale, scale, true, true, true, true, true, Transformation.CENTER);
        pathB.translate(dx, dy);

        // --- 実生成（サンプリング面生成） ---
        // v15系の「元形状と複製形状の間を面でつなぐ」方式に差し替え
        var shadowGroup = generateShadowBySampling(baseItem, dx, dy, scale, false);
        if (!shadowGroup) {
            // 失敗時は複製だけ残さない（プレビューではないため）
            try { pathB.remove(); } catch (_) { }
            alert(L('selectClosed'));
            return;
        }

        // 必要ならパスファインダーで合体（最終確定）
        try {
            doc.selection = null;
            shadowGroup.selected = true;
            app.executeMenuCommand('Live Pathfinder Add');
            app.executeMenuCommand('expandStyle');
        } catch (_) { }

        // 合体後の選択結果をまとめてグループ化（オフセット適用の受け皿）
        var resultItems = null;
        try { resultItems = doc.selection; } catch (_) { resultItems = null; }
        try { doc.selection = null; } catch (_) { }
        var cGroup = null;
        try {
            if (resultItems && resultItems.length) {
                cGroup = groupItems(doc, resultItems);
            } else {
                cGroup = shadowGroup;
            }
        } catch (_) { cGroup = shadowGroup; }

        // expandStyle 後に resultItems を新規グループへ移した場合、元の shadowGroup が空で残ることがあるので削除
        try {
            if (shadowGroup && shadowGroup.isValid && cGroup && cGroup.isValid && shadowGroup !== cGroup) {
                shadowGroup.remove();
            }
        } catch (_) { }

        // 曲線分割で増えたアンカーのカクつきを抑える（曲線近似の点のみスムーズ化）
        smoothPolylineInItem(cGroup, 0.25, 150);

        // オフセット（ライブ効果）をCグループに適用
        applyOffsetEffectToCGroup(cGroup);

        // C を A の背面へ
        try { cGroup.move(originalPath, ElementPlacement.PLACEAFTER); } catch (_) {
            try { cGroup.zOrder(ZOrderMethod.SENDTOBACK); } catch (_) { }
        }

        // プレビュー用の複製（pathB）は不要（確定時は消す）
        try { pathB.remove(); } catch (_) { }

        // 選択解除
        try { doc.selection = null; } catch (_) { }

        app.redraw();
        return;

        } finally {
            // Group/Text 由来の一時ベース（②）が残らないように、どの経路でも必ず掃除する
            if (baseIsTemp) {
                try { baseCleanup(); } catch (_) { }
            }
        }

    }

    /* プレビュー更新 / Update preview */
    function updatePreview() {
        if (isPreviewing) {
            app.undo();
            isPreviewing = false;
        }
        if (chkPreview.value) {
            doExtrude(false);
            isPreviewing = true;
        }
        app.redraw();
    }

    /* イベントリスナー / Event listeners */
    // 値が変わった時
    inputDistance.onChanging = function () { if (chkPreview.value) updatePreview(); };
    inputAngle.onChanging = function () { if (chkPreview.value) updatePreview(); };
    inputScale.onChanging = function () { if (chkPreview.value) updatePreview(); };

    // プレビューチェックボックスがクリックされた時
    chkPreview.onClick = updatePreview;

    // OKボタン
    btnOk.onClick = function () {
        if (isPreviewing) {
            app.undo();
            isPreviewing = false;
        }
        doExtrude(true);
        dlg.close();
    };

    // キャンセルボタン
    btnCancel.onClick = function () {
        if (isPreviewing) {
            app.undo();
            isPreviewing = false;
        }
        dlg.close();
    };

    // ダイアログを開いた時点でプレビューを表示
    if (chkPreview && chkPreview.value) {
        updatePreview();
    }
    dlg.show();

})();
