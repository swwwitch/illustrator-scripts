#target illustrator
#targetengine "pathfinder-palette"
app.preferences.setBooleanPreference("ShowExternalJSXWarning", false);

/*
AiSmartPathfinder.jsx — Smart Pathfinder Palette

選択した複数オブジェクトにパスファインダーを適用する常駐パレット。
アイコンをクリックすると、その操作をメインエンジンへ委譲して即時に実行する。

パネル構成（上から）:
  モード         出力モードを排他ラジオで選択（下記 A/B/C。ショートカット P/C/F）
  形状モード      合体／前面型抜き／交差／中マド（Adobe Pathfinder command 0〜3）
  パスファインダー 分割／刈り込み／合流／切り抜き／アウトライン／背面型抜き（3個×2行, command 4〜9）
  オプション      余分なポイントを削除（RemovePoints）／塗りのないアートワークを削除（ExtractUnpainted, 分割・アウトラインのみ）／［拡張］ボタン

出力モード（モードパネルの排他ラジオ）:
  A パスファインダーを実行  上段・下段とも＝グループ化→XML（Adobe Pathfinder）→拡張→グループ解除（実際にパスへ変換）
  B 複合シェイプを作成      上段のみ＝ダイナミックアクション（ai_compound_shape）で複合シェイプ（下段6ボタンはディム＝無効）
  C 効果として適用          上段・下段とも＝XML ライブ効果のまま（拡張しない）
  ※ ダイナミックアクションの複合シェイプは expandStyle で焼き込めないため、A の上段も XML 方式を使う
  ※ A/C で複数オブジェクトを選択している場合は、効果適用の前に一時的にグループ化する（1つの効果対象にまとめるため）。
     A では拡張後にそのグループを解除して実パスへ戻す。C はライブ効果のためグループのまま残る。

形状モード（上段4ボタン）の Option+クリック:
  出力モードに関係なく複合シェイプを作成する（＝モード B と同じ。拡張はしない）。

［拡張］ボタン（オプションパネル）:
  選択中の複合シェイプ（DOM 上 PluginItem）を通常のパスへ拡張する（ai_expand_compound_shape）。
  パレットは選択変更を受け取れないため常に押せ、複合シェイプの有無はクリック時に判定する。

Persistent palette that applies Pathfinder operations to the current selection.
Panels top-down: Mode (A/B/C radios; shortcuts P/C/F), Shape Mode (Unite/Minus
Front/Intersect/Exclude), Pathfinders (Divide/Trim/Merge/Crop/Outline/Minus
Back), Options (Remove points / Remove unpainted / Expand button).
A = Apply: both rows group, apply the Adobe Pathfinder XML, expand, and ungroup
(bake to real paths). B = Compound shape: Shape Modes only via the
ai_compound_shape action (Pathfinders dimmed/disabled). C = Apply as effect:
both rows keep the live Adobe Pathfinder effect. "Remove unpainted artwork"
affects Divide / Outline only. Option-clicking a Shape Mode button always makes a
compound shape regardless of the output mode. The Expand button turns a selected
compound shape back into plain paths.

構成 / Structure
- 常駐エンジン（#targetengine）でパレット参照を保持し GC を回避
- 表示中の常駐 app は DOM 接続を失うため、DOM 操作は BridgeTalk でメインエンジンへ委譲
- 委譲は worker 関数を toString → encodeURIComponent → eval(decodeURIComponent(...)) で送信
- worker の toString はコメントを取り込み壊すため、送信前に stripWorkerComments で除去（worker 本体にはコメントを書かず、説明は JSDoc に集約する）
- 戻り値はマーカー（OK / NODOC / NOSEL / NEEDTWO / NOCS / ERR:...）→ ローカライズして下部 status に表示
- 選択不足は条件で分離：効果は1つ以上（NOSEL）、実行・複合シェイプは2つ以上（NEEDTWO）
- 複数選択時は効果を1つの対象にまとめるため一時的にグループ化し、A（destructive）では拡張後に解除する（エラー時はその一時グループだけを選択し直して解除）
- UI は buildModePanel / buildShapeModePanel / buildPathfinderRows / buildOptionPanel / buildStatusText と addOperationButton で構築し、setupWindow / setupPanel と PANEL_MARGINS 等の共通変数で統一
- アイコンは onDraw で描画（無効時は dimIconColors でディム表示）
- Option（alt）状態は onDraw では取れないため mousedown で記録し onClick で読む
- 全体を IIFE で閉じ、$.global にはパレット参照（__pfPaletteWindow）だけを残す
- 二重起動回避：既に開いていれば作り直さず前面化して終了
- キーボード：Esc で閉じる／P・C・F で出力モード切替
*/

// =========================================
// バージョン / Version
// =========================================

var SCRIPT_VERSION = "v1.0.5";

/* エンジンのグローバルを汚さないため IIFE で閉じる。パレット参照だけ $.global に残す。
 * Wrap everything in an IIFE; only the palette reference lives on $.global. */
(function () {

/* ============================================================
 * ローカライズ / Localization
 * ============================================================ */

/**
 * 現在の UI 言語を返す。
 * @returns {string} "ja" または "en" / "ja" or "en"
 */
function getCurrentLanguage() {
    return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
}

var currentLanguage = getCurrentLanguage();

var LABELS = {
    dialog: {
        title: { ja: "パスファインダー", en: "Pathfinder" }
    },
    panel: {
        shapeMode:  { ja: "形状モード",     en: "Shape Mode" },
        pathfinder: { ja: "パスファインダー", en: "Pathfinders" },
        mode:       { ja: "モード",         en: "Mode" },
        option:     { ja: "オプション",     en: "Options" }
    },
    mode: {
        unite:      { ja: "合体",                     en: "Unite(Add)" },
        minusFront: { ja: "前面オブジェクトで型抜き", en: "Minus Front" },
        intersect:  { ja: "交差",                     en: "Intersect" },
        exclude:    { ja: "中マド",                   en: "Exclude" }
    },
    pathfinder: {
        divide:    { ja: "分割",                   en: "Divide" },
        trim:      { ja: "刈り込み",               en: "Trim" },
        merge:     { ja: "合流",                   en: "Merge" },
        crop:      { ja: "切り抜き",               en: "Crop" },
        outline:   { ja: "アウトライン",           en: "Outline" },
        minusBack: { ja: "背面オブジェクトで型抜き", en: "Minus Back" }
    },
    apply: {
        execute:  { ja: "パスファインダーを実行", en: "Apply Pathfinder" },
        compound: { ja: "複合シェイプを作成",     en: "Compound shape" },
        effect:   { ja: "効果として適用",         en: "Apply as effect" }
    },
    button: {
        expand: { ja: "複合シェイプを拡張", en: "Expand Compound Shape" }
    },
    option: {
        removePoints:    { ja: "余分なポイントを削除",       en: "Remove redundant points" },
        removeUnpainted: { ja: "塗りのないアートワークを削除", en: "Remove unpainted artwork" }
    },
    status: {
        ready:   { ja: "アイコンをクリックして実行",       en: "Click an icon to apply" },
        applied: { ja: "適用しました",                    en: "Applied." },
        noDoc:   { ja: "ドキュメントが開かれていません",  en: "No document is open." },
        noSel:   { ja: "1つ以上のオブジェクトを選択してください", en: "Select one or more objects." },
        needTwo: { ja: "2つ以上のオブジェクトを選択してください", en: "Select two or more objects." },
        noCompound: { ja: "複合シェイプを選択してください", en: "Select a compound shape." },
        timeout: { ja: "タイムアウトしました",            en: "Timed out." },
        error:   { ja: "エラー: ",                        en: "Error: " }
    },
    tip: {
        esc:             { ja: "Esc: パレットを閉じる",     en: "Esc: close the palette" },
        removeUnpainted: { ja: "分割・アウトラインのみ有効", en: "Divide / Outline only" },
        compoundApply:   { ja: "形状モード（上段）のみ",    en: "Shape Mode (top row) only" },
        optionCompound:  { ja: "Option+クリックで複合シェイプ", en: "Option-click to make a compound shape" },
        expand:          { ja: "選択中の複合シェイプを通常のパスに拡張", en: "Expand the selected compound shape to paths" },
        shortcutExecute:  { ja: "ショートカット: P", en: "Shortcut: P" },
        shortcutCompound: { ja: "ショートカット: C", en: "Shortcut: C" },
        shortcutEffect:   { ja: "ショートカット: F", en: "Shortcut: F" }
    }
};

/**
 * ドットパスでローカライズ文字列を引く。存在しないキーは null 耐性でパス文字列を返す。
 * @param {string} dotPath 例 "mode.unite" / e.g. "mode.unite"
 * @returns {string} ローカライズ済み文字列 / localized string
 */
function getLocalizedText(dotPath) {
    var parts = String(dotPath).split(".");
    var labelNode = LABELS;
    for (var i = 0; i < parts.length; i++) {
        if (labelNode == null) return dotPath;
        labelNode = labelNode[parts[i]];
    }
    if (labelNode == null) return dotPath;
    if (typeof labelNode[currentLanguage] === "string") return labelNode[currentLanguage];
    if (typeof labelNode.en === "string") return labelNode.en;
    return dotPath;
}

/* ============================================================
 * シェイプモード定義 / Shape mode definitions
 * ============================================================ */

/* compoundValue と name は記録済み .aia から採取（ai_compound_shape の enumerated 値）
 * compoundValue    : B（複合シェイプ）用 ai_compound_shape の enumerated 値 / enumerated value for the compound-shape action
 * pathfinderCommand: A/C（実行・効果）用 Pathfinder XML の Command 番号 / Pathfinder XML Command index
 * name             : parameter-1 の /name（.aia に記録された表示名）/ recorded parameter name
 * icon             : onDraw の描画種別 / icon draw type
 * labelKey         : getLocalizedText() のドットパス / dotted label key
 *
 * ★ 番号体系を分離して持つ / NOTE: two separate numbering systems, held in separate fields
 *   複合シェイプの enumerated 値と Pathfinder XML の Command は本来別体系。
 *   0/1/2/3 はたまたま両者で一致しているが、片方の割り当てが変わっても壊れないよう
 *   compoundValue / pathfinderCommand を別フィールドとして明示的に持つ。
 *   下段 PATHFINDER_MODES.command（4〜9）と合わせて Pathfinder Command は 0〜9 の連番になる。
 */
var SHAPE_MODES = [
    { compoundValue: 0, pathfinderCommand: 0, name: "追加",                   icon: "unite",      labelKey: "mode.unite" },
    { compoundValue: 3, pathfinderCommand: 3, name: "前面オブジェクトで型抜き", icon: "minusFront", labelKey: "mode.minusFront" },
    { compoundValue: 1, pathfinderCommand: 1, name: "交差",                   icon: "intersect",  labelKey: "mode.intersect" },
    { compoundValue: 2, pathfinderCommand: 2, name: "中マド",                 icon: "exclude",    labelKey: "mode.exclude" }
];

/* パスファインダー（Adobe Pathfinder ライブ効果）/ Pathfinders (Adobe Pathfinder live effect)
 * command  : ライブ効果の Command 番号 / live effect Command index
 * icon     : onDraw の描画種別 / icon draw type
 * labelKey : getLocalizedText() のドットパス / dotted label key
 * unpainted: 「塗りのないアートワークを削除」が効くか（分割・アウトラインのみ）/ ExtractUnpainted applies (Divide/Outline only)
 */
var PATHFINDER_MODES = [
    { command: 5, icon: "divide",    labelKey: "pathfinder.divide",    unpainted: true },
    { command: 7, icon: "trim",      labelKey: "pathfinder.trim",      unpainted: false },
    { command: 8, icon: "merge",     labelKey: "pathfinder.merge",     unpainted: false },
    { command: 9, icon: "crop",      labelKey: "pathfinder.crop",      unpainted: false },
    { command: 6, icon: "outline",   labelKey: "pathfinder.outline",   unpainted: true },
    { command: 4, icon: "minusBack", labelKey: "pathfinder.minusBack", unpainted: false }
];

/* ============================================================
 * worker 関数（メインエンジンで実行）/ Worker functions (run in main engine)
 * ------------------------------------------------------------
 * ・DOM を触る処理はすべてここに集約し、押下のたびにメインエンジンへ委譲する
 * ・toString は改行を消すため、必ずセミコロンで終える
 * ・関数「本体内」にはコメント（// も /* *\/ も）を書かない。stripWorkerComments は
 *   JSDoc（/**）のみを対象にするため、本体内コメントは除去されず eval を壊しうる
 * ・文字列・正規表現に /** を書かない（誤除去の原因になる）
 * ・追加・分割したら必ず WORKER_FUNCS に登録する
 * ============================================================ */

/**
 * 選択オブジェクトを複合シェイプ化する（メインエンジン用エントリ）。
 * keepCompound が false のときは拡張してフラットなパスにする。
 * @param {number} shapeModeValue enumerated 値 / enumerated value
 * @param {string} shapeModeName parameter-1 の /name / recorded parameter name
 * @param {boolean} keepCompound 複合シェイプのまま残すか（false なら拡張）/ keep as compound shape (false expands)
 * @returns {string} マーカー "OK" / "NODOC" / "NOSEL"（1つも未選択）/ "NEEDTWO"（2つ未満）/ "ERR:..."
 */
function workerApplyCompoundShape(shapeModeValue, shapeModeName, keepCompound) {
    if (app.documents.length === 0) { return "NODOC"; }
    var currentSelection = app.activeDocument.selection;
    if (!currentSelection || currentSelection.length < 1) { return "NOSEL"; }
    if (currentSelection.length < 2) { return "NEEDTWO"; }
    var uniqueToken = "AiSmartPathfinder_"
        + (new Date()).getTime()
        + "_"
        + Math.floor(Math.random() * 100000);
    var actionConfig = {
        setName: uniqueToken + "_set",
        actionName: uniqueToken + "_action",
        internalName: "ai_compound_shape",
        localizedName: "複合シェイプ",
        shapeModeKey: 1851878757,
        shapeModeName: shapeModeName,
        shapeModeValue: shapeModeValue,
        expandKey: 1836016741,
        actionFilePath: Folder.temp.fsName + "/" + uniqueToken + ".aia"
    };
    try {
        var actionSource = buildActionSource(actionConfig);
        playTemporaryAction(actionSource, actionConfig.setName, actionConfig.actionName, actionConfig.actionFilePath);
        if (!keepCompound) { app.executeMenuCommand("expandStyle"); }
        app.redraw();
        return "OK";
    } catch (applyError) {
        return "ERR:" + applyError;
    }
}

/**
 * 選択オブジェクトにパスファインダー（Adobe Pathfinder ライブ効果）を適用する（メインエンジン用エントリ）。
 * 複数選択時は効果対象を1つにまとめるため一時的にグループ化する（効果のまま destructive=false なら単体でも実行可）。
 * destructive が true のときは適用後に拡張し、その一時グループを解除してフラットなパスへ戻す。
 * エラー時は作成した一時グループだけを選択し直して解除する（他の階層・選択には触れない）。
 * ※ この関数は BridgeTalk 委譲で toString 送信されるため、本体にコメントを書かず説明はこの JSDoc に集約する。
 * @param {number} command ライブ効果の Command 番号 / live effect Command index
 * @param {boolean} removeUnpainted 塗りのないアートワークを削除（分割・アウトラインのみ有効）/ remove unpainted (Divide/Outline only)
 * @param {boolean} removePoints 余分なポイントを削除 / remove redundant points
 * @param {boolean} destructive 実際にパスへ変換するか（true: 拡張＋グループ解除 / false: ライブ効果のまま）/ bake to paths
 * @returns {string} マーカー "OK" / "NODOC" / "NOSEL"（1つも未選択）/ "NEEDTWO"（実行モードで2つ未満）/ "ERR:..."
 */
function workerApplyPathfinder(command, removeUnpainted, removePoints, destructive) {
    if (app.documents.length === 0) { return "NODOC"; }
    var currentDocument = app.activeDocument;
    var currentSelection = currentDocument.selection;
    if (!currentSelection || currentSelection.length < 1) { return "NOSEL"; }
    if (destructive && currentSelection.length < 2) { return "NEEDTWO"; }

    var groupedForOperation = false;
    var temporaryGroup = null;
    try {
        var targetItem;
        if (currentSelection.length >= 2) {
            app.executeMenuCommand("group");
            groupedForOperation = true;
            temporaryGroup = currentDocument.selection[0];
            targetItem = temporaryGroup;
        } else {
            targetItem = currentSelection[0];
        }

        targetItem.applyEffect(buildPathfinderXML(command, removeUnpainted, removePoints));
        app.redraw();

        if (destructive) {
            app.executeMenuCommand("expandStyle");
            app.executeMenuCommand("ungroup");
            groupedForOperation = false;
            temporaryGroup = null;
        }

        app.redraw();
        return "OK";
    } catch (pathfinderError) {
        if (groupedForOperation && temporaryGroup) {
            try {
                currentDocument.selection = null;
                temporaryGroup.selected = true;
                if (currentDocument.selection.length === 1
                        && currentDocument.selection[0].typename === "GroupItem") {
                    app.executeMenuCommand("ungroup");
                }
            } catch (rollbackError) { }
        }
        return "ERR:" + pathfinderError;
    }
}

/**
 * Adobe Pathfinder ライブ効果の XML を生成する。
 * @param {number} command Command 番号 / Command index
 * @param {boolean} removeUnpainted ExtractUnpainted の値（塗りのないアートワークを削除）/ ExtractUnpainted flag
 * @param {boolean} removePoints RemovePoints の値（余分なポイントを削除）/ RemovePoints flag
 * @returns {string} ライブ効果 XML / live effect XML
 */
function buildPathfinderXML(command, removeUnpainted, removePoints) {
    var displayNames = ['Add', 'Intersect', 'Exclude', 'Minus Front', 'Minus Back', 'Divide', 'Outline', 'Trim', 'Merge', 'Crop', 'Hard Mix', 'Soft Mix', 'Trap'];
    return '<LiveEffect name="Adobe Pathfinder" isPre="1"><Dict data="I Command ' + command
        + ' B ConvertCustom 1 B ExtractUnpainted ' + (removeUnpainted ? 1 : 0)
        + ' R Mix 0.5 R Precision 10 B RemovePoints ' + (removePoints ? 1 : 0)
        + ' R TrapAspect 1 B TrapConvertCustom 1 R TrapMaxTint 1 B TrapReverse 0 R TrapThickness 0.25 R TrapTint 0.4 R TrapTintTolerance 0.05">'
        + '<Entry name="DisplayString" value="' + displayNames[command] + '" valueType="S"/></Dict></LiveEffect>';
}

/**
 * 複合シェイプ用の一時アクションソースを生成する。
 * @param {object} actionConfig setName / actionName / internalName / localizedName / shapeModeKey / shapeModeName / shapeModeValue / expandKey を持つ設定
 * @returns {string} .aia のソース文字列 / .aia source text
 */
function buildActionSource(actionConfig) {
    return ''
        + '/version 3\n'
        + buildNameLine('/name', actionConfig.setName)
        + '/isOpen 1\n'
        + '/actionCount 1\n'
        + '/action-1 {\n'
        + buildNameLine(' /name', actionConfig.actionName)
        + ' /keyIndex 0\n'
        + ' /colorIndex 0\n'
        + ' /isOpen 0\n'
        + ' /eventCount 1\n'
        + ' /event-1 {\n'
        + ' /useRulersIn1stQuadrant 0\n'
        + ' /internalName (' + actionConfig.internalName + ')\n'
        + buildNameLine(' /localizedName', actionConfig.localizedName)
        + ' /isOpen 1\n'
        + ' /isOn 1\n'
        + ' /hasDialog 0\n'
        + ' /parameterCount 2\n'
        + ' /parameter-1 {\n'
        + ' /key ' + actionConfig.shapeModeKey + '\n'
        + ' /showInPalette 4294967295\n'
        + ' /type (enumerated)\n'
        + buildNameLine(' /name', actionConfig.shapeModeName)
        + ' /value ' + actionConfig.shapeModeValue + '\n'
        + ' }\n'
        + ' /parameter-2 {\n'
        + ' /key ' + actionConfig.expandKey + '\n'
        + ' /showInPalette 4294967295\n'
        + ' /type (integer)\n'
        + ' /value 0\n'
        + ' }\n'
        + ' }\n'
        + '}\n';
}

/**
 * `/name [ <byteCount> <utf8Hex> ]` 形式の1行を生成する。
 * @param {string} prefix 行頭のキー（先頭スペース含む）/ line key prefix
 * @param {string} text 対象文字列 / target string
 * @returns {string} 生成した1行 / generated line
 */
function buildNameLine(prefix, text) {
    var encoded = stringToUtf8Hex(text);
    return prefix + ' [ ' + encoded.byteCount + ' ' + encoded.hex + ' ]\n';
}

/**
 * 文字列を UTF-8 バイト列の16進表記に変換する（.aia の名前フィールド用）。
 * サロゲートペア（U+10000 以上、絵文字等）は結合して 4 バイトで符号化する。
 * @param {string} sourceText 変換対象 / source string
 * @returns {{hex: string, byteCount: number}} 16進文字列とバイト数 / hex text and byte count
 */
function stringToUtf8Hex(sourceText) {
    var hexText = "";
    var byteCount = 0;
    for (var i = 0; i < sourceText.length; i++) {
        var codePoint = sourceText.charCodeAt(i);
        if (codePoint >= 0xD800 && codePoint <= 0xDBFF && i + 1 < sourceText.length) {
            var lowSurrogate = sourceText.charCodeAt(i + 1);
            if (lowSurrogate >= 0xDC00 && lowSurrogate <= 0xDFFF) {
                codePoint = 0x10000 + ((codePoint - 0xD800) << 10) + (lowSurrogate - 0xDC00);
                i++;
            }
        }
        var bytes;
        if (codePoint < 0x80) {
            bytes = [codePoint];
        } else if (codePoint < 0x800) {
            bytes = [0xC0 | (codePoint >> 6), 0x80 | (codePoint & 0x3F)];
        } else if (codePoint < 0x10000) {
            bytes = [0xE0 | (codePoint >> 12), 0x80 | ((codePoint >> 6) & 0x3F), 0x80 | (codePoint & 0x3F)];
        } else {
            bytes = [0xF0 | (codePoint >> 18), 0x80 | ((codePoint >> 12) & 0x3F), 0x80 | ((codePoint >> 6) & 0x3F), 0x80 | (codePoint & 0x3F)];
        }
        for (var byteIndex = 0; byteIndex < bytes.length; byteIndex++) {
            var singleHex = bytes[byteIndex].toString(16);
            if (singleHex.length < 2) { singleHex = "0" + singleHex; }
            hexText += singleHex;
            byteCount++;
        }
    }
    return { hex: hexText, byteCount: byteCount };
}

/**
 * 一時アクションをファイル化 → ロード → 実行し、後始末する。
 * encoding は open より前に設定する（open 後だと書き込みに反映されない場合があるため）。
 * unloadAction はロードに成功したときだけ行う（未ロードのセットを外そうとして例外を出さない）。
 * ※ この関数は BridgeTalk 委譲で toString 送信されるため、本体にコメントを書かず説明はこの JSDoc に集約する。
 * @param {string} actionSource .aia のソース文字列 / .aia source text
 * @param {string} setName アクションセット名 / action set name
 * @param {string} actionName アクション名 / action name
 * @param {string} actionFilePath 一時ファイルパス / temp file path
 * @returns {void}
 */
function playTemporaryAction(actionSource, setName, actionName, actionFilePath) {
    var actionFile = new File(actionFilePath);
    var isActionLoaded = false;
    var isActionFileOpen = false;
    actionFile.encoding = "UTF-8";
    try {
        if (!actionFile.open("w")) { throw new Error("Failed to open temporary action file."); }
        isActionFileOpen = true;
        actionFile.write(actionSource);
        actionFile.close();
        isActionFileOpen = false;
        app.loadAction(actionFile);
        isActionLoaded = true;
        app.doScript(actionName, setName, false);
    } finally {
        if (isActionFileOpen) { try { actionFile.close(); } catch (closeError) { } }
        if (actionFile.exists) { try { actionFile.remove(); } catch (removeError) { } }
        if (isActionLoaded) { try { app.unloadAction(setName, ""); } catch (unloadError2) { } }
    }
}

/**
 * 選択中の複合シェイプを通常のパスへ拡張する（メインエンジン用エントリ）。
 * クリック時に選択を判定し、複合シェイプ（DOM 上 PluginItem）が無ければ "NOCS" を返す。
 * 複合シェイプがあればダイナミックアクション ai_expand_compound_shape を一時アクションとして再生する。
 * @returns {string} マーカー "OK" / "NODOC" / "NOCS"（複合シェイプ未選択）/ "ERR:..."
 */
function workerExpandCompoundShape() {
    if (app.documents.length === 0) { return "NODOC"; }
    var currentSelection = app.activeDocument.selection;
    if (!currentSelection || currentSelection.length < 1) { return "NOCS"; }
    var hasCompoundShape = false;
    for (var selectionIndex = 0; selectionIndex < currentSelection.length; selectionIndex++) {
        if (currentSelection[selectionIndex].typename === "PluginItem") { hasCompoundShape = true; break; }
    }
    if (!hasCompoundShape) { return "NOCS"; }
    var uniqueToken = "AiSmartPathfinder_expand_"
        + (new Date()).getTime()
        + "_"
        + Math.floor(Math.random() * 100000);
    var actionConfig = {
        setName: uniqueToken + "_set",
        actionName: uniqueToken + "_action",
        internalName: "ai_expand_compound_shape",
        localizedName: "複合シェイプを拡張",
        expandParamKey: 2020634212,
        actionFilePath: Folder.temp.fsName + "/" + uniqueToken + ".aia"
    };
    try {
        var actionSource = buildExpandActionSource(actionConfig);
        playTemporaryAction(actionSource, actionConfig.setName, actionConfig.actionName, actionConfig.actionFilePath);
        app.redraw();
        return "OK";
    } catch (expandError) {
        return "ERR:" + expandError;
    }
}

/**
 * 複合シェイプ拡張用の一時アクションソースを生成する（integer パラメータ1個）。
 * @param {object} actionConfig setName / actionName / internalName / localizedName / expandParamKey を持つ設定
 * @returns {string} .aia のソース文字列 / .aia source text
 */
function buildExpandActionSource(actionConfig) {
    return ''
        + '/version 3\n'
        + buildNameLine('/name', actionConfig.setName)
        + '/isOpen 1\n'
        + '/actionCount 1\n'
        + '/action-1 {\n'
        + buildNameLine(' /name', actionConfig.actionName)
        + ' /keyIndex 0\n'
        + ' /colorIndex 0\n'
        + ' /isOpen 0\n'
        + ' /eventCount 1\n'
        + ' /event-1 {\n'
        + ' /useRulersIn1stQuadrant 0\n'
        + ' /internalName (' + actionConfig.internalName + ')\n'
        + buildNameLine(' /localizedName', actionConfig.localizedName)
        + ' /isOpen 0\n'
        + ' /isOn 1\n'
        + ' /hasDialog 0\n'
        + ' /parameterCount 1\n'
        + ' /parameter-1 {\n'
        + ' /key ' + actionConfig.expandParamKey + '\n'
        + ' /showInPalette 4294967295\n'
        + ' /type (integer)\n'
        + ' /value 0\n'
        + ' }\n'
        + ' }\n'
        + '}\n';
}

/* 委譲する worker 関数はすべてここに登録する（登録漏れ防止）/ register every delegated worker function */
var WORKER_FUNCS = [
    workerApplyCompoundShape,
    workerApplyPathfinder,
    workerExpandCompoundShape,
    buildPathfinderXML,
    buildActionSource,
    buildExpandActionSource,
    buildNameLine,
    stringToUtf8Hex,
    playTemporaryAction
];

/* ============================================================
 * BridgeTalk 委譲 / Delegation to the main engine
 * ============================================================ */

/**
 * worker 関数群と呼び出し式をメインエンジンで同期実行し、マーカー文字列を返す。
 * @param {string} callExpression メインエンジンで評価する呼び出し式 / call expression to evaluate
 * @returns {string} マーカー / marker string
 */
function delegateCall(callExpression) {
    var workerSource = "";
    for (var i = 0; i < WORKER_FUNCS.length; i++) {
        workerSource += WORKER_FUNCS[i].toString() + "\n";
    }
    /* ExtendScript の Function.toString() は関数間の JSDoc を末尾に取り込み、
     * さらにコメント終端を欠落させて「未終了コメント」を作る（eval 全体が壊れる）。
     * 全連結後にコメントを除去してから送る（次の function が終端の目印になる）。 */
    workerSource = stripWorkerComments(workerSource);
    var evalSource = workerSource + "\n" + callExpression + ";";

    var resultHolder = { value: "TIMEOUT" };
    var bridge = new BridgeTalk();
    bridge.target = "illustrator";
    bridge.body = 'eval(decodeURIComponent("' + encodeURIComponent(evalSource) + '"));';
    bridge.onResult = function (message) { resultHolder.value = String(message.body); };
    bridge.onError = function (message) { resultHolder.value = "ERR:" + String(message.body); };
    bridge.onTimeout = function () { resultHolder.value = "TIMEOUT"; };
    bridge.send(10);
    return resultHolder.value;
}

/**
 * 文字列を JS 呼び出し式へ安全に埋め込むためのダブルクォート付きリテラルに変換する。
 * @param {string} value 対象文字列 / target string
 * @returns {string} エスケープ済みの "..." リテラル / escaped double-quoted literal
 */
function quoteString(value) {
    return '"' + String(value)
        .replace(/\\/g, "\\\\")
        .replace(/"/g, '\\"')
        .replace(/\r/g, "\\r")
        .replace(/\n/g, "\\n") + '"';
}

function stripWorkerComments(source) {
    var output = "";
    var quoteCharacter = null;
    var isEscaped = false;
    var inBlockComment = false;
    var isJSDocComment = false;

    for (var i = 0; i < source.length; i++) {
        var currentCharacter = source.charAt(i);
        var nextCharacter = (i + 1 < source.length) ? source.charAt(i + 1) : "";
        var nextNextCharacter = (i + 2 < source.length) ? source.charAt(i + 2) : "";

        if (inBlockComment) {
            if (currentCharacter === "*" && nextCharacter === "/") {
                inBlockComment = false;
                isJSDocComment = false;
                i++;
            } else if (isJSDocComment && currentCharacter === "f"
                    && source.substr(i, 8) === "function") {
                inBlockComment = false;
                isJSDocComment = false;
                output += "function";
                i += 7;
            }
            continue;
        }

        if (quoteCharacter !== null) {
            output += currentCharacter;
            if (isEscaped) {
                isEscaped = false;
            } else if (currentCharacter === "\\") {
                isEscaped = true;
            } else if (currentCharacter === quoteCharacter) {
                quoteCharacter = null;
            }
            continue;
        }

        if (currentCharacter === '"' || currentCharacter === "'") {
            quoteCharacter = currentCharacter;
            output += currentCharacter;
            continue;
        }

        if (currentCharacter === "/" && nextCharacter === "*" && nextNextCharacter === "*") {
            inBlockComment = true;
            isJSDocComment = true;
            i += 2;
            continue;
        }

        output += currentCharacter;
    }

    return output;
}

/**
 * 選択オブジェクトを複合シェイプ化する / apply the compound shape
 * @param {number} shapeModeValue enumerated 値 / enumerated value
 * @param {string} shapeModeName parameter-1 の /name / recorded parameter name
 * @param {boolean} keepCompound 複合シェイプのまま残すか（false なら拡張）/ keep as compound shape (false expands)
 * @returns {string} マーカー / marker string
 */
function delegateApply(shapeModeValue, shapeModeName, keepCompound) {
    return delegateCall('workerApplyCompoundShape('
        + shapeModeValue + ', ' + quoteString(shapeModeName) + ', ' + (keepCompound ? 'true' : 'false') + ')');
}

/**
 * 選択オブジェクトにパスファインダー（ライブ効果）を適用する / apply a Pathfinder
 * @param {number} command ライブ効果の Command 番号 / live effect Command index
 * @param {boolean} removeUnpainted 塗りのないアートワークを削除 / remove unpainted
 * @param {boolean} removePoints 余分なポイントを削除 / remove redundant points
 * @param {boolean} destructive 実際にパスへ変換するか（拡張＋グループ解除）/ bake to paths
 * @returns {string} マーカー / marker string
 */
function delegatePathfinder(command, removeUnpainted, removePoints, destructive) {
    return delegateCall('workerApplyPathfinder(' + command + ', '
        + (removeUnpainted ? 'true' : 'false') + ', ' + (removePoints ? 'true' : 'false') + ', '
        + (destructive ? 'true' : 'false') + ')');
}

/**
 * 選択中の複合シェイプを拡張する / expand the selected compound shape
 * @returns {string} マーカー / marker string
 */
function delegateExpandCompoundShape() {
    return delegateCall('workerExpandCompoundShape()');
}

/**
 * マーカー文字列をローカライズした status テキストに変換する。
 * @param {string} marker delegateToMain の戻り値 / marker from delegateToMain
 * @param {object} mode 適用した SHAPE_MODES の要素 / applied shape mode
 * @returns {string} 表示用テキスト / status text
 */
function markerToStatus(marker, mode) {
    if (marker === "OK") { return getLocalizedText(mode.labelKey) + ": " + getLocalizedText("status.applied"); }
    if (marker === "NODOC") { return getLocalizedText("status.noDoc"); }
    if (marker === "NOSEL") { return getLocalizedText("status.noSel"); }
    if (marker === "NEEDTWO") { return getLocalizedText("status.needTwo"); }
    if (marker === "NOCS") { return getLocalizedText("status.noCompound"); }
    if (marker === "TIMEOUT") { return getLocalizedText("status.timeout"); }
    if (marker.indexOf("ERR:") === 0) { return getLocalizedText("status.error") + marker.substring(4); }
    return marker;
}

/* ============================================================
 * アイコン描画 / Icon drawing
 * ============================================================ */

/**
 * UI が明るいテーマかどうかを判定する。
 * @returns {boolean} 明るい UI なら true / true if the UI is light
 */
function isLightUI() {
    try {
        return app.preferences.getRealPreference("uiBrightness") > 0.5;
    } catch (brightnessError) {
        return false; /* 取得失敗時は暗い側にフォールバック / fall back to dark */
    }
}

/**
 * UI の明暗に応じたアイコン色・背景色・中間色を返す。
 * @returns {{icon: number[], bg: number[], muted: number[]}} 描画色（RGBA 0〜1）/ drawing colors (RGBA 0-1)
 */
function getIconColors() {
    if (isLightUI()) {
        return { icon: [0.33, 0.33, 0.33, 1], bg: [0.93, 0.93, 0.93, 1], muted: [0.62, 0.62, 0.62, 1] };
    }
    return { icon: [0.85, 0.85, 0.85, 1], bg: [0.27, 0.27, 0.27, 1], muted: [0.55, 0.55, 0.55, 1] };
}

/**
 * 無効（ディム）表示用に、アイコン色を背景色へ寄せて薄くする。
 * @param {{icon: number[], bg: number[], muted: number[]}} colors 通常色 / normal colors
 * @returns {{icon: number[], bg: number[], muted: number[]}} ディム色 / dimmed colors
 */
function dimIconColors(colors) {
    var towardBg = 0.6; /* 0=そのまま / 1=背景色 / blend factor toward background */
    function blendTowardBackground(color, background) {
        return [
            color[0] + (background[0] - color[0]) * towardBg,
            color[1] + (background[1] - color[1]) * towardBg,
            color[2] + (background[2] - color[2]) * towardBg,
            1
        ];
    }
    return { icon: blendTowardBackground(colors.icon, colors.bg), bg: colors.bg, muted: blendTowardBackground(colors.muted, colors.bg) };
}

/**
 * 矩形を塗る / fill a rectangle
 * @param {object} graphics ScriptUIGraphics
 * @param {object} brush ブラシ / brush
 * @param {number} x 左 / left
 * @param {number} y 上 / top
 * @param {number} rectWidth 幅 / width
 * @param {number} rectHeight 高さ / height
 * @returns {void}
 */
function fillRect(graphics, brush, x, y, rectWidth, rectHeight) {
    graphics.newPath();
    graphics.rectPath(x, y, rectWidth, rectHeight);
    graphics.fillPath(brush);
}

/**
 * 矩形の輪郭を描く / stroke a rectangle
 * @param {object} graphics ScriptUIGraphics
 * @param {object} pen ペン / pen
 * @param {number} x 左 / left
 * @param {number} y 上 / top
 * @param {number} rectWidth 幅 / width
 * @param {number} rectHeight 高さ / height
 * @returns {void}
 */
function strokeRect(graphics, pen, x, y, rectWidth, rectHeight) {
    graphics.newPath();
    graphics.rectPath(x, y, rectWidth, rectHeight);
    graphics.strokePath(pen);
}

/**
 * 直線を描く / draw a line
 * @param {object} graphics ScriptUIGraphics
 * @param {object} pen ペン / pen
 * @param {number} x1 始点 x / start x
 * @param {number} y1 始点 y / start y
 * @param {number} x2 終点 x / end x
 * @param {number} y2 終点 y / end y
 * @returns {void}
 */
function drawLine(graphics, pen, x1, y1, x2, y2) {
    graphics.newPath();
    graphics.moveTo(x1, y1);
    graphics.lineTo(x2, y2);
    graphics.strokePath(pen);
}

/**
 * パスファインダー操作のアイコンを iconbutton に描画する（形状モード・パスファインダー共通）。
 * 左上（back）と右下（front）の2つの正方形の重なりで各操作を表現する。無効時はディム表示。
 * @param {object} control 描画対象の iconbutton / the iconbutton being drawn
 * @param {string} iconType 描画種別（unite / minusFront / intersect / exclude / divide / trim / merge / crop / outline / minusBack）
 * @returns {void}
 */
function drawOperationIcon(control, iconType) {
    var graphics = control.graphics;
    var colors = getIconColors();
    if (!control.enabled) { colors = dimIconColors(colors); }
    var iconBrush = graphics.newBrush(graphics.BrushType.SOLID_COLOR, colors.icon);
    var backgroundBrush = graphics.newBrush(graphics.BrushType.SOLID_COLOR, colors.bg);
    var iconPen = graphics.newPen(graphics.PenType.SOLID_COLOR, colors.icon, 2);
    var backgroundPen = graphics.newPen(graphics.PenType.SOLID_COLOR, colors.bg, 2);
    var mutedPen = graphics.newPen(graphics.PenType.SOLID_COLOR, colors.muted, 2);

    /* 背景を塗ってネイティブ枠を隠す / paint background to hide the native frame */
    fillRect(graphics, backgroundBrush, 0, 0, control.size[0], control.size[1]);

    /* 2つの正方形の配置 / two overlapping squares */
    var side = 22;
    var backX = 7, backY = 7;
    var frontX = 17, frontY = 17;
    var overlapWidth = (backX + side) - frontX;   /* 重なりの幅 / overlap width = 12 */
    var overlapHeight = (backY + side) - frontY;  /* 重なりの高さ / overlap height = 12 */

    if (iconType === "unite" || iconType === "merge") {
        /* 合体・合流：2つを同色で塗り、継ぎ目なしの和集合シルエット */
        fillRect(graphics, iconBrush, backX, backY, side, side);
        fillRect(graphics, iconBrush, frontX, frontY, side, side);
    } else if (iconType === "minusFront") {
        /* 前面型抜き：back を塗り、front を背景色で抜いて輪郭のみ残す */
        fillRect(graphics, iconBrush, backX, backY, side, side);
        fillRect(graphics, backgroundBrush, frontX, frontY, side, side);
        strokeRect(graphics, iconPen, frontX, frontY, side, side);
    } else if (iconType === "intersect") {
        /* 交差：2つは輪郭のみ、重なり部分だけ塗りつぶす */
        strokeRect(graphics, iconPen, backX, backY, side, side);
        strokeRect(graphics, iconPen, frontX, frontY, side, side);
        fillRect(graphics, iconBrush, frontX, frontY, overlapWidth, overlapHeight);
    } else if (iconType === "exclude") {
        /* 中マド：2つを塗り、重なり部分を背景色で抜く */
        fillRect(graphics, iconBrush, backX, backY, side, side);
        fillRect(graphics, iconBrush, frontX, frontY, side, side);
        fillRect(graphics, backgroundBrush, frontX, frontY, overlapWidth, overlapHeight);
    } else if (iconType === "divide") {
        /* 分割：2つを塗り、front の輪郭を背景色の線で入れて分割の継ぎ目を見せる */
        fillRect(graphics, iconBrush, backX, backY, side, side);
        fillRect(graphics, iconBrush, frontX, frontY, side, side);
        strokeRect(graphics, backgroundPen, frontX, frontY, side, side);
    } else if (iconType === "trim") {
        /* 刈り込み：2つを塗り、重なり境界（front の上辺・左辺）に背景色の継ぎ目線 */
        fillRect(graphics, iconBrush, backX, backY, side, side);
        fillRect(graphics, iconBrush, frontX, frontY, side, side);
        drawLine(graphics, backgroundPen, frontX, frontY, backX + side, frontY);
        drawLine(graphics, backgroundPen, frontX, frontY, frontX, backY + side);
    } else if (iconType === "crop") {
        /* 切り抜き：back は中間色の輪郭のみ、front を塗る（最前面で切り抜く） */
        strokeRect(graphics, mutedPen, backX, backY, side, side);
        fillRect(graphics, iconBrush, frontX, frontY, side, side);
    } else if (iconType === "outline") {
        /* アウトライン：2つとも輪郭のみ（線に変換） */
        strokeRect(graphics, mutedPen, backX, backY, side, side);
        strokeRect(graphics, iconPen, frontX, frontY, side, side);
    } else if (iconType === "minusBack") {
        /* 背面型抜き：front を塗り、back の重なりを背景色で抜き、back を中間色の輪郭に */
        fillRect(graphics, iconBrush, frontX, frontY, side, side);
        fillRect(graphics, backgroundBrush, frontX, frontY, overlapWidth, overlapHeight);
        strokeRect(graphics, mutedPen, backX, backY, side, side);
    }
}

/* ============================================================
 * UIレイアウトの共通設定 / Shared UI layout
 * ============================================================ */

/* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
var ICON_PANEL_MARGINS = [8, 16, 8, 8];  /* アイコンボタンのみのパネルは余白を狭く / tighter margins for icon-only panels */
var PANEL_SPACING  = 8;                  /* パネル内の要素間隔 / panel spacing */
var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */

/**
 * ウィンドウの共通設定を適用する / apply shared window layout
 * @param {Window} targetWindow 対象ウィンドウ / target window
 * @param {number} [spacing] 要素間隔（省略時は WINDOW_SPACING）/ spacing override
 * @returns {void}
 */
function setupWindow(targetWindow, spacing) {
    targetWindow.orientation = "column";
    targetWindow.alignChildren = "fill";
    targetWindow.margins = WINDOW_MARGINS;
    targetWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
}

/**
 * パネルの共通設定を適用する / apply shared panel layout
 * @param {Panel} panel 対象パネル / target panel
 * @param {number} [spacing] 要素間隔（省略時は PANEL_SPACING）/ spacing override
 * @returns {void}
 */
function setupPanel(panel, spacing) {
    panel.orientation = "column";
    panel.alignChildren = ["fill", "top"];
    panel.alignment = "fill";
    panel.margins = PANEL_MARGINS;
    panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/* ============================================================
 * パレット UI / Palette UI
 * ============================================================ */

/**
 * onDraw から iconType を束縛したクロージャを返す / bind iconType for onDraw
 * @param {string} iconType 描画種別 / icon draw type
 * @returns {function} onDraw ハンドラ / onDraw handler
 */
function makeIconDrawer(iconType) {
    return function () {
        drawOperationIcon(this, iconType);
    };
}

/**
 * 操作アイコンボタン（46x46・onDraw 描画）を container に追加する。
 * @param {object} container 追加先のパネル/グループ / parent container
 * @param {object} operation icon / labelKey を持つ操作定義 / operation with icon & labelKey
 * @param {function} onClickHandler クリック時の処理 / click handler
 * @returns {object} 追加した iconbutton / the added iconbutton
 */
function addOperationButton(container, operation, onClickHandler) {
    var button = container.add("iconbutton", undefined, undefined, { style: "toolbutton" });
    button.preferredSize = [46, 46];
    button.helpTip = getLocalizedText(operation.labelKey);
    button.onDraw = makeIconDrawer(operation.icon);
    button.onClick = onClickHandler;
    return button;
}

/**
 * モードパネル（出力モードの排他ラジオ）を構築する。
 * @param {Window} parentWindow 親ウィンドウ / parent window
 * @returns {{execute: object, compound: object, effect: object}} ラジオボタン群 / radios
 */
function buildModePanel(parentWindow) {
    var panel = parentWindow.add("panel", undefined, getLocalizedText("panel.mode"));
    setupPanel(panel);
    var radios = {
        execute:  panel.add("radiobutton", undefined, getLocalizedText("apply.execute")),
        compound: panel.add("radiobutton", undefined, getLocalizedText("apply.compound")),
        effect:   panel.add("radiobutton", undefined, getLocalizedText("apply.effect"))
    };
    radios.execute.value = true;
    /* ショートカットは UI ラベルには出さず helpTip に載せる / show shortcuts in helpTip, not in the label */
    radios.execute.helpTip = getLocalizedText("tip.shortcutExecute");
    radios.compound.helpTip = getLocalizedText("tip.compoundApply") + " / " + getLocalizedText("tip.shortcutCompound");
    radios.effect.helpTip = getLocalizedText("tip.shortcutEffect");
    /* ボタン類は広げず左寄せ / keep button-like controls left-aligned */
    radios.execute.alignment = "left";
    radios.compound.alignment = "left";
    radios.effect.alignment = "left";
    return radios;
}

/**
 * 形状モードのアイコンパネル（横1行）を構築する。
 * @param {Window} parentWindow 親ウィンドウ / parent window
 * @returns {object} パネル / the panel
 */
function buildShapeModePanel(parentWindow) {
    var panel = parentWindow.add("panel", undefined, getLocalizedText("panel.shapeMode"));
    panel.orientation = "row";
    panel.alignChildren = "center";
    panel.margins = ICON_PANEL_MARGINS;
    panel.spacing = PANEL_SPACING;
    return panel;
}

/**
 * パスファインダーのアイコン行（3個×2行）を構築する。
 * @param {Window} parentWindow 親ウィンドウ / parent window
 * @returns {object[]} 2つの行グループ / two row groups
 */
function buildPathfinderRows(parentWindow) {
    var panel = parentWindow.add("panel", undefined, getLocalizedText("panel.pathfinder"));
    panel.orientation = "column";
    panel.alignChildren = "center";
    panel.margins = ICON_PANEL_MARGINS;
    panel.spacing = PANEL_SPACING;
    var rows = [panel.add("group"), panel.add("group")];
    for (var rowIndex = 0; rowIndex < rows.length; rowIndex++) {
        rows[rowIndex].orientation = "row";
        rows[rowIndex].alignChildren = "center";
        rows[rowIndex].spacing = PANEL_SPACING;
    }
    return rows;
}

/**
 * オプションパネル（チェックボックス＋拡張ボタン）を構築する。
 * 拡張ボタンは複合シェイプ選択時のみ機能するが、判定はクリック時に worker 側で行うため常に押せる。
 * @param {Window} parentWindow 親ウィンドウ / parent window
 * @returns {{removePoints: object, removeUnpainted: object, expand: object}} 各コントロール / controls
 */
function buildOptionPanel(parentWindow) {
    var panel = parentWindow.add("panel", undefined, getLocalizedText("panel.option"));
    setupPanel(panel);
    var controls = {
        removePoints:    panel.add("checkbox", undefined, getLocalizedText("option.removePoints")),
        removeUnpainted: panel.add("checkbox", undefined, getLocalizedText("option.removeUnpainted")),
        expand:          panel.add("button", undefined, getLocalizedText("button.expand"))
    };
    controls.removePoints.value = true;
    controls.removeUnpainted.value = false;
    controls.removeUnpainted.helpTip = getLocalizedText("tip.removeUnpainted");
    controls.expand.helpTip = getLocalizedText("tip.expand");
    controls.removePoints.alignment = "left";
    controls.removeUnpainted.alignment = "left";
    controls.expand.alignment = "left";
    return controls;
}

/**
 * 下部の状況表示（複数行で折り返し・幅を膨らませない）を構築する。
 * @param {Window} parentWindow 親ウィンドウ / parent window
 * @returns {object} statictext
 */
function buildStatusText(parentWindow) {
    var statusText = parentWindow.add("statictext", undefined, getLocalizedText("status.ready"), { multiline: true });
    statusText.characters = 20;
    statusText.preferredSize.height = 32;
    statusText.helpTip = getLocalizedText("tip.esc");
    return statusText;
}

/**
 * 常駐パレットを表示する。二重起動を回避（既に開いていれば前面化して終了）。
 * @returns {void}
 */
function showPalette() {
    /* 二重起動回避：既存パレットが生きていれば作り直さず前面化して終了 / avoid double launch: reuse existing */
    if ($.global.__pfPaletteWindow) {
        try {
            $.global.__pfPaletteWindow.show();
            try { $.global.__pfPaletteWindow.active = true; } catch (activeError) { }
            return;
        } catch (reuseError) {
            /* 参照が無効なら作り直す / stale reference → recreate */
            $.global.__pfPaletteWindow = null;
        }
    }

    /* 再入防止（BridgeTalk 同期送信中の多重発火を防ぐ）/ re-entrancy guard for this palette session */
    var isBusy = false;

    var paletteWindow = new Window("palette", getLocalizedText("dialog.title") + " " + SCRIPT_VERSION, undefined, { resizeable: false });
    setupWindow(paletteWindow);
    /* コメントエリア（下部ステータス）の下余白を 4 に / bottom margin under the comment area = 4 */
    paletteWindow.margins = [WINDOW_MARGINS, WINDOW_MARGINS, WINDOW_MARGINS, 4];

    /* モードパネル（出力モードの排他ラジオ・最上段）/ Mode panel (output-mode radios, top)
     * A: 実行（実際にパスへ）/ B: 複合シェイプ（上段のみ）/ C: 効果として適用（ライブ）
     */
    var modeRadios = buildModePanel(paletteWindow);
    var modeExecuteRadio = modeRadios.execute;
    var modeCompoundRadio = modeRadios.compound;
    var modeEffectRadio = modeRadios.effect;

    /* 形状モードパネル（アイコンボタンは後段で追加）/ Shape mode panel (buttons added below) */
    var shapeModePanel = buildShapeModePanel(paletteWindow);

    /* パスファインダーパネル（3個×2行, ボタンは後段で追加）/ Pathfinder panel (3 per row × 2, buttons added below) */
    var pathfinderRows = buildPathfinderRows(paletteWindow);

    /* オプションパネル＋コメントエリアをまとめ、両者間の余白をゼロにする / group options + comment area with no gap between them */
    var bottomWrap = paletteWindow.add("group");
    bottomWrap.orientation = "column";
    bottomWrap.alignChildren = "fill";
    bottomWrap.alignment = "fill";
    bottomWrap.margins = 0;
    bottomWrap.spacing = 18;

    /* オプションパネル（チェックボックス＋拡張ボタン）/ Options panel (checkboxes + expand button) */
    var optionControls = buildOptionPanel(bottomWrap);
    var removePointsCheckbox = optionControls.removePoints;
    var removeUnpaintedCheckbox = optionControls.removeUnpainted;
    var expandButton = optionControls.expand;

    /* 状況表示（幅を膨らませないよう複数行で折り返す）/ Status text (multiline, keeps width narrow) */
    var statusText = buildStatusText(bottomWrap);

    /**
     * 状況表示を更新する / update the status line
     * @param {string} message 表示文字列 / message
     * @returns {void}
     */
    function setStatus(message) {
        statusText.text = message;
    }

    /* isBusy ガード付きで委譲を実行し、返り値（status 文字列）を表示する / guarded delegate + status */
    function runExclusive(produceStatus) {
        if (isBusy) { return; }
        isBusy = true;
        try {
            setStatus(produceStatus());
        } catch (delegateError) {
            setStatus(getLocalizedText("status.error") + delegateError);
        } finally {
            isBusy = false;
        }
    }

    /* 形状モード（上段4つ）クリックで即適用する / apply a Shape Mode on click
     * B 複合シェイプ → ダイナミックアクション（複合シェイプのまま）
     * A 実行 / C 効果 → XML ライブ効果（value は Adobe Pathfinder の Command 番号と一致）。A は拡張＋グループ解除、C はライブのまま
     * ※ ダイナミックアクションの複合シェイプは expandStyle で綺麗に焼き込めないため A も XML を使う
     */
    function makeApplyHandler(mode) {
        return function () {
            /* onClick は event を持たないため、直前の mousedown で記録した Option 状態を読む
             * onClick carries no event; read the Option state recorded by the preceding mousedown */
            var withOption = (this.__altPressed === true);
            this.__altPressed = false;
            runExclusive(function () {
                if (modeCompoundRadio.value || withOption) {
                    /* B: ai_compound_shape の enumerated 値を渡す / pass the compound-shape enumerated value */
                    return markerToStatus(delegateApply(mode.compoundValue, mode.name, true), mode);
                }
                var destructive = !modeEffectRadio.value;
                /* A/C: Pathfinder XML の Command 番号を渡す / pass the Pathfinder XML Command index */
                var marker = delegatePathfinder(mode.pathfinderCommand, false, removePointsCheckbox.value, destructive);
                return markerToStatus(marker, mode);
            });
        };
    }

    /* mousedown で Option（alt）状態をボタンに記録する（onClick は event を持たない）
     * record the Option (alt) state on mousedown; onClick has no event to read it from */
    function makeAltRecorder(button) {
        return function (mouseEvent) {
            button.__altPressed = (mouseEvent.altKey === true);
        };
    }

    for (var i = 0; i < SHAPE_MODES.length; i++) {
        var mode = SHAPE_MODES[i];
        var shapeModeButton = addOperationButton(shapeModePanel, mode, makeApplyHandler(mode));
        /* Option+クリックのヒントと mousedown での alt 記録を上乗せ / add Option-click hint & alt recorder */
        shapeModeButton.helpTip = getLocalizedText(mode.labelKey) + " / " + getLocalizedText("tip.optionCompound");
        shapeModeButton.addEventListener("mousedown", makeAltRecorder(shapeModeButton));
    }

    /* クリックで該当パスファインダーを即適用する / apply the Pathfinder immediately on click
     * C 効果 → ライブ効果のまま / A 実行 → 拡張＋グループ解除で実際にパスへ
     * ※ B（複合シェイプ）選択時は下段が無効化されクリックできない。removeUnpainted は分割・アウトラインのみ有効
     */
    function makePathfinderHandler(pathfinder) {
        return function () {
            runExclusive(function () {
                var destructive = !modeEffectRadio.value;
                var shouldRemoveUnpainted = pathfinder.unpainted && removeUnpaintedCheckbox.value;
                var marker = delegatePathfinder(pathfinder.command, shouldRemoveUnpainted, removePointsCheckbox.value, destructive);
                return markerToStatus(marker, pathfinder);
            });
        };
    }

    var pathfinderButtons = [];
    for (var pathfinderIndex = 0; pathfinderIndex < PATHFINDER_MODES.length; pathfinderIndex++) {
        var pathfinder = PATHFINDER_MODES[pathfinderIndex];
        var pathfinderRow = pathfinderRows[Math.floor(pathfinderIndex / 3)]; /* 3個ごとに改行 / 3 per row */
        var pathfinderButton = addOperationButton(pathfinderRow, pathfinder, makePathfinderHandler(pathfinder));
        pathfinderButtons.push(pathfinderButton);
    }

    /* B（複合シェイプ）は形状モード専用 → 選択時は下段パスファインダーを無効化 / disable Pathfinders under mode B */
    function updatePathfinderEnabled() {
        var enabled = !modeCompoundRadio.value;
        for (var buttonIndex = 0; buttonIndex < pathfinderButtons.length; buttonIndex++) {
            pathfinderButtons[buttonIndex].enabled = enabled;
        }
    }
    modeExecuteRadio.onClick = updatePathfinderEnabled;
    modeCompoundRadio.onClick = updatePathfinderEnabled;
    modeEffectRadio.onClick = updatePathfinderEnabled;
    updatePathfinderEnabled();

    /* 拡張ボタンのクリックで選択中の複合シェイプを拡張する（判定はクリック時に worker 側で実施）
     * パレットは選択変更を通知できず enabled のキャッシュは古くなるため、常に押せるようにして
     * クリック時に複合シェイプの有無を判定する（無ければ status に案内を出す）
     * expand the compound shape on click; validate the selection at click time (palettes get no
     * selection-change events, so a cached enabled state goes stale — keep it always clickable) */
    expandButton.onClick = function () {
        runExclusive(function () {
            return markerToStatus(delegateExpandCompoundShape(), { labelKey: "button.expand" });
        });
    };

    /* 出力モードを選択し、パスファインダーの有効状態を更新する / select an output mode and refresh Pathfinder state */
    function selectOutputMode(targetRadio) {
        modeExecuteRadio.value = (targetRadio === modeExecuteRadio);
        modeCompoundRadio.value = (targetRadio === modeCompoundRadio);
        modeEffectRadio.value = (targetRadio === modeEffectRadio);
        updatePathfinderEnabled();
    }

    /* キーボードショートカット / Keyboard shortcuts
     * Esc: 閉じる / P: パスファインダーを実行 / C: 複合シェイプにする / F: 効果として適用
     */
    paletteWindow.addEventListener("keydown", function (kbEvent) {
        if (kbEvent.keyName === "Escape") {
            paletteWindow.close();
        } else if (kbEvent.keyName === "P") {
            selectOutputMode(modeExecuteRadio);
            kbEvent.preventDefault();
        } else if (kbEvent.keyName === "C") {
            selectOutputMode(modeCompoundRadio);
            kbEvent.preventDefault();
        } else if (kbEvent.keyName === "F") {
            selectOutputMode(modeEffectRadio);
            kbEvent.preventDefault();
        }
    });

    /* 閉じたら常駐参照をクリア / clear the persistent reference on close */
    paletteWindow.onClose = function () {
        $.global.__pfPaletteWindow = null;
        return true;
    };

    /* 常駐エンジンに参照を保持して GC を回避 / keep the reference alive to avoid GC */
    $.global.__pfPaletteWindow = paletteWindow;

    paletteWindow.show();
}

showPalette();

})();
