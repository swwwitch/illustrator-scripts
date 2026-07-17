#target illustrator
#targetengine "pathfinder-palette"
app.preferences.setBooleanPreference("ShowExternalJSXWarning", false);

/*
AiSmartPathfinder.jsx — Smart Pathfinder Palette

選択した複数オブジェクトにパスファインダーを適用する常駐パレット。
アイコンをクリックすると、その操作をメインエンジンへ委譲して即時に実行する。

タブ構成: 「基本」／「その他」（EN: Special）の2タブ（tabbedpanel）

「基本」タブ（上から）:
  モード         出力モードを排他ラジオで選択（下記 A/B/C。ショートカット P/C/F）
  形状モード      合体／前面型抜き／交差／中マド（Adobe Pathfinder command 0〜3）
  パスファインダー 分割／刈り込み／合流／切り抜き／アウトライン／背面型抜き（3個×2行, command 4〜9）
  オプション      余分なポイントを削除（RemovePoints, 右に［強制］ボタン）／塗りのないアートワークを削除（ExtractUnpainted, 分割・アウトライン かつ実行モードのみ）／［複合シェイプを拡張］ボタン

「その他」タブ（4パネル構成）:
  マド埋め      ［実行］（複合パス解除＋合体→実パス化）／［効果］（ライブ効果のまま）
  変換          ［線を塗りに変換］（線のアウトライン化・ライブ効果）
  アピアランス  縦並び。［分割］（expandStyle＋可能ならグループ解除）／［効果のみを消去］（消去→塗り・線を復元）／［（完全に）消去］（アピアランスを消去）
  ツール、パネルを表示 ［アピアランス］（Style Palette）／［パスファインダー］（Adobe PathfinderUI）／［シェイプ形成ツール］（selectTool）

  ※ ［強制］は選択パスの直線上の冗長アンカーを削除（PathCleanupTool 相当・許容誤差0.02固定）
  ※ マド埋め［実行］は PathCleanupTool の fillHolesOnSelection と同じ手順（group→noCompoundPath→Live Pathfinder Add→expandStyle→ungroup）
  ※ パスファインダーの6アイコン（2行×3）は unifyIconCellWidths でセル幅を統一し列を揃える

出力モード（モードパネルの排他ラジオ）:
  A パスに変換          上段・下段とも＝グループ化→XML（Adobe Pathfinder）→拡張→グループ解除（実際にパスへ変換）
  B 複合シェイプを作成  上段のみ＝ダイナミックアクション（ai_compound_shape）で複合シェイプ（下段6ボタンはディム＝無効）
  C 効果として適用      上段・下段とも＝XML ライブ効果のまま（拡張しない）
  ※ ダイナミックアクションの複合シェイプは expandStyle で焼き込めないため、A の上段も XML 方式を使う
     （ai_plugin_pathfinder のパネル実行アクションも検討したが、undo が2回必要になるため不採用）
  ※ A/C で複数オブジェクトを選択している場合は、効果適用の前に一時的にグループ化する（1つの効果対象にまとめるため）。
     A では拡張後にそのグループを解除して実パスへ戻す。C はライブ効果のためグループのまま残る。

形状モード（上段4ボタン）の Option+クリック:
  出力モードに関係なく複合シェイプを作成する（＝モード B と同じ。拡張はしない）。

パスファインダー（下段6ボタン）の Option+クリック:
  出力モードに関係なく効果として適用する（＝モード C と同じ。removeUnpainted は強制OFF）。

［複合シェイプを拡張］ボタン（オプションパネル）:
  選択中の複合シェイプ（DOM 上 PluginItem）を通常のパスへ拡張する（ai_expand_compound_shape）。
  Option（Alt）+クリックのときは拡張ではなく解除する（ai_release_compound_shape）。
  パレットは選択変更を受け取れないため常に押せ、複合シェイプの有無はクリック時に判定する（無ければ "NOCS"）。

その他タブのボタン:
  マド埋め［実行］    選択を複合パス解除→ライブパスファインダー合体→拡張→グループ解除（PathCleanupTool の fillHolesOnSelection 相当）。
  マド埋め［効果］    選択を複合パス解除→ライブパスファインダー合体（ライブ効果のまま）。
  ［線を塗りに変換］  線をアウトライン化して1つの塗りにまとめる（ライブ効果）。
  アピアランス［分割］          選択のアピアランス（ライブ効果）を expandStyle で実体化し、可能ならグループ解除する。
  アピアランス［効果のみを消去］ 「アピアランスを消去」を対象単位で再生したうえで、元の塗り・線（テキストは文字塗り）だけを復元する。
                                見た目上はライブ効果だけが消える（ClearAppearance.jsx の「復元する」既定と同等の固定オプション）。
  アピアランス［（完全に）消去］ 「アピアランスを消去」ダイナミックアクション（ai_plugin_appearance / key 1835363957 / value 6）を
                                選択全体に一括再生する。塗り・線などの復元は行わない（アピアランスパネルのメニュー相当）。
  ツール、パネルを表示［アピアランス］     Illustrator のアピアランスパネル（Style Palette）を表示する。
  ツール、パネルを表示［パスファインダー］ Illustrator のパスファインダーパネル（Adobe PathfinderUI）を表示する。
  ツール、パネルを表示［シェイプ形成ツール］ シェイプ形成ツール（Adobe Shape Builder Tool）に切り替える（ドキュメント必須）。

［強制］ボタン（オプションパネル・「余分なポイントを削除」の右）:
  選択パス（グループ・複合パス含む）の直線上の冗長なアンカーポイントを削除する（許容誤差 0.02 固定・2パス）。

Persistent palette that applies Pathfinder operations to the current selection.
Two tabs: "Basic" and "Special". Basic tab, top-down: Mode (A/B/C radios;
shortcuts P/C/F), Shape Mode (Unite/Minus Front/Intersect/Exclude), Pathfinders
(Divide/Trim/Merge/Crop/Outline/Minus Back), Options (Remove points [+ Force
button] / Remove unpainted / Expand Compound Shape button). Special tab:
grouped panels — Fill Holes (Apply / Effect), Convert (Convert Strokes to
Fills), Appearance (Expand / Clear Effects / Clear; stacked vertically), Show
Panel (Appearance / Pathfinder). Clear plays the ai_plugin_appearance "Clear
Appearance" dynamic action on the whole selection (no attribute restoration);
Clear Effects clears each item's appearance and then restores its original
fill/stroke (per-character fill for text) so only the live effects disappear.
A = Apply: both rows group, apply the Adobe Pathfinder XML, expand, and ungroup
(bake to real paths). B = Compound shape: Shape Modes only via the
ai_compound_shape action (Pathfinders dimmed/disabled). C = Apply as effect:
both rows keep the live Adobe Pathfinder effect. "Remove unpainted artwork"
affects Divide / Outline only, and only in the destructive Apply mode (it is
forced off in Apply-as-effect mode). Option-clicking a Shape Mode button always makes a
compound shape regardless of the output mode; Option-clicking a Pathfinder
button always applies it as a live effect. The Expand Compound Shape button
turns a selected compound shape into plain paths (Option-click releases it
instead). The Force button removes redundant collinear anchor points from the
selected paths. The Show Panel buttons show the Appearance / Pathfinder panels.

構成 / Structure
- 常駐エンジン（#targetengine）でパレット参照を保持し GC を回避
- 表示中の常駐 app は DOM 接続を失うため、DOM 操作は BridgeTalk でメインエンジンへ委譲
- 委譲は worker 関数を toString → encodeURIComponent → eval(decodeURIComponent(...)) で送信
- worker の toString はコメントを取り込み壊すため、送信前に stripWorkerComments で除去（worker 本体にはコメントを書かず、説明は JSDoc に集約する）
- 戻り値はマーカー（OK / NODOC / NOSEL / NEEDTWO / NOCS / ERR:...）で受けるが、ステータス表示エリアは持たないため markerToStatus の戻り値は破棄（委譲の副作用のみ利用）
- 選択不足は条件で分離：効果は1つ以上（NOSEL）、実行・複合シェイプは2つ以上（NEEDTWO）
- 複数選択時は効果を1つの対象にまとめるため一時的にグループ化し、A（destructive）では拡張後に解除する（エラー時はその一時グループだけを選択し直して解除）
- UI は tabbedpanel（基本／その他）配下に buildModePanel / buildShapeModePanel / buildPathfinderRows / buildOptionPanel と addOperationButton で構築し、setupWindow / setupPanel と PANEL_MARGINS 等の共通変数で統一。その他タブはマド埋め／変換／アピアランス／ツール、パネルを表示の4パネル（setupPanel 共用。マド埋めは row で横並び、アピアランス・ツール、パネルを表示は縦並び）
- ［ツール、パネルを表示：アピアランス／パスファインダー］は app.executeMenuCommand（Style Palette / Adobe PathfinderUI）を、［シェイプ形成ツール］は app.selectTool（Adobe Shape Builder Tool）をメインエンジンへ委譲
- アイコンは onDraw で描画（無効時は dimIconColors でディム表示）
- Option（alt）状態は onDraw では取れないため mousedown で記録し onClick で読む（形状モードの複合シェイプ化・下段の効果適用・拡張ボタンの解除に共通）
- 全体を IIFE で閉じ、$.global にはパレット参照（__pfPaletteWindow）だけを残す
- 二重起動回避：既に開いていれば作り直さず前面化して終了
- キーボード：Esc で閉じる／P・C・F で出力モード切替

### 紹介記事(note)

https://note.com/dtp_tranist/n/n6909b836221a

*/

// =========================================
// バージョン / Version
// =========================================

var SCRIPT_VERSION = "v1.1.0";

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
    tab: {
        basic:   { ja: "基本",   en: "Basic" },
        special: { ja: "その他", en: "Special" }
    },
    panel: {
        shapeMode:  { ja: "形状モード",     en: "Shape Mode" },
        pathfinder: { ja: "パスファインダー", en: "Pathfinders" },
        mode:       { ja: "モード",         en: "Mode" },
        option:     { ja: "オプション",     en: "Options" },
        fillHoles:  { ja: "マド埋め",       en: "Fill Holes" },
        convert:    { ja: "変換",           en: "Convert" },
        appearance: { ja: "アピアランス",   en: "Appearance" },
        showPanel:  { ja: "ツール、パネルを表示", en: "Show Tools & Panels" }
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
    caption: {
        unite:      { ja: "合体",         en: "Unite" },
        minusFront: { ja: "前面型抜き",   en: "Minus Front" },
        intersect:  { ja: "交差",         en: "Intersect" },
        exclude:    { ja: "中マド",       en: "Exclude" },
        divide:     { ja: "分割",         en: "Divide" },
        trim:       { ja: "刈り込み",     en: "Trim" },
        merge:      { ja: "合流",         en: "Merge" },
        crop:       { ja: "切り抜き",     en: "Crop" },
        outline:    { ja: "アウトライン", en: "Outline" },
        minusBack:  { ja: "背面型抜き",   en: "Minus Back" }
    },
    apply: {
        execute:  { ja: "パスに変換",             en: "Convert to Paths" },
        compound: { ja: "複合シェイプを作成",     en: "Compound shape" },
        effect:   { ja: "効果として適用",         en: "Apply as effect" }
    },
    button: {
        expand:        { ja: "複合シェイプを拡張", en: "Expand Compound Shape" },
        expandRelease: { ja: "複合シェイプを解除", en: "Release Compound Shape" },
        release:       { ja: "解除", en: "Release" },
        appearance:      { ja: "アピアランス",       en: "Appearance" },
        pathfinderPanel: { ja: "パスファインダー",   en: "Pathfinder" },
        shapeBuilder:    { ja: "シェイプ形成ツール", en: "Shape Builder Tool" },
        cleanup:         { ja: "強制",               en: "Force" },
        fillHolesExpand: { ja: "実行",               en: "Apply" },
        fillHolesEffect: { ja: "効果",               en: "Effect" },
        expandAppearance:{ ja: "分割",               en: "Expand" },
        clearEffectsOnly:{ ja: "効果のみを消去",     en: "Clear Effects" },
        clearAppearance: { ja: "（完全に）消去",       en: "Clear" },
        strokeToFill:    { ja: "線を塗りに変換",     en: "Convert Strokes to Fills" }
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
        removeUnpainted: { ja: "分割・アウトラインのみ有効（効果として適用のときはOFF）", en: "Divide / Outline only (off in Apply-as-effect mode)" },
        compoundApply:   { ja: "形状モード（上段）のみ",    en: "Shape Mode (top row) only" },
        optionCompound:  { ja: "Option+クリックで複合シェイプ", en: "Option-click to make a compound shape" },
        expand:          { ja: "選択中の複合シェイプを通常のパスに拡張", en: "Expand the selected compound shape to paths" },
        release:         { ja: "選択中の複合シェイプを解除", en: "Release the selected compound shape" },
        optionRelease:   { ja: "Option+クリックで解除", en: "Option-click to release" },
        cleanup:         { ja: "直線上の冗長なアンカーポイントを削除（許容誤差 0.02）", en: "Remove redundant collinear anchor points (tolerance 0.02)" },
        appearance:      { ja: "アピアランスパネルを表示", en: "Show the Appearance panel" },
        pathfinderPanel: { ja: "パスファインダーパネルを表示", en: "Show the Pathfinder panel" },
        shapeBuilder:    { ja: "シェイプ形成ツールに切り替える", en: "Switch to the Shape Builder tool" },
        fillHolesExpand: { ja: "複合パスを解除して合体し、拡張して実パスにする（マド埋め）", en: "Fill holes and expand to real paths" },
        fillHolesEffect: { ja: "複合パスを解除して合体（ライブ効果のまま／マド埋め）", en: "Fill holes, keep as a live effect" },
        expandAppearance:{ ja: "選択オブジェクトのアピアランスを実体化（可能ならグループ解除）", en: "Expand the appearance of the selection (ungroup if possible)" },
        clearEffectsOnly:{ ja: "アピアランスを消去し、元の塗り・線（テキストは文字塗り）だけを戻す（＝効果のみ消去）", en: "Clear appearance but keep the original fill/stroke (removes effects only)" },
        clearAppearance: { ja: "選択オブジェクトのアピアランスを消去", en: "Clear the appearance of the selection" },
        optionEffect:    { ja: "Option+クリックで効果として適用", en: "Option-click to apply as a live effect" },
        strokeToFill:    { ja: "線をアウトライン化して1つの塗りにまとめる（ライブ効果）", en: "Outline strokes and merge into one fill (live effect)" },
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

/**
 * 選択中の複合シェイプを解除する（メインエンジン用エントリ）。
 * クリック時に選択を判定し、複合シェイプ（DOM 上 PluginItem）が無ければ "NOCS" を返す。
 * 複合シェイプがあればダイナミックアクション ai_release_compound_shape を一時アクションとして再生する。
 * アクション構造は拡張と同じ integer パラメータ1個のため buildExpandActionSource を共用する。
 * @returns {string} マーカー "OK" / "NODOC" / "NOCS"（複合シェイプ未選択）/ "ERR:..."
 */
function workerReleaseCompoundShape() {
    if (app.documents.length === 0) { return "NODOC"; }
    var currentSelection = app.activeDocument.selection;
    if (!currentSelection || currentSelection.length < 1) { return "NOCS"; }
    var hasCompoundShape = false;
    for (var selectionIndex = 0; selectionIndex < currentSelection.length; selectionIndex++) {
        if (currentSelection[selectionIndex].typename === "PluginItem") { hasCompoundShape = true; break; }
    }
    if (!hasCompoundShape) { return "NOCS"; }
    var uniqueToken = "AiSmartPathfinder_release_"
        + (new Date()).getTime()
        + "_"
        + Math.floor(Math.random() * 100000);
    var actionConfig = {
        setName: uniqueToken + "_set",
        actionName: uniqueToken + "_action",
        internalName: "ai_release_compound_shape",
        localizedName: "複合シェイプを解除",
        expandParamKey: 1919710053,
        actionFilePath: Folder.temp.fsName + "/" + uniqueToken + ".aia"
    };
    try {
        var actionSource = buildExpandActionSource(actionConfig);
        playTemporaryAction(actionSource, actionConfig.setName, actionConfig.actionName, actionConfig.actionFilePath);
        app.redraw();
        return "OK";
    } catch (releaseError) {
        return "ERR:" + releaseError;
    }
}

/**
 * アピアランスパネルを表示する（メインエンジン用エントリ）。
 * ドキュメント不要のメニューコマンドでパネルの表示をトグルする。
 * @returns {string} マーカー "OK" / "ERR:..."
 */
function workerShowAppearancePanel() {
    try {
        app.executeMenuCommand("Style Palette");
        return "OK";
    } catch (appearanceError) {
        return "ERR:" + appearanceError;
    }
}

/**
 * マド埋め：選択を複合パス解除→ライブパスファインダー（合体）でマドを埋める（メインエンジン用エントリ）。
 * expand が true のときは expandStyle で実体化してグループ解除する（実パスへ／PathCleanupTool の fillHolesOnSelection 相当）。
 * expand が false のときはライブ効果（アピアランス）のまま残す（グループのまま）。単一で ungroup が失敗しても握りつぶす。
 * @param {boolean} expand 拡張して実パスにするか（false ならライブ効果のまま）
 * @returns {string} マーカー "OK" / "NODOC" / "NOSEL" / "ERR:..."
 */
function workerFillHoles(expand) {
    if (app.documents.length === 0) { return "NODOC"; }
    var currentSelection = app.activeDocument.selection;
    if (!currentSelection || currentSelection.length < 1) { return "NOSEL"; }
    try {
        app.executeMenuCommand("group");
        app.executeMenuCommand("noCompoundPath");
        app.executeMenuCommand("Live Pathfinder Add");
        if (expand) {
            app.executeMenuCommand("expandStyle");
            try {
                app.executeMenuCommand("ungroup");
            } catch (ungroupError) { }
        }
        app.redraw();
        return "OK";
    } catch (fillHolesError) {
        return "ERR:" + fillHolesError;
    }
}

/**
 * アピアランスを分割：選択オブジェクトのアピアランス（ライブ効果）を実体化する（メインエンジン用エントリ）。
 * 分割結果は通常グループになるため、続けてグループ解除を試みる（解除できない選択でも失敗は握りつぶす）。
 * @returns {string} マーカー "OK" / "NODOC" / "NOSEL" / "ERR:..."
 */
function workerExpandAppearance() {
    if (app.documents.length === 0) { return "NODOC"; }
    var currentSelection = app.activeDocument.selection;
    if (!currentSelection || currentSelection.length < 1) { return "NOSEL"; }
    try {
        app.executeMenuCommand("expandStyle");
        try {
            app.executeMenuCommand("ungroup");
        } catch (ungroupError) { }
        app.redraw();
        return "OK";
    } catch (expandAppearanceError) {
        return "ERR:" + expandAppearanceError;
    }
}

/**
 * 線を塗りに：ライブ効果で線をアウトライン化し、合流→合体で1つの塗りにまとめる（メインエンジン用エントリ）。
 * Live Outline Stroke → Live Pathfinder Merge → Live Pathfinder Add（いずれもライブ効果）。
 * @returns {string} マーカー "OK" / "NODOC" / "NOSEL" / "ERR:..."
 */
function workerStrokeToFill() {
    if (app.documents.length === 0) { return "NODOC"; }
    var currentSelection = app.activeDocument.selection;
    if (!currentSelection || currentSelection.length < 1) { return "NOSEL"; }
    try {
        app.executeMenuCommand("Live Outline Stroke");
        app.executeMenuCommand("Live Pathfinder Merge");
        app.executeMenuCommand("Live Pathfinder Add");
        app.redraw();
        return "OK";
    } catch (strokeToFillError) {
        return "ERR:" + strokeToFillError;
    }
}

/**
 * 選択パス（グループ・複合パス含む）から直線上の冗長なアンカーポイントを削除する（メインエンジン用エントリ）。
 * PathCleanupTool.jsx の「直線状のアンカーポイント」削除を許容誤差 0.02 固定で実行する。
 * @returns {string} マーカー "OK" / "NODOC" / "NOSEL" / "ERR:..."
 */
function workerCleanupCollinear() {
    if (app.documents.length === 0) { return "NODOC"; }
    var currentSelection = app.activeDocument.selection;
    if (!currentSelection || currentSelection.length < 1) { return "NOSEL"; }
    try {
        var targets = [];
        for (var i = 0; i < currentSelection.length; i++) {
            collectCleanupPathItems(currentSelection[i], targets);
        }
        removeRedundantAnchorsCollinear(targets, 0.02);
        removeRedundantAnchorsCollinear(targets, 0.02);
        app.redraw();
        return "OK";
    } catch (cleanupError) {
        return "ERR:" + cleanupError;
    }
}

/**
 * ロック・非表示（親・レイヤー含む）を判定する。処理対象から除外するため。
 * @param {object} item PageItem / Layer
 * @returns {boolean} スキップ対象なら true
 */
function isCleanupSkippable(item) {
    var currentItem = item;
    while (currentItem) {
        try {
            if (currentItem.locked === true) { return true; }
            if (currentItem.hidden === true) { return true; }
            if (currentItem.typename === "Layer") {
                if (currentItem.locked === true) { return true; }
                if (currentItem.visible === false) { return true; }
            }
            if (currentItem.layer) {
                try {
                    if (currentItem.layer.locked === true) { return true; }
                    if (currentItem.layer.visible === false) { return true; }
                } catch (layerError) { }
            }
        } catch (accessError) { }
        try {
            currentItem = currentItem.parent;
        } catch (parentError) {
            break;
        }
        if (!currentItem || currentItem.typename === "Document") { break; }
    }
    return false;
}

/**
 * 選択項目から PathItem を再帰的に収集する（GroupItem・CompoundPathItem を展開）。
 * @param {object} item 対象項目
 * @param {object[]} pathItems 収集先の配列
 * @returns {void}
 */
function collectCleanupPathItems(item, pathItems) {
    if (!item) { return; }
    if (isCleanupSkippable(item)) { return; }
    try {
        if (item.typename === "PathItem") {
            pathItems.push(item);
            return;
        }
        if (item.typename === "CompoundPathItem") {
            for (var ci = 0; ci < item.pathItems.length; ci++) {
                if (!isCleanupSkippable(item.pathItems[ci])) { pathItems.push(item.pathItems[ci]); }
            }
            return;
        }
        if (item.typename === "GroupItem") {
            for (var gi = 0; gi < item.pageItems.length; gi++) {
                collectCleanupPathItems(item.pageItems[gi], pathItems);
            }
            return;
        }
        if (item.pageItems && item.pageItems.length) {
            for (var pi = 0; pi < item.pageItems.length; pi++) {
                collectCleanupPathItems(item.pageItems[pi], pathItems);
            }
        }
    } catch (collectError) { }
}

/**
 * 3点が一直線上にあるか（外積の絶対値が許容誤差未満か）を判定する。
 * @param {number[]} pointA 端点1 [x,y]
 * @param {number[]} pointB 中間点 [x,y]
 * @param {number[]} pointC 端点2 [x,y]
 * @param {number} tolerance 許容誤差
 * @returns {boolean} 一直線上なら true
 */
function isCleanupCollinear(pointA, pointB, pointC, tolerance) {
    var area = (pointB[0] - pointA[0]) * (pointC[1] - pointA[1]) - (pointB[1] - pointA[1]) * (pointC[0] - pointA[0]);
    return Math.abs(area) < tolerance;
}

/**
 * 直線上の冗長なアンカーポイント（ハンドルなし・前後アンカーと一直線）を削除する。
 * オープンパスの端点は削除しない。削除でのインデックスずれを避けるため後ろから走査する。
 * @param {object[]} targets PathItem 配列
 * @param {number} tolerance 許容誤差（0.02）
 * @returns {number} 削除したアンカー数
 */
function removeRedundantAnchorsCollinear(targets, tolerance) {
    if (!targets || !targets.length) { return 0; }
    var removedCount = 0;
    for (var s = 0; s < targets.length; s++) {
        var item = targets[s];
        if (!item || isCleanupSkippable(item)) { continue; }
        try {
            var pts = item.pathPoints;
            var isClosed = item.closed;
            var startIndex = isClosed ? (pts.length - 1) : (pts.length - 2);
            var endIndex = isClosed ? 0 : 1;
            for (var i = startIndex; i >= endIndex; i--) {
                var currentLen = pts.length;
                if (currentLen < 3) { break; }
                if (i > currentLen - 1) { i = currentLen - 1; }
                if (i < endIndex) { break; }
                var prevIndex = (i - 1 + currentLen) % currentLen;
                var nextIndex = (i + 1) % currentLen;
                var pA = pts[prevIndex];
                var pB = pts[i];
                var pC = pts[nextIndex];
                var straightLeft = (Math.abs(pB.anchor[0] - pB.leftDirection[0]) + Math.abs(pB.anchor[1] - pB.leftDirection[1])) < tolerance;
                var straightRight = (Math.abs(pB.anchor[0] - pB.rightDirection[0]) + Math.abs(pB.anchor[1] - pB.rightDirection[1])) < tolerance;
                if (straightLeft && straightRight && isCleanupCollinear(pA.anchor, pB.anchor, pC.anchor, tolerance)) {
                    pB.remove();
                    removedCount++;
                }
            }
        } catch (anchorError) { }
    }
    return removedCount;
}

/**
 * パスファインダーパネルを表示する（メインエンジン用エントリ）。
 * ドキュメント不要のメニューコマンドでパネルの表示をトグルする。
 * @returns {string} マーカー "OK" / "ERR:..."
 */
function workerShowPathfinderPanel() {
    try {
        app.executeMenuCommand("Adobe PathfinderUI");
        return "OK";
    } catch (pathfinderPanelError) {
        return "ERR:" + pathfinderPanelError;
    }
}

/**
 * シェイプ形成ツールに切り替える（メインエンジン用エントリ）。
 * ツールの選択にはドキュメントが必要なため、未オープンなら NODOC を返す。
 * @returns {string} マーカー "OK" / "NODOC" / "ERR:..."
 */
function workerSelectShapeBuilderTool() {
    try {
        if (app.documents.length === 0) return "NODOC";
        app.selectTool("Adobe Shape Builder Tool");
        return "OK";
    } catch (shapeBuilderError) {
        return "ERR:" + shapeBuilderError;
    }
}

/**
 * 単一の enumerated パラメータを持つ一時アクションソースを生成する（アピアランスの消去などに使用）。
 * @param {object} actionConfig setName / actionName / internalName / localizedName / enumKey / enumName / enumValue を持つ設定
 * @returns {string} .aia のソース文字列 / .aia source text
 */
function buildEnumeratedActionSource(actionConfig) {
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
        + ' /parameterCount 1\n'
        + ' /parameter-1 {\n'
        + ' /key ' + actionConfig.enumKey + '\n'
        + ' /showInPalette 4294967295\n'
        + ' /type (enumerated)\n'
        + buildNameLine(' /name', actionConfig.enumName)
        + ' /value ' + actionConfig.enumValue + '\n'
        + ' }\n'
        + ' }\n'
        + '}\n';
}

/**
 * アピアランスを解除：選択オブジェクトに「アピアランスを消去」ダイナミックアクションを再生する（メインエンジン用エントリ）。
 * ai_plugin_appearance（key 1835363957 / enumerated「アピアランスを消去」/ value 6）を一時アクションとして選択全体に一括再生する。
 * 記録済み .aia と同一パラメータ（セット名のみユニーク化して既存アクションセットとの衝突を回避）。塗り・線などの復元は行わない。
 * @returns {string} マーカー "OK" / "NODOC" / "NOSEL" / "ERR:..."
 */
function workerClearAppearance() {
    if (app.documents.length === 0) { return "NODOC"; }
    var currentSelection = app.activeDocument.selection;
    if (!currentSelection || currentSelection.length < 1) { return "NOSEL"; }
    var uniqueToken = "AiSmartPathfinder_clearapp_"
        + (new Date()).getTime()
        + "_"
        + Math.floor(Math.random() * 100000);
    var actionConfig = {
        setName: uniqueToken + "_set",
        actionName: uniqueToken + "_action",
        internalName: "ai_plugin_appearance",
        localizedName: "アピアランス",
        enumKey: 1835363957,
        enumName: "アピアランスを消去",
        enumValue: 6,
        actionFilePath: Folder.temp.fsName + "/" + uniqueToken + ".aia"
    };
    try {
        var actionSource = buildEnumeratedActionSource(actionConfig);
        playTemporaryAction(actionSource, actionConfig.setName, actionConfig.actionName, actionConfig.actionFilePath);
        app.redraw();
        return "OK";
    } catch (clearAppearanceError) {
        return "ERR:" + clearAppearanceError;
    }
}

/**
 * 「アピアランスを消去」ダイナミックアクション（ai_plugin_appearance / key 1835363957 / value 6）を
 * 現在の選択に対して一時アクションとして1回再生する（呼び出し前に対象を選択しておくこと）。
 * ClearAppearance.jsx の act_Clear 相当。セット名のみユニーク化して既存アクションセットとの衝突を避ける。
 * @returns {void}
 */
function playClearAppearanceAction() {
    var uniqueToken = "AiSmartPathfinder_cleareffects_"
        + (new Date()).getTime()
        + "_"
        + Math.floor(Math.random() * 100000);
    var actionConfig = {
        setName: uniqueToken + "_set",
        actionName: uniqueToken + "_action",
        internalName: "ai_plugin_appearance",
        localizedName: "アピアランス",
        enumKey: 1835363957,
        enumName: "アピアランスを消去",
        enumValue: 6,
        actionFilePath: Folder.temp.fsName + "/" + uniqueToken + ".aia"
    };
    var actionSource = buildEnumeratedActionSource(actionConfig);
    playTemporaryAction(actionSource, actionConfig.setName, actionConfig.actionName, actionConfig.actionFilePath);
}

/**
 * 効果のみを消去（メインエンジン用エントリ）。
 * 選択オブジェクトのアピアランスを消去したうえで、元の塗り・線（テキストは文字塗り）だけを再適用する。
 * 結果として見た目上はライブ効果だけが消える（ClearAppearance.jsx の「復元する」既定と同等の固定オプション）。
 * 消去はオブジェクト単位で行うため、いったん個別選択→消去→復元を繰り返し、最後に元の選択へ戻す。
 * @returns {string} マーカー "OK" / "NODOC" / "NOSEL" / "ERR:..."
 */
function workerClearEffectsOnly() {
    if (app.documents.length === 0) { return "NODOC"; }
    var currentDocument = app.activeDocument;
    var currentSelection = currentDocument.selection;
    if (!currentSelection || currentSelection.length < 1) { return "NOSEL"; }
    var restoreOptions = {
        fillStroke: true,
        textFillFirst: false,
        textFillPerChar: true,
        strokeSettings: true,
        opacity: true,
        blendingMode: true,
        overprint: true
    };
    try {
        var originalSelection = [];
        for (var i = 0; i < currentSelection.length; i++) { originalSelection.push(currentSelection[i]); }
        processClearEffectsItems(originalSelection, restoreOptions);
        currentDocument.selection = null;
        for (var r = 0; r < originalSelection.length; r++) {
            try { if (originalSelection[r]) { originalSelection[r].selected = true; } } catch (reselectError) { }
        }
        app.redraw();
        return "OK";
    } catch (clearEffectsError) {
        return "ERR:" + clearEffectsError;
    }
}

/**
 * 選択項目を再帰的にたどり、種類ごとに効果のみ消去を適用する。
 * グループ（クリップ以外）は展開して再帰、クリップグループ・複合パスは消去のみ、パスは塗り・線を復元、テキストは文字塗りを復元。
 * @param {object[]} items 対象項目配列 / target items
 * @param {object} restoreOptions 復元オプション / restore options
 * @returns {void}
 */
function processClearEffectsItems(items, restoreOptions) {
    for (var i = 0; i < items.length; i++) {
        var item = items[i];
        if (!item) { continue; }
        switch (item.typename) {
            case "GroupItem":
                if (item.clipped) {
                    clearEffectsAppearanceOnly(item, restoreOptions);
                } else {
                    processClearEffectsItems(item.pageItems, restoreOptions);
                }
                break;
            case "PathItem":
                clearEffectsPreserveFillStroke(item, restoreOptions);
                break;
            case "CompoundPathItem":
                clearEffectsAppearanceOnly(item, restoreOptions);
                break;
            case "TextFrame":
                clearEffectsTextPreserveFill(item, restoreOptions);
                break;
        }
    }
}

/**
 * 対象1つだけを選択する（消去アクションを対象単位で効かせるため）。
 * @param {object} targetItem 対象項目 / target item
 * @returns {void}
 */
function selectOnlyForClear(targetItem) {
    var currentDocument = app.activeDocument;
    currentDocument.selection = null;
    targetItem.selected = true;
}

/**
 * クリップグループ・複合パス向け：アピアランスを消去し、不透明度・描画モードだけ戻す（塗り・線は戻さない）。
 * @param {object} item 対象項目 / target item
 * @param {object} restoreOptions 復元オプション / restore options
 * @returns {void}
 */
function clearEffectsAppearanceOnly(item, restoreOptions) {
    try {
        var savedOpacity = null;
        var savedBlendingMode = null;
        if (restoreOptions && restoreOptions.opacity) {
            try { savedOpacity = item.opacity; } catch (opacityReadError) { }
        }
        if (restoreOptions && restoreOptions.blendingMode) {
            try { savedBlendingMode = item.blendingMode; } catch (blendReadError) { }
        }
        selectOnlyForClear(item);
        playClearAppearanceAction();
        try { if (savedOpacity !== null) { item.opacity = savedOpacity; } } catch (opacityWriteError) { }
        try { if (savedBlendingMode !== null) { item.blendingMode = savedBlendingMode; } } catch (blendWriteError) { }
    } catch (clearOnlyError) { }
}

/**
 * パス向け：アピアランスを消去し、元の塗り・線・線幅（＋設定に応じて線属性・オーバープリント・不透明度・描画モード）を復元する。
 * @param {object} item PathItem
 * @param {object} restoreOptions 復元オプション / restore options
 * @returns {void}
 */
function clearEffectsPreserveFillStroke(item, restoreOptions) {
    try {
        var hasFill = item.filled;
        var hasStroke = item.stroked;
        var savedFill = hasFill ? cloneColorForClear(item.fillColor) : null;
        var savedStroke = hasStroke ? cloneColorForClear(item.strokeColor) : null;
        var savedStrokeWidth = hasStroke ? item.strokeWidth : 1;

        var savedStrokeCap = null;
        var savedStrokeJoin = null;
        var savedStrokeDashes = null;
        var savedStrokeDashOffset = null;
        var savedStrokeMiterLimit = null;
        var savedStrokeOverprint = null;
        var savedFillOverprint = null;

        if (hasFill && restoreOptions.overprint) {
            try { savedFillOverprint = item.fillOverprint; } catch (fillOverprintReadError) { }
        }
        if (hasStroke && restoreOptions.strokeSettings) {
            try { savedStrokeCap = item.strokeCap; } catch (capReadError) { }
            try { savedStrokeJoin = item.strokeJoin; } catch (joinReadError) { }
            try { savedStrokeDashes = item.strokeDashes ? item.strokeDashes.slice(0) : null; } catch (dashReadError) { }
            try { savedStrokeDashOffset = item.strokeDashOffset; } catch (dashOffsetReadError) { }
            try { savedStrokeMiterLimit = item.strokeMiterLimit; } catch (miterReadError) { }
        }
        if (hasStroke && restoreOptions.overprint) {
            try { savedStrokeOverprint = item.strokeOverprint; } catch (strokeOverprintReadError) { }
        }

        var savedOpacity = null;
        var savedBlendingMode = null;
        if (restoreOptions.opacity) {
            try { savedOpacity = item.opacity; } catch (opacityReadError) { }
        }
        if (restoreOptions.blendingMode) {
            try { savedBlendingMode = item.blendingMode; } catch (blendReadError) { }
        }

        selectOnlyForClear(item);
        playClearAppearanceAction();

        if (hasFill && savedFill) {
            item.filled = true;
            item.fillColor = savedFill;
            if (restoreOptions.overprint) {
                try { if (savedFillOverprint !== null) { item.fillOverprint = savedFillOverprint; } } catch (fillOverprintWriteError) { }
            }
        } else {
            item.filled = false;
            item.fillColor = makeNoColorForClear();
        }

        if (hasStroke && savedStroke) {
            item.stroked = true;
            item.strokeColor = savedStroke;
            item.strokeWidth = savedStrokeWidth;
            if (restoreOptions.strokeSettings) {
                try { if (savedStrokeCap !== null) { item.strokeCap = savedStrokeCap; } } catch (capWriteError) { }
                try { if (savedStrokeJoin !== null) { item.strokeJoin = savedStrokeJoin; } } catch (joinWriteError) { }
                try { if (savedStrokeDashes !== null) { item.strokeDashes = savedStrokeDashes; } } catch (dashWriteError) { }
                try { if (savedStrokeDashOffset !== null) { item.strokeDashOffset = savedStrokeDashOffset; } } catch (dashOffsetWriteError) { }
                try { if (savedStrokeMiterLimit !== null) { item.strokeMiterLimit = savedStrokeMiterLimit; } } catch (miterWriteError) { }
            }
            if (restoreOptions.overprint) {
                try { if (savedStrokeOverprint !== null) { item.strokeOverprint = savedStrokeOverprint; } } catch (strokeOverprintWriteError) { }
            }
        } else {
            item.stroked = false;
            item.strokeColor = makeNoColorForClear();
        }

        try { if (savedOpacity !== null) { item.opacity = savedOpacity; } } catch (opacityWriteError) { }
        try { if (savedBlendingMode !== null) { item.blendingMode = savedBlendingMode; } } catch (blendWriteError) { }
    } catch (preserveError) { }
}

/**
 * テキスト向け：アピアランスを消去し、文字単位の塗り（textFillPerChar）を復元する（線は復元しない）。
 * @param {object} textFrame TextFrame
 * @param {object} restoreOptions 復元オプション / restore options
 * @returns {void}
 */
function clearEffectsTextPreserveFill(textFrame, restoreOptions) {
    try {
        var textRange = textFrame.textRange;
        var characters = null;
        var characterCount = 0;
        var characterFills = [];
        var hasCharacters = false;

        try {
            characters = textRange.characters;
            characterCount = characters.length;
            hasCharacters = (characterCount > 0);
        } catch (characterReadError) {
            characters = null;
            characterCount = 0;
            hasCharacters = false;
        }

        if (restoreOptions.textFillPerChar && hasCharacters) {
            for (var i = 0; i < characterCount; i++) {
                characterFills.push(getTextRangeFillCloneForClear(characters[i]));
            }
        }

        var firstCharacterFill = null;
        if (restoreOptions.textFillFirst && hasCharacters) {
            firstCharacterFill = getTextRangeFillCloneForClear(characters[0]);
        }

        var rangeFill = getTextRangeFillCloneForClear(textRange);

        var savedOpacity = null;
        var savedBlendingMode = null;
        if (restoreOptions.opacity) {
            try { savedOpacity = textFrame.opacity; } catch (opacityReadError) { }
        }
        if (restoreOptions.blendingMode) {
            try { savedBlendingMode = textFrame.blendingMode; } catch (blendReadError) { }
        }

        selectOnlyForClear(textFrame);
        playClearAppearanceAction();

        if (restoreOptions.textFillPerChar && hasCharacters) {
            for (var j = 0; j < characterCount; j++) {
                restoreTextFillOnlyForClear(characters[j], characterFills[j]);
            }
        } else if (restoreOptions.textFillFirst) {
            var fillToApply = firstCharacterFill || rangeFill;
            restoreTextFillOnlyForClear(textRange, fillToApply);
        }

        try { if (savedOpacity !== null) { textFrame.opacity = savedOpacity; } } catch (opacityWriteError) { }
        try { if (savedBlendingMode !== null) { textFrame.blendingMode = savedBlendingMode; } } catch (blendWriteError) { }
    } catch (textPreserveError) { }
}

/**
 * カラーオブジェクトを型ごとに安全に複製する。
 * @param {object} color 複製元カラー / source color
 * @returns {object} 複製したカラー（未対応型は null）/ cloned color (null when unsupported)
 */
function cloneColorForClear(color) {
    if (!color) { return null; }
    switch (color.typename) {
        case "RGBColor":
            var rgbColor = new RGBColor();
            rgbColor.red = color.red;
            rgbColor.green = color.green;
            rgbColor.blue = color.blue;
            return rgbColor;
        case "CMYKColor":
            var cmykColor = new CMYKColor();
            cmykColor.cyan = color.cyan;
            cmykColor.magenta = color.magenta;
            cmykColor.yellow = color.yellow;
            cmykColor.black = color.black;
            return cmykColor;
        case "GrayColor":
            var grayColor = new GrayColor();
            grayColor.gray = color.gray;
            return grayColor;
        case "SpotColor":
            var spotColor = new SpotColor();
            spotColor.spot = color.spot;
            spotColor.tint = color.tint;
            return spotColor;
        case "GradientColor":
            var gradientColor = new GradientColor();
            gradientColor.gradient = color.gradient;
            gradientColor.angle = color.angle;
            gradientColor.length = color.length;
            gradientColor.origin = color.origin;
            gradientColor.matrix = color.matrix;
            return gradientColor;
        case "PatternColor":
            var patternColor = new PatternColor();
            patternColor.pattern = color.pattern;
            try { patternColor.matrix = color.matrix; } catch (patternMatrixError) { }
            return patternColor;
        case "NoColor":
            return new NoColor();
        default:
            return null;
    }
}

/**
 * NoColor を生成して返す。
 * @returns {NoColor} NoColor インスタンス / a NoColor instance
 */
function makeNoColorForClear() {
    return new NoColor();
}

/**
 * テキスト範囲の塗りカラーを複製して返す（NoColor・取得失敗時は null）。
 * @param {object} textRange TextRange
 * @returns {object} 複製した塗りカラー、または null / cloned fill color or null
 */
function getTextRangeFillCloneForClear(textRange) {
    var fillColor = null;
    try {
        fillColor = textRange.characterAttributes.fillColor;
    } catch (fillReadError) {
        return null;
    }
    if (fillColor && fillColor.typename && fillColor.typename !== "NoColor") {
        return cloneColorForClear(fillColor);
    }
    return null;
}

/**
 * テキスト範囲に塗りだけを復元する（線は NoColor に）。
 * @param {object} textRange TextRange
 * @param {object} fill 復元する塗りカラー（null なら NoColor）/ fill color to restore (NoColor when null)
 * @returns {void}
 */
function restoreTextFillOnlyForClear(textRange, fill) {
    var attributes = textRange.characterAttributes;
    if (fill) {
        attributes.fillColor = cloneColorForClear(fill);
    } else {
        attributes.fillColor = makeNoColorForClear();
    }
    attributes.strokeColor = makeNoColorForClear();
}

/* 委譲する worker 関数はすべてここに登録する（登録漏れ防止）/ register every delegated worker function */
var WORKER_FUNCS = [
    workerApplyCompoundShape,
    workerApplyPathfinder,
    workerExpandCompoundShape,
    workerReleaseCompoundShape,
    workerShowAppearancePanel,
    workerShowPathfinderPanel,
    workerSelectShapeBuilderTool,
    workerFillHoles,
    workerExpandAppearance,
    workerClearAppearance,
    workerClearEffectsOnly,
    playClearAppearanceAction,
    processClearEffectsItems,
    selectOnlyForClear,
    clearEffectsAppearanceOnly,
    clearEffectsPreserveFillStroke,
    clearEffectsTextPreserveFill,
    cloneColorForClear,
    makeNoColorForClear,
    getTextRangeFillCloneForClear,
    restoreTextFillOnlyForClear,
    buildEnumeratedActionSource,
    workerStrokeToFill,
    workerCleanupCollinear,
    isCleanupSkippable,
    collectCleanupPathItems,
    isCleanupCollinear,
    removeRedundantAnchorsCollinear,
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
 * 選択中の複合シェイプを解除する / release the selected compound shape
 * @returns {string} マーカー / marker string
 */
function delegateReleaseCompoundShape() {
    return delegateCall('workerReleaseCompoundShape()');
}

/**
 * アピアランスパネルを表示する / show the Appearance panel
 * @returns {string} マーカー / marker string
 */
function delegateShowAppearancePanel() {
    return delegateCall('workerShowAppearancePanel()');
}

/**
 * パスファインダーパネルを表示する / show the Pathfinder panel
 * @returns {string} マーカー / marker string
 */
function delegateShowPathfinderPanel() {
    return delegateCall('workerShowPathfinderPanel()');
}

/**
 * シェイプ形成ツールに切り替える / switch to the Shape Builder tool
 * @returns {string} マーカー / marker string
 */
function delegateSelectShapeBuilderTool() {
    return delegateCall('workerSelectShapeBuilderTool()');
}

/**
 * アピアランスを解除（消去＋基本属性の復元）する / clear appearance and restore basic attributes
 * @returns {string} マーカー / marker string
 */
function delegateClearAppearance() {
    return delegateCall('workerClearAppearance()');
}

/**
 * 効果のみを消去（アピアランス消去＋塗り・線の復元）する / clear effects only (clear appearance, keep fill/stroke)
 * @returns {string} マーカー / marker string
 */
function delegateClearEffectsOnly() {
    return delegateCall('workerClearEffectsOnly()');
}

/**
 * マド埋め（複合パス解除＋合体）を実行する / fill holes (release compound + unite)
 * @param {boolean} expand 拡張して実パスにするか（false ならライブ効果のまま）/ bake to paths (false keeps live effect)
 * @returns {string} マーカー / marker string
 */
function delegateFillHoles(expand) {
    return delegateCall('workerFillHoles(' + (expand ? 'true' : 'false') + ')');
}

/**
 * アピアランスを分割（実体化）する / expand appearance
 * @returns {string} マーカー / marker string
 */
function delegateExpandAppearance() {
    return delegateCall('workerExpandAppearance()');
}

/**
 * 線を塗りに（アウトライン化＋合流＋合体）を実行する / stroke to fill (outline + merge + add)
 * @returns {string} マーカー / marker string
 */
function delegateStrokeToFill() {
    return delegateCall('workerStrokeToFill()');
}

/**
 * 選択パスの直線上の冗長アンカーを削除（パスを整形）する / clean up collinear anchor points
 * @returns {string} マーカー / marker string
 */
function delegateCleanupCollinear() {
    return delegateCall('workerCleanupCollinear()');
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

    if (iconType === "unite") {
        /* 合体：2つを同色で塗り、継ぎ目なしの和集合シルエット */
        fillRect(graphics, iconBrush, backX, backY, side, side);
        fillRect(graphics, iconBrush, frontX, frontY, side, side);
    } else if (iconType === "merge") {
        /* 合流：和集合シルエット。back を塗り、front の左辺のすぐ外側（back の下タブ側）に細い縦の白いシームを入れてから
         * front を塗り直し、front（重なり側）が欠けないようにする */
        fillRect(graphics, iconBrush, backX, backY, side, side);
        fillRect(graphics, backgroundBrush, frontX - 2, frontY, 2, overlapHeight);
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
        /* 分割：2つを塗り、重なり部分の輪郭を背景色の線で囲って分割の継ぎ目（小さな窓）を見せる */
        fillRect(graphics, iconBrush, backX, backY, side, side);
        fillRect(graphics, iconBrush, frontX, frontY, side, side);
        strokeRect(graphics, backgroundPen, frontX, frontY, overlapWidth, overlapHeight);
    } else if (iconType === "trim") {
        /* 刈り込み：back を front で切り取った L字にし、front との境界に細い白いシームを残して front を最前面で塗る */
        var trimSeam = 3;
        fillRect(graphics, iconBrush, backX, backY, side, side);
        fillRect(graphics, backgroundBrush, frontX - trimSeam, frontY - trimSeam, side + trimSeam, side + trimSeam);
        fillRect(graphics, iconBrush, frontX, frontY, side, side);
    } else if (iconType === "crop") {
        /* 切り抜き：back は中間色の太い枠。front はベタ塗りしてから内側の右下（重なりを除く）を背景色でくり抜き、
         * 枠と重なりブロックを継ぎ目のない1つの黒い形にする（腕と中央が細い角だけで繋がってすき間に見えるのを防ぐ） */
        var cropBorder = 3;
        var cropBackPen = graphics.newPen(graphics.PenType.SOLID_COLOR, colors.muted, cropBorder);
        strokeRect(graphics, cropBackPen, backX, backY, side, side);
        fillRect(graphics, iconBrush, frontX, frontY, side, side);
        fillRect(graphics, backgroundBrush,
            frontX + overlapWidth, frontY + cropBorder,
            side - cropBorder - overlapWidth, side - cropBorder * 2);
        fillRect(graphics, backgroundBrush,
            frontX + cropBorder, frontY + overlapHeight,
            side - cropBorder * 2, side - cropBorder - overlapHeight);
        /* back の枠と front が交わる2点を背景色で抜き、light の枠と dark の図形を分離して見せる。
         * 抜きは front の枠（cropBorder）を貫く大きさにして、ブロックと腕をつなぐ細いブリッジを残さない */
        var cropGap = cropBorder * 2;
        fillRect(graphics, backgroundBrush,
            (backX + side) - cropGap / 2, frontY - cropGap / 2, cropGap, cropGap);
        fillRect(graphics, backgroundBrush,
            frontX - cropGap / 2, (backY + side) - cropGap / 2, cropGap, cropGap);
        /* 中央の■は欠けさせない：抜きの後に重なりブロックを塗り直して常に完全な正方形にする */
        fillRect(graphics, iconBrush, frontX, frontY, overlapWidth, overlapHeight);
    } else if (iconType === "outline") {
        /* アウトライン：2つとも同色の輪郭のみ（線に変換）。2つの交差点を背景色で抜き、
         * 交差の中心に白い窓を空けて2つの枠が編み込まれて見えるようにする */
        strokeRect(graphics, iconPen, backX, backY, side, side);
        strokeRect(graphics, iconPen, frontX, frontY, side, side);
        var outlineGap = 5;
        fillRect(graphics, backgroundBrush,
            (backX + side) - outlineGap / 2, frontY - outlineGap / 2, outlineGap, outlineGap);
        fillRect(graphics, backgroundBrush,
            frontX - outlineGap / 2, (backY + side) - outlineGap / 2, outlineGap, outlineGap);
    } else if (iconType === "minusBack") {
        /* 背面型抜き：front（背面側＝右下）を塗り、back（前面側＝左上）を背景色で塗って輪郭を付け、重なりを白く抜く */
        fillRect(graphics, iconBrush, frontX, frontY, side, side);
        fillRect(graphics, backgroundBrush, backX, backY, side, side);
        strokeRect(graphics, iconPen, backX, backY, side, side);
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

/* ボタンの高さを指定 px 詰める（レイアウト確定後に呼ぶ）/ Trim a button's height by the given px (call after layout) */
function trimButtonHeight(button, px) {
    try {
        button.size = [button.size.width, button.size.height - px];
    } catch (e) {}
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
 * 操作アイコンボタン（46x46・onDraw 描画）＋直下のキャプションを container に追加する。
 * アイコンとラベルを縦グループにまとめ、キャプションは caption.<icon> のローカライズ短縮名を使う。
 * 有効/無効の切り替えでキャプションも一緒にディムできるよう、button に __caption 参照を持たせる。
 * @param {object} container 追加先のパネル/グループ / parent container
 * @param {object} operation icon / labelKey を持つ操作定義 / operation with icon & labelKey
 * @param {function} onClickHandler クリック時の処理 / click handler
 * @returns {object} 追加した iconbutton（__caption にキャプションを保持）/ the added iconbutton (holds its caption on __caption)
 */
function addOperationButton(container, operation, onClickHandler) {
    var cell = container.add("group");
    cell.orientation = "column";
    cell.alignChildren = "center";
    cell.spacing = 2;
    var button = cell.add("iconbutton", undefined, undefined, { style: "toolbutton" });
    button.preferredSize = [46, 46];
    button.helpTip = getLocalizedText(operation.labelKey);
    button.onDraw = makeIconDrawer(operation.icon);
    button.onClick = onClickHandler;
    var caption = cell.add("statictext", undefined, getLocalizedText("caption." + operation.icon));
    caption.justify = "center";
    button.__caption = caption;
    button.__cell = cell;
    return button;
}

/**
 * アイコンセル群の幅を最大キャプション幅に統一する（列を揃えるため）。
 * 各キャプションの描画幅を measureString で測って最大値を求め、全セル・全キャプションへ適用する。
 * @param {object[]} buttons addOperationButton が返した iconbutton 配列（__cell / __caption 保持）
 * @returns {void}
 */
function unifyIconCellWidths(buttons) {
    var maxWidth = 46;
    for (var i = 0; i < buttons.length; i++) {
        var caption = buttons[i].__caption;
        if (!caption) { continue; }
        var measuredWidth = 0;
        try {
            measuredWidth = caption.graphics.measureString(caption.text, caption.graphics.font, 1000)[0];
        } catch (measureError) {
            measuredWidth = 0;
        }
        if (measuredWidth > maxWidth) { maxWidth = measuredWidth; }
    }
    for (var j = 0; j < buttons.length; j++) {
        if (buttons[j].__cell) { buttons[j].__cell.preferredSize.width = maxWidth; }
        if (buttons[j].__caption) { buttons[j].__caption.preferredSize.width = maxWidth; }
    }
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
 * オプションパネル（チェックボックス＋拡張ボタン＋パスを整形ボタン）を構築する。
 * 拡張ボタンは複合シェイプ選択時のみ機能するが、判定はクリック時に worker 側で行うため常に押せる。
 * Option（Alt）+クリックで拡張ではなく解除する（配線は showPalette 側）。
 * パスを整形ボタンは選択パスの直線上の冗長アンカーを削除する（許容誤差 0.02 固定）。
 * @param {Window} parentWindow 親ウィンドウ / parent window
 * @returns {{removePoints: object, removeUnpainted: object, expand: object, cleanup: object}} 各コントロール / controls
 */
function buildOptionPanel(parentWindow) {
    var panel = parentWindow.add("panel", undefined, getLocalizedText("panel.option"));
    setupPanel(panel);
    /* 「余分なポイントを削除」と［強制］ボタンを同じ行に横並び / removePoints checkbox + Force button on one row */
    var removePointsRow = panel.add("group");
    removePointsRow.orientation = "row";
    removePointsRow.alignChildren = ["left", "center"];
    removePointsRow.alignment = "left";
    removePointsRow.spacing = PANEL_SPACING;
    var controls = {
        removePoints:    removePointsRow.add("checkbox", undefined, getLocalizedText("option.removePoints")),
        cleanup:         removePointsRow.add("button", undefined, getLocalizedText("button.cleanup")),
        removeUnpainted: panel.add("checkbox", undefined, getLocalizedText("option.removeUnpainted")),
        expand:          panel.add("button", undefined, getLocalizedText("button.expand"))
    };
    controls.removePoints.value = true;
    controls.removeUnpainted.value = false;
    controls.removeUnpainted.helpTip = getLocalizedText("tip.removeUnpainted");
    controls.expand.helpTip = getLocalizedText("tip.expand") + " / " + getLocalizedText("tip.optionRelease");
    controls.cleanup.helpTip = getLocalizedText("tip.cleanup");
    controls.removeUnpainted.alignment = "left";
    controls.expand.alignment = "center";
    return controls;
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
    /* タブを入れるので外周余白は小さめにし、各タブ側で内側余白を持たせる / smaller window margin; tabs hold the inner padding */
    paletteWindow.margins = 8;

    /* タブ（基本／Special）/ Tabs (Basic / Special) */
    var tabbedPanel = paletteWindow.add("tabbedpanel");
    tabbedPanel.alignChildren = "fill";
    tabbedPanel.alignment = "fill";

    var basicTab = tabbedPanel.add("tab", undefined, getLocalizedText("tab.basic"));
    basicTab.orientation = "column";
    basicTab.alignChildren = "fill";
    basicTab.margins = [12, 14, 0, 12];
    basicTab.spacing = WINDOW_SPACING;

    var specialTab = tabbedPanel.add("tab", undefined, getLocalizedText("tab.special"));
    specialTab.orientation = "column";
    specialTab.alignChildren = "fill";
    specialTab.margins = [12, 14, 0, 12];
    specialTab.spacing = WINDOW_SPACING;

    tabbedPanel.selection = 0;

    /* モードパネル（出力モードの排他ラジオ・最上段）/ Mode panel (output-mode radios, top)
     * A: 実行（実際にパスへ）/ B: 複合シェイプ（上段のみ）/ C: 効果として適用（ライブ）
     */
    var modeRadios = buildModePanel(basicTab);
    var modeExecuteRadio = modeRadios.execute;
    var modeCompoundRadio = modeRadios.compound;
    var modeEffectRadio = modeRadios.effect;

    /* 形状モードパネル（アイコンボタンは後段で追加）/ Shape mode panel (buttons added below) */
    var shapeModePanel = buildShapeModePanel(basicTab);

    /* パスファインダーパネル（3個×2行, ボタンは後段で追加）/ Pathfinder panel (3 per row × 2, buttons added below) */
    var pathfinderRows = buildPathfinderRows(basicTab);

    /* オプションパネル（チェックボックス＋拡張ボタン）/ Options panel (checkboxes + expand button) */
    var optionControls = buildOptionPanel(basicTab);
    var removePointsCheckbox = optionControls.removePoints;
    var removeUnpaintedCheckbox = optionControls.removeUnpainted;
    var expandButton = optionControls.expand;
    var cleanupButton = optionControls.cleanup;

    /* その他タブ：マド埋め／変換／アピアランスの3パネル / "Special" tab: three grouped panels */
    var fillHolesPanel = specialTab.add("panel", undefined, getLocalizedText("panel.fillHoles"));
    setupPanel(fillHolesPanel);
    var fillHolesRow = fillHolesPanel.add("group");
    fillHolesRow.orientation = "row";
    fillHolesRow.alignment = "left";
    var fillHolesExpandButton = fillHolesRow.add("button", undefined, getLocalizedText("button.fillHolesExpand"));
    fillHolesExpandButton.helpTip = getLocalizedText("tip.fillHolesExpand");
    var fillHolesEffectButton = fillHolesRow.add("button", undefined, getLocalizedText("button.fillHolesEffect"));
    fillHolesEffectButton.helpTip = getLocalizedText("tip.fillHolesEffect");

    var convertPanel = specialTab.add("panel", undefined, getLocalizedText("panel.convert"));
    setupPanel(convertPanel);
    var strokeToFillButton = convertPanel.add("button", undefined, getLocalizedText("button.strokeToFill"));
    strokeToFillButton.helpTip = getLocalizedText("tip.strokeToFill");
    strokeToFillButton.alignment = "left";

    var appearancePanel = specialTab.add("panel", undefined, getLocalizedText("panel.appearance"));
    setupPanel(appearancePanel);
    var appearanceRow = appearancePanel.add("group");
    appearanceRow.orientation = "column";
    appearanceRow.alignment = "left";
    appearanceRow.alignChildren = "fill";
    var expandAppearanceButton = appearanceRow.add("button", undefined, getLocalizedText("button.expandAppearance"));
    expandAppearanceButton.helpTip = getLocalizedText("tip.expandAppearance");
    var clearEffectsOnlyButton = appearanceRow.add("button", undefined, getLocalizedText("button.clearEffectsOnly"));
    clearEffectsOnlyButton.helpTip = getLocalizedText("tip.clearEffectsOnly");
    var clearAppearanceButton = appearanceRow.add("button", undefined, getLocalizedText("button.clearAppearance"));
    clearAppearanceButton.helpTip = getLocalizedText("tip.clearAppearance");

    var showPanelPanel = specialTab.add("panel", undefined, getLocalizedText("panel.showPanel"));
    setupPanel(showPanelPanel);
    var appearanceButton = showPanelPanel.add("button", undefined, getLocalizedText("button.appearance"));
    appearanceButton.helpTip = getLocalizedText("tip.appearance");
    appearanceButton.alignment = "left";
    var pathfinderPanelButton = showPanelPanel.add("button", undefined, getLocalizedText("button.pathfinderPanel"));
    pathfinderPanelButton.helpTip = getLocalizedText("tip.pathfinderPanel");
    pathfinderPanelButton.alignment = "left";
    var shapeBuilderButton = showPanelPanel.add("button", undefined, getLocalizedText("button.shapeBuilder"));
    shapeBuilderButton.helpTip = getLocalizedText("tip.shapeBuilder");
    shapeBuilderButton.alignment = "left";

    /* isBusy ガード付きで委譲を実行する / guarded delegate */
    function runExclusive(produceStatus) {
        if (isBusy) { return; }
        isBusy = true;
        try {
            produceStatus();
        } catch (delegateError) {
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
     * Option+クリック → 出力モードに関係なく効果として適用（＝モード C と同じ）
     * ※ B（複合シェイプ）選択時は下段が無効化されクリックできない。removeUnpainted は分割・アウトラインのみ有効
     */
    function makePathfinderHandler(pathfinder) {
        return function () {
            /* onClick は event を持たないため、直前の mousedown で記録した Option 状態を読む
             * onClick carries no event; read the Option state recorded by the preceding mousedown */
            var withOption = (this.__altPressed === true);
            this.__altPressed = false;
            runExclusive(function () {
                var destructive = !modeEffectRadio.value && !withOption;
                /* 「効果として適用」（ライブ効果＝アピアランスに残る）ときは removeUnpainted を強制OFF
                 * ExtractUnpainted はライブ効果だと分割・アウトラインの見た目を壊すため、実行モードのみ有効
                 * force removeUnpainted OFF in effect mode; only apply it in the destructive (Apply) mode */
                var shouldRemoveUnpainted = pathfinder.unpainted && removeUnpaintedCheckbox.value && destructive;
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
        /* Option+クリックのヒントと mousedown での alt 記録を上乗せ / add Option-click hint & alt recorder */
        pathfinderButton.helpTip = getLocalizedText(pathfinder.labelKey) + " / " + getLocalizedText("tip.optionEffect");
        pathfinderButton.addEventListener("mousedown", makeAltRecorder(pathfinderButton));
        pathfinderButtons.push(pathfinderButton);
    }
    /* 6アイコン（2行×3）をひとつのセットとしてセル幅を統一し、上下の列を揃える
     * unify the 6 pathfinder cells (2 rows × 3) as one set so columns line up */
    unifyIconCellWidths(pathfinderButtons);

    /* B（複合シェイプ）は形状モード専用 → 選択時は下段パスファインダーを無効化 / disable Pathfinders under mode B */
    function updatePathfinderEnabled() {
        var enabled = !modeCompoundRadio.value;
        for (var buttonIndex = 0; buttonIndex < pathfinderButtons.length; buttonIndex++) {
            pathfinderButtons[buttonIndex].enabled = enabled;
            if (pathfinderButtons[buttonIndex].__caption) {
                pathfinderButtons[buttonIndex].__caption.enabled = enabled;
            }
        }
    }
    modeExecuteRadio.onClick = updatePathfinderEnabled;
    modeCompoundRadio.onClick = updatePathfinderEnabled;
    modeEffectRadio.onClick = updatePathfinderEnabled;
    updatePathfinderEnabled();

    /* 拡張ボタンのクリックで選択中の複合シェイプを拡張する（判定はクリック時に worker 側で実施）
     * Option（Alt）+クリックのときは拡張ではなく解除する（直前の mousedown で記録した Option 状態を読む）
     * パレットは選択変更を通知できず enabled のキャッシュは古くなるため、常に押せるようにして
     * クリック時に複合シェイプの有無を判定する（無ければ worker 側で "NOCS" を返す）
     * expand the compound shape on click; Option-click releases it instead (read the Option state
     * recorded by the preceding mousedown). Validate the selection at click time in the worker. */
    expandButton.addEventListener("mousedown", makeAltRecorder(expandButton));
    expandButton.onClick = function () {
        var withOption = (this.__altPressed === true);
        this.__altPressed = false;
        runExclusive(function () {
            if (withOption) {
                return markerToStatus(delegateReleaseCompoundShape(), { labelKey: "button.release" });
            }
            return markerToStatus(delegateExpandCompoundShape(), { labelKey: "button.expand" });
        });
    };

    /* Option（Alt）押下中は拡張ボタンのラベルを「複合シェイプを解除」に切り替える（見た目だけ。実処理は onClick 側で判定）
     * ボタン上での mousemove（Option 状態を含む）と、ウィンドウの keydown/keyup で更新する
     * ※ macOS の ScriptUI では修飾キー単独の keydown が発火しないことがあるため、ボタンにマウスを重ねる mousemove を主トリガにする
     * swap the label to Release while Option is held (cosmetic; the action is decided in onClick) */
    function updateExpandButtonLabel(withOption) {
        expandButton.text = withOption
            ? getLocalizedText("button.expandRelease")
            : getLocalizedText("button.expand");
    }
    expandButton.addEventListener("mousemove", function (mouseEvent) {
        updateExpandButtonLabel(mouseEvent.altKey === true);
    });
    expandButton.addEventListener("mouseout", function () {
        updateExpandButtonLabel(false);
    });
    paletteWindow.addEventListener("keydown", function (kbEvent) {
        if (kbEvent.altKey === true) { updateExpandButtonLabel(true); }
    });
    paletteWindow.addEventListener("keyup", function (kbEvent) {
        updateExpandButtonLabel(kbEvent.altKey === true);
    });

    /* パスを整形ボタンのクリックで直線上の冗長アンカーを削除する（許容誤差 0.02）
     * clean up collinear anchor points on click (tolerance 0.02) */
    cleanupButton.onClick = function () {
        runExclusive(function () {
            return markerToStatus(delegateCleanupCollinear(), { labelKey: "button.cleanup" });
        });
    };

    /* マド埋め（拡張）：複合パス解除＋合体して実パスに拡張 / fill holes and expand to paths */
    fillHolesExpandButton.onClick = function () {
        runExclusive(function () {
            return markerToStatus(delegateFillHoles(true), { labelKey: "button.fillHolesExpand" });
        });
    };

    /* マド埋め（効果）：複合パス解除＋合体をライブ効果のまま残す / fill holes, keep as live effect */
    fillHolesEffectButton.onClick = function () {
        runExclusive(function () {
            return markerToStatus(delegateFillHoles(false), { labelKey: "button.fillHolesEffect" });
        });
    };

    /* 線を塗りに：アウトライン化＋合流＋合体をライブ効果で実行する / stroke to fill on click */
    strokeToFillButton.onClick = function () {
        runExclusive(function () {
            return markerToStatus(delegateStrokeToFill(), { labelKey: "button.strokeToFill" });
        });
    };

    /* アピアランスを分割ボタンのクリックでアピアランスを実体化する / expand appearance on click */
    expandAppearanceButton.onClick = function () {
        runExclusive(function () {
            return markerToStatus(delegateExpandAppearance(), { labelKey: "button.expandAppearance" });
        });
    };

    /* 効果のみを消去：アピアランスを消去し、元の塗り・線（テキストは文字塗り）だけを戻す / clear effects only, keep fill/stroke */
    clearEffectsOnlyButton.onClick = function () {
        runExclusive(function () {
            return markerToStatus(delegateClearEffectsOnly(), { labelKey: "button.clearEffectsOnly" });
        });
    };

    /* （完全に）消去：消去して塗り・線などの基本属性を復元する / clear appearance and restore attributes */
    clearAppearanceButton.onClick = function () {
        runExclusive(function () {
            return markerToStatus(delegateClearAppearance(), { labelKey: "button.clearAppearance" });
        });
    };

    /* アピアランスボタンのクリックでアピアランスパネルを表示する / show the Appearance panel on click */
    appearanceButton.onClick = function () {
        runExclusive(function () {
            return markerToStatus(delegateShowAppearancePanel(), { labelKey: "button.appearance" });
        });
    };

    /* パスファインダーボタンのクリックでパスファインダーパネルを表示する / show the Pathfinder panel on click */
    pathfinderPanelButton.onClick = function () {
        runExclusive(function () {
            return markerToStatus(delegateShowPathfinderPanel(), { labelKey: "button.pathfinderPanel" });
        });
    };

    /* シェイプ形成ツールボタンのクリックでツールを切り替える / switch to the Shape Builder tool on click */
    shapeBuilderButton.onClick = function () {
        runExclusive(function () {
            return markerToStatus(delegateSelectShapeBuilderTool(), { labelKey: "button.shapeBuilder" });
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
     * Esc: 閉じる / P: パスに変換 / C: 複合シェイプにする / F: 効果として適用
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

    /* レイアウトを確定させてから全プッシュボタンの高さのみ -2 で詰める（アイコンボタンは対象外）
     * finalize layout, then trim only the height of every push button by 2px (icon buttons excluded) */
    paletteWindow.layout.layout(true);
    var trimTargetButtons = [
        expandButton, cleanupButton,
        fillHolesExpandButton, fillHolesEffectButton,
        strokeToFillButton,
        expandAppearanceButton, clearEffectsOnlyButton, clearAppearanceButton,
        appearanceButton, pathfinderPanelButton, shapeBuilderButton
    ];
    for (var trimIndex = 0; trimIndex < trimTargetButtons.length; trimIndex++) {
        trimButtonHeight(trimTargetButtons[trimIndex], 2);
    }

    paletteWindow.show();
}

showPalette();

})();
