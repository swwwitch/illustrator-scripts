#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
  グリッド複製ツール / Duplicate in Grid
  Version: SCRIPT_VERSION（変数で管理）

  概要 / Overview
  - 選択中のオブジェクトを、指定した「繰り返し数 × 繰り返し数」のグリッドで複製します。
  - 間隔は現在の定規単位（rulerType）で入力し、内部では pt に変換して処理します。
  - ダイアログの値変更（入力・↑↓キー・Shift/Option 併用）でライブプレビューが更新されます。
  - ローカライズ（日本語/英語）とバージョン表記に対応し、ダイアログタイトルにバージョンを表示します。
  - Fill パネル：［アートボードの端まで］に対応します。UIとして［アートボードいっぱいに］を追加（選択オブジェクトを基準に行列を計算し、OK時に複製を含む全体をアートボード中央へ移動）。
  - 方向（横：右/左、縦：上/下）を指定できます。
  - クリッピングマスクを優先して境界を取得（なければ可視境界）します。
    - プレビュー描画は一時オブジェクトのみ削除（noteタグで管理）し、既存の「_preview」レイヤー上の他要素は削除しません。
  - 複数選択時は自動でグループ化してから処理（以降の挙動は単一オブジェクトと同じ）。/ When multiple objects are selected, they are grouped first and processed like a single object.

  更新日 / Last Updated: 2025-10-29

 - 2025-10-26: 「アートボードいっぱいに」ON時に「端まで」を自動OFF、方向（横・縦）をディム表示に変更（UI挙動のみ、機能変更なし）。
  - 2025-10-29: 複数オブジェクト選択時に自動でグループ化してから実行する挙動を追加（既存機能は変更なし）。
  - 2025-10-29: 既存の「_preview」レイヤー内容が消える不具合を修正（プレビューにnoteタグを付与し、そのタグのみ削除）。
 変更履歴 / Changelog
  - 2025-10-26: 「方向」を左または上に変更した場合、「アートボードの端まで」を自動OFFにする挙動を追加（UI挙動のみ、機能変更なし）。
  - 2025-10-26: 英語ラベルの略語を廃止し「Link Horizontal & Vertical」に統一（UIテキストのみ変更、機能変更なし）。
  - 2025-10-26: 「アートボードいっぱいに」のロジックを実装：選択オブジェクトを基準に行数・列数を算出し、OK時に全体をアートボード中央へ配置。
  - 2025-10-26: Fill パネルに「アートボードいっぱいに」チェックを UI 追加（ロジック未実装、今後対応予定）。
  - 2025-10-26: 英語ラベルを「Fill to Artboard Edge」に変更（機能変更なし）。
  - 2025-10-25: 「アートボードの端まで」実行時に (0,0) セルの複製を常にスキップするよう修正（元オブジェクトの二重化バグ修正）。
  - 2025-10-25: 「アートボードの端まで」選択時に方向を自動で［右・下］に設定。機能追加（既存既定値の明示化）。
  - 2025-10-25: Fill パネルの見出しをローカライズ対応（JA/EN）。機能変更なし。
  - 2025-10-25: オプション「最背面のオブジェクト」「方向を無視して端まで」を削除。ドキュメント更新。機能の他部分は変更なし。
  - 2025-10-25: ドキュメント整備（概要更新、コメントの粒度調整、不要コメント削除、コメントアウトコード整理）。機能変更なし。
  - 2025-10-23: ローカライズ（JA/EN）、SCRIPT_VERSION 導入、ダイアログタイトルにバージョン表示を追加
  - 2025-10-23: rulerType に基づく単位ラベル表示と、単位→pt 変換を実装
  - 2025-10-23: ライブプレビュー（_preview レイヤー）を追加、再描画の安定化（app.redraw）
  - 2025-10-23: ↑↓／Shift／Option による数値増減、繰り返し数の整数化を実装
  - 2025-10-23: UI 調整（ラベル幅統一・右揃え、OK を右側、透明度・位置補正）
*/

/* バージョン / Version */
var SCRIPT_VERSION = "v1.9";

/* 言語判定 / Locale detection */
function getCurrentLang() {
    return ($.locale && $.locale.toLowerCase().indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();
var PREVIEW_TAG = "__grid_preview__";

/* ラベル定義 / Label definitions (JA/EN) */
var LABELS = {
    dialogTitle: { ja: "グリッド状に複製", en: "Duplicate in Grid" },
    repeatCount: { ja: "繰り返し数", en: "Count" },
    repeatCountH:{ ja: "横", en: "Horizontal" },
    repeatCountV:{ ja: "縦", en: "Vertical" },
    gap:         { ja: "間隔", en: "Gap" },
    unitFmt:     { ja: "間隔（{unit}）:", en: "Gap ({unit}):" },
    fillTitle:   { ja: "敷き詰め", en: "Fill" },
    ok:          { ja: "OK", en: "OK" },
    cancel:      { ja: "キャンセル", en: "Cancel" },
    alertNoDoc:  { ja: "ドキュメントが開かれていません。", en: "No document is open." },
    alertNoSel:  { ja: "オブジェクトを選択してください。", en: "Please select an object." },
    alertCountInvalid:{ ja: "繰り返し数は1以上の整数を入力してください。", en: "Enter an integer count of 1 or more." },
    alertGapInvalid:  { ja: "間隔は数値で入力してください。", en: "Enter a numeric gap value." },
    directionTitle:   { ja: "方向", en: "Direction" },
    dirRight:   { ja: "右", en: "Right" },
    dirLeft:    { ja: "左", en: "Left" },
    verticalTitle:{ ja: "縦方向", en: "Vertical" },
    dirUp:      { ja: "上", en: "Up" },
    dirDown:    { ja: "下", en: "Down" }
};

/* ラベル取得関数 / Label resolver */
function L(key, params) {
    var table = LABELS[key];
    var text = (table && table[lang]) ? table[lang] : key;
    if (params) {
        for (var k in params) if (params.hasOwnProperty(k)) {
            text = text.replace(new RegExp("\\{" + k + "\\}", "g"), params[k]);
        }
    }
    return text;
}

/* テキストフィールドの数値操作 / Arrow key increment–decrement for numeric fields */
function changeValueByArrowKey(edittext) {
    edittext.addEventListener("keydown", function(e) {
        var v = Number(edittext.text);
        if (isNaN(v)) return;
        var kb = ScriptUI.environment.keyboardState, d = 1;
        if (kb.shiftKey) {
            d = 10;
            if (e.keyName == "Up") { v = Math.ceil((v + 1)/d)*d; e.preventDefault(); }
            else if (e.keyName == "Down") { v = Math.floor((v - 1)/d)*d; e.preventDefault(); }
        } else if (kb.altKey) {
            d = 0.1;
            if (e.keyName == "Up") { v += d; e.preventDefault(); }
            else if (e.keyName == "Down") { v -= d; e.preventDefault(); }
        } else {
            d = 1;
            if (e.keyName == "Up") { v += d; e.preventDefault(); }
            else if (e.keyName == "Down") { v -= d; e.preventDefault(); }
        }
        v = kb.altKey ? Math.round(v*10)/10 : Math.round(v);
        if (edittext.isInteger) v = Math.max(0, Math.round(v));
        edittext.text = v;
        if (typeof edittext.onChanging === "function") { try { edittext.onChanging(); } catch(_){} }
    });
}

/* 単位ユーティリティ / Units */
var unitLabelMap = {0:"in",1:"mm",2:"pt",3:"pica",4:"cm",5:"Q/H",6:"px",7:"ft/in",8:"m",9:"yd",10:"ft"};
function getCurrentUnitLabel(){ var c = app.preferences.getIntegerPreference("rulerType"); return unitLabelMap[c]||"pt"; }
function getCurrentUnitCode(){ return app.preferences.getIntegerPreference("rulerType"); }
function unitToPoints(unitCode, value) {
    var PT_IN=72, MM_PT=PT_IN/25.4;
    switch(unitCode){
        case 0: return value*PT_IN;           // in
        case 1: return value*MM_PT;           // mm
        case 2: return value;                 // pt
        case 3: return value*12;              // pica
        case 4: return value*(MM_PT*10);      // cm
        case 5: return value*(MM_PT*0.25);    // Q/H
        case 6: return value;                 // px ≒ pt
        case 7: return value*PT_IN;           // ft/in → in
        case 8: return value*(MM_PT*1000);    // m
        case 9: return value*(PT_IN*36);      // yd
        case 10:return value*(PT_IN*12);      // ft
        default:return value;
    }
}

/* クリッピングマスク優先の境界取得 / Bounds preferring clipping mask */
function getMaskedBounds(item){
    try{
        if (item.typename==="GroupItem" && item.clipped){
            for (var i=0;i<item.pageItems.length;i++){
                var pi=item.pageItems[i];
                if (pi.typename==="PathItem" && pi.clipping) return pi.geometricBounds;
                if (pi.typename==="CompoundPathItem" && pi.pathItems.length>0 && pi.pathItems[0].clipping)
                    return pi.pathItems[0].geometricBounds;
            }
        }
        if (item.typename==="PathItem" && item.clipping) return item.geometricBounds;
        if (item.typename==="CompoundPathItem" && item.pathItems.length>0 && item.pathItems[0].clipping)
            return item.pathItems[0].geometricBounds;
    }catch(_){}
    return item.visibleBounds;
}

/* プレビュー用ユーティリティ / Preview utilities */
function getPreviewLayer(doc){
    var name="_preview", lyr;
    try{ lyr=doc.layers.getByName(name); }
    catch(e){ lyr=doc.layers.add(); lyr.name=name; }
    lyr.visible=true; lyr.locked=false;
    try{ lyr.zOrder(ZOrderMethod.BRINGTOFRONT); }catch(_){}
    return lyr;
}
function clearPreview(doc){
    try{
        var lyr = doc.layers.getByName("_preview");
        // Remove only items we created (note tagged)
        for (var i = lyr.pageItems.length - 1; i >= 0; i--) {
            try {
                var it = lyr.pageItems[i];
                if (it.note === PREVIEW_TAG) it.remove();
            } catch(_) {}
        }
        app.redraw();
    }catch(e){}
}
function buildPreview(doc, sourceItem, rows, cols, gapX, gapY, w, h, hDir, vDir, baseL, baseT){
    var lyr=getPreviewLayer(doc);
    clearPreview(doc);
    for (var r=0;r<rows;r++){
        for (var c=0;c<cols;c++){
            if (r===0 && c===0) continue;
            var dup=sourceItem.duplicate(lyr, ElementPlacement.PLACEATBEGINNING);
            dup.note = PREVIEW_TAG;
            var offX=(w+gapX)*c; if (hDir==="left") offX=-offX;
            var offY=(h+gapY)*r;
            var desiredL=baseL+offX;
            var desiredT=(vDir==="up")?(baseT+offY):(baseT-offY);
            var mb=getMaskedBounds(dup); // [L,T,R,B]
            var dx=desiredL-mb[0], dy=desiredT-mb[1];
            dup.left+=dx; dup.top+=dy;
        }
    }
    app.redraw();
}

/* ダイアログ生成＆プレビュー結線 / Build dialog and preview wiring */
function showDialog(doc, sourceItem, w, h){
    var dlg=new Window("dialog", L("dialogTitle")+" "+SCRIPT_VERSION);
    var offsetX=300, dialogOpacity=0.98;
    dlg.onShow=function(){ var loc=dlg.location; dlg.location=[loc[0]+offsetX, loc[1]]; };
    dlg.opacity=dialogOpacity;
    dlg.orientation="column"; dlg.alignChildren="fill";

    var unitCode=getCurrentUnitCode(), unitLabel=getCurrentUnitLabel();

    /* 繰り返し数 / Repeat Count */
    var repeatPanel=dlg.add("panel", undefined, L("repeatCount"));
    repeatPanel.orientation="row"; repeatPanel.alignChildren="top";
    repeatPanel.margins=[20,20,20,10]; repeatPanel.spacing=20;

    var repeatLeftCol=repeatPanel.add("group"); repeatLeftCol.orientation="column"; repeatLeftCol.alignChildren="left";
    var repeatRightCol=repeatPanel.add("group"); repeatRightCol.orientation="column"; repeatRightCol.alignChildren="left"; repeatRightCol.alignment=["left","center"];

    var repeatXGroup=repeatLeftCol.add("group");
    repeatXGroup.add("statictext", undefined, L("repeatCountH")+":");
    var countXInput=repeatXGroup.add("edittext", undefined, "2"); countXInput.characters=4; countXInput.isInteger=true; changeValueByArrowKey(countXInput);

    var repeatYGroup=repeatLeftCol.add("group");
    repeatYGroup.add("statictext", undefined, L("repeatCountV")+":");
    var countYInput=repeatYGroup.add("edittext", undefined, "2"); countYInput.characters=4; countYInput.isInteger=true; changeValueByArrowKey(countYInput);

    var linkGroup=repeatRightCol.add("group"); linkGroup.alignment=["left","center"];
    var linkCheck=linkGroup.add("checkbox", undefined, (lang==="ja"?"連動":"Link Horizontal & Vertical"));
    linkCheck.value=true;
    function syncCounts(){
        if (linkCheck.value){
            countYInput.enabled=false;
            countYInput.text=countXInput.text;
            if (typeof countYInput.onChanging==="function"){ try{ countYInput.onChanging(); }catch(_){} }
        } else { countYInput.enabled=true; }
    }
    linkCheck.onClick=function(){ syncCounts(); applyPreview(); };
    syncCounts();

    /* パネル：Fill（アートボードの端まで／アートボードいっぱいに） / Panel: Fill */
    var fillPanel=dlg.add("panel", undefined, L("fillTitle"));
    fillPanel.orientation="column"; fillPanel.alignChildren="left";
    fillPanel.margins=[20,15,20,10]; fillPanel.spacing=8;

    var fillABGroup=fillPanel.add("group"); fillABGroup.alignment=["left","center"];
    var fillToArtboardCheck=fillABGroup.add("checkbox", undefined, (lang==="ja"?"アートボードの端まで":"Fill to Artboard Edge"));
    fillToArtboardCheck.value=false;
    fillToArtboardCheck.onClick=function(){
        if (fillToArtboardCheck.value){
            // 「アートボードの端まで」選択時は方向を［右・下］に固定
            if (typeof dirRight !== "undefined") { dirRight.value = true; }
            if (typeof dirLeft  !== "undefined") { dirLeft.value  = false; }
            if (typeof dirDown  !== "undefined") { dirDown.value  = true; }
            if (typeof dirUp    !== "undefined") { dirUp.value    = false; }
            // 方向コントロールは有効化
            if (typeof dirRight !== "undefined") dirRight.enabled = true;
            if (typeof dirLeft  !== "undefined") dirLeft.enabled  = true;
            if (typeof dirUp    !== "undefined") dirUp.enabled    = true;
            if (typeof dirDown  !== "undefined") dirDown.enabled  = true;
            linkCheck.value=false; 
            syncCounts();
            recalcCountsForArtboard();
        }
        if (typeof fillToArtboardFullCheck !== "undefined" && fillToArtboardCheck.value) {
            fillToArtboardFullCheck.value = false;
        }
        applyPreview();
    };

    // 追加：アートボードいっぱいに
    var fillABFullGroup = fillPanel.add("group"); fillABFullGroup.alignment=["left","center"];
    var fillToArtboardFullCheck = fillABFullGroup.add("checkbox", undefined, (lang==="ja"?"アートボードいっぱいに":"Fill Full Artboard"));
    fillToArtboardFullCheck.value = false;
    fillToArtboardFullCheck.onClick = function(){
        if (fillToArtboardFullCheck.value){
            // 端までと排他
            if (typeof fillToArtboardCheck !== "undefined") fillToArtboardCheck.value = false;
            // 方向は任意。既定を右・下へ
            if (typeof dirRight !== "undefined") { dirRight.value = true; }
            if (typeof dirLeft  !== "undefined") { dirLeft.value  = false; }
            if (typeof dirDown  !== "undefined") { dirDown.value  = true; }
            if (typeof dirUp    !== "undefined") { dirUp.value    = false; }
            // 方向コントロールをディム（無効化）
            if (typeof dirRight !== "undefined") dirRight.enabled = false;
            if (typeof dirLeft  !== "undefined") dirLeft.enabled  = false;
            if (typeof dirUp    !== "undefined") dirUp.enabled    = false;
            if (typeof dirDown  !== "undefined") dirDown.enabled  = false;
            // 行列をアートボード内に収まる最大値で自動計算
            recalcCountsForArtboardFull();
        } else {
            // チェック解除時は方向コントロールを再有効化
            if (typeof dirRight !== "undefined") dirRight.enabled = true;
            if (typeof dirLeft  !== "undefined") dirLeft.enabled  = true;
            if (typeof dirUp    !== "undefined") dirUp.enabled    = true;
            if (typeof dirDown  !== "undefined") dirDown.enabled  = true;
        }
        applyPreview();
    };

    /* 間隔（現在単位表記） / Gap (show in current ruler units) */
    var gapPanel=dlg.add("panel", undefined, (lang==="ja"?("間隔（"+unitLabel+"）"):("Gap ("+unitLabel+")")));
    gapPanel.orientation="row"; gapPanel.alignChildren="top";
    gapPanel.margins=[20,15,20,10]; gapPanel.spacing=20;

    var gapLeftCol=gapPanel.add("group"); gapLeftCol.orientation="column"; gapLeftCol.alignChildren="left";
    var gapRightCol=gapPanel.add("group"); gapRightCol.orientation="column"; gapRightCol.alignChildren="left"; gapRightCol.alignment=["left","center"];

    var gapXGroup=gapLeftCol.add("group");
    gapXGroup.add("statictext", undefined, (lang==="ja"?"左右:":"Horizontal:"));
    var gapXInput=gapXGroup.add("edittext", undefined, "10"); gapXInput.characters=4; changeValueByArrowKey(gapXInput);

    var gapYGroup=gapLeftCol.add("group");
    gapYGroup.add("statictext", undefined, (lang==="ja"?"上下:":"Vertical:"));
    var gapYInput=gapYGroup.add("edittext", undefined, "10"); gapYInput.characters=4; changeValueByArrowKey(gapYInput);

    var gapLinkGroup=gapRightCol.add("group"); gapLinkGroup.alignment=["left","center"];
    var gapLink=gapLinkGroup.add("checkbox", undefined, (lang==="ja"?"連動":"Link Horizontal & Vertical")); gapLink.value=true;
    function syncGaps(){
        if (gapLink.value){
            gapYInput.enabled=false; gapYInput.text=gapXInput.text;
            if (typeof gapYInput.onChanging==="function"){ try{ gapYInput.onChanging(); }catch(_){} }
        } else { gapYInput.enabled=true; }
    }
    gapLink.onClick=function(){ syncGaps(); applyPreview(); };
    syncGaps();

    if (fillToArtboardCheck.value){ recalcCountsForArtboard(); }

    // 初期状態の有効/無効を同期
    if (fillToArtboardFullCheck.value){
        if (typeof dirRight !== "undefined") dirRight.enabled = false;
        if (typeof dirLeft  !== "undefined") dirLeft.enabled  = false;
        if (typeof dirUp    !== "undefined") dirUp.enabled    = false;
        if (typeof dirDown  !== "undefined") dirDown.enabled  = false;
    } else {
        if (typeof dirRight !== "undefined") dirRight.enabled = true;
        if (typeof dirLeft  !== "undefined") dirLeft.enabled  = true;
        if (typeof dirUp    !== "undefined") dirUp.enabled    = true;
        if (typeof dirDown  !== "undefined") dirDown.enabled  = true;
    }

    /* 方向パネル（横・縦の展開方向を指定） / Direction panel (horizontal & vertical placement) */
    var dirPanel=dlg.add("panel", undefined, L("directionTitle"));
    dirPanel.orientation="column"; dirPanel.alignChildren="left"; dirPanel.margins=[20,15,20,10];

    var hGroup=dirPanel.add("group"); hGroup.orientation="row"; hGroup.alignChildren="left";
    hGroup.add("statictext", undefined, (lang==="ja"?"横方向":"Horizontal"));
    var dirRight=hGroup.add("radiobutton", undefined, L("dirRight"));
    var dirLeft =hGroup.add("radiobutton", undefined, L("dirLeft"));
    dirRight.value = true;
    dirRight.onClick = function(){
        // 右選択時はそのままプレビュー
        applyPreview();
    };
    dirLeft.onClick = function(){
        // 左を選んだら「アートボードの端まで」を自動OFF
        try{
            if (typeof fillToArtboardCheck !== "undefined" && fillToArtboardCheck.value){
                fillToArtboardCheck.value = false;
            }
        }catch(_){}
        applyPreview();
    };

    var vGroup=dirPanel.add("group"); vGroup.orientation="row"; vGroup.alignChildren="left";
    vGroup.add("statictext", undefined, (lang==="ja"?"縦方向":"Vertical"));
    var dirUp  =vGroup.add("radiobutton", undefined, L("dirUp"));
    var dirDown=vGroup.add("radiobutton", undefined, L("dirDown"));
    dirDown.value = true;
    dirUp.onClick = function(){
        // 上を選んだら「アートボードの端まで」を自動OFF
        try{
            if (typeof fillToArtboardCheck !== "undefined" && fillToArtboardCheck.value){
                fillToArtboardCheck.value = false;
            }
        }catch(_){}
        applyPreview();
    };
    dirDown.onClick = function(){
        // 下選択時はそのままプレビュー
        applyPreview();
    };

    function applyPreview(){
        var cx=parseInt(countXInput.text,10), cy=parseInt(countYInput.text,10);
        var gx=parseFloat(gapXInput.text),  gy=parseFloat(gapYInput.text);
        if (isNaN(cx)||cx<1) return; if (isNaN(cy)||cy<1) return; if (isNaN(gx)||isNaN(gy)) return;
        var gptX=unitToPoints(unitCode,gx), gptY=unitToPoints(unitCode,gy);
        var hDir=dirRight.value?"right":"left";
        var vDir=(dirUp&&dirUp.value)?"up":"down";
        var baseMask=getMaskedBounds(sourceItem), baseLeft=baseMask[0], baseTop=baseMask[1];
        buildPreview(doc, sourceItem, cy, cx, gptX, gptY, w, h, hDir, vDir, baseLeft, baseTop);
    }

    // アートボードの端まで：列・行の自動計算
    function recalcCountsForArtboard(){
        try{
            var tgt=doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect; // [L,T,R,B]
            var tgtL=tgt[0], tgtT=tgt[1], tgtR=tgt[2], tgtB=tgt[3];
            var baseMask=getMaskedBounds(sourceItem), baseLeft=baseMask[0], baseTop=baseMask[1];

            var gx=parseFloat(gapXInput.text), gy=parseFloat(gapYInput.text);
            if (isNaN(gx)||isNaN(gy)) return;
            var gptX=unitToPoints(unitCode,gx), gptY=unitToPoints(unitCode,gy);
            var stepX=w+gptX, stepY=h+gptY;

            var hDir=dirRight.value?"right":"left";
            var vDir=(typeof dirUp!=="undefined" && dirUp.value)?"up":"down";

            var cols=1, rows=1;
            if (stepX>0){
                if (hDir==="right"){ var availW=tgtR-baseLeft; cols=Math.floor((availW+gptX)/stepX); }
                else { var availWL=baseLeft-tgtL; cols=Math.floor((availWL+gptX)/stepX); }
                if (cols<1) cols=1;
            }
            if (stepY>0){
                if (vDir==="down"){ var availH=baseTop-tgtB; rows=Math.floor((availH+gptY)/stepY); }
                else { var availHU=tgtT-baseTop; rows=Math.floor((availHU+gptY)/stepY); }
                if (rows<1) rows=1;
            }
            countXInput.text=String(cols);
            countYInput.text=String(rows);
        }catch(e){}
    }

    // アートボードいっぱいに：アートボード内に収まる最大の列・行を計算（方向は無視）
    function recalcCountsForArtboardFull(){
        try{
            var tgt = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect; // [L,T,R,B]
            var tgtW = Math.abs(tgt[2] - tgt[0]);
            var tgtH = Math.abs(tgt[1] - tgt[3]);
            var gx = parseFloat(gapXInput.text), gy = parseFloat(gapYInput.text);
            if (isNaN(gx) || isNaN(gy)) return;
            var gptX = unitToPoints(unitCode, gx), gptY = unitToPoints(unitCode, gy);
            var stepX = w + gptX, stepY = h + gptY;
            var cols = (stepX > 0) ? Math.floor((tgtW + gptX) / stepX) : 1;
            var rows = (stepY > 0) ? Math.floor((tgtH + gptY) / stepY) : 1;
            if (cols < 1) cols = 1;
            if (rows < 1) rows = 1;
            countXInput.text = String(cols);
            countYInput.text = String(rows);
        }catch(e){}
    }

    // 値変更時のプレビュー更新と再計算
    countXInput.onChanging=function(){ if (linkCheck.value) countYInput.text=countXInput.text; applyPreview(); };
    countXInput.onChange  =function(){ if (linkCheck.value) countYInput.text=countXInput.text; applyPreview(); };
    countYInput.onChanging=applyPreview;
    countYInput.onChange  =applyPreview;

    gapXInput.onChanging=function(){
        if (gapLink.value) gapYInput.text=gapXInput.text;
        if (fillToArtboardCheck.value) recalcCountsForArtboard();
        if (fillToArtboardFullCheck.value) recalcCountsForArtboardFull();
        applyPreview();
    };
    gapXInput.onChange=function(){
        if (gapLink.value) gapYInput.text=gapXInput.text;
        if (fillToArtboardCheck.value) recalcCountsForArtboard();
        if (fillToArtboardFullCheck.value) recalcCountsForArtboardFull();
        applyPreview();
    };
    gapYInput.onChanging=function(){
        if (fillToArtboardCheck.value) recalcCountsForArtboard();
        if (fillToArtboardFullCheck.value) recalcCountsForArtboardFull();
        applyPreview();
    };
    gapYInput.onChange=function(){
        if (fillToArtboardCheck.value) recalcCountsForArtboard();
        if (fillToArtboardFullCheck.value) recalcCountsForArtboardFull();
        applyPreview();
    };

    var btnGroup=dlg.add("group"); btnGroup.alignment="center";
    var cancelBtn=btnGroup.add("button", undefined, L("cancel"), {name:"cancel"});
    var okBtn    =btnGroup.add("button", undefined, L("ok"));

    var result=null;
    okBtn.onClick=function(){
        var cx=parseInt(countXInput.text,10), cy=parseInt(countYInput.text,10);
        var gx=parseFloat(gapXInput.text), gy=parseFloat(gapYInput.text);
        if (isNaN(cx)||cx<1){ alert(L("alertCountInvalid")); return; }
        if (isNaN(cy)||cy<1){ alert(L("alertCountInvalid")); return; }
        if (isNaN(gx)||isNaN(gy)){ alert(L("alertGapInvalid")); return; }
        result={
            cols:cx, rows:cy,
            gapX:unitToPoints(unitCode,gx),
            gapY:unitToPoints(unitCode,gy),
            direction:dirRight.value?"right":"left",
            vDirection:(dirUp&&dirUp.value)?"up":"down",
            fillToArtboard:fillToArtboardCheck.value,
            fillFullArtboard: (typeof fillToArtboardFullCheck !== "undefined" ? fillToArtboardFullCheck.value : false)
        };
        clearPreview(doc); dlg.close();
    };
    cancelBtn.onClick=function(){ clearPreview(doc); dlg.close(); };

    var origOnShow=dlg.onShow;
    dlg.onShow=function(){
        if (typeof origOnShow==="function"){ try{ origOnShow(); }catch(_){} }
        $.sleep(0); applyPreview(); countXInput.active=true;
    };

    dlg.show();
    return result;
}

/* メイン：検証→ダイアログ→複製 / Main: validate → dialog → duplicate */
function main(){
    if (app.documents.length===0){ alert(L("alertNoDoc")); return; }
    var doc=app.activeDocument;
    if (doc.selection.length===0){ alert(L("alertNoSel")); return; }

    var sel=doc.selection;
    // 複数選択時は自動でグループ化してから処理 / Auto-group when multiple items are selected
    if (sel.length > 1) {
        try {
            var grp = doc.groupItems.add();
            // 既存の選択配列はライブで変化し得るため、コピーを回して移動
            var toMove = [];
            for (var i = 0; i < sel.length; i++) toMove.push(sel[i]);
            for (var j = 0; j < toMove.length; j++) {
                try { toMove[j].move(grp, ElementPlacement.PLACEATEND); } catch(_) {}
            }
            // グループを選択対象にして以降の処理を単一オブジェクトと同様に
            doc.selection = null;
            grp.selected = true;
            sel = [grp];
        } catch(e) {}
    }

    var bounds=getMaskedBounds(sel[0]);
    var w=bounds[2]-bounds[0], h=bounds[1]-bounds[3];

    var settings=showDialog(doc, sel[0], w, h);
    if (!settings) return;

    var rows=settings.rows, cols=settings.cols;
    var gapX=settings.gapX, gapY=settings.gapY;
    var direction=settings.direction||"right";
    var vDirection=settings.vDirection||"down";

    var baseMask0=getMaskedBounds(sel[0]);
    var baseLeftMain=baseMask0[0], baseTopMain=baseMask0[1];

    var dupItems = [];

    for (var row=0; row<rows; row++){
        for (var col=0; col<cols; col++){
            // 元オブジェクトと同じセルは常にスキップ（ダブり防止）
            if (row===0 && col===0) continue;
            var dup=sel[0].duplicate();
            var offX=(w+gapX)*col; if (direction==="left") offX=-offX;
            var offY=(h+gapY)*row;
            var desiredL=baseLeftMain+offX;
            var desiredT=(vDirection==="up")?(baseTopMain+offY):(baseTopMain-offY);
            var mb=getMaskedBounds(dup);
            var dx=desiredL-mb[0], dy=desiredT-mb[1];
            dup.left += dx; dup.top += dy;
            dupItems.push(dup);
        }
    }

    if (settings.fillFullArtboard){
        try{
            var ab = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect; // [L,T,R,B]
            var abCX = (ab[0] + ab[2]) / 2.0;
            var abCY = (ab[1] + ab[3]) / 2.0;

            // ユニオン境界を計算（選択元 + 複製）
            var unionL = +Infinity, unionT = -Infinity, unionR = -Infinity, unionB = +Infinity;
            function expandBy(bounds){
                if (bounds[0] < unionL) unionL = bounds[0];
                if (bounds[1] > unionT) unionT = bounds[1];
                if (bounds[2] > unionR) unionR = bounds[2];
                if (bounds[3] < unionB) unionB = bounds[3];
            }
            expandBy(getMaskedBounds(sel[0]));
            for (var i=0; i<dupItems.length; i++){
                expandBy(getMaskedBounds(dupItems[i]));
            }

            var grpCX = (unionL + unionR) / 2.0;
            var grpCY = (unionT + unionB) / 2.0;

            var dx = abCX - grpCX;
            var dy = abCY - grpCY;

            // 全アイテムを同量移動（選択元＋複製）
            sel[0].left += dx; sel[0].top += dy;
            for (var j=0; j<dupItems.length; j++){
                dupItems[j].left += dx; dupItems[j].top += dy;
            }
        }catch(e){}
    }
}
main();