#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

ColorToK100Converter.jsx

### 概要

- RGB または CMYK で構成された黒を安定した K100 の黒に変換する Illustrator 用スクリプトです。
- テキスト、パス、スウォッチの塗りおよび線カラーが一括対象になります。

### 主な機能

- RGB 黒（RGB 各値が 39 未満）を K100 に変換
- CMYK の多色ブラック（CMY 全てが 70 以上 または合計が 310 以上）を K100 に変換
- テキストの文字単位での変換対応
- スウォッチカラーも自動変換
- 日本語／英語インターフェース対応

### 処理の流れ

1. テキストの文字単位の塗り・線カラーを変換
2. パス、コンパウンドパスの塗り・線カラーを変換
3. グループ内オブジェクトを再帰的に変換
4. スウォッチのカラー定義を変換

### 更新履歴

- v1.0.0 (20250612) : 初期バージョン
- v1.0.1 (20250615) : 処理構造とコメント整理

---

### Script Name:

ColorToK100Converter.jsx

### Overview

- An Illustrator script that converts RGB or CMYK-based blacks into stable K100 black.
- Targets fills and strokes in text, paths, and swatches all at once.

### Main Features

- Converts RGB blacks (each channel under 39) to K100
- Converts CMYK rich blacks (all CMY ≥ 70 or total ≥ 310) to K100
- Supports per-character conversion for text
- Also converts swatch colors automatically
- Japanese and English UI support

### Process Flow

1. Convert text character fill/stroke colors
2. Convert path and compound path fill/stroke colors
3. Recursively convert objects in groups
4. Convert swatch color definitions

### Update History

- v1.0.0 (20250612): Initial version
- v1.0.1 (20250615): Refined structure and cleaned comments
*/

// ロック／非表示オブジェクトの除外が必要な場合は、ここで判定を追加すると良い（現状は全対象）

var k100Swatch = null;
try {
  k100Swatch = app.activeDocument.swatches.getByName("ブラック");
} catch (e) {
  k100Swatch = null;
}

// 黒判定用の閾値定数オブジェクト
var blackThresholds = {
    cmyk: { cyan: 70, magenta: 70, yellow: 70 },
    rgb: { red: 39, green: 39, blue: 39 }
};

// RGBおよびCMYKの黒をK100に変換する関数
function convertRGBBlackToK100(color) {
  // --- RGBカラー処理 ---
  if (color.typename === "RGBColor") {
    var r = color.red;
    var g = color.green;
    var b = color.blue;
    // すべて39未満なら黒と判定してK100に変換
    if (r < blackThresholds.rgb.red && g < blackThresholds.rgb.green && b < blackThresholds.rgb.blue) {
      if (k100Swatch && k100Swatch.color.typename === "CMYKColor") {
        return k100Swatch.color;
      } else {
        var fallback = new CMYKColor();
        fallback.cyan = 0;
        fallback.magenta = 0;
        fallback.yellow = 0;
        fallback.black = 100;
        return fallback;
      }
    }
    // それ以外は変換せずそのまま返す
    return color;
  }

  // --- CMYKカラー処理 ---
  if (color.typename === "CMYKColor") {
    var c = color.cyan;
    var m = color.magenta;
    var y = color.yellow;
    var k = color.black;
    var total = c + m + y + k;

    // 墨ベタ（K=100 かつ CMYがすべて10以下）はそのまま許容
    if (k === 100 && c <= 10 && m <= 10 && y <= 10) {
      return color;
    }

    // 変換対象条件：CMYKのすべてが60以上 または CMYのすべてが70以上 または 合計310以上
    var isRichBlack = (
      (c >= blackThresholds.cmyk.cyan &&
       m >= blackThresholds.cmyk.magenta &&
       y >= blackThresholds.cmyk.yellow) ||
      (c >= 60 && m >= 60 && y >= 60 && k >= 60) ||
      total >= 310
    );

    if (isRichBlack) {
      if (k100Swatch && k100Swatch.color.typename === "CMYKColor") {
        return k100Swatch.color;
      } else {
        var fallback = new CMYKColor();
        fallback.cyan = 0;
        fallback.magenta = 0;
        fallback.yellow = 0;
        fallback.black = 100;
        return fallback;
      }
    }
    // それ以外は変換せずそのまま返す
    return color;
  }

  // その他のカラー型は変換せずそのまま返す
  return color;
}

function processGroupItem(group) {
  var paths = group.pathItems;
  for (var i = 0; i < paths.length; i++) {
    if (paths[i].locked || paths[i].hidden) continue;
    if (paths[i].filled) paths[i].fillColor = convertRGBBlackToK100(paths[i].fillColor);
    if (paths[i].stroked) paths[i].strokeColor = convertRGBBlackToK100(paths[i].strokeColor);
  }

  var compoundPaths = group.compoundPathItems;
  for (var i = 0; i < compoundPaths.length; i++) {
    var cp = compoundPaths[i];
    if (cp.locked || cp.hidden) continue;
    if (cp.filled) cp.fillColor = convertRGBBlackToK100(cp.fillColor);
    if (cp.stroked) cp.strokeColor = convertRGBBlackToK100(cp.strokeColor);
  }

  var texts = group.textFrames;
  for (var j = 0; j < texts.length; j++) {
    if (texts[j].locked || texts[j].hidden) continue;
    var chars = texts[j].characters;
    for (var k = 0; k < chars.length; k++) {
      chars[k].fillColor = convertRGBBlackToK100(chars[k].fillColor);
      chars[k].strokeColor = convertRGBBlackToK100(chars[k].strokeColor);
    }
  }

  var subGroups = group.groupItems;
  for (var m = 0; m < subGroups.length; m++) {
    processGroupItem(subGroups[m]);
  }
}

function main() {
  // ドキュメントのカラーモードをチェック（RGBの場合は処理を中断）
  if (app.activeDocument.documentColorSpace === DocumentColorSpace.RGB) {
    alert("このスクリプトはCMYKカラーモードのドキュメントでのみ使用できます。");
    return;
  }
  // ドキュメント内すべての TextFrame の各 Character の fillColor/strokeColor を変換
  var doc = app.activeDocument;

  // TextFrame の文字単位での塗り・線カラー変換
  var textFrames = doc.textFrames;
  for (var i = 0; i < textFrames.length; i++) {
    var tf = textFrames[i];
    if (tf.locked || tf.hidden) continue;
    var chars = tf.characters;
    for (var j = 0; j < chars.length; j++) {
      var ch = chars[j];
      ch.fillColor = convertRGBBlackToK100(ch.fillColor);
      ch.strokeColor = convertRGBBlackToK100(ch.strokeColor);
    }
  }

  // PathItem の塗り・線カラー変換
  var paths = doc.pathItems;
  for (var i = 0; i < paths.length; i++) {
    var path = paths[i];
    if (path.locked || path.hidden) continue;
    if (path.filled) path.fillColor = convertRGBBlackToK100(path.fillColor);
    if (path.stroked) path.strokeColor = convertRGBBlackToK100(path.strokeColor);
  }

  // CompoundPathItem の塗り・線カラー変換
  var compoundPaths = doc.compoundPathItems;
  for (var i = 0; i < compoundPaths.length; i++) {
    var compPath = compoundPaths[i];
    if (compPath.locked || compPath.hidden) continue;
    if (compPath.filled) compPath.fillColor = convertRGBBlackToK100(compPath.fillColor);
    if (compPath.stroked) compPath.strokeColor = convertRGBBlackToK100(compPath.strokeColor);
  }

  // GroupItem を再帰的に処理
  var groups = doc.groupItems;
  for (var i = 0; i < groups.length; i++) {
    processGroupItem(groups[i]);
  }

  // スウォッチ内のカラー定義の変換
  var swatches = doc.swatches;
  for (var i = 0; i < swatches.length; i++) {
    var col = swatches[i].color;
    if (col.typename === "RGBColor" || col.typename === "CMYKColor") {
      swatches[i].color = convertRGBBlackToK100(col);
    }
  }
}

main();