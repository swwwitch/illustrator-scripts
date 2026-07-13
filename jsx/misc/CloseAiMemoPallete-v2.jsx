#target illustrator
#targetengine "TextMemoEngine"

/*
  パレット単体クローズのテンプレート / Single-palette close template

  他パレット用に流用するときは、次の 2 か所だけ書き換える /
  To reuse for another palette, change only these two spots:
    1) 2 行目の #targetengine "..." → 対象パレットの常駐エンジン名 / the target palette's engine name
    2) 下の GLOBAL_REF → 対象パレットの $.global 参照名 / the target palette's $.global reference name

  ※ 1 スクリプト＝1 エンジン。#targetengine を複数書いても複数エンジンでは動かない /
     One script runs in exactly one engine; stacking #targetengine does NOT target multiple engines.
  ※ #targetengine はコンパイル時リテラルなので変数化不可。必ず 2 行目を直接書き換える /
     #targetengine must be a compile-time literal — edit line 2 directly, it cannot be a variable.

  対象 / Target: AiMemoPallete
*/

var GLOBAL_REF = "__TextMemoWindow"; // ← 2 行目の #targetengine と対で書き換える / change together with line 2

(function () {
    var paletteRef = $.global[GLOBAL_REF];
    if (!paletteRef) return; // 参照なし＝開いていない / No reference means it is not open

    // Window 直接保持形式と { window: Window } ラッパー形式の両対応 / Support both a Window and a { window } wrapper
    var paletteWindow = (paletteRef.close ? paletteRef : (paletteRef.window ? paletteRef.window : null));
    try { if (paletteWindow) paletteWindow.close(); } catch (e) {}

    $.global[GLOBAL_REF] = null; // 参照を解放 / Release the reference
})();
