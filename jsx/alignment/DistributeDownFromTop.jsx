#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

(function(){
	if(app.documents.length < 1) return
	if(app.activeDocument.selection.length < 2) return

	var selectedItems = sortByVerticalPosition(app.activeDocument.selection)

	// 移動量は環境設定［テキスト］の「サイズ／行送り」キー入力（text/sizeIncrement）を参照
	var verticalGap = app.preferences.getRealPreference("text/sizeIncrement")

	// 最上部のオブジェクトは固定し、以降を verticalGap ずつ下へずらす
	for(var i=1; i<selectedItems.length; i++){
		selectedItems[i].translate(0, -i * verticalGap)
	}

	function sortByVerticalPosition(selection){
		var sorted = []
		for(var i=0; i<selection.length; i++) sorted.push(selection[i])
		sorted.sort(function(a, b){
			return b.position[1] - a.position[1]
		})
		return sorted
	}
})()
