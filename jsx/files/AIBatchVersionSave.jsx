#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### 概要

フォルダー内の Illustrator ファイルを、指定したバージョン形式でまとめて保存するスクリプトです。

- 保存バージョンを選んで、複数の .ai / .svg ファイルを一括保存
- .svg はIllustratorで開き、選択したバージョンの .ai として保存
- 「上書きモード」と「カスタム」のラジオで保存方法を切替
- 上書きモードでは元フォルダーに同じベース名の .ai として保存
- .svg を上書きモードで処理する場合、同名 .ai の上書き可能性を事前確認
- 対象フォルダーと保存先フォルダーをダイアログ内で指定可能
- 「保存先を対象と同じにする」チェック（デフォルトON）でサフィックス付き同フォルダ保存もワンクリック
- 対象パネルで処理対象の拡張子（.ai / .svg）を選択
- 対象フォルダー内に該当ファイルが無い拡張子はチェックボックスを自動的にディム表示
- パス表示はフルパス／~ 短縮／Dropbox 短縮を切替可能
- 保存バージョンや任意の文字列を、ファイル名の末尾に追加可能
- ファイル名内のスペースを「_」または「-」に変換可能
- 出力ファイルの拡張子は常に小文字 .ai に統一
- PDF互換ファイルの作成を切り替え可能
- Illustrator標準の［更新済み］付与を使用可能
- カスタムモードで .ai 入力＆同一フォルダー＆サフィックス未指定時のみ上書き確認ダイアログを表示
- 処理中は進捗ウィンドウで現在のファイル名と件数を表示
- 処理中はIllustratorの警告ダイアログ（カラープロファイル、フォント置換、上書き確認等）を抑制
- 日本語ファイル名・記号（%, スペース）にも対応
- 日本語／英語のUIに対応

### オリジナルアイデア

クロさん（VoostOn）
https://vooston.web.fc2.com/dtp/dtp_a010.html

### Overview

Batch-saves Illustrator files in a folder using the selected save-version format.

- Batch-save multiple .ai / .svg files with a selected Illustrator version
- Open .svg files in Illustrator and save them as .ai in the selected version
- “Overwrite” / “Custom” radio modes for save method
- In Overwrite mode, save to the source folder as .ai using the same base name
- When .svg is included in Overwrite mode, confirm possible same-named .ai overwrites in advance
- Choose source and destination folders directly in the dialog
- “Use same folder as source” checkbox (default ON) for one-click suffix-renamed saving
- Target panel for selecting file types (.ai / .svg)
- Target checkboxes are auto-dimmed when no matching files exist in the source folder
- Path display toggles for full path / ~ shortening / Dropbox shortening
- Optionally append the save version and custom text to filenames
- Replace spaces in filenames with “_” or “-”
- Output extension is always normalized to lowercase .ai
- Toggle PDF-compatible file creation
- Use Illustrator’s built-in [Converted] suffix option
- In Custom mode, confirms before overwriting only when .ai inputs exist, source/destination are the same, and no suffix is set
- Progress window shows current filename and count during processing
- Suppresses Illustrator interaction dialogs (color profile, font substitution, overwrite, etc.) during processing
- Handles Japanese filenames and special characters (%, spaces)
- Supports Japanese and English UI
*/


var SCRIPT_VERSION = "v1.5.0";

// =========================================
// 保存オプション設定（必要に応じて変更）/ Save option switches (tweak as needed)
// =========================================

var SAVE_OPTS = {
	/* ファイル形式 / File format */
	compressed: true,                // 圧縮を使用 / Use compression

	/* 埋め込みリソース / Embedded resources */
	embedLinkedFiles: false,         // 配置画像を埋め込む【7 or later】/ Embed linked images
	embedICCProfile: false,          // ICC プロファイルを埋め込む【10 or later】/ Embed ICC profile
	fontSubsetThreshold: 100,        // フォントサブセット閾値（％）/ Font subset threshold (%)

	/* 出力処理 / Output rendering */
	// PRESERVEPDFOVERPRINT（保持）/ DISCARDPDFOVERPRINT（破棄）
	overprint: PDFOverprint.PRESERVEPDFOVERPRINT,
	// PRESERVEAPPEARANCE（アピアランス保持）/ PRESERVEPATHS（パス保持）
	flattenOutput: OutputFlattening.PRESERVEAPPEARANCE
};

// 注意：保存バージョンやPDF互換の設定は、ダイアログで選択されたものを使用します。/ Note: Version and PDF compatibility are taken from the dialog settings, not here.
var CONVERTED_PREF_KEY = "fileFormatGetFile/ConvertedInFilename";

// Dropbox のローカルマウントパス接頭辞（環境に合わせて書き換える）/ Local Dropbox mount prefix — adjust to your env
var DROPBOX_PREFIX = "/Users/takano/sw Dropbox/takano masahiro/";

/* ロケール判定 / Detect locale */
function getCurrentLang() {
	return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var currentLanguage = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
	dialogTitle: {
		ja: "バージョン指定で一括保存",
		en: "Batch Save As Version"
	},
	selectSourceFolder: {
		ja: "対象ファイル（.ai / .svg）の入っているフォルダを選択してください",
		en: "Select the folder containing target files (.ai / .svg)"
	},
	selectDestFolder: {
		ja: "保存するフォルダを選択してください",
		en: "Select the destination folder"
	},
	saveVersion: {
		ja: "保存バージョン",
		en: "Save version"
	},
	filenameSuffix: {
		ja: "保存バージョンを付加",
		en: "Append save version"
	},
	customSuffix: {
		ja: "任意の文字列を付加",
		en: "Append custom text"
	},
	pdfCompatible: {
		ja: "PDF互換ファイルを作成",
		en: "Create PDF Compatible File"
	},
	appendConverted: {
		ja: "Illustratorの［更新済み］付与を使用",
		en: "Use Illustrator's [Converted] suffix"
	},
	overwrite: {
		ja: "上書きモード",
		en: "Overwrite mode"
	},
	customMode: {
		ja: "カスタム",
		en: "Custom"
	},
	setSource: {
		ja: "対象...",
		en: "Source..."
	},
	sourceNotSet: {
		ja: "未指定",
		en: "Not set"
	},
	setDestination: {
		ja: "保存先...",
		en: "Destination..."
	},
	destinationNotSet: {
		ja: "未指定",
		en: "Not set"
	},
	destinationSameAsSource: {
		ja: "〈対象と同じ〉",
		en: "〈Same as source〉"
	},
	sameAsSource: {
		ja: "保存先を対象と同じにする",
		en: "Use same folder as source"
	},
	tipSameAsSource: {
		ja: "保存先フォルダーを指定せず、対象と同じフォルダーへ保存します。ファイル名にサフィックスや任意文字列を付けて別名保存できます。",
		en: "Save into the source folder without picking a separate destination. Add a suffix or custom text to keep filenames distinct."
	},
	folderPanel: {
		ja: "フォルダー指定",
		en: "Folders"
	},
	targetPanel: {
		ja: "対象",
		en: "Targets"
	},
	tipTargetAi: {
		ja: ".ai ファイルを処理対象に含めます。",
		en: "Include .ai files in the batch."
	},
	tipTargetSvg: {
		ja: ".svg ファイルを処理対象に含めます。開いた後、選択した Illustrator バージョンの .ai として保存します。",
		en: "Include .svg files in the batch. Opened and saved as .ai using the selected Illustrator version."
	},
	saveOptionsPanel: {
		ja: "保存設定",
		en: "Save Settings"
	},
	filenamePanel: {
		ja: "ファイル名の調整",
		en: "Filename Adjustments"
	},
	separator: {
		ja: "区切り文字",
		en: "Separator"
	},
	cancel: {
		ja: "キャンセル",
		en: "Cancel"
	},
	processingComplete: {
		ja: "処理が完了しました。スクリプトを終了します。",
		en: "Processing complete. Exiting script."
	},
	progressProcessing: {
		ja: "処理中",
		en: "Processing"
	},
	noFilesFound: {
		ja: "対象のファイルが見つかりませんでした。",
		en: "No matching files were found."
	},
	noTargetSelected: {
		ja: "処理対象が選択されていません。",
		en: "No target file types are selected."
	},
	noVersionsAvailable: {
		ja: "このIllustratorでは選択可能な保存バージョンがありません。",
		en: "No save-version targets are available in this Illustrator."
	},
	confirmSameFolder: {
		ja: "対象フォルダーと保存先フォルダーが同じで、ファイル名への追加もありません。\n同名で上書きされます（明示的に上書きしたい場合は「上書きモード」をご利用ください）。続行しますか？",
		en: "Source and destination folders are the same, and no filename suffix is set.\nFiles will be overwritten with the same name (use the “Overwrite” mode for explicit overwriting). Continue?"
	},
	confirmOverwriteSvg: {
		ja: "上書きモードで .svg を対象にしています。\n.svg ファイル自体は上書きされませんが、同じフォルダー内に同名の .ai がある場合は上書きされます。続行しますか？",
		en: "Overwrite mode includes .svg files.\nThe .svg files themselves will not be overwritten, but same-named .ai files in the source folder will be overwritten. Continue?"
	},
	tipSaveVersion: {
		ja: "実行中のIllustratorで保存できるバージョンのみ表示します。",
		en: "Lists only versions supported by the running Illustrator."
	},
	tipAppendConverted: {
		ja: "Illustrator標準の［更新済み］付与設定を一時的に切り替えます。処理後は元の設定に戻します。",
		en: "Temporarily changes Illustrator's built-in [Converted] filename setting, then restores it after processing."
	},
	tipPdfCompatible: {
		ja: "PDF互換ファイルを作成します。互換性は上がりますが、ファイルサイズが大きくなる場合があります。",
		en: "Creates a PDF-compatible file. This improves compatibility but may increase file size."
	},
	tipOverwrite: {
		ja: "対象フォルダー内のファイルを、同じファイル名で .ai として保存します（.ai 入力は上書き、.svg 入力は同名の .ai を生成）。保存先フォルダーやファイル名への付加設定は無効になります。",
		en: "Saves files in the source folder as .ai using the same base name (overwrites .ai inputs; .svg inputs generate a same-named .ai). Destination folder and filename append options are disabled."
	},
	tipFilenameSuffix: {
		ja: "選択した保存バージョンに対応する文字列をファイル名の末尾に付加します。",
		en: "Appends text matching the selected save version to the end of the filename."
	},
	tipCustomSuffix: {
		ja: "任意の文字列を、保存バージョンの文字列の後ろに追加します。",
		en: "Adds custom text after the save-version suffix."
	},
	tipSeparator: {
		ja: "元のファイル名内のスペース、および付加文字列との区切りに使う文字です。",
		en: "Used to replace spaces in the original filename and to separate appended text."
	},
	tipSourceFolder: {
		ja: "処理対象のファイル（.ai / .svg）が入っているフォルダーを指定します。未指定の場合はOK後に選択します。",
		en: "Selects the folder containing target files (.ai / .svg). If not set, you will be asked after clicking OK."
	},
	tipDestinationFolder: {
		ja: "別名保存先のフォルダーを指定します。未指定の場合は対象フォルダー選択後に確認します。",
		en: "Selects the destination folder for saved copies. If not set, you will be asked after choosing the source folder."
	},
	fullPathCheck: {
		ja: "フルパス",
		en: "Full path"
	},
	dropboxCheck: {
		ja: "Dropboxパスを短縮",
		en: "Shorten Dropbox path"
	},
	tipFullPath: {
		ja: "ONにすると ~ 短縮を無効にしてフルパス表示。Dropboxパス短縮がONの場合は無効。",
		en: "When ON, disables ~ abbreviation and shows the full path. Ignored when Dropbox shortening is ON."
	},
	tipDropbox: {
		ja: "Dropbox ローカルマウントの接頭辞を取り除いて表示します。",
		en: "Strips the local Dropbox mount prefix from the displayed path."
	}
};

/* ラベル取得 / Get localized label */
function getLabel(labelKey) {
	return LABELS[labelKey][currentLanguage];
}

/* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
function getLabelWithColon(labelKey) {
	return getLabel(labelKey) + (currentLanguage === 'ja' ? '：' : ':');
}

/* ツールチップ設定 / Set tooltip text */
function setHelpTip(uiControl, labelKey) {
	if (uiControl && LABELS[labelKey]) uiControl.helpTip = getLabel(labelKey);
}

/* Illustratorファイル判定 / Check whether a file is an Illustrator file */
function isIllustratorFile(fileObj) {
	return (fileObj instanceof File) && (/\.ai$/i).test(fileObj.name);
}

/* SVGファイル判定 / Check whether a file is an SVG file */
function isSvgFile(fileObj) {
	return (fileObj instanceof File) && (/\.svg$/i).test(fileObj.name);
}

/* 対象拡張子フィルタを生成 / Build a file filter that matches selected target types */

function buildTargetFileFilter(includeAi, includeSvg) {
	return function (fileObj) {
		if (!(fileObj instanceof File)) return false;
		if (includeAi && isIllustratorFile(fileObj)) return true;
		if (includeSvg && isSvgFile(fileObj)) return true;
		return false;
	};
}

/* ファイル一覧に .ai が含まれるか判定 / Check whether a file list contains any .ai files */
function hasIllustratorFiles(fileList) {
	if (!fileList) return false;
	for (var i = 0; i < fileList.length; i++) {
		if (isIllustratorFile(fileList[i])) return true;
	}
	return false;
}

/* フォルダー内の .ai / .svg ファイル数をカウント / Count .ai and .svg files in a folder */
function countTargetFiles(folder) {
	var counts = { ai: 0, svg: 0 };
	if (!folder) return counts;
	try {
		var files = folder.getFiles(function (f) { return f instanceof File; });
		if (!files) return counts;
		for (var i = 0; i < files.length; i++) {
			var name = files[i].name;
			if ((/\.ai$/i).test(name)) counts.ai++;
			else if ((/\.svg$/i).test(name)) counts.svg++;
		}
	} catch (countError) { }
	return counts;
}

/* URIデコード（失敗時は元の文字列）/ URI-decode with fallback */
function decodeFileName(rawName) {
	try {
		return decodeURI(rawName);
	} catch (decodeError) {
		return rawName;
	}
}

/* 区切り文字を考慮してファイル名パーツを連結 / Join filename parts with the chosen separator */
function appendNamePart(currentName, part, separator) {
	if (part === "") return currentName;
	if (currentName === "") return part;
	return currentName + separator + part;
}

/* ホームディレクトリ部分を ~/... に短縮 / Abbreviate home directory to ~/... */
function toTildePath(fsPath) {
	if (!fsPath) return fsPath;
	try {
		var homePath = Folder("~").fsName;
		if (homePath && homePath.length > 0) {
			if (fsPath === homePath) return "~";
			if (fsPath.indexOf(homePath + "/") === 0) return "~" + fsPath.substr(homePath.length);
			if (fsPath.indexOf(homePath + "\\") === 0) return "~" + fsPath.substr(homePath.length);
		}
	} catch (homeError) { }
	return fsPath;
}

/* 表示用パス整形（Dropbox 接頭辞優先、その後 ~ 短縮）/ Format path for display: Dropbox prefix takes priority, then tilde */
function formatDisplayPath(fsPath, useTilde, useDropbox) {
	if (!fsPath) return fsPath;
	if (useDropbox && DROPBOX_PREFIX && fsPath.indexOf(DROPBOX_PREFIX) === 0) {
		return fsPath.substr(DROPBOX_PREFIX.length);
	}
	if (useTilde) return toTildePath(fsPath);
	return fsPath;
}

// =========================================
// UI 共通設定 / UI Common Setup
// =========================================

var PANEL_MARGINS = [15, 20, 15, 10];
var PANEL_SPACING = 12;

/* パネルの共通レイアウト / Apply common panel layout */
function applyPanelLayout(panel, spacing) {
	panel.orientation = "column";
	panel.alignChildren = "left";
	panel.alignment = "fill";
	panel.margins = PANEL_MARGINS;
	panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

// =========================================
// バージョン定義 / Version Options
// =========================================

//　ドロップダウンの並び順＝ここでの定義順。デフォルトは isDefault:true の項目。
//　実行中のIllustratorが対応していないバージョンは自動的に除外される。
var VERSION_OPTIONS = [];

/* 利用可能なバージョンのみ追加 / Push only if Compatibility enum exists in current Illustrator */
function addVersionOption(label, suffix, enumKey, isDefault) {
	try {
		var compatibilityValue = Compatibility[enumKey];
		if (compatibilityValue == null) return;
		VERSION_OPTIONS.push({
			label: label,
			suffix: suffix,
			compatibility: compatibilityValue,
			isDefault: (isDefault === true)
		});
	} catch (e) { }
}

//　Illustrator CC (v17) 以降は内部ファイル形式が変わっていないため v17 に集約 / CC (v17) onward share the same .ai format
addVersionOption("Illustrator CC以降 (v17)", "v17", "ILLUSTRATOR17", true);
addVersionOption("Illustrator CS6 (v16)", "v16", "ILLUSTRATOR16");
addVersionOption("Illustrator CS5 (v15)", "v15", "ILLUSTRATOR15");
addVersionOption("Illustrator CS4 (v14)", "v14", "ILLUSTRATOR14");
addVersionOption("Illustrator CS3 (v13)", "v13", "ILLUSTRATOR13");
addVersionOption("Illustrator CS2 (v12)", "v12", "ILLUSTRATOR12");
addVersionOption("Illustrator CS (v11)", "v11", "ILLUSTRATOR11");
addVersionOption("Illustrator 10 (v10)", "v10", "ILLUSTRATOR10");
addVersionOption("Illustrator 9 (v9)", "v9", "ILLUSTRATOR9");
addVersionOption("Illustrator 8 (v8)", "v8", "ILLUSTRATOR8");

/* デフォルト指定の版が除外されていた場合は先頭の版にフォールバック / Fallback default if marked one was filtered out */
if (VERSION_OPTIONS.length > 0) {
	var hasDefaultVersion = false;
	for (var versionIndex = 0; versionIndex < VERSION_OPTIONS.length; versionIndex++) {
		if (VERSION_OPTIONS[versionIndex].isDefault) { hasDefaultVersion = true; break; }
	}
	if (!hasDefaultVersion) VERSION_OPTIONS[0].isDefault = true;
}

// =========================================
// メイン処理 / Main
// =========================================

(function () {
	if (VERSION_OPTIONS.length === 0) {
		alert(getLabel('noVersionsAvailable'));
		return;
	}

	var dialogSettings = showSettingsDialog();
	if (dialogSettings == null) return;

	//　対象が一つも選択されていない場合は早期終了（フォルダー選択ダイアログを開く前に弾く）/ Early exit when no target type is selected (before any folder picker)
	if (!dialogSettings.targetAi && !dialogSettings.targetSvg) {
		alert(getLabel('noTargetSelected'));
		return;
	}

	//　入力フォルダ決定：ダイアログで指定済→それ／未指定→ここで聞く
	//　Resolve source: picked-in-dialog → use it / otherwise → ask now
	var sourceFolder;
	if (dialogSettings.sourceFolder != null) {
		sourceFolder = dialogSettings.sourceFolder;
	} else {
		sourceFolder = Folder.selectDialog(getLabel('selectSourceFolder'));
		if (sourceFolder == null) return;
	}

	//　保存先決定：上書き or 「対象と同じ」→入力フォルダ／ダイアログで指定済→それ／未指定→ここで聞く
	//　Resolve destination: overwrite or sameAsSource → source / picked-in-dialog → use it / otherwise → ask now
	var destinationFolder;
	if (dialogSettings.overwrite || dialogSettings.sameAsSource) {
		destinationFolder = sourceFolder;
	} else if (dialogSettings.destinationFolder != null) {
		destinationFolder = dialogSettings.destinationFolder;
	} else {
		destinationFolder = Folder.selectDialog(getLabel('selectDestFolder'));
		if (destinationFolder == null) return;
	}

	//　対象パネルで選択された拡張子（.ai / .svg）のファイルを収集 / Collect files matching the selected target extensions
	var targetFileList = sourceFolder.getFiles(
		buildTargetFileFilter(dialogSettings.targetAi, dialogSettings.targetSvg)
	);
	if (targetFileList.length === 0) {
		alert(getLabel('noFilesFound'));
		return;
	}

	//　上書き事故防止：実際に .ai 入力があり、同一フォルダかつサフィックスなしなら確認 / Prevent accidental overwrite only when actual .ai inputs exist
	if (!dialogSettings.overwrite &&
		hasIllustratorFiles(targetFileList) &&
		String(sourceFolder.fsName) === String(destinationFolder.fsName) &&
		dialogSettings.suffix === "" &&
		dialogSettings.customSuffix === "" &&
		!dialogSettings.appendConverted) {
		if (!confirm(getLabel('confirmSameFolder'))) return;
	}

	//　上書きモードで実際に .svg を処理する場合のみ、同名 .ai の上書き可能性を確認 / Confirm only when SVG files actually exist in the batch
	if (dialogSettings.overwrite && dialogSettings.targetSvg) {
		var hasSvgInBatch = false;
		for (var svgIdx = 0; svgIdx < targetFileList.length; svgIdx++) {
			if (isSvgFile(targetFileList[svgIdx])) { hasSvgInBatch = true; break; }
		}
		if (hasSvgInBatch && !confirm(getLabel('confirmOverwriteSvg'))) return;
	}

	/* ［更新済み］付与の環境設定を一時的に変更し、処理後に元へ戻す / Toggle [Converted] pref then restore */
	var originalConvertedPref;
	try {
		originalConvertedPref = app.preferences.getBooleanPreference(CONVERTED_PREF_KEY);
	} catch (e) {
		originalConvertedPref = false;
	}
	app.preferences.setBooleanPreference(CONVERTED_PREF_KEY, dialogSettings.appendConverted);

	//　Illustrator の各種警告ダイアログを抑制（カラープロファイル、フォント置換、上書き確認など）/ Suppress Illustrator interaction dialogs during batch
	var prevInteractionLevel = app.userInteractionLevel;
	app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

	var progressWin = null;
	var isRestoredAfterBatch = false;

	try {
		progressWin = createProgressWindow(targetFileList.length);
		for (var i = 0; i < targetFileList.length; i++) {
			progressWin.update(i + 1, decodeFileName(targetFileList[i].name));
			saveOneFile(targetFileList[i], destinationFolder, dialogSettings);
		}

		//　完了通知の前に進捗ウィンドウを閉じ、環境設定と警告抑制を元に戻す / Restore UI state before showing completion alert
		if (progressWin != null) progressWin.close();
		app.preferences.setBooleanPreference(CONVERTED_PREF_KEY, originalConvertedPref);
		app.userInteractionLevel = prevInteractionLevel;
		isRestoredAfterBatch = true;
		alert(getLabel('processingComplete'));
	} finally {
		//　処理が中断されても進捗ウィンドウを閉じ、環境設定とユーザー操作レベルを元に戻す / Close progress window, restore preference and interaction level even if interrupted
		if (!isRestoredAfterBatch) {
			if (progressWin != null) progressWin.close();
			app.preferences.setBooleanPreference(CONVERTED_PREF_KEY, originalConvertedPref);
			app.userInteractionLevel = prevInteractionLevel;
		}
	}

	/* 1ファイルを保存 / Save one file */
	function saveOneFile(sourceFile, destinationFolder, dialogSettings) {
		var sourceDocument = null;

		try {
			sourceDocument = app.open(sourceFile);

			var outputFile = getOutputFile(sourceFile, destinationFolder, dialogSettings);
			var saveOptions = createSaveOptions(dialogSettings);
			sourceDocument.saveAs(outputFile, saveOptions);
		} finally {
			//　保存中にエラーが出ても、開いたドキュメントを閉じる / Close opened document even if saving fails
			if (sourceDocument != null) {
				try {
					sourceDocument.close(SaveOptions.DONOTSAVECHANGES);
				} catch (closeError) { }
			}
		}
	}

	/* 保存先ファイルを取得 / Get output file */
	function getOutputFile(sourceFile, destinationFolder, dialogSettings) {
		//　File.name はURIエンコード文字列なので先にデコード / File.name is URI-encoded — decode first
		var decodedName = decodeFileName(sourceFile.name);
		var dotIndex = decodedName.lastIndexOf(".");
		var baseName = (dotIndex >= 0) ? decodedName.substr(0, dotIndex) : decodedName;

		if (dialogSettings.overwrite) {
			//　上書きモード：ファイル名は加工せず、拡張子のみ小文字 .ai に統一 / In overwrite mode, keep filename but force lowercase .ai
			return new File(sourceFile.parent.fsName + "/" + baseName + ".ai");
		}

		var separator = dialogSettings.separator;
		//　半角/全角スペース・タブを区切り文字に変換 / Convert half/full-width spaces and tabs to the chosen separator
		var outputName = baseName.replace(/[ \t　]+/g, separator);
		//　末尾の区切り文字を除去して二重連結を防止 / Strip trailing separator to avoid double joins
		while (outputName.length > 0 && outputName.charAt(outputName.length - 1) === separator) {
			outputName = outputName.substr(0, outputName.length - 1);
		}
		outputName = appendNamePart(outputName, dialogSettings.suffix, separator);
		outputName = appendNamePart(outputName, dialogSettings.customSuffix, separator);
		return new File(destinationFolder.fsName + "/" + outputName + ".ai");
	}

	/* 保存オプションを作成 / Create save options */
	function createSaveOptions(dialogSettings) {
		var saveOptions = new IllustratorSaveOptions();

		/* 保存バージョン・PDF互換はダイアログ設定から / Version & PDF compatibility from dialog */
		saveOptions.compatibility = dialogSettings.compatibility;
		saveOptions.pdfCompatible = dialogSettings.pdfCompatible;

		/* それ以外は SAVE_OPTS を参照 / Others from SAVE_OPTS */
		saveOptions.compressed = SAVE_OPTS.compressed;
		saveOptions.embedLinkedFiles = SAVE_OPTS.embedLinkedFiles;
		saveOptions.embedICCProfile = SAVE_OPTS.embedICCProfile;
		saveOptions.fontSubsetThreshold = SAVE_OPTS.fontSubsetThreshold;
		saveOptions.overprint = SAVE_OPTS.overprint;
		saveOptions.flattenOutput = SAVE_OPTS.flattenOutput;

		return saveOptions;
	}

	/* プログレスウィンドウ生成 / Create progress window */
	function createProgressWindow(totalCount) {
		var progressWindow = new Window("palette", getLabel('dialogTitle') + ' ' + SCRIPT_VERSION);
		progressWindow.alignChildren = "fill";
		progressWindow.margins = 16;
		progressWindow.spacing = 8;

		var statusText = progressWindow.add("statictext", undefined, "");
		statusText.preferredSize.width = 420;

		var progressBar = progressWindow.add("progressbar", undefined, 0, totalCount);
		progressBar.preferredSize.width = 420;
		progressBar.preferredSize.height = 8;

		progressWindow.show();

		return {
			update: function (currentIndex, filename) {
				progressBar.value = currentIndex;
				statusText.text = getLabel('progressProcessing') + " (" + currentIndex + "/" + totalCount + "): " + filename;
				progressWindow.update();
			},
			close: function () {
				try { progressWindow.close(); } catch (closeError) { }
			}
		};
	}

	/* 設定ダイアログ表示 / Show settings dialog */
	function showSettingsDialog() {
		var settingsDialog = new Window("dialog", getLabel('dialogTitle') + ' ' + SCRIPT_VERSION);
		settingsDialog.alignChildren = "fill";
		settingsDialog.margins = 16;
		settingsDialog.spacing = 12;

		var dialogUi = {};
		buildModeControls(settingsDialog, dialogUi);
		buildFolderPickerControls(settingsDialog, dialogUi);
		buildTargetPanel(settingsDialog, dialogUi);
		buildSaveOptionsPanel(settingsDialog, dialogUi);
		buildFilenamePanel(settingsDialog, dialogUi);
		buildDialogButtons(settingsDialog);
		bindDialogEvents(dialogUi);

		if (settingsDialog.show() !== 1) return null;
		return readDialogSettings(dialogUi);
	}

	/* フォルダー指定UIを作成 / Build folder picker controls */
	function buildFolderPickerControls(settingsDialog, dialogUi) {
		var FOLDER_BUTTON_WIDTH = 90;
		var FOLDER_LABEL_WIDTH = 320;

		var folderPanel = settingsDialog.add("panel", undefined, getLabel('folderPanel'));
		applyPanelLayout(folderPanel, 6);

		var sourceGroup = folderPanel.add("group");
		dialogUi.sourceButton = sourceGroup.add("button", undefined, getLabel('setSource'));
		dialogUi.sourceButton.preferredSize.width = FOLDER_BUTTON_WIDTH;
		dialogUi.sourcePathLabel = sourceGroup.add("statictext", undefined, getLabel('sourceNotSet'));
		dialogUi.sourcePathLabel.preferredSize.width = FOLDER_LABEL_WIDTH;
		setHelpTip(dialogUi.sourceButton, 'tipSourceFolder');
		setHelpTip(dialogUi.sourcePathLabel, 'tipSourceFolder');
		dialogUi.pickedSourceFolder = null;

		var destinationGroup = folderPanel.add("group");
		dialogUi.destinationButton = destinationGroup.add("button", undefined, getLabel('setDestination'));
		dialogUi.destinationButton.preferredSize.width = FOLDER_BUTTON_WIDTH;
		dialogUi.destinationPathLabel = destinationGroup.add("statictext", undefined, getLabel('destinationSameAsSource'));
		dialogUi.destinationPathLabel.preferredSize.width = FOLDER_LABEL_WIDTH;
		setHelpTip(dialogUi.destinationButton, 'tipDestinationFolder');
		setHelpTip(dialogUi.destinationPathLabel, 'tipDestinationFolder');
		dialogUi.pickedDestinationFolder = null;

		//　保存先＝対象フォルダー切替 / Same-as-source toggle
		var sameAsSourceGroup = folderPanel.add("group");
		sameAsSourceGroup.orientation = "row";
		sameAsSourceGroup.alignChildren = ["left", "center"];
		dialogUi.sameAsSourceCheckbox = sameAsSourceGroup.add("checkbox", undefined, getLabel('sameAsSource'));
		dialogUi.sameAsSourceCheckbox.value = true;
		setHelpTip(dialogUi.sameAsSourceCheckbox, 'tipSameAsSource');

		//　パス表示オプション / Path display options
		var pathOptGroup = folderPanel.add("group");
		pathOptGroup.orientation = "row";
		pathOptGroup.alignChildren = ["left", "center"];
		pathOptGroup.margins = [0, 10, 0, 0];
		dialogUi.fullPathCheckbox = pathOptGroup.add("checkbox", undefined, getLabel('fullPathCheck'));
		dialogUi.fullPathCheckbox.value = false;
		dialogUi.dropboxCheckbox = pathOptGroup.add("checkbox", undefined, getLabel('dropboxCheck'));
		dialogUi.dropboxCheckbox.value = true;
		setHelpTip(dialogUi.fullPathCheckbox, 'tipFullPath');
		setHelpTip(dialogUi.dropboxCheckbox, 'tipDropbox');
		//　Dropbox 短縮 ON のときフルパスは意味がないので無効化 / Disable full-path toggle when Dropbox shortening overrides
		dialogUi.fullPathCheckbox.enabled = !dialogUi.dropboxCheckbox.value;
	}

	/* 対象ファイル種別パネルを作成 / Build target file-type panel */
	function buildTargetPanel(settingsDialog, dialogUi) {
		var targetPanel = settingsDialog.add("panel", undefined, getLabel('targetPanel'));
		applyPanelLayout(targetPanel, 6);

		var targetGroup = targetPanel.add("group");
		targetGroup.orientation = "row";
		targetGroup.alignChildren = ["left", "center"];

		dialogUi.aiCheckbox = targetGroup.add("checkbox", undefined, ".ai");
		dialogUi.aiCheckbox.value = true;
		setHelpTip(dialogUi.aiCheckbox, 'tipTargetAi');

		dialogUi.svgCheckbox = targetGroup.add("checkbox", undefined, ".svg");
		dialogUi.svgCheckbox.value = true;
		setHelpTip(dialogUi.svgCheckbox, 'tipTargetSvg');
	}

	/* 保存設定パネルを作成 / Build save settings panel */
	function buildSaveOptionsPanel(settingsDialog, dialogUi) {
		var saveOptionsPanel = settingsDialog.add("panel", undefined, getLabel('saveOptionsPanel'));
		applyPanelLayout(saveOptionsPanel, 6);

		var versionGroup = saveOptionsPanel.add("group");
		dialogUi.versionLabel = versionGroup.add("statictext", undefined, getLabelWithColon('saveVersion'));
		dialogUi.versionDropdown = versionGroup.add("dropdownlist");
		setHelpTip(dialogUi.versionLabel, 'tipSaveVersion');
		setHelpTip(dialogUi.versionDropdown, 'tipSaveVersion');

		dialogUi.defaultVersionIndex = 0;
		for (var i = 0; i < VERSION_OPTIONS.length; i++) {
			dialogUi.versionDropdown.add("item", VERSION_OPTIONS[i].label);
			if (VERSION_OPTIONS[i].isDefault) dialogUi.defaultVersionIndex = i;
		}
		dialogUi.versionDropdown.selection = dialogUi.defaultVersionIndex;

		dialogUi.convertedCheckbox = saveOptionsPanel.add("checkbox", undefined, getLabel('appendConverted'));
		dialogUi.convertedCheckbox.value = true;
		setHelpTip(dialogUi.convertedCheckbox, 'tipAppendConverted');

		dialogUi.pdfCheckbox = saveOptionsPanel.add("checkbox", undefined, getLabel('pdfCompatible'));
		dialogUi.pdfCheckbox.value = true;
		setHelpTip(dialogUi.pdfCheckbox, 'tipPdfCompatible');
	}

	/* モード切替UI（ダイアログ最上部、中央寄せ）/ Top-level mode toggle (centered) */
	function buildModeControls(settingsDialog, dialogUi) {
		var modeGroup = settingsDialog.add("group");
		modeGroup.orientation = "row";
		modeGroup.alignChildren = ["center", "center"];
		modeGroup.alignment = ["center", "top"];
		dialogUi.overwriteRadio = modeGroup.add("radiobutton", undefined, getLabel('overwrite'));
		dialogUi.customRadio = modeGroup.add("radiobutton", undefined, getLabel('customMode'));
		dialogUi.customRadio.value = true;
		dialogUi.overwriteRadio.value = false;
		setHelpTip(dialogUi.overwriteRadio, 'tipOverwrite');
	}

	/* ファイル名調整パネルを作成 / Build filename adjustment panel */
	function buildFilenamePanel(settingsDialog, dialogUi) {
		var filenamePanel = settingsDialog.add("panel", undefined, getLabel('filenamePanel'));
		applyPanelLayout(filenamePanel, 6);
		buildSuffixControls(filenamePanel, dialogUi);
		buildCustomSuffixControls(filenamePanel, dialogUi);
		buildSeparatorControls(filenamePanel, dialogUi);
	}

	/* 保存バージョン付加UIを作成 / Build save-version suffix controls */
	function buildSuffixControls(filenamePanel, dialogUi) {
		var suffixGroup = filenamePanel.add("group");
		dialogUi.suffixCheckbox = suffixGroup.add("checkbox", undefined, getLabelWithColon('filenameSuffix'));
		dialogUi.suffixCheckbox.value = true;
		dialogUi.suffixInput = suffixGroup.add("edittext", undefined, VERSION_OPTIONS[dialogUi.defaultVersionIndex].suffix);
		dialogUi.suffixInput.characters = 12;
		setHelpTip(dialogUi.suffixCheckbox, 'tipFilenameSuffix');
		setHelpTip(dialogUi.suffixInput, 'tipFilenameSuffix');
	}

	/* 任意文字列付加UIを作成 / Build custom suffix controls */
	function buildCustomSuffixControls(filenamePanel, dialogUi) {
		var customGroup = filenamePanel.add("group");
		dialogUi.customCheckbox = customGroup.add("checkbox", undefined, getLabelWithColon('customSuffix'));
		dialogUi.customCheckbox.value = false;
		dialogUi.customInput = customGroup.add("edittext", undefined, "");
		dialogUi.customInput.characters = 12;
		dialogUi.customInput.enabled = dialogUi.customCheckbox.value;
		setHelpTip(dialogUi.customCheckbox, 'tipCustomSuffix');
		setHelpTip(dialogUi.customInput, 'tipCustomSuffix');
	}

	/* 区切り文字UIを作成 / Build separator controls */
	function buildSeparatorControls(filenamePanel, dialogUi) {
		dialogUi.separatorOptions = ["-", "_"];
		dialogUi.separatorDefaultIndex = 1;

		var separatorGroup = filenamePanel.add("group");
		dialogUi.separatorLabel = separatorGroup.add("statictext", undefined, getLabelWithColon('separator'));
		setHelpTip(dialogUi.separatorLabel, 'tipSeparator');

		dialogUi.separatorRadios = [];
		for (var sepIdx = 0; sepIdx < dialogUi.separatorOptions.length; sepIdx++) {
			var separatorRadio = separatorGroup.add("radiobutton", undefined, dialogUi.separatorOptions[sepIdx]);
			separatorRadio.value = (sepIdx === dialogUi.separatorDefaultIndex);
			setHelpTip(separatorRadio, 'tipSeparator');
			dialogUi.separatorRadios.push(separatorRadio);
		}
	}

	/* ダイアログボタンを作成 / Build dialog buttons */
	function buildDialogButtons(settingsDialog) {
		var buttonGroup = settingsDialog.add("group");
		buttonGroup.alignment = "right";
		buttonGroup.add("button", undefined, getLabel('cancel'), { name: "cancel" });
		buttonGroup.add("button", undefined, "OK", { name: "ok" });
	}

	/* ダイアログイベントを設定 / Bind dialog events */
	function bindDialogEvents(dialogUi) {
		dialogUi.sourceButton.onClick = function () {
			var selectedFolder = Folder.selectDialog(getLabel('selectSourceFolder'));
			if (selectedFolder != null) {
				dialogUi.pickedSourceFolder = selectedFolder;
				//　ツールチップは常にフルパスで確認可能 / Tooltip always shows the full path
				dialogUi.sourcePathLabel.helpTip = selectedFolder.fsName;
				refreshPathDisplays(dialogUi);
				updateTargetAvailability(dialogUi);
			}
		};

		dialogUi.destinationButton.onClick = function () {
			var selectedFolder = Folder.selectDialog(getLabel('selectDestFolder'));
			if (selectedFolder != null) {
				dialogUi.pickedDestinationFolder = selectedFolder;
				dialogUi.destinationPathLabel.helpTip = selectedFolder.fsName;
				refreshPathDisplays(dialogUi);
			}
		};

		dialogUi.fullPathCheckbox.onClick = function () {
			refreshPathDisplays(dialogUi);
		};
		dialogUi.dropboxCheckbox.onClick = function () {
			//　Dropbox 短縮 ON のときフルパスは意味がないので disable / Full-path toggle has no effect when Dropbox shortening is ON
			dialogUi.fullPathCheckbox.enabled = !dialogUi.dropboxCheckbox.value;
			refreshPathDisplays(dialogUi);
		};

		dialogUi.suffixCheckbox.onClick = function () {
			dialogUi.suffixInput.enabled = dialogUi.suffixCheckbox.value;
		};

		dialogUi.versionDropdown.onChange = function () {
			dialogUi.suffixInput.text = VERSION_OPTIONS[dialogUi.versionDropdown.selection.index].suffix;
		};

		dialogUi.customCheckbox.onClick = function () {
			dialogUi.customInput.enabled = dialogUi.customCheckbox.value;
		};

		dialogUi.overwriteRadio.onClick = function () {
			updateOverwriteState(dialogUi);
		};
		dialogUi.customRadio.onClick = function () {
			updateOverwriteState(dialogUi);
		};
		dialogUi.sameAsSourceCheckbox.onClick = function () {
			updateOverwriteState(dialogUi);
		};
		updateOverwriteState(dialogUi);
	}

	/* 対象パネルの enable 状態を更新（該当ファイルが無いとdim、ある場合はON）/ Sync target checkboxes with file availability */
	function updateTargetAvailability(dialogUi) {
		if (!dialogUi.pickedSourceFolder) {
			//　フォルダー未選択時は両方を初期状態（有効＋ON）に戻す / Reset to enabled + ON when no folder is picked
			dialogUi.aiCheckbox.enabled = true;
			dialogUi.svgCheckbox.enabled = true;
			dialogUi.aiCheckbox.value = true;
			dialogUi.svgCheckbox.value = true;
			return;
		}
		var counts = countTargetFiles(dialogUi.pickedSourceFolder);
		var hasAi = counts.ai > 0;
		var hasSvg = counts.svg > 0;
		dialogUi.aiCheckbox.enabled = hasAi;
		dialogUi.svgCheckbox.enabled = hasSvg;
		//　ファイル有無に合わせて value も同期（有ればON、無ければOFF）/ Sync value to availability
		dialogUi.aiCheckbox.value = hasAi;
		dialogUi.svgCheckbox.value = hasSvg;
	}

	/* パス表示の再計算（フルパス／Dropbox 短縮の状態を反映）/ Refresh path labels from current toggle state */
	function refreshPathDisplays(dialogUi) {
		var useTilde = !dialogUi.fullPathCheckbox.value;
		var useDropbox = dialogUi.dropboxCheckbox.value;
		if (dialogUi.pickedSourceFolder) {
			dialogUi.sourcePathLabel.text = formatDisplayPath(dialogUi.pickedSourceFolder.fsName, useTilde, useDropbox);
		}
		refreshDestinationLabel(dialogUi);
	}

	/* 保存先ラベルの表示テキストを更新（sameAsSource ON なら〈対象と同じ〉）/ Update destination label text */
	function refreshDestinationLabel(dialogUi) {
		var isOverwrite = dialogUi.overwriteRadio.value;
		var isSameAsSource = dialogUi.sameAsSourceCheckbox.value;
		if (isSameAsSource && !isOverwrite) {
			dialogUi.destinationPathLabel.text = getLabel('destinationSameAsSource');
			dialogUi.destinationPathLabel.helpTip = dialogUi.pickedSourceFolder ? dialogUi.pickedSourceFolder.fsName : getLabel('tipSameAsSource');
			return;
		}
		if (dialogUi.pickedDestinationFolder) {
			var useTilde = !dialogUi.fullPathCheckbox.value;
			var useDropbox = dialogUi.dropboxCheckbox.value;
			dialogUi.destinationPathLabel.text = formatDisplayPath(dialogUi.pickedDestinationFolder.fsName, useTilde, useDropbox);
			dialogUi.destinationPathLabel.helpTip = dialogUi.pickedDestinationFolder.fsName;
		} else {
			dialogUi.destinationPathLabel.text = getLabel('destinationNotSet');
			dialogUi.destinationPathLabel.helpTip = getLabel('tipDestinationFolder');
		}
	}

	/* モード状態のUIを更新（上書きモード／保存先=対象 を反映）/ Update UI state for mode + same-as-source */
	function updateOverwriteState(dialogUi) {
		var isOverwrite = dialogUi.overwriteRadio.value;
		var isSameAsSource = dialogUi.sameAsSourceCheckbox.value;
		var hideDestination = isOverwrite || isSameAsSource;
		if (isOverwrite) {
			dialogUi.convertedCheckbox.value = false;
			dialogUi.suffixCheckbox.value = false;
			dialogUi.customCheckbox.value = false;
		} else {
			dialogUi.suffixCheckbox.value = true;
		}
		dialogUi.convertedCheckbox.enabled = !isOverwrite;
		dialogUi.suffixCheckbox.enabled = !isOverwrite;
		dialogUi.suffixInput.enabled = !isOverwrite && dialogUi.suffixCheckbox.value;
		dialogUi.customCheckbox.enabled = !isOverwrite;
		dialogUi.customInput.enabled = !isOverwrite && dialogUi.customCheckbox.value;
		dialogUi.separatorLabel.enabled = !isOverwrite;
		for (var i = 0; i < dialogUi.separatorRadios.length; i++) {
			dialogUi.separatorRadios[i].enabled = !isOverwrite;
		}
		//　上書きモードでは sameAsSource は暗黙的に true なので操作不可 / In overwrite mode, sameAsSource is implicit
		dialogUi.sameAsSourceCheckbox.enabled = !isOverwrite;
		dialogUi.destinationButton.enabled = !hideDestination;
		dialogUi.destinationPathLabel.enabled = !hideDestination;
		refreshDestinationLabel(dialogUi);
	}

	/* 選択中の区切り文字を取得 / Get selected separator */
	function getSelectedSeparator(dialogUi) {
		for (var i = 0; i < dialogUi.separatorRadios.length; i++) {
			if (dialogUi.separatorRadios[i].value) return dialogUi.separatorOptions[i];
		}
		return dialogUi.separatorOptions[dialogUi.separatorDefaultIndex];
	}

	/* ダイアログ設定を読み取る / Read dialog settings */
	function readDialogSettings(dialogUi) {
		var selectedVersion = VERSION_OPTIONS[dialogUi.versionDropdown.selection.index];
		var selectedSeparator = getSelectedSeparator(dialogUi);
		return {
			compatibility: selectedVersion.compatibility,
			suffix: dialogUi.suffixCheckbox.value ? dialogUi.suffixInput.text : "",
			customSuffix: dialogUi.customCheckbox.value ? dialogUi.customInput.text : "",
			separator: selectedSeparator,
			pdfCompatible: dialogUi.pdfCheckbox.value,
			appendConverted: dialogUi.convertedCheckbox.value,
			overwrite: dialogUi.overwriteRadio.value,
			sameAsSource: dialogUi.sameAsSourceCheckbox.value,
			targetAi: dialogUi.aiCheckbox.value,
			targetSvg: dialogUi.svgCheckbox.value,
			sourceFolder: dialogUi.pickedSourceFolder,
			destinationFolder: dialogUi.pickedDestinationFolder
		};
	}
})();