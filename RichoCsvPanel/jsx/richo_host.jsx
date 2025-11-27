#target premierepro

// hostscript_whisper_csv.jsx を同じ jsx フォルダから読み込む
(function () {
    try {
        var scriptFolder = new File($.fileName).parent; // jsx フォルダ
        $.evalFile(scriptFolder.fsName + "/hostscript_whisper_csv.jsx");
    } catch (e) {
        alert("hostscript_whisper_csv.jsx の読み込みに失敗しました:\n" + e);
    }
})();

// 単純テスト：アラート
function richo_testAlert() {
    alert("Richo CSV Panel から実行できたよ！");
    return "ok";
}

// ==== ここを自分の MOGRT に合わせて変更 ====
// V12 に載せたいテロップ用 .mogrt ファイル
// 例: D:/AI_ver1/mogrt/richo_caption.mogrt
var RICHO_MOGRT_PATH = "D:/AI_ver1/mogrt/richo_caption.mogrt";
// =============================================

var RICHO_TEXT_PARAM_LABELS = [
    "ソーステキスト",
    "Source Text",
    "テキスト",
    "Text"
];

function _richo_findTextParam(props) {
    for (var i = 0; i < RICHO_TEXT_PARAM_LABELS.length; i++) {
        try {
            var p = props.getParamForDisplayName(RICHO_TEXT_PARAM_LABELS[i]);
            if (p) return p;
        } catch (e) {}
    }
    for (var j = 0; j < props.numItems; j++) {
        var cand = props[j];
        if (!cand) continue;
        try {
            var v = cand.getValue();
            if (v !== null && (typeof v === "string" || v instanceof String)) {
                return cand;
            }
        } catch (e2) {}
    }
    return null;
}

function richo_insertHello() {
    var seq = app.project.activeSequence;
    if (!seq) {
        alert("アクティブなシーケンスがありません。");
        return "no_sequence";
    }

    var vTracks = seq.videoTracks;
    if (vTracks.numTracks <= 11) {
        alert("V12 が存在しません。（ビデオトラックを増やしてください）");
        return "no_V12";
    }

    var mogrtFile = new File(RICHO_MOGRT_PATH);
    if (!mogrtFile.exists) {
        alert("MOGRT が見つかりません:\n" + RICHO_MOGRT_PATH);
        return "no_mogrt";
    }

    var curPos = seq.getPlayerPosition();
    var trackItem = seq.importMGT(mogrtFile.fsName, curPos.ticks, 11, 0);
    if (!trackItem) {
