// hostscript_whisper_csv.jsx
// WhisperX の CSV (start,end,speaker,text) → MOGRT を自動配置する ExtendScript

// ========================
// 設定
// ========================

// WhisperX の出力フォルダ
var WHISPER_OUTPUT_DIR = "D:/AI_ver1/Whisper_1_0/output/";

// ログ出力
var LOG_DIR  = "D:/AI_ver1/Logs/";
var LOG_FILE = LOG_DIR + "cep_whisper_mogrt_log.txt";

// スピーカーごとの MOGRT テンプレ
var MOGRT_FOR_SPEAKER = {
    "SPEAKER_0": "D:/AI_ver1/mogrt/A1_subtitle.mogrt",
    "SPEAKER_1": "D:/AI_ver1/mogrt/A2_subtitle.mogrt",
    "":          "D:/AI_ver1/mogrt/Narration_subtitle.mogrt"
};

// デフォルト MOGRT
var DEFAULT_MOGRT = "D:/AI_ver1/mogrt/A1_subtitle.mogrt";

// Essential Graphics で公開したテキストパラメータ名
var TEXT_PARAM_NAME      = "Text";
// 1つの MOGRT でキャラを切り替えたい場合のスライダー名（使わなければそのままでOK）
var CHAR_TYPE_PARAM_NAME = "CharType";


// ========================
// ログ
// ========================

function whisper_ensureLogDir() {
    var folder = new Folder(LOG_DIR);
    if (!folder.exists) {
        folder.create();
    }
}

function whisper_writeLog(level, code, message) {
    try {
        whisper_ensureLogDir();
        var f = new File(LOG_FILE);
        if (!f.open("a")) {
            return;
        }
        var now = new Date();
        var ts =
            now.getFullYear() + "-" +
            (now.getMonth() + 1) + "-" +
            now.getDate() + " " +
            now.getHours() + ":" +
            now.getMinutes() + ":" +
            now.getSeconds();
        f.write("[" + ts + "][" + level + "][" + code + "] " + message + "\n");
        f.close();
    } catch (e) {
        // ログ失敗は無視
    }
}

function whisper_logInfo(code, msg)  { whisper_writeLog("INFO",  code, msg); }
function whisper_logWarn(code, msg)  { whisper_writeLog("WARN",  code, msg); }
function whisper_logError(code, msg) { whisper_writeLog("ERROR", code, msg); }


// ========================
// ユーティリティ
// ========================

function whisper_secondsToTime(sec) {
    var t = new Time();
    t.seconds = sec;
    return t;
}

function whisper_getStemFromMediaPath(mediaPath) {
    var norm = mediaPath.replace(/\\/g, "/");
    var parts = norm.split("/");
    var name = parts[parts.length - 1];
    return name.replace(/\.[^\.]+$/, "");
}

function whisper_getCsvPathForSelectedItem() {
    var sel = app.project.getSelection();
    if (!sel || sel.length === 0) {
        alert("プロジェクトパネルで元クリップを1つ選択してから実行してね。");
        whisper_logWarn("NO_SELECTION", "No project item selected.");
        return null;
    }
    var item = sel[0];
    if (!item.getMediaPath) {
        alert("このアイテムにはメディアパスがありません。");
        whisper_logError("NO_MEDIA_PATH_FN", "item has no getMediaPath. name=" + item.name);
        return null;
    }
    var mediaPath = item.getMediaPath();
    if (!mediaPath || mediaPath === "") {
        alert("このアイテムにはメディアパスがありません。");
        whisper_logError("NO_MEDIA_PATH", "mediaPath empty. name=" + item.name);
        return null;
    }

    var stem = whisper_getStemFromMediaPath(mediaPath);
    var csvPath = WHISPER_OUTPUT_DIR + stem + ".csv";
    var f = new File(csvPath);
    if (!f.exists) {
        alert("対応する CSV が見つかりませんでした。\n" + csvPath);
        whisper_logError("CSV_NOT_FOUND", "csvPath=" + csvPath);
        return null;
    }

    whisper_logInfo("CSV_FOUND", "csvPath=" + csvPath);
    return csvPath;
}

function whisper_readTextFile(path) {
    var f = new File(path);
    if (!f.exists) {
        whisper_logError("FILE_NOT_FOUND", "path=" + path);
        return "";
    }
    if (!f.open("r")) {
        whisper_logError("FILE_OPEN_FAILED", "path=" + path);
        return "";
    }
    var txt = f.read();
    f.close();
    return txt;
}

// WhisperX の CSV をパース
function whisper_parseWhisperCsv(csvText) {
    var lines = csvText.split(/\r\n|\n|\r/);
    var segments = [];
    var i;

    for (i = 1; i < lines.length; i++) { // 0行目はヘッダ前提
        var line = lines[i];
        if (!line || /^\s*$/.test(line)) {
            continue;
        }

        // 例: 0.000,2.345,SPEAKER_0,"こんにちは, りぃちょです"
        var firstQuote = line.indexOf(",\"");
        if (firstQuote < 0) {
            whisper_logWarn("CSV_PARSE_SKIP", "no ,\" in line " + (i + 1) + ": " + line);
            continue;
        }

        var head     = line.substring(0, firstQuote);
        var textPart = line.substring(firstQuote + 2); // ," の後ろ

        // 先頭と末尾の " を外す
        if (textPart.charAt(0) === "\"") {
            textPart = textPart.substring(1);
        }
        if (textPart.charAt(textPart.length - 1) === "\"") {
            textPart = textPart.substring(0, textPart.length - 1);
        }
        // "" → "
        textPart = textPart.replace(/""/g, "\"");

        var cols = head.split(",");
        if (cols.length < 3) {
            whisper_logWarn("CSV_PARSE_SHORT", "too few columns at line " + (i + 1) + ": " + line);
            continue;
        }

        var startSec = parseFloat(cols[0]);
        var endSec   = parseFloat(cols[1]);
        var speaker  = cols[2];

        if (isNaN(startSec) || isNaN(endSec)) {
            whisper_logWarn("CSV_TIME_PARSE_FAIL", "invalid time at line " + (i + 1) + ": " + line);
            continue;
        }

        var duration = endSec - startSec;
        if (duration <= 0) {
            whisper_logWarn("CSV_NEG_DURATION", "non-positive duration at line " + (i + 1) + ": " + line);
            duration = 0.1;
        }

        segments.push({
            start:    startSec,
            end:      endSec,
            duration: duration,
            speaker:  speaker,
            text:     textPart
        });
    }

    whisper_logInfo("CSV_PARSED", "segments=" + segments.length);
    return segments;
}

function whisper_getMogrtPathForSpeaker(speaker) {
    if (MOGRT_FOR_SPEAKER.hasOwnProperty(speaker)) {
        return MOGRT_FOR_SPEAKER[speaker];
    }
    return DEFAULT_MOGRT;
}

function whisper_setCharTypeBySpeaker(mgtComp, speaker) {
    if (!CHAR_TYPE_PARAM_NAME) {
        return;
    }
    if (!mgtComp || !mgtComp.properties || !mgtComp.properties.getParamForDisplayName) {
        return;
    }
    var p = mgtComp.properties.getParamForDisplayName(CHAR_TYPE_PARAM_NAME);
    if (!p) {
        return;
    }
    var v = 0;
    if (speaker === "SPEAKER_1") {
        v = 1;
    }
    try {
        p.setValue(v, 1);
    } catch (e) {
        whisper_logError("CHAR_PARAM_SET_FAIL", "" + e);
    }
}

function whisper_insertSingleMogrtFromSegment(seq, seg, videoTrackIndex) {
    var mogrtPath = whisper_getMogrtPathForSpeaker(seg.speaker);
    var mogrtFile = new File(mogrtPath);
    if (!mogrtFile.exists) {
        whisper_logError("MOGRT_NOT_FOUND", "mogrtPath=" + mogrtPath);
        alert("MOGRT が見つかりません:\n" + mogrtPath);
        return;
    }

    var startTime = whisper_secondsToTime(seg.start);
    var trackItem = seq.importMGT(mogrtFile.fsName, startTime.ticks, videoTrackIndex, -1);
    if (!trackItem) {
        whisper_logError("MOGRT_IMPORT_FAILED", "mogrtPath=" + mogrtPath);
        alert("importMGT に失敗しました:\n" + mogrtPath);
        return;
    }

    var newEnd = new Time();
    newEnd.seconds = startTime.seconds + seg.duration;
    trackItem.end = newEnd;

    var comp = null;
    if (trackItem.getMGTComponent) {
        comp = trackItem.getMGTComponent();
    }
    if (!comp && trackItem.components && trackItem.components.numItems > 0) {
        comp = trackItem.components[0];
    }
    if (!comp) {
        whisper_logError("MGT_COMPONENT_NULL", "mogrtPath=" + mogrtPath);
        return;
    }

    var props = comp.properties;
    if (!props || !props.getParamForDisplayName) {
        whisper_logError("MGT_PROPERTIES_INVALID", "mogrtPath=" + mogrtPath);
        return;
    }

    var textParam = props.getParamForDisplayName(TEXT_PARAM_NAME);
    if (!textParam) {
        whisper_logWarn("TEXT_PARAM_NOT_FOUND",
            "displayName=" + TEXT_PARAM_NAME + ", mogrtPath=" + mogrtPath);
    } else {
        try {
            textParam.setValue(seg.text, 1);
        } catch (e) {
            whisper_logError("TEXT_PARAM_SET_FAIL", "" + e);
        }
    }

    // CharType があれば speaker に応じて変更
    whisper_setCharTypeBySpeaker(comp, seg.speaker);
}

// エントリポイント：CSV → MOGRT
// videoTrackIndex: 0=V1, 1=V2...
function placeMogrtsFromWhisperCsv(videoTrackIndex) {
    try {
        var seq = app.project.activeSequence;
        if (!seq) {
            alert("アクティブなシーケンスがありません。");
            whisper_logError("NO_SEQUENCE", "no activeSequence");
            return "NO_SEQUENCE";
        }

        var csvPath = whisper_getCsvPathForSelectedItem();
        if (!csvPath) {
            return "NO_CSV";
        }

        var csvText = whisper_readTextFile(csvPath);
        if (!csvText) {
            alert("CSV が空、または読み込みに失敗しました。");
            whisper_logError("CSV_EMPTY", "path=" + csvPath);
            return "CSV_EMPTY";
        }

        var segments = whisper_parseWhisperCsv(csvText);
        if (!segments || segments.length === 0) {
            alert("CSV からセグメントを取得できませんでした。");
            whisper_logError("NO_SEGMENTS", "path=" + csvPath);
            return "NO_SEGMENTS";
        }

        var placed = 0;
        for (var i = 0; i < segments.length; i++) {
            var seg = segments[i];
            if (seg.duration < 0.15) {
                whisper_logWarn("SEGMENT_TOO_SHORT",
                    "start=" + seg.start + ", duration=" + seg.duration);
                continue;
            }
            whisper_insertSingleMogrtFromSegment(seq, seg, videoTrackIndex);
            placed++;
        }

        whisper_logInfo("PLACE_DONE", "placed=" + placed + ", csv=" + csvPath);
        return "OK: placed " + placed + " segments";
    } catch (e) {
        whisper_logError("UNCAUGHT", "" + e);
        alert("placeMogrtsFromWhisperCsv 中にエラー:\n" + e);
        return "ERROR: " + e;
    }
}


// ========================
// ★★★ Debug Wrapper（追加部分）★★★★★
// ========================

function debug_placeMogrtsWrapper(videoTrackIndex) {
    try {
        // 関数が定義されているかチェック
        if (typeof placeMogrtsFromWhisperCsv !== "function") {
            var t = typeof placeMogrtsFromWhisperCsv;
            whisper_logError("NO_FUNC", "typeof placeMogrtsFromWhisperCsv = " + t);
            return "NO_FUNC: typeof placeMogrtsFromWhisperCsv = " + t;
        }

        // 実行
        var result = placeMogrtsFromWhisperCsv(videoTrackIndex);
        whisper_logInfo("WRAPPER_OK", "Result=" + result);
        return "OK_CALL: " + result;

    } catch (e) {
        // ここに来れば ExtendScript の内部エラーをキャッチできる
        whisper_logError("WRAPPER_ERROR", "" + e);
        return "ERROR_IN_placeMogrtsFromWhisperCsv: " + e;
    }
}
