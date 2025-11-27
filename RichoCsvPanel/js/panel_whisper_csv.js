// panel_whisper_csv.js
// WhisperX CSV → MOGRT 自動配置用 CEP パネル側 JS

(function () {
  'use strict';

  var cs = new CSInterface();

  function setLog(msg) {
    var el = document.getElementById("log");
    if (el) {
      el.textContent = msg;
    }
  }

  function placeFromCsv() {
    var select = document.getElementById("vTrackSelect");
    var vIndex = 0;
    if (select) {
      var vTrackValue = select.value;
      vIndex = parseInt(vTrackValue, 10) - 1; // V1→0
      if (isNaN(vIndex) || vIndex < 0) {
        vIndex = 0;
      }
    }

    // まずフロント側が生きてるか確認用
    alert("CSV → MOGRT ボタン押されたよ。\nvideoTrackIndex=" + vIndex);
    setLog("placeMogrtsFromWhisperCsv 実行中…");

    // ExtendScript を呼ぶ（結果はコールバックで受け取る）
    cs.evalScript("placeMogrtsFromWhisperCsv(" + vIndex + ")", function (result) {
      if (result) {
        alert("ExtendScript result: " + result);
        setLog(result);
      } else {
        setLog("完了（戻り値なし）");
      }
    });
  }

  function init() {
    var btn = document.getElementById("btnCsvToMogrt");
    if (btn) {
      btn.addEventListener("click", placeFromCsv);
    }
  }

  document.addEventListener("DOMContentLoaded", init);
}());
