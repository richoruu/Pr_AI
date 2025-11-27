(function () {
  var cs;

  function log(msg) {
    var el = document.getElementById("log");
    el.textContent += msg + "\n";
  }

  function onLoaded() {
    cs = new CSInterface();

    // TEST ボタン：単純にアラートを出す
    document.getElementById("btnTest").addEventListener("click", function () {
      cs.evalScript("richo_testAlert()", function (result) {
        log("testAlert → " + result);
      });
    });

    // HELLO ボタン：V12 に MOGRT を挿入してテキストを「ハロー」に
    document.getElementById("btnHello").addEventListener("click", function () {
      cs.evalScript("richo_insertHello()", function (result) {
        log("insertHello → " + result);
      });
    });
  }

  document.addEventListener("DOMContentLoaded", onLoaded);
})();
