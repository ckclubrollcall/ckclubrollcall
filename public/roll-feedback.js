// 點名勾選的手感：動畫 + 震動回饋
// - Android：navigator.vibrate
// - iOS Safari（17.4+）：藏一個 <input type="checkbox" switch>，點名時用程式觸發它，系統會給觸覺回饋
// 學生列表由 app.js 動態產生，所以用事件委派，app.js 不需修改
(function () {
  'use strict';

  var list = document.getElementById('studentList');
  if (!list) return;

  // ---------- iOS：隱藏的 switch 開關，用來觸發觸覺回饋 ----------
  // 不能用 display:none（會失效），改用移出畫面 + 透明 + 不可點擊
  var hapticLabel = document.createElement('label');
  hapticLabel.setAttribute('aria-hidden', 'true');
  hapticLabel.style.cssText =
    'position:fixed;left:-9999px;top:0;width:1px;height:1px;opacity:0;pointer-events:none;overflow:hidden;';
  var hapticInput = document.createElement('input');
  hapticInput.type = 'checkbox';
  hapticInput.setAttribute('switch', '');
  hapticInput.tabIndex = -1;
  hapticLabel.appendChild(hapticInput);
  document.body.appendChild(hapticLabel);

  function iosHaptic() {
    // 必須在使用者點擊事件內同步呼叫
    try { hapticLabel.click(); } catch (e) { /* 忽略 */ }
  }

  // ---------- 震動（Android） ----------
  function buzz(pattern) {
    if ('vibrate' in navigator) {
      try { navigator.vibrate(pattern); } catch (e) { /* 忽略 */ }
    }
  }

  // ---------- 勾選動畫 ----------
  function pop(cb, checked) {
    var cls = checked ? 'pop-check' : 'pop-uncheck';
    cb.classList.remove('pop-check', 'pop-uncheck');
    void cb.offsetWidth; // 重新觸發動畫
    cb.classList.add(cls);
    cb.addEventListener('animationend', function done() {
      cb.classList.remove(cls);
      cb.removeEventListener('animationend', done);
    });
  }

  // ---------- onclick：點整列（含 label）都會觸發 ----------
  var last = 0;
  list.addEventListener('click', function (e) {
    var row = e.target.closest && e.target.closest('.student-item');
    if (!row) return;

    // 點 label 時會連帶產生一次對 checkbox 的 click，50ms 內只處理一次
    var now = Date.now();
    if (now - last < 50) return;
    last = now;

    iosHaptic(); // iOS：同步觸發（需在點擊手勢內）

    // 等瀏覽器 / app.js 完成切換後，再讀取最終的勾選狀態
    setTimeout(function () {
      var cb = row.querySelector('input[type="checkbox"]');
      if (!cb) return;
      if (cb.checked) {
        buzz(25);   // 勾選：稍明確的一下
        pop(cb, true);
      } else {
        buzz(12);   // 取消：更輕的一下
        pop(cb, false);
      }
    }, 0);
  });
})();
