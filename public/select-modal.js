// 指導老師是否出席：以彈出視窗取代原生下拉選單
// 原生 <select id="teacherPresent"> 仍保留（隱藏），值與 change 事件照舊，app.js 不需修改
(function () {
  'use strict';

  var select = document.getElementById('teacherPresent');
  var btn = document.getElementById('teacherPresentBtn');
  var modal = document.getElementById('teacherPresentModal');
  var list = document.getElementById('teacherPresentOptions');
  if (!select || !btn || !modal || !list) return;

  var items = list.querySelectorAll('.option-item');

  function sync() {
    var v = select.value;
    btn.textContent = v;
    items.forEach(function (el) {
      el.classList.toggle('selected', el.getAttribute('data-value') === v);
    });
  }

  // app.js 若用程式直接設定 select.value（例如重設表單），按鈕文字也會跟著更新
  var proto = Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value');
  Object.defineProperty(select, 'value', {
    configurable: true,
    get: function () { return proto.get.call(this); },
    set: function (val) { proto.set.call(this, val); sync(); }
  });
  select.addEventListener('change', sync);

  function open() { sync(); modal.classList.add('open'); }
  function close() { modal.classList.remove('open'); }

  function choose(value) {
    if (select.value !== value) {
      select.value = value;
      select.dispatchEvent(new Event('change', { bubbles: true }));
    }
    sync();
    close();
  }

  btn.addEventListener('click', open);

  items.forEach(function (el) {
    el.addEventListener('click', function () { choose(el.getAttribute('data-value')); });
    el.addEventListener('keydown', function (e) {
      if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        choose(el.getAttribute('data-value'));
      }
    });
  });

  // 點背景關閉、按 Esc 關閉
  modal.addEventListener('click', function (e) { if (e.target === modal) close(); });
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape' && modal.classList.contains('open')) close();
  });

  sync();
})();
