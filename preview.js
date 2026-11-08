/*
  Print preview window. The panel keeps the page's details in this Mac's
  storage (lsk.preview) and opens this window; it draws the page to scale:
  each label's edge, the used labels shaded, and the text and pictures of the
  labels ready to print, laid out the way Word lays them out. The panel
  updates the details while the window is open, and the window follows.
*/
(function () {
  'use strict';
  var T = window.LSK_TOOLS;
  var KEY = 'lsk.preview';
  var shown = 0;

  function $(id) { return document.getElementById(id); }
  function read() { try { return JSON.parse(localStorage.getItem(KEY) || 'null'); } catch (e) { return null; } }
  function pct(v, of) { return (v / of * 100) + '%'; }

  // As big as the window allows, keeping the page's shape.
  function fitPage(f) {
    var page = $('page'), head = document.querySelector('.top').offsetHeight;
    var ratio = f.page.width / f.page.height;
    page.style.aspectRatio = f.page.width + ' / ' + f.page.height;
    page.style.width = 'min(calc(100vw - 32px), calc((100vh - ' + (head + 24) + 'px) * ' + ratio.toFixed(5) + '))';
  }

  function draw(data) {
    var page = $('page');
    if (!data || !data.fmt) { page.hidden = true; $('empty').hidden = false; return; }
    var f = data.fmt, used = {}, byN = {};
    (data.used || []).forEach(function (n) { used[n] = true; });
    (data.labels || []).forEach(function (l) { byN[l.n] = l; });
    $('sub').textContent = data.name + (data.adjusted ? ' · with your printer adjustment' : '');
    page.hidden = false; $('empty').hidden = true;
    fitPage(f);
    page.innerHTML = '';
    var pw = f.page.width, ph = f.page.height;
    for (var n = 1; n <= f.across * f.down; n++) {
      var r = Math.floor((n - 1) / f.across), c = (n - 1) % f.across, l = byN[n];
      var box = document.createElement('div');
      box.className = 'lab' + (used[n] ? ' used' : '');
      box.style.left = pct(f.firstLabel.left + c * f.pitch.across, pw);
      box.style.top = pct(f.firstLabel.top + r * f.pitch.down, ph);
      box.style.width = pct(f.label.width, pw);
      box.style.height = pct(f.label.height, ph);
      if (f.corner) box.style.borderRadius = (f.corner / pw * 100) + 'cqw';
      box.title = 'Label ' + n + (used[n] ? ', used' : l ? ', prints' : ', open');
      if (l && !used[n]) {
        var t = document.createElement('div');
        t.className = 't';
        box.appendChild(t);
        T.drawLabelText(t, l.specs || [], f);
        if (l.picture) {
          var pic = document.createElement('div');
          pic.className = 'pic'; pic.textContent = 'Picture';
          box.appendChild(pic);
        }
      }
      page.appendChild(box);
    }
    // The small "▲ top" line Word prints in the blank strip at the top.
    if (data.arrow) {
      var a = document.createElement('div'), mid = f.firstLabel.left + ((f.across - 1) * f.pitch.across + f.label.width) / 2;
      a.className = 'arrow';
      a.textContent = data.arrow;
      a.style.left = pct(mid, pw);
      a.style.top = pct(data.arrowPt.before / 72, ph);
      a.style.fontSize = (data.arrowPt.size / 72 / pw * 100) + 'cqw';
      page.appendChild(a);
    }
    shown = data.at || 0;
  }

  function close() {
    try {
      if (window.Office && Office.context && Office.context.ui && Office.context.ui.messageParent) { Office.context.ui.messageParent(JSON.stringify({ type: 'close' })); return; }
    } catch (e) { /* not in Word */ }
    window.close();
  }

  function start() {
    $('close').addEventListener('click', close);
    draw(read());
    window.addEventListener('resize', function () { var d = read(); if (d && d.fmt) fitPage(d.fmt); });
    // Follow the panel's changes.
    setInterval(function () { var d = read(); if (d && d.at && d.at !== shown) draw(d); }, 1200);
    window.LSK_PREVIEW = { draw: draw, read: read };
  }
  if (window.Office && Office.onReady) Office.onReady(function () { /* ready for messageParent */ });
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start); else start();
})();
