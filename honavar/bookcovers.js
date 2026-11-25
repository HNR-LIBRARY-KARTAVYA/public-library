/* bookcovers.js - shows the cover photo of every book on the page (search results, My Library lists, loans).
   Nothing in the page is changed; it only looks for book cards and adds a picture when the book has one. */
(function () {
  var API = 'https://script.google.com/macros/s/AKfycbygwhBJBk9AxPy2elc92uNiU5PjsmeaKqIMtvg_EO0lUlphIa9k6RQnyePyfZIP06uO7g/exec';
  var seen = {}, cache = {}, timer = null;
  var css = document.createElement('style');
  css.textContent = '.bclcv{width:64px;height:90px;object-fit:cover;border-radius:6px;border:1px solid #d9d2bd;background:#f3eee0;flex:0 0 auto;cursor:zoom-in;box-shadow:0 1px 3px rgba(0,0,0,.15)}' +
    '.bclcv-big{position:fixed;inset:0;background:rgba(0,0,0,.78);display:flex;align-items:center;justify-content:center;z-index:99999;cursor:zoom-out}' +
    '.bclcv-big img{max-width:92vw;max-height:90vh;border-radius:8px;background:#fff}';
  document.head.appendChild(css);
  function thumb(id, w) { return 'https://drive.google.com/thumbnail?id=' + encodeURIComponent(id) + '&sz=w' + w; }
  function accOf(el) { return el.getAttribute('data-acc') || el.getAttribute('data-res') || el.getAttribute('data-renew') || el.getAttribute('data-full') || ''; }
  function cardOf(el) { return el.closest('.book-card') || el.closest('.item'); }
  function put(card, id) {
    if (!card || card.querySelector('.bclcv')) return;
    var img = document.createElement('img');
    img.className = 'bclcv'; img.alt = 'Cover'; img.loading = 'lazy'; img.src = thumb(id, 200);
    img.onerror = function () { img.remove(); };
    img.onclick = function (e) {
      e.stopPropagation();
      var ov = document.createElement('div'); ov.className = 'bclcv-big';
      ov.innerHTML = '<img alt="Cover" src="' + thumb(id, 900) + '">';
      ov.onclick = function () { ov.remove(); };
      document.body.appendChild(ov);
    };
    card.insertBefore(img, card.firstChild);
  }
  function scan() {
    var jobs = [], els = document.querySelectorAll('[data-acc],[data-res],[data-renew],[data-full]');
    for (var i = 0; i < els.length; i++) {
      var a = accOf(els[i]), card = cardOf(els[i]);
      if (!a || !card || card.getAttribute('data-bclcv')) continue;
      card.setAttribute('data-bclcv', '1');
      jobs.push({ a: a, card: card });
    }
    if (!jobs.length) return;
    var need = [];
    jobs.forEach(function (j) { if (!(j.a in cache) && need.indexOf(j.a) < 0) need.push(j.a); });
    function paint() { jobs.forEach(function (j) { if (cache[j.a]) put(j.card, cache[j.a]); }); }
    if (!need.length) { paint(); return; }
    need = need.slice(0, 60);
    fetch(API, { method: 'POST', credentials: 'omit', redirect: 'follow', headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ fn: 'bclCovers', args: [need] }) })
      .then(function (r) { return r.json(); })
      .then(function (j) {
        var m = (j && j.ok && j.result) || {};
        need.forEach(function (a) { cache[a] = m[a] || ''; });
        paint();
      })
      .catch(function () { /* no cover - the page stays as it was */ });
  }
  function later() { clearTimeout(timer); timer = setTimeout(scan, 250); }
  new MutationObserver(later).observe(document.documentElement, { childList: true, subtree: true });
  later();
})();
