// URL Web App Apps Script (hasil Deploy > Manage deployments, berakhiran /exec).
const API_URL = 'https://script.google.com/macros/s/AKfycbwSY2zHSIxwDTi7Y1gjjF8SSFade8xJSsjv_9tb47tSqyzx0hNiX38oIFLc1y6oql4OXQ/exec';

// Pengganti google.script.run supaya script.js tetap sama persis.
(function () {
  function panggil(fn, args) {
    return fetch(API_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },   // text/plain = tanpa preflight CORS
      body: JSON.stringify({ fn: fn, args: args }),
      redirect: 'follow'
    })
      .then(function (r) { return r.json(); })
      .then(function (j) {
        if (!j.ok) throw new Error(j.error || 'Terjadi kesalahan di server');
        return j.data;
      });
  }
  function runner(ok, fail) {
    return new Proxy({}, {
      get: function (_, nama) {
        if (nama === 'withSuccessHandler') return function (f) { return runner(f, fail); };
        if (nama === 'withFailureHandler') return function (f) { return runner(ok, f); };
        return function () {
          const args = Array.prototype.slice.call(arguments);
          panggil(String(nama), args).then(
            function (d) { if (ok) ok(d); },
            function (e) { if (fail) fail(e); else console.error(e); }
          );
        };
      }
    });
  }
  window.google = window.google || {};
  window.google.script = window.google.script || {};
  window.google.script.run = runner(null, null);
})();
