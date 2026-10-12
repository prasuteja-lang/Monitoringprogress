// ====================== api.js ======================
// Pengganti google.script.run supaya script.js tetap sama persis.
//
// LANGKAH WAJIB: ganti nilai API_URL di bawah dengan URL Web App Apps Script kamu
// (Deploy > Manage deployments > salin "Web app URL", berakhiran /exec).
const API_URL = 'GANTI_DENGAN_URL_WEB_APP_ANDA';

(function () {
  const PESAN_HTML =
    'Server membalas halaman HTML, bukan data. Cek di Apps Script: ' +
    '(1) Deploy > Manage deployments > Version = "New version" lalu Deploy, ' +
    '(2) "Who has access" = Anyone, (3) URL yang dipakai berakhiran /exec.';

  function panggil(fn, args) {
    if (!API_URL || API_URL.indexOf('https://script.google.com/') !== 0) {
      return Promise.reject(new Error(
        'API_URL di api.js belum diisi dengan URL Web App Apps Script (harus diawali https://script.google.com/ dan berakhiran /exec).'
      ));
    }
    return fetch(API_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },   // text/plain = tanpa preflight CORS
      body: JSON.stringify({ fn: fn, args: args }),
      redirect: 'follow'
    })
      .then(function (r) { return r.text(); })
      .then(function (teks) {
        let j;
        try { j = JSON.parse(teks); }
        catch (e) { throw new Error(/^\s*</.test(teks) ? PESAN_HTML : 'Balasan server tidak valid.'); }
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
