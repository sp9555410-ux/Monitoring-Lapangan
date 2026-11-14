/* Meniru google.script.run memakai fetch ke Apps Script (doPost). */
(function () {
  const URL_EXEC = '__URL_EXEC__';
  function make(ok, fail) {
    return new Proxy({}, { get(_, name) {
      if (name === 'withSuccessHandler') return f => make(f, fail);
      if (name === 'withFailureHandler') return f => make(ok, f);
      return (...args) => {
        let tk = localStorage.getItem('tk') || prompt('Masukkan kode akses');
        if (!tk) { fail && fail(new Error('Kode akses diperlukan.')); return; }
        fetch(URL_EXEC, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' },
          body: JSON.stringify({ fn: name, args, token: tk }) })
          .then(r => r.json())
          .then(j => {
            if (j.error) { if (/akses/i.test(j.error)) localStorage.removeItem('tk'); throw new Error(j.error); }
            localStorage.setItem('tk', tk);
            ok && ok(j.data);
          })
          .catch(e => fail ? fail(e) : console.error(e));
      };
    }});
  }
  window.google = { script: { run: make(null, null) } };
})();
