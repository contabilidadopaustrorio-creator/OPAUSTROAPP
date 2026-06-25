// ─────────────────────────────────────────────────────────────
// auth.js  –  OPAUSTRO módulo de autenticación compartido v4.0
// Incluir en el <head> de cada módulo.
// ─────────────────────────────────────────────────────────────

var OPA = (function () {
  // URL del Apps Script unificado (actualizar tras nuevo despliegue)
  var GAS = 'https://script.google.com/macros/s/AKfycbzBzl8XzPUuffEPu-sbyQXTTvuCAUSL9wUwLTbciUJrfJj7pkOhmZ2p-HAnFRBT5wVr/exec';
  // URL del portal de login
  var PORTAL = 'https://contabilidadopaustrorio-creator.github.io/OPAUSTROAPP/';

  var _sess = null;

  // ── Inicializa auth. Llama desde cada módulo indicando su nombre.
  // module: 'logistica' | 'ventas' | 'gerencia'
  // Devuelve Promise<session> o redirige al portal si no autorizado.
  async function init(module) {
    // 1. Token en hash URL (navegación desde portal o nueva pestaña)
    var hash  = window.location.hash || '';
    var match = hash.match(/[#&]t=([A-Za-z0-9]+)/);
    var token = null;

    if (match) {
      token = match[1];
      // Limpiar hash para no exponer token en barra de dirección
      history.replaceState(null, '', window.location.pathname + window.location.search);
      sessionStorage.setItem('opa_token', token);
    } else {
      token = sessionStorage.getItem('opa_token');
    }

    if (!token) { _goPortal(); return null; }

    // 2. Cache local de sesión (30 min para no llamar GAS en cada página)
    var cached = _getCachedSess();
    if (cached) {
      // Verificar que el módulo está permitido
      if (module && cached.modulos.indexOf(module) === -1) { _goPortal(); return null; }
      _sess = cached;
      return _sess;
    }

    // 3. Validar con GAS
    try {
      var r = await fetch(GAS + '?action=validateSession&token=' + encodeURIComponent(token), { redirect: 'follow' });
      var d = await r.json();
      if (!d.valid) { logout(); return null; }
      if (module && d.modulos.indexOf(module) === -1) { _goPortal(); return null; }

      _sess = {
        token: token, nombre: d.nombre, rol: d.rol,
        modulos: d.modulos, puedeEscribir: d.puedeEscribir,
        exp: Date.now() + 30 * 60 * 1000 // cache 30 min en cliente
      };
      sessionStorage.setItem('opa_sess', JSON.stringify(_sess));
      return _sess;
    } catch (e) {
      // Sin red: usar caché local si existe; si no, portar
      var fallback = _getCachedSess(true);
      if (fallback && (!module || fallback.modulos.indexOf(module) !== -1)) {
        _sess = fallback;
        return _sess;
      }
      _goPortal();
      return null;
    }
  }

  // ── Aplica restricciones visuales según rol ─────────────────
  // Elementos con clase "write-only" o atributo data-write-only
  // se ocultan cuando el usuario no puede escribir.
  function applyRoleUI() {
    if (!_sess) return;
    if (!_sess.puedeEscribir) {
      document.querySelectorAll('.write-only, [data-write-only]').forEach(function (el) {
        el.style.display = 'none';
      });
      document.querySelectorAll('.readonly-badge').forEach(function (el) {
        el.style.display = '';
      });
    }
  }

  // ── Devuelve el token actual ─────────────────────────────────
  function token()         { return _sess ? _sess.token : sessionStorage.getItem('opa_token'); }
  function session()       { return _sess; }
  function puedeEscribir() { return _sess ? _sess.puedeEscribir : false; }

  // ── Cierra sesión ────────────────────────────────────────────
  function logout() {
    var t = token();
    sessionStorage.removeItem('opa_token');
    sessionStorage.removeItem('opa_sess');
    _sess = null;
    if (t) {
      try { fetch(GAS + '?action=logout&token=' + encodeURIComponent(t), { redirect: 'follow', keepalive: true }); } catch (e) {}
    }
    _goPortal();
  }

  // ── Agrega token a un objeto de payload para fetch POST ──────
  function addToken(payload) {
    return Object.assign({}, payload, { token: token() });
  }

  function gasUrl() { return GAS; }

  function _goPortal() {
    window.location.replace(PORTAL);
  }

  function _getCachedSess(allowExpired) {
    try {
      var raw = sessionStorage.getItem('opa_sess');
      if (!raw) return null;
      var s = JSON.parse(raw);
      if (!allowExpired && Date.now() > s.exp) { sessionStorage.removeItem('opa_sess'); return null; }
      return s;
    } catch (e) { return null; }
  }

  // ── Interceptor global: añade token a todos los POST al GAS ──
  // Se activa en cuanto auth.js se carga, antes del app code.
  (function patchFetch() {
    var _orig = window.fetch;
    window.fetch = function(url, opts) {
      if (typeof url === 'string' && url.indexOf('script.google.com') !== -1 &&
          opts && opts.method === 'POST' && opts.body) {
        try {
          var body = JSON.parse(opts.body);
          if (!body.token) {
            body.token = sessionStorage.getItem('opa_token') || '';
            opts = Object.assign({}, opts, { body: JSON.stringify(body) });
          }
        } catch(e) {}
      }
      return _orig.call(this, url, opts);
    };
  })();

  return { init: init, applyRoleUI: applyRoleUI, token: token, session: session, puedeEscribir: puedeEscribir, logout: logout, addToken: addToken, gasUrl: gasUrl };
})();
