/* Mi panel
   - Tareas y secciones: Todoist (API v1, endpoint /sync)
   - Datos propios de la app (etiquetas, ...): almacén local por colecciones,
     sincronizado con OneDrive personal (carpeta de la app) si el usuario lo conecta
*/
(() => {
  'use strict';

  // Si esta página se carga dentro de la ventana de login de Microsoft, no arranca la app:
  // la ventana principal (MSAL) lee la respuesta de la dirección y cierra esta ventana
  if (window.opener && window.opener !== window && /(^|[#?&])(code|error)=/.test(location.hash + location.search)) return;

  const API_URL = 'https://api.todoist.com/api/v1/sync';
  const KEYS = {
    token: 'mp.todoist.token',
    cache: 'mp.todoist.cache',
    datos: 'mp.datos',
    base: 'mp.datos.base',
    onedrive: 'mp.onedrive',
    ui: 'mp.ui'
  };
  const LEGACY_TAGS_KEY = 'mp.tags'; // formato anterior de las etiquetas, se migra al almacén
  const SCHEMA = 1; // versión del formato del almacén
  const GRAPH_FILE = 'https://graph.microsoft.com/v1.0/me/drive/special/approot:/datos.json';
  const OD_SCOPES = ['Files.ReadWrite.AppFolder', 'User.Read'];
  const PALETTE = ['#DBF266', '#7FA2E0', '#F2B35B', '#B39DF0', '#5CCFB8', '#F0857A',
                   '#E68AC4', '#8BD17C', '#6FC3E8', '#D9A77A', '#F5D76E', '#C9D0DC'];
  const SMALL = window.matchMedia('(max-width: 1099px)');

  // ---------- almacenamiento local ----------
  const store = {
    get(key, fallback) {
      try {
        const raw = localStorage.getItem(key);
        return raw ? JSON.parse(raw) : fallback;
      } catch (e) { return fallback; }
    },
    set(key, value) {
      try { localStorage.setItem(key, JSON.stringify(value)); }
      catch (e) { toast('No se pudo guardar en este navegador.', 'error'); }
    }
  };

  // ---------- estado ----------
  const S = {
    token: store.get(KEYS.token, ''),
    data: store.get(KEYS.cache, null),
    datos: loadDatos(),
    tags: null,
    od: Object.assign({ clientId: '', account: '', etag: null, syncedAt: null }, store.get(KEYS.onedrive, {})),
    odBase: store.get(KEYS.base, null),
    odState: 'idle', // idle | syncing | login | error
    ui: Object.assign({ sidebarHidden: false, projectId: null }, store.get(KEYS.ui, {})),
    filter: new Set(),
    syncing: false,
    editing: null,
    sortables: []
  };

  function normalizeTags(t) {
    const out = { list: [], assign: {} };
    if (!t || typeof t !== 'object') return out;
    if (Array.isArray(t.list)) {
      out.list = t.list
        .filter(x => x && typeof x.id === 'string' && typeof x.name === 'string' && isHex(x.color))
        .map(x => ({ id: x.id, name: x.name.slice(0, 40), color: x.color }));
    }
    if (t.assign && typeof t.assign === 'object') {
      for (const [k, v] of Object.entries(t.assign)) {
        if (Array.isArray(v)) out.assign[k] = v.filter(id => typeof id === 'string');
      }
    }
    return out;
  }

  S.tags = tagsFromDoc(S.datos);

  // ---------- almacén de datos de la app ----------
  // Documento: { schema, colecciones: { [nombre]: { [id]: registro } } }.
  // Cada registro lleva su id; la fusión con OneDrive trabaja registro a registro.
  function emptyDoc() { return { schema: SCHEMA, colecciones: {} }; }

  function normalizeDoc(d) {
    const out = emptyDoc();
    if (!d || typeof d !== 'object') return out;
    out.schema = Number.isInteger(d.schema) ? d.schema : SCHEMA;
    const cols = d.colecciones && typeof d.colecciones === 'object' ? d.colecciones : {};
    for (const [name, recs] of Object.entries(cols)) {
      if (!recs || typeof recs !== 'object' || Array.isArray(recs)) continue;
      out.colecciones[name] = {};
      for (const [id, r] of Object.entries(recs)) {
        if (r && typeof r === 'object' && r.id === id) out.colecciones[name][id] = r;
      }
    }
    return out;
  }

  function loadDatos() {
    const saved = store.get(KEYS.datos, null);
    if (saved) return normalizeDoc(saved);
    // Primera vez con el almacén: se traen las etiquetas del formato anterior
    const doc = emptyDoc();
    writeTags(doc, normalizeTags(store.get(LEGACY_TAGS_KEY, null)));
    store.set(KEYS.datos, doc);
    return doc;
  }

  // Etiquetas en el almacén: "etiquetas" ({id, name, color}), "asignaciones" ({id: taskId, tags: [...]})
  // y el orden aparte, en "orden" ({id: 'etiquetas', ids: [...]}), para que reordenar o borrar
  // no cuente como cambio en cada etiqueta al fusionar con otro dispositivo.
  function writeTags(doc, tags) {
    const etiquetas = {}, asignaciones = {};
    for (const t of tags.list) etiquetas[t.id] = { id: t.id, name: t.name, color: t.color };
    for (const [taskId, ids] of Object.entries(tags.assign)) {
      if (taskId.startsWith('tmp-') || !ids.length) continue;
      asignaciones[taskId] = { id: taskId, tags: ids.slice() };
    }
    doc.colecciones.etiquetas = etiquetas;
    doc.colecciones.asignaciones = asignaciones;
    doc.colecciones.orden = Object.assign({}, doc.colecciones.orden, { etiquetas: { id: 'etiquetas', ids: tags.list.map(t => t.id) } });
  }

  function tagsFromDoc(doc) {
    const ord = ((doc.colecciones.orden || {}).etiquetas || {}).ids || [];
    const pos = id => { const i = ord.indexOf(id); return i < 0 ? Infinity : i; };
    const et = Object.values(doc.colecciones.etiquetas || {}).sort((a, b) => pos(a.id) - pos(b.id) || a.name.localeCompare(b.name));
    const assign = {};
    for (const r of Object.values(doc.colecciones.asignaciones || {})) assign[r.id] = r.tags;
    return normalizeTags({ list: et, assign });
  }

  // JSON con las claves ordenadas, para comparar registros sin depender del orden de las propiedades
  function stable(v) {
    if (v === undefined) return '';
    if (Array.isArray(v)) return '[' + v.map(stable).join(',') + ']';
    if (v && typeof v === 'object') return '{' + Object.keys(v).sort().map(k => JSON.stringify(k) + ':' + stable(v[k])).join(',') + '}';
    return JSON.stringify(v);
  }

  // Fusión a tres bandas, registro a registro: si solo cambió un lado gana ese cambio
  // (también los borrados); si cambiaron los dos, gana este dispositivo.
  function mergeDocs(base, local, remote) {
    const out = emptyDoc();
    out.schema = Math.max(local.schema, remote.schema);
    const names = new Set([...Object.keys(base.colecciones), ...Object.keys(local.colecciones), ...Object.keys(remote.colecciones)]);
    for (const name of names) {
      const b = base.colecciones[name] || {}, l = local.colecciones[name] || {}, r = remote.colecciones[name] || {};
      const res = {};
      for (const id of new Set([...Object.keys(b), ...Object.keys(l), ...Object.keys(r)])) {
        const sb = stable(b[id]), sl = stable(l[id]), sr = stable(r[id]);
        const pick = sl === sb ? r[id] : (sr === sb ? l[id] : l[id]);
        if (pick !== undefined) res[id] = pick;
      }
      out.colecciones[name] = res;
    }
    return out;
  }

  const saveDatos = () => { store.set(KEYS.datos, S.datos); odSchedulePush(); };
  const saveOd = () => store.set(KEYS.onedrive, S.od);
  const saveTags = () => {
    writeTags(S.datos, S.tags);
    saveDatos();
  };
  // Tras traer cambios de otro dispositivo; conserva las asignaciones de tareas aún sin crear en Todoist
  function applyDatos(doc) {
    const pendientes = Object.entries(S.tags.assign).filter(([k]) => k.startsWith('tmp-'));
    S.datos = doc;
    store.set(KEYS.datos, S.datos);
    S.tags = tagsFromDoc(S.datos);
    for (const [k, v] of pendientes) S.tags.assign[k] = v;
  }
  const saveUi = () => store.set(KEYS.ui, S.ui);
  const saveCache = () => store.set(KEYS.cache, S.data);

  // ---------- utilidades ----------
  function uid() {
    if (window.crypto && crypto.randomUUID) return crypto.randomUUID();
    return 'id-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2);
  }
  function isHex(c) { return typeof c === 'string' && /^#[0-9a-fA-F]{6}$/.test(c); }

  // Crea elementos del DOM sin usar innerHTML (los textos de Todoist nunca se interpretan como HTML)
  function h(tag, attrs, ...children) {
    const el = document.createElement(tag);
    if (attrs) {
      for (const [k, v] of Object.entries(attrs)) {
        if (v === null || v === undefined || v === false) continue;
        if (k === 'class') el.className = v;
        else if (k === 'style') el.style.cssText = v;
        else if (k === 'dataset') Object.assign(el.dataset, v);
        else if (k.startsWith('on') && typeof v === 'function') el.addEventListener(k.slice(2), v);
        else if (k === 'value') el.value = v;
        else el.setAttribute(k, v === true ? '' : v);
      }
    }
    for (const child of children.flat(Infinity)) {
      if (child === null || child === undefined || child === false) continue;
      el.append(child.nodeType ? child : document.createTextNode(String(child)));
    }
    return el;
  }

  function svg(paths, opts = {}) {
    const ns = 'http://www.w3.org/2000/svg';
    const el = document.createElementNS(ns, 'svg');
    el.setAttribute('viewBox', '0 0 24 24');
    el.setAttribute('aria-hidden', 'true');
    if (opts.cls) el.setAttribute('class', opts.cls);
    for (const p of paths) {
      const node = document.createElementNS(ns, p.tag || 'path');
      for (const [k, v] of Object.entries(p)) if (k !== 'tag') node.setAttribute(k, v);
      el.append(node);
    }
    return el;
  }
  const ICON = {
    plus: () => svg([{ d: 'M12 5v14M5 12h14' }]),
    close: () => svg([{ d: 'M6 6l12 12M18 6L6 18' }]),
    trash: () => svg([{ d: 'M4 7h16M10 11v6M14 11v6M6 7l1 13h10l1-13M9 7V4h6v3' }]),
    edit: () => svg([{ d: 'M4 20h4L19 9l-4-4L4 16v4zM13.5 6.5l4 4' }]),
    comment: () => svg([{ d: 'M20 11.5a7.5 7.5 0 0 1-10.9 6.7L4 19.5l1.3-4.6A7.5 7.5 0 1 1 20 11.5z' }]),
    tag: () => svg([{ d: 'M20.6 13.4l-7.2 7.2a2 2 0 0 1-2.8 0L3 13V3h10l7.6 7.6a2 2 0 0 1 0 2.8z' }, { tag: 'circle', cx: '7.5', cy: '7.5', r: '1.5' }]),
    grip: () => svg([
      { tag: 'circle', cx: '8', cy: '5', r: '2' }, { tag: 'circle', cx: '16', cy: '5', r: '2' },
      { tag: 'circle', cx: '8', cy: '12', r: '2' }, { tag: 'circle', cx: '16', cy: '12', r: '2' },
      { tag: 'circle', cx: '8', cy: '19', r: '2' }, { tag: 'circle', cx: '16', cy: '19', r: '2' }
    ], { cls: 'grip' })
  };

  // Quita el formato markdown de Todoist para mostrar el título limpio
  function cleanTitle(s) {
    return String(s || '')
      .replace(/\[([^\]]+)\]\([^)]*\)/g, '$1')
      .replace(/(\*\*|__)(.+?)\1/g, '$2')
      .replace(/`([^`]+)`/g, '$1')
      .trim();
  }

  function relTime(ts) {
    if (!ts) return 'nunca';
    const min = Math.round((Date.now() - ts) / 60000);
    if (min < 1) return 'hace un momento';
    if (min < 60) return 'hace ' + min + ' min';
    const hrs = Math.round(min / 60);
    if (hrs < 24) return 'hace ' + hrs + ' h';
    const days = Math.round(hrs / 24);
    return days === 1 ? 'ayer' : 'hace ' + days + ' días';
  }

  // ---------- avisos ----------
  function toast(message, kind = 'ok', action = null, ms = 5000) {
    const box = document.getElementById('toasts');
    const el = h('div', { class: 'toast' + (kind === 'error' ? ' error' : ''), role: kind === 'error' ? 'alert' : 'status' },
      h('span', null, message));
    if (action) {
      el.append(h('button', {
        type: 'button',
        onclick: () => { el.remove(); action.fn(); }
      }, action.label));
    }
    box.append(el);
    setTimeout(() => el.remove(), ms);
  }

  // ---------- API de Todoist ----------
  class ApiError extends Error {
    constructor(kind, message) { super(message); this.kind = kind; }
  }

  async function apiCall(params, token = S.token) {
    let res;
    try {
      res = await fetch(API_URL, {
        method: 'POST',
        headers: { Authorization: 'Bearer ' + token },
        body: new URLSearchParams(params)
      });
    } catch (e) {
      throw new ApiError('network', 'No se pudo conectar con Todoist. Revisa la conexión.');
    }
    if (res.status === 401 || res.status === 403) throw new ApiError('auth', 'Todoist no acepta el token. Revisa que esté bien copiado.');
    if (res.status === 429) throw new ApiError('rate', 'Demasiadas peticiones a Todoist. Espera un minuto.');
    if (!res.ok) throw new ApiError('http', 'Todoist respondió con un error (' + res.status + ').');
    return res.json();
  }

  async function fetchAll(token = S.token) {
    const r = await apiCall({
      sync_token: '*',
      resource_types: JSON.stringify(['user', 'projects', 'sections', 'items', 'notes'])
    }, token);
    const projects = (r.projects || [])
      .filter(p => !p.is_deleted && !p.is_archived)
      .sort((a, b) => (b.inbox_project ? 1 : 0) - (a.inbox_project ? 1 : 0) || (a.child_order || 0) - (b.child_order || 0));
    const inboxId = (r.user && r.user.inbox_project_id) || (projects.find(p => p.inbox_project) || {}).id || null;
    const sections = (r.sections || []).filter(s => !s.is_deleted && !s.is_archived);
    const items = (r.items || []).filter(i => !i.is_deleted && !i.checked).map(i => ({
      id: i.id,
      content: i.content,
      project_id: i.project_id,
      section_id: i.section_id || null,
      parent_id: i.parent_id || null,
      child_order: i.child_order || 0,
      recurring: !!(i.due && i.due.is_recurring)
    }));
    // Comentarios de tareas activas
    const active = new Set(items.map(i => i.id));
    const notes = (r.notes || []).filter(n => !n.is_deleted && active.has(n.item_id)).map(n => {
      const f = n.file_attachment;
      return {
        id: n.id,
        item_id: n.item_id,
        content: n.content || '',
        posted_at: n.posted_at || null,
        file: f && /^https:\/\//.test(f.file_url || '') ? { name: f.file_name || 'Archivo adjunto', url: f.file_url } : null
      };
    });
    return {
      projects: projects.map(p => ({ id: p.id, name: p.name, inbox: !!p.inbox_project })),
      sections: sections.map(s => ({ id: s.id, name: s.name, project_id: s.project_id, order: s.section_order || 0 })),
      items,
      notes,
      inboxId,
      syncedAt: Date.now()
    };
  }

  async function runCommands(commands) {
    const r = await apiCall({ commands: JSON.stringify(commands) });
    const status = r.sync_status || {};
    for (const c of commands) {
      const st = status[c.uuid];
      if (st !== 'ok') {
        const msg = st && st.error ? 'Todoist rechazó el cambio: ' + st.error : 'Todoist rechazó el cambio.';
        throw new ApiError('command', msg);
      }
    }
    return r.temp_id_mapping || {};
  }

  async function doSync(opts = {}) {
    if (S.syncing || !S.token) return;
    S.syncing = true;
    renderSyncState();
    try {
      S.data = await fetchAll();
      saveCache();
      if (!opts.silent) toast('Sincronizado con Todoist');
    } catch (e) {
      toast(e.message, 'error');
      if (e.kind === 'auth') openSetup(e.message);
    } finally {
      S.syncing = false;
      renderAll();
    }
  }

  // Si una escritura falla, se vuelve a leer Todoist para dejar el tablero como está de verdad
  async function recover(err) {
    toast(err.message, 'error');
    await doSync({ silent: true });
  }

  // ---------- OneDrive (cuenta personal de Microsoft) ----------
  // Login con MSAL.js; el fichero datos.json vive en la carpeta de la app (/Aplicaciones/<app>).
  // Escritura con If-Match (eTag): si otro dispositivo escribió antes, se vuelve a leer y se fusiona.
  class OdError extends Error {
    constructor(kind, message) { super(message); this.kind = kind; }
  }
  let msalApp = null;     // PublicClientApplication ya inicializada
  let msalLoading = null; // promesa de carga e inicialización
  let odTimer = null;
  let odAgain = false;

  const odConnected = () => !!(S.od.clientId && S.od.account);
  const appUrl = () => location.origin + location.pathname.replace(/index\.html$/, '');

  // Carga msal-browser.min.js solo si el usuario ha configurado OneDrive
  function loadMsal() {
    if (msalLoading) return msalLoading;
    msalLoading = new Promise((resolve, reject) => {
      if (window.msal) return resolve();
      const sc = document.createElement('script');
      sc.src = 'msal-browser.min.js';
      sc.onload = resolve;
      sc.onerror = () => reject(new OdError('load', 'No se pudo cargar el módulo de inicio de sesión de Microsoft.'));
      document.head.append(sc);
    }).then(async () => {
      const app = new msal.PublicClientApplication({
        auth: { clientId: S.od.clientId, authority: 'https://login.microsoftonline.com/consumers', redirectUri: appUrl() },
        cache: { cacheLocation: 'localStorage' }
      });
      await app.initialize();
      msalApp = app;
      // Vuelta de un login por redirección (si el navegador bloqueó la ventana emergente)
      const res = await app.handleRedirectPromise();
      if (res && res.account) odLoggedIn(res.account);
      return app;
    });
    msalLoading.catch(() => { msalLoading = null; });
    return msalLoading;
  }

  function odAccount() {
    if (!msalApp) return null;
    return msalApp.getActiveAccount() || msalApp.getAllAccounts()[0] || null;
  }

  function odLoggedIn(account) {
    msalApp.setActiveAccount(account);
    S.od.account = account.username || account.name || 'Cuenta de Microsoft';
    saveOd();
    S.odState = 'idle';
    renderSyncState();
    syncDatos();
  }

  // Se llama directamente desde un clic: la ventana emergente debe abrirse sin esperas previas
  function odConnect() {
    if (!msalApp) {
      toast('Preparando el inicio de sesión. Vuelve a pulsar en un momento.');
      loadMsal().catch(e => toast(e.message, 'error'));
      return;
    }
    const acc = odAccount();
    const req = { scopes: OD_SCOPES, prompt: 'select_account' };
    if (acc) req.loginHint = acc.username;
    msalApp.loginPopup(req)
      .then(res => odLoggedIn(res.account))
      .catch(e => {
        const code = e && e.errorCode;
        if (code === 'popup_window_error' || code === 'empty_window_error') return msalApp.loginRedirect(req);
        if (code === 'user_cancelled') return;
        toast('No se pudo conectar con OneDrive: ' + (e.message || code), 'error');
      });
  }

  async function odDisconnect() {
    try {
      if (!msalApp) await loadMsal();
      const acc = odAccount();
      if (acc) await msalApp.clearCache({ account: acc });
    } catch (e) { /* aunque falle, se olvida la cuenta en la app */ }
    S.od.account = '';
    S.od.etag = null;
    S.od.syncedAt = null;
    saveOd();
    S.odBase = null;
    localStorage.removeItem(KEYS.base);
    S.odState = 'idle';
    renderSyncState();
  }

  async function odToken() {
    if (!msalApp) await loadMsal();
    const acc = odAccount();
    if (!acc) throw new OdError('login', 'Conecta de nuevo OneDrive.');
    try {
      const r = await msalApp.acquireTokenSilent({ scopes: OD_SCOPES, account: acc });
      return r.accessToken;
    } catch (e) {
      if (e instanceof msal.InteractionRequiredAuthError || (e && /interaction_required|login_required|consent_required|no_tokens_found/.test(e.errorCode || ''))) {
        throw new OdError('login', 'La sesión de OneDrive ha caducado. Vuelve a conectar.');
      }
      throw new OdError('network', 'No se pudo conectar con Microsoft. Revisa la conexión.');
    }
  }

  async function graphFetch(url, opts) {
    try { return await fetch(url, opts); }
    catch (e) { throw new OdError('network', 'No se pudo conectar con OneDrive. Revisa la conexión.'); }
  }

  // Lee datos.json: null si aún no existe
  async function odPull(token) {
    const r = await graphFetch(GRAPH_FILE, { headers: { Authorization: 'Bearer ' + token } });
    if (r.status === 404) return null;
    if (r.status === 401) throw new OdError('login', 'La sesión de OneDrive ha caducado. Vuelve a conectar.');
    if (!r.ok) throw new OdError('http', 'OneDrive respondió con un error (' + r.status + ').');
    const meta = await r.json();
    const url = meta['@microsoft.graph.downloadUrl'];
    if (!url) throw new OdError('http', 'OneDrive no devolvió el fichero de datos.');
    const c = await graphFetch(url);
    if (!c.ok) throw new OdError('http', 'No se pudo descargar el fichero de datos (' + c.status + ').');
    let json;
    try { json = await c.json(); } catch (e) { throw new OdError('format', 'El fichero de datos de OneDrive está dañado.'); }
    return { etag: meta.eTag, doc: normalizeDoc(json) };
  }

  // Escribe datos.json solo si nadie lo ha cambiado desde etag (o si no existe, cuando etag es null)
  async function odPush(token, doc, etag) {
    const headers = { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' };
    if (etag) headers['If-Match'] = etag; else headers['If-None-Match'] = '*';
    const body = JSON.stringify({ app: 'mi-panel', schema: doc.schema, guardado: new Date().toISOString(), colecciones: doc.colecciones });
    const r = await graphFetch(GRAPH_FILE + ':/content', { method: 'PUT', headers, body });
    if (r.status === 412 || r.status === 409) return { conflict: true };
    if (r.status === 401) throw new OdError('login', 'La sesión de OneDrive ha caducado. Vuelve a conectar.');
    if (!r.ok) throw new OdError('http', 'OneDrive no aceptó los cambios (' + r.status + ').');
    const meta = await r.json();
    return { etag: meta.eTag };
  }

  async function syncDatos(opts = {}) {
    if (!odConnected()) return;
    if (S.odState === 'syncing') { odAgain = true; return; }
    clearTimeout(odTimer);
    S.odState = 'syncing';
    renderSyncState();
    try {
      const token = await odToken();
      let done = false;
      for (let i = 0; i < 4 && !done; i++) {
        const remote = await odPull(token);
        if (remote && remote.doc.schema > SCHEMA) {
          throw new OdError('schema', 'Los datos de OneDrive son de una versión más nueva de la app. Recarga la página.');
        }
        const base = normalizeDoc(S.odBase);
        let merged = S.datos;
        if (remote && remote.etag !== S.od.etag) merged = mergeDocs(base, S.datos, remote.doc);
        if (stable(merged.colecciones) !== stable(S.datos.colecciones)) {
          applyDatos(merged);
          renderAll();
          if (S.editing) renderDrawer();
        }
        if (remote && stable(merged.colecciones) === stable(remote.doc.colecciones)) {
          S.odBase = remote.doc;
          S.od.etag = remote.etag;
          done = true;
        } else {
          const res = await odPush(token, merged, remote ? remote.etag : null);
          if (res.conflict) continue; // otro dispositivo escribió entretanto: leer y fusionar otra vez
          S.odBase = normalizeDoc(merged);
          S.od.etag = res.etag;
          done = true;
        }
      }
      if (!done) throw new OdError('conflict', 'OneDrive cambió varias veces seguidas. Vuelve a sincronizar.');
      store.set(KEYS.base, S.odBase);
      S.od.syncedAt = Date.now();
      saveOd();
      S.odState = 'idle';
      if (opts.notify) toast('Datos sincronizados con OneDrive');
    } catch (e) {
      S.odState = e.kind === 'login' ? 'login' : 'error';
      if (e.kind === 'login') toast(e.message, 'error', { label: 'Conectar', fn: odConnect }, 10000);
      else toast(e.message || 'Error al sincronizar con OneDrive.', 'error');
    } finally {
      renderSyncState();
      if (odAgain) { odAgain = false; odSchedulePush(); }
    }
  }

  // Tras un cambio en los datos de la app se sube a OneDrive al poco (agrupando cambios seguidos)
  function odSchedulePush() {
    if (!odConnected() || S.odState === 'login') return;
    if (S.odState === 'syncing') { odAgain = true; return; }
    clearTimeout(odTimer);
    odTimer = setTimeout(() => syncDatos(), 1500);
  }

  const odPending = () => !S.odBase || stable(S.datos.colecciones) !== stable(normalizeDoc(S.odBase).colecciones);

  // ---------- datos derivados ----------
  function currentProjectId() {
    if (!S.data) return null;
    const ids = S.data.projects.map(p => p.id);
    if (S.ui.projectId && ids.includes(S.ui.projectId)) return S.ui.projectId;
    return S.data.inboxId || ids[0] || null;
  }

  function columns() {
    const pid = currentProjectId();
    if (!pid) return [];
    const byOrder = (a, b) => a.child_order - b.child_order;
    const secs = S.data.sections.filter(s => s.project_id === pid).sort((a, b) => a.order - b.order);
    const items = S.data.items.filter(i => i.project_id === pid && !i.parent_id);
    const cols = [];
    const loose = items.filter(i => !i.section_id).sort(byOrder);
    if (loose.length || !secs.length) cols.push({ key: 'none', name: 'Sin sección', items: loose });
    for (const s of secs) {
      cols.push({ key: s.id, name: s.name, items: items.filter(i => i.section_id === s.id).sort(byOrder) });
    }
    return cols;
  }

  const tagById = id => S.tags.list.find(t => t.id === id);
  const assigned = taskId => (S.tags.assign[taskId] || []).filter(id => tagById(id));
  const tagUsage = tagId => Object.values(S.tags.assign).filter(arr => arr.includes(tagId)).length;

  function isVisible(item) {
    if (!S.filter.size) return true;
    return assigned(item.id).some(id => S.filter.has(id));
  }

  // ---------- render ----------
  function renderAll() {
    applyLayout();
    renderProjects();
    renderSyncState();
    renderTagbar();
    renderBoard();
  }

  function applyLayout() {
    const app = document.getElementById('app');
    app.classList.toggle('sb-hidden', !SMALL.matches && S.ui.sidebarHidden);
    if (!SMALL.matches) {
      app.classList.remove('sb-open');
      document.getElementById('sidebarScrim').hidden = true;
    }
  }

  function renderProjects() {
    const sel = document.getElementById('projectSelect');
    sel.replaceChildren();
    if (!S.data) return;
    const pid = currentProjectId();
    for (const p of S.data.projects) {
      sel.append(h('option', { value: p.id, selected: p.id === pid }, p.inbox ? 'Bandeja de entrada' : p.name));
    }
  }

  function renderSyncState() {
    const when = S.data ? relTime(S.data.syncedAt) : 'nunca';
    const busy = S.syncing || S.odState === 'syncing';
    document.getElementById('syncInfo').textContent = busy ? 'Sincronizando…' : 'Sincronizado ' + when;
    document.getElementById('syncSide').textContent = when;
    const od = document.getElementById('syncOd');
    od.hidden = !odConnected();
    if (odConnected()) {
      const st = S.odState === 'syncing' ? 'sincronizando…'
        : S.odState === 'login' ? 'reconectar'
        : S.odState === 'error' ? 'con error'
        : odPending() ? 'cambios sin subir'
        : relTime(S.od.syncedAt);
      document.getElementById('syncOdState').textContent = st;
      od.classList.toggle('warn', S.odState === 'login' || S.odState === 'error');
    }
    const btn = document.getElementById('syncBtn');
    btn.disabled = busy;
    btn.classList.toggle('is-syncing', busy);
  }

  function renderTagbar() {
    const bar = document.getElementById('tagbar');
    bar.replaceChildren();
    bar.classList.toggle('filtering', S.filter.size > 0);
    bar.append(h('span', { class: 'tagbar-label' }, ICON.tag(),
      h('span', null, S.tags.list.length ? 'Arrastra a una tarea · pulsa para filtrar' : 'Aún no tienes etiquetas')));

    for (const t of S.tags.list) {
      const pill = h('button', {
        type: 'button',
        class: 'tag-pill' + (S.filter.has(t.id) ? ' active' : ''),
        style: '--c:' + t.color,
        'aria-pressed': S.filter.has(t.id) ? 'true' : 'false'
      }, ICON.grip(), t.name);
      attachTagDrag(pill, t);
      bar.append(pill);
    }

    const end = h('div', { class: 'tagbar-end' });
    if (S.filter.size) {
      end.append(h('button', { type: 'button', class: 'btn subtle small', onclick: () => { S.filter.clear(); renderTagbar(); renderBoard(); } }, 'Quitar filtro'));
    }
    end.append(h('button', { type: 'button', class: 'btn subtle small', onclick: openTagManager },
      S.tags.list.length ? 'Gestionar' : 'Crear etiquetas'));
    bar.append(end);
  }

  function renderBoard() {
    const board = document.getElementById('board');
    S.sortables.forEach(s => s.destroy());
    S.sortables = [];
    board.replaceChildren();
    if (!S.data) return;

    for (const col of columns()) {
      const visible = col.items.filter(isVisible);
      const list = h('div', { class: 'col-list', dataset: { sectionKey: col.key } }, visible.map(renderCard));
      const column = h('section', { class: 'column', 'aria-label': col.name },
        h('div', { class: 'col-head' },
          h('h2', { title: col.name }, col.name),
          h('span', { class: 'count' }, String(S.filter.size ? visible.length : col.items.length))),
        list,
        !visible.length ? h('p', { class: 'col-empty' }, S.filter.size ? 'Nada con este filtro' : 'Sin tareas') : null,
        h('button', { type: 'button', class: 'col-add', onclick: () => openNew(col.key, false) }, ICON.plus(), 'Añadir tarea')
      );
      board.append(column);

      S.sortables.push(Sortable.create(list, {
        group: 'tasks',
        animation: 160,
        forceFallback: true,
        fallbackClass: 'card-drag',
        fallbackTolerance: 4,
        delayOnTouchOnly: true,
        delay: 180,
        filter: '.check, .pending',
        preventOnFilter: false,
        ghostClass: 'card-ghost',
        chosenClass: 'card-chosen',
        onEnd: onCardDrop
      }));
    }
  }

  function renderCard(item) {
    const tags = assigned(item.id).map(tagById);
    const pending = item.id.startsWith('tmp-');
    const hasNotes = (S.data.notes || []).some(n => n.item_id === item.id);
    return h('div', {
      class: 'card' + (pending ? ' pending' : '') + (hasNotes ? ' has-notes' : ''),
      dataset: { taskId: item.id },
      style: '--strip:' + (tags[0] ? tags[0].color : 'transparent')
    },
      h('button', {
        type: 'button', class: 'check', 'aria-label': 'Completar tarea', disabled: pending,
        onclick: e => { e.stopPropagation(); completeTask(item.id); }
      }),
      h('button', {
        type: 'button', class: 'card-title', title: cleanTitle(item.content),
        onclick: () => { if (!pending) openEdit(item.id); }
      }, cleanTitle(item.content)),
      hasNotes ? h('span', { class: 'card-notes', title: 'Tiene comentarios', 'aria-label': 'Tiene comentarios' }, ICON.comment()) : null,
      tags.length ? h('span', { class: 'card-tags' },
        tags.map(t => h('span', { class: 'mini-tag', style: '--c:' + t.color }, t.name))) : null
    );
  }

  // ---------- acciones sobre tareas ----------
  async function onCardDrop(evt) {
    const id = evt.item.dataset.taskId;
    const fromKey = evt.from.dataset.sectionKey;
    const toKey = evt.to.dataset.sectionKey;
    if (fromKey === toKey && evt.oldIndex === evt.newIndex) return;
    const item = S.data.items.find(i => i.id === id);
    if (!item) return;

    const commands = [];
    if (fromKey !== toKey) {
      const args = { id };
      if (toKey === 'none') args.project_id = currentProjectId();
      else args.section_id = toKey;
      commands.push({ type: 'item_move', uuid: uid(), args });
      item.section_id = toKey === 'none' ? null : toKey;
    }
    // El orden solo se envía sin filtro activo (con filtro no se ven todas las tareas de la columna)
    if (!S.filter.size) {
      const order = [...evt.to.querySelectorAll('.card')]
        .map((el, idx) => ({ id: el.dataset.taskId, child_order: idx + 1 }))
        .filter(o => !o.id.startsWith('tmp-'));
      if (order.length) commands.push({ type: 'item_reorder', uuid: uid(), args: { items: order } });
      for (const o of order) {
        const it = S.data.items.find(x => x.id === o.id);
        if (it) it.child_order = o.child_order;
      }
    }
    saveCache();
    setTimeout(renderBoard, 0);
    if (!commands.length) return;
    try { await runCommands(commands); }
    catch (e) { await recover(e); }
  }

  async function completeTask(id) {
    const idx = S.data.items.findIndex(i => i.id === id);
    if (idx < 0) return;
    const [item] = S.data.items.splice(idx, 1);
    saveCache();
    renderBoard();
    renderProjects();
    const undo = item.recurring ? null : { label: 'Deshacer', fn: () => undoComplete(item) };
    toast(item.recurring ? 'Tarea recurrente completada' : 'Tarea completada', 'ok', undo);
    try {
      await runCommands([{ type: 'item_close', uuid: uid(), args: { id } }]);
      if (item.recurring) await doSync({ silent: true });
    } catch (e) { await recover(e); }
  }

  async function undoComplete(item) {
    S.data.items.push(item);
    saveCache();
    renderBoard();
    renderProjects();
    try { await runCommands([{ type: 'item_uncomplete', uuid: uid(), args: { id: item.id } }]); }
    catch (e) { await recover(e); }
  }

  function assignTag(taskId, tagId) {
    if (!taskId || taskId.startsWith('tmp-')) return;
    const list = S.tags.assign[taskId] || [];
    if (list.includes(tagId)) return;
    S.tags.assign[taskId] = list.concat(tagId);
    saveTags();
    renderBoard();
  }

  // ---------- arrastrar etiquetas a tareas ----------
  function attachTagDrag(pill, tag) {
    pill.addEventListener('pointerdown', e => {
      if (e.button !== 0) return;
      const start = { x: e.clientX, y: e.clientY };
      let ghost = null;
      let target = null;
      let moved = false;
      try { pill.setPointerCapture(e.pointerId); } catch (_) { /* sin captura */ }

      const onMove = ev => {
        if (!moved) {
          if (Math.hypot(ev.clientX - start.x, ev.clientY - start.y) < 6) return;
          moved = true;
          ghost = pill.cloneNode(true);
          ghost.classList.add('tag-ghost');
          ghost.classList.remove('active');
          document.body.append(ghost);
          pill.classList.add('dragging');
          document.body.classList.add('tag-dragging');
        }
        ghost.style.transform = 'translate(' + (ev.clientX + 10) + 'px,' + (ev.clientY + 10) + 'px)';
        const under = document.elementFromPoint(ev.clientX, ev.clientY);
        const card = under ? under.closest('.card:not(.pending)') : null;
        if (card !== target) {
          if (target) target.classList.remove('drop-target');
          target = card;
          if (target) target.classList.add('drop-target');
        }
      };
      const finish = ev => {
        pill.removeEventListener('pointermove', onMove);
        pill.removeEventListener('pointerup', finish);
        pill.removeEventListener('pointercancel', finish);
        if (ghost) ghost.remove();
        pill.classList.remove('dragging');
        document.body.classList.remove('tag-dragging');
        if (target) target.classList.remove('drop-target');
        if (ev.type === 'pointercancel') return;
        if (!moved) {
          if (S.filter.has(tag.id)) S.filter.delete(tag.id); else S.filter.add(tag.id);
          renderTagbar();
          renderBoard();
          return;
        }
        if (target) assignTag(target.dataset.taskId, tag.id);
      };
      pill.addEventListener('pointermove', onMove);
      pill.addEventListener('pointerup', finish);
      pill.addEventListener('pointercancel', finish);
    });
    // El teclado (Enter/Espacio) filtra
    pill.addEventListener('keydown', e => {
      if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        if (S.filter.has(tag.id)) S.filter.delete(tag.id); else S.filter.add(tag.id);
        renderTagbar();
        renderBoard();
      }
    });
  }

  // ---------- panel de tarea (nueva / editar) ----------
  function openNew(sectionKey, fromHeader) {
    if (!S.data) return;
    S.editing = { mode: 'new', sectionKey, fromHeader, name: '', tagIds: [] };
    renderDrawer();
  }

  function openEdit(taskId) {
    const item = S.data.items.find(i => i.id === taskId);
    if (!item) return;
    S.editing = { mode: 'edit', taskId, name: item.content, original: item.content, tagIds: assigned(taskId).slice() };
    renderDrawer();
  }

  function closeDrawer() {
    S.editing = null;
    document.getElementById('drawer').hidden = true;
    document.getElementById('drawerScrim').hidden = true;
  }

  function columnName(key) {
    const col = columns().find(c => c.key === key);
    return col ? col.name : 'Sin sección';
  }

  function renderDrawer() {
    const ed = S.editing;
    const drawer = document.getElementById('drawer');
    drawer.replaceChildren();
    if (!ed) return;

    const item = ed.mode === 'edit' ? S.data.items.find(i => i.id === ed.taskId) : null;
    const where = ed.mode === 'edit'
      ? 'en ' + columnName(item && item.section_id ? item.section_id : 'none')
      : 'en ' + columnName(ed.sectionKey);

    const nameInput = h('input', {
      class: 'input', id: 'taskName', type: 'text', value: ed.name, maxlength: '500',
      placeholder: 'Nombre de la tarea', autocomplete: 'off'
    });
    nameInput.addEventListener('input', () => { ed.name = nameInput.value; });
    nameInput.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); saveDrawer(); } });

    const chips = h('div', { class: 'chips' });
    const drawChips = () => {
      chips.replaceChildren();
      for (const t of S.tags.list) {
        const on = ed.tagIds.includes(t.id);
        chips.append(h('button', {
          type: 'button', class: 'chip' + (on ? ' on' : ''), style: '--c:' + t.color, 'aria-pressed': on ? 'true' : 'false',
          onclick: () => {
            ed.tagIds = on ? ed.tagIds.filter(x => x !== t.id) : ed.tagIds.concat(t.id);
            drawChips();
          }
        }, h('span', { class: 'dot' }), t.name));
      }
    };
    drawChips();

    const commentsField = ed.mode === 'edit' && !ed.taskId.startsWith('tmp-') ? commentsSection(ed.taskId) : null;

    let columnField = null;
    if (ed.mode === 'new' && ed.fromHeader) {
      const sel = h('select', { class: 'select', id: 'taskColumn' },
        columns().map(c => h('option', { value: c.key, selected: c.key === ed.sectionKey }, c.name)));
      sel.addEventListener('change', () => { ed.sectionKey = sel.value; });
      columnField = h('div', { class: 'field' }, h('label', { class: 'field-label', for: 'taskColumn' }, 'Columna'), sel);
    }

    // drawer.append no ignora null (lo pinta como texto), así que se filtran los huecos
    drawer.append(...[
      h('div', { class: 'drawer-head' },
        h('div', null,
          h('h2', null, ed.mode === 'new' ? 'Nueva tarea' : 'Editar tarea'),
          ed.fromHeader ? null : h('p', null, where)),
        h('button', { type: 'button', class: 'icon-btn', 'aria-label': 'Cerrar', onclick: closeDrawer }, ICON.close())),
      h('div', { class: 'field' }, h('label', { class: 'field-label', for: 'taskName' }, 'Nombre'), nameInput),
      columnField,
      h('div', { class: 'field' },
        h('span', { class: 'field-label' }, 'Etiquetas'),
        S.tags.list.length ? chips : h('p', { class: 'hint' }, 'Todavía no has creado etiquetas.'),
        S.tags.list.length
          ? h('p', { class: 'hint' }, 'Pulsa una etiqueta para asignarla o quitarla.')
          : h('div', null, h('button', { type: 'button', class: 'btn ghost small', onclick: openTagManager }, 'Crear etiquetas'))),
      commentsField,
      h('div', { class: 'drawer-foot' },
        ed.mode === 'edit'
          ? h('button', { type: 'button', class: 'btn ghost small', onclick: () => { const id = ed.taskId; closeDrawer(); completeTask(id); } }, 'Completar')
          : null,
        h('span', { class: 'spacer' }),
        h('button', { type: 'button', class: 'btn subtle small', onclick: closeDrawer }, 'Cancelar'),
        h('button', { type: 'button', class: 'btn primary small', onclick: saveDrawer }, ed.mode === 'new' ? 'Crear tarea' : 'Guardar'))
    ].filter(Boolean));

    drawer.hidden = false;
    document.getElementById('drawerScrim').hidden = false;
    setTimeout(() => nameInput.focus(), 50);
  }

  // ---------- comentarios (se guardan en Todoist al momento, aparte del botón Guardar) ----------
  function notesOf(taskId) {
    return (S.data.notes || []).filter(n => n.item_id === taskId)
      .sort((a, b) => String(a.posted_at || '').localeCompare(String(b.posted_at || '')));
  }

  function noteDate(iso) {
    const d = new Date(iso);
    if (!iso || isNaN(d)) return '';
    const opts = { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' };
    if (d.getFullYear() !== new Date().getFullYear()) opts.year = 'numeric';
    return d.toLocaleString('es-ES', opts);
  }

  function commentsSection(taskId) {
    const box = h('div', { class: 'field' });
    let editingId = null;
    // Pinta el cambio optimista al momento y otra vez cuando responde Todoist
    const run = p => { draw(); renderBoard(); p.then(() => { draw(); renderBoard(); }); };

    // Atajos en los cuadros de texto: Ctrl/Cmd+Intro guarda, Escape cancela sin cerrar el panel
    const keys = (area, onSave, onCancel) => area.addEventListener('keydown', e => {
      if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); onSave(); }
      else if (e.key === 'Escape' && onCancel) { e.stopPropagation(); onCancel(); }
    });

    const draw = () => {
      if (!box.isConnected && box.childElementCount) return;
      const notes = notesOf(taskId);
      const list = h('div', { class: 'notes' });
      let focusArea = null;
      for (const n of notes) {
        const pending = n.id.startsWith('tmp-');
        if (n.id === editingId) {
          const area = h('textarea', { class: 'input textarea', rows: '3', 'aria-label': 'Editar comentario' });
          area.value = n.content;
          const save = () => run(updateNote(n, area.value));
          const cancel = () => { editingId = null; draw(); };
          keys(area, () => { editingId = null; save(); }, cancel);
          list.append(h('div', { class: 'note' }, area,
            h('div', { class: 'note-actions' },
              h('button', { type: 'button', class: 'btn subtle small', onclick: cancel }, 'Cancelar'),
              h('button', { type: 'button', class: 'btn primary small', onclick: () => { editingId = null; save(); } }, 'Guardar'))));
          focusArea = area;
          continue;
        }
        list.append(h('div', { class: 'note' + (pending ? ' pending' : '') },
          h('div', { class: 'note-meta' },
            h('span', null, pending ? 'Enviando…' : noteDate(n.posted_at)),
            pending ? null : h('span', { class: 'note-tools' },
              h('button', { type: 'button', class: 'icon-btn mini', 'aria-label': 'Editar comentario', onclick: () => { editingId = n.id; draw(); } }, ICON.edit()),
              h('button', { type: 'button', class: 'icon-btn mini', 'aria-label': 'Borrar comentario', onclick: () => run(deleteNote(n)) }, ICON.trash()))),
          n.content ? h('div', { class: 'note-text' }, n.content) : null,
          n.file ? h('a', { class: 'note-file', href: n.file.url, target: '_blank', rel: 'noopener noreferrer' }, n.file.name) : null));
      }

      const area = h('textarea', { class: 'input textarea', rows: '2', placeholder: 'Escribe un comentario', 'aria-label': 'Nuevo comentario' });
      const add = () => {
        const text = area.value.trim();
        if (!text) { area.focus(); return; }
        run(addNote(taskId, text));
      };
      keys(area, add, null);

      box.replaceChildren(...[
        h('span', { class: 'field-label' }, 'Comentarios'),
        notes.length ? list : null,
        area,
        h('div', { class: 'note-actions' },
          h('span', { class: 'hint' }, 'Se guardan en Todoist al momento.'),
          h('button', { type: 'button', class: 'btn ghost small', onclick: add }, 'Comentar'))
      ].filter(Boolean));
      if (focusArea) {
        focusArea.focus();
        focusArea.setSelectionRange(focusArea.value.length, focusArea.value.length);
      }
    };
    draw();
    return box;
  }

  async function addNote(taskId, content) {
    const tempId = 'tmp-' + uid();
    const note = { id: tempId, item_id: taskId, content, posted_at: new Date().toISOString(), file: null };
    S.data.notes = (S.data.notes || []).concat(note);
    try {
      const map = await runCommands([{ type: 'note_add', temp_id: tempId, uuid: uid(), args: { item_id: taskId, content } }]);
      const realId = map[tempId];
      if (!realId) throw new ApiError('command', 'Todoist no devolvió el identificador del comentario.');
      note.id = realId;
      saveCache();
    } catch (e) {
      S.data.notes = S.data.notes.filter(n => n.id !== tempId);
      await recover(e);
    }
  }

  async function updateNote(note, value) {
    const content = value.trim();
    if (!content || content === note.content) return;
    note.content = content;
    saveCache();
    try { await runCommands([{ type: 'note_update', uuid: uid(), args: { id: note.id, content } }]); }
    catch (e) { await recover(e); }
  }

  async function deleteNote(note) {
    if (!confirm('¿Borrar este comentario? También se borrará en Todoist.')) return;
    S.data.notes = S.data.notes.filter(n => n.id !== note.id);
    saveCache();
    try { await runCommands([{ type: 'note_delete', uuid: uid(), args: { id: note.id } }]); }
    catch (e) { await recover(e); }
  }

  async function saveDrawer() {
    const ed = S.editing;
    if (!ed) return;
    const name = ed.name.trim();
    if (!name) {
      toast('Escribe un nombre para la tarea.', 'error');
      return;
    }
    const tagIds = ed.tagIds.filter(id => tagById(id));
    closeDrawer();

    if (ed.mode === 'edit') {
      const item = S.data.items.find(i => i.id === ed.taskId);
      if (!item) return;
      S.tags.assign[ed.taskId] = tagIds;
      saveTags();
      const renamed = name !== ed.original.trim();
      if (renamed) { item.content = name; saveCache(); }
      renderBoard();
      if (renamed) {
        try { await runCommands([{ type: 'item_update', uuid: uid(), args: { id: item.id, content: name } }]); }
        catch (e) { await recover(e); }
      }
      return;
    }

    // Nueva tarea
    const pid = currentProjectId();
    const tempId = 'tmp-' + uid();
    const sectionId = ed.sectionKey === 'none' ? null : ed.sectionKey;
    const siblings = S.data.items.filter(i => i.project_id === pid && (i.section_id || null) === sectionId && !i.parent_id);
    const item = {
      id: tempId, content: name, project_id: pid, section_id: sectionId, parent_id: null,
      child_order: siblings.reduce((m, i) => Math.max(m, i.child_order), 0) + 1, recurring: false
    };
    S.data.items.push(item);
    S.tags.assign[tempId] = tagIds;
    renderBoard();
    renderProjects();

    const args = { content: name, project_id: pid };
    if (sectionId) args.section_id = sectionId;
    try {
      const map = await runCommands([{ type: 'item_add', temp_id: tempId, uuid: uid(), args }]);
      const realId = map[tempId];
      if (!realId) throw new ApiError('command', 'Todoist no devolvió el identificador de la tarea.');
      item.id = realId;
      S.tags.assign[realId] = S.tags.assign[tempId] || [];
      delete S.tags.assign[tempId];
      saveTags();
      saveCache();
      renderBoard();
    } catch (e) {
      S.data.items = S.data.items.filter(i => i.id !== tempId);
      delete S.tags.assign[tempId];
      saveTags();
      await recover(e);
    }
  }

  // ---------- modales ----------
  function openModal(title, body, opts = {}) {
    const root = document.getElementById('modalRoot');
    const close = () => {
      wrap.remove();
      document.removeEventListener('keydown', onKey);
      if (opts.onClose) opts.onClose();
    };
    const onKey = e => { if (e.key === 'Escape' && !opts.locked) close(); };
    const modal = h('div', { class: 'modal', role: 'dialog', 'aria-modal': 'true', 'aria-label': title },
      h('div', { class: 'modal-head' },
        h('h2', null, title),
        opts.locked ? null : h('button', { type: 'button', class: 'icon-btn', 'aria-label': 'Cerrar', onclick: close }, ICON.close())),
      body);
    const wrap = h('div', { class: 'modal-wrap' }, modal);
    wrap.addEventListener('mousedown', e => { if (e.target === wrap && !opts.locked) close(); });
    document.addEventListener('keydown', onKey);
    root.append(wrap);
    return close;
  }

  // Conectar Todoist
  function openSetup(errorMsg) {
    document.getElementById('modalRoot').replaceChildren();
    const input = h('input', { class: 'input', id: 'tokenInput', type: 'password', placeholder: 'Pega aquí tu token de Todoist', autocomplete: 'off' });
    const error = h('p', { class: 'error', role: 'alert' }, errorMsg || '');
    error.hidden = !errorMsg;
    const btn = h('button', { type: 'button', class: 'btn primary' }, 'Conectar');
    const hasToken = !!S.token;

    const body = h('div', { style: 'display:flex;flex-direction:column;gap:16px' },
      h('p', null, 'Para leer y actualizar tus tareas, la app necesita tu token personal de Todoist.'),
      h('ol', { class: 'steps' },
        h('li', null, 'Abre Todoist en el navegador.'),
        h('li', null, 'Ve a Ajustes → Integraciones → Desarrollador.'),
        h('li', null, 'Copia el token de API y pégalo aquí.')),
      h('div', { class: 'field' }, h('label', { class: 'field-label', for: 'tokenInput' }, 'Token de API'), input),
      error,
      h('p', { class: 'muted' }, 'El token se guarda solo en este navegador. Tendrás que pegarlo una vez en cada dispositivo.'),
      h('div', { class: 'row-actions' }, btn));

    const close = openModal('Conecta tu Todoist', body, { locked: !hasToken });

    const connect = async () => {
      const token = input.value.trim();
      if (!token) { error.textContent = 'Pega el token primero.'; error.hidden = false; return; }
      btn.disabled = true;
      btn.textContent = 'Conectando…';
      try {
        const data = await fetchAll(token);
        S.token = token;
        store.set(KEYS.token, token);
        S.data = data;
        saveCache();
        close();
        renderAll();
        toast('Todoist conectado');
      } catch (e) {
        error.textContent = e.message;
        error.hidden = false;
        btn.disabled = false;
        btn.textContent = 'Conectar';
      }
    };
    btn.addEventListener('click', connect);
    input.addEventListener('keydown', e => { if (e.key === 'Enter') connect(); });
    setTimeout(() => input.focus(), 50);
  }

  // Gestor de etiquetas
  function openTagManager() {
    if (S.editing) closeDrawer();
    const body = h('div', { style: 'display:flex;flex-direction:column;gap:14px' });
    let openPaletteId = null;
    let confirmId = null;
    let confirmTimer = null;
    let listSortable = null;

    const changed = () => { saveTags(); renderTagbar(); renderBoard(); };

    const palette = (current, onPick) => {
      const box = h('div', { class: 'palette' });
      for (const c of PALETTE) {
        box.append(h('button', {
          type: 'button', class: c.toLowerCase() === current.toLowerCase() ? 'sel' : '', style: '--c:' + c,
          'aria-label': 'Color ' + c, onclick: () => onPick(c)
        }));
      }
      const custom = h('input', { type: 'color', value: current, 'aria-label': 'Color personalizado' });
      custom.addEventListener('change', () => { if (isHex(custom.value)) onPick(custom.value); });
      box.append(h('label', null, 'Otro', custom));
      return box;
    };

    const draw = () => {
      body.replaceChildren();
      const list = h('div', { class: 'tm-list' });

      for (const t of S.tags.list) {
        const name = h('input', { class: 'input tm-name', type: 'text', value: t.name, maxlength: '40', 'aria-label': 'Nombre de la etiqueta' });
        name.addEventListener('change', () => {
          const v = name.value.trim();
          if (!v) { name.value = t.name; return; }
          t.name = v;
          changed();
        });
        const uses = tagUsage(t.id);
        const isConfirm = confirmId === t.id;
        const del = h('button', {
          type: 'button', class: 'icon-btn tm-del' + (isConfirm ? ' confirm' : ''),
          'aria-label': 'Borrar etiqueta',
          onclick: () => {
            if (uses > 0 && !isConfirm) {
              confirmId = t.id;
              clearTimeout(confirmTimer);
              confirmTimer = setTimeout(() => { confirmId = null; draw(); }, 4000);
              draw();
              return;
            }
            S.tags.list = S.tags.list.filter(x => x.id !== t.id);
            for (const k of Object.keys(S.tags.assign)) S.tags.assign[k] = S.tags.assign[k].filter(id => id !== t.id);
            S.filter.delete(t.id);
            confirmId = null;
            changed();
            draw();
          }
        }, isConfirm ? 'Borrar (en ' + uses + (uses === 1 ? ' tarea)' : ' tareas)') : ICON.trash());

        const row = h('div', { class: 'tm-row', dataset: { id: t.id } },
          h('div', { class: 'tm-main' },
            h('span', { class: 'tm-handle', 'aria-hidden': 'true' }, ICON.grip()),
            h('button', {
              type: 'button', class: 'swatch', style: '--c:' + t.color, 'aria-label': 'Cambiar color',
              onclick: () => { openPaletteId = openPaletteId === t.id ? null : t.id; draw(); }
            }),
            name, del),
          openPaletteId === t.id ? palette(t.color, c => { t.color = c; openPaletteId = null; changed(); draw(); }) : null);
        list.append(row);
      }
      if (!S.tags.list.length) list.append(h('p', { class: 'muted' }, 'Crea tu primera etiqueta: por ejemplo "Llamar", "Comprar" o "Casa".'));

      // Nueva etiqueta
      const used = S.tags.list.map(t => t.color.toLowerCase());
      let newColor = PALETTE.find(c => !used.includes(c.toLowerCase())) || PALETTE[0];
      let showNewPalette = false;
      const newName = h('input', { class: 'input tm-name', type: 'text', maxlength: '40', placeholder: 'Nueva etiqueta', 'aria-label': 'Nombre de la nueva etiqueta' });
      const newSwatch = h('button', { type: 'button', class: 'swatch', style: '--c:' + newColor, 'aria-label': 'Color de la nueva etiqueta' });
      const newPaletteSlot = h('div');
      newSwatch.addEventListener('click', () => {
        showNewPalette = !showNewPalette;
        newPaletteSlot.replaceChildren(showNewPalette ? palette(newColor, c => {
          newColor = c;
          newSwatch.style.setProperty('--c', c);
          showNewPalette = false;
          newPaletteSlot.replaceChildren();
        }) : '');
      });
      const add = () => {
        const v = newName.value.trim();
        if (!v) { newName.focus(); return; }
        S.tags.list.push({ id: 't-' + uid(), name: v, color: newColor });
        changed();
        draw();
        const again = body.querySelector('input[placeholder="Nueva etiqueta"]');
        if (again) again.focus();
      };
      newName.addEventListener('keydown', e => { if (e.key === 'Enter') add(); });

      body.append(
        list,
        h('div', { class: 'tm-row' },
          h('div', { class: 'tm-new' }, newSwatch, newName,
            h('button', { type: 'button', class: 'btn primary small', onclick: add }, 'Añadir')),
          newPaletteSlot),
        h('p', { class: 'muted' }, 'Arrastra por el asa para cambiar el orden. Las etiquetas se guardan en este navegador y, si lo conectas, en OneDrive.'));

      if (listSortable) listSortable.destroy();
      listSortable = Sortable.create(list, {
        handle: '.tm-handle',
        animation: 150,
        forceFallback: true,
        onEnd: () => {
          const ids = [...list.querySelectorAll('.tm-row')].map(r => r.dataset.id);
          S.tags.list.sort((a, b) => ids.indexOf(a.id) - ids.indexOf(b.id));
          changed();
        }
      });
    };

    draw();
    openModal('Etiquetas', body, {
      onClose: () => {
        if (listSortable) listSortable.destroy();
        clearTimeout(confirmTimer);
        renderAll();
      }
    });
  }

  // Ajustes · OneDrive
  function odSection() {
    const box = h('div', { class: 'modal-section' });
    const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

    const draw = () => {
      const intro = h('p', { class: 'muted' }, 'Guarda los datos de la app en tu OneDrive personal (en una carpeta propia de la app) y los sincroniza entre tus dispositivos.');
      if (!S.od.clientId) {
        const input = h('input', { class: 'input', type: 'text', placeholder: 'Id. de aplicación (cliente)', autocomplete: 'off', spellcheck: 'false', 'aria-label': 'Id. de aplicación (cliente)' });
        const saveId = () => {
          const v = input.value.trim();
          if (!GUID.test(v)) { toast('El Id. de aplicación tiene este formato: 1234abcd-12ab-34cd-56ef-1234567890ab', 'error'); input.focus(); return; }
          S.od.clientId = v;
          saveOd();
          loadMsal().then(draw).catch(e => toast(e.message, 'error'));
          draw();
        };
        input.addEventListener('keydown', e => { if (e.key === 'Enter') saveId(); });
        box.replaceChildren(h('h3', null, 'OneDrive'), intro,
          h('p', { class: 'muted' }, 'Pega el Id. de aplicación (cliente) del registro que creaste en Microsoft Entra.'),
          input,
          h('div', { class: 'row-actions' }, h('button', { type: 'button', class: 'btn ghost small', onclick: saveId }, 'Guardar')));
        return;
      }

      const changeId = h('button', {
        type: 'button', class: 'btn subtle small',
        onclick: async () => {
          if (!confirm('¿Cambiar el Id. de aplicación? Se desconectará OneDrive en este navegador.')) return;
          await odDisconnect();
          S.od.clientId = '';
          saveOd();
          msalApp = null; msalLoading = null;
          location.reload();
        }
      }, 'Cambiar Id. de aplicación');

      if (!odConnected()) {
        box.replaceChildren(h('h3', null, 'OneDrive'), intro,
          h('p', { class: 'muted' }, 'Al conectar, tus datos de este navegador se juntan con los que ya haya en OneDrive; no se pierde nada.'),
          h('div', { class: 'row-actions' },
            h('button', { type: 'button', class: 'btn primary small', onclick: () => { odConnect(); } }, 'Conectar con OneDrive'),
            changeId));
        return;
      }

      const estado = S.odState === 'syncing' ? 'Sincronizando…'
        : S.odState === 'login' ? 'La sesión ha caducado: pulsa Reconectar.'
        : S.odState === 'error' ? 'La última sincronización falló.'
        : odPending() ? 'Hay cambios sin subir.'
        : 'Última sincronización: ' + relTime(S.od.syncedAt) + '.';
      box.replaceChildren(h('h3', null, 'OneDrive'),
        h('p', { class: 'muted' }, 'Conectado como ', h('strong', null, S.od.account), '.'),
        h('p', { class: 'muted' }, estado),
        h('div', { class: 'row-actions' },
          S.odState === 'login'
            ? h('button', { type: 'button', class: 'btn primary small', onclick: () => { odConnect(); } }, 'Reconectar')
            : h('button', { type: 'button', class: 'btn ghost small', onclick: () => syncDatos({ notify: true }) }, 'Sincronizar ahora'),
          h('button', {
            type: 'button', class: 'btn subtle small',
            onclick: async () => {
              if (!confirm('¿Desconectar OneDrive en este navegador? Tus datos se quedan en OneDrive y en este navegador.')) return;
              await odDisconnect();
              draw();
            }
          }, 'Desconectar')));
    };
    draw();
    // Mientras Ajustes está abierto, refleja los cambios de estado (login, sincronización)
    const timer = setInterval(() => { if (!box.isConnected) clearInterval(timer); else if (!(box.contains(document.activeElement) && document.activeElement.tagName === 'INPUT')) draw(); }, 1000);
    return box;
  }

  // Ajustes
  function openSettings() {
    const fileInput = h('input', { type: 'file', accept: 'application/json,.json', hidden: true });
    fileInput.addEventListener('change', async () => {
      const file = fileInput.files && fileInput.files[0];
      if (!file) return;
      try {
        const parsed = JSON.parse(await file.text());
        if (parsed && parsed.datos) {
          // Copia completa del almacén
          const doc = normalizeDoc(parsed.datos);
          if (doc.schema > SCHEMA) throw new Error('versión');
          if (!confirm('Se sustituirán todos los datos de la app por los del fichero. ¿Continuar?')) return;
          applyDatos(doc);
          saveDatos();
        } else {
          // Formato antiguo: solo etiquetas
          const tags = normalizeTags(parsed && parsed.tags ? parsed.tags : parsed);
          if (!tags.list.length) throw new Error('vacío');
          if (!confirm('Se sustituirán tus etiquetas actuales por las del fichero (' + tags.list.length + ' etiquetas). ¿Continuar?')) return;
          S.tags = tags;
          saveTags();
        }
        renderAll();
        toast('Datos importados');
      } catch (e) {
        toast('El fichero no tiene un formato válido.', 'error');
      } finally {
        fileInput.value = '';
      }
    });

    const exportData = () => {
      const blob = new Blob([JSON.stringify({ app: 'mi-panel', version: 2, exportedAt: new Date().toISOString(), datos: S.datos }, null, 2)], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const a = h('a', { href: url, download: 'mi-panel-datos.json' });
      document.body.append(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    };

    const masked = S.token ? '••••••••' + S.token.slice(-4) : 'sin conectar';
    let close = null;
    const body = h('div', { style: 'display:flex;flex-direction:column;gap:16px' },
      h('div', { class: 'modal-section', style: 'border-top:none;padding-top:0' },
        h('h3', null, 'Todoist'),
        h('p', { class: 'muted' }, 'Token: ' + masked),
        h('div', { class: 'row-actions' },
          h('button', { type: 'button', class: 'btn ghost small', onclick: () => { close(); openSetup(); } }, 'Cambiar token'))),
      odSection(),
      h('div', { class: 'modal-section' },
        h('h3', null, 'Copia de seguridad'),
        h('p', { class: 'muted' }, 'Exporta todos los datos de la app (etiquetas y lo que se añada) en un fichero, o importa una copia. Importar sustituye los datos actuales.'),
        h('div', { class: 'row-actions' },
          h('button', { type: 'button', class: 'btn ghost small', onclick: exportData }, 'Exportar'),
          h('button', { type: 'button', class: 'btn ghost small', onclick: () => fileInput.click() }, 'Importar'),
          fileInput)),
      h('div', { class: 'modal-section' },
        h('h3', null, 'Este navegador'),
        h('p', { class: 'muted' }, 'Borra de este navegador el token, los datos de la app, la sesión de OneDrive y la copia local de tareas. En Todoist y OneDrive no se borra nada.'),
        h('div', { class: 'row-actions' },
          h('button', {
            type: 'button', class: 'btn danger small',
            onclick: () => {
              const aviso = odConnected() && !odPending()
                ? '¿Borrar todos los datos de la app en este navegador? Tus datos siguen en OneDrive.'
                : '¿Borrar todos los datos de la app en este navegador? Exporta antes una copia si quieres conservarlos.';
              if (!confirm(aviso)) return;
              Object.values(KEYS).concat(LEGACY_TAGS_KEY).forEach(k => localStorage.removeItem(k));
              Object.keys(localStorage).filter(k => k.startsWith('msal.')).forEach(k => localStorage.removeItem(k));
              location.reload();
            }
          }, 'Borrar datos de este navegador'))));
    close = openModal('Ajustes', body);
  }

  // ---------- menú lateral ----------
  function toggleSidebar(show) {
    const app = document.getElementById('app');
    if (SMALL.matches) {
      app.classList.toggle('sb-open', show);
      document.getElementById('sidebarScrim').hidden = !show;
    } else {
      S.ui.sidebarHidden = !show;
      saveUi();
      applyLayout();
    }
  }

  // ---------- arranque ----------
  function bind() {
    document.getElementById('hideSidebar').addEventListener('click', () => toggleSidebar(false));
    document.getElementById('showSidebar').addEventListener('click', () => toggleSidebar(true));
    document.getElementById('sidebarScrim').addEventListener('click', () => toggleSidebar(false));
    document.getElementById('drawerScrim').addEventListener('click', closeDrawer);
    document.getElementById('syncBtn').addEventListener('click', () => {
      // Si la sesión de OneDrive caducó, este clic sirve para abrir la ventana de login
      if (odConnected() && S.odState === 'login') odConnect();
      else syncDatos();
      doSync();
    });
    document.getElementById('addBtn').addEventListener('click', () => {
      const cols = columns();
      if (cols.length) openNew(cols[0].key, true);
    });
    document.getElementById('openTags').addEventListener('click', () => { if (SMALL.matches) toggleSidebar(false); openTagManager(); });
    document.getElementById('openSettings').addEventListener('click', () => { if (SMALL.matches) toggleSidebar(false); openSettings(); });
    document.getElementById('projectSelect').addEventListener('change', e => {
      S.ui.projectId = e.target.value;
      saveUi();
      S.filter.clear();
      renderAll();
    });
    document.addEventListener('keydown', e => {
      if (e.key === 'Escape' && S.editing && !document.querySelector('.modal-wrap')) closeDrawer();
    });
    SMALL.addEventListener('change', applyLayout);
    setInterval(renderSyncState, 60000);
  }

  function start() {
    bind();
    renderAll();
    // MSAL se prepara al arrancar (sin conectar con nada) para que el login responda al primer clic
    if (S.od.clientId) loadMsal().catch(() => {});
    if (!S.token) openSetup();
    else if (!S.data) doSync({ silent: true });
  }

  if (typeof Sortable === 'undefined') {
    document.body.textContent = 'Falta el fichero Sortable.min.js junto a index.html.';
    return;
  }
  start();
})();
