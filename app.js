/* Mi panel
   - Tareas y secciones: Todoist (API v1, endpoint /sync)
   - Datos propios de la app (etiquetas, ...): almacén local por colecciones,
     sincronizado con OneDrive personal (carpeta de la app) si el usuario lo conecta
*/
(() => {
  'use strict';

  // iPhone/iPad: Safari amplía la página al tocar un campo con letra de menos de 16 px y no la
  // devuelve a su tamaño. Con maximum-scale=1 no lo hace; pellizcar para ampliar sigue funcionando en iOS.
  if (/iPad|iPhone|iPod/.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1)) {
    const vp = document.querySelector('meta[name="viewport"]');
    if (vp && !/maximum-scale/.test(vp.content)) vp.content += ', maximum-scale=1';
    // Safari recuerda el zoom de la pestaña al recargar; reaplicar la etiqueta lo devuelve a 1
    const reset = () => {
      if (!vp) return;
      const base = vp.content.replace(/,\s*maximum-scale=1/, '');
      vp.content = base;
      requestAnimationFrame(() => { vp.content = base + ', maximum-scale=1'; });
    };
    window.addEventListener('pageshow', () => setTimeout(reset, 50));
    setTimeout(reset, 300);
  }

  const API_URL = 'https://api.todoist.com/api/v1/sync';
  const KEYS = {
    token: 'mp.todoist.token',
    cache: 'mp.todoist.cache',
    datos: 'mp.datos',
    base: 'mp.datos.base',
    onedrive: 'mp.onedrive',
    google: 'mp.google',
    planner: 'mp.planner',
    calendar: 'mp.calendar',
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
    gg: Object.assign({ clientId: '', account: '', token: null, exp: 0, syncedAt: null }, store.get(KEYS.google, {})),
    pl: Object.assign({ account: '', plans: [], planId: null, buckets: [], tasks: null, cats: {}, syncedAt: null }, store.get(KEYS.planner, {})),
    plState: 'idle', // idle | syncing | login | error
    cal: Object.assign({ events: [], from: null, to: null, syncedAt: null }, store.get(KEYS.calendar, {})),
    calState: 'idle', // idle | syncing | error
    ggState: 'idle', // idle | syncing | login | error
    ui: Object.assign({ sidebarHidden: false, projectId: null, view: 'personal' }, store.get(KEYS.ui, {})),
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

  const saveDatos = () => { store.set(KEYS.datos, S.datos); odSchedulePush(); calSchedule(); };
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
    file: () => svg([{ d: 'M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z' }, { d: 'M14 3v5h5M9 13h6M9 17h4' }]),
    check: () => svg([{ d: 'M5 12l5 5 9-10' }]),
    calendar: () => svg([{ tag: 'rect', x: '3', y: '5', width: '18', height: '16', rx: '2' }, { d: 'M16 3v4M8 3v4M3 10h18' }]),
    left: () => svg([{ d: 'M15 6l-6 6 6 6' }]),
    right: () => svg([{ d: 'M9 6l6 6-6 6' }]),
    cake: () => svg([{ d: 'M4 21h16M5 21v-6a2 2 0 0 1 2-2h10a2 2 0 0 1 2 2v6M5 17c1.5 1 3 1 4.5 0s3-1 4.5 0 3 1 5 0M12 13V9' }]),
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
      // Vuelta del login de Microsoft (la respuesta llega en la dirección de la página)
      try {
        const res = await app.handleRedirectPromise();
        if (res && res.account) {
          if (res.state === 'planner' || !isPersonalAcc(res.account)) plLoggedIn(res.account);
          else odLoggedIn(res.account, true);
        }
      } catch (e) {
        const cancelled = /access_denied|user_cancelled/.test((e && (e.errorCode + ' ' + e.message)) || '');
        toast(cancelled ? 'Has cancelado el inicio de sesión de Microsoft.' : 'No se pudo iniciar sesión con Microsoft: ' + (e.errorMessage || e.message || e.errorCode), 'error', null, 10000);
      }
      return app;
    });
    msalLoading.catch(() => { msalLoading = null; });
    return msalLoading;
  }

  // En la caché de MSAL conviven la cuenta personal (OneDrive) y la del trabajo (Planner):
  // la personal es la del directorio de cuentas de Microsoft (consumers)
  const CONSUMERS_TID = '9188040d-6c67-4c5b-b112-36a304b66dad';
  const isPersonalAcc = a => a && (a.tenantId === CONSUMERS_TID || (a.idTokenClaims && a.idTokenClaims.tid === CONSUMERS_TID));
  function odAccount() {
    if (!msalApp) return null;
    return msalApp.getAllAccounts().find(isPersonalAcc) || null;
  }

  function odLoggedIn(account, notify) {
    S.od.account = account.username || account.name || 'Cuenta de Microsoft';
    saveOd();
    S.odState = 'idle';
    renderSyncState();
    if (notify) toast('OneDrive conectado');
    syncDatos();
  }

  // Login por redirección: la página va a Microsoft y vuelve; al volver lo procesa loadMsal().
  // (Sin ventana emergente: el login de Microsoft aísla su ventana y MSAL 4 deja de poder leerla.)
  async function odConnect() {
    try {
      if (!msalApp) await loadMsal();
      clearStaleInteraction();
      const acc = odAccount();
      const req = { scopes: OD_SCOPES, prompt: 'select_account' };
      if (acc) req.loginHint = acc.username;
      await msalApp.loginRedirect(req);
    } catch (e) {
      toast('No se pudo conectar con OneDrive: ' + (e.errorMessage || e.message || e.errorCode), 'error', null, 10000);
    }
  }

  // Un intento anterior que no terminó deja la marca de "inicio de sesión en curso" y bloquea
  // los siguientes (interaction_in_progress); al pulsar Conectar se descarta
  function clearStaleInteraction() {
    for (const st of [sessionStorage, localStorage]) {
      try { Object.keys(st).filter(k => k.startsWith('msal.') && k.endsWith('interaction.status')).forEach(k => st.removeItem(k)); }
      catch (e) { /* almacenamiento no disponible */ }
    }
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

  // ---------- Trabajo: Microsoft Planner ----------
  // Misma App Registration y misma MSAL que OneDrive, con la cuenta del trabajo (autoridad
  // "organizations"). Graph v1.0; cada escritura lleva el etag de la tarea (If-Match).
  const PL_SCOPES = ['Tasks.ReadWrite', 'User.Read'];
  const ORG_AUTH = 'https://login.microsoftonline.com/organizations';
  const GRAPH = 'https://graph.microsoft.com/v1.0';
  // Colores aproximados de las 25 etiquetas de Planner (category1 … category25)
  const PL_COLORS = [['Rosa', '#E68AC4'], ['Rojo', '#F0857A'], ['Amarillo', '#F5D76E'], ['Verde', '#8BD17C'], ['Azul', '#7FA2E0'],
    ['Morado', '#B39DF0'], ['Bronce', '#D9A77A'], ['Lima', '#DBF266'], ['Aguamarina', '#6FC3E8'], ['Gris', '#9AA4B5'],
    ['Plata', '#C9D0DC'], ['Marrón', '#B08968'], ['Arándano', '#C9607F'], ['Naranja', '#F2A65A'], ['Melocotón', '#F5B9A0'],
    ['Caléndula', '#F2C14E'], ['Verde claro', '#B5E48C'], ['Verde oscuro', '#5C9E6E'], ['Verde azulado', '#5CCFB8'], ['Azul claro', '#A8C8F0'],
    ['Azul oscuro', '#5F7FC0'], ['Lavanda', '#C9B6F2'], ['Ciruela', '#A7679F'], ['Gris claro', '#B8BFCB'], ['Gris oscuro', '#6E7A8E']];
  class PlError extends Error {
    constructor(kind, message) { super(message); this.kind = kind; }
  }
  const isWork = () => S.ui.view === 'trabajo';
  const savePl = () => store.set(KEYS.planner, S.pl);
  const plConnected = () => !!(S.od.clientId && S.pl.account);
  function plAccount() {
    if (!msalApp) return null;
    return msalApp.getAllAccounts().find(a => !isPersonalAcc(a) && (!S.pl.account || a.username === S.pl.account)) || null;
  }

  async function plConnect() {
    if (!S.od.clientId) { toast('Primero pon el Id. de aplicación de Microsoft en Ajustes → OneDrive.', 'error'); return; }
    try {
      if (!msalApp) await loadMsal();
      clearStaleInteraction();
      const req = { scopes: PL_SCOPES, authority: ORG_AUTH, prompt: 'select_account', state: 'planner' };
      if (S.pl.account) { req.loginHint = S.pl.account; delete req.prompt; }
      await msalApp.loginRedirect(req);
    } catch (e) {
      toast('No se pudo conectar con Planner: ' + (e.errorMessage || e.message || e.errorCode), 'error', null, 10000);
    }
  }

  function plLoggedIn(account) {
    S.pl.account = account.username || account.name || 'Cuenta del trabajo';
    savePl();
    S.plState = 'idle';
    toast('Planner conectado');
    plSync();
  }

  async function plDisconnect() {
    try {
      if (!msalApp) await loadMsal();
      const acc = plAccount();
      if (acc) await msalApp.clearCache({ account: acc });
    } catch (e) { /* se olvida igualmente */ }
    S.pl = { account: '', plans: [], planId: null, buckets: [], tasks: null, cats: {}, syncedAt: null };
    savePl();
    S.plState = 'idle';
    renderAll();
  }

  async function plToken() {
    if (!msalApp) await loadMsal();
    const acc = plAccount();
    if (!acc) throw new PlError('login', 'Conecta de nuevo tu cuenta del trabajo.');
    try {
      const r = await msalApp.acquireTokenSilent({ scopes: PL_SCOPES, account: acc, authority: 'https://login.microsoftonline.com/' + (acc.tenantId || 'organizations') });
      return r.accessToken;
    } catch (e) {
      if (e instanceof msal.InteractionRequiredAuthError || (e && /interaction_required|login_required|consent_required|no_tokens_found/.test(e.errorCode || ''))) {
        throw new PlError('login', 'La sesión del trabajo ha caducado. Vuelve a conectar.');
      }
      throw new PlError('network', 'No se pudo conectar con Microsoft. Revisa la conexión.');
    }
  }

  async function plCall(method, path, body, etag) {
    const token = await plToken();
    const headers = { Authorization: 'Bearer ' + token };
    if (body !== undefined) headers['Content-Type'] = 'application/json';
    if (etag) headers['If-Match'] = etag;
    if (method === 'PATCH') headers.Prefer = 'return=representation';
    let r;
    try { r = await fetch(/^https:/.test(path) ? path : GRAPH + path, { method, headers, body: body === undefined ? undefined : JSON.stringify(body) }); }
    catch (e) { throw new PlError('network', 'No se pudo conectar con Planner. Revisa la conexión.'); }
    if (r.status === 401) throw new PlError('login', 'La sesión del trabajo ha caducado. Vuelve a conectar.');
    if (r.status === 412 || r.status === 409) throw new PlError('conflict', 'La tarea ha cambiado en Planner; se vuelve a leer.');
    if (r.status === 403) throw new PlError('http', 'Planner no permite esta acción (403). Revisa el permiso Tasks.ReadWrite de la app.');
    if (r.status === 204) return {};
    let data = null;
    try { data = await r.json(); } catch (e) { /* sin cuerpo */ }
    if (!r.ok) throw new PlError('http', 'Planner respondió con un error (' + r.status + ')' + (data && data.error && data.error.message ? ': ' + data.error.message : '') + '.');
    return data || {};
  }
  async function plList(path) {
    let out = [], next = path;
    while (next) { const r = await plCall('GET', next); out = out.concat(r.value || []); next = r['@odata.nextLink'] || null; }
    return out;
  }
  const plEtag = o => o['@odata.etag'] || o.etag || null;
  const plTaskFrom = t => ({
    id: t.id, title: t.title || '', bucketId: t.bucketId, planId: t.planId, orderHint: t.orderHint || '',
    cats: Object.keys(t.appliedCategories || {}).filter(k => t.appliedCategories[k]).sort((a, b) => +a.slice(8) - +b.slice(8)),
    hasDescription: !!t.hasDescription, etag: plEtag(t)
  });
  const normTitle = s => String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();

  async function plSync(opts = {}) {
    if (!plConnected() || S.plState === 'syncing') return;
    S.plState = 'syncing';
    renderSyncState();
    try {
      const plans = (await plList('/me/planner/plans')).map(p => ({ id: p.id, title: p.title }));
      plans.sort((a, b) => a.title.localeCompare(b.title, 'es'));
      let planId = plans.some(p => p.id === S.pl.planId) ? S.pl.planId : null;
      if (!planId) planId = (plans.find(p => normTitle(p.title) === 'tareas mias') || plans[0] || {}).id || null;
      let buckets = [], tasks = [], cats = Object.assign({}, S.pl.cats);
      if (planId) {
        const [bk, tk, det] = await Promise.all([
          plList('/planner/plans/' + planId + '/buckets'),
          plList('/planner/plans/' + planId + '/tasks'),
          plCall('GET', '/planner/plans/' + planId + '/details')
        ]);
        buckets = bk.map(b => ({ id: b.id, name: b.name, orderHint: b.orderHint || '' })).sort((a, b) => a.orderHint < b.orderHint ? 1 : a.orderHint > b.orderHint ? -1 : 0); // Planner los muestra de mayor a menor
        tasks = tk.filter(t => (t.percentComplete || 0) < 100).map(plTaskFrom);
        cats[planId] = { etag: plEtag(det), names: det.categoryDescriptions || {} };
      }
      Object.assign(S.pl, { plans, planId, buckets, tasks, cats, syncedAt: Date.now() });
      savePl();
      S.plState = 'idle';
      if (opts.notify) toast('Sincronizado con Planner');
    } catch (e) {
      S.plState = e.kind === 'login' ? 'login' : 'error';
      if (e.kind === 'login') toast(e.message, 'error', { label: 'Conectar', fn: plConnect }, 10000);
      else toast(e.message || 'Error al leer Planner.', 'error');
    } finally {
      renderAll();
    }
  }

  // Si una escritura falla se avisa y se vuelve a leer Planner
  async function plRecover(e) {
    if (e.kind === 'login') { S.plState = 'login'; renderSyncState(); toast(e.message, 'error', { label: 'Conectar', fn: plConnect }, 10000); return; }
    toast(e.message || 'Planner no aceptó el cambio.', 'error');
    await plSync();
  }

  const plTask = id => (S.pl.tasks || []).find(t => t.id === id);
  // Etiquetas del plan: las que tienen nombre o están puestas en alguna tarea
  function plTags() {
    const names = ((S.pl.cats || {})[S.pl.planId] || {}).names || {};
    const used = new Set((S.pl.tasks || []).filter(t => t.planId === S.pl.planId).flatMap(t => t.cats));
    const out = [];
    for (let i = 1; i <= 25; i++) {
      const k = 'category' + i, name = names[k];
      if (name || used.has(k)) out.push({ id: k, name: name || PL_COLORS[i - 1][0], color: PL_COLORS[i - 1][1] });
    }
    return out;
  }
  function plColumns() {
    if (!S.pl.tasks || !S.pl.planId) return [];
    // Planner ordena por orderHint (comparación ordinal; el menor va primero)
    const ord = (a, b) => a.orderHint < b.orderHint ? -1 : a.orderHint > b.orderHint ? 1 : 0;
    const items = S.pl.tasks.filter(t => t.planId === S.pl.planId);
    return S.pl.buckets.map(b => ({ key: b.id, name: b.name, items: items.filter(t => t.bucketId === b.id).sort(ord) }));
  }
  function plReplace(t) {
    const i = S.pl.tasks.findIndex(x => x.id === t.id);
    if (i >= 0) S.pl.tasks[i] = t; else S.pl.tasks.push(t);
    savePl();
  }

  async function plDrop(evt) {
    const id = evt.item.dataset.taskId;
    const toKey = evt.to.dataset.sectionKey, fromKey = evt.from.dataset.sectionKey;
    if (fromKey === toKey && evt.oldIndex === evt.newIndex) return;
    const t = plTask(id);
    if (!t || id.startsWith('tmp-')) return;
    // Nuevo orderHint entre las tarjetas vecinas: "<anterior> <siguiente>!" (Planner calcula el valor)
    const ids = [...evt.to.querySelectorAll('.card')].map(el => el.dataset.taskId);
    const pos = ids.indexOf(id);
    const prev = pos > 0 ? plTask(ids[pos - 1]) : null, next = pos < ids.length - 1 ? plTask(ids[pos + 1]) : null;
    const body = { orderHint: (prev ? prev.orderHint : '') + ' ' + (next ? next.orderHint : '') + '!' };
    if (fromKey !== toKey) { body.bucketId = toKey; t.bucketId = toKey; }
    // en pantalla, al momento: hint provisional entre los vecinos
    t.orderHint = prev && next ? prev.orderHint + '~' : prev ? prev.orderHint + '~' : next ? '' : t.orderHint;
    savePl();
    setTimeout(renderBoard, 0);
    try { plReplace(plTaskFrom(await plCall('PATCH', '/planner/tasks/' + id, body, t.etag))); renderBoard(); }
    catch (e) { await plRecover(e); }
  }

  async function plComplete(id) {
    const t = plTask(id);
    if (!t) return;
    S.pl.tasks = S.pl.tasks.filter(x => x.id !== id);
    savePl();
    renderBoard();
    try {
      const done = plTaskFrom(await plCall('PATCH', '/planner/tasks/' + id, { percentComplete: 100 }, t.etag));
      toast('Tarea completada', 'ok', { label: 'Deshacer', fn: () => plUncomplete(done) });
    } catch (e) { await plRecover(e); }
  }
  async function plUncomplete(t) {
    S.pl.tasks.push(t);
    savePl();
    renderBoard();
    try { plReplace(plTaskFrom(await plCall('PATCH', '/planner/tasks/' + t.id, { percentComplete: 0 }, t.etag))); renderBoard(); }
    catch (e) { await plRecover(e); }
  }

  // Etiquetas de una tarea: appliedCategories con true/false para cada cambio
  async function plSetCats(id, ids) {
    const t = plTask(id);
    if (!t || id.startsWith('tmp-')) return;
    const applied = {};
    for (const k of t.cats) if (!ids.includes(k)) applied[k] = false;
    for (const k of ids) if (!t.cats.includes(k)) applied[k] = true;
    if (!Object.keys(applied).length) return;
    t.cats = ids.slice().sort((a, b) => +a.slice(8) - +b.slice(8));
    savePl();
    renderBoard();
    try { plReplace(plTaskFrom(await plCall('PATCH', '/planner/tasks/' + id, { appliedCategories: applied }, t.etag))); renderBoard(); }
    catch (e) { await plRecover(e); }
  }

  async function plSaveDrawer(ed, name, tagIds) {
    if (ed.mode === 'edit') {
      const t = plTask(ed.taskId);
      if (!t) return;
      const body = {};
      if (name !== ed.original.trim()) { body.title = name; t.title = name; }
      const applied = {};
      for (const k of t.cats) if (!tagIds.includes(k)) applied[k] = false;
      for (const k of tagIds) if (!t.cats.includes(k)) applied[k] = true;
      if (Object.keys(applied).length) { body.appliedCategories = applied; t.cats = tagIds.slice(); }
      const notesChanged = ed.notesLoaded && ed.notes !== ed.notesOrig;
      if (notesChanged) t.hasDescription = !!ed.notes.trim();
      savePl();
      renderBoard();
      try {
        if (Object.keys(body).length) plReplace(plTaskFrom(await plCall('PATCH', '/planner/tasks/' + t.id, body, t.etag)));
        if (notesChanged) {
          await plCall('PATCH', '/planner/tasks/' + t.id + '/details', { description: ed.notes }, ed.detailsEtag);
          const fresh = plTask(t.id);
          if (fresh) { fresh.hasDescription = !!ed.notes.trim(); savePl(); }
        }
        renderBoard();
      } catch (e) { await plRecover(e); }
      return;
    }
    // Nueva tarea, al final de la columna
    const col = plColumns().find(c => c.key === ed.sectionKey) || plColumns()[0];
    if (!col) return;
    const last = col.items[col.items.length - 1];
    const tmp = { id: 'tmp-' + uid(), title: name, bucketId: col.key, planId: S.pl.planId, orderHint: (last ? last.orderHint : '') + '~', cats: tagIds.slice(), hasDescription: false, etag: null };
    S.pl.tasks.push(tmp);
    renderBoard();
    const applied = {};
    for (const k of tagIds) applied[k] = true;
    try {
      const created = await plCall('POST', '/planner/tasks', { planId: S.pl.planId, bucketId: col.key, title: name, orderHint: (last ? last.orderHint : '') + ' !', appliedCategories: applied });
      S.pl.tasks = S.pl.tasks.filter(x => x.id !== tmp.id);
      plReplace(plTaskFrom(created));
      renderBoard();
    } catch (e) {
      S.pl.tasks = S.pl.tasks.filter(x => x.id !== tmp.id);
      await plRecover(e);
    }
  }

  // Notas de la tarea (campo "Notas" de Planner): se leen al abrir el panel y se guardan con Guardar
  function plNotesSection(ed) {
    const area = h('textarea', { class: 'input textarea', rows: '5', placeholder: 'Cargando notas…', 'aria-label': 'Notas', disabled: true });
    area.addEventListener('input', () => { ed.notes = area.value; });
    plCall('GET', '/planner/tasks/' + ed.taskId + '/details').then(d => {
      if (S.editing !== ed) return;
      ed.notes = ed.notesOrig = d.description || '';
      ed.detailsEtag = plEtag(d);
      ed.notesLoaded = true;
      area.value = ed.notes;
      area.disabled = false;
      area.placeholder = 'Sin notas';
    }).catch(e => {
      area.placeholder = e.kind === 'login' ? 'Conecta tu cuenta del trabajo para ver las notas.' : 'No se pudieron leer las notas.';
    });
    return h('div', { class: 'field' }, h('label', { class: 'field-label' }, 'Notas'), area, h('p', { class: 'hint' }, 'Se guardan en Planner con «Guardar».'));
  }

  // Nombres de las etiquetas del plan (los colores los fija Planner)
  function openPlLabels() {
    const det = (S.pl.cats || {})[S.pl.planId];
    if (!det) { toast('Primero sincroniza con Planner.', 'error'); return; }
    const inputs = {};
    const rows = PL_COLORS.map(([color, hex], i) => {
      const k = 'category' + (i + 1);
      const inp = h('input', { class: 'input tm-name', type: 'text', maxlength: '100', placeholder: color + ' (sin nombre)', 'aria-label': 'Nombre de la etiqueta ' + color });
      inp.value = det.names[k] || '';
      inputs[k] = inp;
      return h('div', { class: 'tm-row' }, h('div', { class: 'tm-main' }, h('span', { class: 'swatch static', style: '--c:' + hex, title: color }), inp));
    });
    let close = null;
    const save = async () => {
      const changes = {};
      for (const [k, inp] of Object.entries(inputs)) {
        const v = inp.value.trim();
        if (v !== (det.names[k] || '')) changes[k] = v || null;
      }
      close();
      if (!Object.keys(changes).length) return;
      Object.assign(det.names, changes);
      savePl();
      renderTagbar();
      renderBoard();
      try {
        await plCall('PATCH', '/planner/plans/' + S.pl.planId + '/details', { categoryDescriptions: changes }, det.etag);
        await plSync();
      } catch (e) { await plRecover(e); }
    };
    close = openModal('Etiquetas de Planner', h('div', { class: 'form' },
      h('p', { class: 'hint' }, 'Son las etiquetas de este plan en Planner: el nombre que pongas aquí se verá también allí. Los colores los fija Planner.'),
      h('div', { class: 'tm-list pl-labels' }, rows),
      h('div', { class: 'row-actions modal-foot' }, h('span', { class: 'spacer' }),
        h('button', { type: 'button', class: 'btn subtle small', onclick: () => close() }, 'Cancelar'),
        h('button', { type: 'button', class: 'btn primary small', onclick: save }, 'Guardar'))));
  }

  // ---------- Google (contactos: lectura y escritura) ----------
  // OAuth 2.0 para aplicaciones de navegador (flujo implícito) por redirección: la página va a Google
  // y vuelve con un token de acceso de una hora en el fragmento de la dirección. Sin librerías de Google.
  const GG_SCOPES = 'https://www.googleapis.com/auth/contacts https://www.googleapis.com/auth/calendar.events https://www.googleapis.com/auth/calendar.calendarlist.readonly email';
  const GG_STATE = 'mp.google.state';
  const GG_NEXT = 'mp.google.next'; // qué hacer al volver de Google: { edit: id|null, draft }
  const PERSON_FIELDS = 'names,emailAddresses,phoneNumbers,birthdays,biographies,memberships,addresses,organizations';
  const PEOPLE = 'https://people.googleapis.com/v1/';
  class GgError extends Error {
    constructor(kind, message) { super(message); this.kind = kind; }
  }
  const saveGg = () => store.set(KEYS.google, S.gg);
  const ggConnected = () => !!(S.gg.clientId && S.gg.account);
  const ggTokenOk = () => !!S.gg.token && Date.now() < S.gg.exp - 60000;

  // consent: vuelve a mostrar la pantalla de permisos (si la vez anterior no se marcó el de contactos)
  function ggLogin(consent) {
    const state = uid();
    sessionStorage.setItem(GG_STATE, state);
    const q = new URLSearchParams({ client_id: S.gg.clientId, redirect_uri: appUrl(), response_type: 'token', scope: GG_SCOPES, include_granted_scopes: 'true', state });
    if (S.gg.account) q.set('login_hint', S.gg.account);
    if (consent === true) q.set('prompt', S.gg.account ? 'consent' : 'consent select_account');
    else if (!S.gg.account) q.set('prompt', 'select_account');
    location.assign('https://accounts.google.com/o/oauth2/v2/auth?' + q);
  }

  // Vuelta del login de Google. Solo se procesa si el "state" es el que se guardó al salir
  // (así no se confunde con la vuelta de Microsoft).
  function ggHandleReturn() {
    const hash = location.hash.slice(1);
    if (!/(^|&)(access_token|error)=/.test(hash)) return;
    const q = new URLSearchParams(hash);
    const saved = sessionStorage.getItem(GG_STATE);
    if (!saved || q.get('state') !== saved) return;
    sessionStorage.removeItem(GG_STATE);
    history.replaceState(null, '', location.pathname + location.search);
    if (q.get('error')) {
      toast(q.get('error') === 'access_denied' ? 'Has cancelado la conexión con Google.' : 'Google no permitió la conexión (' + q.get('error') + ').', 'error', null, 10000);
      return;
    }
    const scope = q.get('scope') || '';
    if (!/auth\/contacts(\.readonly)?(\s|$)/.test(scope)) {
      sessionStorage.removeItem(GG_NEXT);
      toast('Google no ha dado permiso para tus contactos: marca la casilla de contactos.', 'error', { label: 'Volver a pedir', fn: () => ggLogin(true) }, 15000);
      return;
    }
    S.gg.write = /auth\/contacts(\s|$)/.test(scope);
    S.gg.cal = /auth\/calendar\.events(\s|$)/.test(scope) && /calendar\.calendarlist/.test(scope);
    S.gg.token = q.get('access_token');
    S.gg.exp = Date.now() + (parseInt(q.get('expires_in'), 10) || 3600) * 1000;
    saveGg();
    ggAfterLogin();
  }

  async function ggAfterLogin() {
    try {
      if (!S.gg.account) {
        const me = await ggGet('https://www.googleapis.com/oauth2/v3/userinfo');
        S.gg.account = me.email || 'Cuenta de Google';
        saveGg();
        toast('Google conectado');
      }
    } catch (e) { /* se reintenta al sincronizar */ }
    S.ggState = 'idle';
    let next = null;
    try { next = JSON.parse(sessionStorage.getItem(GG_NEXT) || 'null'); } catch (e) { /* nada */ }
    sessionStorage.removeItem(GG_NEXT);
    await syncContactos();
    applyView();
    calSync();
    if (next && 'edit' in next) {
      if (!S.gg.write) { toast('Para editar hay que marcar el permiso de editar contactos en Google.', 'error', { label: 'Volver a pedir', fn: () => ggLogin(true) }, 15000); return; }
      if (S.ui.view !== 'contactos') showView('contactos');
      openEditContacto(next.edit, next.draft || null);
    }
  }

  // Antes de escribir: hace falta un permiso vigente que incluya editar. Si no, se va a Google
  // y al volver se continúa con "next" (abrir el formulario, con el borrador si lo había).
  function ggEnsureWrite(next) {
    if (ggTokenOk() && S.gg.write) return true;
    sessionStorage.setItem(GG_NEXT, JSON.stringify(next));
    ggLogin(!S.gg.write);
    return false;
  }

  function ggDisconnect() {
    if (S.gg.token) fetch('https://oauth2.googleapis.com/revoke?token=' + encodeURIComponent(S.gg.token), { method: 'POST', mode: 'no-cors' }).catch(() => {});
    S.gg.account = '';
    S.gg.token = null;
    S.gg.exp = 0;
    S.gg.syncedAt = null;
    saveGg();
    S.ggState = 'idle';
    renderSyncState();
  }

  const ggGet = url => ggCall('GET', url);
  async function ggCall(method, url, body) {
    if (!ggTokenOk()) throw new GgError('login', 'La sesión de Google ha caducado. Vuelve a conectar.');
    const opts = { method, headers: { Authorization: 'Bearer ' + S.gg.token } };
    if (body !== undefined) { opts.headers['Content-Type'] = 'application/json'; opts.body = JSON.stringify(body); }
    let r;
    try { r = await fetch(url, opts); }
    catch (e) { throw new GgError('network', 'No se pudo conectar con Google. Revisa la conexión.'); }
    if (r.status === 401) { S.gg.token = null; saveGg(); throw new GgError('login', 'La sesión de Google ha caducado. Vuelve a conectar.'); }
    let data = null;
    try { data = await r.json(); } catch (e) { /* sin cuerpo */ }
    if (r.ok) return data || {};
    const msg = (data && data.error && (data.error.message || data.error.status)) || '';
    const esCal = url.includes('/calendar/');
    if (r.status === 403 && esCal && /insufficient|scope|PERMISSION/i.test(msg)) { S.gg.cal = false; saveGg(); throw new GgError('scope', 'Google no deja usar el calendario con el permiso actual.'); }
    if (r.status === 403 && esCal) throw new GgError('http', 'Google Calendar no responde (403). Revisa que la Google Calendar API esté activada en tu proyecto.' + (msg ? ' ' + msg : ''));
    if (r.status === 403 && method !== 'GET') { S.gg.write = false; saveGg(); throw new GgError('scope', 'Google no deja modificar tus contactos con el permiso actual.'); }
    if (r.status === 403) throw new GgError('http', 'Google no deja leer los contactos (403). Revisa que la People API esté activada en tu proyecto.');
    if (r.status === 400 && /etag|FAILED_PRECONDITION|out of date/i.test(msg)) throw new GgError('conflict', 'Este contacto ha cambiado en Google mientras lo editabas.');
    throw new GgError('http', 'Google respondió con un error (' + r.status + ')' + (msg ? ': ' + msg : '') + '.');
  }

  // Contacto de Google -> registro de la colección "contactos". Guarda lo necesario para volver a
  // escribirlo sin perder datos: etag (versión), nombre de pila y apellidos, tipos originales y grupos.
  function contactFrom(p, groups) {
    const prim = arr => (arr || []).find(x => x.metadata && x.metadata.primary) || (arr || [])[0];
    const nm = prim(p.names);
    const bd = prim((p.birthdays || []).filter(b => b.date && b.date.month && b.date.day));
    const two = n => String(n).padStart(2, '0');
    const org = prim(p.organizations);
    const id = 'g-' + String(p.resourceName).replace(/^people\//, '');
    const labels = [], grupos = [];
    let favorito = false;
    for (const m of p.memberships || []) {
      const rn = m.contactGroupMembership && m.contactGroupMembership.contactGroupResourceName;
      if (rn === 'contactGroups/starred') favorito = true;
      else if (groups[rn]) { labels.push(groups[rn]); grupos.push(rn); }
    }
    const items = arr => (arr || []).map(x => ({ valor: x.value, tipo: x.type || '', etiqueta: x.formattedType || '' })).filter(x => x.valor);
    return {
      id, resourceName: p.resourceName, etag: p.etag || null,
      nombre: (nm && nm.displayName) || [nm && nm.givenName, nm && nm.familyName].filter(Boolean).join(' ') || (prim(p.emailAddresses) || {}).value || 'Sin nombre',
      nombrePila: (nm && nm.givenName) || '', apellidos: (nm && nm.familyName) || '',
      telefonos: items(p.phoneNumbers),
      emails: items(p.emailAddresses),
      cumple: bd ? (bd.date.year ? bd.date.year + '-' : '--') + two(bd.date.month) + '-' + two(bd.date.day) : null,
      notas: ((p.biographies || [])[0] || {}).value || '',
      etiquetas: labels.sort((a, b) => a.localeCompare(b, 'es')),
      grupos: grupos.sort(),
      favorito,
      direcciones: (p.addresses || []).map(x => ({ valor: x.formattedValue || [x.streetAddress, x.postalCode, x.city].filter(Boolean).join(', '), tipo: x.type || '', etiqueta: x.formattedType || '' })).filter(x => x.valor),
      org: { nombre: (org && org.name) || '', cargo: (org && org.title) || '' },
      empresa: org ? [org.name, org.title].filter(Boolean).join(' · ') : ''
    };
  }

  // Categorías del usuario: colección "gruposContactos" ({id: resourceName, nombre})
  const grupoMap = () => Object.fromEntries(Object.values(S.datos.colecciones.gruposContactos || {}).map(g => [g.id, g.nombre]));

  async function syncContactos(opts = {}) {
    if (!ggConnected() && !S.gg.token) return;
    if (S.ggState === 'syncing') return;
    S.ggState = 'syncing';
    renderSyncState();
    try {
      const groups = {}, grecs = {};
      let page = '';
      do {
        const r = await ggGet(PEOPLE + 'contactGroups?pageSize=1000' + (page ? '&pageToken=' + encodeURIComponent(page) : ''));
        for (const g of r.contactGroups || []) if (g.groupType === 'USER_CONTACT_GROUP') {
          groups[g.resourceName] = g.formattedName || g.name;
          grecs[g.resourceName] = { id: g.resourceName, nombre: groups[g.resourceName] };
        }
        page = r.nextPageToken || '';
      } while (page);
      const recs = {};
      do {
        const r = await ggGet(PEOPLE + 'people/me/connections?personFields=' + PERSON_FIELDS + '&pageSize=1000' + (page ? '&pageToken=' + encodeURIComponent(page) : ''));
        for (const p of r.connections || []) { const c = contactFrom(p, groups); recs[c.id] = c; }
        page = r.nextPageToken || '';
      } while (page);
      if (stable(recs) !== stable(S.datos.colecciones.contactos || {}) || stable(grecs) !== stable(S.datos.colecciones.gruposContactos || {})) {
        S.datos.colecciones.contactos = recs;
        S.datos.colecciones.gruposContactos = grecs;
        saveDatos();
      }
      S.gg.syncedAt = Date.now();
      saveGg();
      S.ggState = 'idle';
      if (S.ui.view === 'contactos' || S.ui.view === 'calendario') applyView();
      if (opts.notify) toast(Object.keys(recs).length + ' contactos actualizados desde Google');
    } catch (e) {
      S.ggState = e.kind === 'login' ? 'login' : 'error';
      if (e.kind === 'login') toast(e.message, 'error', { label: 'Conectar', fn: ggLogin }, 10000);
      else toast(e.message || 'Error al leer los contactos de Google.', 'error');
    } finally {
      renderSyncState();
    }
  }

  // ---- escritura en Google ----
  // f: { nombrePila, apellidos, telefonos, emails, cumple: {dia, mes, anio}|null, notas, org, direcciones, categorias: [nombres], favorito }
  function personBody(f) {
    const list = (arr, key) => arr.filter(x => x.valor.trim()).map(x => ({ [key]: x.valor.trim(), type: x.tipo || 'other' }));
    const body = {
      names: [{ givenName: f.nombrePila.trim(), familyName: f.apellidos.trim() }],
      phoneNumbers: list(f.telefonos, 'value'),
      emailAddresses: list(f.emails, 'value'),
      birthdays: f.cumple ? [{ date: Object.assign({ month: f.cumple.mes, day: f.cumple.dia }, f.cumple.anio ? { year: f.cumple.anio } : {}) }] : [],
      biographies: f.notas.trim() ? [{ value: f.notas.trim(), contentType: 'TEXT_PLAIN' }] : [],
      organizations: f.org.nombre.trim() || f.org.cargo.trim() ? [{ name: f.org.nombre.trim(), title: f.org.cargo.trim() }] : [],
      addresses: list(f.direcciones, 'formattedValue')
    };
    return body;
  }

  async function saveContacto(cur, f) {
    const groups = grupoMap();
    let person;
    if (cur) {
      person = await ggCall('PATCH', PEOPLE + cur.resourceName + ':updateContact?updatePersonFields=names,phoneNumbers,emailAddresses,birthdays,biographies,organizations,addresses&personFields=' + PERSON_FIELDS,
        Object.assign({ etag: cur.etag }, personBody(f)));
    } else {
      person = await ggCall('POST', PEOPLE + 'people:createContact?personFields=' + PERSON_FIELDS, personBody(f));
    }
    const rn = person.resourceName;
    // Categorías: las nuevas se crean en Google; después se añade o quita el contacto de cada una
    const byName = Object.fromEntries(Object.entries(groups).map(([k, v]) => [v, k]));
    const want = [];
    for (const name of f.categorias) {
      let g = byName[name];
      if (!g) {
        const created = await ggCall('POST', PEOPLE + 'contactGroups', { contactGroup: { name } });
        g = created.resourceName;
        groups[g] = created.formattedName || created.name || name;
        S.datos.colecciones.gruposContactos = Object.assign({}, S.datos.colecciones.gruposContactos, { [g]: { id: g, nombre: groups[g] } });
      }
      want.push(g);
    }
    const had = cur ? cur.grupos || [] : [];
    if (f.favorito) want.push('contactGroups/starred');
    if (cur && cur.favorito) had.push('contactGroups/starred');
    const changes = [];
    for (const g of want) if (!had.includes(g)) changes.push([g, { resourceNamesToAdd: [rn] }]);
    for (const g of had) if (!want.includes(g)) changes.push([g, { resourceNamesToRemove: [rn] }]);
    for (const [g, body] of changes) await ggCall('POST', PEOPLE + g + '/members:modify', body);
    if (changes.length) person = await ggGet(PEOPLE + rn + '?personFields=' + PERSON_FIELDS); // etag y grupos al día
    const rec = contactFrom(person, groups);
    S.datos.colecciones.contactos = Object.assign({}, S.datos.colecciones.contactos, { [rec.id]: rec });
    saveDatos();
    return rec;
  }

  async function deleteContacto(c) {
    await ggCall('DELETE', PEOPLE + c.resourceName + ':deleteContact');
    const rest = Object.assign({}, S.datos.colecciones.contactos);
    delete rest[c.id];
    S.datos.colecciones.contactos = rest;
    const marks = Object.assign({}, S.datos.colecciones.cumpledias);
    delete marks[c.id];
    S.datos.colecciones.cumpledias = marks;
    saveDatos();
  }

  // Vuelve a leer un contacto de Google (tras un conflicto)
  async function reloadContacto(c) {
    try {
      const p = await ggGet(PEOPLE + c.resourceName + '?personFields=' + PERSON_FIELDS);
      const rec = contactFrom(p, grupoMap());
      S.datos.colecciones.contactos = Object.assign({}, S.datos.colecciones.contactos, { [rec.id]: rec });
      saveDatos();
      return rec;
    } catch (e) { return null; }
  }

  // Desde un botón: si el permiso caducó, se va a Google (vuelve y sincroniza solo)
  function ggRefresh() {
    if (!S.gg.clientId) { openSettings(); return; }
    if (ggTokenOk()) syncContactos({ notify: true });
    else ggLogin();
  }

  // ---------- Calendario (Google Calendar) ----------
  // Mismo login de Google. El calendario elegido se guarda en el almacén (colección "ajustes",
  // registro "calendario"). La app crea y mantiene sus propios eventos, marcados con
  // extendedProperties.private { mpApp: '1', mpKey }: cumpleaños, cumpledías especiales,
  // gastos marcados y tareas de Personal con fecha. Los demás eventos solo se leen.
  const CALAPI = 'https://www.googleapis.com/calendar/v3/';
  const calCfg = () => ((S.datos.colecciones.ajustes || {}).calendario) || null;
  const calReady = () => ggConnected() && !!S.gg.cal && !!calCfg();
  const saveCal = () => store.set(KEYS.calendar, S.cal);
  const ymdOf = t => new Date(t).toISOString().slice(0, 10);
  const addDays = (ymd, n) => ymdOf(utcDay(ymd) + n * DAY);
  let calTimer = null;

  async function calList() {
    const r = await ggGet(CALAPI + 'users/me/calendarList?minAccessRole=writer&maxResults=250');
    return (r.items || []).map(c => ({ id: c.id, nombre: c.summaryOverride || c.summary, color: c.backgroundColor, primary: !!c.primary }));
  }

  // Regla de repetición de un gasto: meses de cobro y día (el 29-31 pasa al último día si el mes es más corto)
  function gastoRRule(g) {
    const meses = (Array.isArray(g.meses) && g.meses.length ? g.meses : mesesPara(g.periodicidad, 1)).slice().sort((a, b) => a - b).join(',');
    const d = g.dia || 1;
    if (d <= 28) return 'RRULE:FREQ=MONTHLY;BYMONTH=' + meses + ';BYMONTHDAY=' + d;
    const dias = []; for (let i = 28; i <= d; i++) dias.push(i);
    return 'RRULE:FREQ=MONTHLY;BYMONTH=' + meses + ';BYMONTHDAY=' + dias.join(',') + ';BYSETPOS=-1';
  }

  // Nombre corto para el calendario: nombre de pila y primer apellido (o las dos primeras palabras)
  function shortName(c) {
    const pila = (c.nombrePila || '').trim(), ap = (c.apellidos || '').trim();
    if (pila) return ap ? pila + ' ' + ap.split(/\s+/)[0] : pila;
    return c.nombre.trim().split(/\s+/).slice(0, 2).join(' ');
  }

  // Eventos que la app debe tener en el calendario, por clave
  function calDesired() {
    const out = new Map();
    const ev = (key, summary, date, extra) => out.set(key, Object.assign({
      summary, start: { date }, end: { date: addDays(date, 1) }, transparency: 'transparent',
      description: 'Creado por Mi panel.', reminders: { useDefault: true },
      extendedProperties: { private: { mpApp: '1', mpKey: key } }
    }, extra));
    const today = todayUtc(), year = new Date(today).getUTCFullYear();
    // Cumpleaños y cumpledías especiales (múltiplos de 5.000 días) de los contactos seguidos
    for (const c of contactos()) {
      if (!cdMarked(c.id) || !c.cumple) continue;
      const full = /^\d{4}/.test(c.cumple);
      const mmdd = c.cumple.slice(-5);
      const start = full ? c.cumple : year + '-' + (mmdd === '02-29' && !isLeap(year) ? '02-28' : mmdd);
      ev('cumple:' + c.id, shortName(c) + ' · cumpleaños', start, {
        recurrence: [mmdd === '02-29' ? 'RRULE:FREQ=YEARLY;BYMONTH=2;BYMONTHDAY=-1' : 'RRULE:FREQ=YEARLY'], colorId: '4'
      });
      if (!full) continue;
      const birth = utcDay(c.cumple), lived = Math.round((today - birth) / DAY);
      let k = Math.max(5000, Math.ceil(lived / 5000) * 5000);
      for (let i = 0; i < 2; i++, k += 5000) ev('dias:' + c.id + ':' + k, shortName(c) + ' · ' + fmtNum(k) + ' días', ymdOf(birth + k * DAY), { colorId: '6' });
    }
    // Gastos marcados para el calendario (activos)
    for (const g of gastos()) {
      if (!g.calendario || g.activo === false) continue;
      const first = vencimientos(g, year, year + 1)[0];
      if (first === undefined) continue;
      ev('gasto:' + g.id, g.concepto + ' · ' + fmtEur(g.importe), ymdOf(first), { recurrence: [gastoRRule(g)], colorId: '5' });
    }
    // Tareas de Personal con fecha (solo si la tarea sigue en Todoist)
    for (const f of Object.values(S.datos.colecciones.fechasTareas || {})) {
      const t = S.data && S.data.items.find(i => i.id === f.id);
      if (t && isYmd(f.fecha)) ev('tarea:' + f.id, cleanTitle(t.content), f.fecha, { colorId: '2' });
    }
    return out;
  }

  const calSame = (e, d) => e.summary === d.summary && (e.start || {}).date === d.start.date && (e.end || {}).date === d.end.date &&
    (e.recurrence || []).join() === (d.recurrence || []).join() && (e.colorId || '') === (d.colorId || '');

  // Deja los eventos de la app como mandan los datos: crea, cambia o borra solo los suyos
  async function calReconcile() {
    const cid = encodeURIComponent(calCfg().calendarId);
    const mine = [];
    let page = '';
    do {
      const r = await ggGet(CALAPI + 'calendars/' + cid + '/events?privateExtendedProperty=' + encodeURIComponent('mpApp=1') + '&maxResults=2500' + (page ? '&pageToken=' + encodeURIComponent(page) : ''));
      mine.push(...(r.items || []));
      page = r.nextPageToken || '';
    } while (page);
    const want = calDesired();
    const seen = new Set();
    let cambios = 0;
    for (const e of mine) {
      const key = ((e.extendedProperties || {}).private || {}).mpKey;
      // Sin datos de Todoist cargados no se tocan los eventos de tareas
      if (key && key.startsWith('tarea:') && !S.data) { seen.add(key); continue; }
      const d = want.get(key);
      if (!d || seen.has(key)) { await ggCall('DELETE', CALAPI + 'calendars/' + cid + '/events/' + encodeURIComponent(e.id)); cambios++; continue; }
      seen.add(key);
      if (!calSame(e, d)) { await ggCall('PUT', CALAPI + 'calendars/' + cid + '/events/' + encodeURIComponent(e.id), d); cambios++; }
    }
    for (const [key, d] of want) if (!seen.has(key)) { await ggCall('POST', CALAPI + 'calendars/' + cid + '/events', d); cambios++; }
    return cambios;
  }

  // Lee los eventos (todos, ya desplegados) de un intervalo de días para la vista de mes
  async function calFetch(fromYmd, toYmd) {
    const cid = encodeURIComponent(calCfg().calendarId);
    const items = [];
    let page = '';
    do {
      const r = await ggGet(CALAPI + 'calendars/' + cid + '/events?singleEvents=true&orderBy=startTime&maxResults=2500&timeMin=' + encodeURIComponent(fromYmd + 'T00:00:00Z') + '&timeMax=' + encodeURIComponent(toYmd + 'T00:00:00Z') + (page ? '&pageToken=' + encodeURIComponent(page) : ''));
      items.push(...(r.items || []));
      page = r.nextPageToken || '';
    } while (page);
    const evs = items.filter(e => e.status !== 'cancelled').map(e => ({
      id: e.id, title: e.summary || '(sin título)', allDay: !!(e.start && e.start.date),
      start: (e.start && (e.start.date || e.start.dateTime)) || '', end: (e.end && (e.end.date || e.end.dateTime)) || '',
      key: ((e.extendedProperties || {}).private || {}).mpKey || '', link: e.htmlLink || ''
    }));
    // Se sustituye lo leído en ese intervalo y se conserva el resto
    const inRange = e => e.start.slice(0, 10) >= fromYmd && e.start.slice(0, 10) < toYmd;
    S.cal.events = S.cal.events.filter(e => !inRange(e)).concat(evs);
    S.cal.from = S.cal.from && S.cal.from < fromYmd ? S.cal.from : fromYmd;
    S.cal.to = S.cal.to && S.cal.to > toYmd ? S.cal.to : toYmd;
    saveCal();
  }

  async function calSync(opts = {}) {
    if (!calReady() || !ggTokenOk() || S.calState === 'syncing') return;
    clearTimeout(calTimer);
    S.calState = 'syncing';
    renderSyncState();
    try {
      const cambios = await calReconcile();
      const t = new Date(todayUtc());
      const from = ymdOf(Date.UTC(t.getUTCFullYear(), t.getUTCMonth() - 2, 1));
      const to = ymdOf(Date.UTC(t.getUTCFullYear(), t.getUTCMonth() + 13, 1));
      S.cal.events = [];
      await calFetch(from, to);
      S.cal.syncedAt = Date.now();
      saveCal();
      S.calState = 'idle';
      if (opts.notify) toast('Calendario sincronizado' + (cambios ? ' (' + cambios + (cambios === 1 ? ' cambio)' : ' cambios)') : ''));
    } catch (e) {
      S.calState = 'error';
      if (e.kind === 'login') { S.calState = 'idle'; toast(e.message, 'error', { label: 'Conectar', fn: ggLogin }, 10000); }
      else if (e.kind === 'scope') toast(e.message, 'error', { label: 'Dar permiso', fn: () => ggLogin(true) }, 15000);
      else toast(e.message || 'Error con Google Calendar.', 'error');
    } finally {
      renderSyncState();
      if (S.ui.view === 'calendario') renderCal();
    }
  }

  // Tras un cambio en datos que afectan al calendario, se pone al día al poco (si el permiso sigue vigente)
  function calSchedule() {
    if (!calReady() || !ggTokenOk()) return;
    clearTimeout(calTimer);
    calTimer = setTimeout(() => calSync(), 2500);
  }

  // ---------- datos derivados ----------
  function currentProjectId() {
    if (!S.data) return null;
    const ids = S.data.projects.map(p => p.id);
    if (S.ui.projectId && ids.includes(S.ui.projectId)) return S.ui.projectId;
    return S.data.inboxId || ids[0] || null;
  }

  function columns() {
    if (isWork()) return plColumns();
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

  // Etiquetas del tablero activo: las de la app en Personal, las de Planner en Trabajo
  const tagList = () => isWork() ? plTags() : S.tags.list;
  const tagById = id => tagList().find(t => t.id === id);
  const assigned = taskId => (isWork() ? ((plTask(taskId) || {}).cats || []) : (S.tags.assign[taskId] || [])).filter(id => tagById(id));
  const tagUsage = tagId => Object.values(S.tags.assign).filter(arr => arr.includes(tagId)).length;
  const boardReady = () => isWork() ? !!(S.pl.tasks && S.pl.planId) : !!S.data;

  function isVisible(item) {
    if (!S.filter.size) return true;
    return assigned(item.id).some(id => S.filter.has(id));
  }

  // ---------- render ----------
  function renderAll() {
    applyLayout();
    applyView();
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
    if (isWork()) {
      for (const p of S.pl.plans || []) sel.append(h('option', { value: p.id, selected: p.id === S.pl.planId }, p.title));
      return;
    }
    if (!S.data) return;
    const pid = currentProjectId();
    for (const p of S.data.projects) {
      sel.append(h('option', { value: p.id, selected: p.id === pid }, p.inbox ? 'Bandeja de entrada' : p.name));
    }
  }

  // Cabecera: cuándo se sincronizó la fuente de esta pantalla y, si alguna conexión necesita
  // atención (reconectar o error), un aviso que se puede pulsar
  function renderSyncState() {
    const busy = S.syncing || S.odState === 'syncing' || S.plState === 'syncing' || S.ggState === 'syncing' || S.calState === 'syncing';
    const v = S.ui.view;
    const ts = v === 'trabajo' ? S.pl.syncedAt : v === 'personal' ? (S.data && S.data.syncedAt) : v === 'contactos' ? S.gg.syncedAt : v === 'calendario' ? S.cal.syncedAt : S.od.syncedAt;
    document.getElementById('syncInfo').textContent = busy ? 'Sincronizando…' : ts ? 'Sincronizado ' + relTime(ts) : '';
    const warns = [];
    if (odConnected() && (S.odState === 'login' || S.odState === 'error')) warns.push(['OneDrive', S.odState, odConnect]);
    if (plConnected() && (S.plState === 'login' || S.plState === 'error')) warns.push(['Planner', S.plState, plConnect]);
    if (ggConnected() && S.ggState === 'error') warns.push(['Google', 'error', ggRefresh]);
    if (calReady() && S.calState === 'error') warns.push(['Calendario', 'error', openSettings]);
    const w = document.getElementById('syncWarn');
    w.hidden = !warns.length;
    w.replaceChildren(...warns.map(([name, st, fn]) => h('button', { type: 'button', class: 'warn-pill', onclick: st === 'login' ? fn : openSettings, title: st === 'login' ? 'Pulsa para volver a conectar' : 'Revisa la conexión en Ajustes' },
      name + (st === 'login' ? ': reconectar' : ': error'))));
    const btn = document.getElementById('syncBtn');
    btn.disabled = busy;
    btn.classList.toggle('is-syncing', busy);
  }

  function renderTagbar() {
    const bar = document.getElementById('tagbar');
    bar.replaceChildren();
    bar.classList.toggle('filtering', S.filter.size > 0);
    const tags = tagList();
    bar.append(h('span', { class: 'tagbar-label' }, ICON.tag(),
      h('span', null, tags.length ? 'Arrastra a una tarea · pulsa para filtrar' : isWork() ? 'Este plan aún no tiene etiquetas' : 'Aún no tienes etiquetas')));

    for (const t of tags) {
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
    end.append(h('button', { type: 'button', class: 'btn subtle small', onclick: isWork() ? openPlLabels : openTagManager },
      tags.length ? 'Gestionar' : 'Crear etiquetas'));
    bar.append(end);
  }

  function renderBoard() {
    const board = document.getElementById('board');
    S.sortables.forEach(s => s.destroy());
    S.sortables = [];
    board.replaceChildren();
    if (isWork() && !plConnected()) {
      board.append(h('div', { class: 'board-empty' },
        h('p', null, 'Conecta tu cuenta del trabajo para ver aquí tus tareas de Planner.'),
        h('button', { type: 'button', class: 'btn primary small', onclick: S.od.clientId ? plConnect : openSettings }, S.od.clientId ? 'Conectar con Planner' : 'Ir a Ajustes')));
      return;
    }
    if (!boardReady()) return;

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
    const work = isWork();
    const hasNotes = work ? item.hasDescription : (S.data.notes || []).some(n => n.item_id === item.id);
    const hasFecha = !work && !!(S.datos.colecciones.fechasTareas || {})[item.id];
    const title = work ? item.title : cleanTitle(item.content);
    const notesLabel = work ? 'Tiene notas' : 'Tiene comentarios';
    const fechaCal = !work && ((S.datos.colecciones.fechasTareas || {})[item.id] || {}).fecha;
    return h('div', {
      class: 'card' + (pending ? ' pending' : '') + (hasNotes || hasFecha ? ' has-notes' : ''),
      dataset: { taskId: item.id },
      style: '--strip:' + (tags[0] ? tags[0].color : 'transparent')
    },
      h('button', {
        type: 'button', class: 'check', 'aria-label': 'Completar tarea', disabled: pending,
        onclick: e => { e.stopPropagation(); completeTask(item.id); }
      }),
      h('button', {
        type: 'button', class: 'card-title', title,
        onclick: () => { if (!pending) openEdit(item.id); }
      }, title),
      hasNotes || fechaCal ? h('span', { class: 'card-notes' },
        fechaCal ? h('span', { class: 'card-cal', title: 'En el calendario: ' + fmtDay(utcDay(fechaCal)), 'aria-label': 'En el calendario el ' + fmtDay(utcDay(fechaCal)) }, ICON.calendar()) : null,
        hasNotes ? h('span', { title: notesLabel, 'aria-label': notesLabel }, ICON.comment()) : null) : null,
      tags.length ? h('span', { class: 'card-tags' },
        tags.map(t => h('span', { class: 'mini-tag', style: '--c:' + t.color }, t.name))) : null
    );
  }

  // ---------- acciones sobre tareas ----------
  async function onCardDrop(evt) {
    if (isWork()) return plDrop(evt);
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
    if (isWork()) return plComplete(id);
    const idx = S.data.items.findIndex(i => i.id === id);
    if (idx < 0) return;
    const [item] = S.data.items.splice(idx, 1);
    saveCache();
    // Si tenía fecha en el calendario, su evento se borra (se recupera al deshacer)
    const fecha = !item.recurring && (S.datos.colecciones.fechasTareas || {})[id];
    if (fecha) { const rest = Object.assign({}, S.datos.colecciones.fechasTareas); delete rest[id]; S.datos.colecciones.fechasTareas = rest; saveDatos(); item._fecha = fecha; }
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
    const fecha = item._fecha;
    delete item._fecha;
    S.data.items.push(item);
    saveCache();
    if (fecha) { S.datos.colecciones.fechasTareas = Object.assign({}, S.datos.colecciones.fechasTareas, { [item.id]: fecha }); saveDatos(); }
    renderBoard();
    renderProjects();
    try { await runCommands([{ type: 'item_uncomplete', uuid: uid(), args: { id: item.id } }]); }
    catch (e) { await recover(e); }
  }

  function assignTag(taskId, tagId) {
    if (!taskId || taskId.startsWith('tmp-')) return;
    if (isWork()) { const cur = assigned(taskId); if (!cur.includes(tagId)) plSetCats(taskId, cur.concat(tagId)); return; }
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
    if (!boardReady()) return;
    S.editing = { mode: 'new', sectionKey, fromHeader, name: '', tagIds: [], work: isWork(), fecha: '' };
    renderDrawer();
  }

  function openEdit(taskId) {
    const item = isWork() ? plTask(taskId) : S.data.items.find(i => i.id === taskId);
    if (!item) return;
    const name = isWork() ? item.title : item.content;
    const f = !isWork() && (S.datos.colecciones.fechasTareas || {})[taskId];
    S.editing = { mode: 'edit', taskId, name, original: name, tagIds: assigned(taskId).slice(), work: isWork(), fecha: f ? f.fecha : '' };
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

    const item = ed.mode === 'edit' ? (isWork() ? plTask(ed.taskId) : S.data.items.find(i => i.id === ed.taskId)) : null;
    const where = ed.mode === 'edit'
      ? 'en ' + columnName(isWork() ? item && item.bucketId : item && item.section_id ? item.section_id : 'none')
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
      for (const t of tagList()) {
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

    const commentsField = ed.mode === 'edit' && !ed.taskId.startsWith('tmp-') ? (isWork() ? plNotesSection(ed) : commentsSection(ed.taskId)) : null;
    // Fecha en el calendario (Personal): crea un evento de todo el día en Google Calendar
    let fechaField = null;
    if (!isWork()) {
      const inp = h('input', { class: 'input', id: 'taskFecha', type: 'date', value: ed.fecha || '' });
      inp.addEventListener('change', () => { ed.fecha = inp.value; });
      fechaField = h('div', { class: 'field' },
        h('label', { class: 'field-label', for: 'taskFecha' }, 'Fecha en el calendario'),
        h('div', { class: 'row-edit two' }, inp, h('button', { type: 'button', class: 'btn subtle small', onclick: () => { inp.value = ''; ed.fecha = ''; } }, 'Quitar')),
        h('p', { class: 'hint' }, calReady() ? 'Crea un evento con el nombre de la tarea en tu calendario. Al completar la tarea, el evento se borra.' : 'Se guardará y aparecerá en el calendario cuando conectes Google Calendar en Ajustes.'));
    }

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
        tagList().length ? chips : h('p', { class: 'hint' }, isWork() ? 'Este plan aún no tiene etiquetas.' : 'Todavía no has creado etiquetas.'),
        tagList().length
          ? h('p', { class: 'hint' }, 'Pulsa una etiqueta para asignarla o quitarla.')
          : h('div', null, h('button', { type: 'button', class: 'btn ghost small', onclick: isWork() ? openPlLabels : openTagManager }, 'Crear etiquetas'))),
      fechaField,
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
    if (ed.work) return plSaveDrawer(ed, name, tagIds);
    const setFecha = (id, fecha) => {
      const cur = (S.datos.colecciones.fechasTareas || {})[id];
      if ((cur ? cur.fecha : '') === (fecha || '')) return;
      const rest = Object.assign({}, S.datos.colecciones.fechasTareas);
      if (isYmd(fecha)) rest[id] = { id, fecha }; else delete rest[id];
      S.datos.colecciones.fechasTareas = rest;
      saveDatos();
    };

    if (ed.mode === 'edit') {
      const item = S.data.items.find(i => i.id === ed.taskId);
      if (!item) return;
      S.tags.assign[ed.taskId] = tagIds;
      saveTags();
      setFecha(ed.taskId, ed.fecha);
      const renamed = name !== ed.original.trim();
      if (renamed) { item.content = name; saveCache(); }
      renderBoard();
      if (renamed) {
        try { await runCommands([{ type: 'item_update', uuid: uid(), args: { id: item.id, content: name } }]); calSchedule(); }
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
      if (isYmd(ed.fecha)) setFecha(realId, ed.fecha);
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

  // Ajustes · Planner (cuenta del trabajo, misma App Registration que OneDrive)
  function plSection() {
    const box = h('div', { class: 'modal-section' });
    const draw = () => {
      const intro = h('p', { class: 'muted' }, 'Tus tareas de Microsoft Planner en la pantalla Trabajo, con la cuenta del trabajo. Usa el mismo Id. de aplicación que OneDrive.');
      if (!S.od.clientId) {
        box.replaceChildren(h('h3', null, 'Trabajo (Planner)'), intro, h('p', { class: 'hint' }, 'Pon primero el Id. de aplicación en la sección OneDrive.'));
        return;
      }
      if (!plConnected()) {
        box.replaceChildren(h('h3', null, 'Trabajo (Planner)'), intro,
          h('div', { class: 'row-actions' }, h('button', { type: 'button', class: 'btn primary small', onclick: () => { plConnect(); } }, 'Conectar con Planner')));
        return;
      }
      const plan = (S.pl.plans || []).find(p => p.id === S.pl.planId);
      box.replaceChildren(h('h3', null, 'Trabajo (Planner)'),
        h('p', { class: 'muted' }, 'Conectado como ', h('strong', null, S.pl.account), '.'),
        h('p', { class: 'muted' }, (plan ? 'Plan: ' + plan.title + '. ' : '') + (S.plState === 'login' ? 'La sesión ha caducado: pulsa Reconectar.' : S.pl.syncedAt ? 'Última sincronización: ' + relTime(S.pl.syncedAt) + '.' : 'Sin sincronizar todavía.')),
        h('div', { class: 'row-actions' },
          S.plState === 'login'
            ? h('button', { type: 'button', class: 'btn primary small', onclick: () => { plConnect(); } }, 'Reconectar')
            : h('button', { type: 'button', class: 'btn ghost small', onclick: () => plSync({ notify: true }) }, 'Sincronizar ahora'),
          h('button', {
            type: 'button', class: 'btn subtle small',
            onclick: async () => {
              if (!confirm('¿Desconectar Planner en este navegador? En Planner no se borra nada.')) return;
              await plDisconnect();
              draw();
            }
          }, 'Desconectar')));
    };
    draw();
    const timer = setInterval(() => { if (!box.isConnected) clearInterval(timer); else if (!(box.contains(document.activeElement) && document.activeElement.tagName === 'INPUT')) draw(); }, 1000);
    return box;
  }

  // Ajustes · Google
  function ggSection() {
    const box = h('div', { class: 'modal-section' });
    const draw = () => {
      const intro = h('p', { class: 'muted' }, 'Tus contactos (con sus teléfonos, correos, cumpleaños, notas y categorías) y tu calendario de Google.');
      if (!S.gg.clientId) {
        const input = h('input', { class: 'input', type: 'text', placeholder: 'ID de cliente de Google', autocomplete: 'off', spellcheck: 'false', 'aria-label': 'ID de cliente de Google' });
        const saveId = () => {
          const v = input.value.trim();
          if (!/^[\w-]+\.apps\.googleusercontent\.com$/.test(v)) { toast('El ID de cliente termina en .apps.googleusercontent.com', 'error'); input.focus(); return; }
          S.gg.clientId = v;
          saveGg();
          draw();
        };
        input.addEventListener('keydown', e => { if (e.key === 'Enter') saveId(); });
        box.replaceChildren(h('h3', null, 'Google'), intro,
          h('p', { class: 'muted' }, 'Pega el ID de cliente que creaste en Google Cloud (tipo «Aplicación web»).'),
          input, h('div', { class: 'row-actions' }, h('button', { type: 'button', class: 'btn ghost small', onclick: saveId }, 'Guardar')));
        return;
      }
      const changeId = h('button', {
        type: 'button', class: 'btn subtle small',
        onclick: () => {
          if (!confirm('¿Cambiar el ID de cliente de Google? Se desconectará Google en este navegador.')) return;
          ggDisconnect();
          S.gg.clientId = '';
          saveGg();
          draw();
        }
      }, 'Cambiar ID de cliente');
      if (!ggConnected()) {
        box.replaceChildren(h('h3', null, 'Google'), intro,
          h('div', { class: 'row-actions' }, h('button', { type: 'button', class: 'btn primary small', onclick: ggLogin }, 'Conectar con Google'), changeId));
        return;
      }
      const n = Object.keys(S.datos.colecciones.contactos || {}).length;
      const cal = calCfg();
      const calLine = h('p', { class: 'muted' }, 'Calendario: ', cal ? h('strong', null, cal.nombre) : 'sin elegir', '. ',
        !S.gg.cal ? h('button', { type: 'button', class: 'btn subtle small inline', onclick: () => ggLogin(true) }, 'Dar permiso de Calendar')
          : h('button', { type: 'button', class: 'btn subtle small inline', onclick: openCalPicker }, cal ? 'Cambiar' : 'Elegir calendario'));
      box.replaceChildren(h('h3', null, 'Google'),
        h('p', { class: 'muted' }, 'Conectado como ', h('strong', null, S.gg.account), '.'),
        h('p', { class: 'muted' }, n + ' contactos · ' + (S.gg.syncedAt ? 'leídos ' + relTime(S.gg.syncedAt) : 'sin leer todavía') + '. El permiso de Google dura una hora; después, «Actualizar contactos» vuelve a pedirlo.'),
        calLine,
        h('div', { class: 'row-actions' },
          h('button', { type: 'button', class: 'btn ghost small', onclick: ggRefresh }, 'Actualizar contactos'),
          h('button', {
            type: 'button', class: 'btn subtle small',
            onclick: () => {
              if (!confirm('¿Desconectar Google en este navegador? Los contactos ya leídos se quedan en la app.')) return;
              ggDisconnect();
              draw();
            }
          }, 'Desconectar')));
    };
    draw();
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
          // Copia del almacén: se sustituyen solo las colecciones que trae el fichero
          const doc = normalizeDoc(parsed.datos);
          if (doc.schema > SCHEMA) throw new Error('versión');
          const names = Object.keys(doc.colecciones);
          if (!names.length) throw new Error('vacío');
          const resumen = names.map(n => n + ' (' + Object.keys(doc.colecciones[n]).length + ')').join(', ');
          if (!confirm('Se sustituirán estos datos por los del fichero: ' + resumen + '. El resto no se toca. ¿Continuar?')) return;
          const merged = normalizeDoc(S.datos);
          for (const n of names) merged.colecciones[n] = doc.colecciones[n];
          applyDatos(merged);
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
      h('div', { class: 'modal-section' },
        h('h3', null, 'Etiquetas de tareas'),
        h('p', { class: 'muted' }, 'Crea, renombra, ordena y elige el color de las etiquetas que arrastras a las tareas.'),
        h('div', { class: 'row-actions' }, h('button', { type: 'button', class: 'btn ghost small', onclick: () => { close(); openTagManager(); } }, 'Gestionar etiquetas'))),
      odSection(),
      plSection(),
      ggSection(),
      h('div', { class: 'modal-section' },
        h('h3', null, 'Copia de seguridad'),
        h('p', { class: 'muted' }, 'Exporta todos los datos de la app en un fichero, o importa una copia. Al importar se sustituye solo lo que trae el fichero (por ejemplo, gastos y pagos); lo demás no se toca.'),
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
              if (S.gg.token) ggDisconnect();
              Object.values(KEYS).concat(LEGACY_TAGS_KEY).forEach(k => localStorage.removeItem(k));
              Object.keys(localStorage).filter(k => k.startsWith('msal.')).forEach(k => localStorage.removeItem(k));
              location.reload();
            }
          }, 'Borrar datos de este navegador'))));
    close = openModal('Ajustes', body);
  }

  // ---------- pantallas ----------
  // Pantallas propias: sección en index.html (#view<Nombre>), función que la pinta y acción del botón principal
  const VIEWS = {
    personal: { title: 'Personal', add: 'Añadir tarea', board: 'Todoist' },
    trabajo: { title: 'Trabajo', add: 'Añadir tarea', board: 'Planner' },
    calendario: { title: 'Calendario', eyebrow: () => calCfg() ? calCfg().nombre : 'Google Calendar', add: 'Hoy', addIcon: 'calendar', el: 'viewCalendario', render: () => renderCal(), onAdd: () => { calMonth = null; calSel = null; renderCal(); } },
    contactos: { title: 'Contactos', eyebrow: 'Agenda y cumples', add: () => ctTab() === 'agenda' ? 'Nuevo contacto' : 'Elegir contactos', el: 'viewContactos', render: () => renderCumpledias(), onAdd: () => ctTab() === 'agenda' ? openEditContacto(null) : openCdPicker() },
    gastos: { title: 'Gastos', eyebrow: 'Pagos recurrentes', add: 'Nuevo gasto', el: 'viewGastos', render: () => renderGastos(), onAdd: () => openGasto(null) },
    consumos: { title: 'Consumos', eyebrow: 'Luz, gas y agua', add: 'Subir facturas', el: 'viewConsumos', render: () => renderConsumos(), onAdd: () => pickFacturas() }
  };

  function applyView() {
    if (S.ui.view === 'cumpledias') S.ui.view = 'contactos'; // nombre anterior de la pantalla
    if (!VIEWS[S.ui.view]) S.ui.view = 'personal';
    const v = S.ui.view, cfg = VIEWS[v], board = !!cfg.board;
    if (board) document.getElementById('eyebrowSrc').textContent = cfg.board;
    document.querySelectorAll('.nav-item[data-view]').forEach(a => {
      const on = a.dataset.view === v;
      a.classList.toggle('active', on);
      if (on) a.setAttribute('aria-current', 'page'); else a.removeAttribute('aria-current');
    });
    document.getElementById('viewTitle').textContent = cfg.title;
    document.getElementById('eyebrowTodoist').hidden = !board;
    const eb = document.getElementById('eyebrowOther');
    eb.hidden = board;
    eb.textContent = (typeof cfg.eyebrow === 'function' ? cfg.eyebrow() : cfg.eyebrow) || '';
    document.getElementById('addLabel').textContent = typeof cfg.add === 'function' ? cfg.add() : cfg.add;
    const addBtn = document.getElementById('addBtn');
    addBtn.querySelector('svg').replaceWith(ICON[cfg.addIcon || 'plus']());
    document.getElementById('tagbar').hidden = !board;
    document.getElementById('boardWrap').hidden = !board;
    for (const [k, c] of Object.entries(VIEWS)) if (c.el) document.getElementById(c.el).hidden = k !== v;
    if (cfg.render) cfg.render();
  }

  function showView(v) {
    if (S.editing) closeDrawer();
    if (VIEWS[v] && VIEWS[v].board && v !== S.ui.view) S.filter.clear();
    S.ui.view = v;
    saveUi();
    if (SMALL.matches) toggleSidebar(false);
    if (VIEWS[v] && VIEWS[v].board) renderAll(); else applyView();
  }

  // ---------- pantalla Calendario (vista de mes) ----------
  let calMonth = null; // primer día del mes visible (ms UTC)
  let calSel = null;   // día elegido (AAAA-MM-DD)
  let calLoading = false;
  const DIAS_SEM = ['L', 'M', 'X', 'J', 'V', 'S', 'D'];
  const localYmd = d => d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');

  function evColor(e) {
    const k = e.key || '';
    if (k.startsWith('cumple:')) return '#E68AC4';
    if (k.startsWith('dias:')) return '#F2B35B';
    if (k.startsWith('tarea:')) return '#DBF266';
    if (k.startsWith('gasto:')) return '#F5D76E';
    const c = calCfg() && calCfg().color;
    return isHex(c) ? c : '#7FA2E0';
  }
  const evOrigin = e => (e.key || '').startsWith('cumple:') ? 'Cumpleaños' : (e.key || '').startsWith('dias:') ? 'Cumpledías' : (e.key || '').startsWith('gasto:') ? 'Gasto' : (e.key || '').startsWith('tarea:') ? 'Tarea' : '';
  // Día(s) de un evento: los de todo el día ocupan [inicio, fin); los de hora, su día local de inicio
  function evOnDay(e, ymd) {
    if (e.allDay) return e.start <= ymd && ymd < (e.end || addDays(e.start, 1));
    return localYmd(new Date(e.start)) === ymd;
  }
  const evTime = e => e.allDay ? 'Todo el día' : new Date(e.start).toLocaleTimeString('es-ES', { hour: '2-digit', minute: '2-digit' });

  function renderCal() {
    const root = document.getElementById('viewCalendario');
    root.dataset.day = String(todayUtc());
    const box = (msg, label, fn) => root.replaceChildren(h('section', { class: 'panel' }, h('p', { class: 'muted empty-note' }, msg),
      label ? h('div', { class: 'row-actions' }, h('button', { type: 'button', class: 'btn primary small', onclick: fn }, label)) : null));
    if (!ggConnected()) return box('Conecta tu cuenta de Google en Ajustes para ver aquí tu calendario.', 'Ir a Ajustes', openSettings);
    if (!S.gg.cal) return box('Falta el permiso de Google Calendar. Al pulsar, Google te pedirá aceptarlo (marca las casillas de Calendar).', 'Dar permiso', () => ggLogin(true));
    if (!calCfg()) return box('Elige en Ajustes el calendario que quieres usar.', 'Elegir calendario', openCalPicker);

    const today = ymdOf(todayUtc());
    if (calMonth === null) { const t = new Date(todayUtc()); calMonth = Date.UTC(t.getUTCFullYear(), t.getUTCMonth(), 1); }
    const m = new Date(calMonth), Y = m.getUTCFullYear(), M = m.getUTCMonth();
    const first = ymdOf(calMonth), next = ymdOf(Date.UTC(Y, M + 1, 1));
    if (!calSel || calSel.slice(0, 7) !== first.slice(0, 7)) calSel = today.slice(0, 7) === first.slice(0, 7) ? today : first;

    // Mes fuera de lo ya leído: se lee si el permiso sigue vigente
    if ((!S.cal.from || first < S.cal.from || next > S.cal.to) && ggTokenOk() && !calLoading) {
      calLoading = true;
      calFetch(ymdOf(Date.UTC(Y, M - 1, 1)), ymdOf(Date.UTC(Y, M + 2, 1)))
        .catch(e => toast(e.message || 'No se pudo leer el calendario.', 'error'))
        .finally(() => { calLoading = false; if (S.ui.view === 'calendario') renderCal(); });
    }

    const nav = h('div', { class: 'cal-nav' },
      h('button', { type: 'button', class: 'icon-btn', 'aria-label': 'Mes anterior', onclick: () => { calMonth = Date.UTC(Y, M - 1, 1); renderCal(); } }, ICON.left()),
      h('h2', null, cap(new Date(calMonth).toLocaleDateString('es-ES', { month: 'long', year: 'numeric', timeZone: 'UTC' }))),
      h('button', { type: 'button', class: 'icon-btn', 'aria-label': 'Mes siguiente', onclick: () => { calMonth = Date.UTC(Y, M + 1, 1); renderCal(); } }, ICON.right()),
      calLoading ? h('span', { class: 'muted' }, 'Leyendo…') : null,
      h('div', { class: 'legend cal-legend' },
        h('span', { class: 'lg', style: '--c:#E68AC4' }, 'Cumpleaños'), h('span', { class: 'lg', style: '--c:#F2B35B' }, 'Cumpledías'),
        h('span', { class: 'lg', style: '--c:#F5D76E' }, 'Gastos'), h('span', { class: 'lg', style: '--c:#DBF266' }, 'Tareas')));

    const grid = h('div', { class: 'cal-grid' }, DIAS_SEM.map(d => h('span', { class: 'cal-wd' }, d)));
    const offset = (m.getUTCDay() + 6) % 7; // lunes = 0
    const start = calMonth - offset * DAY;
    const weeks = Math.ceil((offset + new Date(Date.UTC(Y, M + 1, 0)).getUTCDate()) / 7);
    for (let i = 0; i < weeks * 7; i++) {
      const ymd = ymdOf(start + i * DAY);
      const evs = S.cal.events.filter(e => evOnDay(e, ymd));
      const cell = h('button', {
        type: 'button', class: 'cal-day' + (ymd.slice(0, 7) !== first.slice(0, 7) ? ' out' : '') + (ymd === today ? ' today' : '') + (ymd === calSel ? ' sel' : ''),
        'aria-label': fmtDay(utcDay(ymd)) + (evs.length ? ', ' + evs.length + (evs.length === 1 ? ' evento' : ' eventos') : ''),
        onclick: () => { calSel = ymd; if (ymd.slice(0, 7) !== first.slice(0, 7)) calMonth = Date.UTC(+ymd.slice(0, 4), +ymd.slice(5, 7) - 1, 1); renderCal(); }
      },
        h('span', { class: 'cal-num' }, String(+ymd.slice(8))),
        h('span', { class: 'cal-evs' }, evs.slice(0, 3).map(e => h('span', { class: 'cal-ev', style: '--c:' + evColor(e) }, e.title)),
          evs.length > 3 ? h('span', { class: 'cal-more' }, '+' + (evs.length - 3)) : null),
        h('span', { class: 'cal-dots' }, evs.slice(0, 4).map(e => h('span', { class: 'cal-dot', style: '--c:' + evColor(e) }))));
      grid.append(cell);
    }

    const dayEvs = S.cal.events.filter(e => evOnDay(e, calSel)).sort((a, b) => (a.allDay === b.allDay ? a.start.localeCompare(b.start) : a.allDay ? -1 : 1));
    const list = h('div', { class: 'ev-list' }, dayEvs.map(e => h('a', { class: 'ev cal-item', href: e.link || '#', target: '_blank', rel: 'noopener noreferrer' },
      h('span', { class: 'cat-dot', style: '--c:' + evColor(e) }),
      h('span', { class: 'ev-main' }, h('span', { class: 'ev-name' }, e.title), h('span', { class: 'ev-what' }, [evTime(e), evOrigin(e) ? 'creado por la app (' + evOrigin(e).toLowerCase() + ')' : ''].filter(Boolean).join(' · '))))));
    if (!dayEvs.length) list.append(h('p', { class: 'muted empty-note' }, 'Sin eventos este día.'));

    root.replaceChildren(
      h('section', { class: 'panel' }, nav, grid),
      h('section', { class: 'panel' }, h('div', { class: 'panel-head' }, h('h2', null, cap(new Date(utcDay(calSel)).toLocaleDateString('es-ES', { weekday: 'long', day: 'numeric', month: 'long', timeZone: 'UTC' })))), list));
  }

  // Elegir el calendario de Google (se guarda en el almacén: vale para todos los dispositivos)
  async function openCalPicker() {
    if (!S.gg.cal) { ggLogin(true); return; }
    if (!ggTokenOk()) { toast('El permiso de Google ha caducado: reconecta y vuelve a elegir.', 'error', { label: 'Conectar', fn: ggLogin }, 10000); return; }
    let cals;
    try { cals = await calList(); } catch (e) { toast(e.message, 'error'); return; }
    if (!cals.length) { toast('No hay calendarios en los que puedas escribir.', 'error'); return; }
    const cur = calCfg();
    let close = null;
    const list = h('div', { class: 'pick-list' }, cals.map(c => h('button', {
      type: 'button', class: 'pick' + (cur && cur.calendarId === c.id ? ' on' : ''),
      onclick: () => {
        S.datos.colecciones.ajustes = Object.assign({}, S.datos.colecciones.ajustes, { calendario: { id: 'calendario', calendarId: c.id, nombre: c.nombre, color: isHex(c.color) ? c.color : null } });
        S.cal = { events: [], from: null, to: null, syncedAt: null };
        saveCal();
        saveDatos();
        close();
        toast('Calendario: ' + c.nombre);
        applyView();
        calSync({ notify: true });
      }
    }, h('span', { class: 'pick-box', style: isHex(c.color) ? 'background:' + c.color + ';border-color:' + c.color : '' }), h('span', { class: 'ev-main' }, h('span', { class: 'ev-name' }, c.nombre), h('span', { class: 'ev-what' }, c.primary ? 'Calendario principal' : '')))));
    close = openModal('Elige el calendario', h('div', { class: 'form' },
      h('p', { class: 'hint' }, 'La app leerá sus eventos y creará en él los suyos: cumpleaños y cumpledías especiales de los contactos que sigues, los gastos marcados y las tareas con fecha.'), list));
  }

  // ---------- contactos: cumpledías y cumpleaños ----------
  // Colección "personas": { id, name, birth: 'AAAA-MM-DD' }. Pensada para crecer
  // (p. ej. un vínculo con contactos de Google más adelante).
  const DAY = 86400000;
  const MILESTONE = 1000;
  let cdFilter = 'todo'; // todo | dias | cumple

  const personas = () => Object.values(S.datos.colecciones.personas || {});
  const utcDay = ymd => { const [y, m, d] = ymd.split('-').map(Number); return Date.UTC(y, m - 1, d); };
  const todayUtc = () => { const n = new Date(); return Date.UTC(n.getFullYear(), n.getMonth(), n.getDate()); };
  const isLeap = y => (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;
  const fmtNum = n => n.toLocaleString('es-ES');
  const fmtDay = t => new Date(t).toLocaleDateString('es-ES', { day: 'numeric', month: 'short', year: 'numeric', timeZone: 'UTC' });
  const isYmd = v => /^\d{4}-\d{2}-\d{2}$/.test(v || '') && !isNaN(utcDay(v));

  function untilText(n) {
    if (n === 0) return 'hoy';
    if (n === 1) return 'mañana';
    if (n < 60) return 'en ' + n + ' días';
    if (n < 730) return 'en ' + Math.round(n / 30.44) + ' meses';
    return 'en ' + (n / 365.25).toFixed(1).replace('.', ',') + ' años';
  }

  // Próximos hitos de una persona (desde hoy, incluido): tres de días de vida y su próximo cumpleaños
  function personEvents(p, today) {
    if (/^--\d{2}-\d{2}$/.test(p.birth || '')) {
      // Sin año (contactos de Google): solo el cumpleaños, sin edad ni días de vida
      const bm = +p.birth.slice(2, 4) - 1, bd = +p.birth.slice(5, 7);
      const onYear = y => Date.UTC(y, bm, bm === 1 && bd === 29 && !isLeap(y) ? 28 : bd);
      let y = new Date(today).getUTCFullYear();
      if (onYear(y) < today) y++;
      return [{ kind: 'cumple', p, value: null, date: onYear(y), until: Math.round((onYear(y) - today) / DAY) }];
    }
    const birth = utcDay(p.birth);
    const lived = Math.round((today - birth) / DAY);
    const out = [];
    let k = Math.max(MILESTONE, Math.ceil(lived / MILESTONE) * MILESTONE);
    for (let i = 0; i < 3; i++, k += MILESTONE) out.push({ kind: 'dias', p, value: k, date: birth + k * DAY });
    const b = new Date(birth), bm = b.getUTCMonth(), bd = b.getUTCDate();
    const onYear = y => Date.UTC(y, bm, bm === 1 && bd === 29 && !isLeap(y) ? 28 : bd); // 29-F: el 28 en años no bisiestos
    let y = new Date(today).getUTCFullYear();
    if (onYear(y) < today) y++;
    out.push({ kind: 'cumple', p, value: y - b.getUTCFullYear(), date: onYear(y) });
    return out.map(e => Object.assign(e, { until: Math.round((e.date - today) / DAY) }));
  }

  // Pantalla Contactos: pestañas Agenda (contactos de Google) y Cumpledías
  const contactos = () => Object.values(S.datos.colecciones.contactos || {});
  const ctTab = () => S.ui.ctTab || (contactos().length ? 'agenda' : 'cumpledias');
  const normName = s => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/\s+/g, ' ').trim();
  // Color de cada categoría de Google: el elegido por el usuario (colección "coloresContactos",
  // {id: resourceName del grupo, color}) o, si no hay, uno estable sacado del nombre
  function catColorC(name) {
    const rn = Object.entries(grupoMap()).find(([, n]) => n === name);
    const elegido = rn && ((S.datos.colecciones.coloresContactos || {})[rn[0]] || {}).color;
    if (isHex(elegido)) return elegido;
    let hsh = 0;
    for (const ch of String(name)) hsh = (hsh * 31 + ch.codePointAt(0)) >>> 0;
    return PALETTE[hsh % PALETTE.length];
  }

  function openCatColors() {
    const grupos = Object.entries(grupoMap()).sort((a, b) => a[1].localeCompare(b[1], 'es'));
    if (!grupos.length) { toast('No hay categorías en tus contactos.', 'error'); return; }
    const list = h('div', { class: 'cat-colors' });
    let open = null;
    const draw = () => list.replaceChildren(...grupos.map(([rn, name]) => {
      const color = catColorC(name);
      const row = h('div', { class: 'tm-row' },
        h('div', { class: 'tm-main' },
          h('button', { type: 'button', class: 'swatch', style: '--c:' + color, 'aria-label': 'Cambiar color de ' + name, onclick: () => { open = open === rn ? null : rn; draw(); } }),
          h('span', { class: 'cat-name' }, name)),
        open === rn ? colorPicker(color, c => {
          S.datos.colecciones.coloresContactos = Object.assign({}, S.datos.colecciones.coloresContactos, { [rn]: { id: rn, color: c } });
          saveDatos();
          open = null;
          draw();
        }) : null);
      return row;
    }));
    draw();
    openModal('Colores de las categorías', h('div', { class: 'form' },
      h('p', { class: 'hint' }, 'Solo cambia cómo se ven en la app; en Google las categorías no tienen color.'), list),
      { onClose: () => { if (S.ui.view === 'contactos') renderCumpledias(); } });
  }

  // Paleta de colores con opción de color libre
  function colorPicker(current, onPick) {
    const box = h('div', { class: 'palette' });
    for (const c of PALETTE) {
      box.append(h('button', { type: 'button', class: c.toLowerCase() === String(current).toLowerCase() ? 'sel' : '', style: '--c:' + c, 'aria-label': 'Color ' + c, onclick: () => onPick(c) }));
    }
    const custom = h('input', { type: 'color', value: isHex(current) ? current : PALETTE[0], 'aria-label': 'Color personalizado' });
    custom.addEventListener('change', () => { if (isHex(custom.value)) onPick(custom.value); });
    box.append(h('label', null, 'Otro', custom));
    return box;
  }
  // Contactos marcados para Cumpledías: colección "cumpledias" ({id: id del contacto})
  const cdMarked = id => !!(S.datos.colecciones.cumpledias || {})[id];
  function setCdMark(id, on) {
    const marks = Object.assign({}, S.datos.colecciones.cumpledias);
    if (on) marks[id] = { id }; else delete marks[id];
    S.datos.colecciones.cumpledias = marks;
    saveDatos();
  }
  let agQuery = '';
  let agCat = null; // null | '★' | nombre de categoría

  // Personas de Cumpledías: los contactos marcados (con cumpleaños) y, si quedan, las añadidas a mano antes
  function cumplePeople() {
    const google = contactos().filter(c => cdMarked(c.id) && c.cumple).map(c => ({ id: c.id, name: c.nombre, birth: c.cumple, src: 'google' }));
    const ya = new Set(google.map(p => normName(p.name)));
    const manual = personas().filter(p => isYmd(p.birth) && !ya.has(normName(p.name))).map(p => ({ id: p.id, name: p.name, birth: p.birth, src: 'manual' }));
    return google.concat(manual).sort((a, b) => a.name.localeCompare(b.name, 'es'));
  }

  // Elegir qué contactos salen en Cumpledías (solo los que tienen cumpleaños)
  function openCdPicker() {
    const all = contactos().filter(c => c.cumple).sort((a, b) => a.nombre.localeCompare(b.nombre, 'es'));
    if (!all.length) { toast(contactos().length ? 'Ningún contacto tiene cumpleaños. Añádeselo al editarlo.' : 'Primero conecta Google y lee tus contactos.', 'error'); return; }
    let q = '';
    const search = h('input', { class: 'input search', type: 'search', placeholder: 'Buscar contacto', 'aria-label': 'Buscar contacto', autocomplete: 'off' });
    const list = h('div', { class: 'pick-list' });
    const draw = () => {
      const nq = normName(q);
      list.replaceChildren(...all.filter(c => !nq || normName(c.nombre).includes(nq)).map(c => {
        const on = cdMarked(c.id);
        return h('button', { type: 'button', class: 'pick' + (on ? ' on' : ''), 'aria-pressed': on ? 'true' : 'false', onclick: () => { setCdMark(c.id, !on); draw(); } },
          h('span', { class: 'pick-box' }, on ? ICON.check() : null),
          h('span', { class: 'ev-main' }, h('span', { class: 'ev-name' }, c.nombre), h('span', { class: 'ev-what' }, cumpleText(c.cumple))));
      }));
    };
    search.addEventListener('input', () => { q = search.value; draw(); });
    draw();
    openModal('Contactos en Cumples', h('div', { class: 'form' },
      h('p', { class: 'hint' }, 'Marca los contactos cuyos cumpleaños y días de vida quieres seguir. Solo aparecen los que tienen cumpleaños en Google. También puedes marcarlos con el pastel de la Agenda.'),
      search, list), { onClose: () => renderCumpledias() });
    setTimeout(() => search.focus(), 50);
  }

  function cumpleText(cumple) {
    if (!cumple) return 'Sin cumpleaños';
    if (/^\d{4}/.test(cumple)) return fmtDay(utcDay(cumple));
    const [mm, dd] = cumple.slice(2).split('-');
    return (+dd) + ' de ' + MESES_LARGOS_CT[+mm - 1] + ' (sin año)';
  }

  function renderCumpledias() {
    const root = document.getElementById('viewContactos');
    const today = todayUtc();
    root.dataset.day = String(today);
    const tab = ctTab();
    const tabs = h('div', { class: 'seg', role: 'group', 'aria-label': 'Vista' },
      [['agenda', 'Agenda'], ['cumpledias', 'Cumples']].map(([k, label]) =>
        h('button', { type: 'button', class: tab === k ? 'on' : '', 'aria-pressed': tab === k ? 'true' : 'false', onclick: () => { S.ui.ctTab = k; saveUi(); applyView(); } }, label)));
    const head = h('div', { class: 'g-toolbar' }, tabs);
    root.replaceChildren(head, ...(tab === 'agenda' ? agendaNodes() : cumpleNodes(today)));
  }

  function agendaNodes() {
    const all = contactos().sort((a, b) => a.nombre.localeCompare(b.nombre, 'es'));
    if (!all.length) {
      return [h('section', { class: 'panel' },
        h('p', { class: 'muted empty-note' }, S.gg.clientId
          ? 'Aún no has leído tus contactos de Google.'
          : 'Conecta tu cuenta de Google en Ajustes para ver aquí tus contactos (teléfonos, correos, cumpleaños, notas y categorías).'),
        h('div', { class: 'row-actions' }, S.gg.clientId
          ? h('button', { type: 'button', class: 'btn primary small', onclick: ggRefresh }, ggConnected() ? 'Leer contactos' : 'Conectar con Google')
          : h('button', { type: 'button', class: 'btn ghost small', onclick: openSettings }, 'Ir a Ajustes')))];
    }
    const cats = [...new Set(all.flatMap(c => c.etiquetas || []))].sort((a, b) => a.localeCompare(b, 'es'));
    if (agCat && agCat !== '★' && agCat !== 'cumples' && !cats.includes(agCat)) agCat = null;
    const search = h('input', { class: 'input search', type: 'search', placeholder: 'Buscar por nombre, teléfono o correo', value: agQuery, 'aria-label': 'Buscar contactos', autocomplete: 'off' });
    const chip = (key, label, color) => h('button', {
      type: 'button', class: 'chip' + (agCat === key ? ' on' : ''), style: '--c:' + color, 'aria-pressed': agCat === key ? 'true' : 'false',
      onclick: () => { agCat = agCat === key ? null : key; renderCumpledias(); }
    }, key === null ? null : h('span', { class: 'dot' }), label);
    const chips = h('div', { class: 'chips' },
      chip(null, 'Todos', '#C9D0DC'),
      all.some(c => c.favorito) ? chip('★', 'Favoritos', '#DBF266') : null,
      all.some(c => cdMarked(c.id)) ? chip('cumples', 'En Cumples', '#E68AC4') : null,
      cats.map(c => chip(c, c, catColorC(c))),
      cats.length ? h('button', { type: 'button', class: 'btn subtle small', onclick: openCatColors, title: 'Elegir los colores de las categorías' }, 'Colores') : null);
    const count = h('span', { class: 'panel-total' });
    const list = h('div', { class: 'ev-list' });
    const draw = () => {
      const q = normName(agQuery), qd = agQuery.replace(/\D/g, '');
      const shown = all.filter(c => (agCat === null || (agCat === '★' ? c.favorito : agCat === 'cumples' ? cdMarked(c.id) : (c.etiquetas || []).includes(agCat))) &&
        (!q || normName(c.nombre).includes(q) || (c.emails || []).some(e => normName(e.valor).includes(q)) || (qd.length >= 3 && (c.telefonos || []).some(t => t.valor.replace(/\D/g, '').includes(qd)))));
      count.textContent = shown.length + (shown.length === 1 ? ' contacto' : ' contactos');
      list.replaceChildren(...shown.slice(0, 400).map(contactRow));
      if (!shown.length) list.append(h('p', { class: 'muted empty-note' }, 'Ningún contacto coincide.'));
    };
    search.addEventListener('input', () => { agQuery = search.value; draw(); });
    draw();
    const refresh = h('button', { type: 'button', class: 'btn ghost small', onclick: ggRefresh, title: S.gg.syncedAt ? 'Leídos ' + relTime(S.gg.syncedAt) : '' }, 'Actualizar desde Google');
    return [h('section', { class: 'panel' },
      h('div', { class: 'panel-head' }, search, count, refresh), chips, list)];
  }

  // Fila de la Agenda: pulsar abre la ficha; el pastel marca o desmarca para Cumples sin abrirla
  function contactRow(c) {
    const color = (c.etiquetas || []).length ? catColorC(c.etiquetas[0]) : '#9AA4B5';
    const ini = c.nombre.split(/\s+/).filter(Boolean).slice(0, 2).map(w => w[0]).join('').toUpperCase();
    const sub = [];
    if ((c.telefonos || [])[0]) sub.push(c.telefonos[0].valor);
    if ((c.emails || [])[0]) sub.push(c.emails[0].valor);
    const cake = h('button', { type: 'button', class: 'icon-btn mini cake-toggle' });
    const paint = () => {
      const on = cdMarked(c.id);
      cake.classList.toggle('on', on);
      cake.setAttribute('aria-pressed', on ? 'true' : 'false');
      cake.setAttribute('aria-label', on ? 'Quitar de Cumples' : 'Seguir en Cumples');
      cake.title = on ? 'En Cumples (pulsa para quitar)' : c.cumple ? 'Seguir en Cumples' : 'Sin cumpleaños: añádeselo con Editar';
    };
    cake.append(ICON.cake());
    cake.addEventListener('click', e => {
      e.stopPropagation();
      const on = !cdMarked(c.id);
      setCdMark(c.id, on);
      paint();
      if (on && !c.cumple) toast('Sin cumpleaños no saldrá en Cumples: añádeselo con Editar.', 'error');
    });
    paint();
    return h('div', { class: 'ev contact', onclick: () => openContacto(c.id) },
      h('span', { class: 'avatar', style: '--c:' + color }, ini || '·'),
      h('button', { type: 'button', class: 'ev-main as-btn' },
        h('span', { class: 'ev-name' }, c.nombre, c.favorito ? h('span', { class: 'star', title: 'Favorito' }, ' ★') : null),
        h('span', { class: 'ev-what' }, sub.join(' · ') || 'Sin teléfono ni correo')),
      h('span', { class: 'cat-dots' }, (c.etiquetas || []).slice(0, 4).map(e => h('span', { class: 'cat-dot', style: '--c:' + catColorC(e), title: e }))),
      cake);
  }

  // Ficha de contacto (solo lectura; se edita en Google Contactos)
  function openContacto(id) {
    const c = (S.datos.colecciones.contactos || {})[id];
    if (!c) return;
    const rows = [];
    const sect = (label, items) => { if (items.length) rows.push(h('div', { class: 'field' }, h('span', { class: 'field-label' }, label), ...items)); };
    const line = (main, tipo, href) => h('div', { class: 'ct-line' },
      href ? h('a', { href, class: 'ct-link' }, main) : h('span', null, main),
      tipo ? h('span', { class: 'ct-type' }, tipo) : null);
    sect('Teléfonos', (c.telefonos || []).map(t => line(t.valor, t.tipo, 'tel:' + t.valor.replace(/[^\d+]/g, ''))));
    sect('Correos', (c.emails || []).map(e => line(e.valor, e.tipo, /^[^\s@]+@[^\s@]+$/.test(e.valor) ? 'mailto:' + e.valor : null)));
    if (c.cumple) {
      const full = /^\d{4}/.test(c.cumple);
      const [, mm, dd] = c.cumple.replace(/^--/, '0000-').split('-');
      const txt = full ? fmtDay(utcDay(c.cumple)) : (+dd) + ' de ' + MESES_LARGOS_CT[+mm - 1];
      sect('Cumpleaños', [line(txt, full ? Math.floor((todayUtc() - utcDay(c.cumple)) / DAY / 365.2425) + ' años' : '', null)]);
    }
    if ((c.etiquetas || []).length) sect('Categorías', [h('div', { class: 'chips' }, c.etiquetas.map(e => h('span', { class: 'chip on static', style: '--c:' + catColorC(e) }, h('span', { class: 'dot' }), e)))]);
    if (c.empresa) sect('Empresa', [line(c.empresa, '', null)]);
    sect('Direcciones', (c.direcciones || []).map(d => line(d.valor, d.tipo, null)));
    if (c.notas) sect('Notas', [h('div', { class: 'note-text' }, c.notas)]);
    const gid = String(c.resourceName || '').replace(/^people\//, '');
    const mark = h('input', { type: 'checkbox', id: 'ctCd', checked: cdMarked(c.id) });
    mark.addEventListener('change', () => { setCdMark(c.id, mark.checked); if (mark.checked && !c.cumple) toast('Sin cumpleaños no saldrá en Cumples: añádeselo con Editar.', 'error'); });
    let close = null;
    const body = h('div', { class: 'form' }, ...rows,
      rows.length ? null : h('p', { class: 'muted' }, 'Este contacto no tiene más datos.'),
      h('label', { class: 'check-line', for: 'ctCd' }, mark, 'Seguir en Cumples'),
      h('div', { class: 'row-actions modal-foot' },
        h('a', { class: 'btn subtle small', href: 'https://contacts.google.com/person/' + encodeURIComponent(gid), target: '_blank', rel: 'noopener noreferrer' }, 'Abrir en Google'),
        h('span', { class: 'spacer' }),
        h('button', { type: 'button', class: 'btn primary small', onclick: () => { close(); openEditContacto(c.id); } }, ICON.edit(), 'Editar')));
    close = openModal(c.nombre + (c.favorito ? ' ★' : ''), body, { onClose: () => { if (S.ui.view === 'contactos') renderCumpledias(); } });
  }
  // Formulario de contacto (crear o editar): se guarda en Google
  const TIPOS_TEL = [['mobile', 'Móvil'], ['home', 'Casa'], ['work', 'Trabajo'], ['other', 'Otro']];
  const TIPOS_MAIL = [['home', 'Personal'], ['work', 'Trabajo'], ['other', 'Otro']];
  const TIPOS_DIR = [['home', 'Casa'], ['work', 'Trabajo'], ['other', 'Otro']];

  async function openEditContacto(id, draft) {
    let cur = id ? (S.datos.colecciones.contactos || {})[id] : null;
    if (id && !cur) { toast('Ese contacto ya no está.', 'error'); return; }
    if (!ggEnsureWrite({ edit: id || null, draft: draft || null })) return;
    if (cur && !cur.etag) {
      // Datos guardados por una versión anterior de la app: se releen antes de editar
      await syncContactos();
      cur = (S.datos.colecciones.contactos || {})[id];
      if (!cur || !cur.etag) { toast('No se pudo preparar el contacto para editarlo. Pulsa «Actualizar desde Google».', 'error'); return; }
    }
    const parts = cur && cur.cumple ? cur.cumple.replace(/^--/, '0-').split('-').map(Number) : null;
    const f = draft || {
      nombrePila: cur ? cur.nombrePila || (cur.apellidos ? '' : cur.nombre) : '', apellidos: cur ? cur.apellidos || '' : '',
      telefonos: cur ? cur.telefonos.map(x => ({ valor: x.valor, tipo: x.tipo, etiqueta: x.etiqueta })) : [{ valor: '', tipo: 'mobile' }],
      emails: cur ? cur.emails.map(x => ({ valor: x.valor, tipo: x.tipo, etiqueta: x.etiqueta })) : [{ valor: '', tipo: 'home' }],
      cumple: parts ? { anio: parts[0] || null, mes: parts[1], dia: parts[2] } : null,
      notas: cur ? cur.notas || '' : '',
      org: cur ? Object.assign({ nombre: '', cargo: '' }, cur.org) : { nombre: '', cargo: '' },
      direcciones: cur ? (cur.direcciones || []).map(x => ({ valor: x.valor, tipo: x.tipo, etiqueta: x.etiqueta })) : [],
      categorias: cur ? (cur.etiquetas || []).slice() : [],
      favorito: cur ? !!cur.favorito : false
    };

    const inp = (attrs, val, onInput) => {
      const el = h('input', Object.assign({ class: 'input', autocomplete: 'off' }, attrs));
      el.value = val || '';
      el.addEventListener('input', () => onInput(el.value));
      return el;
    };
    // Lista editable (teléfonos, correos, direcciones): valor + tipo + quitar
    const rowsEditor = (arr, tipos, attrs, addLabel, defTipo) => {
      const box = h('div', { class: 'rows-edit' });
      const draw = () => {
        box.replaceChildren(...arr.map((it, i) => {
          const opts = tipos.slice();
          if (it.tipo && !opts.some(o => o[0] === it.tipo)) opts.push([it.tipo, it.etiqueta || it.tipo]);
          const sel = h('select', { class: 'select', 'aria-label': 'Tipo' }, opts.map(([v, l]) => h('option', { value: v, selected: v === (it.tipo || defTipo) }, l)));
          sel.addEventListener('change', () => { it.tipo = sel.value; });
          if (!it.tipo) it.tipo = defTipo;
          return h('div', { class: 'row-edit' }, inp(attrs, it.valor, v => { it.valor = v; }), sel,
            h('button', { type: 'button', class: 'icon-btn mini', 'aria-label': 'Quitar', onclick: () => { arr.splice(i, 1); draw(); } }, ICON.close()));
        }), h('button', { type: 'button', class: 'btn subtle small add-row', onclick: () => { arr.push({ valor: '', tipo: defTipo }); draw(); } }, ICON.plus(), addLabel));
      };
      draw();
      return box;
    };

    const dia = inp({ type: 'text', inputmode: 'numeric', placeholder: 'Día', 'aria-label': 'Día', maxlength: '2' }, f.cumple ? String(f.cumple.dia) : '', () => {});
    const mes = h('select', { class: 'select', 'aria-label': 'Mes' }, h('option', { value: '' }, 'Mes'), MESES_LARGOS_CT.map((m, i) => h('option', { value: String(i + 1), selected: f.cumple && f.cumple.mes === i + 1 }, m)));
    const anio = inp({ type: 'text', inputmode: 'numeric', placeholder: 'Año', 'aria-label': 'Año', maxlength: '4' }, f.cumple && f.cumple.anio ? String(f.cumple.anio) : '', () => {});

    const cats = [...new Set(Object.values(grupoMap()).concat(f.categorias))].sort((a, b) => a.localeCompare(b, 'es'));
    const catBox = h('div', { class: 'chips' });
    const drawCats = () => catBox.replaceChildren(...cats.map(c => {
      const on = f.categorias.includes(c);
      return h('button', { type: 'button', class: 'chip' + (on ? ' on' : ''), style: '--c:' + catColorC(c), 'aria-pressed': on ? 'true' : 'false',
        onclick: () => { f.categorias = on ? f.categorias.filter(x => x !== c) : f.categorias.concat(c); drawCats(); } }, h('span', { class: 'dot' }), c);
    }));
    drawCats();
    const nuevaCat = inp({ type: 'text', placeholder: 'Nueva categoría', 'aria-label': 'Nueva categoría', maxlength: '60' }, '', () => {});
    const addCat = () => {
      const v = nuevaCat.value.trim();
      if (!v) return;
      if (!cats.includes(v)) cats.push(v);
      if (!f.categorias.includes(v)) f.categorias.push(v);
      nuevaCat.value = '';
      drawCats();
    };
    nuevaCat.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); addCat(); } });
    const fav = h('input', { type: 'checkbox', id: 'ceFav', checked: f.favorito });
    fav.addEventListener('change', () => { f.favorito = fav.checked; });
    const notas = h('textarea', { class: 'input textarea', rows: '3', 'aria-label': 'Notas' });
    notas.value = f.notas;
    notas.addEventListener('input', () => { f.notas = notas.value; });

    let close = null;
    const saveBtn = h('button', { type: 'button', class: 'btn primary small' }, cur ? 'Guardar' : 'Crear contacto');
    const readCumple = () => {
      const d = dia.value.trim(), m = mes.value, y = anio.value.trim();
      if (!d && !m && !y) return { ok: true, val: null };
      const dn = +d, mn = +m, yn = y ? +y : null;
      const max = new Date(Date.UTC(yn || 2000, mn, 0)).getUTCDate();
      if (!(mn >= 1 && mn <= 12) || !(dn >= 1 && dn <= max) || (yn !== null && !(yn >= 1900 && yn <= new Date().getFullYear()))) return { ok: false };
      return { ok: true, val: { dia: dn, mes: mn, anio: yn } };
    };
    const save = async () => {
      const bd = readCumple();
      if (!bd.ok) { toast('Revisa la fecha de cumpleaños (día y mes; el año es opcional).', 'error'); return; }
      f.cumple = bd.val;
      if (!f.nombrePila.trim() && !f.apellidos.trim()) { toast('Escribe al menos el nombre.', 'error'); return; }
      saveBtn.disabled = true;
      saveBtn.textContent = 'Guardando…';
      try {
        await saveContacto(cur, f);
        close();
        toast(cur ? 'Contacto guardado en Google' : 'Contacto creado en Google');
        if (S.ui.view === 'contactos') renderCumpledias();
      } catch (e) {
        saveBtn.disabled = false;
        saveBtn.textContent = cur ? 'Guardar' : 'Crear contacto';
        if (e.kind === 'login' || e.kind === 'scope') {
          // Se guarda lo escrito y se vuelve aquí después de pasar por Google
          sessionStorage.setItem(GG_NEXT, JSON.stringify({ edit: id || null, draft: f }));
          ggLogin(e.kind === 'scope');
          return;
        }
        if (e.kind === 'conflict') {
          close();
          const fresh = await reloadContacto(cur);
          toast('Este contacto cambió en Google mientras lo editabas: se ha recargado. Vuelve a hacer el cambio.', 'error', null, 10000);
          if (fresh) openEditContacto(fresh.id);
          return;
        }
        toast(e.message || 'No se pudo guardar en Google.', 'error', null, 10000);
      }
    };
    saveBtn.addEventListener('click', save);

    const fld = (label, ...els) => h('div', { class: 'field' }, h('span', { class: 'field-label' }, label), ...els);
    const body = h('div', { class: 'form' },
      h('div', { class: 'form-row' },
        fld('Nombre', inp({ type: 'text', maxlength: '80', 'aria-label': 'Nombre' }, f.nombrePila, v => { f.nombrePila = v; })),
        fld('Apellidos', inp({ type: 'text', maxlength: '80', 'aria-label': 'Apellidos' }, f.apellidos, v => { f.apellidos = v; }))),
      fld('Teléfonos', rowsEditor(f.telefonos, TIPOS_TEL, { type: 'tel', inputmode: 'tel', placeholder: 'Teléfono', 'aria-label': 'Teléfono' }, 'Añadir teléfono', 'mobile')),
      fld('Correos', rowsEditor(f.emails, TIPOS_MAIL, { type: 'email', inputmode: 'email', placeholder: 'Correo', 'aria-label': 'Correo' }, 'Añadir correo', 'home')),
      fld('Cumpleaños', h('div', { class: 'date-parts' }, dia, mes, anio), h('p', { class: 'hint' }, 'Día y mes; el año es opcional.')),
      fld('Categorías', catBox, h('div', { class: 'row-edit two' }, nuevaCat, h('button', { type: 'button', class: 'btn ghost small', onclick: addCat }, 'Añadir'))),
      h('label', { class: 'check-line', for: 'ceFav' }, fav, 'Favorito (destacado en Google)'),
      h('div', { class: 'form-row' },
        fld('Empresa', inp({ type: 'text', maxlength: '120', 'aria-label': 'Empresa' }, f.org.nombre, v => { f.org.nombre = v; })),
        fld('Cargo', inp({ type: 'text', maxlength: '120', 'aria-label': 'Cargo' }, f.org.cargo, v => { f.org.cargo = v; }))),
      fld('Direcciones', rowsEditor(f.direcciones, TIPOS_DIR, { type: 'text', placeholder: 'Dirección', 'aria-label': 'Dirección' }, 'Añadir dirección', 'home')),
      fld('Notas', notas),
      h('div', { class: 'row-actions modal-foot' },
        cur ? h('button', {
          type: 'button', class: 'btn danger small',
          onclick: async () => {
            if (!confirm('¿Borrar a ' + cur.nombre + ' de Google? Se puede recuperar en Google Contactos con «Deshacer cambios» (hasta 30 días).')) return;
            try { await deleteContacto(cur); close(); toast('Contacto borrado de Google'); renderCumpledias(); }
            catch (e) {
              if (e.kind === 'login' || e.kind === 'scope') { sessionStorage.setItem(GG_NEXT, JSON.stringify({ edit: id })); ggLogin(e.kind === 'scope'); return; }
              toast(e.message || 'No se pudo borrar.', 'error');
            }
          }
        }, 'Borrar') : null,
        h('span', { class: 'spacer' }),
        h('button', { type: 'button', class: 'btn subtle small', onclick: () => close() }, 'Cancelar'),
        saveBtn));
    close = openModal(cur ? 'Editar contacto' : 'Nuevo contacto', body);
  }

  const MESES_LARGOS_CT = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];

  function cumpleNodes(today) {
    const people = cumplePeople();
    const all = people.flatMap(p => personEvents(p, today)).sort((a, b) => a.date - b.date || a.p.name.localeCompare(b.p.name, 'es'));
    const events = all.filter(e => cdFilter === 'todo' || e.kind === cdFilter);

    const t = new Date(today), ty = t.getUTCFullYear(), tm = t.getUTCMonth();
    const inMonth = events.filter(e => { const d = new Date(e.date); return d.getUTCFullYear() === ty && d.getUTCMonth() === tm; }).length;
    const inYear = events.filter(e => new Date(e.date).getUTCFullYear() === ty).length;
    const next = events[0];
    const stat = (label, value, sub, cls) => h('div', { class: 'stat' + (cls ? ' ' + cls : '') },
      h('span', { class: 'stat-label' }, label), h('span', { class: 'stat-value' }, value), h('span', { class: 'stat-sub' }, sub));

    const stats = h('div', { class: 'stats' },
      stat('Próximo', next ? (next.until === 0 ? 'Hoy' : fmtNum(next.until)) : '—', next ? (next.until === 0 ? next.p.name : (next.until === 1 ? 'día · ' : 'días · ') + next.p.name) : 'sin personas', 'accent'),
      stat('Este mes', String(inMonth), inMonth === 1 ? 'evento' : 'eventos'),
      stat('Este año', String(inYear), inYear === 1 ? 'evento' : 'eventos'),
      stat('Personas', String(people.length), people.length === 1 ? 'en la lista' : 'en la lista'));

    const filters = h('div', { class: 'seg', role: 'group', 'aria-label': 'Filtrar' },
      [['todo', 'Todo'], ['cumple', 'Cumpleaños'], ['dias', 'Días de vida']].map(([k, label]) =>
        h('button', { type: 'button', class: cdFilter === k ? 'on' : '', 'aria-pressed': cdFilter === k ? 'true' : 'false', onclick: () => { cdFilter = k; renderCumpledias(); } }, label)));

    const list = h('div', { class: 'ev-list' });
    for (const e of events.slice(0, 20)) {
      const redondo = e.kind === 'dias' && e.value % 5000 === 0;
      const badge = e.kind === 'dias' ? fmtNum(e.value / 1000) + 'k' : e.value === null ? '' : String(e.value);
      const what = e.kind === 'dias' ? 'Cumple ' + fmtNum(e.value) + ' días de vida' : e.value === null ? 'Cumpleaños' : 'Cumple ' + e.value + (e.value === 1 ? ' año' : ' años');
      const kind = e.kind === 'cumple' ? 'cumple' : redondo ? 'redondo' : 'dias';
      list.append(h('div', { class: 'ev cd-' + kind + (e.until === 0 ? ' today' : '') },
        h('span', { class: 'ev-badge ' + kind, title: kind === 'cumple' ? 'Cumpleaños' : kind === 'redondo' ? 'Cumpledías redondo (múltiplo de 5.000)' : 'Cumpledías' }, badge || ICON.cake()),
        h('div', { class: 'ev-main' }, h('span', { class: 'ev-name' }, e.p.name), h('span', { class: 'ev-what' }, what)),
        h('div', { class: 'ev-when' },
          h('span', { class: 'ev-date' }, fmtDay(e.date)),
          h('span', { class: 'ev-until' + (e.until === 0 ? ' now' : e.until <= 30 ? ' soon' : '') }, untilText(e.until)))));
    }
    if (!events.length) list.append(h('p', { class: 'muted empty-note' }, people.length ? 'No hay eventos con este filtro.' : 'Elige contactos (con cumpleaños) para ver sus próximos cumpledías y cumpleaños.'));

    const plist = h('div', { class: 'people' });
    for (const p of people) {
      const full = isYmd(p.birth);
      const sub = full ? fmtDay(utcDay(p.birth)) + ' · ' + fmtNum(Math.round((today - utcDay(p.birth)) / DAY)) + ' días' : 'Sin año de nacimiento';
      plist.append(h('button', { type: 'button', class: 'person', onclick: () => p.src === 'google' ? openContacto(p.id) : openPerson(p.id) },
        h('span', { class: 'person-name' }, p.name, p.src === 'manual' ? h('span', { class: 'src-tag' }, 'a mano') : null),
        h('span', { class: 'person-sub' }, sub)));
    }
    if (!people.length) plist.append(h('p', { class: 'muted empty-note' }, 'Todavía no hay nadie: pulsa «Elegir».'));

    return [stats,
      h('div', { class: 'cd-grid' },
        h('section', { class: 'panel' },
          h('div', { class: 'panel-head' }, h('h2', null, 'Próximos'), filters),
          h('div', { class: 'legend cd-legend' },
            h('span', { class: 'lg', style: '--c:#E68AC4' }, 'Cumpleaños'),
            h('span', { class: 'lg', style: '--c:#F2B35B' }, 'Cada 5.000 días'),
            h('span', { class: 'lg', style: '--c:#7FA2E0' }, 'Cada 1.000 días')),
          list),
        h('section', { class: 'panel' },
          h('div', { class: 'panel-head' }, h('h2', null, 'Personas'),
            h('button', { type: 'button', class: 'btn subtle small', onclick: openCdPicker }, ICON.plus(), 'Elegir')),
          plist))];
  }

  function openPerson(id) {
    const cur = id ? (S.datos.colecciones.personas || {})[id] : null;
    const name = h('input', { class: 'input', id: 'pName', type: 'text', maxlength: '60', value: cur ? cur.name : '', placeholder: 'Nombre', autocomplete: 'off' });
    const birth = h('input', { class: 'input', id: 'pBirth', type: 'date', value: cur ? cur.birth : '', max: new Date().toISOString().slice(0, 10), min: '1900-01-01' });
    let close = null;
    const save = () => {
      const n = name.value.trim(), b = birth.value;
      if (!n) { toast('Escribe un nombre.', 'error'); name.focus(); return; }
      if (!isYmd(b) || utcDay(b) > todayUtc() || b < '1900-01-01') { toast('Pon una fecha de nacimiento válida.', 'error'); birth.focus(); return; }
      const rec = { id: cur ? cur.id : 'p-' + uid(), name: n, birth: b };
      S.datos.colecciones.personas = Object.assign({}, S.datos.colecciones.personas, { [rec.id]: Object.assign({}, cur, rec) });
      saveDatos();
      close();
      renderCumpledias();
    };
    name.addEventListener('keydown', e => { if (e.key === 'Enter') save(); });
    const body = h('div', { style: 'display:flex;flex-direction:column;gap:14px' },
      h('div', { class: 'field' }, h('label', { class: 'field-label', for: 'pName' }, 'Nombre'), name),
      h('div', { class: 'field' }, h('label', { class: 'field-label', for: 'pBirth' }, 'Fecha de nacimiento'), birth),
      h('div', { class: 'row-actions modal-foot' },
        cur ? h('button', {
          type: 'button', class: 'btn danger small',
          onclick: () => {
            if (!confirm('¿Borrar a ' + cur.name + '?')) return;
            const rest = Object.assign({}, S.datos.colecciones.personas);
            delete rest[cur.id];
            S.datos.colecciones.personas = rest;
            saveDatos();
            close();
            renderCumpledias();
          }
        }, 'Borrar') : null,
        h('span', { class: 'spacer' }),
        h('button', { type: 'button', class: 'btn subtle small', onclick: () => close() }, 'Cancelar'),
        h('button', { type: 'button', class: 'btn primary small', onclick: save }, cur ? 'Guardar' : 'Añadir')));
    close = openModal(cur ? 'Editar persona' : 'Nueva persona', body);
    setTimeout(() => name.focus(), 50);
  }

  // ---------- gastos y pagos ----------
  // Colecciones: "gastos" (recurrentes) y "pagos" (historial). Cada pago guarda copia del concepto y
  // la categoría para que el historial se conserve aunque se borre el gasto.
  const CATEGORIAS = ['Seguros', 'Impuestos', 'Comunidad', 'Inversiones', 'Suscripciones', 'Consumos', 'Préstamos', 'Alquileres', 'Streaming', 'Otros'];
  const CAT_COLOR = { Seguros: '#7FA2E0', Impuestos: '#F0857A', Comunidad: '#F2B35B', Inversiones: '#8BD17C', Suscripciones: '#B39DF0',
    Consumos: '#F5D76E', 'Préstamos': '#E68AC4', Alquileres: '#D9A77A', Streaming: '#6FC3E8', Otros: '#9AA4B5' };
  const PERIODOS = { mensual: 1, bimestral: 2, trimestral: 3, semestral: 6, anual: 12 };
  const PERIODO_NOMBRE = { mensual: 'Mensual', bimestral: 'Bimestral', trimestral: 'Trimestral', semestral: 'Semestral', anual: 'Anual', personalizado: 'Meses elegidos' };
  const MESES_CORTOS = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];
  const TOL_ANTES = 15, GRACIA = 20; // días: un pago cubre un vencimiento si cae entre 15 días antes y 20 después
  let gTab = 'proximos';   // proximos | historial | resumen
  let rYear = null;        // año del resumen (null: el actual)
  let gChart = null;
  let gCat = 'Todas';
  let hYear = 'Todos';

  const gastos = () => Object.values(S.datos.colecciones.gastos || {});
  const pagos = () => Object.values(S.datos.colecciones.pagos || {});
  const fmtEur = n => (Number(n) || 0).toLocaleString('es-ES', { style: 'currency', currency: 'EUR' });
  const catColor = c => CAT_COLOR[c] || CAT_COLOR.Otros;
  // Importes escritos a mano: admite coma o punto decimal ("25,40", "25.4", "1.250,00")
  const fmtNumInput = n => (Math.round((Number(n) || 0) * 100) / 100).toFixed(2).replace('.', ',');
  function parseAmount(v) {
    let t = String(v || '').trim().replace(/\s|€/g, '');
    if (t.includes(',')) t = t.replace(/\./g, '').replace(',', '.');
    return /^\d+(\.\d+)?$/.test(t) ? parseFloat(t) : NaN;
  }
  const todayYmd = () => { const n = new Date(); return n.getFullYear() + '-' + String(n.getMonth() + 1).padStart(2, '0') + '-' + String(n.getDate()).padStart(2, '0'); };
  const mesesPara = (periodo, desde) => {
    const step = PERIODOS[periodo] || 1;
    if (step === 1) return [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12];
    const out = [];
    for (let m = ((desde - 1) % step) + 1; m <= 12; m += step) out.push(m);
    return out;
  };

  // Vencimientos (días UTC) entre dos años; el día se ajusta al último del mes si no existe (30 en febrero → 28/29)
  function vencimientos(g, y1, y2) {
    const meses = Array.isArray(g.meses) && g.meses.length ? g.meses.slice().sort((a, b) => a - b) : mesesPara(g.periodicidad, 1);
    const out = [];
    for (let y = y1; y <= y2; y++) {
      for (const m of meses) {
        const last = new Date(Date.UTC(y, m, 0)).getUTCDate();
        out.push(Date.UTC(y, m - 1, Math.min(g.dia || 1, last)));
      }
    }
    return out;
  }

  const lastPago = id => pagos().filter(p => p.gastoId === id && isYmd(p.fecha)).map(p => utcDay(p.fecha)).sort((a, b) => b - a)[0] || null;

  // Próximo vencimiento: el último pago cubre su vencimiento y toca el siguiente; sin pago, el primero desde hoy.
  // La fecha fija (proximoPago) manda mientras no haya un pago que la cubra.
  function nextDue(g, today) {
    const last = lastPago(g.id);
    if (isYmd(g.proximoPago) && !(last && last >= utcDay(g.proximoPago) - TOL_ANTES * DAY)) return utcDay(g.proximoPago);
    const y = new Date(today).getUTCFullYear();
    const v = vencimientos(g, y - 1, y + 1);
    if (!v.length) return null;
    if (last) {
      let cubierto = null;
      for (const d of v) if (last >= d - TOL_ANTES * DAY && last <= d + GRACIA * DAY) cubierto = d;
      if (cubierto !== null) {
        const sig = v.find(d => d > cubierto);
        if (sig !== undefined) return sig;
      }
    }
    const fut = v.find(d => d >= today);
    return fut !== undefined ? fut : v[v.length - 1];
  }

  // Fecha corta: sin el año si es el actual
  const shortDay = t => new Date(t).getUTCFullYear() === new Date(todayUtc()).getUTCFullYear() ? fmtDay(t).replace(/ \d{4}$/, '') : fmtDay(t);

  function dueText(n) {
    if (n < 0) return n === -1 ? 'venció ayer' : 'vencido hace ' + (-n) + ' días';
    if (n === 0) return 'hoy';
    if (n === 1) return 'mañana';
    return 'en ' + n + ' días';
  }

  function renderGastos() {
    const root = document.getElementById('viewGastos');
    const today = todayUtc();
    root.dataset.day = String(today);
    const activos = gastos().filter(g => g.activo !== false).map(g => {
      const due = nextDue(g, today);
      return { g, due, until: due === null ? null : Math.round((due - today) / DAY) };
    });
    const t = new Date(today);
    const finMes = Date.UTC(t.getUTCFullYear(), t.getUTCMonth() + 1, 0);
    const hastaFin = activos.filter(x => x.until !== null && x.until >= 0 && x.due <= finMes);
    const vencidos = activos.filter(x => x.until !== null && x.until < 0);
    // Mes que viene: todos los vencimientos de ese mes a partir del próximo de cada gasto
    // (un gasto mensual que toca este mes cuenta también su cuota del mes siguiente)
    const ini = Date.UTC(t.getUTCFullYear(), t.getUTCMonth() + 1, 1), fin = Date.UTC(t.getUTCFullYear(), t.getUTCMonth() + 2, 0);
    const mesQueViene = [];
    for (const x of activos) {
      if (x.due === null) continue;
      const y = t.getUTCFullYear();
      const fechas = new Set(vencimientos(x.g, y, y + 1).filter(d => d >= x.due));
      fechas.add(x.due);
      for (const d of fechas) if (d >= ini && d <= fin) mesQueViene.push(x);
    }
    const nombreMes = new Date(ini).toLocaleDateString('es-ES', { month: 'long', timeZone: 'UTC' });
    const sum = arr => arr.reduce((a, x) => a + (Number(x.g.importe) || 0), 0);
    const stat = (label, value, sub, cls) => h('div', { class: 'stat' + (cls ? ' ' + cls : '') },
      h('span', { class: 'stat-label' }, label), h('span', { class: 'stat-value' }, value), h('span', { class: 'stat-sub' }, sub));

    const tabs = h('div', { class: 'seg', role: 'group', 'aria-label': 'Vista' },
      [['proximos', 'Próximos'], ['historial', 'Historial'], ['resumen', 'Resumen']].map(([k, label]) =>
        h('button', { type: 'button', class: gTab === k ? 'on' : '', 'aria-pressed': gTab === k ? 'true' : 'false', onclick: () => { gTab = k; renderGastos(); } }, label)));

    const head = h('div', { class: 'g-toolbar' }, tabs,
      h('button', { type: 'button', class: 'btn ghost small', onclick: () => openPago(null) }, ICON.plus(), 'Registrar pago'));

    if (gChart) { gChart.destroy(); gChart = null; }
    if (gTab === 'resumen') {
      root.replaceChildren(head, ...resumenPanels());
      return;
    }
    if (gTab === 'historial') {
      root.replaceChildren(head, historialPanel());
      return;
    }

    const cats = ['Todas'].concat(CATEGORIAS.filter(c => gastos().some(g => g.categoria === c)));
    if (!cats.includes(gCat)) gCat = 'Todas';
    const chips = h('div', { class: 'chips' }, cats.map(c => h('button', {
      type: 'button', class: 'chip' + (gCat === c ? ' on' : ''), style: '--c:' + (c === 'Todas' ? '#C9D0DC' : catColor(c)),
      'aria-pressed': gCat === c ? 'true' : 'false', onclick: () => { gCat = c; renderGastos(); }
    }, c === 'Todas' ? null : h('span', { class: 'dot' }), c)));

    const inCat = x => gCat === 'Todas' || x.g.categoria === gCat;
    const list = h('div', { class: 'ev-list' });
    const rows = activos.filter(inCat).sort((a, b) => (a.due ?? Infinity) - (b.due ?? Infinity) || a.g.concepto.localeCompare(b.g.concepto, 'es'));
    for (const x of rows) list.append(gastoRow(x.g, x.due, x.until));
    const pausados = gastos().filter(g => g.activo === false).filter(g => inCat({ g }));
    if (pausados.length) {
      list.append(h('p', { class: 'list-sep' }, 'Pausados'));
      for (const g of pausados.sort((a, b) => a.concepto.localeCompare(b.concepto, 'es'))) list.append(gastoRow(g, null, null, true));
    }
    if (!rows.length && !pausados.length) list.append(h('p', { class: 'muted empty-note' }, gastos().length ? 'No hay gastos en esta categoría.' : 'Crea tu primer gasto recurrente con «Nuevo gasto».'));

    root.replaceChildren(
      h('div', { class: 'stats three' },
        stat('Hasta fin de mes', fmtEur(sum(hastaFin)), hastaFin.length + (hastaFin.length === 1 ? ' gasto' : ' gastos'), 'accent'),
        stat('Mes que viene', fmtEur(sum(mesQueViene)), mesQueViene.length + (mesQueViene.length === 1 ? ' cobro' : ' cobros') + ' en ' + nombreMes),
        stat('Vencidos', String(vencidos.length), vencidos.length ? fmtEur(sum(vencidos)) + ' sin registrar' : 'todo al día', vencidos.length ? 'danger' : '')),
      head,
      h('section', { class: 'panel' }, chips, list));
  }

  function gastoRow(g, due, until, paused) {
    const sub = [g.categoria || 'Otros', PERIODO_NOMBRE[g.periodicidad] || 'Mensual', 'día ' + (g.dia || 1)];
    if (g.proveedor) sub.push(g.proveedor);
    if (g.cobro === 'automatico') sub.push('domiciliado');
    const chipCls = until === null ? '' : until < 0 ? ' late' : until === 0 ? ' now' : until <= 7 ? ' soon' : '';
    return h('div', { class: 'ev gasto' + (paused ? ' paused' : '') + (until !== null && until < 0 ? ' late' : '') },
      h('span', { class: 'cat-dot', style: '--c:' + catColor(g.categoria), title: g.categoria || 'Otros' }),
      h('button', { type: 'button', class: 'ev-main as-btn', onclick: () => openGasto(g.id), title: 'Editar' },
        h('span', { class: 'ev-name' }, g.concepto),
        h('span', { class: 'ev-what' }, sub.join(' · '))),
      h('div', { class: 'ev-when' },
        h('span', { class: 'ev-amount' }, fmtEur(g.importe)),
        paused ? h('span', { class: 'ev-until' }, 'pausado')
          : h('span', { class: 'ev-until' + chipCls, title: due ? fmtDay(due) : '' }, due ? dueText(until) + ' · ' + shortDay(due) : 'sin fecha')),
      paused ? h('span') : h('button', { type: 'button', class: 'btn ghost small', onclick: () => openPago(g.id) }, 'Pagar'));
  }

  // Resumen anual: lo pagado (historial) por mes y categoría, y previsión hasta fin de año
  function resumenPanels() {
    const today = todayUtc(), yNow = new Date(today).getUTCFullYear();
    const all = pagos().filter(p => isYmd(p.fecha));
    const years = [...new Set(all.map(p => +p.fecha.slice(0, 4)).concat(yNow))].sort((a, b) => b - a);
    if (!years.includes(rYear)) rYear = yNow;
    const y = rYear;
    const delAno = all.filter(p => +p.fecha.slice(0, 4) === y);
    const pagado = delAno.reduce((a, p) => a + (Number(p.importe) || 0), 0);
    const meses = y < yNow ? 12 : y === yNow ? new Date(today).getUTCMonth() + 1 : 0;
    // Pendiente hasta el 31 de diciembre: los vencimientos de cada gasto activo desde su próximo (vencidos incluidos)
    let pendiente = 0, nPend = 0;
    if (y === yNow) {
      const fin = Date.UTC(y, 11, 31);
      for (const g of gastos().filter(x => x.activo !== false)) {
        const due = nextDue(g, today);
        if (due === null) continue;
        const fechas = new Set(vencimientos(g, y, y).filter(d => d >= due));
        if (due <= fin) fechas.add(due);
        for (const d of fechas) if (d <= fin) { pendiente += Number(g.importe) || 0; nPend++; }
      }
    }
    const stat = (label, value, sub, cls) => h('div', { class: 'stat' + (cls ? ' ' + cls : '') },
      h('span', { class: 'stat-label' }, label), h('span', { class: 'stat-value' }, value), h('span', { class: 'stat-sub' }, sub));

    const yearChips = h('div', { class: 'chips' }, years.map(yr => h('button', {
      type: 'button', class: 'chip' + (yr === y ? ' on' : ''), style: '--c:#C9D0DC', 'aria-pressed': yr === y ? 'true' : 'false',
      onclick: () => { rYear = yr; renderGastos(); }
    }, String(yr))));

    const stats = h('div', { class: 'stats three' },
      stat('Pagado en ' + y, fmtEur(pagado), delAno.length + (delAno.length === 1 ? ' pago' : ' pagos'), 'accent'),
      stat('Media mensual', meses ? fmtEur(pagado / meses) : '—', meses ? 'en ' + meses + (meses === 1 ? ' mes' : ' meses') : 'año sin empezar'),
      y === yNow
        ? stat('Previsión del año', fmtEur(pagado + pendiente), 'quedan ' + fmtEur(pendiente) + ' en ' + nPend + (nPend === 1 ? ' cobro' : ' cobros'))
        : stat('Previsión del año', fmtEur(pagado), y < yNow ? 'año cerrado' : 'sin datos'));

    // Por categoría
    const porCat = {};
    for (const p of delAno) porCat[p.categoria || 'Otros'] = (porCat[p.categoria || 'Otros'] || 0) + (Number(p.importe) || 0);
    const cats = Object.entries(porCat).sort((a, b) => b[1] - a[1]);
    const max = cats.length ? cats[0][1] : 0;
    const catList = h('div', { class: 'cat-list' }, cats.map(([c, v]) => h('div', { class: 'cat-row' },
      h('span', { class: 'cat-dot', style: '--c:' + catColor(c) }),
      h('span', { class: 'cat-name' }, c),
      h('span', { class: 'cat-bar' }, h('span', { style: 'width:' + (max ? Math.max(2, v / max * 100) : 0).toFixed(1) + '%;--c:' + catColor(c) })),
      h('span', { class: 'cat-pct' }, (pagado ? Math.round(v / pagado * 100) : 0) + ' %'),
      h('span', { class: 'ev-amount' }, fmtEur(v)))));
    if (!cats.length) catList.append(h('p', { class: 'muted empty-note' }, 'No hay pagos registrados en ' + y + '.'));

    // Gráfico: barras apiladas por categoría y línea con el total mensual del año anterior
    const porMes = cat => Array.from({ length: 12 }, (_, m) => {
      const v = delAno.filter(p => (p.categoria || 'Otros') === cat && +p.fecha.slice(5, 7) === m + 1).reduce((a, p) => a + (Number(p.importe) || 0), 0);
      return v ? Math.round(v * 100) / 100 : null;
    });
    const prevTotals = Array.from({ length: 12 }, (_, m) => {
      const v = all.filter(p => +p.fecha.slice(0, 4) === y - 1 && +p.fecha.slice(5, 7) === m + 1).reduce((a, p) => a + (Number(p.importe) || 0), 0);
      return v ? Math.round(v * 100) / 100 : null;
    });
    const hasPrev = prevTotals.some(v => v !== null);
    const canvas = h('canvas', { 'aria-label': 'Gasto por mes en ' + y + ' por categoría' });
    const chartPanel = h('section', { class: 'panel' },
      h('div', { class: 'panel-head' }, h('h2', null, 'Por mes · ' + y),
        hasPrev ? h('div', { class: 'legend' }, h('span', { class: 'lg line' }, 'Total ' + (y - 1))) : null),
      cats.length ? h('div', { class: 'chart-box' }, canvas) : h('p', { class: 'muted empty-note' }, 'Sin datos para el gráfico.'));
    if (cats.length) {
      loadChart().then(() => {
        if (!canvas.isConnected) return;
        if (gChart) gChart.destroy();
        const datasets = cats.map(([c]) => ({ type: 'bar', label: c, data: porMes(c), backgroundColor: catColor(c), stack: 'pagos', borderRadius: 3, order: 2 }));
        if (hasPrev) datasets.push({ type: 'line', label: 'Total ' + (y - 1), data: prevTotals, borderColor: 'rgba(201, 208, 220, 0.75)', backgroundColor: 'rgba(201, 208, 220, 0.75)', borderDash: [5, 4], borderWidth: 1.5, pointRadius: 2, tension: 0.25, spanGaps: true, order: 1 });
        gChart = new Chart(canvas, {
          data: { labels: MESES_CORTOS, datasets },
          options: {
            responsive: true, maintainAspectRatio: false, animation: false,
            interaction: { mode: 'index', intersect: false },
            plugins: { legend: { display: false }, tooltip: {
              filter: it => it.raw !== null,
              callbacks: { label: ctx => ctx.dataset.label + ': ' + fmtEur(ctx.raw), footer: items => { const t = items.filter(i => i.dataset.type === 'bar').reduce((a, i) => a + (i.raw || 0), 0); return t ? 'Total ' + y + ': ' + fmtEur(t) : ''; } }
            } },
            scales: {
              x: { stacked: true, ticks: { color: '#8A94A6', font: { family: 'Manrope', size: 11 } }, grid: { display: false }, border: { display: false } },
              y: { stacked: true, beginAtZero: true, ticks: { color: '#8A94A6', font: { family: 'Manrope', size: 11 }, maxTicksLimit: 5, callback: v => v + ' €' }, grid: { color: 'rgba(255, 255, 255, 0.06)' }, border: { display: false } }
            }
          }
        });
      }).catch(e => toast(e.message, 'error'));
    }

    return [yearChips, stats, chartPanel,
      h('section', { class: 'panel' }, h('div', { class: 'panel-head' }, h('h2', null, 'Por categoría · ' + y), h('span', { class: 'panel-total' }, fmtEur(pagado))), catList)];
  }

  function historialPanel() {
    const all = pagos().filter(p => isYmd(p.fecha)).sort((a, b) => b.fecha.localeCompare(a.fecha) || a.concepto.localeCompare(b.concepto, 'es'));
    const years = ['Todos'].concat([...new Set(all.map(p => p.fecha.slice(0, 4)))].sort().reverse());
    if (!years.includes(hYear)) hYear = 'Todos';
    const cats = ['Todas'].concat(CATEGORIAS.filter(c => all.some(p => p.categoria === c)));
    if (!cats.includes(gCat)) gCat = 'Todas';
    const shown = all.filter(p => (hYear === 'Todos' || p.fecha.startsWith(hYear)) && (gCat === 'Todas' || p.categoria === gCat));
    const total = shown.reduce((a, p) => a + (Number(p.importe) || 0), 0);

    const chips = h('div', { class: 'chips' },
      cats.map(c => h('button', {
        type: 'button', class: 'chip' + (gCat === c ? ' on' : ''), style: '--c:' + (c === 'Todas' ? '#C9D0DC' : catColor(c)),
        'aria-pressed': gCat === c ? 'true' : 'false', onclick: () => { gCat = c; renderGastos(); }
      }, c === 'Todas' ? null : h('span', { class: 'dot' }), c)),
      h('span', { class: 'chips-sep' }),
      years.map(y => h('button', {
        type: 'button', class: 'chip' + (hYear === y ? ' on' : ''), style: '--c:#C9D0DC', 'aria-pressed': hYear === y ? 'true' : 'false',
        onclick: () => { hYear = y; renderGastos(); }
      }, y)));

    const list = h('div', { class: 'ev-list' });
    for (const p of shown) {
      list.append(h('div', { class: 'ev pago' },
        h('span', { class: 'cat-dot', style: '--c:' + catColor(p.categoria), title: p.categoria || 'Otros' }),
        h('div', { class: 'ev-main' },
          h('span', { class: 'ev-name' }, p.concepto),
          h('span', { class: 'ev-what' }, [fmtDay(utcDay(p.fecha)), p.categoria || 'Otros'].concat(p.notas ? [p.notas] : []).join(' · '))),
        h('span', { class: 'ev-amount' }, fmtEur(p.importe)),
        h('button', {
          type: 'button', class: 'icon-btn mini', 'aria-label': 'Borrar pago',
          onclick: () => {
            if (!confirm('¿Borrar el pago de «' + p.concepto + '» del ' + fmtDay(utcDay(p.fecha)) + '?')) return;
            const rest = Object.assign({}, S.datos.colecciones.pagos);
            delete rest[p.id];
            S.datos.colecciones.pagos = rest;
            saveDatos();
            renderGastos();
          }
        }, ICON.trash())));
    }
    if (!shown.length) list.append(h('p', { class: 'muted empty-note' }, all.length ? 'No hay pagos con estos filtros.' : 'Aún no hay pagos registrados.'));

    return h('section', { class: 'panel' }, chips,
      h('div', { class: 'panel-head' }, h('h2', null, shown.length + (shown.length === 1 ? ' pago' : ' pagos')), h('span', { class: 'panel-total' }, fmtEur(total))),
      list);
  }

  function putRecord(col, rec) {
    S.datos.colecciones[col] = Object.assign({}, S.datos.colecciones[col], { [rec.id]: rec });
    saveDatos();
  }

  function openGasto(id) {
    const cur = id ? (S.datos.colecciones.gastos || {})[id] : null;
    const st = {
      periodicidad: cur && (PERIODOS[cur.periodicidad] || cur.periodicidad === 'personalizado') ? cur.periodicidad : 'mensual',
      meses: cur && Array.isArray(cur.meses) && cur.meses.length ? cur.meses.slice() : mesesPara('mensual', 1),
      cobro: cur ? (cur.cobro === 'manual' ? 'manual' : 'automatico') : 'automatico'
    };
    const inp = (idAttr, attrs) => h('input', Object.assign({ class: 'input', id: idAttr, autocomplete: 'off' }, attrs));
    const concepto = inp('gConcepto', { type: 'text', maxlength: '80', value: cur ? cur.concepto : '', placeholder: 'Por ejemplo: Seguro del coche' });
    const proveedor = inp('gProveedor', { type: 'text', maxlength: '80', value: cur ? cur.proveedor || '' : '', placeholder: 'Opcional' });
    const importe = inp('gImporte', { type: 'text', inputmode: 'decimal', value: cur ? fmtNumInput(cur.importe) : '', placeholder: '0,00' });
    const dia = inp('gDia', { type: 'number', min: '1', max: '31', step: '1', inputmode: 'numeric', value: String(cur ? cur.dia || 1 : new Date().getDate()) });
    const categoria = h('select', { class: 'select', id: 'gCategoria' }, CATEGORIAS.map(c => h('option', { value: c, selected: (cur ? cur.categoria : 'Otros') === c }, c)));
    const periodo = h('select', { class: 'select', id: 'gPeriodo' }, Object.entries(PERIODO_NOMBRE).map(([k, v]) => h('option', { value: k, selected: st.periodicidad === k }, v)));
    const proximo = inp('gProximo', { type: 'date', value: cur && isYmd(cur.proximoPago) ? cur.proximoPago : '' });
    const notas = h('textarea', { class: 'input textarea', id: 'gNotas', rows: '2', maxlength: '500' });
    notas.value = cur ? cur.notas || '' : '';
    const pausado = h('input', { type: 'checkbox', id: 'gPausado', checked: cur ? cur.activo === false : false });
    const enCal = h('input', { type: 'checkbox', id: 'gCal', checked: cur ? !!cur.calendario : false });

    const mesesBox = h('div', { class: 'months' });
    const drawMeses = () => mesesBox.replaceChildren(...MESES_CORTOS.map((n, i) => {
      const m = i + 1, on = st.meses.includes(m);
      return h('button', {
        type: 'button', class: 'month' + (on ? ' on' : ''), 'aria-pressed': on ? 'true' : 'false',
        onclick: () => { st.meses = on ? st.meses.filter(x => x !== m) : st.meses.concat(m).sort((a, b) => a - b); drawMeses(); }
      }, n);
    }));
    periodo.addEventListener('change', () => {
      st.periodicidad = periodo.value;
      if (st.periodicidad !== 'personalizado') st.meses = mesesPara(st.periodicidad, st.meses[0] || new Date().getMonth() + 1);
      drawMeses();
    });
    drawMeses();
    const cobroSeg = h('div', { class: 'seg' });
    const drawCobro = () => cobroSeg.replaceChildren(...[['automatico', 'Domiciliado'], ['manual', 'Pago manual']].map(([k, v]) =>
      h('button', { type: 'button', class: st.cobro === k ? 'on' : '', 'aria-pressed': st.cobro === k ? 'true' : 'false', onclick: () => { st.cobro = k; drawCobro(); } }, v)));
    drawCobro();

    let close = null;
    const save = () => {
      const c = concepto.value.trim();
      const imp = parseAmount(importe.value);
      const d = parseInt(dia.value, 10);
      if (!c) { toast('Escribe el concepto.', 'error'); concepto.focus(); return; }
      if (!(imp >= 0)) { toast('Pon un importe válido.', 'error'); importe.focus(); return; }
      if (!(d >= 1 && d <= 31)) { toast('El día debe estar entre 1 y 31.', 'error'); dia.focus(); return; }
      if (!st.meses.length) { toast('Marca al menos un mes de cobro.', 'error'); return; }
      const rec = Object.assign({}, cur, {
        id: cur ? cur.id : 'g-' + uid(), concepto: c, proveedor: proveedor.value.trim(), categoria: categoria.value,
        importe: Math.round(imp * 100) / 100, periodicidad: st.periodicidad, meses: st.meses.slice(), dia: d,
        proximoPago: isYmd(proximo.value) ? proximo.value : null, cobro: st.cobro, notas: notas.value.trim(),
        activo: !pausado.checked, calendario: enCal.checked
      });
      if (!cur) rec.creado = Date.now();
      putRecord('gastos', rec);
      close();
      renderGastos();
    };
    const field = (label, forId, el, hint) => h('div', { class: 'field' }, h('label', { class: 'field-label', for: forId }, label), el, hint ? h('p', { class: 'hint' }, hint) : null);
    const body = h('div', { class: 'form' },
      field('Concepto', 'gConcepto', concepto),
      h('div', { class: 'form-row' }, field('Importe (€)', 'gImporte', importe), field('Categoría', 'gCategoria', categoria)),
      h('div', { class: 'form-row' }, field('Periodicidad', 'gPeriodo', periodo), field('Día de cobro', 'gDia', dia)),
      h('div', { class: 'field' }, h('span', { class: 'field-label' }, 'Meses de cobro'), mesesBox,
        h('p', { class: 'hint' }, 'Se marcan según la periodicidad; pulsa un mes para cambiarlo.')),
      field('Próximo pago fijo', 'gProximo', proximo, 'Opcional. Úsalo si te han cambiado la fecha de cobro; deja de aplicarse al registrar el pago.'),
      h('div', { class: 'form-row' }, field('Proveedor', 'gProveedor', proveedor), h('div', { class: 'field' }, h('span', { class: 'field-label' }, 'Cobro'), cobroSeg)),
      field('Notas', 'gNotas', notas),
      h('label', { class: 'check-line', for: 'gCal' }, enCal, 'Mostrar en el calendario (un evento en cada vencimiento)'),
      h('label', { class: 'check-line', for: 'gPausado' }, pausado, 'Pausado (no cuenta en vencimientos ni totales)'),
      h('div', { class: 'row-actions modal-foot' },
        cur ? h('button', {
          type: 'button', class: 'btn danger small',
          onclick: () => {
            if (!confirm('¿Borrar «' + cur.concepto + '»? Los pagos ya registrados se quedan en el historial.')) return;
            const rest = Object.assign({}, S.datos.colecciones.gastos);
            delete rest[cur.id];
            S.datos.colecciones.gastos = rest;
            saveDatos();
            close();
            renderGastos();
          }
        }, 'Borrar') : null,
        h('span', { class: 'spacer' }),
        h('button', { type: 'button', class: 'btn subtle small', onclick: () => close() }, 'Cancelar'),
        h('button', { type: 'button', class: 'btn primary small', onclick: save }, cur ? 'Guardar' : 'Crear gasto')));
    close = openModal(cur ? 'Editar gasto' : 'Nuevo gasto', body);
    if (!cur) setTimeout(() => concepto.focus(), 50);
  }

  function openPago(gastoId) {
    const lista = gastos().sort((a, b) => a.concepto.localeCompare(b.concepto, 'es'));
    if (!lista.length) { toast('Primero crea un gasto.', 'error'); return; }
    const sel = h('select', { class: 'select', id: 'pgGasto' }, lista.map(g => h('option', { value: g.id, selected: g.id === gastoId }, g.concepto + (g.activo === false ? ' (pausado)' : ''))));
    const g0 = () => (S.datos.colecciones.gastos || {})[sel.value];
    const importe = h('input', { class: 'input', id: 'pgImporte', type: 'text', inputmode: 'decimal', autocomplete: 'off', value: fmtNumInput(g0().importe) });
    const fecha = h('input', { class: 'input', id: 'pgFecha', type: 'date', value: todayYmd() });
    const notas = h('input', { class: 'input', id: 'pgNotas', type: 'text', maxlength: '200', placeholder: 'Opcional', autocomplete: 'off' });
    // Si el importe no es el del gasto: ¿solo esta vez o es el importe nuevo a partir de ahora?
    let alcance = 'temporal';
    const alcanceSeg = h('div', { class: 'seg' });
    const drawAlcance = () => alcanceSeg.replaceChildren(...[['temporal', 'Solo este pago'], ['fijo', 'Nuevo importe fijo']].map(([k, v]) =>
      h('button', { type: 'button', class: alcance === k ? 'on' : '', 'aria-pressed': alcance === k ? 'true' : 'false', onclick: () => { alcance = k; drawAlcance(); } }, v)));
    drawAlcance();
    const alcanceHint = h('p', { class: 'hint' });
    const alcanceBox = h('div', { class: 'field', hidden: true }, h('span', { class: 'field-label' }, 'Este importe es distinto del habitual'), alcanceSeg, alcanceHint);
    const checkImporte = () => {
      const g = g0(), imp = parseAmount(importe.value);
      const distinto = g && imp >= 0 && Math.abs(imp - (Number(g.importe) || 0)) >= 0.005;
      alcanceBox.hidden = !distinto;
      if (distinto) alcanceHint.textContent = 'Importe habitual de «' + g.concepto + '»: ' + fmtEur(g.importe) + '. Con «Nuevo importe fijo» se actualiza el gasto para los próximos cobros.';
    };
    importe.addEventListener('input', checkImporte);
    sel.addEventListener('change', () => { importe.value = fmtNumInput(g0().importe); alcance = 'temporal'; drawAlcance(); checkImporte(); });
    let close = null;
    const save = () => {
      const g = g0();
      const imp = parseAmount(importe.value);
      if (!g) return;
      if (!(imp >= 0)) { toast('Pon un importe válido.', 'error'); importe.focus(); return; }
      if (!isYmd(fecha.value)) { toast('Pon una fecha válida.', 'error'); fecha.focus(); return; }
      const cantidad = Math.round(imp * 100) / 100;
      putRecord('pagos', { id: 'pg-' + uid(), gastoId: g.id, concepto: g.concepto, categoria: g.categoria || 'Otros', fecha: fecha.value, importe: cantidad, notas: notas.value.trim() });
      const fijo = !alcanceBox.hidden && alcance === 'fijo';
      if (fijo) putRecord('gastos', Object.assign({}, g, { importe: cantidad }));
      close();
      renderGastos();
      toast('Pago registrado: ' + g.concepto + (fijo ? '. Nuevo importe: ' + fmtEur(cantidad) : ''));
    };
    notas.addEventListener('keydown', e => { if (e.key === 'Enter') save(); });
    const field = (label, forId, el) => h('div', { class: 'field' }, h('label', { class: 'field-label', for: forId }, label), el);
    const body = h('div', { class: 'form' },
      field('Gasto', 'pgGasto', sel),
      h('div', { class: 'form-row' }, field('Fecha', 'pgFecha', fecha), field('Importe (€)', 'pgImporte', importe)),
      alcanceBox,
      field('Notas', 'pgNotas', notas),
      h('div', { class: 'row-actions modal-foot' },
        h('span', { class: 'spacer' }),
        h('button', { type: 'button', class: 'btn subtle small', onclick: () => close() }, 'Cancelar'),
        h('button', { type: 'button', class: 'btn primary small', onclick: save }, 'Registrar pago')));
    close = openModal('Registrar pago', body);
  }

  // ---------- consumos: luz, gas y agua ----------
  // Colección "consumos": un registro por factura, id "<tipo>-<fin>". Cada factura cuenta en el mes
  // en que termina el consumo (anio/mes = mes de "fin"). Los PDF van a OneDrive: facturas/AAAA/<tipo>_<fin>.pdf
  const TIPOS = {
    luz: { nombre: 'Luz', unidad: 'kWh', campo: 'kwh', color: '#F5D76E' },
    gas: { nombre: 'Gas', unidad: 'kWh', campo: 'kwh', color: '#8BD17C' },
    agua: { nombre: 'Agua', unidad: 'm³', campo: 'm3', color: '#6FC3E8' }
  };
  const MESES_LARGOS = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
  let cTipo = 'luz';
  let cMetric = 'coste'; // coste | consumo
  let chartObj = null;
  let chartLoading = null;
  let pdfLoading = null;

  const consumos = () => Object.values(S.datos.colecciones.consumos || {});
  const fmtCant = (n, u) => n === null || n === undefined ? '—' : (Math.round(n * 100) / 100).toLocaleString('es-ES') + ' ' + u;
  const cap = s => s.charAt(0).toUpperCase() + s.slice(1);

  // ---- lectores de facturas (texto extraído con pdf.js; se quitan los espacios antes de buscar) ----
  // Número en formato español: "1.234,56" -> 1234.56
  function numEs(s) {
    if (s === undefined || s === null) return null;
    let t = String(s).trim();
    if (t.includes(',')) t = t.replace(/\./g, '').replace(',', '.');
    const n = parseFloat(t);
    return isFinite(n) ? n : null;
  }
  const dmyToYmd = s => { const m = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/.exec(s || ''); if (!m) return null; const y = m[3].length === 2 ? '20' + m[3] : m[3]; return y + '-' + m[2].padStart(2, '0') + '-' + m[1].padStart(2, '0'); };
  const grab = (re, t, i = 1) => { const m = re.exec(t); return m ? m[i] : undefined; };
  // En los patrones, # es un número ("1.234,56")
  const R = s => new RegExp(s.replace(/#/g, '([\\d.]+(?:,\\d+)?)'));

  function parseFactura(text) {
    const t = text.replace(/\s+/g, '');
    const emision = dmyToYmd(grab(/Fechaemisiónfactura:?(\d{1,2}\/\d{1,2}\/\d{4})/, t) || grab(/FechaEmisión:?(\d{1,2}\/\d{1,2}\/\d{4})/, t));
    let d = null;
    if (/DATOSDELAFACTURADEELECTRICIDAD|INFORMACIÓNDELCONSUMOELÉCTRICO|RESUMENDELAFACTURADEELECTRICIDAD/.test(t)) {
      const per = /del(\d{2}\/\d{2}\/\d{4})a(\d{2}\/\d{2}\/\d{4})/.exec(t);
      d = { tipo: 'luz', inicio: per && dmyToYmd(per[1]), fin: per && dmyToYmd(per[2]),
        kwh: numEs(grab(R('ConsumoTotal#kWh'), t)),
        precioKwh: numEs(grab(R('Consumo#kWhx#Eur\\/kWh'), t, 2)),
        costeEnergia: numEs(grab(R('FacturadeElectricidad#€'), t)),
        costeServicios: numEs(grab(R('Facturadeservicios#€'), t)),
        coste: numEs(grab(R('Totalimporteapagar#€'), t) || grab(R('TOTAL#€'), t)) };
      // Formato antiguo sin factura de servicios: todo es energía
      if (d.costeEnergia === null && d.costeServicios === null && !/FACTURADESERVICIOS/.test(t)) d.costeEnergia = d.coste;
    } else if (/DATOSDELAFACTURADEGAS|RESUMENDELAFACTURADEGAS|INFORMACIÓNDELCONSUMOGAS|GASNATURA/.test(t)) {
      const per = /del(\d{2}\/\d{2}\/\d{4})al(\d{2}\/\d{2}\/\d{4})/.exec(t);
      d = { tipo: 'gas', inicio: per && dmyToYmd(per[1]), fin: per && dmyToYmd(per[2]),
        kwh: numEs(grab(R('ConsumoTotal#kWh'), t)),
        m3: numEs(grab(R('Consumo#m[³3]'), t)),
        precioKwh: numEs(grab(R('TérminoEnergíaGas#kWhx#Eur\\/kWh'), t, 2)),
        costeEnergia: numEs(grab(R('FacturadeGas#€'), t) || grab(R('TOTALIMPORTEFACTURADEGAS#€'), t) || grab(R('TOTALIMPORTEFACTURA#€'), t)),
        costeServicios: numEs(grab(R('Facturadeservicios#€'), t)),
        coste: numEs(grab(R('(?:Totalimporteapagar|TOTALAPAGAR)#€'), t)) };
    } else if (/EMIVASA|FACTURASERVICIOSMEDIOAMBIENTALES/.test(t)) {
      const per = /Lecturaanterior.*?Lecturaactual.*?(\d{2}\/\d{2}\/\d{4})(\d{2}\/\d{2}\/\d{4})/.exec(t);
      const corto = /CUOTADECONSUMOAGUA(\d{2}\/\d{2}\/\d{2})-(\d{2}\/\d{2}\/\d{2})/.exec(t);
      d = { tipo: 'agua', inicio: per ? dmyToYmd(per[1]) : corto && dmyToYmd(corto[1]), fin: per ? dmyToYmd(per[2]) : corto && dmyToYmd(corto[2]),
        m3: numEs(grab(R('Consumo#m3'), t)),
        precioM3: numEs(grab(R('CUOTADECONSUMOAGUA[\\d\\/-]+#M3x#€\\/M3'), t, 2)),
        costeEnergia: numEs(grab(R('AGUAPOTABLESubtotal:#'), t)),
        coste: numEs(grab(R('Totalfactura:?#€'), t) || grab(R('TotalFactura:?#€'), t)) };
    }
    if (!d) return null;
    // Sin resumen de servicios: se suman las facturas de servicios detalladas
    if (d.tipo !== 'agua' && d.costeServicios === null) {
      const sv = [...t.matchAll(/DETALLEDELAFACTURADESERVICIOS.*?TOTALIMPORTEFACTURA([\d.]+(?:,\d+)?)€/g)].map(m => numEs(m[1]));
      if (sv.length) d.costeServicios = Math.round(sv.reduce((a, b) => a + b, 0) * 100) / 100;
    }
    if (d.costeEnergia === null && d.coste !== null && d.costeServicios !== null) d.costeEnergia = Math.round((d.coste - d.costeServicios) * 100) / 100;
    if (d.coste === null && (d.costeEnergia !== null || d.costeServicios !== null)) d.coste = Math.round(((d.costeEnergia || 0) + (d.costeServicios || 0)) * 100) / 100;
    d.emision = emision;
    return d;
  }

  // Librerías bajo demanda (incluidas en el repo): Chart.js para gráficos, pdf.js para leer facturas
  function loadChart() {
    if (window.Chart) return Promise.resolve();
    if (!chartLoading) {
      chartLoading = new Promise((resolve, reject) => {
        const sc = document.createElement('script');
        sc.src = 'chart.umd.min.js';
        sc.onload = resolve;
        sc.onerror = () => { chartLoading = null; reject(new Error('No se pudo cargar la librería de gráficos.')); };
        document.head.append(sc);
      });
    }
    return chartLoading;
  }
  function loadPdf() {
    if (!pdfLoading) {
      pdfLoading = import('./pdf.min.mjs').then(lib => {
        lib.GlobalWorkerOptions.workerSrc = new URL('pdf.worker.min.mjs', location.href).href;
        return lib;
      });
      pdfLoading.catch(() => { pdfLoading = null; });
    }
    return pdfLoading;
  }
  async function pdfText(file) {
    const lib = await loadPdf();
    const task = lib.getDocument({ data: new Uint8Array(await file.arrayBuffer()), isEvalSupported: false });
    try {
      const doc = await task.promise;
      let out = '';
      for (let i = 1; i <= doc.numPages; i++) {
        const tc = await (await doc.getPage(i)).getTextContent();
        for (const it of tc.items) out += it.str + (it.hasEOL ? '\n' : ' ');
        out += '\n';
      }
      return out;
    } finally {
      task.destroy();
    }
  }

  // Ficheros en la carpeta de la app de OneDrive
  const graphPath = path => 'https://graph.microsoft.com/v1.0/me/drive/special/approot:/' + path.split('/').map(encodeURIComponent).join('/');
  async function odPutFile(path, blob) {
    const token = await odToken();
    const r = await graphFetch(graphPath(path) + ':/content', { method: 'PUT', headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/pdf' }, body: blob });
    if (r.status === 401) throw new OdError('login', 'La sesión de OneDrive ha caducado. Vuelve a conectar.');
    if (!r.ok) throw new OdError('http', 'OneDrive no aceptó el fichero (' + r.status + ').');
  }
  async function odOpenFile(path) {
    const w = window.open('', '_blank'); // se abre ya, con el clic, para que el navegador no la bloquee
    try {
      const token = await odToken();
      const r = await graphFetch(graphPath(path), { headers: { Authorization: 'Bearer ' + token } });
      if (r.status === 404) throw new OdError('http', 'La factura ya no está en OneDrive.');
      if (!r.ok) throw new OdError('http', 'No se pudo abrir la factura (' + r.status + ').');
      const url = (await r.json())['@microsoft.graph.downloadUrl'];
      const blob = await (await graphFetch(url)).blob();
      const obj = URL.createObjectURL(new Blob([blob], { type: 'application/pdf' }));
      if (w) w.location.href = obj; else location.href = obj;
      setTimeout(() => URL.revokeObjectURL(obj), 120000);
    } catch (e) {
      if (w) w.close();
      if (e.kind === 'login') toast(e.message, 'error', { label: 'Conectar', fn: odConnect }, 10000);
      else toast(e.message || 'No se pudo abrir la factura.', 'error');
    }
  }

  function pickFacturas() {
    const input = h('input', { type: 'file', accept: 'application/pdf,.pdf', multiple: true, hidden: true });
    input.addEventListener('change', () => { const fs = [...input.files]; input.remove(); if (fs.length) uploadFacturas(fs); });
    document.body.append(input);
    input.click();
  }

  async function uploadFacturas(files) {
    const nuevas = [], actualizadas = [], fallos = [];
    let sinPdf = 0;
    toast(files.length === 1 ? 'Leyendo la factura…' : 'Leyendo ' + files.length + ' facturas…');
    for (const file of files) {
      try {
        if (file.size > 4 * 1024 * 1024) throw new Error('ocupa más de 4 MB');
        const d = parseFactura(await pdfText(file));
        if (!d || !d.fin) throw new Error('no se reconoce el formato');
        const id = d.tipo + '-' + d.fin;
        const prev = (S.datos.colecciones.consumos || {})[id];
        const rec = Object.assign({ notas: '', factura: null }, prev, { id, tipo: d.tipo, anio: +d.fin.slice(0, 4), mes: +d.fin.slice(5, 7) });
        for (const [k, v] of Object.entries(d)) if (k !== 'tipo' && v !== null && v !== undefined) rec[k] = v;
        if (odConnected()) {
          const path = 'facturas/' + rec.anio + '/' + d.tipo + '_' + d.fin + '.pdf';
          try { await odPutFile(path, file); rec.factura = path; }
          catch (e) { sinPdf++; if (e.kind === 'login') S.odState = 'login'; }
        } else sinPdf++;
        S.datos.colecciones.consumos = Object.assign({}, S.datos.colecciones.consumos, { [id]: rec });
        (prev ? actualizadas : nuevas).push(rec);
      } catch (e) {
        fallos.push(file.name + ' (' + (e.message || 'error') + ')');
      }
    }
    if (nuevas.length || actualizadas.length) saveDatos();
    // Si todas las facturas son del mismo tipo, se muestra ese tipo
    const tipos = new Set(nuevas.concat(actualizadas).map(r => r.tipo));
    if (tipos.size === 1) cTipo = [...tipos][0];
    renderConsumos();
    const partes = [];
    if (nuevas.length) partes.push(nuevas.length + (nuevas.length === 1 ? ' factura añadida' : ' facturas añadidas'));
    if (actualizadas.length) partes.push(actualizadas.length + (actualizadas.length === 1 ? ' actualizada' : ' actualizadas'));
    if (partes.length) toast(cap(partes.join(', ')) + '.');
    if (sinPdf) toast(odConnected() ? 'No se pudieron guardar ' + sinPdf + ' PDF en OneDrive; los datos sí se han guardado.' : 'Datos guardados. Conecta OneDrive para guardar también los PDF.', 'error', null, 8000);
    if (fallos.length) toast('No se pudieron leer: ' + fallos.slice(0, 3).join(', ') + (fallos.length > 3 ? ' y ' + (fallos.length - 3) + ' más' : '') + '.', 'error', null, 10000);
  }

  // Próxima factura: fecha por la duración habitual del periodo; importe por el mismo periodo del año
  // pasado y por la media de las tres últimas
  function estimate(list) {
    const dated = list.filter(c => isYmd(c.fin)).sort((a, b) => a.fin.localeCompare(b.fin));
    if (dated.length < 2) return null;
    const last = dated[dated.length - 1];
    const gaps = dated.slice(-6).map((c, i, a) => i ? (utcDay(c.fin) - utcDay(a[i - 1].fin)) / DAY : null).filter(Boolean).sort((a, b) => a - b);
    const gap = gaps[Math.floor(gaps.length / 2)];
    const next = utcDay(last.fin) + gap * DAY;
    const yearAgo = next - 365 * DAY;
    const prevYear = dated.filter(c => c.coste !== null && Math.abs(utcDay(c.fin) - yearAgo) <= Math.min(25, gap / 2) * DAY)
      .sort((a, b) => Math.abs(utcDay(a.fin) - yearAgo) - Math.abs(utcDay(b.fin) - yearAgo))[0] || null;
    const last3 = dated.filter(c => c.coste !== null).slice(-3);
    const avg = last3.length ? last3.reduce((a, c) => a + c.coste, 0) / last3.length : null;
    const value = prevYear && avg !== null ? (prevYear.coste + avg) / 2 : prevYear ? prevYear.coste : avg;
    return { next, gap, prevYear, avg, n: last3.length, value };
  }

  function renderConsumos() {
    const root = document.getElementById('viewConsumos');
    root.dataset.day = String(todayUtc());
    const all = consumos();
    const byTipo = t => all.filter(c => c.tipo === t);
    const latest = t => byTipo(t).filter(c => isYmd(c.fin)).sort((a, b) => b.fin.localeCompare(a.fin))[0] || null;

    const cards = h('div', { class: 'stats three picks' }, Object.entries(TIPOS).map(([k, T]) => {
      const l = latest(k);
      return h('button', { type: 'button', class: 'stat pick' + (cTipo === k ? ' on' : ''), style: '--c:' + T.color, 'aria-pressed': cTipo === k ? 'true' : 'false', onclick: () => { cTipo = k; renderConsumos(); } },
        h('span', { class: 'stat-label' }, T.nombre),
        h('span', { class: 'stat-value' }, l && l.coste !== null ? fmtEur(l.coste) : '—'),
        h('span', { class: 'stat-sub' }, l ? cap(MESES_LARGOS[l.mes - 1]) + ' ' + l.anio + (l[T.campo] !== null && l[T.campo] !== undefined ? ' · ' + fmtCant(l[T.campo], T.unidad) : '') : 'sin datos'));
    }));

    const T = TIPOS[cTipo];
    const list = byTipo(cTipo);
    const est = estimate(list);
    const estPanel = est && est.value !== null ? h('section', { class: 'panel estimate' },
      h('div', { class: 'est-main' },
        h('span', { class: 'stat-label' }, 'Próxima factura (estimada)'),
        h('span', { class: 'est-value' }, '≈ ' + fmtEur(est.value)),
        h('span', { class: 'stat-sub' }, 'Consumo hasta ~' + fmtDay(est.next).replace(/ \d{4}$/, '') + ' · periodo habitual de ' + Math.round(est.gap) + ' días')),
      h('div', { class: 'est-parts' },
        h('span', null, 'Mismo periodo del año pasado: ', h('strong', null, est.prevYear ? fmtEur(est.prevYear.coste) : '—')),
        h('span', null, 'Media de las ' + est.n + ' últimas: ', h('strong', null, est.avg !== null ? fmtEur(est.avg) : '—')))) : null;

    // Gráfico: este año frente al anterior, por mes
    const y = new Date(todayUtc()).getUTCFullYear();
    const field = cMetric === 'coste' ? 'coste' : T.campo;
    const serie = yr => Array.from({ length: 12 }, (_, m) => {
      const vals = list.filter(c => c.anio === yr && c.mes === m + 1 && c[field] !== null && c[field] !== undefined).map(c => c[field]);
      return vals.length ? Math.round(vals.reduce((a, b) => a + b, 0) * 100) / 100 : null;
    });
    const cur = serie(y), prev = serie(y - 1);
    const hasData = cur.some(v => v !== null) || prev.some(v => v !== null);
    const metricSeg = h('div', { class: 'seg', role: 'group', 'aria-label': 'Ver' },
      [['coste', 'Coste'], ['consumo', 'Consumo']].map(([k, label]) => h('button', { type: 'button', class: cMetric === k ? 'on' : '', 'aria-pressed': cMetric === k ? 'true' : 'false', onclick: () => { cMetric = k; renderConsumos(); } }, label)));
    const canvas = h('canvas', { 'aria-label': 'Gráfico de ' + T.nombre.toLowerCase() + ': ' + y + ' frente a ' + (y - 1) });
    const chartPanel = h('section', { class: 'panel' },
      h('div', { class: 'panel-head' }, h('h2', null, T.nombre + ' · ' + (y - 1) + ' y ' + y),
        h('div', { class: 'legend' }, h('span', { class: 'lg prev' }, String(y - 1)), h('span', { class: 'lg cur', style: '--c:' + T.color }, String(y))),
        metricSeg),
      hasData ? h('div', { class: 'chart-box' }, canvas) : h('p', { class: 'muted empty-note' }, cMetric === 'coste' ? 'No hay importes de estos dos años.' : 'No hay datos de consumo de estos dos años.'));

    // Historial
    const rows = h('div', { class: 'ev-list' });
    const sorted = list.slice().sort((a, b) => (b.fin || '').localeCompare(a.fin || '') || b.anio - a.anio || b.mes - a.mes);
    for (const c of sorted) {
      const periodo = isYmd(c.inicio) && isYmd(c.fin)
        ? fmtDay(utcDay(c.inicio)).replace(/ \d{4}$/, '') + ' → ' + fmtDay(utcDay(c.fin)) + ' · ' + Math.round((utcDay(c.fin) - utcDay(c.inicio)) / DAY) + ' días'
        : isYmd(c.fin) ? 'Hasta el ' + fmtDay(utcDay(c.fin)) : '';
      const detalle = [];
      if (c[T.campo] !== null && c[T.campo] !== undefined) detalle.push(fmtCant(c[T.campo], T.unidad));
      const precio = cTipo === 'agua' ? c.precioM3 : c.precioKwh;
      if (precio) detalle.push(precio.toLocaleString('es-ES', { maximumFractionDigits: 6 }) + ' €/' + T.unidad);
      rows.append(h('div', { class: 'ev consumo' },
        h('div', { class: 'ev-main' },
          h('span', { class: 'ev-name' }, cap(MESES_LARGOS[c.mes - 1]) + ' ' + c.anio),
          h('span', { class: 'ev-what' }, [periodo].concat(detalle).filter(Boolean).join(' · ') || '—')),
        h('span', { class: 'ev-amount' }, c.coste !== null && c.coste !== undefined ? fmtEur(c.coste) : '—'),
        h('div', { class: 'row-tools' },
          c.factura ? h('button', { type: 'button', class: 'icon-btn mini', 'aria-label': 'Ver factura', title: 'Ver factura', onclick: () => { if (!odConnected()) return toast('Conecta OneDrive para ver la factura.', 'error'); odOpenFile(c.factura); } }, ICON.file()) : null,
          h('button', { type: 'button', class: 'icon-btn mini', 'aria-label': 'Editar', title: 'Editar', onclick: () => openConsumo(c.id) }, ICON.edit()),
          h('button', {
            type: 'button', class: 'icon-btn mini', 'aria-label': 'Borrar', title: 'Borrar',
            onclick: () => {
              if (!confirm('¿Borrar el registro de ' + T.nombre.toLowerCase() + ' de ' + MESES_LARGOS[c.mes - 1] + ' de ' + c.anio + '?' + (c.factura ? ' El PDF se queda en OneDrive.' : ''))) return;
              const rest = Object.assign({}, S.datos.colecciones.consumos);
              delete rest[c.id];
              S.datos.colecciones.consumos = rest;
              saveDatos();
              renderConsumos();
            }
          }, ICON.trash()))));
    }
    if (!sorted.length) rows.append(h('p', { class: 'muted empty-note' }, 'Sube tus facturas en PDF o añade un consumo a mano.'));

    const toolbar = h('div', { class: 'g-toolbar' }, h('span'),
      h('div', { class: 'row-actions' },
        h('button', { type: 'button', class: 'btn ghost small', onclick: () => openConsumo(null) }, ICON.plus(), 'Añadir a mano'),
        h('button', { type: 'button', class: 'btn ghost small', onclick: pickFacturas }, ICON.file(), 'Subir facturas')));

    root.replaceChildren(...[cards, estPanel, chartPanel,
      h('section', { class: 'panel' }, h('div', { class: 'panel-head' }, h('h2', null, 'Historial · ' + T.nombre), toolbar), rows)].filter(Boolean));

    if (chartObj) { chartObj.destroy(); chartObj = null; }
    if (hasData) {
      loadChart().then(() => {
        if (!canvas.isConnected) return;
        const unit = cMetric === 'coste' ? '€' : T.unidad;
        const fmt = v => cMetric === 'coste' ? fmtEur(v) : fmtCant(v, unit);
        chartObj = new Chart(canvas, {
          type: 'bar',
          data: { labels: MESES_CORTOS, datasets: [
            { label: String(y - 1), data: prev, backgroundColor: 'rgba(154, 164, 181, 0.32)', borderRadius: 4 },
            { label: String(y), data: cur, backgroundColor: T.color, borderRadius: 4 }
          ] },
          options: {
            responsive: true, maintainAspectRatio: false, animation: false,
            interaction: { mode: 'index', intersect: false },
            plugins: { legend: { display: false }, tooltip: { callbacks: { label: ctx => ctx.dataset.label + ': ' + (ctx.raw === null ? '—' : fmt(ctx.raw)) } } },
            scales: {
              x: { ticks: { color: '#8A94A6', font: { family: 'Manrope', size: 11 } }, grid: { display: false }, border: { display: false } },
              y: { beginAtZero: true, ticks: { color: '#8A94A6', font: { family: 'Manrope', size: 11 }, maxTicksLimit: 5, callback: v => cMetric === 'coste' ? v + ' €' : v }, grid: { color: 'rgba(255, 255, 255, 0.06)' }, border: { display: false } }
            }
          }
        });
      }).catch(e => toast(e.message, 'error'));
    }
  }

  function openConsumo(id) {
    const cur = id ? (S.datos.colecciones.consumos || {})[id] : null;
    const st = { tipo: cur ? cur.tipo : cTipo };
    const inp = (idAttr, attrs) => h('input', Object.assign({ class: 'input', id: idAttr, autocomplete: 'off' }, attrs));
    const fin = inp('cFin', { type: 'date', value: cur && isYmd(cur.fin) ? cur.fin : '' });
    const inicio = inp('cInicio', { type: 'date', value: cur && isYmd(cur.inicio) ? cur.inicio : '' });
    const coste = inp('cCoste', { type: 'text', inputmode: 'decimal', placeholder: '0,00', value: cur && cur.coste !== null && cur.coste !== undefined ? fmtNumInput(cur.coste) : '' });
    const cantidad = inp('cCantidad', { type: 'text', inputmode: 'decimal', placeholder: 'Opcional' });
    const notas = inp('cNotas', { type: 'text', maxlength: '200', placeholder: 'Opcional', value: cur ? cur.notas || '' : '' });
    const cantLabel = h('label', { class: 'field-label', for: 'cCantidad' });
    const setCant = () => {
      const T = TIPOS[st.tipo];
      cantLabel.textContent = 'Consumo (' + T.unidad + ')';
      const v = cur && cur.tipo === st.tipo ? cur[T.campo] : null;
      cantidad.value = v !== null && v !== undefined ? String(v).replace('.', ',') : '';
    };
    const seg = h('div', { class: 'seg' });
    const drawSeg = () => seg.replaceChildren(...Object.entries(TIPOS).map(([k, T]) => h('button', {
      type: 'button', class: st.tipo === k ? 'on' : '', 'aria-pressed': st.tipo === k ? 'true' : 'false',
      onclick: () => { st.tipo = k; drawSeg(); setCant(); }
    }, T.nombre)));
    drawSeg(); setCant();
    let close = null;
    const save = () => {
      if (!isYmd(fin.value)) { toast('Pon la fecha en que termina el consumo.', 'error'); fin.focus(); return; }
      if (inicio.value && (!isYmd(inicio.value) || inicio.value >= fin.value)) { toast('El inicio debe ser anterior al fin.', 'error'); inicio.focus(); return; }
      const c = coste.value.trim() ? parseAmount(coste.value) : null;
      const q = cantidad.value.trim() ? parseAmount(cantidad.value) : null;
      if (Number.isNaN(c) || Number.isNaN(q)) { toast('Revisa los números.', 'error'); return; }
      const newId = st.tipo + '-' + fin.value;
      const cols = Object.assign({}, S.datos.colecciones.consumos);
      if (newId !== (cur && cur.id) && cols[newId] && !confirm('Ya hay un registro de ' + TIPOS[st.tipo].nombre.toLowerCase() + ' que termina ese día. ¿Sustituirlo?')) return;
      if (cur && cur.id !== newId) delete cols[cur.id];
      const T = TIPOS[st.tipo];
      const rec = Object.assign({ kwh: null, m3: null, costeEnergia: null, costeServicios: null, precioKwh: null, precioM3: null, factura: null }, cur && cur.tipo === st.tipo ? cur : {}, {
        id: newId, tipo: st.tipo, anio: +fin.value.slice(0, 4), mes: +fin.value.slice(5, 7),
        inicio: inicio.value || null, fin: fin.value, coste: c, notas: notas.value.trim()
      });
      rec[T.campo] = q;
      cols[newId] = rec;
      S.datos.colecciones.consumos = cols;
      saveDatos();
      cTipo = st.tipo;
      close();
      renderConsumos();
    };
    const fld = (label, forId, el) => h('div', { class: 'field' }, h('label', { class: 'field-label', for: forId }, label), el);
    const body = h('div', { class: 'form' },
      h('div', { class: 'field' }, h('span', { class: 'field-label' }, 'Tipo'), seg),
      h('div', { class: 'form-row' }, fld('Inicio del consumo', 'cInicio', inicio), fld('Fin del consumo', 'cFin', fin)),
      h('p', { class: 'hint' }, 'El registro cuenta en el mes en que termina el consumo.'),
      h('div', { class: 'form-row' }, fld('Importe total (€)', 'cCoste', coste), h('div', { class: 'field' }, cantLabel, cantidad)),
      fld('Notas', 'cNotas', notas),
      h('div', { class: 'row-actions modal-foot' },
        h('span', { class: 'spacer' }),
        h('button', { type: 'button', class: 'btn subtle small', onclick: () => close() }, 'Cancelar'),
        h('button', { type: 'button', class: 'btn primary small', onclick: save }, cur ? 'Guardar' : 'Añadir')));
    close = openModal(cur ? 'Editar consumo' : 'Añadir consumo', body);
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
      // Si la sesión de OneDrive caducó, este clic lleva al login de Microsoft
      if (odConnected() && S.odState === 'login') odConnect();
      else syncDatos();
      // Contactos de Google: solo si el permiso sigue vigente (no se sale de la página sin pedirlo)
      if (ggConnected() && ggTokenOk()) syncContactos().then(() => calSync({ notify: S.ui.view === 'calendario' }));
      // Planner: si la sesión del trabajo caducó, este clic lleva a reconectar
      if (plConnected()) { if (S.plState === 'login') plConnect(); else plSync(); }
      doSync();
    });
    document.getElementById('addBtn').addEventListener('click', () => {
      const cfg = VIEWS[S.ui.view];
      if (cfg && cfg.onAdd) return cfg.onAdd();
      const cols = columns();
      if (cols.length) openNew(cols[0].key, true);
    });
    document.querySelectorAll('.nav-item[data-view]').forEach(a => a.addEventListener('click', e => {
      e.preventDefault();
      showView(a.dataset.view);
    }));
    document.getElementById('openSettings').addEventListener('click', () => { if (SMALL.matches) toggleSidebar(false); openSettings(); });
    document.getElementById('projectSelect').addEventListener('change', e => {
      S.filter.clear();
      if (isWork()) {
        S.pl.planId = e.target.value;
        savePl();
        plSync();
        return;
      }
      S.ui.projectId = e.target.value;
      saveUi();
      renderAll();
    });
    document.addEventListener('keydown', e => {
      if (e.key === 'Escape' && S.editing && !document.querySelector('.modal-wrap')) closeDrawer();
    });
    SMALL.addEventListener('change', applyLayout);
    setInterval(() => {
      renderSyncState();
      // Al cambiar de día se recalculan los cumpledías
      const cfg = VIEWS[S.ui.view];
      const el = cfg && cfg.el && document.getElementById(cfg.el);
      if (el && !el.hidden && el.dataset.day !== String(todayUtc())) cfg.render();
    }, 60000);
  }

  function start() {
    bind();
    renderAll();
    ggHandleReturn();
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
