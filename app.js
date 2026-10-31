/* Mi panel — v1 (Personal · Todoist)
   - Tareas y secciones: Todoist (API v1, endpoint /sync)
   - Etiquetas y colores: solo en este navegador (localStorage)
*/
(() => {
  'use strict';

  const API_URL = 'https://api.todoist.com/api/v1/sync';
  const KEYS = {
    token: 'mp.todoist.token',
    cache: 'mp.todoist.cache',
    tags: 'mp.tags',
    ui: 'mp.ui'
  };
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
    tags: normalizeTags(store.get(KEYS.tags, null)),
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

  const saveTags = () => store.set(KEYS.tags, S.tags);
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
    const count = S.data.items.filter(i => i.project_id === pid && !i.parent_id).length;
    document.getElementById('taskCount').textContent = String(count);
  }

  function renderSyncState() {
    const when = S.data ? relTime(S.data.syncedAt) : 'nunca';
    document.getElementById('syncInfo').textContent = S.syncing ? 'Sincronizando…' : 'Sincronizado ' + when;
    document.getElementById('syncSide').textContent = when;
    const btn = document.getElementById('syncBtn');
    btn.disabled = S.syncing;
    btn.classList.toggle('is-syncing', S.syncing);
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
        h('p', { class: 'muted' }, 'Arrastra por el asa para cambiar el orden. Las etiquetas se guardan en este navegador.'));

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

  // Ajustes
  function openSettings() {
    const fileInput = h('input', { type: 'file', accept: 'application/json,.json', hidden: true });
    fileInput.addEventListener('change', async () => {
      const file = fileInput.files && fileInput.files[0];
      if (!file) return;
      try {
        const parsed = JSON.parse(await file.text());
        const tags = normalizeTags(parsed && parsed.tags ? parsed.tags : parsed);
        if (!tags.list.length) throw new Error('vacío');
        if (!confirm('Se sustituirán tus etiquetas actuales por las del fichero (' + tags.list.length + ' etiquetas). ¿Continuar?')) return;
        S.tags = tags;
        saveTags();
        renderAll();
        toast('Etiquetas importadas');
      } catch (e) {
        toast('El fichero no tiene un formato válido.', 'error');
      } finally {
        fileInput.value = '';
      }
    });

    const exportTags = () => {
      const blob = new Blob([JSON.stringify({ app: 'mi-panel', version: 1, exportedAt: new Date().toISOString(), tags: S.tags }, null, 2)], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const a = h('a', { href: url, download: 'mi-panel-etiquetas.json' });
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
        h('h3', null, 'Copia de las etiquetas'),
        h('p', { class: 'muted' }, 'Tus etiquetas solo están en este navegador. Exporta una copia para llevarlas a otro dispositivo o como respaldo.'),
        h('div', { class: 'row-actions' },
          h('button', { type: 'button', class: 'btn ghost small', onclick: exportTags }, 'Exportar'),
          h('button', { type: 'button', class: 'btn ghost small', onclick: () => fileInput.click() }, 'Importar'),
          fileInput)),
      h('div', { class: 'modal-section' },
        h('h3', null, 'Este navegador'),
        h('p', { class: 'muted' }, 'Borra el token, las etiquetas y la copia local de tareas de este navegador. En Todoist no se borra nada.'),
        h('div', { class: 'row-actions' },
          h('button', {
            type: 'button', class: 'btn danger small',
            onclick: () => {
              if (!confirm('¿Borrar todos los datos de la app en este navegador? Exporta antes tus etiquetas si quieres conservarlas.')) return;
              Object.values(KEYS).forEach(k => localStorage.removeItem(k));
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
    document.getElementById('syncBtn').addEventListener('click', () => doSync());
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
    if (!S.token) openSetup();
    else if (!S.data) doSync({ silent: true });
  }

  if (typeof Sortable === 'undefined') {
    document.body.textContent = 'Falta el fichero Sortable.min.js junto a index.html.';
    return;
  }
  start();
})();
