/*
 * Simulación mínima de los servicios de Apps Script que usa Code.gs
 * (SpreadsheetApp, CalendarApp, DriveApp, Utilities, LockService…).
 * Sirve para las pruebas en Node (test/) y para la vista previa local (dev/preview.js).
 * Los métodos de formato que no importan para la lógica no hacen nada y se pueden encadenar.
 */
(function (root) {
  'use strict';

  /** Envuelve un objeto: cualquier método inexistente es un no-op encadenable. */
  function blando(obj) {
    const proxy = new Proxy(obj, {
      get(o, k) {
        if (k in o) return o[k];
        if (typeof k === 'symbol' || k === 'then' || k === 'toJSON') return undefined;
        return function () { return proxy; };
      },
    });
    return proxy;
  }

  const vacio = (v) => v === '' || v === null || v === undefined;

  function Hoja(nombre, ss) {
    this._nombre = nombre;
    this._ss = ss;
    this._d = [];
    this._gid = ss._sigGid++;
  }
  Hoja.prototype = {
    getName() { return this._nombre; },
    setName(n) { this._nombre = n; return this; },
    getSheetId() { return this._gid; },
    getLastRow() {
      for (let i = this._d.length - 1; i >= 0; i--) if ((this._d[i] || []).some((v) => !vacio(v))) return i + 1;
      return 0;
    },
    getLastColumn() {
      let m = 0;
      this._d.forEach((r) => {
        for (let j = (r || []).length - 1; j >= 0; j--) if (!vacio(r[j])) { m = Math.max(m, j + 1); break; }
      });
      return m;
    },
    getMaxRows() { return Math.max(1000, this._d.length); },
    getMaxColumns() { return 26; },
    getRange(fila, col, nf, nc) {
      if (typeof fila !== 'number' || fila < 1 || col < 1) throw new Error('getRange: coordenadas inválidas ' + fila + ',' + col);
      return blando(new Rango(this, fila, col, nf || 1, nc || 1));
    },
    appendRow(arr) { this._d[this.getLastRow()] = arr.slice(); return this; },
    deleteRow(f) { this._d.splice(f - 1, 1); return this; },
    deleteRows(f, n) { this._d.splice(f - 1, n); return this; },
    getBandings() { return []; },
  };

  function Rango(hoja, fila, col, nf, nc) {
    if (nf < 1 || nc < 1) throw new Error('getRange: el número de filas/columnas debe ser ≥ 1');
    Object.assign(this, { _h: hoja, _f: fila, _c: col, _nf: nf, _nc: nc });
  }
  Rango.prototype = {
    getValues() {
      const out = [];
      for (let i = 0; i < this._nf; i++) {
        const r = this._h._d[this._f - 1 + i] || [];
        const fila = [];
        for (let j = 0; j < this._nc; j++) {
          const v = r[this._c - 1 + j];
          fila.push(vacio(v) ? '' : v);
        }
        out.push(fila);
      }
      return out;
    },
    getValue() { return this.getValues()[0][0]; },
    setValues(v) {
      if (v.length !== this._nf || v.some((r) => r.length !== this._nc)) {
        throw new Error(`setValues: dimensiones ${v.length}x${v[0] && v[0].length} no coinciden con ${this._nf}x${this._nc}`);
      }
      v.forEach((r, i) => {
        const fi = this._f - 1 + i;
        if (!this._h._d[fi]) this._h._d[fi] = [];
        r.forEach((x, j) => { this._h._d[fi][this._c - 1 + j] = x; });
      });
      return this;
    },
    setValue(v) {
      const fi = this._f - 1;
      if (!this._h._d[fi]) this._h._d[fi] = [];
      this._h._d[fi][this._c - 1] = v;
      return this;
    },
    getNumberFormats() {
      return Array.from({ length: this._nf }, () => Array.from({ length: this._nc }, () => 'General'));
    },
    setNumberFormats(f) {
      if (f.length !== this._nf || f.some((r) => r.length !== this._nc)) throw new Error('setNumberFormats: dimensiones no coinciden');
      return this;
    },
  };

  function Libro(nombre) {
    this._nombre = nombre;
    this._hojas = [];
    this._sigGid = 0;
    this._id = 'demo-' + Math.random().toString(36).slice(2, 10);
    this.insertSheet('Hoja 1');
    return blando(this);
  }
  Libro.prototype = {
    getId() { return this._id; },
    getUrl() { return 'https://docs.google.com/spreadsheets/d/' + this._id + '/edit'; },
    getName() { return this._nombre; },
    getSheets() { return this._hojas.slice(); },
    getSheetByName(n) { return this._hojas.find((h) => h.getName() === n) || null; },
    insertSheet(n) {
      if (this.getSheetByName(n)) throw new Error('Ya existe la hoja ' + n);
      const h = blando(new Hoja(n, this));
      this._hojas.push(h);
      return h;
    },
    deleteSheet(h) { this._hojas = this._hojas.filter((x) => x !== h); },
    setActiveSheet(h) { this._activa = h; return h; },
    moveActiveSheet(pos) {
      const i = this._hojas.indexOf(this._activa);
      if (i >= 0) { this._hojas.splice(i, 1); this._hojas.splice(pos - 1, 0, this._activa); }
    },
  };

  /** Fecha → texto con el patrón de Utilities.formatDate (usa la zona local del proceso). */
  function formatDate(d, _tz, patron) {
    const p = (n) => String(n).padStart(2, '0');
    return patron.replace(/'([^']*)'|yyyy|MM|dd|HH|mm|ss/g, (m, lit) => {
      if (lit !== undefined) return lit;
      return { yyyy: d.getFullYear(), MM: p(d.getMonth() + 1), dd: p(d.getDate()), HH: p(d.getHours()), mm: p(d.getMinutes()), ss: p(d.getSeconds()) }[m];
    });
  }

  /** Eventos de ejemplo para cualquier día. */
  function eventosDemo(ini) {
    const base = [
      [8, 0, 60, 'María Fernández - Control', 'Consultorio Escazú'],
      [9, 30, 90, 'Primera vez Laura Jiménez', 'Escazú'],
      [11, 0, 120, 'Ana Solís y Luis Mora · Pareja primera vez', 'Curridabat'],
      [14, 0, 60, 'Control Sofía Rojas Curridabat', ''],
      [16, 0, 60, 'Mediciones - Carlos Vargas', 'Domicilio'],
      [18, 0, 60, 'Gimnasio', ''],
    ];
    return base.map(([h, m, dur, titulo, lugar], i) => {
      const s = new Date(ini.getFullYear(), ini.getMonth(), ini.getDate(), h, m);
      const e = new Date(s.getTime() + dur * 60000);
      return {
        getId: () => 'evt' + i + '@google.com',
        getTitle: () => titulo,
        getLocation: () => lugar,
        getStartTime: () => s,
        getEndTime: () => e,
        isAllDayEvent: () => false,
      };
    });
  }

  function crearEntornoGAS() {
    const libro = new Libro('Bitácora');
    const creados = [];
    const calendario = { getEvents: (ini) => eventosDemo(ini) };
    const props = {};
    const icsPorUrl = {};
    return {
      _libro: libro,
      _creados: creados,
      SpreadsheetApp: blando({
        getActiveSpreadsheet: () => libro,
        create: (n) => { const l = new Libro(n); creados.push(l); return l; },
        flush: () => {},
        BorderStyle: { SOLID: 'SOLID' },
        BandingTheme: { LIGHT_GREEN: 'LIGHT_GREEN' },
      }),
      CalendarApp: {
        getDefaultCalendar: () => calendario,
        getCalendarById: () => calendario,
        getCalendarsByName: () => [calendario],
      },
      DriveApp: blando({
        getFoldersByName: () => ({ hasNext: () => false }),
        createFolder: () => blando({}),
        getFileById: () => blando({}),
      }),
      Utilities: {
        formatDate,
        parseDate: (texto) => {
          const m = /^(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2}):(\d{2})$/.exec(texto);
          return new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +m[6]);
        },
        base64Encode: (s) => (typeof btoa === 'function' ? btoa(unescape(encodeURIComponent(s))) : Buffer.from(s).toString('base64')),
        getUuid: () => 'xxxxxxxx-xxxx'.replace(/x/g, () => Math.floor(Math.random() * 16).toString(16)),
      },
      CacheService: { getScriptCache: () => ({ get: () => null, put: () => {} }) },
      UrlFetchApp: {
        fetch: (url) => {
          const texto = icsPorUrl[url];
          return { getResponseCode: () => (texto ? 200 : 404), getContentText: () => texto || '' };
        },
      },
      _icsPorUrl: icsPorUrl,
      LockService: { getScriptLock: () => ({ waitLock() {}, tryLock: () => true, releaseLock() {} }) },
      PropertiesService: {
        getScriptProperties: () => ({ getProperty: (k) => props[k] || null, setProperty: (k, v) => { props[k] = v; } }),
      },
      ScriptApp: { getService: () => ({ getUrl: () => '' }) },
      HtmlService: blando({}),
    };
  }

  if (typeof module !== 'undefined' && module.exports) module.exports = { crearEntornoGAS };
  else root.crearEntornoGAS = crearEntornoGAS;
})(typeof window !== 'undefined' ? window : globalThis);
