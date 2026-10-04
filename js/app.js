/*
 * Rendu de la feuille « Hypothèses » et gestion des saisies.
 * Dépend de : config.js, formulas.js, format.js
 *
 * API publique (pour brancher d'autres onglets plus tard) :
 *   HYPO.getValues()   → toutes les valeurs (entrées + calculées)
 *   HYPO.setInput(ref, value)
 *   HYPO.reset()
 *   document 'hypotheses:change' (CustomEvent, detail = valeurs)
 */
(function () {
  'use strict';

  var fmt = HYPO.format;
  var inputs = {};      // valeurs saisies (colonne C)
  var values = {};      // valeurs complètes après calcul
  var inputDefs = {};   // ref → définition de ligne (cellules d'entrée)
  var calcFormats = Object.assign({}, HYPO.CALC_FORMATS);

  // ── Défauts & persistance ──────────────────────────────────────────────────
  HYPO.ROWS.forEach(function (row) {
    if (row.type === 'param') {
      if (row.input) inputDefs[row.cell] = row; else calcFormats[row.cell] = row.fmt;
    }
    if (row.type === 'check') { calcFormats[row.c] = row.fmt; calcFormats[row.g] = row.fmt; }
  });

  function defaults() {
    var d = {};
    Object.keys(inputDefs).forEach(function (ref) { d[ref] = inputDefs[ref].def; });
    return d;
  }

  function load() {
    var d = defaults();
    try {
      var saved = JSON.parse(localStorage.getItem(HYPO.STORAGE_KEY) || '{}');
      Object.keys(d).forEach(function (ref) {
        var def = inputDefs[ref], x = saved[ref];
        if (x === undefined) return;
        if (def.options && def.options.indexOf(x) === -1) return;
        if (def.fmt !== 'text' && typeof x !== 'number') return;
        d[ref] = x;
      });
    } catch (e) { /* stockage indisponible : valeurs par défaut */ }
    return d;
  }

  function save() {
    try { localStorage.setItem(HYPO.STORAGE_KEY, JSON.stringify(inputs)); } catch (e) { /* ignoré */ }
  }

  // ── Construction du DOM ────────────────────────────────────────────────────
  function el(tag, cls, text) {
    var e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text !== undefined) e.textContent = text;
    return e;
  }
  function td(cls, text, colspan) {
    var c = el('td', cls, text);
    if (colspan) c.colSpan = colspan;
    return c;
  }
  function calcCell(ref, cls) {
    var c = td('cell ' + cls);
    c.dataset.ref = ref;
    return c;
  }

  function inputCell(row) {
    var c = td('cell input');
    var ctrl;
    if (row.options) {
      ctrl = el('select');
      row.options.forEach(function (o) {
        var opt = el('option', null, String(o));
        opt.value = String(o);
        ctrl.appendChild(opt);
      });
      ctrl.value = String(inputs[row.cell]);
      ctrl.addEventListener('change', function () {
        var raw = ctrl.value;
        setInput(row.cell, row.fmt === 'text' ? raw : Number(raw));
      });
    } else {
      ctrl = el('input');
      ctrl.type = 'text';
      ctrl.inputMode = 'decimal';
      ctrl.autocomplete = 'off';
      ctrl.value = fmt.display(inputs[row.cell], row.fmt);
      ctrl.addEventListener('focus', function () {
        ctrl.value = fmt.editable(inputs[row.cell], row.fmt);
        ctrl.select();
      });
      ctrl.addEventListener('input', function () {
        var x = fmt.parse(ctrl.value, row.fmt);
        ctrl.classList.toggle('invalid', x === null);
        if (x !== null) setInput(row.cell, x);
      });
      ctrl.addEventListener('blur', function () {
        ctrl.classList.remove('invalid');
        ctrl.value = fmt.display(inputs[row.cell], row.fmt);
      });
      ctrl.addEventListener('keydown', function (e) {
        if (e.key === 'Enter') ctrl.blur();
        if (e.key === 'Escape') { ctrl.value = fmt.editable(inputs[row.cell], row.fmt); ctrl.blur(); }
      });
    }
    ctrl.id = 'in-' + row.cell;
    ctrl.dataset.ref = row.cell;
    ctrl.setAttribute('aria-label', row.label);
    c.appendChild(ctrl);
    return c;
  }

  var RENDER = {
    title: function (row, tr) { tr.appendChild(td('title', row.text, 9)); },
    subtitle: function (row, tr) { tr.appendChild(td('subtitle', row.text, 9)); },
    spacer: function (row, tr) { tr.appendChild(td('spacer', '', 9)); },
    section: function (row, tr) { tr.appendChild(td('section', row.text, 9)); },
    info: function (row, tr) {
      var c = td('info', row.text, 9);
      c.style.height = (row.span * 14.5) + 'pt';
      tr.appendChild(c);
    },
    param: function (row, tr) {
      tr.appendChild(td('label', row.label));                                  // A
      tr.appendChild(td());                                                    // B
      tr.appendChild(row.input ? inputCell(row) : calcCell(row.cell, 'calc')); // C
      if (row.conv) {
        tr.appendChild(calcCell(row.conv, 'conv'));                            // D
        tr.appendChild(td('unit', row.unit));                                  // E
      } else {
        tr.appendChild(td('unit', row.unit || ''));                            // D
        tr.appendChild(td());                                                  // E
      }
      tr.appendChild(td());                                                    // F
      tr.appendChild(td('note', row.note || '', 3));                           // G:I
    },
    check: function (row, tr) {
      tr.appendChild(td('label-light', row.label));
      tr.appendChild(td());
      tr.appendChild(calcCell(row.c, 'check'));
      tr.appendChild(td('unit-sm', row.cLabel));
      tr.appendChild(td()); tr.appendChild(td());
      tr.appendChild(calcCell(row.g, 'check'));
      tr.appendChild(td('unit-sm', row.gLabel));
      tr.appendChild(td());
    },
    recapHeader: function (row, tr) {
      tr.appendChild(td()); tr.appendChild(td());
      tr.appendChild(td('recap-head', row.c));
      tr.appendChild(td()); tr.appendChild(td()); tr.appendChild(td());
      tr.appendChild(td('recap-head', row.g));
      tr.appendChild(td()); tr.appendChild(td());
    },
    recap: function (row, tr) {
      var b = row.bold ? ' bold' : '';
      tr.appendChild(td(row.bold ? 'label' : 'label-normal', row.label));
      tr.appendChild(td());
      tr.appendChild(calcCell(row.c, 'recap' + b));
      tr.appendChild(td()); tr.appendChild(td()); tr.appendChild(td());
      tr.appendChild(calcCell(row.g, 'recap' + b));
      tr.appendChild(td()); tr.appendChild(td());
    }
  };

  function render(tbody) {
    tbody.innerHTML = '';
    HYPO.ROWS.forEach(function (row) {
      var tr = el('tr', 'row-' + row.type);
      tr.dataset.row = row.r;
      if (row.h) tr.style.height = row.h + 'pt';
      RENDER[row.type](row, tr);
      tbody.appendChild(tr);
    });
  }

  // ── Calcul & mise à jour ───────────────────────────────────────────────────
  function refresh() {
    values = HYPO.compute(inputs);
    document.querySelectorAll('td[data-ref]').forEach(function (c) {
      var ref = c.dataset.ref, val = values[ref];
      c.textContent = fmt.display(val, calcFormats[ref]);
      c.classList.toggle('error', !!(val && val.error));
    });
    document.dispatchEvent(new CustomEvent('hypotheses:change', { detail: values }));
  }

  function setInput(ref, value) {
    if (!inputDefs[ref]) throw new Error('Cellule non saisissable : ' + ref);
    inputs[ref] = value;
    save();
    refresh();
  }

  function syncControls() {
    Object.keys(inputDefs).forEach(function (ref) {
      var ctrl = document.getElementById('in-' + ref);
      if (!ctrl) return;
      ctrl.value = ctrl.tagName === 'SELECT' ? String(inputs[ref]) : fmt.display(inputs[ref], inputDefs[ref].fmt);
      ctrl.classList.remove('invalid');
    });
  }

  function reset() {
    inputs = defaults();
    save();
    syncControls();
    refresh();
  }

  function exportJson() {
    var data = { feuille: 'Hypothèses', date: new Date().toISOString(), entrees: inputs, resultats: values };
    var blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
    var a = el('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'hypotheses.json';
    document.body.appendChild(a);
    a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 0);
  }

  // ── API publique ───────────────────────────────────────────────────────────
  HYPO.getValues = function () { return Object.assign({}, values); };
  HYPO.setInput = function (ref, value) { setInput(ref, value); syncControls(); };
  HYPO.reset = reset;

  // ── Démarrage ──────────────────────────────────────────────────────────────
  document.addEventListener('DOMContentLoaded', function () {
    inputs = load();
    render(document.getElementById('sheet-body'));
    refresh();
    document.getElementById('btn-reset').addEventListener('click', function () {
      if (confirm('Réinitialiser toutes les saisies aux valeurs par défaut ?')) reset();
    });
    document.getElementById('btn-export').addEventListener('click', exportJson);
    document.getElementById('btn-print').addEventListener('click', function () { window.print(); });
  });
})();
