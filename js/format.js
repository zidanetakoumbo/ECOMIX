/*
 * Formats numériques équivalents aux formats Excel de la feuille (affichage fr-FR).
 *   int       → 0
 *   thousands → #,##0
 *   dec2      → #,##0.00
 *   dec6      → 0.000000
 *   pct1      → 0.0%
 *   pct2      → 0.00%
 *   k         → #,##0" k"
 *   text      → texte brut
 */
window.HYPO = window.HYPO || {};

HYPO.format = (function () {
  var cache = {};
  function nf(decimals, grouping) {
    var key = decimals + '|' + grouping;
    if (!cache[key]) {
      cache[key] = new Intl.NumberFormat('fr-FR', {
        minimumFractionDigits: decimals, maximumFractionDigits: decimals, useGrouping: grouping
      });
    }
    return cache[key];
  }

  var FORMATS = {
    int:       { dec: 0, group: false },
    thousands: { dec: 0, group: true },
    dec2:      { dec: 2, group: true },
    dec6:      { dec: 6, group: false },
    pct1:      { dec: 1, group: true, pct: true },
    pct2:      { dec: 2, group: true, pct: true },
    k:         { dec: 0, group: true, suffix: ' k' }
  };

  /** Valeur → texte affiché dans la cellule. */
  function display(value, fmt) {
    if (value && value.error) return value.error;
    if (value === '' || value === null || value === undefined) return '';
    if (typeof value !== 'number' || !FORMATS[fmt]) return String(value);
    var f = FORMATS[fmt];
    var x = f.pct ? value * 100 : value;
    return nf(f.dec, f.group).format(x) + (f.pct ? '%' : '') + (f.suffix || '');
  }

  /** Valeur → texte éditable dans le champ de saisie (pourcentages saisis en %). */
  function editable(value, fmt) {
    if (typeof value !== 'number') return value == null ? '' : String(value);
    var f = FORMATS[fmt] || {};
    var x = f.pct ? value * 100 : value;
    return String(+x.toPrecision(12)).replace('.', ',');
  }

  /**
   * Texte saisi → valeur stockée. Accepte « 1 000 000 », « 5,5 », « 5.5 », « 5 % ».
   * Renvoie null si la saisie n'est pas un nombre valide.
   */
  function parse(text, fmt) {
    if (fmt === 'text') return String(text).trim();
    var s = String(text).replace(/[\s  %]/g, '').replace(',', '.');
    if (s === '' || !/^[-+]?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?$/i.test(s)) return null;
    var x = parseFloat(s);
    return (FORMATS[fmt] && FORMATS[fmt].pct) ? x / 100 : x;
  }

  return { display: display, editable: editable, parse: parse };
})();
