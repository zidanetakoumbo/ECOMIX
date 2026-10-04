/*
 * Formules de la feuille « Hypothèses », transcrites 1:1 depuis Excel.
 * L'ordre des clés est l'ordre d'évaluation (dépendances résolues de haut en bas).
 * Chaque formule reçoit v(ref) qui renvoie la valeur courante d'une cellule.
 * Les erreurs Excel (#DIV/0!, #VALUE!) sont levées via XLError et affichées telles quelles.
 */
window.HYPO = window.HYPO || {};

HYPO.XLError = function (code) { this.code = code; };

HYPO.xl = {
  /** Division Excel : lève #DIV/0! si diviseur nul. */
  div: function (a, b) {
    var x = HYPO.xl.num(a), y = HYPO.xl.num(b);
    if (y === 0) throw new HYPO.XLError('#DIV/0!');
    return x / y;
  },
  /** Coercition numérique Excel : cellule vide = 0, texte non numérique = #VALUE!. */
  num: function (a) {
    if (a === '' || a === null || a === undefined) return 0;
    if (typeof a === 'number') return a;
    throw new HYPO.XLError('#VALUE!');
  },
  /** IFERROR(expr, fallback) */
  iferror: function (fn, fallback) {
    try { return fn(); } catch (e) { if (e instanceof HYPO.XLError) return fallback; throw e; }
  }
};

HYPO.FORMULAS = (function () {
  var xl = HYPO.xl, n = xl.num, div = xl.div;
  return {
    // =IF(C6="FCFA",1,IF(C6="EUR",1/C7,(1/C7)*C8))
    C9: function (v) { return v('C6') === 'FCFA' ? 1 : (v('C6') === 'EUR' ? div(1, v('C7')) : div(1, v('C7')) * n(v('C8'))); },
    // =IF(C6="FCFA","FCFA",IF(C6="EUR","€","$"))
    C10: function (v) { return v('C6') === 'FCFA' ? 'FCFA' : (v('C6') === 'EUR' ? '€' : '$'); },

    // =IFERROR((C18*C20*(1-C21)*C27)/C25,"")
    C30: function (v) { return xl.iferror(function () { return div(n(v('C18')) * n(v('C20')) * (1 - n(v('C21'))) * n(v('C27')), v('C25')); }, ''); },
    // =IFERROR((C18*C20*(1-C21)*C28)/C25,"")
    G30: function (v) { return xl.iferror(function () { return div(n(v('C18')) * n(v('C20')) * (1 - n(v('C21'))) * n(v('C28')), v('C25')); }, ''); },

    // Conversion dans la devise d'affichage : =Cxx*$C$9
    D38: function (v) { return n(v('C38')) * n(v('C9')); },
    D40: function (v) { return n(v('C40')) * n(v('C9')); },
    D45: function (v) { return n(v('C45')) * n(v('C9')); },
    D55: function (v) { return n(v('C55')) * n(v('C9')); },

    // Section 9 — Scénario 1
    C75: function (v) { return n(v('D45')) * n(v('C18')) / 1000; },   // =D45*C18/1000
    C76: function () { return 0; },                                    // =0
    C77: function (v) { return n(v('C75')); },                         // =C75
    C78: function (v) { return n(v('C77')) * n(v('C46')); },           // =C77*C46
    C79: function (v) { return n(v('C77')) - n(v('C78')); },           // =C77-C78

    // Section 9 — Scénario 2
    G75: function (v) { return n(v('D45')) * n(v('C18')) / 1000; },   // =D45*C18/1000
    G76: function (v) { return n(v('D55')) * n(v('C53')) / 1000; },   // =D55*C53/1000
    G77: function (v) { return n(v('G75')) + n(v('G76')); },           // =G75+G76
    G78: function (v) { return n(v('G77')) * n(v('C46')); },           // =G77*C46
    G79: function (v) { return n(v('G77')) - n(v('G78')); }            // =G77-G78
  };
})();

/**
 * Recalcule toutes les formules à partir des saisies.
 * @param {Object} inputs  valeurs des cellules d'entrée { C6: 'FCFA', C7: 655.957, ... }
 * @returns {Object}       toutes les valeurs (entrées + calculs) ; erreurs => { error: '#DIV/0!' }
 */
HYPO.compute = function (inputs) {
  var values = Object.assign({}, inputs);
  var v = function (ref) {
    var x = values[ref];
    if (x && x.error) throw new HYPO.XLError(x.error); // propagation des erreurs comme Excel
    return x === undefined ? '' : x;
  };
  Object.keys(HYPO.FORMULAS).forEach(function (ref) {
    try { values[ref] = HYPO.FORMULAS[ref](v); }
    catch (e) {
      if (!(e instanceof HYPO.XLError)) throw e;
      values[ref] = { error: e.code };
    }
  });
  return values;
};
