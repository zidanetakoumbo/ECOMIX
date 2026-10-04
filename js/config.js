/*
 * Définition de la feuille « Hypothèses » (structure, libellés, valeurs par défaut, formats).
 * Reproduit ligne à ligne la feuille Excel « Modele_Financier_Solaire_3_Scenarios_GP V4.xlsx ».
 *
 * Types de lignes :
 *   title / subtitle  : bandeau d'en-tête (A:I fusionnées)
 *   spacer            : ligne vide (hauteur en pt)
 *   section           : titre de section vert (A:I fusionnées)
 *   param             : libellé (A) + cellule C (saisie ou calcul) + unité + commentaire (G:I)
 *   check             : ligne de vérification (C et G calculées)
 *   info              : bloc de texte explicatif (A:I fusionnées)
 *   recapHeader/recap : tableau récapitulatif section 9 (colonnes C et G)
 *
 * Formats (cf. js/format.js) : text, int, thousands, dec2, dec6, pct1, pct2, k
 * Toute cellule « input » est une entrée utilisateur (case jaune, colonne C).
 */
window.HYPO = window.HYPO || {};

HYPO.STORAGE_KEY = 'ecomix.hypotheses.v1';

HYPO.ROWS = [
  { r: 1, type: 'title', h: 30, text: 'HYPOTHÈSES — TECHNIQUE, CONSOMMATION DU SITE & FINANCEMENT' },
  { r: 2, type: 'subtitle', h: 18, text: 'Green Power Technologie — Autoconsommation solaire avec revente du surplus' },
  { r: 3, type: 'spacer', h: 6 },
  { r: 4, type: 'spacer' },

  // ── 0. DEVISE ─────────────────────────────────────────────────────────────
  { r: 5, type: 'section', text: "0. DEVISE D'AFFICHAGE — CONVERSION AUTOMATIQUE DU CLASSEUR" },
  { r: 6, type: 'param', h: 50.15, label: "Devise d'affichage (liste déroulante)",
    cell: 'C6', input: true, fmt: 'text', options: ['EUR', 'USD', 'FCFA'], def: 'FCFA',
    note: "Convertit tout le classeur (Hypothèses, Tarifs, Scénarios 1/2/3, Synthèse). Saisissez vos coûts et tarifs en FCFA (case jaune, devise de référence du classeur) ; la case adjacente affiche l'équivalent dans la devise choisie ici." },
  { r: 7, type: 'param', h: 28, label: 'Parité FCFA/EUR (fixe, BCEAO/BEAC)',
    cell: 'C7', input: true, fmt: 'dec2', def: 655.957, unit: 'FCFA/EUR',
    note: "Parité fixe du Franc CFA (XOF/BCEAO et XAF/BEAC) à l'euro depuis 1999. Ne pas modifier sauf changement de régime monétaire." },
  { r: 8, type: 'param', h: 39, label: 'Taux de change EUR → USD',
    cell: 'C8', input: true, fmt: 'dec2', def: 1.08, unit: 'USD/EUR',
    note: "Taux indicatif modifiable ; mettez à jour selon le taux de change en vigueur à la date de votre analyse (le FCFA n'étant pas directement parifié au USD, la conversion FCFA→USD passe par l'EUR)." },
  { r: 9, type: 'param', h: 15.65, label: 'Facteur de conversion appliqué depuis le FCFA (informatif)',
    cell: 'C9', fmt: 'dec6' },
  { r: 10, type: 'param', h: 15.65, label: 'Symbole / code devise affiché',
    cell: 'C10', fmt: 'text' },
  { r: 11, type: 'info', span: 2,
    text: "💱 Le facteur de conversion et le symbole ci-dessus se recalculent automatiquement selon la devise choisie ; ils alimentent l'ensemble des cellules monétaires du classeur (Hypothèses, Tarifs, Scénario 1, Scénario 2, Scénario 3, Synthèse)." },

  // ── 1. DURÉE & ACTUALISATION ──────────────────────────────────────────────
  { r: 13, type: 'section', text: '1. DURÉE DU PROJET & ACTUALISATION — PARAMÈTRE PILOTE' },
  { r: 14, type: 'param', h: 39, label: 'Durée du projet retenue (liste déroulante)',
    cell: 'C14', input: true, fmt: 'int', options: [10, 15, 20, 25, 30], def: 15, unit: 'ans',
    note: "⚑ Pilote l'ensemble des calculs. Choisissez 20 ans (durée mini de contrat), 25 ans (durée de vie standard des modules) ou 30 ans (exploitation prolongée)." },
  { r: 15, type: 'param', h: 50.15, label: "Taux d'actualisation projet — WACC nominal",
    cell: 'C15', input: true, fmt: 'pct2', def: 0.05, unit: '%/an',
    note: "Moyenne pondérée du coût de la dette après IS et du coût des fonds propres, pondérée par le gearing cible. Repère Afrique de l'Ouest/Centrale (IRENA/CATF) : 8 à 17% selon le pays et le profil de risque." },
  { r: 16, type: 'spacer' },

  // ── 2. CENTRALE PV ────────────────────────────────────────────────────────
  { r: 17, type: 'section', text: '2. CENTRALE SOLAIRE PV — PARAMÈTRES TECHNIQUES' },
  { r: 18, type: 'param', h: 39, label: 'Puissance installée DC',
    cell: 'C18', input: true, fmt: 'thousands', def: 500, unit: 'kWc',
    note: "Puissance crête DC totale prévue au projet, issue de l'étude de dimensionnement ou du cahier des charges EPC (bordereau de chiffrage équipements)." },
  { r: 19, type: 'param', h: 39, label: 'Ratio DC/AC (informatif)',
    cell: 'C19', input: true, fmt: 'dec2', def: 1.2, unit: 'x',
    note: "Surdimensionnement DC/AC retenu par le bureau d'études ; valeur usuelle 1,10-1,30 pour optimiser le productible sans écrêtage excessif." },
  { r: 20, type: 'param', h: 39, label: 'Productible spécifique — année 1',
    cell: 'C20', input: true, fmt: 'thousands', def: 1650, unit: 'kWh/kWc',
    note: "Issu d'une étude d'irradiation (PVsyst/SAM) réalisée pour le site exact du projet ; irradiation POA × Performance Ratio. Repère Afrique de l'Ouest ensoleillée : 1500-1750 kWh/kWc/an." },
  { r: 21, type: 'param', h: 39, label: 'Pertes de première année (LID)',
    cell: 'C21', input: true, fmt: 'pct1', def: 0, unit: '%',
    note: "Pertes de dégradation initiale (Light-Induced Degradation), indiquées sur la fiche technique du fabricant de modules (souvent 1-2%)." },
  { r: 22, type: 'param', h: 28, label: 'Dégradation annuelle (après année 1)',
    cell: 'C22', input: true, fmt: 'pct2', def: 0.001, unit: '%/an',
    note: 'Taux garanti par le fabricant (garantie de performance linéaire de la fiche technique), typiquement 0,3-0,55%/an.' },
  { r: 23, type: 'spacer' },

  // ── 3. CONSOMMATION ───────────────────────────────────────────────────────
  { r: 24, type: 'section', text: '3. CONSOMMATION DU SITE & AUTOCONSOMMATION' },
  { r: 25, type: 'param', h: 39, label: 'Consommation annuelle du site — année 1',
    cell: 'C25', input: true, fmt: 'thousands', def: 1000000, unit: 'kWh',
    note: "À établir à partir des factures d'électricité réelles du client sur 12 mois (somme des kWh facturés). C'est la donnée la plus fiable et la plus importante à collecter avant tout chiffrage." },
  { r: 26, type: 'param', h: 39, label: 'Taux de croissance annuel de la consommation du site',
    cell: 'C26', input: true, fmt: 'pct1', def: 0.01, unit: '%/an',
    note: "Anticipation de la croissance d'activité du client (extension, nouvelles lignes de production...) ; à défaut d'information, retenir une hypothèse prudente 0-2%/an." },
  { r: 27, type: 'param', h: 105, label: "Taux d'autoconsommation DIRECTE (immédiate, sans stockage) — commun Scénarios 1 et 2",
    cell: 'C27', input: true, fmt: 'pct1', def: 0.6, unit: '% de la production',
    note: "Part de la production PV consommée INSTANTANÉMENT par le site, au moment où elle est produite (journée), sans transiter par une batterie. Dépend de la correspondance entre la courbe de production (cloche solaire) et la courbe de charge du site. Repère pré-faisabilité pour un site tertiaire/industriel actif en journée : 35-55% ; à affiner par une simulation horaire (8760h) du profil de charge réel. Cette part est identique pour les Scénarios 1 et 2 (même site, même PV) : seul le Scénario 2 y ajoute l'autoconsommation via batterie ci-dessous." },
  { r: 28, type: 'param', h: 83.15, label: "Taux d'autoconsommation TOTALE — Scénario 2 (directe + restituée par la batterie)",
    cell: 'C28', input: true, fmt: 'pct1', def: 0.38, unit: '% de la production',
    note: "Part de la production totale finalement autoconsommée une fois la batterie ajoutée (directe + énergie stockée en journée puis restituée au site, en priorité pour substituer le groupe électrogène — cf. section 6). Toujours ≥ au taux direct ci-dessus. Repère pré-faisabilité : 65-85% selon le dimensionnement de la batterie vis-à-vis du profil de charge ; à confirmer par la même simulation horaire." },
  { r: 29, type: 'spacer' },
  { r: 30, type: 'check', h: 15.65, label: 'Vérification — Autoconsommation implicite vs consommation du site (année 1)',
    c: 'C30', cLabel: 'Scénario 1', g: 'G30', gLabel: 'Scénario 2', fmt: 'pct1' },
  { r: 31, type: 'info', span: 2,
    text: "Ce ratio doit rester ≤ 100% (l'autoconsommation ne peut pas dépasser la consommation réelle du site à aucun moment) : s'il dépasse largement 60-70% en base annuelle, le taux d'autoconsommation saisi ci-dessus est probablement trop optimiste au regard de la taille du champ PV par rapport à la consommation du site — à corriger avant de poursuivre le chiffrage." },
  { r: 33, type: 'spacer' },

  // ── 3bis. SURPLUS ─────────────────────────────────────────────────────────
  { r: 34, type: 'section', text: "3bis. RÉPARTITION DU SURPLUS — L'ÉNERGIE PRODUITE N'A PAS D'OBLIGATION D'INJECTION AU RÉSEAU" },
  { r: 35, type: 'param', h: 129, label: "Taux d'energie solaire injecte sur le reseau",
    cell: 'C35', input: true, fmt: 'pct1', def: 0.05, unit: '% du surplus',
    note: "⚑ Une fois l'autoconsommation directe (et, Scénario 2, le stockage batterie) déduits, le reliquat de production N'A PAS À ÊTRE injecté au réseau : ce taux fixe la part du surplus réellement exportée et vendue au tarif de rachat. Mettez 100% si tout le surplus est injecté (cas standard avec contrat d'injection), ou une valeur inférieure si le site choisit de plafonner/écrêter l'injection (absence d'accord réseau, limite technique, choix stratégique de ne pas dépendre du rachat) : le reste est alors écrêté, sans coût ni revenu, ni pour le Scénario 1 ni pour le Scénario 2." },
  { r: 36, type: 'spacer' },

  // ── 4. OPEX & AMORTISSEMENT ───────────────────────────────────────────────
  { r: 37, type: 'section', text: '4. OPEX SOLAIRE & AMORTISSEMENT FISCAL' },
  { r: 38, type: 'param', h: 39, label: 'OPEX solaire initial (O&M, assurance, monitoring)',
    cell: 'C38', input: true, fmt: 'thousands', def: 5000, conv: 'D38', unit: '/kWc/an',
    note: "À chiffrer à partir de l'offre du prestataire O&M, de l'assurance dommages/RC et du monitoring à distance. Repère marché Afrique de l'Ouest : 1-2% du CAPEX solaire par an." },
  { r: 39, type: 'param', h: 28, label: 'Indexation OPEX',
    cell: 'C39', input: true, fmt: 'pct1', def: 0.01, unit: '%/an',
    note: "Indexation contractuelle de l'OPEX (contrat O&M) ; à défaut de contrat signé, retenir l'inflation prévisionnelle locale." },
  { r: 40, type: 'param', h: 39, label: 'Provision démantèlement (annuelle, provisionnée)',
    cell: 'C40', input: true, fmt: 'thousands', def: 0, conv: 'D40', unit: '/kWc/an',
    note: "Rarement exigée pour une installation en toiture C&I (à la différence d'un parc au sol) ; mettre à 0 si non applicable, ou estimer via un devis de dépose/recyclage." },
  { r: 41, type: 'param', h: 39, label: "Durée d'amortissement fiscal solaire",
    cell: 'C41', input: true, fmt: 'int', def: 15, unit: 'ans',
    note: "Durée d'amortissement comptable retenue par votre expert-comptable pour les installations de production d'électricité (souvent 20 ans, parfois accéléré selon le régime fiscal local)." },
  { r: 42, type: 'param', label: 'Senario A,B ou C',
    cell: 'C42', input: true, fmt: 'text', options: ['A', 'B', 'C'], def: 'A' },
  { r: 43, type: 'spacer' },

  // ── 5. FINANCEMENT ────────────────────────────────────────────────────────
  { r: 44, type: 'section', text: '5. FINANCEMENT — DETTE & FONDS PROPRES' },
  { r: 45, type: 'param', h: 50.15, label: 'CAPEX solaire (équipements + installation, cf. chiffrage détaillé)',
    cell: 'C45', input: true, fmt: 'thousands', def: 500000, conv: 'D45', unit: '/kWc',
    note: "À reprendre du bordereau de chiffrage détaillé (modules, structures, câbles, onduleurs, protections, local technique, main d'œuvre). Repère marché Afrique de l'Ouest 2025-2026 : ≈ 400 000-550 000 FCFA/kWc pour une installation C&I clé en main." },
  { r: 46, type: 'param', h: 39, label: 'Gearing — quote-part dette senior',
    cell: 'C46', input: true, fmt: 'pct1', def: 0.8, unit: '% du CAPEX',
    note: "Quote-part de dette acceptée par les prêteurs, déterminée par le DSCR minimum exigé (souvent 1,20-1,35x) ; à confirmer avec votre term sheet bancaire local." },
  { r: 47, type: 'param', h: 39, label: 'Coût de la dette senior',
    cell: 'C47', input: true, fmt: 'pct2', def: 0.05, unit: '%/an',
    note: "Taux proposé par les prêteurs locaux (term sheet bancaire), fonction du profil de risque et de la maturité de la dette ; repère marché bancaire UEMOA/CEMAC 6-10%." },
  { r: 48, type: 'param', h: 39, label: 'Durée de la dette (tenor), annuité constante',
    cell: 'C48', input: true, fmt: 'int', def: 6, unit: 'ans',
    note: "Durée du prêt proposée par les prêteurs, généralement inférieure à la durée de vie du projet (marge de sécurité en fin de contrat)." },
  { r: 49, type: 'param', h: 28, label: 'Coût des fonds propres exigé',
    cell: 'C49', input: true, fmt: 'pct2', def: 0.15, unit: '%/an',
    note: "Rendement minimum exigé par les investisseurs (hurdle rate) sur ce marché, fonction du profil de risque du projet et du pays." },
  { r: 50, type: 'param', h: 39, label: "Taux d'Impôt sur les Sociétés (IS)",
    cell: 'C50', input: true, fmt: 'pct1', def: 0.3, unit: '%',
    note: "Taux d'IS applicable dans le pays d'implantation du projet (ex. 30% au Sénégal), à adapter selon la fiscalité locale ou un régime incitatif éventuel (code des investissements)." },
  { r: 51, type: 'spacer' },

  // ── 6. BATTERIE ───────────────────────────────────────────────────────────
  { r: 52, type: 'section', text: '6. STOCKAGE BATTERIE (BESS) — SCÉNARIO 2 UNIQUEMENT' },
  { r: 53, type: 'param', h: 39, label: 'Capacité utile de stockage',
    cell: 'C53', input: true, fmt: 'thousands', def: 723, unit: 'kWh',
    note: "Capacité utile (après réservation DoD) issue du dimensionnement technique, à faire valider par l'intégrateur BESS au regard du profil de charge du site." },
  { r: 54, type: 'param', h: 39, label: 'Puissance de la batterie',
    cell: 'C54', input: true, fmt: 'thousands', def: 324, unit: 'kW',
    note: "Associée à sa durée de décharge (ex. 2h) ; à définir avec l'intégrateur selon l'usage visé (décalage de charge, écrêtage de pointe). Technologie LFP (Lithium Fer Phosphate) usuelle." },
  { r: 55, type: 'param', h: 39, label: 'CAPEX batterie clé en main (PCS inclus)',
    cell: 'C55', input: true, fmt: 'thousands', def: 108880, conv: 'D55', unit: '/kWh',
    note: "À demander à votre intégrateur BESS (offre clé en main incluant PCS, conteneur/armoire, installation). Repère marché : BESS turnkey ≈ 150 000-200 000 FCFA/kWh (2025-2026)." },
  { r: 56, type: 'param', h: 28, label: 'OPEX batterie (O&M, monitoring, assurance)',
    cell: 'C56', input: true, fmt: 'pct1', def: 0.02, unit: '% CAPEX/an',
    note: "À demander au fournisseur BESS (contrat O&M/monitoring dédié), généralement 1,5-3%/an du CAPEX initial." },
  { r: 57, type: 'param', h: 28, label: 'Rendement aller-retour (RTE)',
    cell: 'C57', input: true, fmt: 'pct1', def: 0.95, unit: '%',
    note: "Rendement garanti par le fabricant (fiche technique), pertes onduleur/PCS incluses ; typiquement 88-95% pour du LFP récent." },
  { r: 58, type: 'param', h: 39, label: 'Cycles équivalents pleine charge',
    cell: 'C58', input: true, fmt: 'thousands', def: 365, unit: 'cycles/an',
    note: "Limité par le surplus solaire quotidien disponible pour la charge (pas de charge réseau) et par le profil de charge du site ; à affiner par la même simulation horaire." },
  { r: 59, type: 'param', h: 39, label: 'Année de remplacement (mi-vie)',
    cell: 'C59', input: true, fmt: 'int', def: 14, unit: 'année n°',
    note: "Année anticipée selon la durée de vie garantie du fabricant (cycles ou calendaire, première limite atteinte) ; souvent à mi-vie du projet." },
  { r: 60, type: 'param', h: 28, label: 'Coût de remplacement (% du CAPEX initial /kWh)',
    cell: 'C60', input: true, fmt: 'pct1', def: 0.5, unit: '%',
    note: "Anticipe la baisse du coût des cellules à l'horizon du remplacement ; à ajuster selon les courbes d'apprentissage technologique." },
  { r: 61, type: 'param', h: 28, label: "Durée d'amortissement fiscal batterie",
    cell: 'C61', input: true, fmt: 'int', def: 15, unit: 'ans',
    note: "Durée d'amortissement comptable de la batterie retenue par votre expert-comptable, souvent plus courte que celle du solaire." },
  { r: 62, type: 'param', h: 94, label: 'Priorité de substitution — part de la décharge batterie affectée au groupe électrogène',
    cell: 'C62', input: true, fmt: 'pct1', def: 0.8, unit: "% de l'énergie restituée",
    note: "⚑ Spécifique au Scénario 2 : l'énergie stockée en journée puis restituée par la batterie sert D'ABORD à substituer le groupe électrogène (la source la plus coûteuse), le réseau en appoint pour le reste. Ce paramètre fixe la part de la décharge batterie valorisée au tarif groupe électrogène (le solde est valorisé au tarif réseau) ; à ajuster selon la part réelle de la consommation nocturne/de secours couverte par le groupe électrogène sur le site (cf. section 8 ci-dessous)." },
  { r: 63, type: 'spacer' },

  // ── 7. BÉNÉFICES SECONDAIRES ──────────────────────────────────────────────
  { r: 64, type: 'section', text: '7. BÉNÉFICES SECONDAIRES DE LA BATTERIE (SCÉNARIO 2, OPTIONNEL — mettre à 0 si non retenu)' },
  { r: 65, type: 'param', h: 50.15, label: "Réduction de puissance souscrite permise par l'écrêtage",
    cell: 'C65', input: true, fmt: 'thousands', def: 0, unit: 'kW',
    note: "Baisse de la puissance de pointe appelée au réseau, estimée par l'intégrateur BESS selon le profil de charge et le dimensionnement batterie ; mettre à 0 si le contrat de fourniture ne facture pas de prime fixe liée à la puissance souscrite." },
  { r: 66, type: 'spacer' },

  // ── 8. GROUPE ÉLECTROGÈNE ─────────────────────────────────────────────────
  { r: 67, type: 'section', text: "8. GROUPE ÉLECTROGÈNE — MIX ÉNERGÉTIQUE DE RÉFÉRENCE (AVANT SOLAIRE, S'APPLIQUE AUX TROIS SCÉNARIOS)" },
  { r: 68, type: 'param', h: 105, label: 'Part de la consommation du site couverte par le groupe électrogène',
    cell: 'C68', input: true, fmt: 'pct1', def: 0.05, unit: '% de la consommation',
    note: "⚑ De nombreux sites C&I en Afrique de l'Ouest/Centrale couvrent une partie de leur consommation par un groupe électrogène diesel en permanence (délestage récurrent, zone non couverte, secours programmé), indépendamment du solaire. À établir à partir des relevés de fonctionnement du groupe (heures de marche × puissance ÷ consommation totale du site sur 12 mois) ; mettre à 0 si le site est en 100% réseau. Cette part est identique pour les Scénarios 1, 2 et 3 : c'est une caractéristique du site, pas du système solaire." },
  { r: 69, type: 'info', span: 3,
    text: "💡 Le kWh diesel étant structurellement le plus cher des trois sources (cf. onglet Tarifs, section 3), chaque kWh autoconsommé qui aurait dû venir du groupe électrogène génère une économie plus importante qu'un kWh qui aurait été acheté au réseau. Le modèle applique donc, à l'autoconsommation, un tarif évité moyen pondéré entre tarif réseau et tarif groupe électrogène, selon leurs poids respectifs dans le mix ci-dessus (voir onglets Scénario 1 / Scénario 2 / Scénario 3)." },
  { r: 72, type: 'spacer' },

  // ── 9. RÉCAPITULATIF CAPEX & FINANCEMENT ───────────────────────────────────
  { r: 73, type: 'section', text: "9. CAPEX & PLAN DE FINANCEMENT CALCULÉS (récapitulatif — Scénarios 1 & 2 ; le Scénario 3 n'a pas de CAPEX)" },
  { r: 74, type: 'recapHeader', h: 15, c: 'Scénario 1', g: 'Scénario 2' },
  { r: 75, type: 'recap', h: 15.65, label: 'CAPEX solaire (k)', c: 'C75', g: 'G75' },
  { r: 76, type: 'recap', h: 15.65, label: 'CAPEX batterie (k)', c: 'C76', g: 'G76' },
  { r: 77, type: 'recap', h: 15.65, label: 'CAPEX total projet (k)', c: 'C77', g: 'G77', bold: true },
  { r: 78, type: 'recap', h: 15.65, label: 'Dette senior tirée (k)', c: 'C78', g: 'G78' },
  { r: 79, type: 'recap', h: 15.65, label: 'Fonds propres investis (k)', c: 'C79', g: 'G79', bold: true },
  { r: 80, type: 'spacer' },
  { r: 81, type: 'info', span: 2,
    text: "Les CAPEX sont exprimés en milliers de la devise sélectionnée ci-dessus (section 0). Le CAPEX solaire (réf. /kWc × kWc) est identique dans les deux scénarios ; le Scénario 2 ajoute le CAPEX batterie (réf. /kWh × kWh). Le gearing (quote-part dette) est appliqué au même taux sur le CAPEX total des deux scénarios. Le Scénario 3 (référence, sans solaire) n'a aucun CAPEX." }
];

/* Formats d'affichage des cellules calculées hors lignes « param » (colonne D convertie, récap). */
HYPO.CALC_FORMATS = {
  D38: 'thousands', D40: 'thousands', D45: 'thousands', D55: 'thousands',
  C75: 'k', G75: 'k', C76: 'k', G76: 'k', C77: 'k', G77: 'k',
  C78: 'k', G78: 'k', C79: 'k', G79: 'k'
};
