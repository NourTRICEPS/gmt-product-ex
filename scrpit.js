
(function () {
  'use strict';
  var VERSION = '2026-11-09-b', prev = window.__gmtScraper;
  if (prev && prev.version === VERSION) { prev.open(); if (!prev.state.busy) prev.collectAll(); return; }
  if (prev) {                                 
    try { prev.state.stop = true; } catch (e) {}
    var oh = document.getElementById('gmt-scraper-host'); if (oh) oh.remove();
    window.__gmtScraper = null;
  }
var T = {
    ar: {
      title: 'مستخرج المنتجات', scan: 'مسح هذه الصفحة', all: 'جمع كل الصفحات', stop: 'إيقاف',
      xlsx: 'تحميل Excel', csv: 'تحميل CSV', copy: 'نسخ للصق في Excel', clear: 'مسح الكل', close: 'إغلاق',
      search: 'بحث بالاسم أو المرجع…', brand: 'كل العلامات', stock: 'كل الحالات', inStock: 'متوفر', out: 'غير متوفر', order: 'عند الطلب',
      min: 'أدنى سعر', max: 'أعلى سعر', promo: 'العروض فقط', pages: 'أقصى عدد صفحات', shown: 'المعروض', of: 'من',
      found: 'منتج جديد', none: 'لم أجد منتجات في هذه الصفحة. افتح صفحة قسم فيها قائمة منتجات بأسعار.',
      working: 'جارٍ الجمع…', done: 'اكتمل الجمع', copied: 'تم النسخ — الصقه في Excel', empty: 'لا توجد بيانات للتحميل',
      cols: ['#', 'صورة', 'المنتج', 'العلامة', 'المرجع', 'ذكي؟', 'النظام', 'السعر', 'السعر القديم', 'الخصم %', 'التوفر', 'القسم', 'الموقع'],
      sel: 'المحدد فقط', page: 'صفحة', delay: 'تأخير (ث)'
    },
    fr: {
      title: 'Extracteur de produits', scan: 'Scanner cette page', all: 'Collecter toutes les pages', stop: 'Arrêter',
      xlsx: 'Télécharger Excel', csv: 'Télécharger CSV', copy: 'Copier pour Excel', clear: 'Tout effacer', close: 'Fermer',
      search: 'Rechercher nom ou réf…', brand: 'Toutes les marques', stock: 'Tous les stocks', inStock: 'En stock', out: 'Rupture', order: 'Sur commande',
      min: 'Prix min', max: 'Prix max', promo: 'Promos seulement', pages: 'Pages max', shown: 'Affichés', of: 'sur',
      found: 'nouveaux produits', none: "Aucun produit trouvé sur cette page. Ouvrez une page de catégorie avec une liste de produits et prix.",
      working: 'Collecte en cours…', done: 'Collecte terminée', copied: 'Copié — collez dans Excel', empty: 'Aucune donnée à télécharger',
      cols: ['#', 'Image', 'Produit', 'Marque', 'Réf.', 'Smart TV', 'Système', 'Prix', 'Ancien prix', 'Remise %', 'Stock', 'Catégorie', 'Site'],
      sel: 'Sélection seulement', page: 'page', delay: 'Délai (s)'
    },
    en: {
      title: 'Product Extractor', scan: 'Scan this page', all: 'Collect all pages', stop: 'Stop',
      xlsx: 'Download Excel', csv: 'Download CSV', copy: 'Copy for Excel', clear: 'Clear all', close: 'Close',
      search: 'Search name or SKU…', brand: 'All brands', stock: 'All stock', inStock: 'In stock', out: 'Out of stock', order: 'On order',
      min: 'Min price', max: 'Max price', promo: 'Promos only', pages: 'Max pages', shown: 'Shown', of: 'of',
      found: 'new products', none: 'No products found on this page. Open a category page that lists products with prices.',
      working: 'Collecting…', done: 'Collection finished', copied: 'Copied — paste into Excel', empty: 'Nothing to download',
      cols: ['#', 'Image', 'Product', 'Brand', 'SKU', 'Smart TV', 'OS', 'Price', 'Old price', 'Discount %', 'Stock', 'Category', 'Site'],
      sel: 'Selected only', page: 'page', delay: 'Delay (s)'
    }
  };
  var S = { lang: 'en', rows: [], seen: {}, view: [], sort: { k: null, d: 1 }, busy: false, stop: false, sel: {} };
  try { S.lang = localStorage.getItem('gmt_sc_lang2') || 'en'; } catch (e) {}
  var STORE = 'gmt_sc_rows_v1';
  try {
    var saved = JSON.parse(localStorage.getItem(STORE) || '[]');
    saved.forEach(function (r) { var k = keyOf(r); if (!S.seen[k]) { S.seen[k] = 1; S.rows.push(r); } });
  } catch (e) {}
  function t(k) { return (T[S.lang] || T.ar)[k]; }
  var CUR = '(?:DT|TND|د\\.?\\s?ت|دينار|dinars?|€|EUR)';
  var PRICE_FULL = new RegExp('^\\s*(?:' + CUR + '\\s*)?\\d[\\d\\s\\u00a0\\u202f.,]*\\s*(?:' + CUR + ')?\\s*(?:TTC|HT|TVA)?\\s*$', 'i');
  var HAS_CUR = new RegExp(CUR, 'i');
  var NBSP = /[   ]/g;
  var sleep = function (ms) { return new Promise(function (r) { setTimeout(r, ms); }); };
  function txt(el) { return (el && el.textContent || '').replace(NBSP, ' ').replace(/\s+/g, ' ').trim(); }

  function parsePrice(text) {
    text = String(text).replace(NBSP, ' ');
    var m = text.match(/\d[\d\s.,]*\d|\d/);
    if (!m) return null;
    var s = m[0].replace(/\s+/g, '');
    var hasDot = s.indexOf('.') > -1, hasCom = s.indexOf(',') > -1;
    if (hasDot && hasCom) {
      var dec = s.lastIndexOf('.') > s.lastIndexOf(',') ? '.' : ',';
      var th = dec === '.' ? ',' : '.';
      s = s.split(th).join('').replace(dec, '.');
    } else if (hasCom) {
      var pc = s.split(',');
      s = pc.length > 2 ? pc.join('') : s.replace(',', '.');
    } else if (hasDot) {
      var pd = s.split('.');
      if (pd.length > 2) s = pd.join('');
    }
    var v = parseFloat(s);
    return isFinite(v) ? v : null;
  }

  function sigOf(el) {
    var c = (el.getAttribute && el.getAttribute('class') || '').split(/\s+/).filter(function (x) {
      return x && !/^(first|last|odd|even|active|selected|item-?\d+|\d+|is-.*|js-.*|slick-.*|swiper-slide-.*)$/i.test(x);
    }).sort().join('.');
    return el.tagName + '.' + c;
  }

  /* ------------------------------------------------------------------ */
  /* 2) اكتشاف المنتجات داخل أي وثيقة                                   */
  /* ------------------------------------------------------------------ */
  function priceElements(root) {
    var out = [];
    var all = root.querySelectorAll('body *');
    for (var i = 0; i < all.length; i++) {
      var el = all[i], tag = el.tagName;
      if (tag === 'SCRIPT' || tag === 'STYLE' || tag === 'NOSCRIPT' || tag === 'SVG' || tag === 'OPTION') continue;
      var s = txt(el);
      if (!s || s.length > 40 || !/\d/.test(s)) continue;
      if (!PRICE_FULL.test(s) || !HAS_CUR.test(s)) continue;
      var childHit = false;
      for (var c = 0; c < el.children.length; c++) {
        var cs = txt(el.children[c]);
        if (cs && /\d/.test(cs) && PRICE_FULL.test(cs) && HAS_CUR.test(cs)) { childHit = true; break; }
      }
      if (!childHit) out.push(el);
    }
    return out;
  }

  function findCard(el, body) {
    var a = el;
    while (a && a !== body && a.parentElement) {
      var par = a.parentElement, sig = sigOf(a), n = 0;
      for (var i = 0; i < par.children.length; i++) if (sigOf(par.children[i]) === sig) n++;
      if (n >= 2 && (a.matches('a[href]') || a.querySelector('a[href]'))) return a;
      a = par;
    }
    return null;
  }

  function detectCards(doc) {
    var body = doc.body; if (!body) return [];
    var pes = priceElements(body), groups = {}, seenCards = [];
    pes.forEach(function (pe) {
      var card = findCard(pe, body);
      if (!card || seenCards.indexOf(card) > -1) return;
      seenCards.push(card);
      var sig = sigOf(card);
      (groups[sig] = groups[sig] || []).push(card);
    });
    var best = null;
    Object.keys(groups).forEach(function (k) { if (!best || groups[k].length > groups[best].length) best = k; });
    return best ? groups[best] : [];
  }

  /* ------------------------------------------------------------------ */
  /* 3) data                                           */
  /* ------------------------------------------------------------------ */
  var BRANDS = ['Samsung', 'LG', 'Sony', 'TCL', 'Hisense', 'Haier', 'Hoover', 'Candy', 'Beko', 'Bosch', 'Siemens', 'Whirlpool', 'Zerowatt',
    'Fisher & Paykel', 'GE Appliances', 'GE', 'Philips', 'Panasonic', 'Sharp', 'Toshiba', 'Xiaomi', 'Huawei', 'Apple', 'HP', 'Dell', 'Lenovo',
    'Asus', 'Acer', 'MSI', 'Condor', 'Telefunken', 'Brandt', 'Arcelik', 'Arçelik', 'Midea', 'Ariston', 'Indesit', 'Electrolux', 'Vestel', 'Tornado',
    'Iris', 'Bic', 'Canon', 'Epson', 'Brother', 'Dyson', 'Moulinex', 'Tefal', 'Rowenta', 'Kenwood', 'Braun', 'Nikai', 'Hitachi', 'Oppo', 'Realme',
    'Infinix', 'Tecno', 'Nokia', 'Motorola', 'JBL', 'Anker', 'Logitech', 'Corsair', 'Gigabyte', 'Nvidia', 'Amazon', 'Google', 'Gree', 'Carrier',
    'Daikin', 'Mitsubishi', 'Hyundai', 'Kiowa', 'Newstar', 'Saba', 'Techwood', 'Akai', 'Skyworth', 'Cecotec', 'Black+Decker', 'Bomann', 'Blaupunkt', 'Grundig', 'Medion', 'Sagemcom', 'Thomson', 'Nokia', 'Honor', 'Poco', 'Vivo', 'ZTE', 'Alcatel', 'Itel', 'Lenco', 'Smeg', 'Liebherr',
    'Miele', 'Gorenje', 'Zanussi', 'AEG', 'Franke', 'Teka', 'Domo', 'Daewoo', 'Eurosonic', 'Sonashi', 'Starlight', 'Orient', 'Nasco', 'Mika', 'Stream System',
    'Premium', 'Cosmos', 'Ecran', 'Hikvision', 'Dahua', 'TP-Link', 'D-Link', 'Tenda', 'Netgear', 'Cisco', 'Ubiquiti', 'Western Digital', 'WD', 'Seagate',
    'SanDisk', 'Kingston', 'Crucial', 'Lexar', 'Transcend', 'Razer', 'SteelSeries', 'HyperX', 'Redragon', 'Havit', 'Trust', 'Genius', 'A4Tech', 'Microsoft',
    'Fujifilm', 'GoPro', 'DJI', 'Nikon', 'Garmin', 'Fitbit', 'Amazfit', 'Bose', 'Marshall', 'Harman Kardon', 'Sennheiser', 'Yamaha', 'Pioneer', 'Denon',
    'Onkyo', 'Xtreme', 'Soundcore', 'QCY', 'Baseus', 'Ugreen', 'Aukey', 'Belkin', 'Vention', 'Bluestork', 'Advance', 'Techmaster', 'Mac', 'Alienware', 'Razer', 'Vega', 'Biolux', 'Evertek', 'Newline', 'Kaiser', 'Heller', 'Gaggenau', 'Terim', 'Maxwell', 'Techno', 'Orbit', 'Geepas', 'Olsenmark', 'Mebashi', 'Binatone', 'Sinbo', 'Royalty Line', 'Fakir', 'Hotpoint', 'Maytag', 'Frigidaire', 'Kelvinator' ];

  function absUrl(u, doc) { try { return new URL(u, doc.baseURI).href; } catch (e) { return u || ''; } }

  function imgOf(card, doc) {
    var im = card.querySelector('img');
    if (!im) return '';
    var s = im.getAttribute('data-src') || im.getAttribute('data-lazy-src') || im.getAttribute('data-original') || im.getAttribute('data-lazy') ||
      im.currentSrc || im.getAttribute('src') || '';
    if (!s || /^data:/.test(s)) { var ss = im.getAttribute('data-srcset') || im.getAttribute('srcset') || ''; s = ss.split(',')[0].trim().split(' ')[0] || s; }
    return s && !/^data:/.test(s) ? absUrl(s, doc) : '';
  }

  function isOldPrice(el, live) {
    if (el.closest('del,s,strike')) return true;
    var cls = '';
    for (var p = el; p && p.nodeType === 1; p = p.parentElement) {
      cls += ' ' + (p.getAttribute('class') || '');
      if (p.tagName === 'ARTICLE' || p.tagName === 'LI') break;
    }
    if (/(^|[\s_-])(old|was|ancien|before|barre|strike|line-through|crossed)([\s_-]|$)/i.test(cls) && !/special|final|new/i.test(el.getAttribute('class') || '')) return true;
    if (live) { try { if (/line-through/.test(getComputedStyle(el).textDecorationLine || '')) return true; } catch (e) {} }
    return false;
  }

  var AV_OUT = /rupture|[ée]puis|indisponible|non[\s-]?disponible|hors[\s-]?stock|out[\s-]?of[\s-]?stock|outofstock|sold[\s-]?out|soldout|unavailable|no[\s-]?stock|stock[\s-]?out|غير متوفر|غير متاح|نفد/i;
  var AV_ORD = /sur[\s-]?commande|pr[ée]-?commande|pre-?order|preorder|on[\s-]?order|back-?order|en arrivage|arrivage|bient[ôo]t (?:disponible|en stock)|coming[\s-]?soon|عند الطلب|قريب/i;
  var AV_IN = /en[\s-]?stock|in[\s-]?stock|instock|disponible|available|متوفر|متاح/i;
  function avFrom(str) {
    if (!str) return '';
    if (AV_OUT.test(str)) return 'Rupture';
    if (AV_ORD.test(str)) return 'Sur commande';
    if (AV_IN.test(str)) return 'En stock';
    return '';
  }
  function availability(card) {
    var scoped = '';
    [].forEach.call(card.querySelectorAll('[class*="stock"],[class*="avail"],[class*="dispo"],[class*="rupture"],[class*="status"],[class*="etat"],[class*="label"],[class*="badge"],[class*="flag"],[itemprop="availability"],[data-availability]'), function (e) {
      scoped += ' | ' + (e.getAttribute('class') || '') + ' ' + txt(e) + ' ' + (e.getAttribute('href') || e.getAttribute('content') || '') + ' ' + (e.getAttribute('data-availability') || '');
    });
    var r = avFrom(scoped) || avFrom(txt(card));
    if (r) return r;
    var btns = card.querySelectorAll('button,a,input[type="submit"],input[type="button"]'), buy = false, off = false, notify = false;
    [].forEach.call(btns, function (b) {
      var x = (txt(b) || b.value || b.getAttribute('title') || b.getAttribute('aria-label') || '') + ' ' + (b.getAttribute('class') || '');
      if (/me pr[ée]venir|pr[ée]venez|notify|notif|alerte|avertir/i.test(x)) notify = true;
      if (/ajouter au panier|add to cart|add-to-cart|addtocart|tocart|acheter|buy now|أضف|اشتر/i.test(x)) { buy = true; if (b.disabled || /disabled/i.test(b.getAttribute('class') || '') || b.getAttribute('aria-disabled') === 'true') off = true; }
    });
    if (notify || (buy && off)) return 'Rupture';
    if (buy) return 'En stock';
    return '';
  }

  /* ---------- العلامة التجارية ---------- */
  var BR_LOW = null;
  function brLow() {
    if (!BR_LOW) { BR_LOW = {}; BRANDS.forEach(function (b) { BR_LOW[b.toLowerCase()] = b; }); }
    return BR_LOW;
  }
  function canonBrand(b) {
    b = String(b || '').replace(/^\s*(?:marque|brand|fabricant|manufacturer|vendor|الماركة|العلامة)(?:\s*[:：]\s*|\s+-\s*|\s+)/i, '').replace(/\s+/g, ' ').trim();
    if (!b || b.length > 40 || /^(n\/?a|-|—|aucune?|none)$/i.test(b)) return '';
    var k = brLow()[b.toLowerCase()];
    if (k) return k;
    if (b.length > 3 && (b === b.toUpperCase() || b === b.toLowerCase())) b = b.charAt(0).toUpperCase() + b.slice(1).toLowerCase();
    return b;
  }
  function norm(x) { return ' ' + String(x || '').toLowerCase().replace(/[^a-z0-9àâçéèêëîïôûùüÿñæœ&+\- ]/g, ' ').replace(/\s+/g, ' ') + ' '; }
  function findKnown(text, pool) {
    var low = norm(text), best = '', bi = 1e9;
    pool.forEach(function (b) {
      var i = low.indexOf(' ' + b.toLowerCase() + ' ');
      if (i < 0) i = low.indexOf(' ' + b.toLowerCase().replace(/[- ]/g, '') + ' ');
      if (i > -1 && (i < bi || (i === bi && b.length > best.length))) { bi = i; best = b; }
    });
    return best;
  }
  var GENERIC = /^(tv|t[ée]l[ée]viseurs?|t[ée]l[ée]s?|smart|[ée]cran|ecran|pc|portable|ordinateur|laptop|smartphone|t[ée]l[ée]phone|mobile|r[ée]frig[ée]rateur|cong[ée]lateur|lave|lave-linge|lave-vaisselle|machine|climatiseur|split|four|cuisini[èe]re|micro|micro-ondes|aspirateur|imprimante|tablette|casque|enceinte|barre|moniteur|souris|clavier|disque|ssd|c[âa]ble|chargeur|montre|appareil|cafeti[èe]re|bouilloire|mixeur|robot|fer|s[èe]che|s[èe]che-linge|ventilateur|chauffe|chauffe-eau|pack|kit|nouveau|new|led|oled|qled|uhd|4k|8k|full|hd|android|google|webos|pouces?|cm|de|du|la|le|les|et|pour|avec|sans)$/i;
  function guessFirstWord(name) {
    var toks = String(name).replace(/["'’()\[\],;:]/g, ' ').split(/\s+/);
    for (var i = 0; i < toks.length && i < 5; i++) {
      var w = toks[i];
      if (!w || GENERIC.test(w) || /\d/.test(w) || w.length < 2 || w.length > 20 || !/^[A-ZÀ-Ý]/.test(w) || !/^[A-Za-zÀ-ÿ&+\-]+$/.test(w)) continue;
      return w;
    }
    return '';
  }
  function spacedText(el) {
    var out = [];
    (function walk(n) {
      for (var c = n.firstChild; c; c = c.nextSibling) {
        if (c.nodeType === 3) { var v = c.nodeValue.replace(NBSP, ' ').trim(); if (v) out.push(v); }
        else if (c.nodeType === 1 && !/^(SCRIPT|STYLE|NOSCRIPT)$/.test(c.tagName)) walk(c);
      }
    })(el);
    return out.join(' | ');
  }
  function validCode(x) { x = String(x || '').trim(); return x.length >= 3 && /[A-Za-z0-9]{3}/.test(x) && !/^(null|none|n\/a|ref|sku|code)$/i.test(x); }
  /* نوع التلفاز: Smart / Non smart + نظام التشغيل (للتلفزيونات فقط) */
  function tvInfo(name, cardText, category) {
    var n = String(name || ''), all = n + ' ' + String(cardText || '');
    var isTV = /(^|[\s|(])(t[ée]l[ée]viseurs?|t[ée]l[ée]s?|tv|television|تلفاز|تلفزيون|شاشة\s*تلفاز)(?=$|[\s|)\-])/i.test(n) || /t[ée]l[ée]viseur|televis|تلفاز|تلفزيون/i.test(String(category || ''));
    if (isTV && /support|t[ée]l[ée]commande|remote|c[aâ]ble|antenne|d[ée]codeur|r[ée]cepteur|barre\s*de\s*son|soundbar|projecteur|adaptateur|dongle|stick|\bbox\b|meuble|fixation|mural|stand|حامل/i.test(n) && !/^\s*(t[ée]l[ée]viseur|tv|smart tv|led tv)/i.test(n)) isTV = false;
    if (!isTV) return { smart: '', os: '' };
    var os = '';
    if (/google\s*tv/i.test(all)) os = 'Google TV';
    else if (/android\s*tv|android\s*\d*(\.\d+)?|androidtv/i.test(all)) os = 'Android TV';
    else if (/web\s*os|webos/i.test(all)) os = 'webOS';
    else if (/tizen/i.test(all)) os = 'Tizen';
    else if (/vidaa/i.test(all)) os = 'VIDAA';
    else if (/roku/i.test(all)) os = 'Roku TV';
    else if (/fire\s*tv/i.test(all)) os = 'Fire TV';
    else if (/titan\s*os/i.test(all)) os = 'Titan OS';
    else if (/linux/i.test(all)) os = 'Linux';
    else if (/netgem|opera\s*tv|whale\s*os|\bsaphi\b|coolita/i.test(all)) os = (all.match(/netgem|opera\s*tv|whale\s*os|saphi|coolita/i) || [''])[0];
    var smart = os || /smart|wi-?fi|wifi|netflix|youtube|ذكي|ذكية/i.test(all);
    var non = /non[\s-]*smart|not smart|\bnormal\b|غير\s*ذكي/i.test(all);
    return { smart: (smart && !non) ? 'Smart' : 'Non smart', os: os };
  }
  function brandOf(card, name, href) {
    var b = '', el;
    var attrs = ['data-brand', 'data-product-brand', 'data-manufacturer', 'data-marque', 'data-vendor'];
    for (var i = 0; i < attrs.length && !b; i++) {
      var sel = '[' + attrs[i] + ']', e2 = (card.matches && card.matches(sel)) ? card : card.querySelector(sel);
      if (e2) b = canonBrand(e2.getAttribute(attrs[i]));
    }
    if (!b) {
      el = card.querySelector('[itemprop="brand"],[class*="brand"],[class*="marque"],[class*="manufacturer"],[class*="vendor"],[class*="fabricant"]');
      if (el) {
        b = canonBrand(el.getAttribute('content') || txt(el));
        if (!b) { var im = el.querySelector('img') || (el.tagName === 'IMG' ? el : null); if (im) b = canonBrand(im.getAttribute('alt') || im.getAttribute('title') || ''); }
        if (!b && el.getAttribute) b = canonBrand(el.getAttribute('title') || el.getAttribute('aria-label') || '');
      }
    }
    if (!b) b = findKnown(name, BRANDS);
    if (!b && href) { try { b = findKnown(decodeURIComponent(href).replace(/[\/._]/g, ' ').replace(/-/g, ' '), BRANDS); } catch (e) {} }
    if (!b) {
      var alts = ''; [].forEach.call(card.querySelectorAll('img[alt],img[title]'), function (m) { alts += ' ' + (m.getAttribute('alt') || '') + ' ' + (m.getAttribute('title') || ''); });
      b = findKnown(alts, BRANDS);
    }
    return b;
  }
  /* تعبئة العلامات الناقصة بعد الجمع: من العلامات المعروفة في باقي الصفوف + أول كلمة تتكرر */
  function fillBrands() {
    var learned = {}, guesses = {};
    S.rows.forEach(function (r) {
      if (r.brand && r.brand.length < 3 && !brLow()[r.brand.toLowerCase()]) r.brand = '';   /* broken value such as "t" (old bug with Brandt) */
      if (!r.brand) r.brand = findKnown(r.name, BRANDS) || '';
    });
    S.rows.forEach(function (r) { if (r.smart === undefined) { var ti = tvInfo(r.name, '', r.category); r.smart = ti.smart; r.os = ti.os; } });
    S.rows.forEach(function (r) { if (r.brand) learned[r.brand] = 1; else if (r._bg) guesses[r._bg] = (guesses[r._bg] || 0) + 1; });
    var pool = Object.keys(learned);
    S.rows.forEach(function (r) {
      if (r.brand) return;
      var b = pool.length ? findKnown(r.name, pool) : '';
      if (!b && r._bg) b = canonBrand(r._bg);
      if (b) r.brand = b;
    });
  }

  function extract(card, doc, ctx) {
    var live = doc === document;
    var pes = priceElements(card).map(function (el) { return { el: el, v: parsePrice(txt(el)), old: isOldPrice(el, live) }; })
      .filter(function (p) { return p.v !== null && p.v > 0; });
    var price = null, oldp = null;
    if (pes.length) {
      var olds = pes.filter(function (p) { return p.old; }), news = pes.filter(function (p) { return !p.old; });
      if (olds.length && news.length) { price = news[0].v; oldp = Math.max.apply(null, olds.map(function (p) { return p.v; })); }
      else {
        var vals = pes.map(function (p) { return p.v; });
        price = Math.min.apply(null, vals);
        var mx = Math.max.apply(null, vals);
        if (mx > price) oldp = mx;
      }
      if (oldp !== null && oldp <= price) oldp = null;
    }
    /* الاسم + الرابط */
    var nameEl = card.querySelector('h1 a,h2 a,h3 a,h4 a,h5 a,[class*="title"] a,[class*="name"] a,[itemprop="name"] a');
    var name = '', href = '';
    if (nameEl) { name = txt(nameEl); href = nameEl.getAttribute('href') || ''; }
    if (!name) {
      var hd = card.querySelector('h1,h2,h3,h4,h5,[class*="title"],[class*="name"],[itemprop="name"]');
      if (hd) name = txt(hd);
    }
    if (!href) {
      var best = null;
      [].forEach.call(card.querySelectorAll('a[href]'), function (a) {
        var s = txt(a);
        if (/^(#|javascript:)/i.test(a.getAttribute('href') || '')) return;
        if (!best || s.length > txt(best).length) best = a;
      });
      if (best) { href = best.getAttribute('href'); if (!name) name = txt(best); }
      else if (card.matches && card.matches('a[href]')) href = card.getAttribute('href');
    }
    if (!name) { var im = card.querySelector('img[alt]'); if (im) name = (im.getAttribute('alt') || '').trim(); }
    if (!name) return null;
    /* العلامة */
    var brand = brandOf(card, name, href), bg = brand ? '' : guessFirstWord(name);
    /* المرجع */
    var sku = '';
    var sEl = card.querySelector('[class*="sku"],[class*="reference"],[class*="-ref"],[itemprop="sku"]');
    if (sEl) sku = txt(sEl).replace(/^(r[ée]f(?:[ée]rence)?|sku|code)\s*[:.]?\s*/i, '');
    if (!sku) { var rm = txt(card).match(/\b(?:r[ée]f(?:[ée]rence)?|sku|code)\b\s*[:.]\s*([A-Za-z0-9][\w\-\/.]{2,})/i); if (rm) sku = rm[1]; }
    var av = availability(card), tvi = tvInfo(name, spacedText(card), ctx.category);
    return {
      source: ctx.host, category: ctx.category, brand: brand, _bg: bg, sku: validCode(sku) ? sku.slice(0, 60) : '', smart: tvi.smart, os: tvi.os, name: name,
      price: price, oldPrice: oldp, discount: (oldp && price) ? Math.round((1 - price / oldp) * 1000) / 10 : null,
      availability: av, url: href ? absUrl(href, doc) : '', image: imgOf(card, doc), page: ctx.url, date: ctx.date
    };
  }

  function categoryOf(doc) {
    var bc = doc.querySelector('[class*="breadcrumb"] li:last-child,[class*="breadcrumb"] a:last-child,nav[aria-label*="read" i] li:last-child,[class*="breadcrumb"] span:last-child');
    var s = bc ? txt(bc) : '';
    if (!s) { var h = doc.querySelector('h1'); s = h ? txt(h) : ''; }
    return (s || (doc.title || '')).slice(0, 80);
  }

  function keyOf(r) { return (r.url || r.name) + '|' + (r.sku || '') + '|' + r.name; }

  function scanDoc(doc, url) {
    var ctx = { host: location.hostname.replace(/^www\./, ''), category: categoryOf(doc), url: url || location.href, date: new Date().toISOString().slice(0, 10) };
    var cards = detectCards(doc), added = 0;
    cards.forEach(function (c) {
      var r = extract(c, doc, ctx);
      if (!r) return;
      var k = keyOf(r);
      if (S.seen[k]) return;
      S.seen[k] = 1; S.rows.push(r); added++;
    });
    fillBrands();
    persist();
    return added;
  }
  function persist() { try { localStorage.setItem(STORE, JSON.stringify(S.rows)); } catch (e) {} }

  /* ------------------------------------------------------------------ */
  /* 4) التنقل بين الصفحات                                              */
  /* ------------------------------------------------------------------ */
  function disabled(el) {
    return !!(el.closest('.disabled,[aria-disabled="true"],.inactive') || el.getAttribute('aria-disabled') === 'true' || el.disabled);
  }
  function findNext(doc) {
    var lr = doc.querySelector('link[rel~="next"][href]');
    if (lr) { var lh = lr.getAttribute('href'); if (lh && !/^(#|javascript:)/i.test(lh)) return { href: absUrl(lh, doc), el: null }; }
    var sels = ['li.active + li > a[href]', 'li.current + li > a[href]', '.page-item.active + .page-item a[href]', '.current + a[href]', '.active + a[href]',
      'a[class*="next" i][href]:not([class*="prev" i])', 'a[id*="next" i][href]', 'li[class*="next" i] > a[href]', 'a[data-page][class*="next" i]',
     'a[rel~="next"]', 'li.next a', '.pagination .next a', '.pagination a.next', 'a.next', '.pages-item-next a', 'a.action.next',
      'a[aria-label*="next" i]', 'a[aria-label*="suiv" i]', 'a[title*="suiv" i]', 'a[title*="next" i]', '.next-page a', 'a.next-page'];
    for (var i = 0; i < sels.length; i++) {
      var els = doc.querySelectorAll(sels[i]);
      for (var j = 0; j < els.length; j++) {
        var el = els[j], h = el.getAttribute('href');
        if (h && !/^(#|javascript:)/i.test(h) && !disabled(el)) return { href: absUrl(h, doc), el: el };
      }
    }
    var as = doc.querySelectorAll('a[href]');
    for (var k = 0; k < as.length; k++) {
      var s = txt(as[k]);
      if (/^(suivant|suiv\.?|next|page suivante|التالي|›|»|>|→)$/i.test(s) && !disabled(as[k]) && !/^(#|javascript:)/i.test(as[k].getAttribute('href'))) {
        return { href: absUrl(as[k].getAttribute('href'), doc), el: as[k] };
      }
    }
    /* ترقيم بالأرقام: ابحث عن الصفحة الحالية ثم رقم أكبر منها بواحد */
    var curEl = doc.querySelector('.pagination .current, .pagination .active, .pages .current, .pages .item.current, .page-numbers.current, [aria-current="page"], li.current, li.active');
    if (curEl) {
      var cn = parseInt(txt(curEl).replace(/\D+/g, ''), 10);
      if (cn) {
        var box = curEl.closest('ul,ol,nav,div') || doc, links = box.querySelectorAll('a[href]');
        for (var q = 0; q < links.length; q++) {
          if (parseInt(txt(links[q]).replace(/\D+/g, ''), 10) === cn + 1 && /^\s*\D*\d+\D*\s*$/.test(txt(links[q])) && !/^(#|javascript:)/i.test(links[q].getAttribute('href'))) {
            return { href: absUrl(links[q].getAttribute('href'), doc), el: links[q] };
          }
        }
      }
    }
    return null;
  }
  /* آخر حل: زيادة رقم الصفحة في الرابط (?p=2 أو ?page=2 أو /page/2) */
  function guessNextUrl(href, n) {
    try {
      var u = new URL(href), keys = ['p', 'page', 'pg', 'pageNumber', 'page_number', 'paged'];
      for (var i = 0; i < keys.length; i++) {
        if (u.searchParams.has(keys[i])) { u.searchParams.set(keys[i], String((parseInt(u.searchParams.get(keys[i]), 10) || 1) + 1)); return u.href; }
      }
      var m = u.pathname.match(/^(.*\/page[\/-])(\d+)(\/?.*)$/i);
      if (m) { u.pathname = m[1] + (parseInt(m[2], 10) + 1) + m[3]; return u.href; }
      if (n === 1) { u.searchParams.set('p', '2'); return u.href; }
    } catch (e) {}
    return null;
  }
  function findLoadMore() {
    var els = document.querySelectorAll('button,a,[role="button"],input[type="button"]');
    for (var i = 0; i < els.length; i++) {
      var s = txt(els[i]) || els[i].value || '';
      if (/(voir plus|charger plus|afficher plus|plus de produits|load more|show more|view more|المزيد|عرض المزيد)/i.test(s) && !disabled(els[i]) && els[i].offsetParent) return els[i];
    }
    return null;
  }
  function countPrices() { return priceElements(document.body).length; }
  async function waitGrowth(before, ms) {
    var t0 = Date.now();
    while (Date.now() - t0 < ms) { await sleep(300); if (countPrices() > before) { await sleep(500); return true; } }
    return false;
  }
  async function tryMore() {
    var btn = findLoadMore(), before = countPrices();
    if (btn) { btn.scrollIntoView({ block: 'center' }); btn.click(); return waitGrowth(before, 7000); }
    window.scrollTo(0, document.body.scrollHeight);
    return waitGrowth(before, 2500);
  }


  /* تحميل الصفحة في إطار مخفي (للمواقع التي ترسم المنتجات بالجافاسكربت) */
  function loadFrame(url, ms) {
    return new Promise(function (res) {
      var f = document.createElement('iframe');
      f.style.cssText = 'position:fixed;left:-99999px;top:0;width:1280px;height:900px;border:0;opacity:0;pointer-events:none';
      f.src = url; document.body.appendChild(f);
      var t0 = Date.now();
      (function poll() {
        var d = null; try { d = f.contentDocument; } catch (e) {}
        if (d && d.readyState === 'complete' && d.body && priceElements(d.body).length > 0) {
          try { f.contentWindow.scrollTo(0, d.body.scrollHeight); } catch (e) {}
          setTimeout(function () { res({ doc: d, frame: f }); }, 900); return;
        }
        if (Date.now() - t0 > ms) { res({ doc: d, frame: f }); return; }
        setTimeout(poll, 400);
      })();
    });
  }
  function guessList(href, n) {
    var out = [];
    try {
      var u = new URL(href), keys = ['p', 'page', 'pg', 'pageNumber', 'page_number', 'paged'];
      for (var i = 0; i < keys.length; i++) {
        if (u.searchParams.has(keys[i])) { u.searchParams.set(keys[i], String((parseInt(u.searchParams.get(keys[i]), 10) || 1) + 1)); return [u.href]; }
      }
      var m = u.pathname.match(/^(.*\/page[\/-])(\d+)(\/?.*)$/i);
      if (m) { u.pathname = m[1] + (parseInt(m[2], 10) + 1) + m[3]; return [u.href]; }
      if (n === 1) {
        ['p', 'page', 'pg', 'pageNumber'].forEach(function (k) { var v = new URL(href); v.searchParams.set(k, '2'); out.push(v.href); });
        var v2 = new URL(href); v2.pathname = v2.pathname.replace(/\/$/, '') + '/page/2'; out.push(v2.href);
      }
    } catch (e) {}
    return out;
  }
  async function collectAll() {
    if (S.busy) return;
    S.busy = true; S.stop = false; refresh();
    var maxPages = parseInt(ui('#max').value, 10) || 200, delay = Math.max(0.3, parseFloat(ui('#delay').value) || 1) * 1000;
    var total = scanDoc(document, location.href), pages = 1, cur = document, seen = {}, lastHref = location.href, why = '', frame = null;
    seen[location.href] = 1;
    status(t('working') + ' ' + t('page') + ' 1 · +' + total);
    async function grab(url, useFrame) {
      var added = 0, d = null;
      try {
        var res = await fetch(url, { credentials: 'include' });
        var html = await res.text();
        d = new DOMParser().parseFromString(html, 'text/html');
        var base = d.createElement('base'); base.href = url; d.head.insertBefore(base, d.head.firstChild);
        added = scanDoc(d, url);
      } catch (e) { d = null; }
      if (!added && useFrame) {                       /* المنتجات تُرسم بالجافاسكربت: حمّل الصفحة في إطار */
        if (frame) { try { frame.remove(); } catch (e) {} frame = null; }
        var fr = await loadFrame(url, 14000); frame = fr.frame;
        if (fr.doc) { var a2 = scanDoc(fr.doc, url); if (a2) { added = a2; d = fr.doc; } }
      }
      return { added: added, doc: d };
    }
    try {
      while (!S.stop && pages < maxPages) {
        var nx = findNext(cur), progressed = false, added = 0;
        if (nx && !seen[nx.href]) {
          seen[nx.href] = 1; await sleep(delay);
          var r1 = await grab(nx.href, true);
          if (r1.added) { progressed = true; added = r1.added; cur = r1.doc; lastHref = nx.href; }
          else why = 'next link found but its page had no new products';
        }
        if (!progressed) {                           /* لا يوجد «التالي»: خمّن رابط الصفحة التالية */
          var gl = guessList(lastHref, pages);
          for (var gi = 0; gi < gl.length && !progressed && !S.stop; gi++) {
            if (seen[gl[gi]]) continue;
            seen[gl[gi]] = 1; await sleep(delay);
            var r2 = await grab(gl[gi], false);
            if (r2.added) { progressed = true; added = r2.added; cur = r2.doc; lastHref = gl[gi]; }
          }
        }
        if (!progressed) {                           /* «عرض المزيد» أو التمرير اللانهائي */
          var grew = await tryMore();
          if (!grew) { why = why || 'no next link, no "load more" button, and scrolling loaded nothing'; break; }
          added = scanDoc(document, location.href);
          if (!added) { why = 'page grew but no new products found'; break; }
          cur = document;
        }
        pages++; total += added;
        status(t('working') + ' ' + t('page') + ' ' + pages + ' · ' + S.rows.length);
        refresh();
      }
    } catch (e) { status('⚠ ' + e.message); }
    if (frame) { try { frame.remove(); } catch (e) {} }
    S.busy = false;
    status(t('done') + ' — ' + pages + ' ' + t('page') + ' · ' + total + ' ' + t('found') + (why && pages < maxPages && !S.stop ? '  (stopped: ' + why + ')' : ''));
    refresh();
  }

  /* ------------------------------------------------------------------ */
  /* 5) تصدير XLSX (بدون مكتبات) و CSV                                  */
  /* ------------------------------------------------------------------ */
  var HEAD = ['Photo', 'Source', 'Category', 'Brand', 'SKU', 'Smart TV', 'OS', 'Product', 'Price (TND)', 'Old price (TND)', 'Discount %', 'Availability', 'URL', 'Image', 'Date'];
  function rowVals(r) { return ['', r.source, r.category, r.brand, r.sku, r.smart, r.os, r.name, r.price, r.oldPrice, r.discount, r.availability, r.url, r.image, r.date]; }

  function crc32(buf) {
    var tb = crc32.tb;
    if (!tb) { tb = crc32.tb = []; for (var n = 0, c; n < 256; n++) { c = n; for (var k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; tb[n] = c >>> 0; } }
    var c2 = -1; for (var i = 0; i < buf.length; i++) c2 = tb[(c2 ^ buf[i]) & 255] ^ (c2 >>> 8);
    return (c2 ^ -1) >>> 0;
  }
  function zip(files) {
    var enc = new TextEncoder(), chunks = [], central = [], offset = 0;
    files.forEach(function (f) {
      var nb = enc.encode(f.name), data = typeof f.data === 'string' ? enc.encode(f.data) : f.data, crc = crc32(data), size = data.length;
      var lh = new DataView(new ArrayBuffer(30));
      lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true); lh.setUint16(12, 0x21, true);
      lh.setUint32(14, crc, true); lh.setUint32(18, size, true); lh.setUint32(22, size, true); lh.setUint16(26, nb.length, true);
      chunks.push(new Uint8Array(lh.buffer), nb, data);
      var ch = new DataView(new ArrayBuffer(46));
      ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true); ch.setUint16(8, 0x0800, true); ch.setUint16(14, 0x21, true);
      ch.setUint32(16, crc, true); ch.setUint32(20, size, true); ch.setUint32(24, size, true); ch.setUint16(28, nb.length, true); ch.setUint32(42, offset, true);
      central.push(new Uint8Array(ch.buffer), nb);
      offset += 30 + nb.length + size;
    });
    var cd = central.reduce(function (a, c) { return a + c.length; }, 0);
    var end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true); end.setUint16(8, files.length, true); end.setUint16(10, files.length, true); end.setUint32(12, cd, true); end.setUint32(16, offset, true);
    return new Blob(chunks.concat(central, [new Uint8Array(end.buffer)]), { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }
  function esc(s) { return String(s).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
  function colL(i) { return String.fromCharCode(65 + i); }

  function buildXlsx(rows, photos) {
    photos = photos || [];
    var widths = [14, 18, 26, 14, 16, 12, 16, 60, 14, 14, 11, 14, 50, 40, 12];
    var x = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>';
    widths.forEach(function (w, i) { x += '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + w + '" customWidth="1"/>'; });
    x += '</cols><sheetData><row r="1">';
    HEAD.forEach(function (h, i) { x += '<c r="' + colL(i) + '1" s="1" t="inlineStr"><is><t>' + esc(h) + '</t></is></c>'; });
    x += '</row>';
    rows.forEach(function (r, ri) {
      var n = ri + 2; x += '<row r="' + n + '"' + (photos[ri] ? ' ht="66" customHeight="1"' : '') + '>';
      rowVals(r).forEach(function (v, i) {
        if (v === null || v === undefined || v === '') return;
        var ref = colL(i) + n;
        if (typeof v === 'number') x += '<c r="' + ref + '"' + (i === 8 || i === 9 ? ' s="2"' : i === 10 ? ' s="3"' : '') + '><v>' + v + '</v></c>';
        else x += '<c r="' + ref + '" t="inlineStr"><is><t xml:space="preserve">' + esc(v) + '</t></is></c>';
      });
      x += '</row>';
    });
    var pics = [];
    photos.forEach(function (p, ri) { if (p) pics.push({ p: p, ri: ri }); });
    x += '</sheetData><autoFilter ref="A1:O' + (rows.length + 1) + '"/>' + (pics.length ? '<drawing r:id="rId1"/>' : '') + '</worksheet>';
    var files = [
      { name: '[Content_Types].xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>' },
      { name: '_rels/.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>' },
      { name: 'xl/workbook.xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Products" sheetId="1" r:id="rId1"/></sheets></workbook>' },
      { name: 'xl/_rels/workbook.xml.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>' },
      { name: 'xl/styles.xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="2"><numFmt numFmtId="164" formatCode="#,##0.000"/><numFmt numFmtId="165" formatCode="0.0"/></numFmts><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFD6336C"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>' },
      { name: 'xl/worksheets/sheet1.xml', data: x }
    ];
    if (pics.length) {
      var EMU = 9525, dx = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
      var dr = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
      pics.forEach(function (o, k) {
        var id = k + 1, w = Math.round(o.p.w * EMU), h = Math.round(o.p.h * EMU);
        files.push({ name: 'xl/media/image' + id + '.jpeg', data: o.p.bytes });
        dr += '<Relationship Id="rId' + id + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image' + id + '.jpeg"/>';
        dx += '<xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>' + Math.round(((100 - o.p.w) / 2 + 4) * EMU) + '</xdr:colOff><xdr:row>' + (o.ri + 1) + '</xdr:row><xdr:rowOff>' + Math.round(Math.max(2, (88 - o.p.h) / 2) * EMU) + '</xdr:rowOff></xdr:from><xdr:ext cx="' + w + '" cy="' + h + '"/>' +
          '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="' + (id + 1) + '" name="Photo ' + id + '"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId' + id + '"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>' +
          '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + w + '" cy="' + h + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>';
      });
      files.push({ name: 'xl/drawings/drawing1.xml', data: dx + '</xdr:wsDr>' });
      files.push({ name: 'xl/drawings/_rels/drawing1.xml.rels', data: dr + '</Relationships>' });
      files.push({ name: 'xl/worksheets/_rels/sheet1.xml.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>' });
      files[0].data = files[0].data.replace('<Default Extension="xml"', '<Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="xml"').replace('</Types>', '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>');
    }
    return zip(files);
  }
  function csvCell(v) { v = v === null || v === undefined ? '' : String(v); return /[",;\n\r]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; }
  function buildCsv(rows) { return '﻿' + [HEAD].concat(rows.map(rowVals)).map(function (r) { return r.slice(1).map(csvCell).join(';'); }).join('\r\n'); }
  function save(blob, name) {
    var a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click(); setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
  }
  function fname(ext) { return 'produits_' + location.hostname.replace(/^www\./, '') + '_' + new Date().toISOString().slice(0, 10) + '.' + ext; }
  /* ---------- Photos for the Excel file: download each product image, shrink it to a small JPEG ---------- */
  function loadBitmap(blob) {
    if (window.createImageBitmap) return createImageBitmap(blob);
    return new Promise(function (res, rej) { var u = URL.createObjectURL(blob), im = new Image(); im.onload = function () { res(im); }; im.onerror = rej; im.src = u; });
  }
  function toThumb(src) {
    var w = src.width || src.naturalWidth, h = src.height || src.naturalHeight;
    if (!w || !h) return Promise.reject(new Error('empty'));
    var k = Math.min(1, 200 / Math.max(w, h)), cw = Math.max(1, Math.round(w * k)), ch = Math.max(1, Math.round(h * k));
    var cv = document.createElement('canvas'); cv.width = cw; cv.height = ch;
    var cx = cv.getContext('2d'); cx.fillStyle = '#fff'; cx.fillRect(0, 0, cw, ch); cx.drawImage(src, 0, 0, cw, ch);
    var dk = 80 / Math.max(cw, ch);   /* size shown in Excel: max 80 px */
    return new Promise(function (res, rej) {
      cv.toBlob(function (b) {
        if (!b) return rej(new Error('blob'));
        b.arrayBuffer().then(function (ab) { res({ bytes: new Uint8Array(ab), w: Math.round(cw * dk), h: Math.round(ch * dk) }); }, rej);
      }, 'image/jpeg', 0.85);
    });
  }
  function fetchPhoto(url) {
    var timeout = new Promise(function (_, rej) { setTimeout(function () { rej(new Error('timeout')); }, 20000); });
    var viaFetch = fetch(url, { credentials: 'omit' }).then(function (r) { if (!r.ok) throw new Error('http ' + r.status); return r.blob(); }).then(loadBitmap).then(toThumb);
    var viaImg = function () {
      return new Promise(function (res, rej) {
        var im = new Image(); im.crossOrigin = 'anonymous';
        im.onload = function () { toThumb(im).then(res, rej); }; im.onerror = function () { rej(new Error('img')); }; im.src = url;
      });
    };
    return Promise.race([viaFetch.catch(viaImg), timeout]);
  }
  function fetchPhotos(rows, progress) {
    var out = new Array(rows.length), i = 0, done = 0, ok = 0, S2 = S.imgCache = S.imgCache || {};
    function worker() {
      if (i >= rows.length) return Promise.resolve();
      var idx = i++, url = rows[idx].image;
      var p = !url ? Promise.resolve(null) : (S2[url] !== undefined ? Promise.resolve(S2[url]) : fetchPhoto(url).then(function (t) { S2[url] = t; return t; }, function () { S2[url] = null; return null; }));
      return p.then(function (t) { out[idx] = t; done++; if (t) ok++; progress(done, rows.length, ok); return worker(); });
    }
    var ws = []; for (var k = 0; k < 4; k++) ws.push(worker());
    return Promise.all(ws).then(function () { return { photos: out, ok: ok }; });
  }

  function exportRows() {
    try { fillBrands(); } catch (e) {}
    var rows = ui('#selOnly').checked ? S.view.filter(function (r) { return S.sel[keyOf(r)]; }) : S.view;
    return rows;
  }

  /* ------------------------------------------------------------------ */
  /* 6) الواجهة                                                         */
  /* ------------------------------------------------------------------ */
  var host = document.createElement('div');
  host.id = 'gmt-scraper-host';
  host.style.cssText = 'all:initial;position:fixed;inset:0;z-index:2147483647;pointer-events:none';
  var root = host.attachShadow({ mode: 'open' });
  var CSS = '*{box-sizing:border-box;font-family:Segoe UI,Tahoma,Arial,sans-serif}' +
    '.wrap{pointer-events:auto;position:fixed;top:2vh;left:2vw;right:2vw;bottom:2vh;background:#fff;color:#1a1a1a;border-radius:14px;box-shadow:0 20px 70px rgba(0,0,0,.45);display:flex;flex-direction:column;overflow:hidden;border:2px solid #D6336C}' +
    '.wrap.min{bottom:auto;height:54px}.wrap.min .bd{display:none}' +
    '.hd{display:flex;align-items:center;gap:10px;padding:10px 16px;background:#D6336C;color:#fff}.hd b{font-size:17px;flex:1}' +
    'button{cursor:pointer;border:0;border-radius:7px;padding:9px 14px;font-weight:700;font-size:13px;background:#D6336C;color:#fff}button:hover{background:#FFD21F;color:#111}' +
    'button.sec{background:#fff;color:#D6336C;border:1.5px solid #D6336C}button.sec:hover{background:#FFD21F;color:#111;border-color:#FFD21F}' +
    'button.hdb{background:rgba(255,255,255,.2)}button:disabled{opacity:.45;cursor:not-allowed}' +
    '.bar{display:flex;flex-wrap:wrap;gap:8px;align-items:center;padding:10px 14px;border-bottom:1px solid #eee;background:#fafafa}' +
    'input,select{padding:8px 10px;border:1.5px solid #ddd;border-radius:7px;font-size:13px;background:#fff;color:#111}input[type=number]{width:96px}' +
    'label{font-size:13px;display:flex;align-items:center;gap:5px}.st{padding:6px 16px;font-size:13px;color:#555;background:#fff;border-bottom:1px solid #eee;min-height:30px}' +
    '.tb{flex:1;overflow:auto}table{border-collapse:collapse;width:100%;font-size:13px}th{position:sticky;top:0;background:#222;color:#fff;padding:9px 8px;text-align:start;cursor:pointer;white-space:nowrap;z-index:1}' +
    'td{padding:7px 8px;border-bottom:1px solid #eee;vertical-align:middle}tr:hover td{background:#fff4f8}td img{width:42px;height:42px;object-fit:contain;border-radius:6px;background:#f2f2f2}' +
    'a{color:#1b5fd1;text-decoration:none}.pr{font-weight:800;white-space:nowrap}.old{color:#999;text-decoration:line-through;white-space:nowrap}.dc{color:#fff;background:#D6336C;border-radius:10px;padding:1px 7px;font-weight:700}' +
    '.av-In{color:#0a8a3c;font-weight:700}.av-Ru{color:#c62828;font-weight:700}.av-Su{color:#b36b00;font-weight:700}.spacer{flex:1}';
  function ui(sel) { return root.querySelector(sel); }
  var htmlShell = '<style>' + CSS + '</style><div class="wrap" id="wrap"><div class="hd"><b id="ttl"></b>' +
    '<button class="hdb" id="lg">FR/AR/EN</button><button class="hdb" id="mn">—</button><button class="hdb" id="cl">✕</button></div>' +
    '<div class="bd" style="display:flex;flex-direction:column;flex:1;min-height:0">' +
    '<div class="bar"><button id="scan"></button><button id="all"></button><button id="stop" class="sec"></button>' +
    '<label><span id="lMax"></span><input id="max" type="number" value="200" min="1" max="1000" style="width:70px"></label>' +
    '<label><span id="lDelay"></span><input id="delay" type="number" value="1" min="0.3" step="0.1" style="width:70px"></label><span class="spacer"></span>' +
    '<button id="xl"></button><button id="cv" class="sec"></button><button id="cp" class="sec"></button><button id="rm" class="sec"></button></div>' +
    '<div class="bar"><input id="q" style="min-width:220px"><select id="br"></select><select id="av"></select>' +
    '<input id="pmin" type="number"><input id="pmax" type="number"><label><input id="promo" type="checkbox"><span id="lPromo"></span></label>' +
    '<label><input id="selOnly" type="checkbox"><span id="lSel"></span></label><span class="spacer"></span><b id="cnt"></b></div>' +
    '<div class="st" id="st"></div><div class="tb"><table><thead><tr id="th"></tr></thead><tbody id="tbody"></tbody></table></div></div></div>';
  root.innerHTML = htmlShell;
  document.documentElement.appendChild(host);

  function status(s) { ui('#st').textContent = s; }
  function applyLang() {
    var L = T[S.lang];
    ui('#wrap').setAttribute('dir', S.lang === 'ar' ? 'rtl' : 'ltr');
    ui('#ttl').textContent = '🛒 ' + L.title;
    ui('#scan').textContent = L.scan; ui('#all').textContent = L.all; ui('#stop').textContent = L.stop;
    ui('#lMax').textContent = L.pages; ui('#lDelay').textContent = L.delay;
    ui('#xl').textContent = '⬇ ' + L.xlsx; ui('#cv').textContent = L.csv; ui('#cp').textContent = L.copy; ui('#rm').textContent = L.clear;
    ui('#q').placeholder = L.search; ui('#pmin').placeholder = L.min; ui('#pmax').placeholder = L.max;
    ui('#lPromo').textContent = L.promo; ui('#lSel').textContent = L.sel;
    ui('#th').innerHTML = '<th><input type="checkbox" id="selAll"></th>' + L.cols.map(function (c, i) { return '<th data-i="' + i + '">' + c + '</th>'; }).join('');
    refresh();
  }
  var SORTKEY = { 2: 'name', 3: 'brand', 4: 'sku', 5: 'smart', 6: 'os', 7: 'price', 8: 'oldPrice', 9: 'discount', 10: 'availability', 11: 'category', 12: 'source' };

  function fmt(n) { return n === null || n === undefined ? '' : n.toLocaleString('fr-FR', { minimumFractionDigits: 3, maximumFractionDigits: 3 }); }
  function refresh() {
    /* العلامات */
    var brands = {}; S.rows.forEach(function (r) { if (r.brand) brands[r.brand] = (brands[r.brand] || 0) + 1; });
    var br = ui('#br'), cur = br.value;
    br.innerHTML = '<option value="">' + t('brand') + '</option>' + Object.keys(brands).sort().map(function (b) { return '<option value="' + esc(b) + '">' + esc(b) + ' (' + brands[b] + ')</option>'; }).join('');
    br.value = cur;
    var L2 = T[S.lang], avc0 = {}; S.rows.forEach(function (r) { var k = r.availability || '__none'; avc0[k] = (avc0[k] || 0) + 1; });
    var avs = ui('#av'), avcur = avs.value, UNK = { ar: 'غير معروف', fr: 'Inconnu', en: 'Unknown' }[S.lang];
    avs.innerHTML = '<option value="">' + L2.stock + '</option>' +
      [['En stock', L2.inStock], ['Sur commande', L2.order], ['Rupture', L2.out], ['__none', UNK]].map(function (o) { return '<option value="' + o[0] + '">' + o[1] + ' (' + (avc0[o[0]] || 0) + ')</option>'; }).join('');
    avs.value = avcur;
    var q = ui('#q').value.toLowerCase().trim(), b = br.value, av = ui('#av').value,
      pmin = parseFloat(ui('#pmin').value), pmax = parseFloat(ui('#pmax').value), promo = ui('#promo').checked;
    S.view = S.rows.filter(function (r) {
      if (q && (r.name + ' ' + r.sku + ' ' + (r.os || '') + ' ' + r.brand).toLowerCase().indexOf(q) < 0) return false;
      if (b && r.brand !== b) return false;
      if (av && (av === '__none' ? !!r.availability : r.availability !== av)) return false;
      if (!isNaN(pmin) && (r.price === null || r.price < pmin)) return false;
      if (!isNaN(pmax) && (r.price === null || r.price > pmax)) return false;
      if (promo && !r.discount) return false;
      return true;
    });
    if (S.sort.k) {
      var k = S.sort.k, d = S.sort.d;
      S.view.sort(function (a, c) {
        var x = a[k], y = c[k];
        if (x === null || x === '' || x === undefined) return 1;
        if (y === null || y === '' || y === undefined) return -1;
        return (typeof x === 'number' ? x - y : String(x).localeCompare(String(y))) * d;
      });
    }
    var h = '', shown = S.view.slice(0, 1500);
    shown.forEach(function (r, i) {
      var key = keyOf(r), avc = r.availability ? { 'En stock': 'av-In', 'Rupture': 'av-Ru', 'Sur commande': 'av-Su' }[r.availability] || '' : '';
      h += '<tr><td><input type="checkbox" data-k="' + esc(key) + '"' + (S.sel[key] ? ' checked' : '') + '></td><td>' + (i + 1) + '</td>' +
        '<td>' + (r.image ? '<img loading="lazy" src="' + esc(r.image) + '">' : '') + '</td>' +
        '<td>' + (r.url ? '<a target="_blank" rel="noopener" href="' + esc(r.url) + '">' + esc(r.name) + '</a>' : esc(r.name)) + '</td>' +
        '<td>' + esc(r.brand) + '</td><td>' + esc(r.sku) + '</td><td>' + esc(r.smart) + '</td><td>' + esc(r.os) + '</td><td class="pr">' + fmt(r.price) + '</td><td class="old">' + fmt(r.oldPrice) + '</td>' +
        '<td>' + (r.discount ? '<span class="dc">-' + r.discount + '%</span>' : '') + '</td><td class="' + avc + '">' + esc(r.availability) + '</td>' +
        '<td>' + esc(r.category) + '</td><td>' + esc(r.source) + '</td></tr>';
    });
    ui('#tbody').innerHTML = h;
    ui('#cnt').textContent = t('shown') + ' ' + S.view.length + ' ' + t('of') + ' ' + S.rows.length;
    ui('#all').disabled = S.busy; ui('#scan').disabled = S.busy; ui('#stop').disabled = !S.busy;
  }

  root.addEventListener('input', function (e) { if (['q', 'pmin', 'pmax'].indexOf(e.target.id) > -1) refresh(); });
  root.addEventListener('change', function (e) {
    var id = e.target.id;
    if (['br', 'av', 'promo', 'selOnly'].indexOf(id) > -1) refresh();
    if (id === 'selAll') { S.view.forEach(function (r) { S.sel[keyOf(r)] = e.target.checked; }); refresh(); }
    if (e.target.dataset && e.target.dataset.k) S.sel[e.target.dataset.k] = e.target.checked;
  });
  root.addEventListener('click', function (e) {
    var th = e.target.closest('th[data-i]');
    if (th && SORTKEY[th.dataset.i]) {
      var k = SORTKEY[th.dataset.i]; S.sort = { k: k, d: S.sort.k === k ? -S.sort.d : 1 }; refresh();
    }
  });
  ui('#scan').onclick = function () {
    var n = scanDoc(document, location.href);
    status(n ? n + ' ' + t('found') : t('none')); refresh();
  };
  ui('#all').onclick = collectAll;
  ui('#stop').onclick = function () { S.stop = true; };
  ui('#xl').onclick = function () {
    var r = exportRows(); if (!r.length) return status(t('empty'));
    var btn = ui('#xl'); if (btn.disabled) return; btn.disabled = true;
    fetchPhotos(r, function (d, n, ok) { status('📷 ' + d + ' / ' + n + ' (' + ok + ' OK)'); }).then(function (res) {
      save(buildXlsx(r, res.photos), fname('xlsx'));
      var withUrl = r.filter(function (x) { return x.image; }).length;
      status('✔ Excel: ' + res.ok + ' / ' + withUrl + ' photos' + (res.ok < withUrl ? ' — the rest could not be downloaded (the site blocks it); their links are in the Image column' : ''));
    }, function (e) { save(buildXlsx(r, []), fname('xlsx')); status('Excel saved without photos (' + e + ')'); }).then(function () { btn.disabled = false; });
  };
  ui('#cv').onclick = function () { var r = exportRows(); if (!r.length) return status(t('empty')); save(new Blob([buildCsv(r)], { type: 'text/csv;charset=utf-8' }), fname('csv')); };
  ui('#cp').onclick = function () {
    var r = exportRows(); if (!r.length) return status(t('empty'));
    var tsv = [HEAD].concat(r.map(rowVals)).map(function (x) { return x.slice(1).map(function (v) { return v === null || v === undefined ? '' : String(v).replace(/\t|\n/g, ' '); }).join('\t'); }).join('\n');
    (navigator.clipboard ? navigator.clipboard.writeText(tsv) : Promise.reject()).then(function () { status(t('copied')); }, function () {
      var ta = document.createElement('textarea'); ta.value = tsv; document.body.appendChild(ta); ta.select(); document.execCommand('copy'); ta.remove(); status(t('copied'));
    });
  };
  ui('#rm').onclick = function () { S.rows = []; S.seen = {}; S.sel = {}; persist(); status(''); refresh(); };
  ui('#cl').onclick = function () { host.style.display = 'none'; };
  ui('#mn').onclick = function () { ui('#wrap').classList.toggle('min'); };
  ui('#lg').onclick = function () {
    S.lang = S.lang === 'en' ? 'fr' : S.lang === 'fr' ? 'ar' : 'en';
    try { localStorage.setItem('gmt_sc_lang2', S.lang); } catch (e) {}
    applyLang();
  };

  window.__gmtScraper = {
    open: function () { host.style.display = 'block'; },
    rows: function () { return S.rows; },
    scan: function () { return scanDoc(document, location.href); },
    version: VERSION, collectAll: collectAll, buildXlsx: buildXlsx, buildCsv: buildCsv, parsePrice: parsePrice, state: S
  };
  applyLang();
var first = scanDoc(document, location.href);
  status(first ? first + ' ' + t('found') + ' — ' + t('all') + ' ↑' : t('none'));
  refresh();
  setTimeout(function () { collectAll(); }, 500);
})();