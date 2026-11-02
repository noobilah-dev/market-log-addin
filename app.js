/* Market Log — Outlook add-in for logging market-situation info from emails.
   Written in plain ES5 so it also runs in older Outlook Web App / Internet Explorer.
   Everything stays in this browser: nothing is sent to any server. */
(function () {
  'use strict';

  // ---------------------------------------------------------------- categories
  var CATEGORIES = [
    { id: 'delay', name: 'Vessel delay / schedule change',
      kw: ['delay', 'delays', 'delayed', 'delaying', 'late arrival', 'arrive late', 'arriving late', 'behind schedule',
           'revised eta', 'eta revised', 'new eta', 'eta changed', 'eta change', 'eta postponed', 'postpone', 'postponed',
           'schedule change', 'schedule changes', 'schedule revised', 'schedule reliability', 'phase in', 'phase out',
           '延誤', '延遲', '延期', '晚到', '船期'] },
    { id: 'cancel', name: 'Shipment cancellation',
      kw: ['cancel', 'cancels', 'cancelled', 'canceled', 'cancelling', 'canceling', 'cancellation', 'withdraw',
           'withdrawn', 'call off', 'called off', 'not proceed', 'not proceeding', 'hold the shipment', 'on hold',
           '取消', '撤單', '退關'] },
    { id: 'space', name: 'Space / overbooking / rollover',
      kw: ['space', 'spaces', 'overbook', 'overbooked', 'overbooking', 'fully booked', 'full booked', 'allocation',
           'roll over', 'rolled over', 'rollover', 'rolled', 'shut out', 'shutout', 'shortshipped', 'short shipped',
           'short-shipped', 'left behind', 'no space', 'space tight', 'tight space',
           '艙位', '爆艙', '甩櫃', '滿艙'] },
    { id: 'blank', name: 'Blank sailing / service change',
      kw: ['blank sailing', 'blank sailings', 'void sailing', 'void sailings', 'skip', 'skipped', 'skipping', 'omit',
           'omitted', 'omission', 'port omission', 'suspend', 'suspended', 'suspension', 'service change',
           'new service', 'route change', 'rotation change', 'transhipment', 'transshipment', 'tranship',
           '跳港', '停航', '停班', '航線調整'] },
    { id: 'congestion', name: 'Port congestion / terminal',
      kw: ['congestion', 'congested', 'berthing', 'berth', 'anchorage', 'waiting time', 'waiting for berth',
           'yard full', 'terminal', 'productivity', 'port closure', 'port closed',
           '塞港', '擁塞', '壅塞', '碼頭', '等泊'] },
    { id: 'rate', name: 'Freight rate / surcharge',
      kw: ['rate', 'rates', 'freight', 'gri', 'pss', 'surcharge', 'surcharges', 'bunker', 'baf', 'ebs', 'caf',
           'price', 'pricing', 'increase', 'increased', 'decrease', 'decreased', 'cheaper', 'expensive', 'quotation',
           'quote', 'tariff', 'competitor', 'competitors', 'competitive',
           '運價', '運費', '漲價', '降價', '附加費', '報價'] },
    { id: 'equipment', name: 'Equipment / container shortage',
      kw: ['equipment', 'container shortage', 'box shortage', 'shortage', 'no empty', 'empty container',
           'empty containers', 'empties', 'reefer plug', 'reefer plugs', 'chassis',
           '缺櫃', '空櫃', '櫃子不足'] },
    { id: 'demand', name: 'Cargo demand / volume / market',
      kw: ['volume', 'volumes', 'demand', 'market', 'peak season', 'slack season', 'slow season', 'slowdown',
           'booking drop', 'bookings', 'forecast', 'orders', 'inventory', 'tariffs', 'trade war', 'economy',
           '市況', '貨量', '需求', '旺季', '淡季'] },
    { id: 'disruption', name: 'Weather / strike / geopolitical',
      kw: ['typhoon', 'storm', 'weather', 'monsoon', 'fog', 'strike', 'strikes', 'protest', 'war', 'conflict',
           'red sea', 'suez', 'hormuz', 'middle east', 'panama canal', 'cape of good hope', 'force majeure',
           'piracy', 'sanction', 'sanctions', 'holiday', 'chinese new year', 'cny', 'golden week',
           '颱風', '罷工', '戰爭', '紅海', '中東', '天氣', '假期', '春節'] },
    { id: 'customs', name: 'Customs / regulation',
      kw: ['customs', 'inspection', 'inspected', 'regulation', 'regulations', 'permit', 'detained', 'detention by customs',
           'held by customs', 'customs hold', '海關', '查驗', '法規'] }
  ];
  var CAT_BY_ID = {};
  var i, j;
  for (i = 0; i < CATEGORIES.length; i++) {
    var c = CATEGORIES[i];
    c.res = [];
    for (j = 0; j < c.kw.length; j++) {
      var k = c.kw[j];
      var src = k.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/ /g, '\\s+');
      c.res.push(/[a-z]/i.test(k) ? new RegExp('(^|[^a-z0-9])' + src + '(?![a-z0-9])', 'i') : new RegExp(src));
    }
    CAT_BY_ID[c.id] = c;
  }

  // --------------------------------------------------------------- locations
  var TLD_COUNTRY = {
    tw: 'Taiwan', cn: 'China', hk: 'Hong Kong', mo: 'Macau', jp: 'Japan', kr: 'Korea', sg: 'Singapore',
    my: 'Malaysia', th: 'Thailand', vn: 'Vietnam', id: 'Indonesia', ph: 'Philippines', kh: 'Cambodia',
    mm: 'Myanmar', 'in': 'India', lk: 'Sri Lanka', pk: 'Pakistan', bd: 'Bangladesh', ae: 'UAE', sa: 'Saudi Arabia',
    qa: 'Qatar', om: 'Oman', kw: 'Kuwait', bh: 'Bahrain', eg: 'Egypt', tr: 'Turkey', il: 'Israel', za: 'South Africa',
    ke: 'Kenya', ng: 'Nigeria', au: 'Australia', nz: 'New Zealand', uk: 'UK', gb: 'UK', ie: 'Ireland', de: 'Germany',
    nl: 'Netherlands', be: 'Belgium', fr: 'France', it: 'Italy', es: 'Spain', pt: 'Portugal', pl: 'Poland',
    se: 'Sweden', dk: 'Denmark', no: 'Norway', fi: 'Finland', gr: 'Greece', ru: 'Russia', us: 'USA', ca: 'Canada',
    mx: 'Mexico', br: 'Brazil', ar: 'Argentina', cl: 'Chile', pe: 'Peru', co: 'Colombia', pa: 'Panama'
  };
  var CITIES = ['Taipei', 'Kaohsiung', 'Keelung', 'Taichung', 'Taoyuan', 'Shanghai', 'Ningbo', 'Shenzhen', 'Shekou',
    'Yantian', 'Guangzhou', 'Nansha', 'Qingdao', 'Xiamen', 'Tianjin', 'Xingang', 'Dalian', 'Fuzhou', 'Nanjing',
    'Lianyungang', 'Zhuhai', 'Hong Kong', 'Tokyo', 'Yokohama', 'Osaka', 'Kobe', 'Nagoya', 'Hakata', 'Busan',
    'Seoul', 'Incheon', 'Port Klang', 'Kuala Lumpur', 'Penang', 'Pasir Gudang', 'Tanjung Pelepas', 'Johor',
    'Kuching', 'Kota Kinabalu', 'Jakarta', 'Surabaya', 'Semarang', 'Belawan', 'Bangkok', 'Laem Chabang',
    'Ho Chi Minh', 'Haiphong', 'Hai Phong', 'Hanoi', 'Da Nang', 'Cai Mep', 'Manila', 'Cebu', 'Phnom Penh',
    'Yangon', 'Chennai', 'Mumbai', 'Nhava Sheva', 'Mundra', 'Kolkata', 'Delhi', 'Colombo', 'Karachi', 'Chittagong',
    'Chattogram', 'Dhaka', 'Dubai', 'Jebel Ali', 'Abu Dhabi', 'Dammam', 'Jeddah', 'Doha', 'Sohar', 'Istanbul',
    'Sydney', 'Melbourne', 'Brisbane', 'Fremantle', 'Auckland', 'Los Angeles', 'Long Beach', 'Oakland', 'Seattle',
    'Tacoma', 'New York', 'New Jersey', 'Savannah', 'Houston', 'Charleston', 'Norfolk', 'Vancouver', 'Toronto',
    'Montreal', 'Rotterdam', 'Hamburg', 'Bremerhaven', 'Antwerp', 'Felixstowe', 'Southampton', 'London',
    'Le Havre', 'Genoa', 'La Spezia', 'Barcelona', 'Valencia', 'Piraeus', 'Gdansk', 'Durban', 'Cape Town',
    'Mombasa', 'Lagos', 'Santos', 'Manzanillo', 'Panama', 'Singapore'];

  var PERSONAL_DOMAINS = /^(gmail|googlemail|yahoo|hotmail|outlook|live|msn|icloud|me|qq|163|126|sina|sohu|naver|singnet|pacific|aol|yeah|foxmail)\./i;
  var SUFFIX_PARTS = /^(com|net|org|co|edu|gov|biz|info|ltd|plc|ac|or|ne|go)$/i;
  var COMPANY_LINE = /(pte\.?\s*ltd|sdn\.?\s*bhd|co\.,?\s*ltd|company\s+limited|\blimited\b|\bltd\b|\binc\b|\bcorp\b|corporation|\bllc\b|\bgmbh\b|\bb\.v\.|\bs\.a\.|logistics|shipping|freight|forwarding|trading|enterprise|industries|有限公司|股份)/i;
  var NOT_SIG_LINE = /(tel|fax|mobile|mob|hp|phone|email|e-mail|www\.|http|address|disclaimer|confidential|this e-?mail|intended|recipient|regards|thank|best|cheers|sent from)/i;

  // ------------------------------------------------------------------ storage
  var memStore = {};
  var storageOk = true;
  function sGet(key, fallback) {
    var raw = null;
    try { raw = window.localStorage.getItem(key); } catch (e) { storageOk = false; raw = memStore[key] || null; }
    if (!raw) { return fallback; }
    try { return JSON.parse(raw); } catch (e2) { return fallback; }
  }
  function sSet(key, val) {
    var raw = JSON.stringify(val);
    try { window.localStorage.setItem(key, raw); } catch (e) { storageOk = false; memStore[key] = raw; }
  }
  var K_ENTRIES = 'mlog.entries.v1', K_SETTINGS = 'mlog.settings.v1', K_CO = 'mlog.companyMap.v1', K_LOC = 'mlog.locMap.v1';
  function getSettings() {
    var s = sGet(K_SETTINGS, {});
    if (!s.egDomains) { s.egDomains = 'evergreen'; }
    return s;
  }
  function egWords() {
    var parts = getSettings().egDomains.split(','), out = [], x;
    for (x = 0; x < parts.length; x++) { var w = trim(parts[x]).toLowerCase(); if (w) { out.push(w); } }
    return out;
  }

  // ------------------------------------------------------------------ helpers
  function $(id) { return document.getElementById(id); }
  function trim(s) { return (s || '').replace(/^\s+|\s+$/g, ''); }
  function esc(s) {
    return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }
  function on(el, ev, fn) { if (el.addEventListener) { el.addEventListener(ev, fn, false); } else { el.attachEvent('on' + ev, fn); } }
  function show(el, yes) { if (yes) { el.className = el.className.replace(/\s*hidden/g, ''); } else if (!/\bhidden\b/.test(el.className)) { el.className += ' hidden'; } }
  function hasClass(el, c) { return new RegExp('(^|\\s)' + c + '(\\s|$)').test(el.className); }
  function toggleClass(el, c, yes) {
    if (yes && !hasClass(el, c)) { el.className += ' ' + c; }
    if (!yes) { el.className = el.className.replace(new RegExp('(^|\\s)' + c + '(?=\\s|$)', 'g'), ''); }
  }
  function pad2(n) { return n < 10 ? '0' + n : '' + n; }
  var MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  function fmtDate(d) { if (!d) { return ''; } d = new Date(d); if (isNaN(d.getTime())) { return ''; } return pad2(d.getDate()) + '-' + MONTHS[d.getMonth()] + '-' + d.getFullYear(); }
  function fmtDateTime(d) { d = new Date(d); return fmtDate(d) + ' ' + pad2(d.getHours()) + ':' + pad2(d.getMinutes()); }
  function titleCase(s) { return s.replace(/[-_]+/g, ' ').replace(/\b([a-z])/g, function (m) { return m.toUpperCase(); }); }
  function uid() { return 'e' + new Date().getTime().toString(36) + Math.floor(Math.random() * 1e6).toString(36); }

  function domainOf(email) { var m = /@([^>\s]+)/.exec(email || ''); return m ? m[1].toLowerCase() : ''; }
  function isEvergreen(email) {
    var d = domainOf(email), w = egWords(), x;
    if (!d) { return false; }
    for (x = 0; x < w.length; x++) { if (d.indexOf(w[x]) >= 0) { return true; } }
    return false;
  }
  function companyFromDomain(d) {
    if (!d || PERSONAL_DOMAINS.test(d)) { return ''; }
    var parts = d.split('.'), x;
    for (x = parts.length - 1; x >= 0; x--) {
      if (parts[x].length <= 2 || SUFFIX_PARTS.test(parts[x])) { continue; }
      return titleCase(parts[x]);
    }
    return '';
  }
  function countryFromDomain(d) {
    if (!d) { return ''; }
    var parts = d.split('.'), last = parts[parts.length - 1];
    return TLD_COUNTRY[last] || '';
  }
  function lastLines(text, n) {
    var lines = text.split(/\r?\n/), out = [], x;
    for (x = 0; x < lines.length; x++) { if (trim(lines[x])) { out.push(trim(lines[x])); } }
    return out.slice(Math.max(0, out.length - n));
  }
  // Returns { city, country } found in the signature (last lines of a message).
  function locationFromSignature(text) {
    var tail = lastLines(text, 25).join('\n'), out = { city: '', country: '' };
    var m = /evergreen[^\n(]{0,60}\(([A-Za-z .]{2,30})\)/i.exec(tail);
    if (m) { out.country = trim(m[1]); }
    var bestPos = -1, x;
    for (x = 0; x < CITIES.length; x++) {
      var re = new RegExp('(^|[^a-z])' + CITIES[x].replace(/ /g, '\\s*') + '(?![a-z])', 'i');
      var mm = re.exec(tail);
      if (mm && mm.index > bestPos) { out.city = CITIES[x]; bestPos = mm.index; }
    }
    return out;
  }
  function joinLoc(city, country) {
    if (city && country && city.toLowerCase() !== country.toLowerCase()) { return city + ', ' + country; }
    return city || country || '';
  }
  function companyFromSignature(text) {
    var lines = lastLines(text, 18), x;
    for (x = lines.length - 1; x >= 0; x--) {
      var l = lines[x];
      if (l.length < 4 || l.length > 90 || NOT_SIG_LINE.test(l) || /@/.test(l)) { continue; }
      if (/evergreen/i.test(l)) { continue; }
      if (COMPANY_LINE.test(l)) { return l.replace(/[|•]+.*$/, '').replace(/^[^A-Za-z0-9\u4e00-\u9fff]+/, ''); }
    }
    return '';
  }

  // ---------------------------------------------------------- email parsing
  var HEADER_RE = /^[ \t>]*(?:From|寄件者|發件人|发件人)\s*[:：][ \t]*(.+)$/gim;
  var OTHER_MARKERS = [/^[ \t>]*-{2,}\s*Original Message\s*-{2,}/im, /^[ \t>]*-{2,}\s*Forwarded message\s*-{2,}/im,
                       /^[ \t>]*On .{5,120} wrote:\s*$/im, /^[ \t>]*_{8,}\s*$/m];

  function parseFromLine(s) {
    s = trim(s.replace(/\s*\[mailto:[^\]]*\]/i, function (m) { return ' <' + m.replace(/^\s*\[mailto:|\]$/gi, '') + '>'; }));
    var m = /^"?([^"<]*?)"?\s*<([^>]+@[^>]+)>/.exec(s);
    if (m) { return { name: trim(m[1]) || m[2], email: trim(m[2]) }; }
    m = /([^\s<>"]+@[^\s<>"]+)/.exec(s);
    if (m) { var nm = trim(s.replace(m[1], '').replace(/[<>"()]/g, '')); return { name: nm || m[1], email: m[1] }; }
    return { name: trim(s.replace(/["']/g, '')), email: '' };
  }

  // Split the body into segments: [0] = the newest message, then one per quoted "From:" header.
  function splitThread(body) {
    var heads = [], m;
    HEADER_RE.lastIndex = 0;
    while ((m = HEADER_RE.exec(body)) !== null) { heads.push({ idx: m.index, end: m.index + m[0].length, from: parseFromLine(m[1]) }); }
    var firstCut = heads.length ? heads[0].idx : body.length, x;
    for (x = 0; x < OTHER_MARKERS.length; x++) { var mm = OTHER_MARKERS[x].exec(body); if (mm && mm.index < firstCut) { firstCut = mm.index; } }
    var segs = [{ from: null, text: body.slice(0, firstCut) }];
    for (x = 0; x < heads.length; x++) {
      var stop = x + 1 < heads.length ? heads[x + 1].idx : body.length;
      var txt = body.slice(heads[x].end, stop);
      // drop the Sent/To/Cc/Subject header lines at the top of the quoted block
      txt = txt.replace(/^(\s*(?:Sent|Date|To|Cc|Subject|寄件日期|傳送時間|收件者|副本|主旨|日期|收件人|抄送|主题)\s*[:：].*\r?\n)+/i, '');
      segs.push({ from: heads[x].from, text: txt });
    }
    return segs;
  }

  function sentencesOf(text) {
    var t = text.replace(/\r/g, '').replace(/^[ \t]*>+[ \t]?/gm, '');
    t = t.replace(/([.!?。！？;；])[ \t]+/g, '$1\n');
    var raw = t.split(/\n+/), out = [], x;
    for (x = 0; x < raw.length; x++) {
      var s = trim(raw[x]).replace(/\s+/g, ' ');
      s = trim(s.replace(/^(dear|hi|hello|hey|good (morning|afternoon|day))\b[^,，:]{0,40}[,，:]\s*/i, ''));
      if (s.length < 6 || s.length > 400) { continue; }
      if (/^(dear|hi|hello|regards|best regards|kind regards|thanks|thank you|tel|fax|mobile|email|sent:|to:|cc:|subject:)\b/i.test(s) && s.length < 40) { continue; }
      if (/^(tel|fax|mobile|email|sent|to|cc|subject)\s*[:：]/i.test(s)) { continue; }
      if (/(disclaimer|confidential|intended recipient|privileged|virus)/i.test(s)) { continue; }
      out.push(s);
    }
    return out;
  }
  function catsOfSentence(s) {
    var hit = [], x, y;
    for (x = 0; x < CATEGORIES.length; x++) {
      for (y = 0; y < CATEGORIES[x].res.length; y++) { if (CATEGORIES[x].res[y].test(s)) { hit.push(CATEGORIES[x].id); break; } }
    }
    return hit;
  }
  function findVessel(text) {
    var found = [], seen = {}, m;
    var re = /\b(EVER\s+[A-Z]{3,}(?:\s+[A-Z]{3,})?)(?:\s*(?:V\.?|VOY\.?|VOYAGE)?\s*(\d{3,4}-\d{2,3}[NSEWnsew]?))?/g;
    while ((m = re.exec(text)) !== null) {
      var v = trim(m[1].replace(/\s+(VOY|VOYAGE|V)$/i, '')) + (m[2] ? ' ' + m[2].toUpperCase() : '');
      var key = v.split(' ').slice(0, 2).join(' ').toUpperCase();
      if (/^EVER\s+(GREEN|GREENLINE|SINCE)$/i.test(m[1])) { continue; }
      if (!seen[key]) { seen[key] = 1; found.push(v); } else if (m[2]) {
        for (var x = 0; x < found.length; x++) { if (found[x].toUpperCase().indexOf(key) === 0 && found[x].length < v.length) { found[x] = v; } }
      }
    }
    if (!found.length) {
      var vv = /\b(\d{4}-\d{3}[NSEW])\b/.exec(text);
      if (vv) { found.push(vv[1]); }
    }
    return found.slice(0, 3).join(', ');
  }

  // Build a person record with best guesses for role, company and location.
  function personFrom(from, segText, me) {
    var email = from.email || '', d = domainOf(email);
    var coMap = sGet(K_CO, {}), locMap = sGet(K_LOC, {});
    var sigLoc = locationFromSignature(segText);
    var eg = email ? isEvergreen(email) : /evergreen/i.test(lastLines(segText, 12).join(' '));
    var p = { name: from.name || email, email: email, role: eg ? 'e' : (email ? 'c' : 'k') };
    if (me && email && email.toLowerCase() === me.toLowerCase()) { p.isMe = true; }
    if (p.role === 'e') {
      p.location = locMap[email.toLowerCase()] || joinLoc(sigLoc.city, sigLoc.country || countryFromDomain(d));
    } else {
      p.company = coMap[d] || companyFromSignature(segText) || companyFromDomain(d);
      if (p.role === 'k') { p.location = joinLoc(sigLoc.city, sigLoc.country); }
    }
    return p;
  }

  function analyse(body, sender, me) {
    var segs = splitThread(body || ''), people = [], seen = {}, x, y;
    var picks = [], pickSeen = {}, catCount = {}, firstHitSeg = -1;
    for (x = 0; x < segs.length; x++) {
      var from = x === 0 ? sender : segs[x].from;
      var p = from ? personFrom(from, segs[x].text, me) : null;
      if (p && !p.isMe) {
        var key = (p.email || p.name).toLowerCase();
        if (!seen[key]) { seen[key] = p; people.push(p); }
        segs[x].person = seen[key];
      }
      if (p && p.isMe) { segs[x].hit = false; continue; } // skip what I wrote myself
      var sents = sentencesOf(segs[x].text), segHit = false;
      for (y = 0; y < sents.length; y++) {
        var cs = catsOfSentence(sents[y]);
        if (!cs.length) { continue; }
        segHit = true;
        for (var z = 0; z < cs.length; z++) { catCount[cs[z]] = (catCount[cs[z]] || 0) + 1; }
        var sk = sents[y].toLowerCase();
        if (!pickSeen[sk] && picks.length < 5) { pickSeen[sk] = 1; picks.push(sents[y]); }
      }
      segs[x].hit = segHit;
      if (segHit && firstHitSeg < 0) { firstHitSeg = x; }
    }
    // Who said it: the sender of the newest message, plus the first person of the other side
    // who actually wrote market-related sentences further down the thread.
    var cust = null, eg = null;
    function assign(p) {
      if (!p || p.isMe) { return; }
      if (p.role === 'e' && !eg) { eg = p; } else if (p.role !== 'e' && !cust) { cust = p; }
    }
    if (segs[0].person) { assign(segs[0].person); }
    for (x = 1; x < segs.length; x++) { if (segs[x].hit && segs[x].person) { assign(segs[x].person); } }
    if (!segs[0].hit && firstHitSeg > 0 && segs[firstHitSeg].person) { assign(segs[firstHitSeg].person); }
    var cats = [];
    for (x = 0; x < CATEGORIES.length; x++) { if (catCount[CATEGORIES[x].id]) { cats.push(CATEGORIES[x].id); } }
    return { cats: cats, picks: picks, people: people, cust: cust, eg: eg, vessel: findVessel(body || '') };
  }

  // ------------------------------------------------------------------- state
  var current = { itemId: '', subject: '', date: null, sender: null, people: [], cats: {} };

  // --------------------------------------------------------------------- UI
  function renderCats() {
    var html = '', x;
    for (x = 0; x < CATEGORIES.length; x++) {
      var c = CATEGORIES[x];
      html += '<button class="chip' + (current.cats[c.id] ? ' on' : '') + '" data-cat="' + c.id + '">' + esc(c.name) + '</button>';
    }
    html += '<button class="chip' + (current.cats.other ? ' on' : '') + '" data-cat="other">Other</button>';
    $('cats').innerHTML = html;
  }
  function renderPeople() {
    var html = '', x;
    if (!current.people.length) { $('people').innerHTML = '<div class="hint">No one found — type the names below.</div>'; return; }
    for (x = 0; x < current.people.length; x++) {
      var p = current.people[x];
      var tag = p.role === 'e' ? '<span class="tag e">Evergreen</span>' : p.role === 'c' ? '<span class="tag c">Customer</span>' : '<span class="tag k">Not sure</span>';
      var meta = [p.email, p.company, p.location];
      var mt = [], y; for (y = 0; y < meta.length; y++) { if (meta[y]) { mt.push(esc(meta[y])); } }
      html += '<div class="person"><div class="pname">' + esc(p.name) + tag + '</div>' +
        '<div class="pmeta">' + mt.join(' · ') + '</div>' +
        '<div class="pbtns"><button data-p="' + x + '" data-as="c"' + (p.role !== 'e' ? ' class="guess"' : '') + '>Use as customer</button>' +
        '<button data-p="' + x + '" data-as="e"' + (p.role === 'e' ? ' class="guess"' : '') + '>Use as Evergreen colleague</button></div></div>';
    }
    $('people').innerHTML = html;
  }
  function fillCustomer(p) {
    $('fCustPic').value = p ? p.name : ''; $('fCustCo').value = p ? (p.company || '') : ''; $('fCustEmail').value = p ? p.email : '';
  }
  function fillEvergreen(p) {
    $('fEgName').value = p ? p.name : ''; $('fEgOffice').value = p ? (p.location || '') : ''; $('fEgEmail').value = p ? p.email : '';
  }
  function setStatus(msg) { $('status').innerHTML = msg; }

  function applyAnalysis(a) {
    current.people = a.people;
    current.cats = {};
    var x, names = [];
    for (x = 0; x < a.cats.length; x++) { current.cats[a.cats[x]] = true; names.push(CAT_BY_ID[a.cats[x]].name); }
    renderCats();
    renderPeople();
    fillCustomer(a.cust);
    fillEvergreen(a.eg);
    $('fVessel').value = a.vessel || '';
    $('fDetails').value = a.picks.join('\n');
    $('fRemarks').value = '';
    show($('saveMsg'), false);
    if (a.cats.length) {
      setStatus('<b>Market info found:</b> ' + esc(names.join(', ')) + '. Check the details and press <b>Save to log</b>.');
    } else {
      setStatus('No market-situation words found in this email. You can still tick a category and save it.');
    }
    // already logged?
    var entries = sGet(K_ENTRIES, []), prev = null;
    for (x = 0; x < entries.length; x++) { if (current.itemId && entries[x].itemId === current.itemId) { prev = entries[x]; } }
    if (prev) { $('dupWarn').innerHTML = 'You already logged this email on ' + esc(fmtDateTime(prev.loggedAt)) + '. Saving again adds a second row.'; }
    show($('dupWarn'), !!prev);
  }

  function readItem() {
    var mb = window.Office && Office.context && Office.context.mailbox;
    var item = mb && mb.item;
    if (!item) { setStatus('Open an email to log it.'); return; }
    current.itemId = item.itemId || '';
    current.subject = item.subject || '';
    current.date = item.dateTimeCreated || new Date();
    var f = item.from || item.sender;
    current.sender = f ? { name: f.displayName || f.emailAddress, email: f.emailAddress || '' } : null;
    current.me = (mb.userProfile && mb.userProfile.emailAddress) || '';
    if (item.body && item.body.getAsync) {
      item.body.getAsync(Office.CoercionType.Text, function (res) {
        if (res.status === Office.AsyncResultStatus.Succeeded) {
          show($('pasteBox'), false);
          applyAnalysis(analyse(res.value || '', current.sender, current.me));
        } else { needPaste(); }
      });
    } else { needPaste(); }
  }
  function needPaste() {
    show($('pasteBox'), true);
    applyAnalysis(analyse('', current.sender, current.me));
    setStatus('Your Outlook version does not let add-ins read the email text directly. Paste it below (Ctrl+A, Ctrl+C in the email).');
  }

  function save() {
    var cats = [], x;
    for (x = 0; x < CATEGORIES.length; x++) { if (current.cats[CATEGORIES[x].id]) { cats.push(CATEGORIES[x].name); } }
    if (current.cats.other) { cats.push('Other'); }
    var e = {
      id: uid(), itemId: current.itemId, subject: current.subject,
      emailDate: current.date ? new Date(current.date).getTime() : null, loggedAt: new Date().getTime(),
      categories: cats,
      custPic: trim($('fCustPic').value), custCo: trim($('fCustCo').value), custEmail: trim($('fCustEmail').value),
      egName: trim($('fEgName').value), egOffice: trim($('fEgOffice').value), egEmail: trim($('fEgEmail').value),
      vessel: trim($('fVessel').value), details: trim($('fDetails').value), remarks: trim($('fRemarks').value)
    };
    if (!cats.length) { setStatus('<b>Tick at least one category</b> before saving.'); return; }
    if (!e.custPic && !e.custCo && !e.egName && !e.egOffice) { setStatus('<b>Fill in who said it</b> (customer or Evergreen colleague) before saving.'); return; }
    var entries = sGet(K_ENTRIES, []);
    entries.push(e);
    sSet(K_ENTRIES, entries);
    // learn for next time
    if (e.custEmail && e.custCo) { var co = sGet(K_CO, {}); var d = domainOf(e.custEmail); if (d && !PERSONAL_DOMAINS.test(d)) { co[d] = e.custCo; sSet(K_CO, co); } }
    if (e.egEmail && e.egOffice) { var lo = sGet(K_LOC, {}); lo[e.egEmail.toLowerCase()] = e.egOffice; sSet(K_LOC, lo); }
    $('saveMsg').innerHTML = 'Saved. You now have ' + entries.length + ' log' + (entries.length === 1 ? '' : 's') + '.';
    show($('saveMsg'), true);
    show($('storageWarn'), !storageOk);
    updateCount();
  }

  // --------------------------------------------------------------- list tab
  function updateCount() { $('count').innerHTML = sGet(K_ENTRIES, []).length; }
  function infoFrom(e) {
    var c = e.custPic || e.custCo, g = e.egName || e.egOffice;
    return c && g ? 'Customer + Evergreen' : c ? 'Customer' : 'Evergreen';
  }
  function renderList() {
    var entries = sGet(K_ENTRIES, []), cat = $('filterCat').value, q = trim($('search').value).toLowerCase();
    var html = '', shown = 0, x;
    for (x = entries.length - 1; x >= 0; x--) {
      var e = entries[x];
      if (cat && e.categories.join('|').indexOf(cat) < 0) { continue; }
      if (q) {
        var hay = [e.custPic, e.custCo, e.egName, e.egOffice, e.vessel, e.details, e.subject, e.remarks].join(' ').toLowerCase();
        if (hay.indexOf(q) < 0) { continue; }
      }
      shown++;
      var who = [];
      if (e.custPic || e.custCo) { who.push(esc(e.custPic) + (e.custCo ? ' (' + esc(e.custCo) + ')' : '')); }
      if (e.egName || e.egOffice) { who.push('EMC ' + esc(e.egOffice || '') + (e.egName ? ' – ' + esc(e.egName) : '')); }
      var cs = '', y; for (y = 0; y < e.categories.length; y++) { cs += '<span>' + esc(e.categories[y]) + '</span>'; }
      html += '<div class="entry"><div class="eh"><span>' + esc(fmtDate(e.emailDate)) + (e.vessel ? ' · ' + esc(e.vessel) : '') + '</span>' +
        '<button class="del" data-del="' + e.id + '">Delete</button></div>' +
        '<div class="ecats">' + cs + '</div><div class="who">' + who.join(' / ') + '</div>' +
        '<div class="det">' + esc(e.details) + '</div>' + (e.remarks ? '<div class="hint">Remarks: ' + esc(e.remarks) + '</div>' : '') + '</div>';
    }
    if (!shown) { html = '<div class="empty">' + (entries.length ? 'Nothing matches the filter.' : 'No logs yet. Open an email and press Save to log.') + '</div>'; }
    $('list').innerHTML = html;
  }
  var pendingDelete = null;
  function onListClick(ev) {
    var t = ev.target || ev.srcElement, id = t.getAttribute && t.getAttribute('data-del');
    if (!id) { return; }
    if (pendingDelete !== id) { pendingDelete = id; t.innerHTML = 'Tap again to delete'; return; }
    var entries = sGet(K_ENTRIES, []), out = [], x;
    for (x = 0; x < entries.length; x++) { if (entries[x].id !== id) { out.push(entries[x]); } }
    sSet(K_ENTRIES, out); pendingDelete = null; renderList(); updateCount();
  }

  // ---------------------------------------------------------------- export
  var COLS = ['No.', 'Email date', 'Category', 'Info from', 'Customer PIC', 'Customer company', 'Evergreen colleague',
              'Evergreen office / location', 'Vessel / voyage', 'What they said', 'Email subject', 'My remarks', 'Logged on'];
  var WIDTHS = [5, 12, 26, 18, 20, 26, 20, 20, 22, 70, 40, 30, 17];
  function rowsForExport() {
    var entries = sGet(K_ENTRIES, []).slice(0), rows = [], x;
    entries.sort(function (a, b) { return (a.emailDate || a.loggedAt) - (b.emailDate || b.loggedAt); });
    for (x = 0; x < entries.length; x++) {
      var e = entries[x];
      rows.push([x + 1, e.emailDate ? new Date(e.emailDate) : '', e.categories.join(', '), infoFrom(e), e.custPic, e.custCo,
                 e.egName, e.egOffice, e.vessel, e.details, e.subject, e.remarks, new Date(e.loggedAt)]);
    }
    return rows;
  }

  // Minimal .xlsx writer (uncompressed zip) — no external library needed.
  var CRC_TABLE = (function () {
    var t = [], n, k, c;
    for (n = 0; n < 256; n++) { c = n; for (k = 0; k < 8; k++) { c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; } t[n] = c >>> 0; }
    return t;
  }());
  function utf8(s) {
    var b = unescape(encodeURIComponent(s)), a = new Uint8Array(b.length), x;
    for (x = 0; x < b.length; x++) { a[x] = b.charCodeAt(x); }
    return a;
  }
  function crc32(a) { var c = 0xFFFFFFFF, x; for (x = 0; x < a.length; x++) { c = CRC_TABLE[(c ^ a[x]) & 0xFF] ^ (c >>> 8); } return (c ^ 0xFFFFFFFF) >>> 0; }
  function zip(files) {
    var parts = [], central = [], offset = 0, x;
    function u16(v) { return [v & 255, (v >>> 8) & 255]; }
    function u32(v) { return [v & 255, (v >>> 8) & 255, (v >>> 16) & 255, (v >>> 24) & 255]; }
    for (x = 0; x < files.length; x++) {
      var name = utf8(files[x].name), data = utf8(files[x].data), crc = crc32(data);
      var head = [].concat([0x50, 0x4b, 0x03, 0x04], u16(20), u16(0x0800), u16(0), u16(0), u16(0x21), u32(crc), u32(data.length), u32(data.length), u16(name.length), u16(0));
      parts.push(new Uint8Array(head), name, data);
      central.push(new Uint8Array([].concat([0x50, 0x4b, 0x01, 0x02], u16(20), u16(20), u16(0x0800), u16(0), u16(0), u16(0x21), u32(crc), u32(data.length), u32(data.length), u16(name.length), u16(0), u16(0), u16(0), u16(0), u32(0), u32(offset))), name);
      offset += head.length + name.length + data.length;
    }
    var cdSize = 0; for (x = 0; x < central.length; x++) { cdSize += central[x].length; }
    var end = new Uint8Array([].concat([0x50, 0x4b, 0x05, 0x06], u16(0), u16(0), u16(files.length), u16(files.length), u32(cdSize), u32(offset), u16(0)));
    return parts.concat(central, [end]);
  }
  function xmlEsc(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, ''); }
  function colName(n) { var s = ''; n++; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
  function excelSerial(d) { return (d.getTime() - d.getTimezoneOffset() * 60000) / 86400000 + 25569; }
  function buildXlsx(rows) {
    var sheet = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>';
    var x, y;
    for (x = 0; x < WIDTHS.length; x++) { sheet += '<col min="' + (x + 1) + '" max="' + (x + 1) + '" width="' + WIDTHS[x] + '" customWidth="1"/>'; }
    sheet += '</cols><sheetData><row r="1">';
    for (x = 0; x < COLS.length; x++) { sheet += '<c r="' + colName(x) + '1" t="inlineStr" s="1"><is><t>' + xmlEsc(COLS[x]) + '</t></is></c>'; }
    sheet += '</row>';
    for (y = 0; y < rows.length; y++) {
      var r = y + 2; sheet += '<row r="' + r + '">';
      for (x = 0; x < rows[y].length; x++) {
        var v = rows[y][x], ref = colName(x) + r;
        if (v instanceof Date) { sheet += '<c r="' + ref + '" s="' + (x === 12 ? 4 : 2) + '"><v>' + excelSerial(v) + '</v></c>'; }
        else if (typeof v === 'number') { sheet += '<c r="' + ref + '" s="3"><v>' + v + '</v></c>'; }
        else if (v) { sheet += '<c r="' + ref + '" t="inlineStr" s="3"><is><t xml:space="preserve">' + xmlEsc(v) + '</t></is></c>'; }
      }
      sheet += '</row>';
    }
    sheet += '</sheetData><autoFilter ref="A1:' + colName(COLS.length - 1) + (rows.length + 1) + '"/></worksheet>';
    var styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<numFmts count="2"><numFmt numFmtId="164" formatCode="dd-mmm-yyyy"/><numFmt numFmtId="165" formatCode="dd-mmm-yyyy hh:mm"/></numFmts>' +
      '<fonts count="2"><font><sz val="10"/><name val="Arial"/></font><font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Arial"/></font></fonts>' +
      '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF007A3D"/><bgColor indexed="64"/></patternFill></fill></fills>' +
      '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right><top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom><diagonal/></border></borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="5">' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
      '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>' +
      '<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="top"/></xf>' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>' +
      '<xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="top"/></xf>' +
      '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
    return zip([
      { name: '[Content_Types].xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>' },
      { name: '_rels/.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>' },
      { name: 'xl/workbook.xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Market Log" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">\'Market Log\'!$A$1:$' + colName(COLS.length - 1) + '$' + (rows.length + 1) + '</definedName></definedNames></workbook>' },
      { name: 'xl/_rels/workbook.xml.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>' },
      { name: 'xl/worksheets/sheet1.xml', data: sheet },
      { name: 'xl/styles.xml', data: styles }
    ]);
  }
  function download() {
    var rows = rowsForExport();
    if (!rows.length) { $('list').innerHTML = '<div class="empty">No logs to download yet.</div>'; return; }
    var blob = new Blob(buildXlsx(rows), { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    var d = new Date(), name = 'Market Log ' + d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate()) + '.xlsx';
    if (window.navigator && window.navigator.msSaveOrOpenBlob) { window.navigator.msSaveOrOpenBlob(blob, name); return; }
    var a = document.createElement('a');
    a.href = (window.URL || window.webkitURL).createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click();
    setTimeout(function () { document.body.removeChild(a); }, 1000);
  }
  function copyForExcel() {
    var rows = rowsForExport(), lines = [COLS.join('\t')], x, y;
    for (x = 0; x < rows.length; x++) {
      var cells = [];
      for (y = 0; y < rows[x].length; y++) {
        var v = rows[x][y];
        if (v instanceof Date) { v = y === 12 ? fmtDateTime(v) : fmtDate(v); }
        v = String(v == null ? '' : v);
        if (/[\t\n"]/.test(v)) { v = '"' + v.replace(/"/g, '""').replace(/\t/g, ' ') + '"'; }
        cells.push(v);
      }
      lines.push(cells.join('\t'));
    }
    var ta = $('copyArea'); ta.value = lines.join('\r\n'); show(ta, true);
    ta.focus(); ta.select();
    try { document.execCommand('copy'); } catch (e) { /* user can copy manually */ }
    show($('copyHelp'), true);
  }

  // --------------------------------------------------------------- settings
  function renderSettings() {
    $('sEgDomains').value = getSettings().egDomains;
    function mapHtml(m, kind) {
      var keys = [], k, html = '', x;
      for (k in m) { if (m.hasOwnProperty(k)) { keys.push(k); } }
      keys.sort();
      for (x = 0; x < keys.length; x++) { html += '<div class="m"><span>' + esc(keys[x]) + ' → <b>' + esc(m[keys[x]]) + '</b></span><button data-kind="' + kind + '" data-key="' + esc(keys[x]) + '">Remove</button></div>'; }
      return html || '<div class="hint">Nothing yet.</div>';
    }
    $('sCoList').innerHTML = mapHtml(sGet(K_CO, {}), 'co');
    $('sLocList').innerHTML = mapHtml(sGet(K_LOC, {}), 'loc');
  }
  function onMapClick(ev) {
    var t = ev.target || ev.srcElement, kind = t.getAttribute && t.getAttribute('data-kind');
    if (!kind) { return; }
    var key = kind === 'co' ? K_CO : K_LOC, m = sGet(key, {});
    delete m[t.getAttribute('data-key')]; sSet(key, m); renderSettings();
  }

  // ------------------------------------------------------------------- tabs
  function switchTab(name) {
    var tabs = document.querySelectorAll('.tab'), x;
    for (x = 0; x < tabs.length; x++) { toggleClass(tabs[x], 'active', tabs[x].getAttribute('data-tab') === name); }
    show($('tab-log'), name === 'log'); show($('tab-list'), name === 'list'); show($('tab-settings'), name === 'settings');
    if (name === 'list') { renderList(); }
    if (name === 'settings') { renderSettings(); }
  }

  function wire() {
    var tabs = document.querySelectorAll('.tab'), x;
    for (x = 0; x < tabs.length; x++) { on(tabs[x], 'click', function (ev) { switchTab((ev.target || ev.srcElement).getAttribute('data-tab') || 'list'); }); }
    on($('cats'), 'click', function (ev) {
      var t = ev.target || ev.srcElement, id = t.getAttribute && t.getAttribute('data-cat');
      if (!id) { return; }
      current.cats[id] = !current.cats[id]; toggleClass(t, 'on', current.cats[id]);
    });
    on($('people'), 'click', function (ev) {
      var t = ev.target || ev.srcElement, idx = t.getAttribute && t.getAttribute('data-p');
      if (idx === null || idx === undefined || idx === '') { return; }
      var p = current.people[+idx];
      if (t.getAttribute('data-as') === 'e') {
        var loc = p.location || sGet(K_LOC, {})[(p.email || '').toLowerCase()] || countryFromDomain(domainOf(p.email));
        fillEvergreen({ name: p.name, email: p.email, location: loc });
      } else {
        fillCustomer({ name: p.name, email: p.email, company: p.company || companyFromDomain(domainOf(p.email)) });
      }
    });
    on($('saveBtn'), 'click', save);
    on($('pasteGo'), 'click', function () { applyAnalysis(analyse($('pasteText').value, current.sender, current.me)); });
    on($('dlBtn'), 'click', download);
    on($('copyBtn'), 'click', copyForExcel);
    on($('filterCat'), 'change', renderList);
    on($('search'), 'keyup', renderList);
    on($('list'), 'click', onListClick);
    on($('sCoList'), 'click', onMapClick);
    on($('sLocList'), 'click', onMapClick);
    on($('sSave'), 'click', function () {
      var s = getSettings(); s.egDomains = trim($('sEgDomains').value) || 'evergreen'; sSet(K_SETTINGS, s);
      $('sMsg').innerHTML = 'Saved. Reopen the email to re-check it with the new setting.'; show($('sMsg'), true);
    });
    on($('bkCopy'), 'click', function () {
      $('bkArea').value = JSON.stringify({ entries: sGet(K_ENTRIES, []), co: sGet(K_CO, {}), loc: sGet(K_LOC, {}), settings: getSettings() });
      $('bkArea').focus(); $('bkArea').select();
      try { document.execCommand('copy'); } catch (e) { /* ignore */ }
      $('bkMsg').innerHTML = 'Backup code copied. Paste it into a Notepad file or email it to yourself.'; show($('bkMsg'), true);
    });
    on($('bkRestore'), 'click', function () {
      var data;
      try { data = JSON.parse($('bkArea').value); } catch (e) { $('bkMsg').innerHTML = 'That backup code is not valid.'; show($('bkMsg'), true); return; }
      var entries = sGet(K_ENTRIES, []), have = {}, added = 0, y, k;
      for (y = 0; y < entries.length; y++) { have[entries[y].id] = 1; }
      var inc = data.entries || [];
      for (y = 0; y < inc.length; y++) { if (inc[y] && inc[y].id && !have[inc[y].id]) { entries.push(inc[y]); added++; } }
      sSet(K_ENTRIES, entries);
      var co = sGet(K_CO, {}), lo = sGet(K_LOC, {});
      for (k in (data.co || {})) { if (data.co.hasOwnProperty(k) && !co[k]) { co[k] = data.co[k]; } }
      for (k in (data.loc || {})) { if (data.loc.hasOwnProperty(k) && !lo[k]) { lo[k] = data.loc[k]; } }
      sSet(K_CO, co); sSet(K_LOC, lo);
      $('bkMsg').innerHTML = 'Restored ' + added + ' log' + (added === 1 ? '' : 's') + '.'; show($('bkMsg'), true);
      updateCount(); renderSettings();
    });
    var opts = '<option value="">All categories</option>';
    for (x = 0; x < CATEGORIES.length; x++) { opts += '<option>' + esc(CATEGORIES[x].name) + '</option>'; }
    $('filterCat').innerHTML = opts + '<option>Other</option>';
  }

  function start() {
    wire();
    renderCats();
    updateCount();
    try { window.localStorage.setItem('mlog.test', '1'); window.localStorage.removeItem('mlog.test'); } catch (e) { storageOk = false; }
    show($('storageWarn'), !storageOk);
    readItem();
    try {
      if (Office.context.mailbox.addHandlerAsync && Office.EventType && Office.EventType.ItemChanged) {
        Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, function () { switchTab('log'); readItem(); });
      }
    } catch (e) { /* older Outlook: pane reloads per email anyway */ }
  }

  // expose parsing for testing
  window.MarketLog = { analyse: analyse, buildXlsx: buildXlsx, rowsForExport: rowsForExport };

  if (window.Office) {
    Office.initialize = function () { start(); };
  } else {
    on(window, 'load', function () { setStatus('Outlook is not available — open this inside Outlook.'); wire(); renderCats(); updateCount(); });
  }
}());
