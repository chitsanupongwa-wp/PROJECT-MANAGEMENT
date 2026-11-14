/**
 * Project To-Do Dashboard — Google Apps Script server
 * Data lives in the Google Sheet this script is attached to.
 * Deploy: Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone
 */

var TZ = 'Asia/Bangkok';
var SESSION_SECONDS = 21600; // 6 hours, renewed on every call

var T = {
  Users:    ['id','u','name','title','role','status','salt','hash','createdAt','lastLogin','photo'],
  Requests: ['id','u','name','title','salt','hash','at'],
  Projects: ['id','name','client','scope','start','finish','status','note','order','createdAt','links','weekly','photo'],
  Access:   ['projectId','userId'],
  Team:     ['name','role'],
  Tasks:    ['id','projectId','title','owner','due','priority','status','progress','done','doneAt','by','createdAt','note','checklist','comments'],
  Issues:   ['id','projectId','no','title','category','severity','location','owner','due','status','desc','resolution','resolvedAt','taskId','riskId','by','createdAt'],
  Risks:    ['id','projectId','no','title','category','L','I','response','plan','trigger','owner','review','status','history','createdAt'],
  Reports:  ['id','projectId','week','rev','status','summary','highlights','help','snap','by','createdBy','createdAt','submittedAt','approvedBy','approvedAt','returnReason'],
  Meetings: ['id','projectId','no','title','date','time','location','attendees','agenda','minutes','actions','status','by','createdAt'],
  Logins:   ['id','userId','u','name','at','type']
};
var KIDS = { tasks: 'Tasks', issues: 'Issues', risks: 'Risks', reports: 'Reports', meetings: 'Meetings' };
var JSONF = { checklist: 1, comments: 1, history: 1, snap: 1, actions: 1, links: 1, weekly: 1 };
var NUMF = { progress: 1, L: 1, I: 1, rev: 1, order: 1 };
var BOOLF = { done: 1 };

/* ---------- web app ---------- */
function doGet() {
  ensureSheets_();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Project To-Do Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------- JSON API (for the Vercel front end) ---------- */
var API_ = { bootstrap: 1, setupAdmin: 1, login: 1, signup: 1, logout: 1, ping: 1, getState: 1, saveProject: 1, deleteProject: 1, saveTeam: 1, removeTeamMember: 1, adminCreateUser: 1, approveRequest: 1, rejectRequest: 1, setUserTitle: 1, setUserAccess: 1, resetPassword: 1, changeMyPassword: 1, deleteUser: 1, updateMyProfile: 1 };
function doPost(e) {
  var out;
  try {
    ensureSheets_();
    var req = JSON.parse(e.postData.contents || '{}');
    if (!API_[req.fn]) throw new Error('Unknown function');
    out = { ok: true, data: this[req.fn].apply(null, req.args || []) };
  } catch (err) { out = { ok: false, error: String(err && err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* works both as a Sheet-bound script and as a standalone project */
var DB_CACHE_;
function db_() {
  if (DB_CACHE_) return DB_CACHE_;
  var ss = SpreadsheetApp.getActive();
  if (!ss) {
    var props = PropertiesService.getScriptProperties(), id = props.getProperty('DB_ID');
    if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
    if (!ss) { ss = SpreadsheetApp.create('PJM Database'); props.setProperty('DB_ID', ss.getId()); }
  }
  return (DB_CACHE_ = ss);
}

/* ---------- sheet helpers ---------- */
function sheet_(name) {
  var ss = db_();
  var s = ss.getSheetByName(name);
  var f = T[name];
  if (!s) {
    s = ss.insertSheet(name);
    s.getRange(1, 1, s.getMaxRows(), Math.max(f.length, 1)).setNumberFormat('@');
    s.getRange(1, 1, 1, f.length).setValues([f]).setFontWeight('bold');
    s.setFrozenRows(1);
  }
  return s;
}
function ensureSheets_() { Object.keys(T).forEach(sheet_); }
function dec_(k, v) {
  if (v instanceof Date) v = Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  if (JSONF[k]) { try { return v === '' ? (k === 'snap' ? null : []) : JSON.parse(v); } catch (e) { return k === 'snap' ? null : []; } }
  if (NUMF[k]) return v === '' ? 0 : Number(v);
  if (BOOLF[k]) return v === true || String(v).toLowerCase() === 'true';
  return v === null || v === undefined ? '' : String(v);
}
function enc_(k, v) {
  if (v === null || v === undefined) return '';
  if (JSONF[k] || typeof v === 'object') return JSON.stringify(v);
  return String(v);
}
function read_(name) {
  var s = sheet_(name), f = T[name], n = s.getLastRow() - 1;
  if (n < 1) return [];
  var v = s.getRange(2, 1, n, f.length).getValues();
  return v.filter(function (r) { return r[0] !== ''; }).map(function (r) {
    var o = {}; f.forEach(function (k, i) { o[k] = dec_(k, r[i]); }); return o;
  });
}
function write_(name, rows) {
  var s = sheet_(name), f = T[name], n = s.getLastRow() - 1;
  if (n > 0) s.getRange(2, 1, n, f.length).clearContent();
  if (rows.length) {
    var need = rows.length + 1 - s.getMaxRows();
    if (need > 0) s.insertRowsAfter(s.getMaxRows(), need);
    s.getRange(2, 1, rows.length, f.length).setNumberFormat('@')
      .setValues(rows.map(function (o) { return f.map(function (k) { return enc_(k, o[k]); }); }));
  }
}
function lock_(fn) {
  var l = LockService.getScriptLock(); l.waitLock(25000);
  try { return fn(); } finally { l.releaseLock(); }
}
function now_() { return new Date().toISOString(); }
function uid_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 12); }

/* ---------- passwords & sessions ---------- */
function hash_(pw, salt) {
  var h = salt + '|' + pw;
  for (var i = 0; i < 300; i++) h = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + salt, Utilities.Charset.UTF_8));
  return h;
}
function checkUsername_(u) {
  if (!/^[A-Za-z0-9._-]{3,30}$/.test(u || '')) throw new Error('Username must be 3–30 characters: letters, numbers, . _ -');
}
function checkPw_(pw) { if (!pw || String(pw).length < 6) throw new Error('Password must be at least 6 characters'); }
function taken_(u) {
  u = String(u).toLowerCase();
  return read_('Users').some(function (x) { return x.u.toLowerCase() === u; }) ||
         read_('Requests').some(function (x) { return x.u.toLowerCase() === u; });
}
function newSession_(user) {
  var t = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put('tok_' + t, user.id, SESSION_SECONDS);
  return t;
}
function me_(token) {
  var c = CacheService.getScriptCache(), id = token && c.get('tok_' + token);
  if (!id) throw new Error('SESSION_EXPIRED');
  var u = read_('Users').filter(function (x) { return x.id === id && x.status !== 'inactive'; })[0];
  if (!u) throw new Error('SESSION_EXPIRED');
  c.put('tok_' + token, id, SESSION_SECONDS);
  return u;
}
function adminOnly_(u) { if (u.role !== 'admin') throw new Error('Only the Project Manager can do this'); }
function pub_(u) { return { id: u.id, u: u.u, name: u.name, title: u.title, role: u.role }; }

/* ---------- public calls ---------- */
function bootstrap() { ensureSheets_(); return { needsSetup: read_('Users').length === 0 }; }

function setupAdmin(u, name, pw) {
  return lock_(function () {
    if (read_('Users').length) throw new Error('Setup is already done — please sign in');
    checkUsername_(u); checkPw_(pw);
    var salt = uid_() + uid_();
    var user = { id: uid_(), u: u, name: name || u, title: 'Project Manager', role: 'admin', status: 'active', salt: salt, hash: hash_(pw, salt), createdAt: now_(), lastLogin: now_() };
    write_('Users', [user]);
    var team = read_('Team'); team.push({ name: user.name, role: 'Project Manager' }); write_('Team', team);
    logLogin_(user, 'in');
    return { token: newSession_(user), state: state_(user) };
  });
}

function login(u, pw) {
  var c = CacheService.getScriptCache(), k = 'fail_' + String(u).toLowerCase(), f = Number(c.get(k) || 0);
  if (f >= 8) throw new Error('Too many attempts — try again in 15 minutes');
  var users = read_('Users');
  var user = users.filter(function (x) { return x.u.toLowerCase() === String(u).trim().toLowerCase(); })[0];
  if (!user || user.status === 'inactive' || hash_(pw, user.salt) !== user.hash) {
    if (!user && read_('Requests').some(function (x) { return x.u.toLowerCase() === String(u).trim().toLowerCase(); }))
      throw new Error('This account is waiting for approval');
    c.put(k, String(f + 1), 900);
    throw new Error('Incorrect username or password');
  }
  c.remove(k);
  lock_(function () { var all = read_('Users'); all.forEach(function (x) { if (x.id === user.id) x.lastLogin = now_(); }); write_('Users', all); logLogin_(user, 'in'); });
  user.lastLogin = now_();
  return { token: newSession_(user), state: state_(user) };
}

function signup(u, name, title, pw) {
  return lock_(function () {
    checkUsername_(u); checkPw_(pw);
    if (!name) throw new Error('Please enter your name');
    if (taken_(u)) throw new Error('This username is already taken or pending approval');
    var salt = uid_() + uid_();
    var r = read_('Requests');
    r.push({ id: uid_(), u: u, name: name, title: title || '', salt: salt, hash: hash_(pw, salt), at: now_() });
    write_('Requests', r);
    return true;
  });
}

function logout(token) {
  if (!token) return true;
  try { var u = me_(token); setPresence_(u.id, null); lock_(function () { logLogin_(u, 'out'); }); } catch (e) {}
  CacheService.getScriptCache().remove('tok_' + token); return true;
}

/* ---------- presence & login log ---------- */
var ONLINE_MS = 150000;
function logLogin_(u, type) {
  var L = read_('Logins'); L.push({ id: uid_(), userId: u.id, u: u.u, name: u.name, at: now_(), type: type });
  if (L.length > 500) L = L.slice(L.length - 500);
  write_('Logins', L);
}
function getPresence_() { try { return JSON.parse(CacheService.getScriptCache().get('presence') || '{}'); } catch (e) { return {}; } }
function setPresence_(id, ts) {
  var c = CacheService.getScriptCache(), m = getPresence_(), now = Date.now();
  if (ts) m[id] = ts; else delete m[id];
  Object.keys(m).forEach(function (k) { if (now - m[k] > 3600000) delete m[k]; });
  c.put('presence', JSON.stringify(m), 21600);
}
function presence_(user, users) {
  var m = getPresence_(), last = m[user.id] || 0, now = Date.now();
  if (now - last > 30000) { setPresence_(user.id, now); m[user.id] = now; }
  return users.filter(function (u) { return u.status !== 'inactive'; }).map(function (u) {
    return { u: u.u, name: u.name, title: u.title, online: !!m[u.id] && now - m[u.id] < ONLINE_MS, seen: m[u.id] ? new Date(m[u.id]).toISOString() : '', lastLogin: u.lastLogin || '' };
  });
}
function ping(token) { var u = me_(token); return { presence: presence_(u, read_('Users')), logins: u.role === 'admin' ? read_('Logins').slice(-100).reverse().map(function (l) { return { u: l.u, name: l.name, at: l.at, type: l.type }; }) : [] }; }

/* ---------- state ---------- */
function visibleIds_(user) {
  var all = read_('Projects').map(function (p) { return p.id; });
  if (user.role === 'admin') return all;
  var mine = {}; read_('Access').forEach(function (a) { if (a.userId === user.id) mine[a.projectId] = 1; });
  return all.filter(function (id) { return mine[id]; });
}
function project_(p, kids, access, users) {
  var o = JSON.parse(JSON.stringify(p));
  Object.keys(KIDS).forEach(function (k) { o[k] = kids[k].filter(function (x) { return x.projectId === p.id; }).map(function (x) { var y = JSON.parse(JSON.stringify(x)); delete y.projectId; return y; }); });
  if (access) {
    var byId = {}; users.forEach(function (u) { byId[u.id] = u; });
    o.access = access.filter(function (a) { return a.projectId === p.id && byId[a.userId]; }).map(function (a) { return byId[a.userId].u; });
  }
  return o;
}
function state_(user) {
  var ids = visibleIds_(user), set = {}; ids.forEach(function (i) { set[i] = 1; });
  var kids = {}; Object.keys(KIDS).forEach(function (k) { kids[k] = read_(KIDS[k]).filter(function (x) { return set[x.projectId]; }); });
  var isA = user.role === 'admin', users = read_('Users'), access = isA ? read_('Access') : null;
  var projects = read_('Projects').filter(function (p) { return set[p.id]; })
    .sort(function (a, b) { return (a.order || 0) - (b.order || 0) || String(a.createdAt).localeCompare(String(b.createdAt)); })
    .map(function (p) { return project_(p, kids, access, users); });
  return {
    me: pub_(user),
    team: read_('Team'),
    projects: projects,
    users: isA ? users.filter(function (u) { return u.status !== 'inactive'; }).map(pub_) : users.filter(function (u) { return u.status !== 'inactive'; }).map(function (u) { return { u: u.u, name: u.name, title: u.title, role: u.role }; }),
    requests: isA ? read_('Requests').map(function (r) { return { id: r.id, u: r.u, name: r.name, role: r.title, at: r.at }; }) : [],
    presence: presence_(user, users),
    logins: isA ? read_('Logins').slice(-100).reverse().map(function (l) { return { u: l.u, name: l.name, at: l.at, type: l.type }; }) : [],
    photos: photos_(users),
    serverTime: now_()
  };
}
function getState(token) { return state_(me_(token)); }

/* ---------- saving ---------- */
function saveProject(token, p, loadedAt) {
  var user = me_(token);
  return lock_(function () {
    var projects = read_('Projects'), cur = projects.filter(function (x) { return x.id === p.id; })[0];
    if (!cur) {
      adminOnly_(user);
      cur = { id: p.id || uid_(), createdAt: now_(), order: projects.length };
      projects.push(cur);
    } else if (user.role !== 'admin' && visibleIds_(user).indexOf(p.id) < 0) {
      throw new Error('You do not have access to this project');
    }
    if (user.role === 'admin') { ['name','client','scope','start','finish','status','note'].forEach(function (k) { cur[k] = p[k] || ''; }); if (p.photo !== undefined) { var ph = String(p.photo || ''); if (!ph || (/^data:image\/(jpeg|png|webp);base64,[A-Za-z0-9+\/=]+$/.test(ph) && ph.length <= 48000)) cur.photo = ph; } cur.weekly = Array.isArray(p.weekly) ? p.weekly.filter(function (w) { return w && /^\d{4}-\d{2}-\d{2}$/.test(String(w.week || '')); }).slice(-260).map(function (w) { var n = function (v) { v = Number(v); return isFinite(v) ? Math.max(0, Math.min(100, Math.round(v * 10) / 10)) : null; }; return { week: String(w.week), plan: n(w.plan), actual: n(w.actual), note: String(w.note || '').slice(0, 300) }; }) : (cur.weekly || []); cur.links = Array.isArray(p.links) ? p.links.filter(function (l) { return l && /^https?:\/\//i.test(String(l.url || '')); }).slice(0, 20).map(function (l) { return { label: String(l.label || 'Link').slice(0, 60), url: String(l.url).slice(0, 2000) }; }) : (cur.links || []); }
    write_('Projects', projects);
    Object.keys(KIDS).forEach(function (k) {
      var all = read_(KIDS[k]);
      var server = all.filter(function (x) { return x.projectId === cur.id; });
      var others = all.filter(function (x) { return x.projectId !== cur.id; });
      var client = (p[k] || []).map(function (x) { var y = JSON.parse(JSON.stringify(x)); y.projectId = cur.id; return y; });
      var ids = {}; client.forEach(function (x) { ids[x.id] = 1; });
      // keep items someone else added after this user loaded the page
      var keep = server.filter(function (x) { return !ids[x.id] && String(x.createdAt || '') > String(loadedAt || ''); });
      write_(KIDS[k], others.concat(client, keep));
    });
    var st = state_(user);
    return { project: st.projects.filter(function (x) { return x.id === cur.id; })[0], serverTime: st.serverTime };
  });
}
function deleteProject(token, id) {
  var user = me_(token); adminOnly_(user);
  return lock_(function () {
    write_('Projects', read_('Projects').filter(function (x) { return x.id !== id; }));
    write_('Access', read_('Access').filter(function (x) { return x.projectId !== id; }));
    Object.keys(KIDS).forEach(function (k) { write_(KIDS[k], read_(KIDS[k]).filter(function (x) { return x.projectId !== id; })); });
    return true;
  });
}
function saveTeam(token, team) {
  me_(token);
  return lock_(function () {
    var cur = read_('Team'), by = {};
    cur.forEach(function (m) { by[m.name] = m; });
    (team || []).forEach(function (m) { if (m && m.name) by[m.name] = { name: m.name, role: m.role || '' }; });
    write_('Team', Object.keys(by).map(function (k) { return by[k]; }));
    return true;
  });
}
function removeTeamMember(token, name) {
  var user = me_(token); adminOnly_(user);
  return lock_(function () { write_('Team', read_('Team').filter(function (m) { return m.name !== name; })); return true; });
}

/* ---------- users (Project Manager) ---------- */
function setAccess_(userId, projectIds) {
  var a = read_('Access').filter(function (x) { return x.userId !== userId; });
  (projectIds || []).forEach(function (pid) { a.push({ projectId: pid, userId: userId }); });
  write_('Access', a);
}
function addTeam_(name, role) {
  var t = read_('Team'), m = t.filter(function (x) { return x.name === name; })[0];
  if (m) m.role = role || m.role; else t.push({ name: name, role: role || '' });
  write_('Team', t);
}
function adminCreateUser(token, o) {
  var admin = me_(token); adminOnly_(admin);
  return lock_(function () {
    checkUsername_(o.u); checkPw_(o.pw);
    if (taken_(o.u)) throw new Error('Username already exists or is pending');
    var salt = uid_() + uid_();
    var user = { id: uid_(), u: o.u, name: o.name || o.u, title: o.title || '', role: o.role === 'admin' ? 'admin' : 'member', status: 'active', salt: salt, hash: hash_(o.pw, salt), createdAt: now_(), lastLogin: '' };
    var users = read_('Users'); users.push(user); write_('Users', users);
    if (user.role !== 'admin') setAccess_(user.id, o.projects || []);
    addTeam_(user.name, user.title);
    return pub_(user);
  });
}
function approveRequest(token, reqId, role, title) {
  var admin = me_(token); adminOnly_(admin);
  return lock_(function () {
    var reqs = read_('Requests'), r = reqs.filter(function (x) { return x.id === reqId; })[0];
    if (!r) throw new Error('Request not found');
    var user = { id: uid_(), u: r.u, name: r.name, title: title || r.title || '', role: role === 'admin' ? 'admin' : 'member', status: 'active', salt: r.salt, hash: r.hash, createdAt: now_(), lastLogin: '' };
    var users = read_('Users'); users.push(user); write_('Users', users);
    write_('Requests', reqs.filter(function (x) { return x.id !== reqId; }));
    addTeam_(user.name, user.title);
    return pub_(user);
  });
}
function rejectRequest(token, reqId) {
  var admin = me_(token); adminOnly_(admin);
  return lock_(function () { write_('Requests', read_('Requests').filter(function (x) { return x.id !== reqId; })); return true; });
}
function setUserTitle(token, userId, title) {
  var admin = me_(token); adminOnly_(admin);
  return lock_(function () {
    var users = read_('Users'), u = users.filter(function (x) { return x.id === userId; })[0];
    if (!u) throw new Error('User not found');
    u.title = title; write_('Users', users); addTeam_(u.name, title); return true;
  });
}
function setUserAccess(token, userId, projectIds) {
  var admin = me_(token); adminOnly_(admin);
  return lock_(function () { setAccess_(userId, projectIds); return true; });
}
function resetPassword(token, userId, pw) {
  var admin = me_(token); adminOnly_(admin); checkPw_(pw);
  return lock_(function () {
    var users = read_('Users'), u = users.filter(function (x) { return x.id === userId; })[0];
    if (!u) throw new Error('User not found');
    u.salt = uid_() + uid_(); u.hash = hash_(pw, u.salt); write_('Users', users); return true;
  });
}
function changeMyPassword(token, pw) {
  var me = me_(token); checkPw_(pw);
  return lock_(function () {
    var users = read_('Users'), u = users.filter(function (x) { return x.id === me.id; })[0];
    u.salt = uid_() + uid_(); u.hash = hash_(pw, u.salt); write_('Users', users); return true;
  });
}
/* ---------- my profile (name + photo) ---------- */
function photos_(users) { var m = {}; users.forEach(function (u) { if (u.status !== 'inactive' && u.photo) m[u.name] = u.photo; }); return m; }
function updateMyProfile(token, o) {
  var me = me_(token); o = o || {};
  var name = String(o.name == null ? me.name : o.name).trim().slice(0, 60);
  if (!name) throw new Error('Please enter your name');
  var photo = o.photo == null ? (me.photo || '') : String(o.photo);
  if (photo && (!/^data:image\/(jpeg|png|webp);base64,[A-Za-z0-9+\/=]+$/.test(photo) || photo.length > 45000)) throw new Error('Photo is too large or not an image');
  return lock_(function () {
    var users = read_('Users'), u = users.filter(function (x) { return x.id === me.id; })[0], old = u.name;
    if (name !== old && users.some(function (x) { return x.id !== me.id && x.name === name; })) throw new Error('This name is already used by another user');
    u.name = name; u.photo = photo; write_('Users', users);
    if (name !== old) {   // keep assignments pointing at the renamed person
      var team = read_('Team'); team.forEach(function (m) { if (m.name === old) m.name = name; }); write_('Team', team);
      ['Tasks', 'Issues', 'Risks'].forEach(function (sh) { var rows = read_(sh), ch = false; rows.forEach(function (r) { if (r.owner === old) { r.owner = name; ch = true; } }); if (ch) write_(sh, rows); });
      var ms = read_('Meetings'), mc = false;
      ms.forEach(function (m) { (m.actions || []).forEach(function (a) { if (a.owner === old) { a.owner = name; mc = true; } }); }); if (mc) write_('Meetings', ms);
    }
    return state_(u);
  });
}
function deleteUser(token, userId) {
  var admin = me_(token); adminOnly_(admin);
  if (userId === admin.id) throw new Error('You cannot delete yourself');
  return lock_(function () {
    write_('Users', read_('Users').filter(function (x) { return x.id !== userId; }));
    write_('Access', read_('Access').filter(function (x) { return x.userId !== userId; }));
    return true;
  });
}
