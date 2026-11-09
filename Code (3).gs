/**
 * మన ఊరు వేదిక – Google Apps Script బ్యాకెండ్
 * Deploy: Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone
 */
const SHEET_ID  = '12Wx-4wjiY2yaqPCCZDHP1ki7EJ1zIaPV3uNyNCEPk-o';
const FOLDER_ID = '1zVP1saTzNI6-OpvVyqIA3iS4lKZQRqu5';
const ADMIN_PHONE = '4056830968';
const ADMIN_PASS  = 'Dvv@1417';   // ఈ ఫైల్ ని ఎవరికీ షేర్ చేయకండి

const H = {
  Users:    ['phone','name','photo','passHash','secQ','secAHash','token','isAdmin','created'],
  Posts:    ['id','type','category','title','details','address','photo','name','phone','showId','likedBy','date'],
  Comments: ['id','postId','name','phone','text','date']
};

function sh(n) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let s = ss.getSheetByName(n);
  if (!s) {
    s = ss.insertSheet(n);
    s.appendRow(H[n]);
    s.getRange('A:Z').setNumberFormat('@');   // ఫోన్ నెంబర్లు టెక్స్ట్ గా
    s.setFrozenRows(1);
  }
  return s;
}
function rows(n) {
  const v = sh(n).getDataRange().getValues(), h = v.shift();
  return v.map((r, i) => { const o = { _row: i + 2 }; h.forEach((k, j) => o[k] = String(r[j])); return o; });
}
function setCell(n, row, key, val) { sh(n).getRange(row, H[n].indexOf(key) + 1).setValue(val); }
function add(n, o) { sh(n).appendRow(H[n].map(k => o[k] === undefined ? '' : String(o[k]))); }
function hash(s) { return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s)); }
function pw(phone, p) { return hash(phone + ':' + p); }
function now() { return Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd-MM-yyyy HH:mm'); }
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function upload(b64, name) {
  if (!b64) return '';
  const m = String(b64).match(/^data:(.+?);base64,(.+)$/);
  if (!m) return '';
  const f = DriveApp.getFolderById(FOLDER_ID).createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name));
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return f.getId();
}

function auth(d) {
  const u = rows('Users').find(x => x.phone === String(d.phone) && x.token && x.token === d.token);
  if (!u) throw new Error('మళ్ళీ లాగిన్ అవ్వండి');
  return u;
}
const pub = u => ({ name: u.name, phone: u.phone, photo: u.photo, isAdmin: u.isAdmin === 'true' });

function doGet() { return out({ ok: true, msg: 'Village API running' }); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    return out(Object.assign({ ok: true }, route(d)));
  } catch (err) {
    return out({ ok: false, error: String(err.message || err) });
  } finally { lock.releaseLock(); }
}

function route(d) {
  const phone = String(d.phone || '').trim();
  switch (d.action) {

    case 'signup': {
      if (!/^\d{10}$/.test(phone)) throw new Error('10 అంకెల ఫోన్ నెంబర్ ఇవ్వండి');
      if (phone === ADMIN_PHONE || rows('Users').some(u => u.phone === phone)) throw new Error('ఈ నెంబర్ తో ఖాతా ఇప్పటికే ఉంది');
      if (!d.name || !d.pass || d.pass.length < 6) throw new Error('పేరు, కనీసం 6 అక్షరాల పాస్‌వర్డ్ ఇవ్వండి');
      if (!d.secQ || !d.secA) throw new Error('సెక్యూరిటీ ప్రశ్న, సమాధానం ఇవ్వండి');
      add('Users', { phone, name: d.name, photo: upload(d.photo, 'user_' + phone + '.jpg'), passHash: pw(phone, d.pass),
        secQ: d.secQ, secAHash: pw(phone, String(d.secA).trim().toLowerCase()), token: '', isAdmin: 'false', created: now() });
      return {};
    }

    case 'login': {
      let u = rows('Users').find(x => x.phone === phone);
      if (phone === ADMIN_PHONE) {
        if (d.pass !== ADMIN_PASS) throw new Error('ఫోన్ నెంబర్ లేదా పాస్‌వర్డ్ తప్పు');
        if (!u) { add('Users', { phone, name: 'మెయిన్ అడ్మిన్', isAdmin: 'true', created: now() }); u = rows('Users').find(x => x.phone === phone); }
      } else if (!u || u.passHash !== pw(phone, d.pass || '')) throw new Error('ఫోన్ నెంబర్ లేదా పాస్‌వర్డ్ తప్పు');
      const token = Utilities.getUuid();
      setCell('Users', u._row, 'token', token);
      return { token, user: pub(u) };
    }

    case 'question': {
      const u = rows('Users').find(x => x.phone === phone);
      if (!u || phone === ADMIN_PHONE) throw new Error('ఈ నెంబర్ తో ఖాతా లేదు');
      return { secQ: u.secQ };
    }

    case 'reset': {
      const u = rows('Users').find(x => x.phone === phone);
      if (!u || phone === ADMIN_PHONE || u.secAHash !== pw(phone, String(d.secA || '').trim().toLowerCase())) throw new Error('సమాధానం సరిపోలలేదు');
      if (!d.newPass || d.newPass.length < 6) throw new Error('కనీసం 6 అక్షరాల పాస్‌వర్డ్ ఇవ్వండి');
      setCell('Users', u._row, 'passHash', pw(phone, d.newPass));
      setCell('Users', u._row, 'token', '');
      return {};
    }

    case 'list': {
      const me = auth(d), admin = me.isAdmin === 'true';
      const cm = rows('Comments');
      const posts = rows('Posts').reverse().map(p => {
        const show = p.showId === 'true' || admin || p.phone === me.phone;
        const likes = p.likedBy ? p.likedBy.split(',') : [];
        return { id: p.id, type: p.type, category: p.category, title: p.title, details: p.details, address: p.address,
          photo: p.photo, name: show ? p.name : '', phone: show ? p.phone : '', anon: !show, date: p.date,
          likes: likes.length, liked: likes.includes(me.phone), canDelete: admin || p.phone === me.phone,
          comments: cm.filter(c => c.postId === p.id).map(c => ({ name: c.name, text: c.text, date: c.date })) };
      });
      return { posts };
    }

    case 'addPost': {
      const me = auth(d);
      if (!d.title || !d.details) throw new Error('శీర్షిక, వివరాలు ఇవ్వండి');
      const id = String(Date.now());
      add('Posts', { id, type: d.type, category: d.category, title: d.title, details: d.details, address: d.address,
        photo: upload(d.photo, 'post_' + id + '.jpg'), name: d.name || me.name, phone: d.pphone || me.phone,
        showId: d.showId ? 'true' : 'false', likedBy: '', date: now() });
      return {};
    }

    case 'like': {
      const me = auth(d), p = rows('Posts').find(x => x.id === String(d.id));
      if (!p) throw new Error('పోస్ట్ లేదు');
      let l = p.likedBy ? p.likedBy.split(',') : [];
      l = l.includes(me.phone) ? l.filter(x => x !== me.phone) : l.concat(me.phone);
      setCell('Posts', p._row, 'likedBy', l.join(','));
      return {};
    }

    case 'comment': {
      const me = auth(d);
      const text = String(d.text || '').trim().slice(0, 500);
      if (!text) throw new Error('కామెంట్ ఖాళీగా ఉంది');
      add('Comments', { id: Date.now(), postId: d.id, name: me.name, phone: me.phone, text, date: now() });
      return {};
    }

    case 'delete': {
      const me = auth(d), p = rows('Posts').find(x => x.id === String(d.id));
      if (!p || (me.isAdmin !== 'true' && p.phone !== me.phone)) throw new Error('అనుమతి లేదు');
      sh('Posts').deleteRow(p._row);
      return {};
    }
  }
  throw new Error('తెలియని చర్య');
}

/** ఒకసారి ఇది Run చేసి Sheets/Drive అనుమతులు ఇవ్వండి */
function setup() { Object.keys(H).forEach(sh); DriveApp.getFolderById(FOLDER_ID); }
