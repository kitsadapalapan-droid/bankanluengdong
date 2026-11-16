/**
 * ระบบหลังบ้าน เว็บไซต์โรงเรียนบ้านก้านเหลืองดง
 * - คลังแบบฟอร์ม/เอกสารดาวน์โหลด (kind = 'form')
 * - ระบบส่งงานครู: แผนการสอน ผลงาน (kind = 'work')
 * ไฟล์เก็บใน Google Drive ของบัญชีที่ติดตั้ง · รายการเก็บใน Google Sheets
 *
 * ติดตั้ง: เรียก setup() หนึ่งครั้ง แล้ว Deploy > Web app
 *   Execute as: Me · Who has access: Anyone
 */

const ROOT_NAME = 'เว็บโรงเรียนบ้านก้านเหลืองดง (ไฟล์)';
const SHEET_NAME = 'ฐานข้อมูลแบบฟอร์มและงานครู';
const HEAD = ['id', 'kind', 'title', 'category', 'teacher', 'subject', 'grade', 'term', 'note', 'fileName', 'mime', 'size', 'fileId', 'link', 'createdAt'];
const MAX_BYTES = 20 * 1024 * 1024;

function setup() {
  const p = PropertiesService.getScriptProperties();
  let root = p.getProperty('ROOT_ID') ? DriveApp.getFolderById(p.getProperty('ROOT_ID')) : DriveApp.createFolder(ROOT_NAME);
  p.setProperty('ROOT_ID', root.getId());
  ['FORMS', 'WORKS'].forEach(function (k) {
    if (!p.getProperty(k + '_ID')) p.setProperty(k + '_ID', root.createFolder(k === 'FORMS' ? 'แบบฟอร์ม' : 'งานครู').getId());
  });
  if (!p.getProperty('SHEET_ID')) {
    const ss = SpreadsheetApp.create(SHEET_NAME);
    DriveApp.getFileById(ss.getId()).moveTo(root);
    const sh = ss.getSheets()[0];
    sh.setName('items');
    sh.appendRow(HEAD);
    sh.setFrozenRows(1);
    p.setProperty('SHEET_ID', ss.getId());
  }
  if (!p.getProperty('PIN')) p.setProperty('PIN', String(Math.floor(100000 + Math.random() * 900000)));
  if (!p.getProperty('ADMIN_PIN')) p.setProperty('ADMIN_PIN', String(Math.floor(100000 + Math.random() * 900000)));
  Logger.log('PIN ครู: ' + p.getProperty('PIN') + ' | PIN ผู้ดูแล: ' + p.getProperty('ADMIN_PIN'));
  Logger.log('โฟลเดอร์: ' + root.getUrl());
  Logger.log('ชีต: ' + SpreadsheetApp.openById(p.getProperty('SHEET_ID')).getUrl());
}

function sheet_() {
  return SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID')).getSheetByName('items');
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function rows_() {
  const v = sheet_().getDataRange().getValues();
  const h = v.shift();
  return v.filter(function (r) { return r[0]; }).map(function (r) {
    const o = {};
    h.forEach(function (k, i) { o[k] = r[i] instanceof Date ? r[i].toISOString() : r[i]; });
    if (o.fileId) {
      o.url = 'https://drive.google.com/file/d/' + o.fileId + '/view';
      o.download = 'https://drive.google.com/uc?export=download&id=' + o.fileId;
    } else o.url = o.link;
    delete o.fileId;
    return o;
  });
}

function doGet(e) {
  try {
    return out_({ ok: true, items: rows_().reverse() });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

function clean_(s, n) {
  return String(s == null ? '' : s).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, n || 200);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const b = JSON.parse(e.postData.contents);
    const p = PropertiesService.getScriptProperties();
    const isAdmin = b.pin && b.pin === p.getProperty('ADMIN_PIN');
    const isTeacher = isAdmin || (b.pin && b.pin === p.getProperty('PIN'));
    if (b.action === 'check') return out_({ ok: !!isTeacher, admin: !!isAdmin });
    if (!isTeacher) return out_({ ok: false, error: 'PIN ไม่ถูกต้อง' });

    if (b.action === 'upload') {
      const kind = b.kind === 'form' ? 'form' : 'work';
      if (kind === 'form' && !isAdmin) return out_({ ok: false, error: 'เพิ่มแบบฟอร์มได้เฉพาะผู้ดูแล' });
      const title = clean_(b.title, 200);
      if (!title) return out_({ ok: false, error: 'กรุณากรอกชื่อเรื่อง' });
      let fileId = '', fileName = '', mime = '', size = 0, link = '';
      if (b.data) {
        const bytes = Utilities.base64Decode(b.data);
        if (bytes.length > MAX_BYTES) return out_({ ok: false, error: 'ไฟล์ใหญ่เกิน 20 MB' });
        fileName = clean_(b.fileName, 150) || 'file';
        mime = clean_(b.mime, 100) || 'application/octet-stream';
        size = bytes.length;
        let folder = DriveApp.getFolderById(p.getProperty(kind === 'form' ? 'FORMS_ID' : 'WORKS_ID'));
        if (kind === 'work') {
          const t = clean_(b.teacher, 80) || 'ไม่ระบุชื่อ';
          const it = folder.getFoldersByName(t);
          folder = it.hasNext() ? it.next() : folder.createFolder(t);
        }
        const f = folder.createFile(Utilities.newBlob(bytes, mime, fileName));
        f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        fileId = f.getId();
      } else if (b.link && /^https:\/\//.test(b.link)) {
        link = clean_(b.link, 500);
      } else return out_({ ok: false, error: 'แนบไฟล์หรือใส่ลิงก์' });
      const id = Utilities.getUuid().slice(0, 8);
      lock.waitLock(20000);
      sheet_().appendRow([id, kind, title, clean_(b.category, 80), clean_(b.teacher, 80), clean_(b.subject, 80), clean_(b.grade, 30), clean_(b.term, 20), clean_(b.note, 500), fileName, mime, size, fileId, link, new Date()]);
      return out_({ ok: true, id: id });
    }

    if (b.action === 'delete') {
      lock.waitLock(20000);
      const sh = sheet_();
      const v = sh.getDataRange().getValues();
      for (let i = 1; i < v.length; i++) {
        if (v[i][0] === b.id) {
          if (!isAdmin && v[i][4] !== b.teacher) return out_({ ok: false, error: 'ลบได้เฉพาะงานของตนเอง' });
          if (v[i][12]) { try { DriveApp.getFileById(v[i][12]).setTrashed(true); } catch (x) {} }
          sh.deleteRow(i + 1);
          return out_({ ok: true });
        }
      }
      return out_({ ok: false, error: 'ไม่พบรายการ' });
    }
    return out_({ ok: false, error: 'คำสั่งไม่ถูกต้อง' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}
