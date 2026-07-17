// ============================================================
// 愛酷揪團 - Google Apps Script Backend
// ============================================================

const SHEET_GROUPS = 'groups';
const SHEET_PARTICIPANTS = 'participants';

// GET → 讀取全部資料
function doGet(e) {
  const result = getAllData();
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// POST → 所有寫入操作
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    let result;
    switch (data.action) {
      case 'createGroup':
        result = createGroup(data.category, data.creator, data.title, data.content,
          data.fee, data.companySubsidy, data.minParticipants, data.imageUrl); break;
      case 'updateGroup':
        result = updateGroup(data.id, data.creator, data.title, data.content,
          data.fee, data.minParticipants, data.isOpen, data.imageUrl); break;
      case 'joinGroup':
        result = joinGroup(data.groupId, data.name); break;
      case 'leaveGroup':
        result = leaveGroup(data.groupId, data.name); break;
      case 'deleteGroup':
        result = deleteGroup(data.id, data.creator); break;
      case 'uploadImage':
        result = uploadImageToDrive(data.base64, data.mimeType, data.filename); break;
      case 'disbandGroup':
        result = disbandGroup(data.id, data.creator); break;
      default:
        result = { success: false, error: 'Unknown action: ' + data.action };
    }
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (e) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: e.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Spreadsheet init ──────────────────────────────────────────
// 容器繫結模式：直接使用綁定的試算表，並在第一次執行時自動建立工作表

function getOrCreateSS() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 建立 groups 工作表（若不存在）
  if (!ss.getSheetByName(SHEET_GROUPS)) {
    const gs = ss.getSheets()[0].setName(SHEET_GROUPS);
    gs.appendRow(['id', 'category', 'creator', 'title', 'content', 'fee', 'company_subsidy', 'min_participants', 'is_open', 'created_at']);
    gs.setFrozenRows(1);
  }

  // 建立 participants 工作表（若不存在）
  if (!ss.getSheetByName(SHEET_PARTICIPANTS)) {
    const ps = ss.insertSheet(SHEET_PARTICIPANTS);
    ps.appendRow(['id', 'group_id', 'name', 'joined_at']);
    ps.setFrozenRows(1);
  }

  return ss;
}

function getSheet(name) {
  return getOrCreateSS().getSheetByName(name);
}

// ── Row parsers ───────────────────────────────────────────────

function rowToGroup(row) {
  return {
    id: String(row[0]),
    category: String(row[1]),
    creator: String(row[2]),
    title: String(row[3]),
    content: String(row[4]),
    fee: Number(row[5]) || 0,
    company_subsidy: Number(row[6]) || 0,
    min_participants: Number(row[7]) || 1,
    is_open: row[8] === true || String(row[8]).toUpperCase() === 'TRUE',
    created_at: String(row[9]),
    image_url: row[10] ? String(row[10]) : '',
    disbanded: row[11] === true || String(row[11] || '').toUpperCase() === 'TRUE'
  };
}

function rowToParticipant(row) {
  return {
    id: String(row[0]),
    group_id: String(row[1]),
    name: String(row[2]),
    joined_at: String(row[3])
  };
}

// ── Read all data ─────────────────────────────────────────────

function getAllData() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('all_data');
  if (hit) return JSON.parse(hit);

  const gData = getSheet(SHEET_GROUPS).getDataRange().getValues();
  const pData = getSheet(SHEET_PARTICIPANTS).getDataRange().getValues();

  const groups = [];
  for (let i = 1; i < gData.length; i++) {
    if (gData[i][0]) groups.push(rowToGroup(gData[i]));
  }

  const participants = [];
  for (let i = 1; i < pData.length; i++) {
    if (pData[i][0]) participants.push(rowToParticipant(pData[i]));
  }

  const result = { groups, participants };
  cache.put('all_data', JSON.stringify(result), 30);
  return result;
}

function invalidateCache() {
  CacheService.getScriptCache().remove('all_data');
}

// ── Create group ──────────────────────────────────────────────

function createGroup(category, creator, title, content, fee, companySubsidy, minParticipants, imageUrl) {
  invalidateCache();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const gs = getSheet(SHEET_GROUPS);
    const ps = getSheet(SHEET_PARTICIPANTS);
    const now = new Date().toISOString();
    const groupId = Utilities.getUuid();

    gs.appendRow([groupId, category, creator, title, content, Number(fee),
      Number(companySubsidy), Number(minParticipants), true, now, imageUrl || '', false]);

    // Auto-join creator unless single-join constraint blocks them
    let autoJoined = true;
    if (category === 'Q2團建' || category === 'Q3活動') {
      const gData = gs.getDataRange().getValues();
      const pData = ps.getDataRange().getValues();
      outer:
      for (let i = 1; i < pData.length; i++) {
        if (pData[i][2] === creator) {
          for (let j = 1; j < gData.length; j++) {
            if (gData[j][0] === pData[i][1] && gData[j][1] === category) {
              autoJoined = false;
              break outer;
            }
          }
        }
      }
    }

    if (autoJoined) {
      ps.appendRow([Utilities.getUuid(), groupId, creator, now]);
    }

    return { success: true, id: groupId, autoJoined };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ── Update group ──────────────────────────────────────────────

function updateGroup(id, creator, title, content, fee, minParticipants, isOpen, imageUrl) {
  invalidateCache();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getSheet(SHEET_GROUPS);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id && data[i][2] === creator) {
        sheet.getRange(i + 1, 4).setValue(title);
        sheet.getRange(i + 1, 5).setValue(content);
        sheet.getRange(i + 1, 6).setValue(Number(fee));
        sheet.getRange(i + 1, 8).setValue(Number(minParticipants));
        sheet.getRange(i + 1, 9).setValue(Boolean(isOpen));
        sheet.getRange(i + 1, 11).setValue(imageUrl || '');
        return { success: true };
      }
    }
    return { success: false, error: '找不到此團或無權限' };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ── Join group ────────────────────────────────────────────────

function joinGroup(groupId, name) {
  invalidateCache();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const gs = getSheet(SHEET_GROUPS);
    const ps = getSheet(SHEET_PARTICIPANTS);
    const gData = gs.getDataRange().getValues();
    const pData = ps.getDataRange().getValues();

    let groupRow = null;
    for (let i = 1; i < gData.length; i++) {
      if (gData[i][0] === groupId) { groupRow = gData[i]; break; }
    }
    if (!groupRow) return { success: false, error: '找不到此團' };

    const isOpen = groupRow[8] === true || String(groupRow[8]).toUpperCase() === 'TRUE';
    if (!isOpen) return { success: false, error: '此團已停止報名' };

    const isDisbanded = groupRow[11] === true || String(groupRow[11] || '').toUpperCase() === 'TRUE';
    if (isDisbanded) return { success: false, error: '此團已流團' };

    const minParticipants = Number(groupRow[7]) || 1;
    const category = groupRow[1];

    let currentCount = 0;
    for (let i = 1; i < pData.length; i++) {
      if (pData[i][1] === groupId) currentCount++;
    }
    if (currentCount >= minParticipants) return { success: false, error: '此團已達成團人數，報名已截止' };

    for (let i = 1; i < pData.length; i++) {
      if (pData[i][1] === groupId && pData[i][2] === name) {
        return { success: false, error: '你已加入此團' };
      }
    }

    if (category === 'Q2團建' || category === 'Q3活動') {
      for (let i = 1; i < pData.length; i++) {
        if (pData[i][2] === name) {
          for (let j = 1; j < gData.length; j++) {
            if (gData[j][0] === pData[i][1] && gData[j][1] === category) {
              return { success: false, error: category + ' 每人只能參加一個團！' };
            }
          }
        }
      }
    }

    ps.appendRow([Utilities.getUuid(), groupId, name, new Date().toISOString()]);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ── Leave group ───────────────────────────────────────────────

function leaveGroup(groupId, name) {
  invalidateCache();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getSheet(SHEET_PARTICIPANTS);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === groupId && data[i][2] === name) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: '找不到報名記錄' };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ── Delete group ──────────────────────────────────────────────

function deleteGroup(id, creator) {
  invalidateCache();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const gs = getSheet(SHEET_GROUPS);
    const gData = gs.getDataRange().getValues();

    for (let i = 1; i < gData.length; i++) {
      if (gData[i][0] === id && gData[i][2] === creator) {
        gs.deleteRow(i + 1);
        const ps = getSheet(SHEET_PARTICIPANTS);
        const pData = ps.getDataRange().getValues();
        for (let j = pData.length - 1; j >= 1; j--) {
          if (pData[j][1] === id) ps.deleteRow(j + 1);
        }
        return { success: true };
      }
    }
    return { success: false, error: '找不到此團或無權限' };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ── Disband group ─────────────────────────────────────────────

function disbandGroup(id, creator) {
  invalidateCache();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getSheet(SHEET_GROUPS);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id && data[i][2] === creator) {
        sheet.getRange(i + 1, 12).setValue(true);  // disbanded = true
        sheet.getRange(i + 1, 9).setValue(false);   // is_open = false
        return { success: true };
      }
    }
    return { success: false, error: '找不到此團或無權限' };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ── Image upload ──────────────────────────────────────────────

function uploadImageToDrive(base64Data, mimeType, filename) {
  try {
    const bytes = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(bytes, mimeType, filename || 'image.jpg');
    const folder = getOrCreateImageFolder();
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const fileId = file.getId();
    return { success: true, url: 'https://drive.google.com/uc?export=view&id=' + fileId };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function getOrCreateImageFolder() {
  return DriveApp.getFolderById('1RQDHiiZ0K0phUkcM0Bg3YtK7Ui_9NrWv');
}
