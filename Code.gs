function doGet(e) {
  var payload = {
    ok: true,
    message: 'CE Dept backend is live and reachable.',
    deployedActions: ['login','getQuestions','saveQuestions','addSubmission','getSubmissions','changePassword','getUsers','getSettings','saveSettings','hasSubmitted','deleteAllSubmissions','startNewForm'],
    checkedAt: new Date().toISOString()
  };
  return ContentService.createTextOutput(JSON.stringify(payload, null, 2)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var body = JSON.parse(e.postData.contents);
  var action = body.action, p = body.payload || {}, result;
  try {
    if (action === 'login') result = handleLogin(ss, p);
    else if (action === 'getQuestions') result = getQuestions(ss);
    else if (action === 'saveQuestions') result = saveQuestions(ss, p);
    else if (action === 'addSubmission') result = addSubmission(ss, p);
    else if (action === 'getSubmissions') result = getSubmissions(ss);
    else if (action === 'changePassword') result = changePassword(ss, p);
    else if (action === 'getUsers') result = getUsers(ss);
    else if (action === 'getSettings') result = getSettings(ss);
    else if (action === 'saveSettings') result = saveSettings(ss, p);
    else if (action === 'hasSubmitted') result = hasSubmitted(ss, p);
    else if (action === 'deleteAllSubmissions') result = deleteAllSubmissions(ss);
    else if (action === 'startNewForm') result = startNewForm(ss);
    else result = { ok:false, error:'Unknown action' };
  } catch (err) { result = { ok:false, error:String(err) }; }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function getSheet(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(headers); }
  return sh;
}

function getSubmissionsSheet(ss) {
  var sh = getSheet(ss, 'Submissions', ['ID','Time','Role','Mobile','Name','AnswersJSON','FormVersion']);
  if (sh.getRange(1,7).getValue() !== 'FormVersion') {
    sh.getRange(1,7).setValue('FormVersion');
  }
  return sh;
}

function handleLogin(ss, p) {
  var sh = getSheet(ss, 'Users', ['Mobile','PasswordHash','Role','Name']);
  var data = sh.getDataRange().getValues();
  for (var i=1;i<data.length;i++){
    if (String(data[i][0])===p.mobile && data[i][2]===p.role) {
      if (data[i][1]!==p.passwordHash) return { ok:false, error:'Incorrect password.' };
      return { ok:true, user:{ mobile:data[i][0], role:data[i][2], name:data[i][3] } };
    }
  }
  if (p.role==='admin') {
    if (data.length===1) {
      sh.appendRow([p.mobile, p.passwordHash, 'admin', p.name||'Admin']);
      return { ok:true, user:{ mobile:p.mobile, role:'admin', name:p.name||'Admin' }, seeded:true };
    }
    return { ok:false, error:'Incorrect mobile number or password.' };
  }
  if (!p.name) return { ok:false, error:'NEED_NAME' };
  sh.appendRow([p.mobile, p.passwordHash, p.role, p.name]);
  return { ok:true, user:{ mobile:p.mobile, role:p.role, name:p.name } };
}

function getQuestions(ss) {
  var sh = getSheet(ss, 'Questions', ['Role','QuestionsJSON']);
  var data = sh.getDataRange().getValues();
  var out = { student: [], teacher: [] };
  for (var i=1;i<data.length;i++){
    if (data[i][0]==='student') out.student = JSON.parse(data[i][1]||'[]');
    if (data[i][0]==='teacher') out.teacher = JSON.parse(data[i][1]||'[]');
  }
  return { ok:true, questions: out };
}

function saveQuestions(ss, p) {
  var sh = getSheet(ss, 'Questions', ['Role','QuestionsJSON']);
  var data = sh.getDataRange().getValues();
  var rowIdx = -1;
  for (var i=1;i<data.length;i++){ if (data[i][0]===p.role) rowIdx = i+1; }
  var json = JSON.stringify(p.questions);
  if (rowIdx>0) sh.getRange(rowIdx,2).setValue(json); else sh.appendRow([p.role, json]);
  return { ok:true };
}

function getSettingValue(ss, key) {
  var sh = getSheet(ss, 'Settings', ['Key','Value']);
  var data = sh.getDataRange().getValues();
  for (var i=1;i<data.length;i++){ if (data[i][0]===key) return data[i][1]; }
  return '';
}

function getFormVersion(ss) {
  var v = getSettingValue(ss, 'formVersion');
  return v ? v : 'v1';
}

function startNewForm(ss) {
  var newVersion = 'v' + Date.now();
  var sh = getSheet(ss, 'Settings', ['Key','Value']);
  var data = sh.getDataRange().getValues();
  var rowIdx = -1;
  for (var i=1;i<data.length;i++){ if (data[i][0]==='formVersion') rowIdx = i+1; }
  if (rowIdx>0) sh.getRange(rowIdx,2).setValue(newVersion); else sh.appendRow(['formVersion', newVersion]);

  // Clear the live question sets so the admin defines fresh ones for the new form.
  // Past submissions are untouched — each already stores its own question labels
  // alongside its answers, so old responses stay fully readable on their own.
  var qsh = getSheet(ss, 'Questions', ['Role','QuestionsJSON']);
  var qdata = qsh.getDataRange().getValues();
  for (var j=1;j<qdata.length;j++){
    if (qdata[j][0]==='student' || qdata[j][0]==='teacher') {
      qsh.getRange(j+1,2).setValue('[]');
    }
  }

  return { ok:true, version:newVersion };
}

function addSubmission(ss, p) {
  var closeAt = getSettingValue(ss, 'formCloseAt');
  if (closeAt) {
    var deadline = new Date(closeAt);
    if (!isNaN(deadline.getTime()) && new Date() > deadline) {
      return { ok:false, error:'CLOSED' };
    }
  }
  var answersJson = JSON.stringify(p.answers);
  if (answersJson.length > 45000) {
    return { ok:false, error:'TOO_LARGE' };
  }
  var version = getFormVersion(ss);
  var sh = getSubmissionsSheet(ss);
  var data = sh.getDataRange().getValues();
  for (var i=1;i<data.length;i++){
    if (String(data[i][3])===p.mobile && data[i][6]===version) {
      return { ok:false, error:'ALREADY_SUBMITTED' };
    }
  }
  sh.appendRow([p.id, p.time, p.role, p.mobile, p.name, answersJson, version]);
  return { ok:true };
}

function hasSubmitted(ss, p) {
  var version = getFormVersion(ss);
  var sh = getSubmissionsSheet(ss);
  var data = sh.getDataRange().getValues();
  for (var i=1;i<data.length;i++){
    if (String(data[i][3])===p.mobile && data[i][6]===version) {
      return { ok:true, submitted:true };
    }
  }
  return { ok:true, submitted:false };
}

function getSubmissions(ss) {
  var sh = getSubmissionsSheet(ss);
  var data = sh.getDataRange().getValues();
  var out = [];
  for (var i=1;i<data.length;i++){
    out.push({ id:data[i][0], time:data[i][1], role:data[i][2], mobile:data[i][3], name:data[i][4], answers: JSON.parse(data[i][5]||'{}'), formVersion:data[i][6] });
  }
  return { ok:true, submissions: out };
}

function deleteAllSubmissions(ss) {
  var sh = getSubmissionsSheet(ss);
  var lastRow = sh.getLastRow();
  if (lastRow > 1) {
    sh.deleteRows(2, lastRow - 1);
  }
  return { ok:true };
}

function changePassword(ss, p) {
  var sh = getSheet(ss, 'Users', ['Mobile','PasswordHash','Role','Name']);
  var data = sh.getDataRange().getValues();
  for (var i=1;i<data.length;i++){
    if (String(data[i][0])===p.mobile && data[i][2]===p.role) {
      sh.getRange(i+1,2).setValue(p.newPasswordHash);
      return { ok:true };
    }
  }
  return { ok:false, error:'User not found.' };
}

function getUsers(ss) {
  var sh = getSheet(ss, 'Users', ['Mobile','PasswordHash','Role','Name']);
  var data = sh.getDataRange().getValues();
  var out = [];
  for (var i=1;i<data.length;i++){
    out.push({ mobile:data[i][0], role:data[i][2], name:data[i][3] });
  }
  return { ok:true, users: out };
}

function getSettings(ss) {
  var sh = getSheet(ss, 'Settings', ['Key','Value']);
  var data = sh.getDataRange().getValues();
  var out = {};
  for (var i=1;i<data.length;i++){
    out[data[i][0]] = data[i][1];
  }
  return { ok:true, settings: out };
}

function saveSettings(ss, p) {
  var sh = getSheet(ss, 'Settings', ['Key','Value']);
  var data = sh.getDataRange().getValues();
  var rowIdx = -1;
  for (var i=1;i<data.length;i++){ if (data[i][0]===p.key) rowIdx = i+1; }
  if (rowIdx>0) sh.getRange(rowIdx,2).setValue(p.value); else sh.appendRow([p.key, p.value]);
  return { ok:true };
}
