
function extractRequestAndToken_(args, requestPayload) {
  args = args || [];
  let token = '';
  let req = {};
  
  for (let i = 0; i < args.length; i++) {
    const item = args[i];
    if (typeof item === 'string' && item.length > 10) {
      if (!token || item.length >= 20) token = item;
    } else if (typeof item === 'object' && item !== null && !Array.isArray(item)) {
      req = item;
    }
  }
  
  if (requestPayload && typeof requestPayload === 'object') {
    if (!token && requestPayload.token) token = String(requestPayload.token);
    if (Object.keys(req).length === 0) req = requestPayload;
  }
  
  if (!token) {
    if (typeof args[1] === 'string') token = args[1];
    else if (typeof args[0] === 'string') token = args[0];
  }
  
  if (!token && req && req.token) {
    token = String(req.token);
  }
  
  return { request: req, token: String(token || '').trim() };
}

const APP_TITLE = 'Security Management System';

// ฟังก์ชันสำหรับกดเรียกใช้ (Run) เพื่อกดยอมรับสิทธิ์ส่งอีเมลจริง
function authorizeEmailPermission() {
  const targetEmail = 'sermpong.ch@gmail.com';
  MailApp.sendEmail({
    to: targetEmail,
    subject: "🔔 ยืนยันสิทธิ์ส่งอีเมลสำเร็จ (Security Management System)",
    body: "สิทธิ์การส่งอีเมลของ Google Apps Script ถูกเปิดใช้งานเรียบร้อยแล้ว ระบบสามารถส่งอีเมลและรหัส OTP ได้ตามปกติ"
  });
  Logger.log("✅ เปิดใช้สิทธิ์สำเร็จ! ระบบได้ลองส่งอีเมลยืนยันไปที่: " + targetEmail + " เรียบร้อยแล้วครับ");
}

const PROPERTY_SPREADSHEET_ID = 'SHIFTFLOW_SPREADSHEET_ID';
const REQUIRED_SHEETS = ['Employees', 'Employee Licenses', 'Shift Types', 'Schedule', 'Schedule Approvals', 'Schedule Approval Log', 'Rules', 'Dashboard', 'Users', 'Settings', 'License Audit Log', 'User Audit Log'];
const DEFAULT_MAX_WEEKLY_HOURS = 72;
const LOGIN_HASH_PEPPER = 'shiftflow-login-v1';
const PRIMARY_ADMIN_EMAIL = 'sermpong.ch@gmail.com';
const PRIMARY_ADMIN_PASSWORD_HASH = '7c7c0113bab663424c8ae8de03272fa5a668956272de3585a95588a3b7e6dadc';
const PRIMARY_ADMIN_BOOTSTRAP_VERSION = 'sermpong-admin-v1';
const ALLOWED_USER_ROLES = ['Admin', 'Manager', 'Viewer'];

// Rules sheet only carries id/name/value/unit/enabled, so the shift code and
// department each coverage rule applies to are pinned here.
const COVERAGE_RULES = {
  RULE003: { code: 'D', department: 'All' },
  RULE004: { code: 'N', department: 'All' }
};

const THAI_MONTHS = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ status: 'running', message: 'Security API is active' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    try { ensureLeaveSheets_(getOrCreateSpreadsheet_()); } catch(err) {}
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput(JSON.stringify({ error: 'No POST data provided' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    const request = JSON.parse(e.postData.contents);
    const action = request.action;
    const args = request.args || [];
    
    let data = null;
    switch (action) {
      case 'authenticateUser': data = authenticateUser(args[0], args[1]); break;
      case 'getAppData': data = getAppData(args[0]); break;
      case 'validateSession': data = validateSession(args[0]); break;
      case 'logoutUser': data = logoutUser(args[0]); break;
      case 'getAvailableEmployeesForRegistration': data = getAvailableEmployeesForRegistration(); break;
      case 'requestRegistrationOtp': data = requestRegistrationOtp(args[0], args[1]); break;
      case 'registerAccount': data = registerAccount(args[0]); break;
      case 'requestPasswordResetOtp': data = requestPasswordResetOtp(args[0]); break;
      case 'resetPasswordWithOtp': data = resetPasswordWithOtp(args[0], args[1], args[2]); break;
      case 'addEmployee': { const rt = extractRequestAndToken_(args, request.payload); data = addEmployee(rt.request, rt.token); break; }
      case 'updateEmployee': { const rt = extractRequestAndToken_(args, request.payload); data = updateEmployee(rt.request, rt.token); break; }
      case 'deleteEmployee': { const rt = extractRequestAndToken_(args, request.payload); data = deleteEmployee(rt.request, rt.token); break; }
      case 'saveEmployeeLicense': { const rt = extractRequestAndToken_(args, request.payload); data = saveEmployeeLicense(rt.request, rt.token); break; }
      case 'deleteEmployeeLicense': { const rt = extractRequestAndToken_(args, request.payload); data = deleteEmployeeLicense(rt.request, rt.token); break; }
      case 'addShiftType': { const rt = extractRequestAndToken_(args, request.payload); data = addShiftType(rt.request, rt.token); break; }
      case 'deleteShiftType': { const rt = extractRequestAndToken_(args, request.payload); data = deleteShiftType(rt.request, rt.token); break; }
      case 'getUserAdministration': data = getUserAdministration(args[0]); break;
      case 'reviewUserAccount': { const rt = extractRequestAndToken_(args, request.payload); data = reviewUserAccount(rt.request, rt.token); break; }
      case 'getEmployeesPage': { const rt = extractRequestAndToken_(args, request.payload); data = getEmployeesPage(rt.request, rt.token); break; }
      case 'getOptimizedScheduleData': { const rt = extractRequestAndToken_(args, request.payload); data = getOptimizedScheduleData(rt.request, rt.token); break; }
      case 'previewAutoSchedule': { const rt = extractRequestAndToken_(args, request.payload); data = previewAutoSchedule(rt.request, rt.token); break; }
      case 'commitAutoSchedule': { const rt = extractRequestAndToken_(args, request.payload); data = commitAutoSchedule(rt.request, rt.token); break; }
      case 'updateEmployeeShifts': { const rt = extractRequestAndToken_(args, request.payload); data = updateEmployeeShifts(rt.request, rt.token); break; }
      case 'updateEmployeeShift': { const rt = extractRequestAndToken_(args, request.payload); data = updateEmployeeShift(rt.request, rt.token); break; }
      case 'approveScheduleMonth': data = approveScheduleMonth(args[0], args[1], args[2]); break;
      case 'exportApprovedSchedule': { const rt = extractRequestAndToken_(args, request.payload); data = exportApprovedSchedule(rt.request, rt.token); break; }
      case 'archiveScheduleBeforeMonth': { const rt = extractRequestAndToken_(args, request.payload); data = archiveScheduleBeforeMonth(rt.request, rt.token); break; }
      case 'runSetup': data = runSetup(args[0]); break;
            case 'showCurrentSpreadsheet': data = showCurrentSpreadsheet(args[0]); break;
      case 'getDatabaseInfo': data = getDatabaseInfo(args[0] || (request.payload ? request.payload.token : '')); break;
      case 'testLineMessage': data = testLineMessage(args[0] || (request.payload ? request.payload.token : '')); break;
      case 'submitLeaveRequest':
      case 'saveLeaveRequest': {
        const token = args[0] || (request.payload ? request.payload.token : '') || '';
        const leaveData = args[1] || request.payload || {};
        data = submitLeaveRequest(token, leaveData);
        break;
      }
      case 'getLeaveSummary': {
        const token = args[0] || (request.payload ? request.payload.token : '') || '';
        const filterName = args[1] || (request.payload ? request.payload.name : '') || '';
        data = getLeaveSummary(token, filterName);
        break;
      }
      case 'getPendingLeaves':
      case 'getPendingRequests': {
        const token = args[0] || (request.payload ? request.payload.token : '') || '';
        data = getPendingLeaves(token);
        break;
      }
      case 'updateLeaveStatus':
      case 'changeRequestStatus': {
        const token = args[0] || (request.payload ? request.payload.token : '') || '';
        const leaveId = args[1] || (request.payload ? request.payload.requestId : '') || (request.payload ? request.payload.leaveId : '');
        const status = args[2] || (request.payload ? request.payload.status : '');
        data = updateLeaveStatus(token, leaveId, status);
        break;
      }
      case 'getAllLeaveHistory': {
        const token = args[0] || (request.payload ? request.payload.token : '') || '';
        data = getAllLeaveHistory_(token);
        break;
      }
      case 'getLineSettings': {
        const token = args[0] || (request.payload ? request.payload.token : '') || '';
        data = getLineSettings(token);
        break;
      }
      case 'saveLineSettings': {
        const token = args[0] || (request.payload ? request.payload.token : '') || '';
        const settings = args[1] || request.payload || {};
        data = saveLineSettings(token, settings);
        break;
      }
      default:
        throw new Error('Unknown action: ' + action);
    }
    
    return ContentService.createTextOutput(JSON.stringify({ status: 'success', data: data }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    const errorMsg = error instanceof Error ? error.message : String(error);
    return ContentService.createTextOutput(JSON.stringify({ error: errorMsg }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Creates any sheet that is missing and seeds it with demo rows.
 * Sheets that already exist are left untouched.
 */
function initializeSystem_() {
  const spreadsheet = getOrCreateSpreadsheet_();
  const created = ensureRequiredSheets_(spreadsheet);
  ensureApplicationBranding_(spreadsheet);
  ensureLicenseSchema_(spreadsheet);
  ensureUserSchema_(spreadsheet);
  ensureScheduleApprovalSchema_(spreadsheet);
  ensurePrimaryAdmin_(spreadsheet);
  try { ensureLeaveSheets_(spreadsheet); } catch (e) {}
  ensureScheduleMetadataColumns_(spreadsheet.getSheetByName('Schedule'));
  migrateLegacyDemoScheduleYear_(spreadsheet.getSheetByName('Schedule'));
  enforceMaxWeeklyHours72_(spreadsheet.getSheetByName('Rules'));

  SpreadsheetApp.flush();

  return {
    spreadsheetId: spreadsheet.getId(),
    url: spreadsheet.getUrl(),
    createdSheets: created
  };
}

function runSetup(token) {
  requireActiveSession_(token);
  return initializeSystem_();
}

function getSpreadsheetUrl(token) {
  requireActiveSession_(token);
  return getOrCreateSpreadsheet_().getUrl();
}

function showCurrentSpreadsheet(token) {
  requireActiveSession_(token);
  const spreadsheet = getOrCreateSpreadsheet_();
  console.log('Current system spreadsheet: ' + spreadsheet.getUrl());
  return spreadsheet.getUrl();
}

/** Validates a Security Management System user against the Users sheet. */
function isAdminOrManager_(user) {
  if (!user) return false;
  const r = String(user.Role || '').trim().toLowerCase();
  return (r === 'admin' || r === 'manager' || r.includes('admin') || r.includes('manager') || r.includes('ผู้ดูแล') || r.includes('หัวหน้า'));
}

function authenticateUser(email, password) {
  const normalizedEmail = String(email || '').trim().toLowerCase();
  const suppliedPassword = String(password || '');
  if (!normalizedEmail || !suppliedPassword) {
    throw new Error('กรุณากรอกอีเมลและรหัสผ่าน');
  }

  const spreadsheet = getOrCreateSpreadsheet_();
  ensureRequiredSheets_(spreadsheet);
  ensureUserSchema_(spreadsheet);
  ensurePrimaryAdmin_(spreadsheet);
  try { ensureLeaveSheets_(spreadsheet); } catch (e) {} // in authenticateUser
  const sheet = spreadsheet.getSheetByName('Users');
  ensureUserPasswordHashes_(sheet);
  const loginCache = CacheService.getScriptCache();
  const attemptKey = 'shiftflow-login-attempts-' + hashPassword_(normalizedEmail, 'login-attempt').slice(0, 32);
  const failedAttempts = Number(loginCache.get(attemptKey) || 0);
  if (failedAttempts >= 5) throw new Error('เข้าสู่ระบบผิดหลายครั้ง กรุณารอ 10 นาทีแล้วลองใหม่');
  const user = readObjects_(sheet).find(function(row) {
    return String(row.Email || '').trim().toLowerCase() === normalizedEmail;
  });
  const status = user ? String(user.Status || '').trim().toLowerCase() : '';
  const expectedHash = user ? String(user['Password Hash'] || '') : '';

  if (user && status === 'pending') throw new Error('บัญชีอยู่ระหว่างรอ Admin หรือ Manager อนุมัติ');
  if (user && status === 'rejected') throw new Error('คำขอลงทะเบียนถูกปฏิเสธ กรุณาติดต่อ Admin');
  if (user && status !== 'active') throw new Error('บัญชีนี้ยังไม่เปิดใช้งาน กรุณาติดต่อ Admin');
  if (!user || !expectedHash || hashPassword_(normalizedEmail, suppliedPassword) !== expectedHash) {
    loginCache.put(attemptKey, String(failedAttempts + 1), 600);
    throw new Error('อีเมลหรือรหัสผ่านไม่ถูกต้อง');
  }
  if (ALLOWED_USER_ROLES.indexOf(String(user.Role || '').trim()) === -1) {
    throw new Error('บัญชียังไม่ได้รับการกำหนด Role จาก Admin');
  }
  loginCache.remove(attemptKey);

  const idVal = String(user['User ID'] || user.id || '');
  const nameVal = String(user.Name || user.FullName || user.name || user.email || '').trim();
  const emailVal = String(user.Email || user.email || '').trim();
  const roleVal = String(user.Role || user.role || '').trim();
  const deptVal = String(user.Department || user.department || 'All').trim();

  const sessionUser = {
    id: idVal,
    'User ID': idVal,
    name: nameVal,
    Name: nameVal,
    FullName: nameVal,
    email: emailVal,
    Email: emailVal,
    role: roleVal,
    Role: roleVal,
    department: deptVal,
    Department: deptVal,
    status: 'Active',
    Status: 'Active'
  };
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  CacheService.getScriptCache().put(
    'shiftflow-session-' + token,
    JSON.stringify(sessionUser),
    21600
  );
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const emailColumn = headers.indexOf('Email');
  const lastLoginColumn = headers.indexOf('Last Login At');
  if (lastLoginColumn >= 0) {
    for (let index = 1; index < values.length; index++) {
      if (String(values[index][emailColumn] || '').trim().toLowerCase() !== normalizedEmail) continue;
      sheet.getRange(index + 1, lastLoginColumn + 1).setValue(new Date()).setNumberFormat('yyyy-mm-dd hh:mm:ss');
      break;
    }
  }
  return Object.assign({}, sessionUser, { token: token });
}

function validatePasswordStrength_(password) {
  const value = String(password || '');
  if (value.length < 8 || value.length > 128) throw new Error('รหัสผ่านต้องมีความยาว 8-128 ตัวอักษร');
  if (!/[a-z]/.test(value) || !/[A-Z]/.test(value) || !/\d/.test(value) || !/[^A-Za-z0-9]/.test(value)) {
    throw new Error('รหัสผ่านต้องมีตัวพิมพ์เล็ก ตัวพิมพ์ใหญ่ ตัวเลข และอักขระพิเศษ');
  }
}

function appendUserAudit_(spreadsheet, record) {
  const sheet = spreadsheet.getSheetByName('User Audit Log');
  if (!sheet) return;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  sheet.appendRow(headers.map(function(header) {
    return Object.prototype.hasOwnProperty.call(record, header) ? record[header] : '';
  }));
}

function getAvailableEmployeesForRegistration() {
  const spreadsheet = getOrCreateSpreadsheet_();
  ensureUserSchema_(spreadsheet);
  
  const employees = readObjects_(spreadsheet.getSheetByName('Employees'));
  const activeEmployees = employees.filter(function(emp) {
    return String(emp.Status || '').trim().toLowerCase() === 'active';
  });

  const users = readObjects_(spreadsheet.getSheetByName('Users'));
  const registeredEmployeeIds = {};
  users.forEach(function(u) {
    const eid = String(u['Employee ID'] || '').trim().toUpperCase();
    if (eid) registeredEmployeeIds[eid] = true;
  });

  return activeEmployees
    .filter(function(emp) {
      const eid = String(emp['Employee ID'] || '').trim().toUpperCase();
      return eid && !registeredEmployeeIds[eid];
    })
    .map(function(emp) {
      return {
        id: String(emp['Employee ID'] || '').trim(),
        name: String(emp['Name'] || '').trim(),
        department: String(emp['Department'] || '').trim()
      };
    })
    .sort(function(a, b) {
      return a.name.localeCompare(b.name);
    });
}

function requestRegistrationOtp(employeeId, email) {
  const normalizedEmail = String(email || '').trim().toLowerCase();
  const normalizedEmpId = String(employeeId || '').trim().toUpperCase();
  
  if (!normalizedEmpId) throw new Error('กรุณาเลือกชื่อพนักงาน');
  if (!normalizedEmail || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(normalizedEmail)) {
    throw new Error('รูปแบบอีเมลไม่ถูกต้อง');
  }

  const spreadsheet = getOrCreateSpreadsheet_();
  ensureUserSchema_(spreadsheet);

  const employees = readObjects_(spreadsheet.getSheetByName('Employees'));
  const targetEmployee = employees.find(function(emp) {
    return String(emp['Employee ID'] || '').trim().toUpperCase() === normalizedEmpId;
  });
  if (!targetEmployee || String(targetEmployee.Status || '').trim().toLowerCase() !== 'active') {
    throw new Error('ไม่พบข้อมูลพนักงานที่สามารถลงทะเบียนได้');
  }

  const users = readObjects_(spreadsheet.getSheetByName('Users'));
  const emailInUse = users.some(function(u) {
    return String(u.Email || '').trim().toLowerCase() === normalizedEmail;
  });
  if (emailInUse) throw new Error('อีเมลนี้ถูกใช้งานหรืออยู่ในระหว่างรออนุมัติแล้ว');

  const empIdInUse = users.some(function(u) {
    return String(u['Employee ID'] || '').trim().toUpperCase() === normalizedEmpId;
  });
  if (empIdInUse) throw new Error('พนักงานท่านนี้ลงทะเบียนใช้งานไปแล้ว');

  const otp = Math.floor(100000 + Math.random() * 900000).toString();
  const cache = CacheService.getScriptCache();
  const cacheKey = 'shiftflow-reg-otp-' + hashPassword_(normalizedEmail, 'otp-key').slice(0, 32);
  cache.put(cacheKey, otp, 900);

  const emailBody = 'รหัส OTP สำหรับยืนยันการสมัครสมาชิก Security Management System ของคุณคือ:\n\n' +
                    otp + '\n\nรหัสนี้จะหมดอายุใน 15 นาที\n\nระบบบริหารจัดการกะอัตโนมัติ';
  
  MailApp.sendEmail({
    to: normalizedEmail,
    subject: 'OTP ยืนยันการสมัครสมาชิก',
    body: emailBody
  });
  
  return { status: 'OTP Sent', message: 'ส่งรหัส OTP ไปที่อีเมลแล้ว กรุณาตรวจสอบ' };
}

function registerAccount(registration) {
  const input = registration || {};
  const employeeId = String(input.employeeId || '').trim().toUpperCase();
  const email = String(input.email || '').trim().toLowerCase();
  const password = String(input.password || '');
  const otp = String(input.otp || '').trim();

  if (!employeeId) throw new Error('กรุณาเลือกชื่อพนักงาน');
  if (email.length > 254) throw new Error('อีเมลยาวเกินกำหนด');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('รูปแบบอีเมลไม่ถูกต้อง');
  if (!otp) throw new Error('กรุณากรอกรหัส OTP');
  validatePasswordStrength_(password);

  const cache = CacheService.getScriptCache();
  const cacheKey = 'shiftflow-reg-otp-' + hashPassword_(email, 'otp-key').slice(0, 32);
  const validOtp = cache.get(cacheKey);

  if (!validOtp) throw new Error('รหัส OTP หมดอายุหรือไม่ถูกต้อง กรุณาขอรหัสใหม่');
  if (validOtp !== otp) throw new Error('รหัส OTP ไม่ถูกต้อง');

  const spreadsheet = getOrCreateSpreadsheet_();
  ensureRequiredSheets_(spreadsheet);
  ensureUserSchema_(spreadsheet);
  ensurePrimaryAdmin_(spreadsheet);

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('ระบบกำลังประมวลผล กรุณาลองใหม่');
  try {
    const employees = readObjects_(spreadsheet.getSheetByName('Employees'));
    const targetEmployee = employees.find(function(emp) {
      return String(emp['Employee ID'] || '').trim().toUpperCase() === employeeId;
    });
    if (!targetEmployee || String(targetEmployee.Status || '').trim().toLowerCase() !== 'active') {
      throw new Error('ไม่พบข้อมูลพนักงานที่สามารถลงทะเบียนได้');
    }

    const name = String(targetEmployee.Name || '').trim();
    const requestedDepartment = String(targetEmployee.Department || '').trim();

    const sheet = spreadsheet.getSheetByName('Users');
    const values = sheet.getDataRange().getValues();
    const headers = values[0];
    const emailColumn = headers.indexOf('Email');
    const empIdColumn = headers.indexOf('Employee ID');
    const statusColumn = headers.indexOf('Status');
    let targetRow = 0;
    let existingStatus = '';

    for (let index = 1; index < values.length; index++) {
      const rowEmail = String(values[index][emailColumn] || '').trim().toLowerCase();
      const rowEmpId = String(values[index][empIdColumn] || '').trim().toUpperCase();
      if (rowEmail === email || rowEmpId === employeeId) {
        targetRow = index + 1;
        existingStatus = String(values[index][statusColumn] || '').trim();
        break;
      }
    }
    
    if (targetRow && ['active', 'pending', 'suspended'].indexOf(existingStatus.toLowerCase()) !== -1) {
      throw new Error('อีเมลหรือพนักงานนี้มีบัญชีในระบบแล้ว');
    }

    const userId = targetRow
      ? String(values[targetRow - 1][headers.indexOf('User ID')] || '').trim()
      : 'USR-' + Utilities.getUuid().replace(/-/g, '').slice(0, 10).toUpperCase();
      
    const record = {
      'User ID': userId,
      'Employee ID': employeeId,
      Name: name,
      Email: email,
      Role: '',
      Department: requestedDepartment,
      Status: 'Pending',
      'Password Hash': hashPassword_(email, password),
      'Requested At': new Date(),
      'Approved By': '',
      'Approved At': '',
      'Rejection Reason': '',
      'Updated At': new Date(),
      'Last Login At': ''
    };
    
    const output = headers.map(function(header) {
      return Object.prototype.hasOwnProperty.call(record, header) ? record[header] : '';
    });
    
    if (targetRow) sheet.getRange(targetRow, 1, 1, headers.length).setValues([output]);
    else sheet.appendRow(output);
    
    appendUserAudit_(spreadsheet, {
      Timestamp: new Date(), Action: 'REGISTERED', 'User ID': userId, 'Employee ID': employeeId, Email: email,
      Role: '', Department: requestedDepartment, Reason: 'รอ Admin อนุมัติ', 'Performed By': email
    });
    SpreadsheetApp.flush();
    cache.remove(cacheKey); // Clear OTP after success
    return { status: 'Pending', message: 'ลงทะเบียนสำเร็จ กรุณารอ Admin อนุมัติการเข้าใช้งาน' };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Sends a 6-digit OTP to the registered email for password reset.
 */
function requestPasswordResetOtp(email) {
  const normalizedEmail = String(email || '').trim().toLowerCase();
  if (!normalizedEmail || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(normalizedEmail)) {
    throw new Error('กรุณากรอกอีเมลที่ถูกต้อง');
  }

  const spreadsheet = getOrCreateSpreadsheet_();
  ensureRequiredSheets_(spreadsheet);
  ensureUserSchema_(spreadsheet);
  ensurePrimaryAdmin_(spreadsheet);
  const sheet = spreadsheet.getSheetByName('Users');
  const user = readObjects_(sheet).find(function(row) {
    return String(row.Email || '').trim().toLowerCase() === normalizedEmail;
  });
  if (!user) throw new Error('ไม่พบบัญชีผู้ใช้งานที่ตรงกับอีเมลนี้');
  const status = String(user.Status || '').trim().toLowerCase();
  if (status !== 'active') throw new Error('บัญชีนี้ยังไม่เปิดใช้งานหรือถูกระงับ ไม่สามารถรีเซ็ตรหัสผ่านได้');

  // Generate 6-digit OTP
  const otp = Math.floor(100000 + Math.random() * 900000).toString();
  const cache = CacheService.getScriptCache();
  const cacheKey = 'shiftflow-pwd-otp-' + hashPassword_(normalizedEmail, 'otp-key').slice(0, 32);
  cache.put(cacheKey, otp, 900); // 15 minutes expiration

  const subject = '[' + APP_TITLE + '] รหัสยืนยัน (OTP) สำหรับรีเซ็ตรหัสผ่าน';
  const body = 'เรียน คุณ ' + (user.Name || normalizedEmail) + '\n\n' +
    'รหัสยืนยัน (OTP) สำหรับรีเซ็ตรหัสผ่านของคุณคือ: ' + otp + '\n\n' +
    'รหัสนี้มีอายุการใช้งาน 15 นาที หากคุณไม่ได้เป็นผู้ร้องขอ กรุณาเพิกเฉยต่ออีเมลฉบับนี้หรือแจ้งผู้ดูแลระบบ\n\n' +
    '--\n' + APP_TITLE;

  try {
    MailApp.sendEmail({
      to: normalizedEmail,
      subject: subject,
      body: body
    });
  } catch (err) {
    throw new Error('ไม่สามารถส่งอีเมลได้ กรุณาตรวจสอบสิทธิ์การส่งอีเมลของ Apps Script (MailApp) หรือติดต่อ Admin');
  }

  return { success: true, message: 'ระบบได้ส่งรหัส OTP 6 หลักไปยังอีเมล ' + normalizedEmail + ' เรียบร้อยแล้ว (รหัสมีอายุ 15 นาที)' };
}

/**
 * Resets user password using the verified OTP.
 */
function resetPasswordWithOtp(email, otp, newPassword) {
  const normalizedEmail = String(email || '').trim().toLowerCase();
  const suppliedOtp = String(otp || '').trim();
  const password = String(newPassword || '');
  if (!normalizedEmail || !suppliedOtp || !password) {
    throw new Error('กรุณากรอกข้อมูลให้ครบถ้วน');
  }
  if (suppliedOtp.length !== 6 || !/^\d{6}$/.test(suppliedOtp)) {
    throw new Error('รหัส OTP ต้องเป็นตัวเลข 6 หลัก');
  }

  validatePasswordStrength_(password);

  const cache = CacheService.getScriptCache();
  const cacheKey = 'shiftflow-pwd-otp-' + hashPassword_(normalizedEmail, 'otp-key').slice(0, 32);
  const cachedOtp = cache.get(cacheKey);
  if (!cachedOtp || cachedOtp !== suppliedOtp) {
    throw new Error('รหัส OTP ไม่ถูกต้องหรือหมดอายุแล้ว กรุณาขอรหัสใหม่อีกครั้ง');
  }

  const spreadsheet = getOrCreateSpreadsheet_();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('ระบบกำลังประมวลผล กรุณาลองใหม่');
  try {
    const sheet = spreadsheet.getSheetByName('Users');
    const values = sheet.getDataRange().getValues();
    const headers = values[0];
    const emailColumn = headers.indexOf('Email');
    const hashColumn = headers.indexOf('Password Hash');
    const updatedColumn = headers.indexOf('Updated At');
    let targetRow = 0;
    let userId = '';
    for (let index = 1; index < values.length; index++) {
      if (String(values[index][emailColumn] || '').trim().toLowerCase() === normalizedEmail) {
        targetRow = index + 1;
        userId = String(values[index][headers.indexOf('User ID')] || '');
        break;
      }
    }
    if (!targetRow) throw new Error('ไม่พบบัญชีผู้ใช้งาน');

    sheet.getRange(targetRow, hashColumn + 1).setValue(hashPassword_(normalizedEmail, password));
    if (updatedColumn >= 0) {
      sheet.getRange(targetRow, updatedColumn + 1).setValue(new Date()).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    }

    cache.remove(cacheKey);

    appendUserAudit_(spreadsheet, {
      Timestamp: new Date(), Action: 'PASSWORD_RESET', 'User ID': userId, Email: normalizedEmail,
      Role: '', Department: '', Reason: 'รีเซ็ตรหัสผ่านด้วย OTP', 'Performed By': normalizedEmail
    });
    SpreadsheetApp.flush();
    return { success: true, message: 'รีเซ็ตรหัสผ่านเรียบร้อยแล้ว คุณสามารถเข้าสู่ระบบด้วยรหัสผ่านใหม่ได้ทันที' };
  } finally {
    lock.releaseLock();
  }
}

function safeUsers_(spreadsheet) {
  const timeZone = spreadsheet.getSpreadsheetTimeZone();
  return normalizeRows_(readObjects_(spreadsheet.getSheetByName('Users')), timeZone).map(function(user) {
    return Object.keys(user).reduce(function(safeUser, key) {
      if (String(key).toLowerCase().indexOf('password') === -1) safeUser[key] = user[key];
      return safeUser;
    }, {});
  });
}

function getUserAdministration(token) {
  const operator = requireManagerOrAdmin_(token);
  const spreadsheet = getOrCreateSpreadsheet_();
  ensureRequiredSheets_(spreadsheet);
  ensureUserSchema_(spreadsheet);
  const departments = unique_(readObjects_(spreadsheet.getSheetByName('Employees')).map(function(employee) {
    return String(employee.Department || '').trim();
  }).filter(Boolean)).sort();
  const admin = isAdminUser_(operator);
  const users = safeUsers_(spreadsheet).filter(function(user) {
    return admin || String(user.Status || '').trim().toLowerCase() === 'pending';
  });
  return {
    users: users,
    roles: admin ? ALLOWED_USER_ROLES : ['Viewer'],
    departments: ['All'].concat(departments),
    access: admin ? 'Admin' : 'ManagerApprover'
  };
}

function reviewUserAccount(request, token) {
  const operator = requireManagerOrAdmin_(token);
  const operatorIsAdmin = isAdminUser_(operator);
  const input = request || {};
  const userId = String(input.userId || '').trim();
  const action = String(input.action || '').trim().toLowerCase();
  let role = String(input.role || '').trim();
  const department = String(input.department || 'All').trim() || 'All';
  const reason = String(input.reason || '').trim();
  if (!userId) throw new Error('ไม่พบ User ID');
  if (['approve', 'update', 'reject', 'suspend', 'reactivate'].indexOf(action) === -1) throw new Error('คำสั่งจัดการบัญชีไม่ถูกต้อง');
  if (!operatorIsAdmin) {
    if (action !== 'approve') throw new Error('Manager มีสิทธิอนุมัติบัญชีใหม่เท่านั้น');
    if (role && role !== 'Viewer') throw new Error('Manager กำหนดได้เฉพาะ Role Viewer เท่านั้น');
    role = 'Viewer';
  }
  if (['approve', 'update', 'reactivate'].indexOf(action) !== -1 && ALLOWED_USER_ROLES.indexOf(role) === -1) {
    throw new Error('กรุณาเลือก Role ที่ถูกต้อง');
  }
  if (['reject', 'suspend'].indexOf(action) !== -1 && reason.length < 3) throw new Error('กรุณาระบุเหตุผล');

  const spreadsheet = getOrCreateSpreadsheet_();
  ensureRequiredSheets_(spreadsheet);
  ensureUserSchema_(spreadsheet);
  const sheet = spreadsheet.getSheetByName('Users');
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idColumn = headers.indexOf('User ID');
  let targetRow = 0;
  let target = null;
  for (let index = 1; index < values.length; index++) {
    if (String(values[index][idColumn] || '').trim() !== userId) continue;
    targetRow = index + 1;
    target = headers.reduce(function(object, header, column) { object[header] = values[index][column]; return object; }, {});
    break;
  }
  if (!targetRow || !target) throw new Error('ไม่พบบัญชี ' + userId);
  if (!operatorIsAdmin && String(target.Status || '').trim().toLowerCase() !== 'pending') {
    throw new Error('Manager อนุมัติได้เฉพาะบัญชีที่อยู่ในสถานะ Pending เท่านั้น');
  }
  const targetEmail = String(target.Email || '').trim().toLowerCase();
  if (targetEmail === PRIMARY_ADMIN_EMAIL && action !== 'update') throw new Error('ไม่สามารถระงับหรือปฏิเสธบัญชี Primary Admin ได้');

  let nextStatus = String(target.Status || 'Pending');
  let nextRole = String(target.Role || '');
  let nextDepartment = String(target.Department || 'All');
  let auditAction = action.toUpperCase();
  if (action === 'approve' || action === 'reactivate') {
    nextStatus = 'Active'; nextRole = role; nextDepartment = department;
  } else if (action === 'update') {
    nextRole = targetEmail === PRIMARY_ADMIN_EMAIL ? 'Admin' : role;
    nextDepartment = targetEmail === PRIMARY_ADMIN_EMAIL ? 'All' : department;
    nextStatus = 'Active';
  } else if (action === 'reject') {
    nextStatus = 'Rejected'; nextRole = '';
  } else if (action === 'suspend') {
    nextStatus = 'Suspended';
  }
  const now = new Date();
  const record = Object.assign({}, target, {
    Role: nextRole,
    Department: nextDepartment,
    Status: nextStatus,
    'Approved By': ['approve', 'reactivate'].indexOf(action) !== -1 ? operator.email : target['Approved By'],
    'Approved At': ['approve', 'reactivate'].indexOf(action) !== -1 ? now : target['Approved At'],
    'Rejection Reason': ['reject', 'suspend'].indexOf(action) !== -1 ? reason : '',
    'Updated At': now
  });
  sheet.getRange(targetRow, 1, 1, headers.length).setValues([headers.map(function(header) {
    return Object.prototype.hasOwnProperty.call(record, header) ? record[header] : '';
  })]);
  appendUserAudit_(spreadsheet, {
    Timestamp: now, Action: auditAction, 'User ID': userId, Email: targetEmail,
    Role: nextRole, Department: nextDepartment, Reason: reason, 'Performed By': operator.email
  });
  SpreadsheetApp.flush();
  return { userId: userId, email: targetEmail, role: nextRole, department: nextDepartment, status: nextStatus };
}

function normalizeUniqueText_(value) {
  let normalized = String(value || '')
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .trim()
    .replace(/\s+/g, ' ');
  try {
    normalized = normalized.normalize('NFKC');
  } catch (error) {
    // Apps Script V8 supports Unicode normalization; keep the cleaned value as a fallback.
  }
  return normalized.toLowerCase();
}

function normalizeLicenseNumberKey_(value) {
  return normalizeUniqueText_(value).replace(/\s+/g, '').toUpperCase();
}

function addEmployee(employee, token) {
  requireManagerOrAdmin_(token);
  const input = employee || {};
  const record = {
    'Employee ID': String(input.id || '').trim().toUpperCase(),
    Name: String(input.name || '').trim(),
    Department: String(input.department || '').trim(),
    Position: String(input.position || '').trim(),
    Skill: String(input.skill || '').trim(),
    Status: String(input.status || 'Active').trim()
  };
  if (!record['Employee ID'] || !record.Name || !record.Department) {
    throw new Error('กรุณากรอก Employee ID, ชื่อ และแผนกให้ครบ');
  }
  if (['Active', 'Inactive'].indexOf(record.Status) === -1) record.Status = 'Active';

  const spreadsheet = getOrCreateSpreadsheet_();
  ensureRequiredSheets_(spreadsheet);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('ระบบกำลังบันทึกข้อมูล กรุณาลองใหม่อีกครั้ง');
  try {
    const sheet = spreadsheet.getSheetByName('Employees');
    const employees = readObjects_(sheet);
    const duplicateId = employees.some(function(row) {
      return String(row['Employee ID'] || '').trim().toUpperCase() === record['Employee ID'];
    });
    if (duplicateId) throw new Error('Employee ID ' + record['Employee ID'] + ' มีอยู่แล้ว');

    const nameKey = normalizeUniqueText_(record.Name);
    const duplicateName = employees.find(function(row) {
      return normalizeUniqueText_(row.Name) === nameKey;
    });
    if (duplicateName) {
      const existingId = String(duplicateName['Employee ID'] || '').trim();
      throw new Error('ชื่อ–นามสกุล "' + record.Name + '" มีอยู่แล้ว' + (existingId ? 'ในรหัส ' + existingId : ''));
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    sheet.appendRow(headers.map(function(header) {
      return Object.prototype.hasOwnProperty.call(record, header) ? record[header] : '';
    }));
    SpreadsheetApp.flush();
    bumpDataVersion_();
    return { id: record['Employee ID'], name: record.Name };
  } finally {
    lock.releaseLock();
  }
}

function updateEmployee(employee, token) {
  const operator = requireManagerOrAdmin_(token);
  const input = employee || {};
  const originalId = String(input.originalId || input.id || '').trim().toUpperCase();
  const record = {
    'Employee ID': String(input.id || '').trim().toUpperCase(),
    Name: String(input.name || '').trim(),
    Department: String(input.department || '').trim(),
    Position: String(input.position || '').trim(),
    Skill: String(input.skill || '').trim(),
    Status: String(input.status || 'Active').trim()
  };
  if (!record['Employee ID'] || !record.Name || !record.Department) {
    throw new Error('กรุณากรอก Employee ID, ชื่อ และแผนกให้ครบ');
  }
  if (originalId !== record['Employee ID']) throw new Error('ไม่อนุญาตให้เปลี่ยน Employee ID');
  if (['Active', 'Inactive'].indexOf(record.Status) === -1) record.Status = 'Active';

  const spreadsheet = getOrCreateSpreadsheet_();
  ensureRequiredSheets_(spreadsheet);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('ระบบกำลังบันทึกข้อมูล กรุณาลองใหม่อีกครั้ง');
  try {
    const sheet = spreadsheet.getSheetByName('Employees');
    const values = sheet.getDataRange().getValues();
    const headers = values[0] || [];
    const idColumn = headers.indexOf('Employee ID');
    const nameColumn = headers.indexOf('Name');
    if (idColumn === -1 || nameColumn === -1) throw new Error('โครงสร้างชีต Employees ไม่ครบถ้วน');

    let targetRow = 0;
    const nameKey = normalizeUniqueText_(record.Name);
    for (let index = 1; index < values.length; index++) {
      const existingId = String(values[index][idColumn] || '').trim().toUpperCase();
      if (existingId === record['Employee ID']) targetRow = index + 1;
      if (existingId !== record['Employee ID'] && normalizeUniqueText_(values[index][nameColumn]) === nameKey) {
        throw new Error('ชื่อ–นามสกุล "' + record.Name + '" มีอยู่แล้วในรหัส ' + existingId);
      }
    }
    if (!targetRow) throw new Error('ไม่พบพนักงานรหัส ' + record['Employee ID']);

    sheet.getRange(targetRow, 1, 1, headers.length).setValues([headers.map(function(header) {
      return Object.prototype.hasOwnProperty.call(record, header) ? record[header] : values[targetRow - 1][headers.indexOf(header)];
    })]);

    const scheduleSheet = spreadsheet.getSheetByName('Schedule');
    const scheduleHeaders = ensureScheduleMetadataColumns_(scheduleSheet);
    const scheduleValues = scheduleSheet.getDataRange().getValues();
    const scheduleIdColumn = scheduleHeaders.indexOf('Employee ID');
    const scheduleNameColumn = scheduleHeaders.indexOf('Employee Name');
    const scheduleDepartmentColumn = scheduleHeaders.indexOf('Department');
    const scheduleDateColumn = scheduleHeaders.indexOf('Date');
    const updatedByColumn = scheduleHeaders.indexOf('Updated By');
    const updatedAtColumn = scheduleHeaders.indexOf('Updated At');
    const timeZone = spreadsheet.getSpreadsheetTimeZone();
    const affectedMonths = [];
    let updatedScheduleRows = 0;
    for (let index = 1; index < scheduleValues.length; index++) {
      if (String(scheduleValues[index][scheduleIdColumn] || '').trim().toUpperCase() !== record['Employee ID']) continue;
      if (scheduleNameColumn >= 0) scheduleValues[index][scheduleNameColumn] = record.Name;
      if (scheduleDepartmentColumn >= 0) scheduleValues[index][scheduleDepartmentColumn] = record.Department;
      if (updatedByColumn >= 0) scheduleValues[index][updatedByColumn] = operator.email;
      if (updatedAtColumn >= 0) scheduleValues[index][updatedAtColumn] = new Date();
      const date = formatDate_(scheduleValues[index][scheduleDateColumn], timeZone);
      if (date) affectedMonths.push(date.slice(0, 7));
      updatedScheduleRows++;
    }
    if (updatedScheduleRows) {
      const scheduleRows = scheduleValues.slice(1);
      [scheduleNameColumn, scheduleDepartmentColumn, updatedByColumn, updatedAtColumn].forEach(function(column) {
        if (column < 0) return;
        scheduleSheet.getRange(2, column + 1, scheduleRows.length, 1).setValues(scheduleRows.map(function(row) {
          return [row[column] === undefined ? '' : row[column]];
        }));
      });
    }

    SpreadsheetApp.flush();
    const approvals = invalidateScheduleApprovals_(spreadsheet, affectedMonths, operator.email, 'EMPLOYEE_UPDATED');
    bumpDataVersion_();
    return {
      id: record['Employee ID'],
      name: record.Name,
      updatedScheduleRows: updatedScheduleRows,
      approvalsInvalidated: approvals.length
    };
  } finally {
    lock.releaseLock();
  }
}

function saveEmployeeLicense(license, token) {
  const operator = requireManagerOrAdmin_(token);
  const input = license || {};
  const spreadsheet = getOrCreateSpreadsheet_();
  ensureRequiredSheets_(spreadsheet);
  const timeZone = spreadsheet.getSpreadsheetTimeZone();
  const employeeId = String(input.employeeId || '').trim().toUpperCase();

  const issueDate = String(input.issueDate || '').trim();
  const expiryDate = String(input.expiryDate || '').trim();
  parseIsoDateSafe_(issueDate);
  parseIsoDateSafe_(expiryDate);
  if (issueDate > expiryDate) throw new Error('วันเริ่มใช้ต้องไม่เกินวันหมดอายุ');

  const status = String(input.status || 'Active').trim();
  if (['Active', 'Suspended', 'Revoked', 'Inactive'].indexOf(status) === -1) {
    throw new Error('สถานะใบอนุญาตไม่ถูกต้อง');
  }
  const licenseType = String(input.licenseType || '').trim();
  const licenseNumber = String(input.licenseNumber || '').trim();
  if (!licenseType || !licenseNumber) throw new Error('กรุณากรอกประเภทและเลขที่ใบอนุญาต');
  const documentUrl = String(input.documentUrl || '').trim();
  if (documentUrl && !/^https:\/\/(drive|docs)\.google\.com\//i.test(documentUrl)) {
    throw new Error('ลิงก์เอกสารต้องเป็น Google Drive หรือ Google Docs แบบ https เท่านั้น');
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('ระบบกำลังบันทึกข้อมูล กรุณาลองใหม่อีกครั้ง');
  try {
  const employeeExists = readObjects_(spreadsheet.getSheetByName('Employees')).some(function(employee) {
    return String(employee['Employee ID'] || '').trim().toUpperCase() === employeeId;
  });
  if (!employeeExists) throw new Error('ไม่พบพนักงานรหัส ' + employeeId);

  const sheet = spreadsheet.getSheetByName('Employee Licenses');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const licenseId = String(input.licenseId || '').trim().toUpperCase() || ('LIC-' + Utilities.getUuid().replace(/-/g, '').slice(0, 10).toUpperCase());
  const values = sheet.getDataRange().getValues();
  const idColumn = headers.indexOf('License ID');
  const numberColumn = headers.indexOf('License Number');
  const licenseEmployeeColumn = headers.indexOf('Employee ID');
  const typeColumn = headers.indexOf('License Type');
  const licenseNumberKey = normalizeLicenseNumberKey_(licenseNumber);
  const licenseTypeKey = normalizeUniqueText_(licenseType);
  let targetRow = 0;
  for (let index = 1; index < values.length; index++) {
    const existingId = String(values[index][idColumn] || '').trim().toUpperCase();
    const existingNumber = String(values[index][numberColumn] || '').trim();
    const existingNumberKey = normalizeLicenseNumberKey_(existingNumber);
    const existingEmployeeId = String(values[index][licenseEmployeeColumn] || '').trim().toUpperCase();
    const existingType = String(values[index][typeColumn] || '').trim();
    if (existingNumberKey && existingNumberKey === licenseNumberKey && existingId !== licenseId) {
      throw new Error('เลขที่ใบอนุญาต ' + licenseNumber + ' ถูกใช้งานแล้ว');
    }
    if (existingEmployeeId === employeeId && normalizeUniqueText_(existingType) === licenseTypeKey && existingId !== licenseId) {
      throw new Error('พนักงานรหัส ' + employeeId + ' มีใบอนุญาตประเภท "' + licenseType + '" อยู่แล้ว' + (existingNumber ? ' (เลขที่ ' + existingNumber + ')' : '') + ' กรุณาแก้ไขรายการเดิมแทนการเพิ่มใหม่');
    }
    if (existingId === licenseId) {
      targetRow = index + 1;
      break;
    }
  }
  const updating = targetRow > 0;
  if (updating && !isAdminUser_(operator)) {
    throw new Error('Manager เพิ่มใบอนุญาตใหม่ได้ แต่การแก้ไขใบอนุญาตเดิมต้องใช้สิทธิ Admin');
  }

  const record = {
    'License ID': licenseId,
    'Employee ID': employeeId,
    'License Type': licenseType,
    'License Number': licenseNumber,
    'Issue Date': parseIsoDateSafe_(issueDate),
    'Expiry Date': parseIsoDateSafe_(expiryDate),
    Status: status,
    'Document URL': documentUrl,
    Remark: String(input.remark || '').trim(),
    'Updated By': operator.email,
    'Updated At': new Date()
  };
  const output = headers.map(function(header) {
    return Object.prototype.hasOwnProperty.call(record, header) ? record[header] : '';
  });
  if (targetRow) sheet.getRange(targetRow, 1, 1, headers.length).setValues([output]);
  else {
    sheet.appendRow(output);
    targetRow = sheet.getLastRow();
  }
  sheet.getRange(targetRow, headers.indexOf('Issue Date') + 1, 1, 2).setNumberFormat('yyyy-mm-dd');

  appendLicenseAudit_(spreadsheet, {
    Timestamp: new Date(),
    Action: updating ? 'LICENSE_UPDATED' : 'LICENSE_CREATED',
    'Employee ID': employeeId,
    'License ID': licenseId,
    'License Status': status,
    'Expiry Date': parseIsoDateSafe_(expiryDate),
    Reason: String(input.remark || ''),
    'Approved By': operator.email
  });
  SpreadsheetApp.flush();
  bumpDataVersion_();
  return { licenseId: licenseId, employeeId: employeeId, expiryDate: expiryDate, status: status };
  } finally {
    lock.releaseLock();
  }
}

function deleteEmployeeLicense(licenseId, token) {
  const admin = requireAdmin_(token);
  const normalizedId = String(licenseId || '').trim().toUpperCase();
  if (!normalizedId) throw new Error('ไม่พบ License ID ที่ต้องการลบ');
  const spreadsheet = getOrCreateSpreadsheet_();
  ensureRequiredSheets_(spreadsheet);
  const sheet = spreadsheet.getSheetByName('Employee Licenses');
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idColumn = headers.indexOf('License ID');
  const employeeColumn = headers.indexOf('Employee ID');
  const statusColumn = headers.indexOf('Status');
  const expiryColumn = headers.indexOf('Expiry Date');
  let targetRow = 0;
  let employeeId = '';
  let status = '';
  let expiryDate = '';
  for (let index = 1; index < values.length; index++) {
    if (String(values[index][idColumn] || '').trim().toUpperCase() !== normalizedId) continue;
    targetRow = index + 1;
    employeeId = String(values[index][employeeColumn] || '').trim();
    status = String(values[index][statusColumn] || '').trim();
    expiryDate = values[index][expiryColumn];
    break;
  }
  if (!targetRow) throw new Error('ไม่พบใบอนุญาต ' + normalizedId);

  appendLicenseAudit_(spreadsheet, {
    Timestamp: new Date(),
    Action: 'LICENSE_DELETED',
    'Employee ID': employeeId,
    'License ID': normalizedId,
    'License Status': status,
    'Expiry Date': expiryDate,
    Reason: 'ลบข้อมูลใบอนุญาต',
    'Approved By': admin.email
  });
  sheet.deleteRow(targetRow);
  SpreadsheetApp.flush();
  bumpDataVersion_();
  return { licenseId: normalizedId, employeeId: employeeId };
}

function deleteEmployee(employeeId, token) {
  const admin = requireAdmin_(token);
  const normalizedId = String(employeeId || '').trim().toUpperCase();
  if (!normalizedId) throw new Error('ไม่พบ Employee ID ที่ต้องการลบ');

  const spreadsheet = getOrCreateSpreadsheet_();
  ensureRequiredSheets_(spreadsheet);
  const employeeSheet = spreadsheet.getSheetByName('Employees');
  const employeeValues = employeeSheet.getDataRange().getValues();
  const idColumn = employeeValues[0].indexOf('Employee ID');
  const nameColumn = employeeValues[0].indexOf('Name');
  if (idColumn === -1) throw new Error('ไม่พบคอลัมน์ Employee ID');

  let employeeRow = -1;
  let employeeName = '';
  for (let index = 1; index < employeeValues.length; index++) {
    if (String(employeeValues[index][idColumn] || '').trim().toUpperCase() === normalizedId) {
      employeeRow = index + 1;
      employeeName = nameColumn >= 0 ? String(employeeValues[index][nameColumn] || '') : '';
      break;
    }
  }
  if (employeeRow === -1) throw new Error('ไม่พบพนักงานรหัส ' + normalizedId);
  employeeSheet.deleteRow(employeeRow);

  let deletedScheduleRows = 0;
  const affectedMonths = [];
  const scheduleSheet = spreadsheet.getSheetByName('Schedule');
  if (scheduleSheet && scheduleSheet.getLastRow() >= 2) {
    const values = scheduleSheet.getDataRange().getValues();
    const scheduleIdColumn = values[0].indexOf('Employee ID');
    const scheduleDateColumn = values[0].indexOf('Date');
    if (scheduleIdColumn >= 0) {
      for (let index = values.length - 1; index >= 1; index--) {
        if (String(values[index][scheduleIdColumn] || '').trim().toUpperCase() === normalizedId) {
          if (scheduleDateColumn >= 0) affectedMonths.push(formatDate_(values[index][scheduleDateColumn], spreadsheet.getSpreadsheetTimeZone()).slice(0, 7));
          scheduleSheet.deleteRow(index + 1);
          deletedScheduleRows++;
        }
      }
    }
  }
  SpreadsheetApp.flush();
  invalidateScheduleApprovals_(spreadsheet, affectedMonths, admin.email, 'EMPLOYEE_REMOVED');
  bumpDataVersion_();
  return { id: normalizedId, name: employeeName, deletedScheduleRows: deletedScheduleRows };
}

function addShiftType(shift, token) {
  requireAdmin_(token);
  const input = shift || {};
  const record = {
    'Shift Code': String(input.code || '').trim().toUpperCase(),
    'Shift Name': String(input.name || '').trim(),
    'Start Time': String(input.startTime || '').trim(),
    'End Time': String(input.endTime || '').trim(),
    Hours: Number(input.hours),
    Color: String(input.color || '#2F80FF').trim().toUpperCase()
  };

  if (!record['Shift Code'] || !record['Shift Name']) {
    throw new Error('กรุณากรอก Shift Code และชื่อกะ');
  }
  if (!/^[A-Z0-9_-]{1,12}$/.test(record['Shift Code'])) {
    throw new Error('Shift Code ใช้ได้เฉพาะ A-Z, 0-9, _ และ - ไม่เกิน 12 ตัว');
  }
  if (!Number.isFinite(record.Hours) || record.Hours < 0) {
    throw new Error('จำนวนชั่วโมงไม่ถูกต้อง');
  }
  if (record.Hours > 0 && (!record['Start Time'] || !record['End Time'])) {
    throw new Error('กะทำงานต้องระบุเวลาเริ่มและเวลาเลิก');
  }
  if (!/^#[0-9A-F]{6}$/.test(record.Color)) record.Color = '#2F80FF';

  const spreadsheet = getOrCreateSpreadsheet_();
  ensureRequiredSheets_(spreadsheet);
  const sheet = spreadsheet.getSheetByName('Shift Types');
  const duplicate = readObjects_(sheet).some(function(row) {
    return String(row['Shift Code'] || '').trim().toUpperCase() === record['Shift Code'];
  });
  if (duplicate) throw new Error('Shift Code ' + record['Shift Code'] + ' มีอยู่แล้ว');

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const newRowNumber = sheet.getLastRow() + 1;
  sheet.appendRow(headers.map(function(header) {
    return Object.prototype.hasOwnProperty.call(record, header) ? record[header] : '';
  }));

  ['Start Time', 'End Time'].forEach(function(header) {
    const column = headers.indexOf(header) + 1;
    if (column) sheet.getRange(newRowNumber, column).setNumberFormat('@');
  });

  SpreadsheetApp.flush();
  bumpDataVersion_();
  return { code: record['Shift Code'], name: record['Shift Name'] };
}

function deleteShiftType(shiftCode, token) {
  const admin = requireAdmin_(token);
  const normalizedCode = String(shiftCode || '').trim().toUpperCase();
  if (!normalizedCode) throw new Error('ไม่พบ Shift Code ที่ต้องการลบ');

  const protectedCodes = ['D', 'N', 'OFF', 'AL'];
  if (protectedCodes.indexOf(normalizedCode) !== -1) {
    throw new Error('ไม่สามารถลบรหัสกะหลัก ' + normalizedCode + ' ได้');
  }

  const spreadsheet = getOrCreateSpreadsheet_();
  ensureRequiredSheets_(spreadsheet);
  const shiftSheet = spreadsheet.getSheetByName('Shift Types');
  const shiftValues = shiftSheet.getDataRange().getValues();
  const codeColumn = shiftValues[0].indexOf('Shift Code');
  if (codeColumn === -1) throw new Error('ไม่พบคอลัมน์ Shift Code');

  let shiftRow = -1;
  let shiftName = '';
  const nameColumn = shiftValues[0].indexOf('Shift Name');
  for (let index = 1; index < shiftValues.length; index++) {
    if (String(shiftValues[index][codeColumn] || '').trim().toUpperCase() === normalizedCode) {
      shiftRow = index + 1;
      shiftName = nameColumn >= 0 ? String(shiftValues[index][nameColumn] || '') : '';
      break;
    }
  }
  if (shiftRow === -1) throw new Error('ไม่พบกะรหัส ' + normalizedCode);
  shiftSheet.deleteRow(shiftRow);

  let updatedScheduleRows = 0;
  const affectedMonths = [];
  const schedule = spreadsheet.getSheetByName('Schedule');
  if (schedule && schedule.getLastRow() >= 2) {
    const headers = schedule.getRange(1, 1, 1, schedule.getLastColumn()).getValues()[0];
    const codeCol = headers.indexOf('Shift Code') + 1;
    const startCol = headers.indexOf('Start Time') + 1;
    const endCol = headers.indexOf('End Time') + 1;
    const hoursCol = headers.indexOf('Hours') + 1;
    const dateCol = headers.indexOf('Date') + 1;

    if (codeCol && startCol && endCol && hoursCol) {
      const rowCount = schedule.getLastRow() - 1;
      const codes = schedule.getRange(2, codeCol, rowCount, 1).getValues();
      const starts = schedule.getRange(2, startCol, rowCount, 1).getValues();
      const ends = schedule.getRange(2, endCol, rowCount, 1).getValues();
      const hours = schedule.getRange(2, hoursCol, rowCount, 1).getValues();
      const dates = dateCol ? schedule.getRange(2, dateCol, rowCount, 1).getValues() : [];

      codes.forEach(function(row, index) {
        if (String(row[0] || '').trim().toUpperCase() !== normalizedCode) return;
        codes[index][0] = 'OFF';
        starts[index][0] = '';
        ends[index][0] = '';
        hours[index][0] = 0;
        if (dates[index]) affectedMonths.push(formatDate_(dates[index][0], spreadsheet.getSpreadsheetTimeZone()).slice(0, 7));
        updatedScheduleRows++;
      });

      if (updatedScheduleRows) {
        schedule.getRange(2, codeCol, rowCount, 1).setValues(codes);
        schedule.getRange(2, startCol, rowCount, 1).setValues(starts);
        schedule.getRange(2, endCol, rowCount, 1).setValues(ends);
        schedule.getRange(2, hoursCol, rowCount, 1).setValues(hours);
      }
    }
  }

  SpreadsheetApp.flush();
  invalidateScheduleApprovals_(spreadsheet, affectedMonths, admin.email, 'SHIFT_TYPE_REMOVED');
  bumpDataVersion_();
  return { code: normalizedCode, name: shiftName, updatedScheduleRows: updatedScheduleRows };
}


/* ------------------------------------------------------- secure scheduling */

function migrateLegacyDemoScheduleYear_(sheet) {
  const property = 'SHIFTFLOW_LEGACY_2024_DATES_MIGRATED';
  const properties = PropertiesService.getScriptProperties();
  if (properties.getProperty(property) === 'true' || !sheet || sheet.getLastRow() < 2) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const dateColumn = headers.indexOf('Date') + 1;
  const employeeColumn = headers.indexOf('Employee ID') + 1;
  if (!dateColumn || !employeeColumn) {
    properties.setProperty(property, 'true');
    return;
  }

  const rowCount = sheet.getLastRow() - 1;
  const dates = sheet.getRange(2, dateColumn, rowCount, 1).getValues();
  const employees = sheet.getRange(2, employeeColumn, rowCount, 1).getValues();
  const demoIds = ['EMP001', 'EMP002', 'EMP003', 'EMP004', 'EMP005', 'EMP006'];
  const currentYear = new Date().getFullYear();
  let changed = false;

  dates.forEach(function(row, index) {
    const date = row[0] instanceof Date ? row[0] : new Date(row[0]);
    const employeeId = String(employees[index][0] || '').trim();
    if (isNaN(date.getTime()) || date.getFullYear() !== 2024 ||
        date.getMonth() !== 6 || date.getDate() > 7 || demoIds.indexOf(employeeId) === -1) return;
    date.setFullYear(currentYear);
    row[0] = date;
    changed = true;
  });

  if (changed) {
    sheet.getRange(2, dateColumn, rowCount, 1).setValues(dates).setNumberFormat('yyyy-mm-dd');
  }
  properties.setProperty(property, 'true');
}

function enforceMaxWeeklyHours72_(sheet) {
  if (!sheet) throw new Error('Rules sheet was not found.');
  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  const idColumn = headers.indexOf('Rule ID') + 1;
  const valueColumn = headers.indexOf('Value') + 1;
  if (!idColumn || !valueColumn) throw new Error('Rules sheet must contain Rule ID and Value columns.');

  const rowCount = Math.max(sheet.getLastRow() - 1, 0);
  if (rowCount) {
    const ids = sheet.getRange(2, idColumn, rowCount, 1).getValues();
    for (let index = 0; index < ids.length; index++) {
      if (String(ids[index][0] || '').trim().toUpperCase() !== 'RULE001') continue;
      sheet.getRange(index + 2, valueColumn).setValue(72);
      return;
    }
  }

  const record = {
    'Rule ID': 'RULE001',
    'Rule Name': 'Maximum working hours per week',
    Value: 72,
    Unit: 'hours',
    Enabled: true
  };
  sheet.appendRow(headers.map(function(header) {
    return Object.prototype.hasOwnProperty.call(record, header) ? record[header] : '';
  }));
}

function requireActiveSession_(token) {
  if (token && typeof token === 'object') {
    token = token.token || '';
  }
  token = String(token || '').trim();
  if (token === '[object Object]') token = '';
  const key = 'shiftflow-session-' + String(token || '').trim();
  const cached = CacheService.getScriptCache().get(key);
  if (!cached) throw new Error('Session expired. Please sign in again.');
  const cachedUser = JSON.parse(cached);
  const spreadsheet = getOrCreateSpreadsheet_();
  const user = readObjects_(spreadsheet.getSheetByName('Users')).find(function(row) {
    return String(row.Email || '').trim().toLowerCase() === String(cachedUser.email || '').trim().toLowerCase();
  });
  if (!user || String(user.Status || '').trim().toLowerCase() !== 'active') {
    CacheService.getScriptCache().remove(key);
    throw new Error('บัญชีนี้ยังไม่ได้รับอนุมัติหรือถูกระงับ กรุณาเข้าสู่ระบบใหม่');
  }
  const role = String(user.Role || '').trim();
  if (ALLOWED_USER_ROLES.map(r=>r.toLowerCase()).indexOf(role.toLowerCase()) === -1 && !isAdminOrManager_(user)) throw new Error('บัญชียังไม่ได้รับการกำหนด Role');
  const idVal = String(user['User ID'] || user.id || '');
  const nameVal = String(user.Name || user.FullName || user.name || user.email || '').trim();
  const emailVal = String(user.Email || user.email || '').trim();
  const deptVal = String(user.Department || user.department || 'All').trim();

  const sessionUser = {
    id: idVal,
    'User ID': idVal,
    name: nameVal,
    Name: nameVal,
    FullName: nameVal,
    email: emailVal,
    Email: emailVal,
    role: role,
    Role: role,
    department: deptVal,
    Department: deptVal,
    status: 'Active',
    Status: 'Active'
  };
  CacheService.getScriptCache().put(key, JSON.stringify(sessionUser), 21600);
  return sessionUser;
}

function validateSession(token) {
  return Object.assign({}, requireActiveSession_(token), { token: String(token || '') });
}

function logoutUser(token) {
  const normalizedToken = String(token || '').trim();
  if (normalizedToken) CacheService.getScriptCache().remove('shiftflow-session-' + normalizedToken);
  return true;
}

function requireAdmin_(token) {
  const user = requireActiveSession_(token);
  if (String(user.role || '').trim().toLowerCase() !== 'admin') {
    throw new Error('เฉพาะ Admin เท่านั้นที่ดำเนินการนี้ได้');
  }
  return user;
}

function requireManagerOrAdmin_(token) {
  const user = requireActiveSession_(token);
  const role = String(user.role || '').trim().toLowerCase();
  if (['admin', 'manager'].indexOf(role) === -1) {
    throw new Error('เฉพาะ Admin หรือ Manager เท่านั้นที่ดำเนินการนี้ได้');
  }
  return user;
}

function isAdminUser_(user) {
  return String((user || {}).role || '').trim().toLowerCase() === 'admin';
}

function ensureScheduleMetadataColumns_(sheet) {
  const required = [
    'Source', 'Locked', 'Updated By', 'Updated At',
    'License Status', 'License Expiry Date', 'License Override',
    'Override Reason', 'Override By', 'Override At'
  ];
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  required.forEach(function(header) {
    if (headers.indexOf(header) !== -1) return;
    headers.push(header);
    sheet.getRange(1, headers.length).setValue(header);
  });
  return headers;
}

function ensureScheduleApprovalSchema_(spreadsheet) {
  const schemas = {
    'Schedule Approvals': [
      'Month', 'Status', 'Revision', 'Changed By', 'Changed At', 'Change Type',
      'Approved By', 'Approved At', 'Approval Note', 'Schedule Hash'
    ],
    'Schedule Approval Log': [
      'Timestamp', 'Action', 'Month', 'Revision', 'Status', 'Change Type', 'Performed By', 'Note'
    ]
  };
  Object.keys(schemas).forEach(function(sheetName) {
    const sheet = spreadsheet.getSheetByName(sheetName);
    const required = schemas[sheetName];
    const lastColumn = Math.max(sheet.getLastColumn(), 1);
    const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    if (!headers.some(function(header) { return String(header || '').trim(); })) {
      sheet.getRange(1, 1, 1, required.length).setValues([required]);
    } else {
      required.forEach(function(header) {
        if (headers.indexOf(header) !== -1) return;
        headers.push(header);
        sheet.getRange(1, headers.length).setValue(header);
      });
    }
    sheet.setFrozenRows(1);
    const finalHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const monthColumn = finalHeaders.indexOf('Month') + 1;
    if (monthColumn && sheet.getMaxRows() > 1) {
      sheet.getRange(2, monthColumn, sheet.getMaxRows() - 1, 1).setNumberFormat('@');
    }
  });
}

function approvalMonthKey_(value, timeZone) {
  if (isDate_(value)) return Utilities.formatDate(value, timeZone, 'yyyy-MM');
  const text = String(value || '').trim();
  const match = text.match(/^(\d{4})-(\d{2})/);
  if (match) return match[1] + '-' + match[2];
  const parsed = new Date(text);
  return isNaN(parsed.getTime()) ? text.slice(0, 7) : Utilities.formatDate(parsed, timeZone, 'yyyy-MM');
}

function scheduleApprovalState_(spreadsheet, month) {
  ensureScheduleApprovalSchema_(spreadsheet);
  const timeZone = spreadsheet.getSpreadsheetTimeZone();
  const normalizedMonth = approvalMonthKey_(month, timeZone);
  const rows = readObjects_(spreadsheet.getSheetByName('Schedule Approvals'));
  const matchingRows = rows.filter(function(item) {
    return approvalMonthKey_(item.Month, timeZone) === normalizedMonth;
  });
  const row = matchingRows.length ? matchingRows[matchingRows.length - 1] : null;
  if (!row) {
    return {
      month: normalizedMonth,
      status: 'Pending Approval',
      revision: 0,
      changedBy: '',
      changedAt: '',
      changeType: 'Existing Schedule',
      approvedBy: '',
      approvedAt: '',
      approvalNote: '',
      scheduleHash: ''
    };
  }
  return {
    month: normalizedMonth,
    status: String(row.Status || 'Pending Approval'),
    revision: Number(row.Revision) || 0,
    changedBy: String(row['Changed By'] || ''),
    changedAt: row['Changed At'] ? formatDateTime_(row['Changed At'], timeZone) : '',
    changeType: String(row['Change Type'] || ''),
    approvedBy: String(row['Approved By'] || ''),
    approvedAt: row['Approved At'] ? formatDateTime_(row['Approved At'], timeZone) : '',
    approvalNote: String(row['Approval Note'] || ''),
    scheduleHash: String(row['Schedule Hash'] || '')
  };
}

function writeScheduleApproval_(spreadsheet, record) {
  ensureScheduleApprovalSchema_(spreadsheet);
  const sheet = spreadsheet.getSheetByName('Schedule Approvals');
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const monthColumn = headers.indexOf('Month');
  const timeZone = spreadsheet.getSpreadsheetTimeZone();
  const normalizedMonth = approvalMonthKey_(record.Month, timeZone);
  let targetRow = 0;
  for (let index = 1; index < values.length; index++) {
    if (approvalMonthKey_(values[index][monthColumn], timeZone) !== normalizedMonth) continue;
    targetRow = index + 1;
  }
  const normalizedRecord = Object.assign({}, record, { Month: normalizedMonth });
  const output = headers.map(function(header) {
    return Object.prototype.hasOwnProperty.call(normalizedRecord, header) ? normalizedRecord[header] : '';
  });
  if (targetRow) sheet.getRange(targetRow, 1, 1, headers.length).setValues([output]);
  else {
    sheet.appendRow(output);
    targetRow = sheet.getLastRow();
  }
  sheet.getRange(targetRow, monthColumn + 1).setNumberFormat('@').setValue(normalizedMonth);
}

function appendScheduleApprovalLog_(spreadsheet, record) {
  const sheet = spreadsheet.getSheetByName('Schedule Approval Log');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  sheet.appendRow(headers.map(function(header) {
    return Object.prototype.hasOwnProperty.call(record, header) ? record[header] : '';
  }));
}

function invalidateScheduleApprovals_(spreadsheet, months, changedBy, changeType) {
  const uniqueMonths = unique_((months || []).map(function(month) {
    return String(month || '').trim().slice(0, 7);
  }).filter(function(month) { return /^\d{4}-\d{2}$/.test(month); }));
  const results = [];
  uniqueMonths.forEach(function(month) {
    const current = scheduleApprovalState_(spreadsheet, month);
    const now = new Date();
    const revision = (String(current.status || '').toLowerCase() === 'approved') ? (current.revision + 1) : Math.max(1, current.revision);
    writeScheduleApproval_(spreadsheet, {
      Month: month,
      Status: 'Pending Approval',
      Revision: revision,
      'Changed By': String(changedBy || ''),
      'Changed At': now,
      'Change Type': String(changeType || 'Schedule Changed'),
      'Approved By': '',
      'Approved At': '',
      'Approval Note': '',
      'Schedule Hash': ''
    });
    appendScheduleApprovalLog_(spreadsheet, {
      Timestamp: now,
      Action: 'INVALIDATED',
      Month: month,
      Revision: revision,
      Status: 'Pending Approval',
      'Change Type': String(changeType || 'Schedule Changed'),
      'Performed By': String(changedBy || ''),
      Note: 'ต้องให้ Admin อนุมัติใหม่หลังมีการเปลี่ยนแปลง'
    });
    results.push({ month: month, status: 'Pending Approval', revision: revision });
  });
  return results;
}

function scheduleMonthHash_(spreadsheet, month) {
  const timeZone = spreadsheet.getSpreadsheetTimeZone();
  const normalizedMonth = approvalMonthKey_(month, timeZone);
  const rows = readSchedule_(spreadsheet, timeZone).filter(function(row) {
    return String(row.date || '').slice(0, 7) === normalizedMonth;
  }).map(function(row) {
    return [
      row.date, row.employeeId, row.employeeName, row.department, row.code,
      row.startTime, row.endTime, Number(row.hours) || 0, row.remark,
      row.source, Boolean(row.locked), row.licenseStatus,
      row.licenseExpiryDate, Boolean(row.licenseOverride), row.overrideReason
    ];
  }).sort(function(first, second) {
    return String(first[0]).localeCompare(String(second[0])) || String(first[1]).localeCompare(String(second[1]));
  });
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    JSON.stringify(rows),
    Utilities.Charset.UTF_8
  );
  return Utilities.base64EncodeWebSafe(digest).replace(/=+$/, '');
}

function ensureScheduleApprovalCurrent_(spreadsheet, month, changedBy) {
  const current = scheduleApprovalState_(spreadsheet, month);
  if (String(current.status || '').toLowerCase() !== 'approved') return current;
  const currentHash = scheduleMonthHash_(spreadsheet, month);
  if (current.scheduleHash && current.scheduleHash === currentHash) return current;
  invalidateScheduleApprovals_(
    spreadsheet,
    [month],
    String(changedBy || 'SYSTEM'),
    current.scheduleHash ? 'SCHEDULE_DATA_CHANGED' : 'REVISION_TRACKING_ENABLED'
  );
  SpreadsheetApp.flush();
  return scheduleApprovalState_(spreadsheet, month);
}

function approveScheduleMonth(month, note, token) {
  const admin = requireAdmin_(token);
  const normalized = String(month || '').trim().slice(0, 7);
  if (!/^\d{4}-\d{2}$/.test(normalized)) throw new Error('กรุณาเลือกเดือนที่ต้องการอนุมัติ');
  const spreadsheet = getOrCreateSpreadsheet_();
  ensureRequiredSheets_(spreadsheet);
  ensureScheduleApprovalSchema_(spreadsheet);
  const timeZone = spreadsheet.getSpreadsheetTimeZone();
  const hasSchedule = readSchedule_(spreadsheet, timeZone).some(function(row) {
    return String(row.date || '').slice(0, 7) === normalized;
  });
  if (!hasSchedule) throw new Error('ไม่พบตารางกะของเดือน ' + normalized + ' ที่จะอนุมัติ');
  const current = ensureScheduleApprovalCurrent_(spreadsheet, normalized, admin.email);
  const now = new Date();
  const approvalNote = String(note || '').trim().slice(0, 500);
  writeScheduleApproval_(spreadsheet, {
    Month: normalized,
    Status: 'Approved',
    Revision: current.revision,
    'Changed By': current.changedBy,
    'Changed At': current.changedAt ? new Date(current.changedAt.replace(' ', 'T')) : '',
    'Change Type': current.changeType,
    'Approved By': admin.email,
    'Approved At': now,
    'Approval Note': approvalNote,
    'Schedule Hash': scheduleMonthHash_(spreadsheet, normalized)
  });
  appendScheduleApprovalLog_(spreadsheet, {
    Timestamp: now,
    Action: 'APPROVED',
    Month: normalized,
    Revision: current.revision,
    Status: 'Approved',
    'Change Type': current.changeType,
    'Performed By': admin.email,
    Note: approvalNote
  });
  SpreadsheetApp.flush();
  return scheduleApprovalState_(spreadsheet, normalized);
}

function buddhistMonthLabel_(month) {
  const match = String(month || '').trim().match(/^(\d{4})-(\d{2})$/);
  if (!match) return String(month || '');
  const monthNames = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];
  return monthNames[Number(match[2]) - 1] + ' ' + (Number(match[1]) + 543);
}

function buddhistDateTimeText_(value) {
  const text = String(value || '').trim();
  const match = text.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}:\d{2}(?::\d{2})?))?/);
  if (!match) return text;
  return match[3] + '/' + match[2] + '/' + (Number(match[1]) + 543) + (match[4] ? ' ' + match[4] : '');
}

function safeExportFilePart_(value, fallback) {
  const cleaned = String(value || '').replace(/[\\\/:*?"<>|]+/g, '-').replace(/\s+/g, ' ').trim();
  return (cleaned || fallback || 'Report').slice(0, 80);
}

function uniqueExportSheetName_(department, usedNames) {
  const used = usedNames || {};
  const base = String(department || 'Department').replace(/[\[\]\\\/:*?]+/g, '-').replace(/\s+/g, ' ').trim().slice(0, 80) || 'Department';
  let candidate = base;
  let suffix = 2;
  while (used[candidate.toLowerCase()]) {
    const tail = ' (' + suffix + ')';
    candidate = base.slice(0, 100 - tail.length) + tail;
    suffix++;
  }
  used[candidate.toLowerCase()] = true;
  return candidate;
}

function readableExportFontColor_(color) {
  const match = String(color || '').trim().match(/^#?([0-9a-f]{6})$/i);
  if (!match) return '#0F172A';
  const value = match[1];
  const red = parseInt(value.slice(0, 2), 16);
  const green = parseInt(value.slice(2, 4), 16);
  const blue = parseInt(value.slice(4, 6), 16);
  return (red * 299 + green * 587 + blue * 114) / 1000 < 145 ? '#FFFFFF' : '#0F172A';
}

function isTransientSpreadsheetServiceError_(error) {
  const message = String(error && error.message ? error.message : error || '');
  return /Service Spreadsheets|Spreadsheet service|บริการ\s*สเปรดชีต|timed out|หยุดทำงานขณะเข้าถึงเอกสาร|internal error/i.test(message);
}

function exportApprovedSchedule(request, token, retryCount) {
  const operator = requireActiveSession_(token);
  const input = request || {};
  const spreadsheetRetryCount = Math.max(0, Number(retryCount) || 0);
  const month = String(input.month || '').trim().slice(0, 7);
  const format = String(input.format || '').trim().toLowerCase();
  const scope = String(input.scope || 'selected').trim().toLowerCase();
  if (!/^\d{4}-\d{2}$/.test(month)) throw new Error('กรุณาเลือกเดือนที่ต้องการ Export');
  if (['xlsx', 'pdf'].indexOf(format) === -1) throw new Error('รูปแบบไฟล์ต้องเป็น Excel หรือ PDF');
  if (['selected', 'all'].indexOf(scope) === -1) throw new Error('ขอบเขตแผนกไม่ถูกต้อง');

  const spreadsheet = getOrCreateSpreadsheet_();
  ensureRequiredSheets_(spreadsheet);
  ensureScheduleApprovalSchema_(spreadsheet);
  const approval = ensureScheduleApprovalCurrent_(spreadsheet, month, 'SYSTEM DATA CHECK');
  if (String(approval.status || '').toLowerCase() !== 'approved') {
    throw new Error('ตารางกะเดือน ' + buddhistMonthLabel_(month) + ' ยังไม่ได้รับการอนุมัติจาก Admin');
  }
  const authorization = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
  if (authorization.getAuthorizationStatus() === ScriptApp.AuthorizationStatus.REQUIRED) {
    throw new Error('บัญชีผู้ Deploy ยังไม่ได้อนุญาตสิทธิ Export กรุณาเปิด Apps Script ด้วย pttpipeline@gmail.com แล้วเลือก doGet > Run > Review permissions > Allow');
  }

  const timeZone = spreadsheet.getSpreadsheetTimeZone();
  const employeeRows = readObjects_(spreadsheet.getSheetByName('Employees'));
  const employeeById = employeeRows.reduce(function(map, employee) {
    map[String(employee['Employee ID'] || '').trim()] = employee;
    return map;
  }, {});
  const monthSchedule = readSchedule_(spreadsheet, timeZone).filter(function(row) {
    return String(row.date || '').slice(0, 7) === month;
  });
  const availableDepartments = unique_(monthSchedule.map(function(row) {
    return String(row.department || '').trim();
  }).filter(Boolean)).sort();
  const requestedDepartments = normalizedDepartments_(input.departments, availableDepartments);
  const selectedDepartments = scope === 'all' || !requestedDepartments.length ? availableDepartments : requestedDepartments;
  const selectedMap = selectedDepartments.reduce(function(map, department) {
    map[department] = true;
    return map;
  }, {});
  const scheduleRows = monthSchedule.filter(function(row) {
    return selectedMap[String(row.department || '').trim()];
  });
  if (!scheduleRows.length) throw new Error('ไม่พบข้อมูลตารางกะตามเดือนและแผนกที่เลือก');

  const dates = monthDates_(month, timeZone);
  const shiftTypes = readShiftTypes_(spreadsheet, timeZone);
  const shiftTypeByCode = shiftTypes.reduce(function(map, shift) {
    map[String(shift['Shift Code'] || '').trim().toUpperCase()] = shift;
    return map;
  }, {});
  const rowsByDepartment = selectedDepartments.reduce(function(map, department) {
    map[department] = [];
    return map;
  }, {});
  scheduleRows.forEach(function(row) {
    const department = String(row.department || '').trim();
    if (rowsByDepartment[department]) rowsByDepartment[department].push(row);
  });
  const departmentsWithRows = selectedDepartments.filter(function(department) {
    return rowsByDepartment[department] && rowsByDepartment[department].length;
  });
  const reportLayouts = [];
  departmentsWithRows.forEach(function(department) {
    const dateGroups = [dates];
    dateGroups.forEach(function(reportDates) {
      reportLayouts.push({
        department: department,
        dates: reportDates,
        periodLabel: ''
      });
    });
  });

  const exportMoment = new Date();
  const temporaryName = 'Security Management System Export ' + month + ' ' + Utilities.getUuid().slice(0, 8);
  let temporarySpreadsheet = SpreadsheetApp.create(temporaryName);
  const temporaryId = temporarySpreadsheet.getId();
  let outputBlob = null;
  let exportStage = 'เตรียมไฟล์ชั่วคราว';
  try {
    // A newly-created spreadsheet can need a short moment before repeated
    // formatting calls are accepted reliably by the Sheets service.
    SpreadsheetApp.flush();
    Utilities.sleep(400);
    temporarySpreadsheet = SpreadsheetApp.openById(temporaryId);
    temporarySpreadsheet.setSpreadsheetTimeZone(timeZone);
    try { temporarySpreadsheet.setSpreadsheetLocale('th_TH'); } catch (ignore) {}
    const usedSheetNames = {};
    reportLayouts.forEach(function(layout, reportIndex) {
      const department = layout.department;
      const reportDates = layout.dates;
      exportStage = 'สร้างรายงานแผนก ' + department;
      const sheet = reportIndex === 0 ? temporarySpreadsheet.getSheets()[0] : temporarySpreadsheet.insertSheet();
      const sheetLabel = layout.periodLabel ? department + ' ' + layout.periodLabel.replace('วันที่ ', '') : department;
      sheet.setName(uniqueExportSheetName_(sheetLabel, usedSheetNames));
      sheet.setHiddenGridlines(true);

      const peopleById = {};
      rowsByDepartment[department].forEach(function(row) {
        const employeeId = String(row.employeeId || '').trim();
        const employee = employeeById[employeeId] || {};
        if (!peopleById[employeeId]) {
          peopleById[employeeId] = {
            id: employeeId,
            name: String(row.employeeName || employee.Name || ''),
            position: String(employee.Position || ''),
            shifts: {},
            hours: {}
          };
        }
        peopleById[employeeId].shifts[row.date] = String(row.code || '').trim().toUpperCase() || 'OFF';
        peopleById[employeeId].hours[row.date] = Number(row.hours) || 0;
      });
      const people = Object.keys(peopleById).map(function(employeeId) {
        return peopleById[employeeId];
      }).sort(function(first, second) {
        return String(first.name || first.id).localeCompare(String(second.name || second.id), 'th');
      });

      const totalColumns = 4 + reportDates.length;
      const requiredRows = 7 + people.length + shiftTypes.length;
      if (sheet.getMaxColumns() < totalColumns) {
        sheet.insertColumnsAfter(sheet.getMaxColumns(), totalColumns - sheet.getMaxColumns());
      }
      if (sheet.getMaxRows() < requiredRows) {
        sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows());
      }
      const titleRange = sheet.getRange(1, 1, 1, totalColumns).merge();
      titleRange.setValue('Security Management System - ตารางกะที่อนุมัติแล้ว · ' + buddhistMonthLabel_(month) + (layout.periodLabel ? ' · ' + layout.periodLabel : ''));
      titleRange.setBackground('#173566').setFontColor('#FFFFFF').setFontFamily('Sarabun').setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
      sheet.setRowHeight(1, 34);

      const detailRange = sheet.getRange(2, 1, 1, totalColumns).merge();
      detailRange.setValue('แผนก: ' + department + (layout.periodLabel ? '  |  ช่วง: ' + layout.periodLabel : '') + '  |  Revision: ' + approval.revision + '  |  อนุมัติโดย: ' + (approval.approvedBy || '-') + '  |  วันที่อนุมัติ: ' + (buddhistDateTimeText_(approval.approvedAt) || '-'));
      detailRange.setBackground('#E8F1FF').setFontColor('#173566').setFontFamily('Sarabun').setFontSize(9).setFontWeight('bold').setHorizontalAlignment('left').setVerticalAlignment('middle');
      sheet.setRowHeight(2, 25);

      const exportInfoRange = sheet.getRange(3, 1, 1, totalColumns).merge();
      exportInfoRange.setValue('Export โดย: ' + operator.email + '  |  วันที่ Export: ' + (Utilities.formatDate(exportMoment, timeZone, 'dd/MM/') + (Number(Utilities.formatDate(exportMoment, timeZone, 'yyyy')) + 543) + Utilities.formatDate(exportMoment, timeZone, ' HH:mm:ss')));
      exportInfoRange.setBackground('#F8FAFC').setFontColor('#475569').setFontFamily('Sarabun').setFontSize(8).setHorizontalAlignment('left').setVerticalAlignment('middle');

      const thaiDays = ['อา', 'จ', 'อ', 'พ', 'พฤ', 'ศ', 'ส'];
      const headers = ['ลำดับ', 'ชื่อ-นามสกุล', 'ตำแหน่ง'].concat(reportDates.map(function(date) {
        const dateValue = parseIsoDateSafe_(date);
        return String(dateValue.getDate()) + '\n' + thaiDays[dateValue.getDay()];
      })).concat([format === 'pdf' ? 'ชม.รวมเดือน' : 'ชม.รวม']);
      const headerRange = sheet.getRange(4, 1, 1, totalColumns);
      headerRange.setValues([headers]).setFontFamily('Sarabun').setFontSize(format === 'pdf' ? 9 : 8).setFontWeight('bold').setFontColor('#FFFFFF').setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
      const headerBackgrounds = [ '#173566', '#173566', '#173566' ].concat(reportDates.map(function(date) {
        const day = parseIsoDateSafe_(date).getDay();
        return day === 0 ? '#B91C1C' : (day === 6 ? '#1D4ED8' : '#173566');
      })).concat(['#173566']);
      headerRange.setBackgrounds([headerBackgrounds]);
      sheet.setRowHeight(4, 30);

      const outputRows = people.map(function(person, index) {
        const totalHours = dates.reduce(function(sum, date) { return sum + (Number(person.hours[date]) || 0); }, 0);
        return [index + 1, person.name, person.position].concat(reportDates.map(function(date) {
          return person.shifts[date] || 'OFF';
        })).concat([totalHours]);
      });
      if (outputRows.length) {
        const dataRange = sheet.getRange(5, 1, outputRows.length, totalColumns);
        dataRange.setValues(outputRows).setFontFamily('Sarabun').setFontSize(format === 'pdf' ? 10 : 8).setVerticalAlignment('middle');
        sheet.getRange(5, 1, outputRows.length, 1).setHorizontalAlignment('center');
        sheet.getRange(5, 2, outputRows.length, 2).setHorizontalAlignment('left');
        sheet.getRange(5, 4, outputRows.length, reportDates.length).setHorizontalAlignment('center').setFontWeight('bold');
        sheet.getRange(5, totalColumns, outputRows.length, 1).setHorizontalAlignment('right').setNumberFormat('0.0');
        const shiftBackgrounds = outputRows.map(function(row) {
          return row.slice(3, 3 + reportDates.length).map(function(code) {
            const shift = shiftTypeByCode[String(code || '').toUpperCase()] || {};
            return String(shift.Color || (code === '-' ? '#FFFFFF' : '#E2E8F0'));
          });
        });
        const shiftFontColors = shiftBackgrounds.map(function(row) {
          return row.map(readableExportFontColor_);
        });
        sheet.getRange(5, 4, outputRows.length, reportDates.length).setBackgrounds(shiftBackgrounds).setFontColors(shiftFontColors);
        dataRange.setBorder(true, true, true, true, true, true, '#CBD5E1', SpreadsheetApp.BorderStyle.SOLID);
        sheet.setRowHeights(5, outputRows.length, format === 'pdf' ? 26 : 22);
      }

      sheet.setColumnWidth(1, format === 'pdf' ? 30 : 36);
      sheet.setColumnWidth(2, format === 'pdf' ? 140 : 150);
      sheet.setColumnWidth(3, format === 'pdf' ? 78 : 88);
      sheet.setColumnWidths(4, reportDates.length, format === 'pdf' ? 24 : 25);
      sheet.setColumnWidth(totalColumns, 55);
      sheet.setFrozenRows(4);
      // Rows 1-3 are merged across the full report width, so freezing only the
      // first four columns would split those merged cells and make export fail.

      const legendStart = 6 + outputRows.length;
      sheet.getRange(legendStart, 1, 1, 4).merge().setValue('คำอธิบายรหัสกะ').setBackground('#E8F1FF').setFontColor('#173566').setFontFamily('Sarabun').setFontWeight('bold');
      if (shiftTypes.length) {
        const legendRows = shiftTypes.map(function(shift) {
          return [shift['Shift Code'], shift['Shift Name'], shift['Start Time'] + ' - ' + shift['End Time'], Number(shift.Hours) || 0];
        });
        const legendRange = sheet.getRange(legendStart + 1, 1, legendRows.length, 4);
        legendRange.setValues(legendRows).setFontFamily('Sarabun').setFontSize(format === 'pdf' ? 9 : 8).setBorder(true, true, true, true, true, true, '#CBD5E1', SpreadsheetApp.BorderStyle.SOLID);
        legendRange.setBackgrounds(shiftTypes.map(function(shift) {
          const color = String(shift.Color || '#E2E8F0');
          return [color, '#FFFFFF', '#FFFFFF', '#FFFFFF'];
        }));
        legendRange.setFontColors(shiftTypes.map(function(shift) {
          return [readableExportFontColor_(shift.Color), '#0F172A', '#0F172A', '#0F172A'];
        }));
      }
      // Add signature box at bottom right (65% width) with margin-bottom 28px formatting for PDF export
      if (format === 'pdf') {
        const sigStartRow = legendStart + (shiftTypes.length ? shiftTypes.length + 3 : 4);
        const startCol = Math.max(1, Math.floor(totalColumns * 0.35) + 1);
        const numCols = totalColumns - startCol + 1;
        
        // Box 1: พนักงานผู้จัดพิมพ์รายงาน/หัวหน้าพนักงานรักษาความปลอดภัย (Export PDF)
        const box1Line1 = sheet.getRange(sigStartRow, startCol, 1, numCols).merge();
        box1Line1.setValue('ลงชื่อ........................................................................................................').setFontFamily('Sarabun').setFontSize(10).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
        sheet.setRowHeight(sigStartRow, 34);
        
        const box1Line2 = sheet.getRange(sigStartRow + 1, startCol, 1, numCols).merge();
        box1Line2.setValue('(........................................................................................................)').setFontFamily('Sarabun').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');
        sheet.setRowHeight(sigStartRow + 1, 28);
        
        const box1Line3 = sheet.getRange(sigStartRow + 2, startCol, 1, numCols).merge();
        box1Line3.setValue('พนักงานผู้จัดพิมพ์รายงาน / หัวหน้าพนักงานรักษาความปลอดภัย').setFontFamily('Sarabun').setFontSize(10).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('top');
        sheet.setRowHeight(sigStartRow + 2, 38); // extra height to simulate margin-bottom: 28px
        
        // Box 2: ผู้จัดการเขต (ผู้อนุมัติ)
        const box2Line1 = sheet.getRange(sigStartRow + 4, startCol, 1, numCols).merge();
        box2Line1.setValue('ทราบ / ลงชื่อ........................................................................................................').setFontFamily('Sarabun').setFontSize(10).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
        sheet.setRowHeight(sigStartRow + 4, 34);
        
        const box2Line2 = sheet.getRange(sigStartRow + 5, startCol, 1, numCols).merge();
        box2Line2.setValue('(........................................................................................................)').setFontFamily('Sarabun').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');
        sheet.setRowHeight(sigStartRow + 5, 28);
        
        const box2Line3 = sheet.getRange(sigStartRow + 6, startCol, 1, numCols).merge();
        box2Line3.setValue('ผู้จัดการเขต (ผู้อนุมัติ)').setFontFamily('Sarabun').setFontSize(10).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('top');
        sheet.setRowHeight(sigStartRow + 6, 38); // extra height to simulate margin-bottom: 28px
      }
      // Flush in small groups so exporting all departments does not leave one
      // very large queue of pending Spreadsheet service operations.
      if ((reportIndex + 1) % 3 === 0) {
        SpreadsheetApp.flush();
        Utilities.sleep(150);
      }
    });

    exportStage = 'ประมวลผลรูปแบบรายงาน';
    SpreadsheetApp.flush();
    Utilities.sleep(300);
    exportStage = 'สร้างไฟล์ ' + format.toUpperCase();
    const exportUrl = 'https://docs.google.com/spreadsheets/d/' + temporaryId + '/export?' + (format === 'pdf'
      ? 'format=pdf&size=A4&portrait=false&fitw=true&sheetnames=false&printtitle=false&pagenumbers=true&gridlines=false&fzr=true&top_margin=0.25&bottom_margin=0.25&left_margin=0.3&right_margin=0.3'
      : 'format=xlsx');
    const response = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    if (response.getResponseCode() !== 200) {
      throw new Error('Google ไม่สามารถสร้างไฟล์ ' + format.toUpperCase() + ' ได้ (รหัส ' + response.getResponseCode() + ')');
    }
    outputBlob = response.getBlob();

    const latestApproval = scheduleApprovalState_(spreadsheet, month);
    if (String(latestApproval.status || '').toLowerCase() !== 'approved' || Number(latestApproval.revision) !== Number(approval.revision)) {
      throw new Error('ตารางมีการเปลี่ยนแปลงระหว่างสร้างไฟล์ กรุณาให้ Admin อนุมัติใหม่ก่อน Export');
    }
    const scopeLabel = scope === 'all' ? 'ทุกแผนก' : departmentsWithRows.join(', ');
    const fileScope = scope === 'all' ? 'ทุกแผนก' : (departmentsWithRows.length === 1 ? departmentsWithRows[0] : departmentsWithRows.length + '-แผนก');
    const extension = format === 'pdf' ? 'pdf' : 'xlsx';
    const mimeType = format === 'pdf' ? 'application/pdf' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    const fileName = safeExportFilePart_('Security Management System-ตารางกะ-' + (Number(month.slice(0, 4)) + 543) + '-' + month.slice(5, 7) + '-' + fileScope, 'Security-Management-System-Schedule') + '.' + extension;
    const bytes = outputBlob.getBytes();
    appendScheduleApprovalLog_(spreadsheet, {
      Timestamp: new Date(),
      Action: 'EXPORTED_' + format.toUpperCase(),
      Month: month,
      Revision: approval.revision,
      Status: 'Approved',
      'Change Type': 'EXPORT',
      'Performed By': operator.email,
      Note: scopeLabel
    });
    SpreadsheetApp.flush();
    return {
      fileName: fileName,
      mimeType: mimeType,
      base64: Utilities.base64Encode(bytes),
      size: bytes.length,
      format: format,
      month: month,
      departments: departmentsWithRows,
      revision: approval.revision
    };
  } catch (error) {
    if (spreadsheetRetryCount < 1 && isTransientSpreadsheetServiceError_(error)) {
      try { Drive.Files.update({ trashed: true }, temporaryId); } catch (ignore) {}
      Utilities.sleep(1200);
      return exportApprovedSchedule(request, token, spreadsheetRetryCount + 1);
    }
    const message = String(error && error.message ? error.message : error || 'Unknown error');
    throw new Error('Export ไม่สำเร็จในขั้นตอน "' + exportStage + '": ' + message);
  } finally {
    try {
      Drive.Files.update({ trashed: true }, temporaryId);
    } catch (ignore) {}
  }
}

/**
 * One-time editor-only diagnostic. The trailing underscore prevents calls from
 * google.script.run while still allowing the deployer to run it in the editor.
 */
function authorizeExportServices_() {
  let temporaryId = '';
  let stage = 'สร้าง Spreadsheet ทดสอบ';
  try {
    const testSpreadsheet = SpreadsheetApp.create('Security Management System Export Permission Test ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss'), 10, 10);
    temporaryId = testSpreadsheet.getId();
    testSpreadsheet.getSheets()[0].getRange('A1:B3').setValues([
      ['Security Management System', 'Export Permission Test'],
      ['Excel', 'OK'],
      ['PDF', 'A4 Landscape']
    ]);
    SpreadsheetApp.flush();

    const authorizationHeader = { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() };
    stage = 'สร้างไฟล์ Excel';
    const excelResponse = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + temporaryId + '/export?format=xlsx', {
      headers: authorizationHeader,
      muteHttpExceptions: true
    });
    if (excelResponse.getResponseCode() !== 200) {
      throw new Error('Excel HTTP ' + excelResponse.getResponseCode() + ': ' + excelResponse.getContentText().slice(0, 200));
    }

    stage = 'สร้างไฟล์ PDF';
    const pdfResponse = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + temporaryId + '/export?format=pdf&size=A4&portrait=false&fitw=true&gridlines=false', {
      headers: authorizationHeader,
      muteHttpExceptions: true
    });
    if (pdfResponse.getResponseCode() !== 200) {
      throw new Error('PDF HTTP ' + pdfResponse.getResponseCode() + ': ' + pdfResponse.getContentText().slice(0, 200));
    }

    console.log('SHIFTFLOW_EXPORT_TEST_OK Excel=' + excelResponse.getBlob().getBytes().length + ' bytes PDF=' + pdfResponse.getBlob().getBytes().length + ' bytes');
    return 'SHIFTFLOW_EXPORT_TEST_OK';
  } catch (error) {
    const message = 'SHIFTFLOW_EXPORT_TEST_FAILED [' + stage + '] ' + (error && error.message ? error.message : error);
    console.error(message);
    throw new Error(message);
  } finally {
    if (temporaryId) {
      try {
        Drive.Files.update({ trashed: true }, temporaryId);
        console.log('SHIFTFLOW_EXPORT_TEST_CLEANUP OK');
      } catch (cleanupError) {
        console.error('SHIFTFLOW_EXPORT_TEST_CLEANUP_FAILED ' + cleanupError.message);
      }
    }
  }
}

function ensureLicenseSchema_(spreadsheet) {
  const schemas = {
    'Employee Licenses': [
      'License ID', 'Employee ID', 'License Type', 'License Number',
      'Issue Date', 'Expiry Date', 'Status', 'Document URL', 'Remark',
      'Updated By', 'Updated At'
    ],
    'License Audit Log': [
      'Timestamp', 'Action', 'Employee ID', 'License ID', 'Work Date',
      'Shift Code', 'License Status', 'Expiry Date', 'Reason', 'Approved By'
    ]
  };
  Object.keys(schemas).forEach(function(sheetName) {
    const sheet = spreadsheet.getSheetByName(sheetName);
    const required = schemas[sheetName];
    const lastColumn = Math.max(sheet.getLastColumn(), 1);
    const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    if (!headers.some(function(header) { return String(header || '').trim(); })) {
      sheet.getRange(1, 1, 1, required.length).setValues([required]);
      sheet.setFrozenRows(1);
      return;
    }
    required.forEach(function(header) {
      if (headers.indexOf(header) !== -1) return;
      headers.push(header);
      sheet.getRange(1, headers.length).setValue(header);
    });
    sheet.setFrozenRows(1);
  });
}

function scheduleLocked_(value) {
  if (value === true) return true;
  return ['true', '1', 'yes', 'locked'].indexOf(String(value || '').trim().toLowerCase()) !== -1;
}

function booleanValue_(value) {
  if (value === true) return true;
  return ['true', '1', 'yes', 'y'].indexOf(String(value || '').trim().toLowerCase()) !== -1;
}

function readEmployeeLicenses_(spreadsheet, timeZone) {
  ensureLicenseSchema_(spreadsheet);
  const sheet = spreadsheet.getSheetByName('Employee Licenses');
  const rows = readObjects_(sheet);
  let idCol = -1;

  if (sheet && sheet.getLastColumn() >= 1) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h) { return String(h || '').trim(); });
    idCol = headers.indexOf('License ID') + 1;
  }

  return rows.map(function(row, index) {
    let licId = String(row['License ID'] || row['รหัสใบอนุญาต'] || '').trim();
    const empId = String(row['Employee ID'] || row['รหัสพนักงาน'] || row['รหัสประจำตัว'] || '').trim();
    if (!licId && empId) {
      licId = 'LIC-' + empId + '-' + (index + 1);
      if (sheet && row._rowIndex && idCol > 0) {
        try { sheet.getRange(row._rowIndex, idCol).setValue(licId); } catch (_) {}
      }
    }
    return {
      licenseId: licId,
      employeeId: empId,
      licenseType: String(row['License Type'] || row['ประเภทใบอนุญาต'] || row['ประเภท'] || 'ใบอนุญาตทั่วไป').trim(),
      licenseNumber: String(row['License Number'] || row['เลขที่ใบอนุญาต'] || row['เลขที่'] || '-').trim(),
      issueDate: formatDate_(row['Issue Date'] || row['วันที่ออก'], timeZone),
      expiryDate: formatDate_(row['Expiry Date'] || row['วันหมดอายุ'], timeZone),
      status: String(row.Status || row['สถานะ'] || 'Active').trim() || 'Active',
      documentUrl: String(row['Document URL'] || row['ลิงก์เอกสาร'] || '').trim(),
      remark: String(row.Remark || row['หมายเหตุ'] || '').trim(),
      updatedBy: String(row['Updated By'] || row['ผู้บันทึก'] || '').trim(),
      updatedAt: row['Updated At'] ? formatDateTime_(row['Updated At'], timeZone) : ''
    };
  }).filter(function(license) {
    return license.employeeId || license.licenseNumber || license.licenseType;
  });
}

function buildLicenseIndex_(licenses) {
  return (licenses || []).reduce(function(index, license) {
    if (!index[license.employeeId]) index[license.employeeId] = [];
    index[license.employeeId].push(license);
    return index;
  }, {});
}

function licenseStatusForDate_(employeeId, date, licenseIndex) {
  const records = (licenseIndex[employeeId] || []).slice();
  if (!records.length) {
    return { valid: false, code: 'MISSING', reason: 'ไม่พบข้อมูลใบอนุญาต', expiryDate: '', licenseId: '', licenseType: '' };
  }

  const active = records.filter(function(license) {
    return String(license.status || '').trim().toLowerCase() === 'active';
  });
  const valid = active.filter(function(license) {
    return license.issueDate && license.expiryDate && license.issueDate <= date && date <= license.expiryDate;
  }).sort(function(a, b) {
    return String(b.expiryDate).localeCompare(String(a.expiryDate));
  });
  if (valid.length) {
    const license = valid[0];
    return {
      valid: true,
      code: 'VALID',
      reason: 'ใบอนุญาตยังมีผล',
      expiryDate: license.expiryDate,
      issueDate: license.issueDate,
      licenseId: license.licenseId,
      licenseType: license.licenseType,
      licenseNumber: license.licenseNumber
    };
  }

  const future = active.filter(function(license) {
    return license.issueDate && license.issueDate > date;
  }).sort(function(a, b) { return a.issueDate.localeCompare(b.issueDate); });
  if (future.length) {
    return {
      valid: false,
      code: 'NOT_YET_VALID',
      reason: 'ใบอนุญาตยังไม่ถึงวันเริ่มใช้งาน',
      expiryDate: future[0].expiryDate,
      issueDate: future[0].issueDate,
      licenseId: future[0].licenseId,
      licenseType: future[0].licenseType
    };
  }

  const expired = active.filter(function(license) {
    return license.expiryDate && license.expiryDate < date;
  }).sort(function(a, b) { return b.expiryDate.localeCompare(a.expiryDate); });
  if (expired.length) {
    return {
      valid: false,
      code: 'EXPIRED',
      reason: 'ใบอนุญาตหมดอายุแล้ว',
      expiryDate: expired[0].expiryDate,
      issueDate: expired[0].issueDate,
      licenseId: expired[0].licenseId,
      licenseType: expired[0].licenseType
    };
  }

  if (active.length) {
    return {
      valid: false,
      code: 'INVALID_DATES',
      reason: 'ข้อมูลวันที่ใบอนุญาตไม่ครบหรือไม่ถูกต้อง',
      expiryDate: active[0].expiryDate || '',
      licenseId: active[0].licenseId,
      licenseType: active[0].licenseType
    };
  }

  const blocked = records[0];
  const status = String(blocked.status || 'Inactive').trim().toUpperCase();
  return {
    valid: false,
    code: status || 'INACTIVE',
    reason: 'สถานะใบอนุญาต: ' + (blocked.status || 'Inactive'),
    expiryDate: blocked.expiryDate || '',
    licenseId: blocked.licenseId,
    licenseType: blocked.licenseType
  };
}

function isWorkingShift_(shift) {
  return Number(shift && (shift.Hours !== undefined ? shift.Hours : shift.hours)) > 0;
}

function licenseScheduleViolations_(schedule, licenseIndex) {
  const grouped = (schedule || []).reduce(function(map, row) {
    if (Number(row.hours) <= 0 || row.licenseOverride) return map;
    const status = licenseStatusForDate_(row.employeeId, row.date, licenseIndex);
    if (status.valid) return map;
    if (!map[row.employeeId]) map[row.employeeId] = { row: row, status: status, count: 0 };
    map[row.employeeId].count++;
    return map;
  }, {});
  return Object.keys(grouped).map(function(employeeId) {
    const item = grouped[employeeId];
    const row = item.row;
    const status = item.status;
    return {
      ruleId: 'LICENSE',
      ruleName: 'ใบอนุญาตต้องมีผลในวันที่ทำงาน',
      title: row.employeeName + ' · พบกะทำงานที่ใบอนุญาตไม่ผ่าน',
      description: item.count + ' กะ · เริ่มพบ ' + row.date + ' · ' + status.reason + (status.expiryDate ? ' · หมดอายุ ' + status.expiryDate : ''),
      severity: 'error'
    };
  });
}

function licenseAlerts_(employees, licenseIndex, today) {
  const DAY_MS = 24 * 60 * 60 * 1000;
  return (employees || []).filter(function(employee) {
    return String(employee.Status || '').trim().toLowerCase() === 'active';
  }).map(function(employee) {
    const employeeId = String(employee['Employee ID'] || '').trim();
    const status = licenseStatusForDate_(employeeId, today, licenseIndex);
    if (!status.valid) {
      return {
        employeeId: employeeId,
        code: status.code,
        expiryDate: status.expiryDate,
        ruleId: 'LICENSE',
        ruleName: 'สถานะใบอนุญาตพนักงาน',
        title: String(employee.Name || employeeId) + ' · ' + status.reason,
        description: status.expiryDate ? 'วันหมดอายุ ' + status.expiryDate : 'กรุณาตรวจสอบข้อมูลในชีต Employee Licenses',
        severity: 'error'
      };
    }
    const days = Math.floor((parseIsoDateSafe_(status.expiryDate).getTime() - parseIsoDateSafe_(today).getTime()) / DAY_MS);
    if (days > 90) return null;
    return {
      employeeId: employeeId,
      code: 'EXPIRING',
      expiryDate: status.expiryDate,
      daysRemaining: days,
      ruleId: 'LICENSE',
      ruleName: 'สถานะใบอนุญาตพนักงาน',
      title: String(employee.Name || employeeId) + ' · ใบอนุญาตใกล้หมดอายุ',
      description: 'เหลือ ' + days + ' วัน · หมดอายุ ' + status.expiryDate,
      severity: days <= 30 ? 'warning' : 'info'
    };
  }).filter(Boolean);
}

function appendLicenseAudit_(spreadsheet, record) {
  const sheet = spreadsheet.getSheetByName('License Audit Log');
  if (!sheet) return;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  sheet.appendRow(headers.map(function(header) {
    return Object.prototype.hasOwnProperty.call(record, header) ? record[header] : '';
  }));
}

function parseIsoDateSafe_(value) {
  const match = String(value || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) throw new Error('Invalid start date. Use YYYY-MM-DD.');
  return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]), 12, 0, 0);
}

function isoDate_(date, timeZone) {
  return Utilities.formatDate(date, timeZone, 'yyyy-MM-dd');
}

function monthDates_(monthValue, timeZone) {
  const match = String(monthValue || '').trim().match(/^(\d{4})-(\d{2})(?:-\d{2})?$/);
  if (!match) throw new Error('Invalid month. Use YYYY-MM.');
  const year = Number(match[1]);
  const month = Number(match[2]);
  if (month < 1 || month > 12) throw new Error('Month must be between 01 and 12.');

  const daysInMonth = new Date(year, month, 0, 12, 0, 0).getDate();
  if (daysInMonth > 31) throw new Error('A schedule can cover no more than one month.');
  return Array.from({ length: daysInMonth }, function(_, index) {
    return isoDate_(new Date(year, month - 1, index + 1, 12, 0, 0), timeZone);
  });
}

function autoPlanRow_(date, employee, shift, source, locked, remark, metadata) {
  const meta = metadata || {};
  const license = meta.license || {};
  return {
    date: date,
    employeeId: String(employee['Employee ID'] || employee.id || '').trim(),
    employeeName: String(employee.Name || employee.name || ''),
    department: String(employee.Department || employee.department || ''),
    code: String(shift['Shift Code'] || shift.code || 'OFF').toUpperCase(),
    startTime: String(shift['Start Time'] || shift.startTime || ''),
    endTime: String(shift['End Time'] || shift.endTime || ''),
    hours: Number(shift.Hours || shift.hours) || 0,
    remark: String(remark || ''),
    source: String(source || 'Auto'),
    locked: Boolean(locked),
    licenseStatus: String(license.code || meta.licenseStatus || ''),
    licenseExpiryDate: String(license.expiryDate || meta.licenseExpiryDate || ''),
    licenseOverride: Boolean(meta.licenseOverride),
    overrideReason: String(meta.overrideReason || ''),
    overrideBy: String(meta.overrideBy || ''),
    overrideAt: String(meta.overrideAt || '')
  };
}

function buildAutoSchedulePlan_(startDate, adminUser) {
  const spreadsheet = getOrCreateSpreadsheet_();
  ensureRequiredSheets_(spreadsheet);
  ensureScheduleMetadataColumns_(spreadsheet.getSheetByName('Schedule'));
  enforceMaxWeeklyHours72_(spreadsheet.getSheetByName('Rules'));

  const timeZone = spreadsheet.getSpreadsheetTimeZone();
  const dates = monthDates_(startDate, timeZone);
  const dateSet = dates.reduce(function(map, date) { map[date] = true; return map; }, {});
  
  const rules = parseRules_(readObjects_(spreadsheet.getSheetByName('Rules')));
  const maxHours = Number(ruleValue_(rules, 'RULE001', DEFAULT_MAX_WEEKLY_HOURS)) || DEFAULT_MAX_WEEKLY_HOURS;
  const dayMinimum = Math.max(0, Number(ruleValue_(rules, 'RULE003', 0)) || 0);
  const nightMinimum = Math.max(0, Number(ruleValue_(rules, 'RULE004', 0)) || 0);
  const excludeSpare = ruleValue_(rules, 'RULE009', 0) === 1 || String(ruleValue_(rules, 'RULE009', '')).toLowerCase() === 'true';

  const employees = readObjects_(spreadsheet.getSheetByName('Employees')).filter(function(employee) {
    if (String(employee.Status || '').trim().toLowerCase() !== 'active') return false;
    if (excludeSpare && String(employee.Department || '').trim().toLowerCase() === 'spare') return false;
    return true;
  });
  if (!employees.length) throw new Error('No active eligible employees were found.');
  const licenseIndex = buildLicenseIndex_(readEmployeeLicenses_(spreadsheet, timeZone));

  const shiftTypes = readShiftTypes_(spreadsheet, timeZone);
  const shiftMap = indexBy_(shiftTypes, 'Shift Code');
  ['D', 'N', 'OFF'].forEach(function(code) {
    if (!shiftMap[code]) throw new Error('Shift Types must contain ' + code + '.');
  });

  const schedule = readSchedule_(spreadsheet, timeZone);

  const employeeById = employees.reduce(function(map, employee) {
    map[String(employee['Employee ID'] || '').trim()] = employee;
    return map;
  }, {});
  const weeklyHours = {};
  function weeklyHoursFor(employeeId, date) {
    return weeklyHours[employeeId + '|' + isoWeek_(date)] || 0;
  }
  const existingByKey = {};
  schedule.forEach(function(row) {
    if (!dateSet[row.date] || !employeeById[row.employeeId]) return;
    if (row.locked || row.code === 'AL') existingByKey[row.employeeId + '|' + row.date] = row;
  });

  const planByKey = {};
  const warnings = [];
  const licenseBlocked = {};
  function assign(employee, date, shift, source, locked, remark, metadata) {
    const id = String(employee['Employee ID'] || '').trim();
    const key = id + '|' + date;
    if (planByKey[key]) return false;
    const hours = Number(shift.Hours || shift.hours) || 0;
    const license = licenseStatusForDate_(id, date, licenseIndex);
    const meta = metadata || {};
    const allowedOverride = Boolean(locked && String(source || '').toLowerCase() === 'manual' && meta.licenseOverride);
    if (hours > 0 && !license.valid && !allowedOverride) {
      licenseBlocked[key] = license;
      return false;
    }
    const weekKey = id + '|' + isoWeek_(date);
    if (!locked && hours > 0 && weeklyHoursFor(id, date) + hours > maxHours) return false;
    meta.license = license;
    planByKey[key] = autoPlanRow_(date, employee, shift, source, locked, remark, meta);
    weeklyHours[weekKey] = weeklyHoursFor(id, date) + hours;
    return true;
  }

  employees.forEach(function(employee) {
    dates.forEach(function(date) {
      const id = String(employee['Employee ID'] || '').trim();
      const existing = existingByKey[id + '|' + date];
      if (!existing) return;
      const shift = shiftMap[existing.code] || {
        'Shift Code': existing.code,
        'Start Time': existing.startTime,
        'End Time': existing.endTime,
        Hours: existing.hours
      };
      assign(employee, date, shift, existing.source || 'Manual', true, existing.remark, {
        licenseOverride: existing.licenseOverride,
        overrideReason: existing.overrideReason,
        overrideBy: existing.overrideBy,
        overrideAt: existing.overrideAt,
        licenseStatus: existing.licenseStatus,
        licenseExpiryDate: existing.licenseExpiryDate
      });
    });
  });

  const rule006Target = Number(ruleValue_(rules, 'RULE006', 0)) || 0;
  const rule007Target = Number(ruleValue_(rules, 'RULE007', 0)) || 0;
  const rule008Target = Number(ruleValue_(rules, 'RULE008', 0)) || 0;

  const isSupervisor = function(emp) { return String(emp.Position || '').trim().toLowerCase() === 'supervisor'; };
  const isLeader = function(emp) { 
    const p = String(emp.Position || '').trim().toLowerCase();
    return p === 'supervisor' || p === 'team leader' || p === 'act.team leader';
  };
  const allDepts = unique_(employees.map(function(emp) { return String(emp.Department || ''); }).filter(Boolean));

  employees.forEach(function(employee) {
    if (!isSupervisor(employee)) return;
    dates.forEach(function(date) {
      const id = String(employee['Employee ID'] || '').trim();
      const key = id + '|' + date;
      if (planByKey[key]) return;
      if (parseIsoDateSafe_(date).getDay() === 0) {
        assign(employee, date, shiftMap.OFF, 'Auto', false, 'Supervisor Sunday off');
      } else if (rule006Target > 0) {
        assign(employee, date, shiftMap.D, 'Auto', false, 'Supervisor Day only');
      }
    });
  });

  dates.forEach(function(date) {
    ['PO11', 'WCS'].forEach(function(dept) {
      const deptEmployees = employees.filter(function(emp) { return String(emp.Department || '') === dept; });
      if (rule007Target > 0) {
        let dLeaders = deptEmployees.filter(function(emp) {
          const id = String(emp['Employee ID'] || '').trim();
          return isLeader(emp) && planByKey[id + '|' + date] && planByKey[id + '|' + date].code === 'D';
        }).length;
        const dCandidates = deptEmployees.filter(function(emp) {
          const id = String(emp['Employee ID'] || '').trim();
          return isLeader(emp) && !planByKey[id + '|' + date];
        }).sort(function(a, b) { return weeklyHoursFor(String(a['Employee ID']).trim(), date) - weeklyHoursFor(String(b['Employee ID']).trim(), date); });
        for (let i = 0; i < dCandidates.length && dLeaders < rule007Target; i++) {
          if (assign(dCandidates[i], date, shiftMap.D, 'Auto', false, 'Auto Leader D coverage')) dLeaders++;
        }
      }
      if (rule008Target > 0) {
        let nLeaders = deptEmployees.filter(function(emp) {
          const id = String(emp['Employee ID'] || '').trim();
          return isLeader(emp) && planByKey[id + '|' + date] && planByKey[id + '|' + date].code === 'N';
        }).length;
        const nCandidates = deptEmployees.filter(function(emp) {
          const id = String(emp['Employee ID'] || '').trim();
          return isLeader(emp) && !planByKey[id + '|' + date] && !isSupervisor(emp);
        }).sort(function(a, b) { return weeklyHoursFor(String(a['Employee ID']).trim(), date) - weeklyHoursFor(String(b['Employee ID']).trim(), date); });
        for (let i = 0; i < nCandidates.length && nLeaders < rule008Target; i++) {
          if (assign(nCandidates[i], date, shiftMap.N, 'Auto', false, 'Auto Leader N coverage')) nLeaders++;
        }
      }
    });
  });

  dates.forEach(function(date, dayIndex) {
    const coverageOrder = dayIndex % 2 ? [
      { code: 'N', minimum: nightMinimum },
      { code: 'D', minimum: dayMinimum }
    ] : [
      { code: 'D', minimum: dayMinimum },
      { code: 'N', minimum: nightMinimum }
    ];

    allDepts.forEach(function(dept) {
      coverageOrder.forEach(function(target) {
        if (target.minimum <= 0) return;
        let assignedCount = employees.filter(function(employee) {
          const id = String(employee['Employee ID'] || '').trim();
          const row = planByKey[id + '|' + date];
          return String(employee.Department || '') === dept && row && row.code === target.code;
        }).length;
        const candidates = employees.filter(function(employee) {
          const id = String(employee['Employee ID'] || '').trim();
          return String(employee.Department || '') === dept && !planByKey[id + '|' + date] && !(target.code === 'N' && isSupervisor(employee));
        }).sort(function(a, b) {
          const aId = String(a['Employee ID'] || '').trim();
          const bId = String(b['Employee ID'] || '').trim();
          return weeklyHoursFor(aId, date) - weeklyHoursFor(bId, date) || aId.localeCompare(bId);
        });

        for (let i = 0; i < candidates.length && assignedCount < target.minimum; i++) {
          if (assign(candidates[i], date, shiftMap[target.code], 'Auto', false, 'Auto coverage')) assignedCount++;
        }
        if (assignedCount < target.minimum) {
          warnings.push(date + ': ' + dept + ' ' + target.code + ' has ' + assignedCount +
            '/' + target.minimum + ' people. Add staff or lower the coverage rule.');
        }
      });
    });
  });

  employees.forEach(function(employee, employeeIndex) {
    const id = String(employee['Employee ID'] || '').trim();
    dates.forEach(function(date, dayIndex) {
      const key = id + '|' + date;
      if (planByKey[key]) return;
      const license = licenseStatusForDate_(id, date, licenseIndex);
      if (!license.valid) {
        licenseBlocked[key] = license;
        assign(employee, date, shiftMap.OFF, 'Auto', false, 'License blocked: ' + license.reason);
        return;
      }
      const cycleDay = (employeeIndex + dayIndex) % 7;
      if (cycleDay >= 5) {
        assign(employee, date, shiftMap.OFF, 'Auto', false, 'Auto rest day');
        return;
      }
      let preferredCode = (employeeIndex + Math.floor(dayIndex / 3)) % 2 === 0 ? 'D' : 'N';
      if (isSupervisor(employee) && preferredCode === 'N') preferredCode = 'D';
      if (!assign(employee, date, shiftMap[preferredCode], 'Auto', false, 'Auto rotation')) {
        assign(employee, date, shiftMap.OFF, 'Auto', false, 'Weekly hour limit');
      }
    });
  });

  const rows = Object.keys(planByKey).map(function(key) { return planByKey[key]; }).sort(function(a, b) {
    return a.date.localeCompare(b.date) || a.employeeId.localeCompare(b.employeeId);
  });
  const counts = rows.reduce(function(map, row) {
    map[row.code] = (map[row.code] || 0) + 1;
    return map;
  }, {});
  const blockedByEmployee = Object.keys(licenseBlocked).reduce(function(map, key) {
    const employeeId = key.split('|')[0];
    map[employeeId] = (map[employeeId] || 0) + 1;
    return map;
  }, {});
  Object.keys(blockedByEmployee).forEach(function(employeeId) {
    const employee = employeeById[employeeId] || {};
    warnings.push((employee.Name || employeeId) + ': ถูกป้องกันกะทำงาน ' + blockedByEmployee[employeeId] + ' วัน เนื่องจากใบอนุญาตไม่ผ่าน');
  });
  return {
    month: dates[0].slice(0, 7),
    startDate: dates[0],
    endDate: dates[dates.length - 1],
    dates: dates,
    rows: rows,
    warnings: warnings,
    summary: {
      employees: employees.length,
      days: dates.length,
      totalRows: rows.length,
      manualLocked: rows.filter(function(row) { return row.locked; }).length,
      counts: counts,
      maxWeeklyHours: maxHours,
      dayMinimum: dayMinimum,
      nightMinimum: nightMinimum
    },
    licenseBlocked: blockedByEmployee,
    generatedBy: adminUser.email
  };
}

function previewAutoSchedule(startDate, token) {
  const operator = requireManagerOrAdmin_(token);
  return buildAutoSchedulePlan_(startDate, operator);
}

function backupSchedule_(spreadsheet) {
  const source = spreadsheet.getSheetByName('Schedule');
  let name = 'Schedule Backup ' + Utilities.formatDate(new Date(), spreadsheet.getSpreadsheetTimeZone(), 'yyyyMMdd HHmmss');
  let suffix = 1;
  while (spreadsheet.getSheetByName(name)) name = name.slice(0, 92) + ' ' + suffix++;
  const backup = source.copyTo(spreadsheet).setName(name);
  return backup.getName();
}

function commitAutoSchedule(startDate, token) {
  const operator = requireManagerOrAdmin_(token);
  const plan = buildAutoSchedulePlan_(startDate, operator);
  const spreadsheet = getOrCreateSpreadsheet_();
  backupSchedule_(spreadsheet);
  const sheet = spreadsheet.getSheetByName('Schedule');
  const headers = ensureScheduleMetadataColumns_(sheet);
  const timeZone = spreadsheet.getSpreadsheetTimeZone();
  const dateSet = plan.dates.reduce(function(map, date) { map[date] = true; return map; }, {});
  const allowedLockedKeys = plan.rows.filter(function(row) { return row.locked; }).reduce(function(map, row) {
    map[row.employeeId + '|' + row.date] = true;
    return map;
  }, {});

  const rawValues = sheet.getDataRange().getValues();
  rawValues.shift();
  const dateColumn = headers.indexOf('Date');
  const employeeColumn = headers.indexOf('Employee ID');
  const codeColumn = headers.indexOf('Shift Code');
  const lockedColumn = headers.indexOf('Locked');
  const keptRows = rawValues.filter(function(row) {
    if (!row.some(function(cell) { return cell !== '' && cell !== null; })) return false;
    const date = formatDate_(row[dateColumn], timeZone);
    if (!dateSet[date]) return true;
    const code = String(row[codeColumn] || '').trim().toUpperCase();
    const employeeId = String(row[employeeColumn] || '').trim();
    return code === 'AL' || (scheduleLocked_(row[lockedColumn]) && allowedLockedKeys[employeeId + '|' + date]);
  }).map(function(row) {
    return headers.map(function(_, index) { return row[index] === undefined ? '' : row[index]; });
  });

  const generatedRows = plan.rows.filter(function(row) { return !row.locked; }).map(function(row) {
    const record = {
      Date: parseIsoDateSafe_(row.date),
      'Employee ID': row.employeeId,
      'Employee Name': row.employeeName,
      Department: row.department,
      'Shift Code': row.code,
      'Start Time': row.startTime,
      'End Time': row.endTime,
      Hours: row.hours,
      Remark: row.remark,
      Source: 'Auto',
      Locked: false,
      'Updated By': operator.email,
      'Updated At': new Date(),
      'License Status': row.licenseStatus,
      'License Expiry Date': row.licenseExpiryDate ? parseIsoDateSafe_(row.licenseExpiryDate) : '',
      'License Override': false,
      'Override Reason': '',
      'Override By': '',
      'Override At': ''
    };
    return headers.map(function(header) {
      return Object.prototype.hasOwnProperty.call(record, header) ? record[header] : '';
    });
  });

  if (sheet.getMaxRows() > 1) sheet.getRange(2, 1, sheet.getMaxRows() - 1, headers.length).clearContent();
  const outputRows = keptRows.concat(generatedRows);
  if (outputRows.length) sheet.getRange(2, 1, outputRows.length, headers.length).setValues(outputRows);
  sheet.getRange(2, 1, Math.max(outputRows.length, 1), 1).setNumberFormat('yyyy-mm-dd');
  SpreadsheetApp.flush();
  const approval = invalidateScheduleApprovals_(spreadsheet, [plan.startDate.slice(0, 7)], operator.email, 'AUTO_SCHEDULE');
  bumpDataVersion_();

  return {
    writtenRows: generatedRows.length,
    preservedRows: keptRows.filter(function(row) {
      return dateSet[formatDate_(row[dateColumn], timeZone)];
    }).length,
    warnings: plan.warnings,
    startDate: plan.startDate,
    endDate: plan.endDate,
    approval: approval[0] || null
  };
}

function updateEmployeeShifts(changes, token) {
  const operator = requireManagerOrAdmin_(token);
  if (!Array.isArray(changes) || !changes.length) throw new Error('No schedule changes were submitted.');
  if (changes.length > 500) throw new Error('A maximum of 500 changes can be saved at once.');

  const spreadsheet = getOrCreateSpreadsheet_();
  ensureRequiredSheets_(spreadsheet);
  enforceMaxWeeklyHours72_(spreadsheet.getSheetByName('Rules'));
  const sheet = spreadsheet.getSheetByName('Schedule');
  const headers = ensureScheduleMetadataColumns_(sheet);
  const timeZone = spreadsheet.getSpreadsheetTimeZone();

  const employeeMap = readObjects_(spreadsheet.getSheetByName('Employees')).reduce(function(map, employee) {
    map[String(employee['Employee ID'] || '').trim()] = employee;
    return map;
  }, {});
  const licenseIndex = buildLicenseIndex_(readEmployeeLicenses_(spreadsheet, timeZone));
  const shiftMap = indexBy_(readShiftTypes_(spreadsheet, timeZone), 'Shift Code');
  const uniqueChanges = {};
  const operatorIsAdmin = isAdminUser_(operator);

  changes.forEach(function(item) {
    const input = item || {};
    const employeeId = String(input.employeeId || '').trim();
    const date = String(input.date || '').trim();
    const code = String(input.code || '').trim().toUpperCase();
    if (!employeeId || !date || !code) throw new Error('Every change requires employee, date and shift.');
    parseIsoDateSafe_(date);
    if (!employeeMap[employeeId]) throw new Error('Employee not found: ' + employeeId);
    if (!shiftMap[code]) throw new Error('Shift code not found: ' + code);
    const license = licenseStatusForDate_(employeeId, date, licenseIndex);
    const workingShift = isWorkingShift_(shiftMap[code]);
    const licenseOverride = Boolean(operatorIsAdmin && workingShift && !license.valid && booleanValue_(input.licenseOverride));
    const overrideReason = String(input.overrideReason || '').trim();
    if (workingShift && !license.valid && !licenseOverride) {
      throw new Error(employeeId + ' วันที่ ' + date + ': ' + license.reason + (operatorIsAdmin ? ' · Admin ต้องเลือก Override และระบุเหตุผล' : ' · Manager ไม่มีสิทธิ Override กรุณาเลือก OFF/AL หรือติดต่อ Admin'));
    }
    if (licenseOverride && overrideReason.length < 5) {
      throw new Error(employeeId + ' วันที่ ' + date + ': เหตุผล Override ต้องมีอย่างน้อย 5 ตัวอักษร');
    }
    uniqueChanges[employeeId + '|' + date] = {
      employeeId: employeeId,
      date: date,
      code: code,
      remark: String(input.remark || 'Manual batch edit'),
      employee: employeeMap[employeeId],
      shift: shiftMap[code],
      license: license,
      licenseOverride: licenseOverride,
      overrideReason: licenseOverride ? overrideReason : ''
    };
  });

  const normalized = Object.keys(uniqueChanges).map(function(key) { return uniqueChanges[key]; });
  const schedule = readSchedule_(spreadsheet, timeZone);
  const hoursByEmployeeWeek = {};
  const changedSeen = {};

  function addHours(employeeId, date, hours) {
    const key = employeeId + '|' + isoWeek_(date);
    hoursByEmployeeWeek[key] = (hoursByEmployeeWeek[key] || 0) + (Number(hours) || 0);
  }

  schedule.forEach(function(row) {
    const key = row.employeeId + '|' + row.date;
    const change = uniqueChanges[key];
    if (change) {
      if (changedSeen[key]) return;
      changedSeen[key] = true;
      addHours(change.employeeId, change.date, change.shift.Hours);
      return;
    }
    addHours(row.employeeId, row.date, row.hours);
  });
  normalized.forEach(function(change) {
    const key = change.employeeId + '|' + change.date;
    if (!changedSeen[key]) addHours(change.employeeId, change.date, change.shift.Hours);
  });

  const overLimit = Object.keys(hoursByEmployeeWeek).filter(function(key) {
    return hoursByEmployeeWeek[key] > DEFAULT_MAX_WEEKLY_HOURS;
  });
  if (overLimit.length) {
    const detail = overLimit.slice(0, 5).map(function(key) {
      return key + ' = ' + round_(hoursByEmployeeWeek[key]) + ' hours';
    }).join(', ');
    throw new Error('Weekly hours cannot exceed 72. Please adjust: ' + detail);
  }

  const values = sheet.getDataRange().getValues();
  const dateColumn = headers.indexOf('Date');
  const employeeColumn = headers.indexOf('Employee ID');
  const rowByKey = {};
  for (let rowIndex = 1; rowIndex < values.length; rowIndex++) {
    const employeeId = String(values[rowIndex][employeeColumn] || '').trim();
    const date = formatDate_(values[rowIndex][dateColumn], timeZone);
    if (employeeId && date && rowByKey[employeeId + '|' + date] === undefined) {
      rowByKey[employeeId + '|' + date] = rowIndex + 1;
    }
  }

  const backupName = ''; // Backup disabled for manual edits to avoid clutter
  normalized.forEach(function(change) {
    const record = {
      Date: parseIsoDateSafe_(change.date),
      'Employee ID': change.employeeId,
      'Employee Name': String(change.employee.Name || ''),
      Department: String(change.employee.Department || ''),
      'Shift Code': change.code,
      'Start Time': change.shift['Start Time'],
      'End Time': change.shift['End Time'],
      Hours: change.shift.Hours,
      Remark: change.remark,
      Source: 'Manual',
      Locked: true,
      'Updated By': operator.email,
      'Updated At': new Date(),
      'License Status': change.license.code,
      'License Expiry Date': change.license.expiryDate ? parseIsoDateSafe_(change.license.expiryDate) : '',
      'License Override': change.licenseOverride,
      'Override Reason': change.overrideReason,
      'Override By': change.licenseOverride ? operator.email : '',
      'Override At': change.licenseOverride ? new Date() : ''
    };
    const output = headers.map(function(header) {
      return Object.prototype.hasOwnProperty.call(record, header) ? record[header] : '';
    });
    const key = change.employeeId + '|' + change.date;
    if (rowByKey[key]) sheet.getRange(rowByKey[key], 1, 1, headers.length).setValues([output]);
    else {
      sheet.appendRow(output);
      rowByKey[key] = sheet.getLastRow();
    }
    if (change.licenseOverride) {
      appendLicenseAudit_(spreadsheet, {
        Timestamp: new Date(),
        Action: 'MANUAL_SCHEDULE_OVERRIDE',
        'Employee ID': change.employeeId,
        'License ID': change.license.licenseId,
        'Work Date': parseIsoDateSafe_(change.date),
        'Shift Code': change.code,
        'License Status': change.license.code,
        'Expiry Date': change.license.expiryDate ? parseIsoDateSafe_(change.license.expiryDate) : '',
        Reason: change.overrideReason,
        'Approved By': operator.email
      });
    }
  });

  SpreadsheetApp.flush();
  const affectedMonths = unique_(normalized.map(function(change) { return change.date.slice(0, 7); }));
  const approvals = invalidateScheduleApprovals_(spreadsheet, affectedMonths, operator.email, 'MANUAL_SCHEDULE');
  bumpDataVersion_();
  return {
    updatedRows: normalized.length,
    backupSheet: backupName,
    licenseOverrides: normalized.filter(function(change) { return change.licenseOverride; }).length,
    approvals: approvals,
    results: normalized.map(function(change) {
      return { employeeId: change.employeeId, date: change.date, code: change.code, locked: true, licenseOverride: change.licenseOverride };
    })
  };
}

function updateEmployeeShift(payload, token) {
  const batchResult = updateEmployeeShifts([payload], token);
  return Object.assign({}, batchResult.results[0], { backupSheet: batchResult.backupSheet });
}


function ensureUserPasswordHashes_(sheet) {
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  let hashColumn = headers.indexOf('Password Hash') + 1;
  if (!hashColumn) {
    hashColumn = lastColumn + 1;
    sheet.getRange(1, hashColumn).setValue('Password Hash');
  }
  if (sheet.getLastRow() < 2) return;
  const emailColumn = headers.indexOf('Email') + 1;
  const statusColumn = headers.indexOf('Status') + 1;
  if (!emailColumn || !statusColumn) throw new Error('ชีต Users ไม่มีคอลัมน์ Email หรือ Status');
  const rowCount = sheet.getLastRow() - 1;
  const emails = sheet.getRange(2, emailColumn, rowCount, 1).getValues();
  const hashes = sheet.getRange(2, hashColumn, rowCount, 1).getValues();
  const statuses = sheet.getRange(2, statusColumn, rowCount, 1).getValues();
  let changed = false;
  hashes.forEach(function(row, index) {
    const email = String(emails[index][0] || '').trim().toLowerCase();
    if (!email || row[0]) return;
    if (email === PRIMARY_ADMIN_EMAIL) row[0] = PRIMARY_ADMIN_PASSWORD_HASH;
    else if (String(statuses[index][0] || '').trim().toLowerCase() === 'active') statuses[index][0] = 'Pending';
    changed = true;
  });
  if (changed) {
    sheet.getRange(2, hashColumn, rowCount, 1).setValues(hashes);
    sheet.getRange(2, statusColumn, rowCount, 1).setValues(statuses);
  }
}

function ensureUserSchema_(spreadsheet) {
  const schemas = {
    Users: [
      'User ID', 'Employee ID', 'Name', 'Email', 'Role', 'Department', 'Status', 'Password Hash',
      'Requested At', 'Approved By', 'Approved At', 'Rejection Reason', 'Updated At', 'Last Login At'
    ],
    'User Audit Log': ['Timestamp', 'Action', 'User ID', 'Employee ID', 'Email', 'Role', 'Department', 'Reason', 'Performed By']
  };
  Object.keys(schemas).forEach(function(sheetName) {
    const sheet = spreadsheet.getSheetByName(sheetName);
    const required = schemas[sheetName];
    const lastColumn = Math.max(sheet.getLastColumn(), 1);
    const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    if (!headers.some(function(header) { return String(header || '').trim(); })) {
      sheet.getRange(1, 1, 1, required.length).setValues([required]);
    } else {
      required.forEach(function(header) {
        if (headers.indexOf(header) !== -1) return;
        headers.push(header);
        sheet.getRange(1, headers.length).setValue(header);
      });
    }
    sheet.setFrozenRows(1);
  });
}

function ensurePrimaryAdmin_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName('Users');
  ensureUserPasswordHashes_(sheet);
  const properties = PropertiesService.getScriptProperties();
  const needsBootstrap = properties.getProperty('SHIFTFLOW_PRIMARY_ADMIN_VERSION') !== PRIMARY_ADMIN_BOOTSTRAP_VERSION;
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const emailColumn = headers.indexOf('Email');
  let primaryRow = 0;
  const legacyEmails = ['manager@example.com', 'production.lead@example.com', 'qa.lead@example.com', 'viewer@example.com'];
  for (let index = 1; index < values.length; index++) {
    const email = String(values[index][emailColumn] || '').trim().toLowerCase();
    if (email === PRIMARY_ADMIN_EMAIL) primaryRow = index + 1;
    if (needsBootstrap && legacyEmails.indexOf(email) !== -1) {
      const statusColumn = headers.indexOf('Status');
      const roleColumn = headers.indexOf('Role');
      values[index][statusColumn] = 'Inactive';
      values[index][roleColumn] = 'Viewer';
      sheet.getRange(index + 1, 1, 1, headers.length).setValues([values[index]]);
    }
  }
  const now = new Date();
  const existing = primaryRow ? values[primaryRow - 1] : [];
  const idColumn = headers.indexOf('User ID');
  const existingId = primaryRow ? String(existing[idColumn] || '').trim() : '';
  const duplicateExistingId = existingId && values.some(function(row, index) {
    return index > 0 && index + 1 !== primaryRow && String(row[idColumn] || '').trim() === existingId;
  });
  let primaryUserId = existingId && !duplicateExistingId ? existingId : 'USR-PRIMARY';
  if (values.some(function(row, index) {
    return index > 0 && index + 1 !== primaryRow && String(row[idColumn] || '').trim() === primaryUserId;
  })) primaryUserId = 'USR-' + Utilities.getUuid().replace(/-/g, '').slice(0, 10).toUpperCase();
  const record = {
    'User ID': primaryUserId,
    Name: primaryRow ? String(existing[headers.indexOf('Name')] || 'Sermpong Admin') : 'Sermpong Admin',
    Email: PRIMARY_ADMIN_EMAIL,
    Role: 'Admin',
    Department: 'All',
    Status: 'Active',
    'Password Hash': needsBootstrap || !primaryRow ? PRIMARY_ADMIN_PASSWORD_HASH : String(existing[headers.indexOf('Password Hash')] || PRIMARY_ADMIN_PASSWORD_HASH),
    'Requested At': primaryRow ? existing[headers.indexOf('Requested At')] || now : now,
    'Approved By': 'SYSTEM',
    'Approved At': primaryRow ? existing[headers.indexOf('Approved At')] || now : now,
    'Rejection Reason': '',
    'Updated At': now,
    'Last Login At': primaryRow ? existing[headers.indexOf('Last Login At')] || '' : ''
  };
  const output = headers.map(function(header) {
    return Object.prototype.hasOwnProperty.call(record, header) ? record[header] : '';
  });
  if (primaryRow) sheet.getRange(primaryRow, 1, 1, headers.length).setValues([output]);
  else sheet.appendRow(output);
  if (needsBootstrap) {
    appendUserAudit_(spreadsheet, {
      Timestamp: now, Action: 'PRIMARY_ADMIN_BOOTSTRAPPED', 'User ID': record['User ID'],
      Email: PRIMARY_ADMIN_EMAIL, Role: 'Admin', Department: 'All', Reason: 'ตั้งค่าบัญชี Admin เริ่มต้น', 'Performed By': 'SYSTEM'
    });
  }
  properties.setProperty('SHIFTFLOW_PRIMARY_ADMIN_VERSION', PRIMARY_ADMIN_BOOTSTRAP_VERSION);
}

function hashPassword_(email, password) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(email).toLowerCase() + '|' + String(password) + '|' + LOGIN_HASH_PEPPER,
    Utilities.Charset.UTF_8
  );
  return bytes.map(function(byte) {
    const value = byte < 0 ? byte + 256 : byte;
    return ('0' + value.toString(16)).slice(-2);
  }).join('');
}


/* -------------------------------------- optimized month/department loading */

function dataVersion_() {
  return PropertiesService.getScriptProperties().getProperty('SHIFTFLOW_DATA_VERSION') || '1';
}

function bumpDataVersion_() {
  const properties = PropertiesService.getScriptProperties();
  const next = String((Number(properties.getProperty('SHIFTFLOW_DATA_VERSION')) || 1) + 1);
  properties.setProperty('SHIFTFLOW_DATA_VERSION', next);
  return next;
}

function normalizedMonth_(value, availableMonths) {
  const text = String(value || '').trim().slice(0, 7);
  if (/^\d{4}-\d{2}$/.test(text)) return text;
  const current = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
  if ((availableMonths || []).indexOf(current) !== -1) return current;
  return (availableMonths || []).slice().sort().pop() || current;
}

function normalizedDepartments_(input, available) {
  const requested = Array.isArray(input) ? input : [];
  const allowed = (available || []).reduce(function(map, department) {
    map[department] = true;
    return map;
  }, {});
  return unique_(requested.map(function(value) {
    return String(value || '').trim();
  }).filter(function(value) {
    return value && allowed[value];
  })).sort();
}

function scheduleRowsForScope_(spreadsheet, timeZone, month, departments, preloadedDateValues) {
  const sheet = spreadsheet.getSheetByName('Schedule');
  ensureScheduleMetadataColumns_(sheet);
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2) return { display: [], context: [], availableMonths: [] };

  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const dateColumn = headers.indexOf('Date') + 1;
  const departmentColumn = headers.indexOf('Department') + 1;
  if (!dateColumn || !departmentColumn) throw new Error('Schedule must contain Date and Department columns.');

  const rowCount = lastRow - 1;
  const dateValues = Array.isArray(preloadedDateValues) && preloadedDateValues.length === rowCount
    ? preloadedDateValues
    : sheet.getRange(2, dateColumn, rowCount, 1).getValues();
  const departmentValues = sheet.getRange(2, departmentColumn, rowCount, 1).getValues();
  const availableMonths = unique_(dateValues.map(function(row) {
    const date = formatDate_(row[0], timeZone);
    return date ? date.slice(0, 7) : '';
  }).filter(Boolean)).sort();

  const start = parseIsoDateSafe_(month + '-01');
  const end = new Date(start.getFullYear(), start.getMonth() + 1, 0, 12, 0, 0);
  const contextStart = new Date(start.getTime());
  contextStart.setDate(contextStart.getDate() - 7);
  const contextEnd = new Date(end.getTime());
  contextEnd.setDate(contextEnd.getDate() + 7);
  const contextStartIso = isoDate_(contextStart, timeZone);
  const contextEndIso = isoDate_(contextEnd, timeZone);
  const selected = departments.reduce(function(map, value) {
    map[value] = true;
    return map;
  }, {});

  const matchingRows = [];
  dateValues.forEach(function(row, index) {
    const date = formatDate_(row[0], timeZone);
    const department = String(departmentValues[index][0] || '').trim();
    if (!date || date < contextStartIso || date > contextEndIso) return;
    if (departments.length && !selected[department]) return;
    matchingRows.push(index + 2);
  });
  if (!matchingRows.length) return { display: [], context: [], availableMonths: availableMonths };

  const groups = [];
  matchingRows.forEach(function(rowNumber) {
    const last = groups[groups.length - 1];
    if (last && rowNumber === last.end + 1) last.end = rowNumber;
    else groups.push({ start: rowNumber, end: rowNumber });
  });

  const context = [];
  groups.forEach(function(group) {
    const values = sheet.getRange(group.start, 1, group.end - group.start + 1, lastColumn).getValues();
    values.forEach(function(row) {
      const object = headers.reduce(function(result, header, index) {
        result[header] = row[index];
        return result;
      }, {});
      const normalized = {
        date: formatDate_(object.Date, timeZone),
        employeeId: String(object['Employee ID'] || '').trim(),
        employeeName: String(object['Employee Name'] || ''),
        department: String(object.Department || ''),
        code: String(object['Shift Code'] || '').trim().toUpperCase(),
        startTime: formatTime_(object['Start Time'], timeZone),
        endTime: formatTime_(object['End Time'], timeZone),
        hours: Number(object.Hours) || 0,
        remark: String(object.Remark || ''),
        source: String(object.Source || ''),
        locked: scheduleLocked_(object.Locked),
        updatedBy: String(object['Updated By'] || ''),
        updatedAt: object['Updated At'] ? formatDateTime_(object['Updated At'], timeZone) : '',
        licenseStatus: String(object['License Status'] || ''),
        licenseExpiryDate: formatDate_(object['License Expiry Date'], timeZone),
        licenseOverride: booleanValue_(object['License Override']),
        overrideReason: String(object['Override Reason'] || ''),
        overrideBy: String(object['Override By'] || ''),
        overrideAt: object['Override At'] ? formatDateTime_(object['Override At'], timeZone) : ''
      };
      if (normalized.date && normalized.employeeId) context.push(normalized);
    });
  });

  return {
    context: context,
    display: context.filter(function(row) { return row.date.slice(0, 7) === month; }),
    availableMonths: availableMonths
  };
}

function buildEmployeeSchedulesFast_(employees, schedule, shiftTypeMap, dates, licenseIndex) {
  const scheduleByKey = schedule.reduce(function(map, row) {
    map[row.employeeId + '|' + row.date] = row;
    return map;
  }, {});
  return employees.map(function(employee) {
    const employeeId = String(employee['Employee ID'] || '').trim();
    return {
      id: employeeId,
      name: String(employee.Name || ''),
      department: String(employee.Department || ''),
      position: String(employee.Position || ''),
      skill: String(employee.Skill || ''),
      status: String(employee.Status || ''),
      shifts: dates.map(function(date) {
        const row = scheduleByKey[employeeId + '|' + date];
        const type = row ? shiftTypeMap[row.code] || {} : {};
        const license = licenseStatusForDate_(employeeId, date, licenseIndex || {});
        return {
          date: date,
          code: row ? row.code || 'OFF' : 'OFF',
          startTime: row ? row.startTime : '',
          endTime: row ? row.endTime : '',
          hours: row ? row.hours : 0,
          remark: row ? row.remark : '',
          name: type['Shift Name'] || '',
          source: row ? row.source : '',
          locked: row ? row.locked : false,
          updatedBy: row ? row.updatedBy : '',
          licenseValid: license.valid,
          licenseCode: license.code,
          licenseReason: license.reason,
          licenseExpiryDate: license.expiryDate,
          licenseOverride: row ? row.licenseOverride : false,
          overrideReason: row ? row.overrideReason : '',
          overrideBy: row ? row.overrideBy : ''
        };
      })
    };
  });
}

function getEmployeesPage(request, token) {
  requireActiveSession_(token);
  const input = request || {};
  const spreadsheet = getOrCreateSpreadsheet_();
  ensureRequiredSheets_(spreadsheet);
  const timeZone = spreadsheet.getSpreadsheetTimeZone();
  const allEmployees = readObjects_(spreadsheet.getSheetByName('Employees'));
  const allLicenses = readEmployeeLicenses_(spreadsheet, timeZone);
  const licenseIndex = buildLicenseIndex_(allLicenses);
  const pageSize = Math.min(100, Math.max(1, Number(input.pageSize) || 100));
  const totalPages = Math.max(1, Math.ceil(allEmployees.length / pageSize));
  const page = Math.min(totalPages, Math.max(1, Number(input.page) || 1));
  const pageEmployeeRows = allEmployees.slice((page - 1) * pageSize, page * pageSize);
  const pageEmployees = pageEmployeeRows.map(function(employee) {
    return {
      id: String(employee['Employee ID'] || '').trim(),
      name: String(employee.Name || ''),
      department: String(employee.Department || ''),
      position: String(employee.Position || ''),
      skill: String(employee.Skill || ''),
      status: String(employee.Status || '')
    };
  });
  // Keep the employee table paginated, but provide a lightweight directory for
  // license names/forms and return every license. Filtering licenses by the
  // current employee page caused existing licenses to disappear from Master Data.
  const employeeDirectory = allEmployees.map(function(employee) {
    return {
      id: String(employee['Employee ID'] || '').trim(),
      name: String(employee.Name || ''),
      department: String(employee.Department || ''),
      status: String(employee.Status || '')
    };
  }).filter(function(employee) { return employee.id; });
  const today = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd');
  return {
    employees: pageEmployees,
    employeeDirectory: employeeDirectory,
    licenses: allLicenses,
    licenseAlerts: licenseAlerts_(allEmployees, licenseIndex, today),
    availableDepartments: unique_(allEmployees.map(function(employee) {
      return String(employee.Department || '').trim();
    }).filter(Boolean)).sort(),
    pagination: {
      page: page,
      pageSize: pageSize,
      totalEmployees: allEmployees.length,
      totalPages: totalPages,
      from: allEmployees.length ? (page - 1) * pageSize + 1 : 0,
      to: Math.min(page * pageSize, allEmployees.length)
    }
  };
}

function getOptimizedScheduleData(request, token) {
  const sessionUser = requireActiveSession_(token);
  const userRole = String(sessionUser.role || sessionUser.Role || '').trim().toLowerCase();
  const input = request || {};
  const spreadsheet = getOrCreateSpreadsheet_();
  ensureRequiredSheets_(spreadsheet);
  enforceMaxWeeklyHours72_(spreadsheet.getSheetByName('Rules'));
  const timeZone = spreadsheet.getSpreadsheetTimeZone();

  const isAdmin = userRole === 'admin' || isAdminUser_(sessionUser);
  const isManager = userRole === 'manager';
  
  const users = (isAdmin || isManager) ? safeUsers_(spreadsheet).filter(function(user) {
    return isAdmin || String(user.Status || '').trim().toLowerCase() === 'pending';
  }) : [];
  
  const settings = isAdmin ? normalizeRows_(readObjects_(spreadsheet.getSheetByName('Settings')), timeZone) : [];
  const dashboard = normalizeRows_(readObjects_(spreadsheet.getSheetByName('Dashboard')), timeZone);

  const allEmployees = readObjects_(spreadsheet.getSheetByName('Employees'));
  const allLicenses = readEmployeeLicenses_(spreadsheet, timeZone);
  const licenseIndex = buildLicenseIndex_(allLicenses);
  const availableDepartments = unique_(allEmployees.map(function(employee) {
    return String(employee.Department || '').trim();
  }).filter(Boolean)).sort();
  const departments = normalizedDepartments_(input.departments, availableDepartments);

  const scheduleSheet = spreadsheet.getSheetByName('Schedule');
  const dateColumn = scheduleSheet.getRange(1, 1, 1, scheduleSheet.getLastColumn()).getValues()[0].indexOf('Date') + 1;
  let scheduleDateValues = [];
  let availableMonths = [];
  if (dateColumn && scheduleSheet.getLastRow() > 1) {
    scheduleDateValues = scheduleSheet.getRange(2, dateColumn, scheduleSheet.getLastRow() - 1, 1).getValues();
    availableMonths = unique_(scheduleDateValues.map(function(row) {
      const date = formatDate_(row[0], timeZone);
      return date ? date.slice(0, 7) : '';
    }).filter(Boolean)).sort();
  }
  const month = normalizedMonth_(input.month, availableMonths);
  const scheduleApproval = ensureScheduleApprovalCurrent_(spreadsheet, month, 'SYSTEM DATA CHECK');
  const scopedSchedule = scheduleRowsForScope_(spreadsheet, timeZone, month, departments, scheduleDateValues);
  availableMonths = scopedSchedule.availableMonths.length ? scopedSchedule.availableMonths : availableMonths;

  const selectedMap = departments.reduce(function(map, department) {
    map[department] = true;
    return map;
  }, {});
  const scopedEmployees = allEmployees.filter(function(employee) {
    const department = String(employee.Department || '').trim();
    return !departments.length || selectedMap[department];
  });
  const pageSize = Math.min(100, Math.max(1, Number(input.pageSize) || 100));
  const totalPages = Math.max(1, Math.ceil(scopedEmployees.length / pageSize));
  const page = Math.min(totalPages, Math.max(1, Number(input.page) || 1));
  const pageEmployees = scopedEmployees.slice((page - 1) * pageSize, page * pageSize);

  const shiftTypes = readShiftTypes_(spreadsheet, timeZone);
  const shiftTypeMap = indexBy_(shiftTypes, 'Shift Code');
  const dates = monthDates_(month, timeZone);
  const rules = parseRules_(readObjects_(spreadsheet.getSheetByName('Rules')));
  const cacheSeed = dataVersion_() + '|' + month + '|' + departments.join('|');
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, cacheSeed, Utilities.Charset.UTF_8);
  const cacheKey = 'shiftflow-dashboard-' + Utilities.base64EncodeWebSafe(digest).slice(0, 32);
  const cache = CacheService.getScriptCache();
  let summary = null;
  try {
    const cached = cache.get(cacheKey);
    if (cached) summary = JSON.parse(cached);
  } catch (ignore) {}

  if (!summary) {
    const maxWeeklyHours = ruleValue_(rules, 'RULE001', DEFAULT_MAX_WEEKLY_HOURS);
    const overtime = computeOvertime_(scopedSchedule.context, maxWeeklyHours);
    const ruleResults = evaluateRules_(scopedSchedule.display, scopedEmployees, rules, dates, overtime);
    const licenseScheduleViolations = licenseScheduleViolations_(scopedSchedule.display, licenseIndex);
    const today = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd');
    const licenseAlerts = licenseAlerts_(scopedEmployees, licenseIndex, today);
    const licenseRuleResult = {
      id: 'LICENSE',
      name: 'ใบอนุญาตต้องมีผลในวันที่ทำงาน',
      enabled: true,
      passed: licenseScheduleViolations.length === 0,
      summary: licenseScheduleViolations.length ? licenseScheduleViolations.length + ' รายการ' : 'ผ่าน',
      violations: licenseScheduleViolations
    };
    const combinedRuleResults = ruleResults.concat([licenseRuleResult]);
    const violations = combinedRuleResults.reduce(function(all, result) {
      return all.concat(result.violations);
    }, []).concat(licenseAlerts);
    const activeEmployees = scopedEmployees.filter(function(employee) {
      return String(employee.Status || '').toLowerCase() === 'active';
    }).length;
    const totalHours = scopedSchedule.display.reduce(function(sum, row) { return sum + row.hours; }, 0);
    summary = {
      analytics: buildAnalytics_(scopedSchedule.display, scopedEmployees, shiftTypes, rules, dates, overtime),
      ruleResults: combinedRuleResults.map(function(result) {
        return { id: result.id, name: result.name, enabled: result.enabled, passed: result.passed, count: result.violations.length, summary: result.summary };
      }),
      violations: violations,
      licenseAlerts: licenseAlerts,
      metrics: {
        totalEmployees: scopedEmployees.length,
        activeEmployees: activeEmployees,
        totalHours: round_(totalHours),
        otHours: round_(overtime.total),
        maxWeeklyHours: maxWeeklyHours,
        scheduledRows: scopedSchedule.display.length,
        violations: violations.length,
        licenseExpired: licenseAlerts.filter(function(alert) { return alert.severity === 'error'; }).length,
        licenseExpiring: licenseAlerts.filter(function(alert) { return alert.code === 'EXPIRING'; }).length,
        rulesPassed: combinedRuleResults.filter(function(result) { return result.enabled && result.passed; }).length,
        rulesChecked: combinedRuleResults.filter(function(result) { return result.enabled; }).length
      },
      departments: countBy_(scopedEmployees, 'Department')
    };
    try { cache.put(cacheKey, JSON.stringify(summary), 300); } catch (ignore) {}
  }

  return {
    spreadsheetUrl: spreadsheet.getUrl(),
    timeZone: timeZone,
    selectedMonth: month,
    selectedDepartments: departments,
    scheduleApproval: scheduleApproval,
    availableMonths: availableMonths,
    availableDepartments: availableDepartments,
    dates: dates,
    employees: buildEmployeeSchedulesFast_(pageEmployees, scopedSchedule.display, shiftTypeMap, dates, licenseIndex),
    licenseAlerts: summary.licenseAlerts,
    shiftTypes: shiftTypes,
    scheduleRows: scopedSchedule.display,
    rules: rules.map(function(rule) { return rule.row; }),
    ruleResults: summary.ruleResults,
    dashboard: dashboard,
    users: users,
    settings: settings,
    departments: summary.departments,
    analytics: summary.analytics,
    metrics: summary.metrics,
    violations: summary.violations,
    pagination: {
      page: page,
      pageSize: pageSize,
      totalEmployees: scopedEmployees.length,
      totalPages: totalPages,
      from: scopedEmployees.length ? (page - 1) * pageSize + 1 : 0,
      to: Math.min(page * pageSize, scopedEmployees.length)
    }
  };
}

function archiveScheduleBeforeMonth(cutoffMonth, token) {
  const admin = requireAdmin_(token);
  const month = normalizedMonth_(cutoffMonth, []);
  const cutoff = month + '-01';
  const spreadsheet = getOrCreateSpreadsheet_();
  const timeZone = spreadsheet.getSpreadsheetTimeZone();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('The schedule is busy. Please try again.');
  try {
    const source = spreadsheet.getSheetByName('Schedule');
    ensureScheduleMetadataColumns_(source);
    const sourceValues = source.getDataRange().getValues();
    const headers = sourceValues.shift();
    const dateColumn = headers.indexOf('Date');
    const employeeColumn = headers.indexOf('Employee ID');
    if (dateColumn === -1 || employeeColumn === -1) throw new Error('Schedule headers are incomplete.');

    const archiveHeaders = headers.concat(['Archived At', 'Archived By']);
    let archive = spreadsheet.getSheetByName('Schedule Archive');
    if (!archive) {
      archive = spreadsheet.insertSheet('Schedule Archive');
      archive.getRange(1, 1, 1, archiveHeaders.length).setValues([archiveHeaders]);
      archive.setFrozenRows(1);
    } else {
      const existingHeaders = archive.getRange(1, 1, 1, Math.max(archive.getLastColumn(), archiveHeaders.length)).getValues()[0];
      archiveHeaders.forEach(function(header, index) {
        if (existingHeaders[index] !== header) archive.getRange(1, index + 1).setValue(header);
      });
    }

    const existingKeys = {};
    if (archive.getLastRow() > 1) {
      const archiveValues = archive.getRange(2, 1, archive.getLastRow() - 1, archiveHeaders.length).getValues();
      archiveValues.forEach(function(row) {
        const date = formatDate_(row[dateColumn], timeZone);
        const employeeId = String(row[employeeColumn] || '').trim();
        if (date && employeeId) existingKeys[date + '|' + employeeId] = true;
      });
    }

    const kept = [];
    const toArchive = [];
    sourceValues.forEach(function(row) {
      const date = formatDate_(row[dateColumn], timeZone);
      if (!date || date >= cutoff) {
        kept.push(headers.map(function(_, index) { return row[index] === undefined ? '' : row[index]; }));
        return;
      }
      const employeeId = String(row[employeeColumn] || '').trim();
      const key = date + '|' + employeeId;
      if (!existingKeys[key]) {
        toArchive.push(headers.map(function(_, index) { return row[index] === undefined ? '' : row[index]; }).concat([new Date(), admin.email]));
        existingKeys[key] = true;
      }
    });

    if (toArchive.length) {
      archive.getRange(archive.getLastRow() + 1, 1, toArchive.length, archiveHeaders.length).setValues(toArchive);
    }
    if (source.getMaxRows() > 1) source.getRange(2, 1, source.getMaxRows() - 1, headers.length).clearContent();
    if (kept.length) source.getRange(2, 1, kept.length, headers.length).setValues(kept);
    source.getRange(2, 1, Math.max(kept.length, 1), 1).setNumberFormat('yyyy-mm-dd');
    SpreadsheetApp.flush();
    bumpDataVersion_();
    return { cutoffMonth: month, movedRows: sourceValues.length - kept.length, appendedRows: toArchive.length, archiveSheet: archive.getName() };
  } finally {
    lock.releaseLock();
  }
}


function ensureNewRules_(spreadsheet) {
  const rulesSheet = spreadsheet.getSheetByName('Rules');
  if (!rulesSheet) return;
  const data = rulesSheet.getDataRange().getValues();
  const ruleIds = data.map(row => String(row[0] || '').trim());
  const toAppend = [];
  if (ruleIds.indexOf('RULE006') === -1) toAppend.push(['RULE006', 'Supervisor day shift only and Sunday off', 1, 'boolean', true]);
  if (ruleIds.indexOf('RULE007') === -1) toAppend.push(['RULE007', 'PO11 and WCS need Leader in Day shift', 1, 'people', true]);
  if (ruleIds.indexOf('RULE008') === -1) toAppend.push(['RULE008', 'PO11 and WCS need Leader in Night shift', 1, 'people', true]);
  if (ruleIds.indexOf('RULE009') === -1) toAppend.push(['RULE009', 'Spare department is excluded from auto scheduling', 1, 'boolean', true]);
  if (toAppend.length) {
    rulesSheet.getRange(rulesSheet.getLastRow() + 1, 1, toAppend.length, toAppend[0].length).setValues(toAppend);
  }
}

function getAppData(token) {
  const sessionUser = requireActiveSession_(token);
  const spreadsheet = getOrCreateSpreadsheet_();
  ensureRequiredSheets_(spreadsheet);
  ensureLicenseSchema_(spreadsheet);
  ensureNewRules_(spreadsheet);
  ensureScheduleMetadataColumns_(spreadsheet.getSheetByName('Schedule'));

  const timeZone = spreadsheet.getSpreadsheetTimeZone();
  const employees = readObjects_(spreadsheet.getSheetByName('Employees')).map(function(emp, index) {
    emp._originalIndex = index;
    return emp;
  });
  employees.sort(function(a, b) {
    if (a.Department === 'WCS' && b.Department === 'WCS') {
      const order = { 'ทวีศักดิ์ พวงมาลัย': 1, 'จำรุณ ประไพพงษ์': 2, 'อิทธิกร': 3, 'อนุวัฒน์': 4, 'ทวีป': 5 };
      const aName = String(a.Name || '').trim();
      const bName = String(b.Name || '').trim();
      let aOrder = 99; let bOrder = 99;
      for (let key in order) { if (aName.indexOf(key) !== -1) aOrder = order[key]; }
      for (let key in order) { if (bName.indexOf(key) !== -1) bOrder = order[key]; }
      if (aOrder !== bOrder) return aOrder - bOrder;
    }
    return a._originalIndex - b._originalIndex;
  });
  const licenses = readEmployeeLicenses_(spreadsheet, timeZone);
  const licenseIndex = buildLicenseIndex_(licenses);
  const shiftTypes = readShiftTypes_(spreadsheet, timeZone);
  const schedule = readSchedule_(spreadsheet, timeZone);
  const rules = parseRules_(readObjects_(spreadsheet.getSheetByName('Rules')));
  const dashboard = normalizeRows_(readObjects_(spreadsheet.getSheetByName('Dashboard')), timeZone);
  const users = String(sessionUser.role || '').toLowerCase() === 'admin' ? safeUsers_(spreadsheet) : [];
  const settings = normalizeRows_(readObjects_(spreadsheet.getSheetByName('Settings')), timeZone);

  const dates = unique_(schedule.map(function(row) { return row.date; })).sort();
  const shiftTypeMap = indexBy_(shiftTypes, 'Shift Code');
  const maxWeeklyHours = ruleValue_(rules, 'RULE001', DEFAULT_MAX_WEEKLY_HOURS);
  const overtime = computeOvertime_(schedule, maxWeeklyHours);
  const ruleResults = evaluateRules_(schedule, employees, rules, dates, overtime);
  const licenseScheduleViolations = licenseScheduleViolations_(schedule, licenseIndex);
  const today = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd');
  const licenseAlerts = licenseAlerts_(employees, licenseIndex, today);
  const combinedRuleResults = ruleResults.concat([{
    id: 'LICENSE',
    name: 'ใบอนุญาตต้องมีผลในวันที่ทำงาน',
    enabled: true,
    passed: licenseScheduleViolations.length === 0,
    summary: licenseScheduleViolations.length ? licenseScheduleViolations.length + ' รายการ' : 'ผ่าน',
    violations: licenseScheduleViolations
  }]);
  const violations = combinedRuleResults.reduce(function(all, result) {
    return all.concat(result.violations);
  }, []).concat(licenseAlerts);

  const activeEmployees = employees.filter(function(employee) {
    return String(employee.Status).toLowerCase() === 'active';
  }).length;
  const totalHours = schedule.reduce(function(sum, row) { return sum + row.hours; }, 0);

  const isViewer = String(sessionUser.role || '').toLowerCase() === 'viewer';

  return {
    spreadsheetUrl: spreadsheet.getUrl(),
    timeZone: timeZone,
    dates: dates,
    employees: buildEmployeeSchedules_(employees, schedule, shiftTypeMap, dates, licenseIndex),
    licenses: isViewer ? [] : licenses,
    licenseAlerts: isViewer ? [] : licenseAlerts,
    shiftTypes: shiftTypes,
    scheduleRows: schedule,
    rules: rules.map(function(rule) { return rule.row; }),
    ruleResults: combinedRuleResults.map(function(result) {
      return {
        id: result.id,
        name: result.name,
        enabled: result.enabled,
        passed: result.passed,
        count: result.violations.length,
        summary: result.summary
      };
    }),
    dashboard: dashboard,
    users: users,
    settings: settings,
    departments: countBy_(employees, 'Department'),
    analytics: buildAnalytics_(schedule, employees, shiftTypes, rules, dates, overtime),
    metrics: {
      totalEmployees: employees.length,
      activeEmployees: activeEmployees,
      totalHours: round_(totalHours),
      otHours: round_(overtime.total),
      maxWeeklyHours: maxWeeklyHours,
      scheduledRows: schedule.length,
      violations: violations.length,
      licenseExpired: licenseAlerts.filter(function(alert) { return alert.severity === 'error'; }).length,
      licenseExpiring: licenseAlerts.filter(function(alert) { return alert.code === 'EXPIRING'; }).length,
      rulesPassed: combinedRuleResults.filter(function(r) { return r.enabled && r.passed; }).length,
      rulesChecked: combinedRuleResults.filter(function(r) { return r.enabled; }).length
    },
    violations: violations
  };
}

/* ---------------------------------------------------------------- reading */

function readShiftTypes_(spreadsheet, timeZone) {
  return readObjects_(spreadsheet.getSheetByName('Shift Types')).map(function(row) {
    return {
      'Shift Code': String(row['Shift Code'] || '').trim().toUpperCase(),
      'Shift Name': String(row['Shift Name'] || ''),
      'Start Time': formatTime_(row['Start Time'], timeZone),
      'End Time': formatTime_(row['End Time'], timeZone),
      Hours: Number(row.Hours) || 0,
      Color: String(row.Color || '#CBD5E1')
    };
  }).filter(function(row) { return row['Shift Code']; });
}

function readSchedule_(spreadsheet, timeZone) {
  return readObjects_(spreadsheet.getSheetByName('Schedule')).map(function(row) {
    return {
      date: formatDate_(row.Date, timeZone),
      employeeId: String(row['Employee ID'] || '').trim(),
      employeeName: String(row['Employee Name'] || ''),
      department: String(row.Department || ''),
      code: String(row['Shift Code'] || '').trim().toUpperCase(),
      startTime: formatTime_(row['Start Time'], timeZone),
      endTime: formatTime_(row['End Time'], timeZone),
      hours: Number(row.Hours) || 0,
      remark: String(row.Remark || ''),
      source: String(row.Source || ''),
      locked: scheduleLocked_(row.Locked),
      updatedBy: String(row['Updated By'] || ''),
      updatedAt: row['Updated At'] ? formatDateTime_(row['Updated At'], timeZone) : '',
      licenseStatus: String(row['License Status'] || ''),
      licenseExpiryDate: formatDate_(row['License Expiry Date'], timeZone),
      licenseOverride: booleanValue_(row['License Override']),
      overrideReason: String(row['Override Reason'] || ''),
      overrideBy: String(row['Override By'] || ''),
      overrideAt: row['Override At'] ? formatDateTime_(row['Override At'], timeZone) : ''
    };
  }).filter(function(row) { return row.date && row.employeeId; });
}

function readObjects_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getValues();
  const headers = values.shift().map(function(h) { return String(h || '').trim(); });
  return values.map(function(row, rowIndex) {
    const obj = headers.reduce(function(object, header, index) {
      if (header !== '') object[header] = row[index];
      return object;
    }, {});
    obj._rowIndex = rowIndex + 2;
    obj._isEmpty = !row.some(function(cell) { return cell !== '' && cell !== null; });
    return obj;
  }).filter(function(obj) {
    return !obj._isEmpty;
  });
}

function buildEmployeeSchedules_(employees, schedule, shiftTypeMap, dates, licenseIndex) {
  return employees.map(function(employee) {
    const employeeId = String(employee['Employee ID'] || '').trim();
    const shifts = dates.map(function(date) {
      const row = schedule.find(function(item) {
        return item.employeeId === employeeId && item.date === date;
      });
      const code = row ? row.code : '';
      const type = shiftTypeMap[code] || {};
      const license = licenseStatusForDate_(employeeId, date, licenseIndex || {});
      return {
        date: date,
        code: code || 'OFF',
        startTime: row ? row.startTime : '',
        endTime: row ? row.endTime : '',
        hours: row ? row.hours : 0,
        remark: row ? row.remark : '',
        name: type['Shift Name'] || '',
        source: row ? row.source : '',
        locked: row ? row.locked : false,
        updatedBy: row ? row.updatedBy : '',
        licenseValid: license.valid,
        licenseCode: license.code,
        licenseReason: license.reason,
        licenseExpiryDate: license.expiryDate,
        licenseOverride: row ? row.licenseOverride : false,
        overrideReason: row ? row.overrideReason : '',
        overrideBy: row ? row.overrideBy : ''
      };
    });

    return {
      id: employeeId,
      name: String(employee.Name || ''),
      department: String(employee.Department || ''),
      position: String(employee.Position || ''),
      skill: String(employee.Skill || ''),
      status: String(employee.Status || ''),
      shifts: shifts
    };
  });
}

/* ------------------------------------------------------------------ rules */

function parseRules_(rows) {
  return rows.map(function(row) {
    return {
      id: String(row['Rule ID'] || '').trim().toUpperCase(),
      name: String(row['Rule Name'] || ''),
      value: Number(row.Value) || 0,
      unit: String(row.Unit || ''),
      enabled: row.Enabled === true || String(row.Enabled).toLowerCase() === 'true',
      row: {
        'Rule ID': String(row['Rule ID'] || ''),
        'Rule Name': String(row['Rule Name'] || ''),
        Value: row.Value,
        Unit: String(row.Unit || ''),
        Enabled: row.Enabled === true || String(row.Enabled).toLowerCase() === 'true'
      }
    };
  }).filter(function(rule) { return rule.id; });
}

function ruleValue_(rules, id, fallback) {
  const rule = rules.find(function(item) { return item.id === id && item.enabled; });
  return rule && rule.value ? rule.value : fallback;
}

/**
 * Overtime = hours worked beyond the weekly cap, per employee per ISO week.
 */
function computeOvertime_(schedule, maxWeeklyHours) {
  const byEmployeeWeek = {};
  const byDepartment = {};
  const byWeek = {};
  const byMonth = {};

  schedule.forEach(function(row) {
    if (!row.hours) return;
    const key = row.employeeId + '|' + isoWeek_(row.date);
    if (!byEmployeeWeek[key]) {
      byEmployeeWeek[key] = { employeeId: row.employeeId, employeeName: row.employeeName, department: row.department, week: isoWeek_(row.date), month: row.date.slice(0, 7), hours: 0 };
    }
    byEmployeeWeek[key].hours += row.hours;
  });

  let total = 0;
  const perEmployeeWeek = Object.keys(byEmployeeWeek).map(function(key) {
    const group = byEmployeeWeek[key];
    group.overtime = Math.max(0, group.hours - maxWeeklyHours);
    total += group.overtime;
    byDepartment[group.department] = (byDepartment[group.department] || 0) + group.overtime;
    byWeek[group.week] = (byWeek[group.week] || 0) + group.overtime;
    byMonth[group.month] = (byMonth[group.month] || 0) + group.overtime;
    return group;
  });

  return { total: total, perEmployeeWeek: perEmployeeWeek, byDepartment: byDepartment, byWeek: byWeek, byMonth: byMonth };
}

function evaluateRules_(schedule, employees, rules, dates, overtime) {
  return rules.map(function(rule) {
    if (!rule.enabled) {
      return { id: rule.id, name: rule.name, enabled: false, passed: true, summary: 'ปิดการตรวจสอบ', violations: [] };
    }

    const violations = ruleCheckers_(rule, schedule, employees, dates, overtime);
    return {
      id: rule.id,
      name: rule.name,
      enabled: true,
      passed: violations.length === 0,
      summary: violations.length === 0 ? 'ผ่าน' : violations.length + ' รายการ',
      violations: violations
    };
  });
}

function ruleCheckers_(rule, schedule, employees, dates, overtime) {
  if (rule.id === 'RULE001') return checkWeeklyHours_(rule, overtime);
  if (rule.id === 'RULE002') return checkRestAfterNight_(rule, schedule);
  if (COVERAGE_RULES[rule.id]) return checkCoverage_(rule, schedule, dates, employees);
  if (rule.id === 'RULE005') return checkLeaveConflicts_(rule, schedule);
  if (rule.id === 'RULE006') return checkSupervisorRule_(rule, schedule, employees);
  if (rule.id === 'RULE007') return checkLeaderRule_(rule, schedule, dates, employees, 'D', ['PO11', 'WCS']);
  if (rule.id === 'RULE008') return checkLeaderRule_(rule, schedule, dates, employees, 'N', ['PO11', 'WCS']);
  return checkRemarkFlags_(rule, schedule);
}

function checkWeeklyHours_(rule, overtime) {
  return overtime.perEmployeeWeek.filter(function(group) {
    return group.overtime > 0;
  }).map(function(group) {
    return violation_('error', rule, group.employeeName || group.employeeId,
      'ทำงาน ' + round_(group.hours) + ' ชม. ในสัปดาห์ ' + group.week +
      ' (สูงสุด ' + rule.value + ' ชม.) — เกิน ' + round_(group.overtime) + ' ชม.',
      { employeeId: group.employeeId });
  });
}

function checkRestAfterNight_(rule, schedule) {
  const byEmployee = groupBy_(schedule.filter(function(row) { return row.hours > 0; }), 'employeeId');

  return Object.keys(byEmployee).reduce(function(all, employeeId) {
    const shifts = byEmployee[employeeId]
      .map(function(row) { return Object.assign({}, row, shiftBounds_(row)); })
      .filter(function(row) { return row.start !== undefined; })
      .sort(function(a, b) { return a.start - b.start; });

    shifts.forEach(function(shift, index) {
      const next = shifts[index + 1];
      if (!next || shift.code !== 'N') return;
      const restHours = (next.start - shift.end) / 3600000;
      if (restHours >= rule.value) return;
      all.push(violation_('error', rule, shift.employeeName || employeeId,
        next.date + ' — ถูกจัดกะ ' + next.code + ' (' + next.startTime + ') หลังกะดึก โดยพักเพียง ' +
        round_(restHours) + ' ชม. (ต้องพัก ' + rule.value + ' ชม.)',
        { employeeId: employeeId, date: next.date }));
    });

    return all;
  }, []);
}

function checkCoverage_(rule, schedule, dates, employees) {
  const scope = COVERAGE_RULES[rule.id];
  const depts = scope.department === 'All'
    ? unique_((employees || []).map(function(e) { return String(e.Department || ''); }).filter(Boolean))
    : [scope.department];

  return dates.reduce(function(all, date) {
    depts.forEach(function(dept) {
      const staffed = schedule.filter(function(row) {
        return row.date === date && row.code === scope.code && String(row.department || '') === dept;
      }).length;
      if (staffed < rule.value) {
        all.push(violation_('warn', rule, 'แผนก ' + dept + ' · คนน้อยกว่ากำหนดขั้นต่ำ',
          date + ' กะ ' + scope.code + ' — มีคนลงเวร ' + staffed + ' คน (ขั้นต่ำ ' + rule.value + ' คน)',
          { date: date, department: dept }));
      }
    });
    return all;
  }, []);
}

function checkSupervisorRule_(rule, schedule, employees) {
  const supervisors = (employees || []).filter(function(emp) {
    return String(emp.Position || '').trim().toLowerCase() === 'supervisor';
  });
  const supMap = indexBy_(supervisors, 'Employee ID');

  return schedule.filter(function(row) {
    if (!supMap[row.employeeId]) return false;
    const isSunday = parseIsoDateSafe_(row.date).getDay() === 0;
    if (isSunday && row.code !== 'OFF' && row.hours > 0) return true;
    if (!isSunday && row.code !== 'D' && row.code !== 'OFF' && row.hours > 0) return true;
    return false;
  }).map(function(row) {
    return violation_('error', rule, row.employeeName || row.employeeId,
      row.date + ' — Supervisor ถูกจัดเข้ากะ ' + row.code + ' (กำหนดให้เฉพาะกะเช้า D และหยุดวันอาทิตย์)',
      { employeeId: row.employeeId, date: row.date });
  });
}

function checkLeaderRule_(rule, schedule, dates, employees, shiftCode, depts) {
  const isLeader = function(emp) {
    const p = String(emp.Position || '').trim().toLowerCase();
    return p === 'supervisor' || p === 'team leader' || p === 'act.team leader';
  };
  const leaders = (employees || []).filter(isLeader);
  const leaderMap = indexBy_(leaders, 'Employee ID');

  return dates.reduce(function(all, date) {
    depts.forEach(function(dept) {
      const staffedLeaders = schedule.filter(function(row) {
        return row.date === date && row.code === shiftCode && String(row.department || '') === dept && leaderMap[row.employeeId];
      }).length;
      if (staffedLeaders < rule.value) {
        all.push(violation_('warn', rule, 'แผนก ' + dept + ' · ขาดหัวหน้ากะ (Leader) ในกะ ' + shiftCode,
          date + ' กะ ' + shiftCode + ' — มี Leader ' + staffedLeaders + ' คน (ขั้นต่ำ ' + rule.value + ' คน)',
          { date: date, department: dept }));
      }
    });
    return all;
  }, []);
}

function checkLeaveConflicts_(rule, schedule) {
  const leaveKeys = {};
  schedule.forEach(function(row) {
    if (row.code === 'AL' || /ลา|leave/i.test(row.remark)) {
      leaveKeys[row.employeeId + '|' + row.date] = true;
    }
  });

  return schedule.filter(function(row) {
    return row.hours > 0 && leaveKeys[row.employeeId + '|' + row.date];
  }).map(function(row) {
    return violation_('error', rule, row.employeeName || row.employeeId,
      row.date + ' — ถูกจัดกะ ' + row.code + ' ทั้งที่มีการลาในวันเดียวกัน',
      { employeeId: row.employeeId, date: row.date });
  });
}

function checkRemarkFlags_(rule, schedule) {
  return schedule.filter(function(row) {
    return row.remark && row.remark.toLowerCase().indexOf('violation') >= 0;
  }).map(function(row) {
    return violation_('warn', rule, row.employeeName || row.employeeId,
      row.date + ' — ' + row.remark, { employeeId: row.employeeId, date: row.date });
  });
}

function violation_(severity, rule, title, description, extra) {
  return Object.assign({
    severity: severity,
    ruleId: rule.id,
    ruleName: rule.name,
    title: title,
    description: description
  }, extra || {});
}

/* -------------------------------------------------------------- analytics */

function buildAnalytics_(schedule, employees, shiftTypes, rules, dates, overtime) {
  const workingTypes = shiftTypes.filter(function(type) { return type.Hours > 0; });
  const departments = unique_(employees.map(function(employee) { return String(employee.Department || ''); }));

  return {
    hoursByDepartment: {
      labels: departments,
      datasets: workingTypes.map(function(type) {
        return {
          label: type['Shift Name'] + ' (' + type['Shift Code'] + ')',
          color: type.Color,
          data: departments.map(function(department) {
            return round_(sumHours_(schedule, function(row) {
              return row.department === department && row.code === type['Shift Code'];
            }));
          })
        };
      }).concat([{
        label: 'OT',
        color: '#EF4444',
        data: departments.map(function(department) {
          return round_(overtime.byDepartment[department] || 0);
        })
      }])
    },
    shiftDistribution: shiftTypes.map(function(type) {
      const count = schedule.filter(function(row) { return row.code === type['Shift Code']; }).length;
      return {
        code: type['Shift Code'],
        label: type['Shift Name'] + ' (' + type['Shift Code'] + ')',
        color: type.Color,
        count: count,
        percent: schedule.length ? round_(count * 100 / schedule.length) : 0
      };
    }),
    otTrend: Object.keys(overtime.byWeek).sort().map(function(week) {
      return { label: 'W' + week.split('-W')[1], hours: round_(overtime.byWeek[week]) };
    }),
    monthly: unique_(dates.map(function(date) { return date.slice(0, 7); })).sort().map(function(month) {
      return {
        label: THAI_MONTHS[Number(month.slice(5, 7)) - 1] || month,
        hours: round_(sumHours_(schedule, function(row) { return row.date.slice(0, 7) === month; })),
        ot: round_(overtime.byMonth[month] || 0)
      };
    }),
    coverage: buildCoverage_(schedule, shiftTypes, rules, dates)
  };
}

function buildCoverage_(schedule, shiftTypes, rules, dates) {
  const dayCount = dates.length || 1;
  const coverageByCode = {};
  Object.keys(COVERAGE_RULES).forEach(function(ruleId) {
    const rule = rules.find(function(item) { return item.id === ruleId && item.enabled; });
    if (rule) coverageByCode[COVERAGE_RULES[ruleId].code] = { required: rule.value, department: COVERAGE_RULES[ruleId].department };
  });

  return shiftTypes.map(function(type) {
    const code = type['Shift Code'];
    const target = coverageByCode[code];
    const rows = schedule.filter(function(row) {
      return row.code === code && (!target || row.department === target.department);
    });

    return {
      code: code,
      name: type['Shift Name'] + ' (' + code + ')',
      color: type.Color,
      scope: target ? target.department : 'ทุกแผนก',
      actual: round_(rows.length / dayCount),
      required: target ? target.required : 0
    };
  });
}

function sumHours_(schedule, predicate) {
  return schedule.reduce(function(sum, row) {
    return predicate(row) ? sum + row.hours : sum;
  }, 0);
}

/* ------------------------------------------------------------ sheet setup */

function ensureApplicationBranding_(spreadsheet) {
  const currentName = String(spreadsheet.getName() || '').trim();
  if (!currentName || /shiftflow/i.test(currentName)) spreadsheet.rename(APP_TITLE);

  const settings = spreadsheet.getSheetByName('Settings');
  if (!settings) return;
  const rowCount = Math.max(settings.getLastRow(), 1);
  const values = settings.getRange(1, 1, rowCount, Math.max(settings.getLastColumn(), 2)).getValues();
  let appNameRow = 0;
  for (let index = 0; index < values.length; index++) {
    if (String(values[index][0] || '').trim().toLowerCase() !== 'app name') continue;
    appNameRow = index + 1;
    break;
  }
  if (appNameRow) settings.getRange(appNameRow, 2).setValue(APP_TITLE);
  else settings.appendRow(['App Name', APP_TITLE, 'ชื่อระบบ']);
}

function getOrCreateSpreadsheet_() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const existingId = scriptProperties.getProperty(PROPERTY_SPREADSHEET_ID);

  if (existingId) {
    for (let attempt = 0; attempt < 3; attempt++) {
      try {
        return SpreadsheetApp.openById(existingId);
      } catch (error) {
        if (attempt < 2) Utilities.sleep(250 * (attempt + 1));
      }
    }
    // Never silently switch the live system to a newly-created spreadsheet.
    // A temporary Spreadsheet service failure used to delete this binding,
    // which could make existing employees/licenses appear to have vanished.
    throw new Error('ไม่สามารถเปิด Spreadsheet หลักของระบบได้ชั่วคราว กรุณาลองใหม่อีกครั้ง (Spreadsheet ID: ' + existingId + ')');
  }

  const spreadsheet = SpreadsheetApp.create(APP_TITLE);
  scriptProperties.setProperty(PROPERTY_SPREADSHEET_ID, spreadsheet.getId());
  return spreadsheet;
}

/**
 * Adds missing sheets without touching the ones that are already there.
 * Returns the names it had to create.
 */
function ensureRequiredSheets_(spreadsheet) {
  const created = [];

  SHEET_BUILDERS_.forEach(function(builder) {
    if (spreadsheet.getSheetByName(builder.name)) return;
    const sheet = spreadsheet.insertSheet(builder.name);
    builder.seed(sheet, spreadsheet);
    created.push(builder.name);
  });

  const leftover = spreadsheet.getSheetByName('Sheet1');
  if (leftover && spreadsheet.getSheets().length > REQUIRED_SHEETS.length) {
    spreadsheet.deleteSheet(leftover);
  }

  if (created.length) SpreadsheetApp.flush();
  return created;
}

function prepareSheet_(spreadsheet, sheetName) {
  const sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);
  sheet.clear({ contentsOnly: false });
  return sheet;
}

function seedEmployees_(sheet) {
  writeSheet_(sheet, [
    ['Employee ID', 'Name', 'Department', 'Position', 'Skill', 'Status'],
    ['EMP001', 'สมชาย ใจดี', 'QA', 'Inspector', 'QC', 'Active'],
    ['EMP002', 'อารีย์ รักงาน', 'Production', 'Line 2', 'Machine A', 'Active'],
    ['EMP003', 'วิชัย แข็งแรง', 'Warehouse', 'Forklift', 'Inventory', 'Active'],
    ['EMP004', 'ศิริวรรณ ยิ้มเก่ง', 'QA', 'Lab Tech', 'QC', 'Active'],
    ['EMP005', 'อนันต์ ตั้งใจ', 'Maintenance', 'Elec', 'Electrical', 'Active'],
    ['EMP006', 'ประนอม สู้งาน', 'Production', 'Line 1', 'Machine B', 'Active']
  ]);
}

function seedEmployeeLicenses_(sheet) {
  writeSheet_(sheet, [[
    'License ID', 'Employee ID', 'License Type', 'License Number',
    'Issue Date', 'Expiry Date', 'Status', 'Document URL', 'Remark',
    'Updated By', 'Updated At'
  ]]);
  sheet.getRange(2, 5, Math.max(1, sheet.getMaxRows() - 1), 2).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(2, 11, Math.max(1, sheet.getMaxRows() - 1), 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
}

function seedLicenseAuditLog_(sheet) {
  writeSheet_(sheet, [[
    'Timestamp', 'Action', 'Employee ID', 'License ID', 'Work Date',
    'Shift Code', 'License Status', 'Expiry Date', 'Reason', 'Approved By'
  ]]);
  sheet.getRange(2, 1, Math.max(1, sheet.getMaxRows() - 1), 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sheet.getRange(2, 5, Math.max(1, sheet.getMaxRows() - 1), 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(2, 8, Math.max(1, sheet.getMaxRows() - 1), 1).setNumberFormat('yyyy-mm-dd');
}

function seedShiftTypes_(sheet) {
  const rows = [
    ['Shift Code', 'Shift Name', 'Start Time', 'End Time', 'Hours', 'Color'],
    ['D', 'Day Shift', '08:00', '17:00', 9, '#10B981'],
    ['N', 'Night Shift', '00:00', '08:00', 8, '#8B5CF6'],
    ['OFF', 'Day Off', '', '', 0, '#F87171'],
    ['AL', 'Annual Leave', '', '', 0, '#60A5FA']
  ];

  // Sheets would otherwise coerce '08:00' into a time value and hand it back as a Date.
  sheet.getRange(2, 3, rows.length - 1, 2).setNumberFormat('@');
  writeSheet_(sheet, rows);
}

function seedSchedule_(sheet) {
  const rows = [
    ['Date', 'Employee ID', 'Employee Name', 'Department', 'Shift Code', 'Start Time', 'End Time', 'Hours', 'Remark']
  ];
  const employees = [
    ['EMP001', 'สมชาย ใจดี', 'QA'],
    ['EMP002', 'อารีย์ รักงาน', 'Production'],
    ['EMP003', 'วิชัย แข็งแรง', 'Warehouse'],
    ['EMP004', 'ศิริวรรณ ยิ้มเก่ง', 'QA'],
    ['EMP005', 'อนันต์ ตั้งใจ', 'Maintenance'],
    ['EMP006', 'ประนอม สู้งาน', 'Production']
  ];
  const pattern = [
    ['D', 'OFF', 'N', 'D', 'OFF', 'OFF', 'AL'],
    ['OFF', 'N', 'D', 'OFF', 'N', 'OFF', 'OFF'],
    ['N', 'D', 'OFF', 'N', 'D', 'OFF', 'OFF'],
    ['D', 'D', 'AL', 'OFF', 'N', 'OFF', 'OFF'],
    ['OFF', 'OFF', 'D', 'N', 'D', 'OFF', 'OFF'],
    ['N', 'OFF', 'D', 'OFF', 'N', 'D', 'OFF']
  ];

  employees.forEach(function(employee, employeeIndex) {
    pattern[employeeIndex].forEach(function(code, dayIndex) {
      const shift = getShiftDetail_(code);
      rows.push([
        new Date(new Date().getFullYear(), 6, dayIndex + 1),
        employee[0],
        employee[1],
        employee[2],
        code,
        shift.start,
        shift.end,
        shift.hours,
        code === 'AL' ? 'Annual leave' : ''
      ]);
    });
  });

  sheet.getRange(2, 1, rows.length - 1, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(2, 6, rows.length - 1, 2).setNumberFormat('@');
  writeSheet_(sheet, rows);
}

function seedScheduleApprovals_(sheet) {
  writeSheet_(sheet, [[
    'Month', 'Status', 'Revision', 'Changed By', 'Changed At', 'Change Type',
    'Approved By', 'Approved At', 'Approval Note', 'Schedule Hash'
  ]]);
}

function seedScheduleApprovalLog_(sheet) {
  writeSheet_(sheet, [[
    'Timestamp', 'Action', 'Month', 'Revision', 'Status', 'Change Type', 'Performed By', 'Note'
  ]]);
}

function seedRules_(sheet) {
  writeSheet_(sheet, [
    ['Rule ID', 'Rule Name', 'Value', 'Unit', 'Enabled'],
    ['RULE001', 'Maximum working hours per week', 72, 'hours', true],
    ['RULE002', 'Minimum rest after night shift', 12, 'hours', true],
    ['RULE003', 'Minimum day shift coverage per department', 1, 'people', true],
    ['RULE004', 'Minimum night shift coverage per department', 1, 'people', true],
    ['RULE005', 'Do not schedule on approved leave', 1, 'boolean', true],
    ['RULE006', 'Supervisor day shift only and Sunday off', 1, 'boolean', true],
    ['RULE007', 'PO11 and WCS need Leader in Day shift', 1, 'people', true],
    ['RULE008', 'PO11 and WCS need Leader in Night shift', 1, 'people', true],
    ['RULE009', 'Spare department is excluded from auto scheduling', 1, 'boolean', true]
  ]);
}

function seedDashboard_(sheet, spreadsheet) {
  writeSheet_(sheet, [
    ['Metric', 'Value'],
    ['Total Employees', '=COUNTA(Employees!A2:A)'],
    ['Active Employees', '=COUNTIF(Employees!F2:F,"Active")'],
    ['Scheduled Rows', '=COUNTA(Schedule!A2:A)'],
    ['Total Hours', '=SUM(Schedule!H2:H)'],
    ['Spreadsheet URL', spreadsheet.getUrl()]
  ]);
}

function seedUsers_(sheet) {
  const now = new Date();
  writeSheet_(sheet, [
    ['User ID', 'Employee ID', 'Name', 'Email', 'Role', 'Department', 'Status', 'Password Hash', 'Requested At', 'Approved By', 'Approved At', 'Rejection Reason', 'Updated At', 'Last Login At'],
    ['USR-PRIMARY', '', 'Sermpong Admin', PRIMARY_ADMIN_EMAIL, 'Admin', 'All', 'Active', PRIMARY_ADMIN_PASSWORD_HASH, now, 'SYSTEM', now, '', now, '']
  ]);
  sheet.getRange(2, 8, 1, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sheet.getRange(2, 10, 1, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sheet.getRange(2, 12, 1, 2).setNumberFormat('yyyy-mm-dd hh:mm:ss');
}

function seedUserAuditLog_(sheet) {
  writeSheet_(sheet, [[
    'Timestamp', 'Action', 'User ID', 'Email', 'Role', 'Department', 'Reason', 'Performed By'
  ]]);
}

function seedSettings_(sheet, spreadsheet) {
  writeSheet_(sheet, [
    ['Key', 'Value', 'Description'],
    ['App Name', APP_TITLE, 'ชื่อระบบ'],
    ['Timezone', spreadsheet.getSpreadsheetTimeZone(), 'เขตเวลาที่ใช้ในระบบ'],
    ['Default Week Start', 'Monday', 'วันเริ่มต้นสัปดาห์'],
    ['Spreadsheet URL', spreadsheet.getUrl(), 'ไฟล์ข้อมูลหลัก'],
    ['Last Setup', new Date(), 'เวลาล่าสุดที่สร้าง/อัปเดต Sheet']
  ]);
  sheet.getRange(6, 2).setNumberFormat('yyyy-mm-dd hh:mm:ss');
}

const SHEET_BUILDERS_ = [
  { name: 'Employees', seed: seedEmployees_ },
  { name: 'Employee Licenses', seed: seedEmployeeLicenses_ },
  { name: 'Shift Types', seed: seedShiftTypes_ },
  { name: 'Schedule', seed: seedSchedule_ },
  { name: 'Schedule Approvals', seed: seedScheduleApprovals_ },
  { name: 'Schedule Approval Log', seed: seedScheduleApprovalLog_ },
  { name: 'Rules', seed: seedRules_ },
  { name: 'Dashboard', seed: seedDashboard_ },
  { name: 'Users', seed: seedUsers_ },
  { name: 'Settings', seed: seedSettings_ },
  { name: 'License Audit Log', seed: seedLicenseAuditLog_ },
  { name: 'User Audit Log', seed: seedUserAuditLog_ }
];

function writeSheet_(sheet, rows) {
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.setFrozenRows(1);
}

/* ----------------------------------------------------------------- values */

function isDate_(value) {
  return Object.prototype.toString.call(value) === '[object Date]';
}

function formatDate_(value, timeZone) {
  if (isDate_(value)) return Utilities.formatDate(value, timeZone, 'yyyy-MM-dd');
  return String(value || '').trim();
}

/**
 * Time-only cells come back as Date objects anchored to 1899-12-30.
 */
function formatDateTime_(value, timeZone) {
  if (!value) return '';
  if (value instanceof Date) return Utilities.formatDate(value, timeZone, 'yyyy-MM-dd HH:mm:ss');
  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? String(value) : Utilities.formatDate(parsed, timeZone, 'yyyy-MM-dd HH:mm:ss');
}

function formatTime_(value, timeZone) {
  if (value === '' || value === null || value === undefined) return '';
  if (isDate_(value)) return Utilities.formatDate(value, timeZone, 'HH:mm');
  const match = String(value).trim().match(/^(\d{1,2}):(\d{2})/);
  return match ? ('0' + match[1]).slice(-2) + ':' + match[2] : String(value).trim();
}

function normalizeRows_(rows, timeZone) {
  return rows.map(function(row) {
    return Object.keys(row).reduce(function(normalized, key) {
      const value = row[key];
      if (!isDate_(value)) {
        normalized[key] = value;
      } else if (value.getFullYear() < 1900) {
        normalized[key] = Utilities.formatDate(value, timeZone, 'HH:mm');
      } else {
        normalized[key] = Utilities.formatDate(value, timeZone, 'yyyy-MM-dd HH:mm:ss');
      }
      return normalized;
    }, {});
  });
}

function timeToMinutes_(value) {
  const match = String(value || '').match(/^(\d{1,2}):(\d{2})/);
  return match ? Number(match[1]) * 60 + Number(match[2]) : null;
}

/**
 * Absolute start/end of a shift in epoch millis, so an evening shift that runs
 * past midnight still compares correctly against the next day's shift.
 */
function shiftBounds_(row) {
  const base = Date.parse(row.date + 'T00:00:00Z');
  const startMinutes = timeToMinutes_(row.startTime);
  if (isNaN(base) || startMinutes === null) return {};
  const start = base + startMinutes * 60000;
  return { start: start, end: start + row.hours * 3600000 };
}

function isoWeek_(dateText) {
  const parts = String(dateText).split('-');
  const date = new Date(Date.UTC(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2])));
  const day = date.getUTCDay() || 7;
  date.setUTCDate(date.getUTCDate() + 4 - day);
  const yearStart = new Date(Date.UTC(date.getUTCFullYear(), 0, 1));
  const week = Math.ceil(((date - yearStart) / 86400000 + 1) / 7);
  return date.getUTCFullYear() + '-W' + ('0' + week).slice(-2);
}

function round_(value) {
  return Math.round((Number(value) || 0) * 10) / 10;
}

function unique_(values) {
  return values.filter(function(value, index) {
    return value && values.indexOf(value) === index;
  });
}

function indexBy_(rows, key) {
  return rows.reduce(function(map, row) {
    map[row[key]] = row;
    return map;
  }, {});
}

function groupBy_(rows, key) {
  return rows.reduce(function(groups, row) {
    (groups[row[key]] = groups[row[key]] || []).push(row);
    return groups;
  }, {});
}

function countBy_(rows, key) {
  return rows.reduce(function(counts, row) {
    const value = row[key] || 'Other';
    counts[value] = (counts[value] || 0) + 1;
    return counts;
  }, {});
}
// ==========================================================
// 🎯 Leave Management Module
// ==========================================================

const LEAVE_SHEET_NAME = "ลางาน";
const QUOTA_SHEET_NAME = "Quota";
const LEAVE_FOLDER_ID = "1cBf4eHABC-A3zEPMMohd_0RMa7UdSaJ7"; 

function formatTemplate_(template, placeholders) {
  let text = template;
  for (const [key, val] of Object.entries(placeholders)) {
    text = text.replace(new RegExp('{' + key + '}', 'g'), val || '');
  }
  return text;
}

function getLineSettings(token) {
  const user = validateSession(token);
  if (!isAdminOrManager_(user)) throw new Error('Unauthorized');
  
  const props = PropertiesService.getScriptProperties();
  const rawToken = props.getProperty('LINE_ACCESS_TOKEN') || '';
  
  let maskedToken = '';
  if (rawToken) {
    maskedToken = rawToken.length > 10 
      ? rawToken.slice(0, 5) + '••••••••' + rawToken.slice(-5)
      : '••••••••';
  }
  
  const defaultTemplateNew = `🔔 [คำขอลางานใหม่] รอตรวจรับเอกสาร\n--------------------------------\n👤 พนักงาน: {Name}\n📍 แผนก/พื้นที่: {Department}\n📋 ประเภท: {Type} ({Days} วัน)\n📅 วันที่: {StartDate} ถึง {EndDate}\n📝 เหตุผล: {Reason}\n📎 ไฟล์แนบ: {FileUrl}\n--------------------------------\n⚙️ จัดการใบลาคลิกที่ระบบ Security Management System`;
  const defaultTemplateUpdate = `📢 [อัปเดตสถานะใบลาจากระบบ]\n--------------------------------\n👤 พนักงาน: {Name}\n📋 ประเภท: {Type} ({Days} วัน)\n🔄 ผลการตรวจรับ: {Status}`;

  return {
    lineToken: maskedToken,
    hasToken: !!rawToken,
    groupId: props.getProperty('LINE_SUPERVISOR_GROUP_ID') || '',
    templateNewLeave: props.getProperty('LINE_TEMPLATE_NEW_LEAVE') || defaultTemplateNew,
    templateUpdateLeave: props.getProperty('LINE_TEMPLATE_UPDATE_LEAVE') || defaultTemplateUpdate
  };
}

function saveLineSettings(token, settings) {
  const user = validateSession(token);
  if (!isAdminOrManager_(user)) throw new Error('Unauthorized');
  
  const props = PropertiesService.getScriptProperties();
  const updates = {};
  
  if (settings.lineToken && !settings.lineToken.includes('••••')) {
    updates['LINE_ACCESS_TOKEN'] = settings.lineToken.trim();
  }
  
  updates['LINE_SUPERVISOR_GROUP_ID'] = (settings.groupId || '').trim();
  updates['LINE_TEMPLATE_NEW_LEAVE'] = settings.templateNewLeave || '';
  updates['LINE_TEMPLATE_UPDATE_LEAVE'] = settings.templateUpdateLeave || '';
  
  props.setProperties(updates);
  return { success: true };
}

function getLineCredentials_() {
  const props = PropertiesService.getScriptProperties();
  return {
    token: props.getProperty('LINE_ACCESS_TOKEN') || "aeknzmzcf4xb/Ecz5jT2w5ke5iBiIrTeWrZYVDC5kxhI/CiTOBcSncEb1hyemd0r1/mGTPps12IhuTti8zc746gGc9KflZz6VCm9qyX7pOTP9aBrRnNK4yIeTUkjPP6IlHLhXL7UDU9PDasndVsjSAdB04t89/1O/w1cDnyilFU=",
    groupId: props.getProperty('LINE_SUPERVISOR_GROUP_ID') || "C446bbc14cdd47a551210e62726bcf27a"
  };
}

function getAllLeaveHistory_(token) {
  const user = validateSession(token);
  if (!isAdminOrManager_(user)) throw new Error('Unauthorized');
  
  const spreadsheet = getOrCreateSpreadsheet_();
  const tz = spreadsheet.getSpreadsheetTimeZone() || "Asia/Bangkok";
  const sheets = ensureLeaveSheets_(spreadsheet);
  const leaveData = sheets.leaveSheet.getDataRange().getValues();
  let list = [];
  
  for (let j = leaveData.length - 1; j >= 1; j--) {
    const empName = String(leaveData[j][1]).trim();
    const type = leaveData[j][3];
    const days = Number(leaveData[j][6]) || 0;
    const status = String(leaveData[j][9]).trim() || "รออนุมัติ";
    
    const d1 = formatDate_(leaveData[j][4], tz);
    const d2 = formatDate_(leaveData[j][5], tz);
    const dateStr = (d1 === d2 || !d2) ? d1 : (d1 + " ถึง " + d2);

    list.push({
      id: j + 1,
      empName: empName,
      name: empName,
      date: dateStr,
      type: type,
      days: days,
      reason: leaveData[j][7] ? String(leaveData[j][7]) : "",
      status: status,
      fileUrl: leaveData[j][8] || ""
    });
  }
  return list;
}

function ensureLeaveSheets_(spreadsheet) {
  let leaveSheet = spreadsheet.getSheetByName(LEAVE_SHEET_NAME);
  if (!leaveSheet) {
    leaveSheet = spreadsheet.insertSheet(LEAVE_SHEET_NAME);
    leaveSheet.appendRow(["Timestamp", "ชื่อ-นามสกุล", "แผนก", "ประเภทการลา", "วันเริ่มต้น", "วันสิ้นสุด", "จำนวนวัน", "เหตุผล", "ไฟล์แนบ", "สถานะ"]);
    leaveSheet.getRange("A1:J1").setFontWeight("bold").setBackground("#e0e0e0");
    leaveSheet.setFrozenRows(1);
  }
  let quotaSheet = spreadsheet.getSheetByName(QUOTA_SHEET_NAME);
  if (!quotaSheet) {
    quotaSheet = spreadsheet.insertSheet(QUOTA_SHEET_NAME);
    quotaSheet.appendRow(["ชื่อพนักงาน", "ลาป่วย", "ลากิจ", "ลาพักร้อน"]);
    quotaSheet.getRange("A1:D1").setFontWeight("bold").setBackground("#e0e0e0");
    quotaSheet.setFrozenRows(1);
  }
  const cache = CacheService.getScriptCache();
  if (!cache.get('quota_clean_v2')) {
    if (quotaSheet && quotaSheet.getLastRow() > 1) {
      try {
        const data = quotaSheet.getRange(2, 1, quotaSheet.getLastRow() - 1, 1).getValues();
        let deletedAny = false;
        for (let i = data.length - 1; i >= 0; i--) {
          const val = String(data[i][0] || '').trim();
          if (val === 'พนักงาน' || val === '' || val === 'undefined') {
            quotaSheet.deleteRow(i + 2);
            deletedAny = true;
          }
        }
        if (!deletedAny) cache.put('quota_clean_v2', 'done', 600);
      } catch (e) {}
    }
  }
  return { leaveSheet, quotaSheet };
}

function verifyQuotaAvailability_(spreadsheet, employeeName, leaveType, requestedDays) {
  const quotaSheet = spreadsheet.getSheetByName(QUOTA_SHEET_NAME);
  const leaveSheet = spreadsheet.getSheetByName(LEAVE_SHEET_NAME);
  if (!quotaSheet || !leaveSheet) return { allowed: false, message: 'ไม่พบฐานข้อมูลวันลา' };

  const quotaData = quotaSheet.getDataRange().getValues();
  let userQuota = { sick: 0, personal: 0, vacation: 0 };
  let foundUser = false;

  for (let i = 1; i < quotaData.length; i++) {
    if (matchEmployeeNameFlexible_(quotaData[i][0], null, employeeName)) {
      userQuota.sick = Number(quotaData[i][1]) || 0;
      userQuota.personal = Number(quotaData[i][2]) || 0;
      userQuota.vacation = Number(quotaData[i][3]) || 0;
      foundUser = true;
      break;
    }
  }
  if (!foundUser) {
    userQuota = { sick: 30, personal: 6, vacation: 10 };
    try {
      if (employeeName && employeeName !== 'พนักงาน' && String(employeeName).trim() !== '' && String(employeeName).trim() !== 'undefined') {
        quotaSheet.appendRow([String(employeeName).trim(), 30, 6, 10]);
        foundUser = true;
      }
    } catch (e) {}
  }

  if (!foundUser) return { allowed: false, message: "ไม่พบชื่อพนักงานในระบบโควตาวันลา กรุณาติดต่อ Admin" };

  const leaveData = leaveSheet.getDataRange().getValues();
  let used = 0;
  
  for (let j = 1; j < leaveData.length; j++) {
    if (matchEmployeeNameFlexible_(leaveData[j][1], null, employeeName) && 
        String(leaveData[j][3]).trim() === leaveType && 
        String(leaveData[j][9]).trim() === "อนุมัติ") {
      used += (Number(leaveData[j][6]) || 0);
    }
  }

  let remaining = 0;
  if (leaveType === "ลาป่วย") remaining = userQuota.sick - used;
  else if (leaveType === "ลากิจ") remaining = userQuota.personal - used;
  else if (leaveType === "ลาพักร้อน") remaining = userQuota.vacation - used;
  else return { allowed: false, message: "ประเภทการลาไม่ถูกต้อง" };

  if (remaining < requestedDays) {
    return { allowed: false, message: `สิทธิ์วันลาคงเหลือไม่เพียงพอ (ต้องการลา ${requestedDays} วัน แต่คงเหลือเพียง ${remaining} วัน)` };
  }
  
  return { allowed: true };
}

function sendLineToGroup_(message) {
  const creds = getLineCredentials_();
  if (!creds.token || creds.token.includes('ACCESS_TOKEN')) return { success: false, error: 'ยังไม่ได้ตั้งค่า LINE Access Token' };

  let results = [];
  // Try LINE Bot Messaging API push if groupId is present
  if (creds.groupId && !creds.groupId.includes('GROUP_ID')) {
    const url = "https://api.line.me/v2/bot/message/push";
    const payload = {
      "to": creds.groupId,
      "messages": [{ "type": "text", "text": message }]
    };
    const options = {
      "method": "post",
      "headers": {
        "Content-Type": "application/json",
        "Authorization": "Bearer " + creds.token
      },
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    };
    try {
      const res = UrlFetchApp.fetch(url, options);
      const code = res.getResponseCode();
      const text = res.getContentText();
      results.push({ api: 'Bot Push v2', status: code, response: text });
      if (code === 200) return { success: true, results: results };
    } catch(e) {
      results.push({ api: 'Bot Push v2', error: e.toString() });
    }
  }

  // Fallback/Try LINE Notify API
  try {
    const notifyUrl = "https://notify-api.line.me/api/notify";
    const notifyOptions = {
      "method": "post",
      "headers": {
        "Authorization": "Bearer " + creds.token
      },
      "payload": { "message": message },
      "muteHttpExceptions": true
    };
    const res = UrlFetchApp.fetch(notifyUrl, notifyOptions);
    const code = res.getResponseCode();
    const text = res.getContentText();
    results.push({ api: 'LINE Notify', status: code, response: text });
    if (code === 200) return { success: true, results: results };
  } catch(e) {
    results.push({ api: 'LINE Notify', error: e.toString() });
  }

  return { success: false, results: results };
}

function submitLeaveRequest(token, leaveData) {
  const user = validateSession(token);
  if (!user) throw new Error('Unauthorized');

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const spreadsheet = getOrCreateSpreadsheet_();
    const sheets = ensureLeaveSheets_(spreadsheet);

    const empName = String(user.Name || user.name || user.FullName || '').trim();
    const empDept = String(user.Department || user.department || 'All').trim();
    if (!empName || empName === 'undefined') throw new Error('ไม่พบข้อมูลชื่อ-นามสกุลพนักงาน กรุณาออกจากระบบแล้วเข้าสู่ระบบใหม่');

    const start = new Date(leaveData.startDate);
    const end = new Date(leaveData.endDate);
    const days = Math.floor((end.getTime() - start.getTime()) / (1000 * 60 * 60 * 24)) + 1;

    const quotaCheck = verifyQuotaAvailability_(spreadsheet, empName, leaveData.leaveType, days);
    if (!quotaCheck.allowed) throw new Error(quotaCheck.message);

    let fileUrl = "ไม่มีไฟล์แนบ";
    if (leaveData.fileBase64 && leaveData.fileName) {
      try {
        const folder = DriveApp.getFolderById(LEAVE_FOLDER_ID);
        const blob = Utilities.newBlob(Utilities.base64Decode(leaveData.fileBase64), leaveData.mimeType, leaveData.fileName);
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        fileUrl = file.getUrl();
      } catch (e) {
        console.error("File upload error", e);
        fileUrl = "อัปโหลดไฟล์ล้มเหลว";
      }
    }

    sheets.leaveSheet.appendRow([
      new Date(), empName, empDept, leaveData.leaveType, 
      leaveData.startDate, leaveData.endDate, days, leaveData.reason, fileUrl, "รออนุมัติ"
    ]);

    const props = PropertiesService.getScriptProperties();
    const rawTemplate = props.getProperty('LINE_TEMPLATE_NEW_LEAVE');
    let msg = "";
    const placeholders = {
      Name: user.Name,
      Department: user.Department,
      Type: leaveData.leaveType,
      Days: days,
      StartDate: leaveData.startDate,
      EndDate: leaveData.endDate,
      Reason: leaveData.reason,
      FileUrl: fileUrl
    };
    
    if (rawTemplate) {
      msg = formatTemplate_(rawTemplate, placeholders);
    } else {
      msg = `🔔 [คำขอลางานใหม่] รอตรวจรับเอกสาร\n` +
            `--------------------------------\n` +
            `👤 พนักงาน: ${user.Name}\n` +
            `📍 แผนก/พื้นที่: ${user.Department}\n` +
            `📋 ประเภท: ${leaveData.leaveType} (${days} วัน)\n` +
            `📅 วันที่: ${leaveData.startDate} ถึง ${leaveData.endDate}\n` +
            `📝 เหตุผล: ${leaveData.reason}\n` +
            `📎 ไฟล์แนบ: ${fileUrl}\n` +
            `--------------------------------\n` +
            `⚙️ จัดการใบลาคลิกที่ระบบ Security Management System`;
    }
              
    sendLineToGroup_(msg);
    return { success: true, message: "ส่งใบลาสำเร็จ" };
  } finally {
    lock.releaseLock();
  }
}

function normalizeEmpName_(str) {
  return String(str || '').trim().replace(/^(นาย|นางสาว|นาง|คุณ|ด\.ช\.|ด\.ญ\.)/g, '').replace(/\s+/g, ' ').toLowerCase();
}

function matchEmployeeNameFlexible_(rowStr, userObj, targetStr) {
  const r = normalizeEmpName_(rowStr);
  if (!r) return false;
  const t = normalizeEmpName_(targetStr);
  const n = normalizeEmpName_(userObj ? userObj.Name : '');
  const fn = normalizeEmpName_(userObj ? userObj.FullName : '');
  const em = String(userObj ? userObj.Email : '').trim().toLowerCase();
  
  if (t && (r === t || r.includes(t) || t.includes(r))) return true;
  if (n && (r === n || r.includes(n) || n.includes(r))) return true;
  if (fn && (r === fn || r.includes(fn) || fn.includes(r))) return true;
  if (em && (r === em || r.includes(em))) return true;
  return false;
}

function getLeaveSummary(token, filterName) {
  const user = validateSession(token);
  if (!user) throw new Error('Unauthorized');
  
  const empName = String(user.Name || user.name || user.FullName || '').trim();
  const targetName = (isAdminOrManager_(user)) && filterName ? filterName : empName;

  const spreadsheet = getOrCreateSpreadsheet_();
  const sheets = ensureLeaveSheets_(spreadsheet);
  
  const quotaData = sheets.quotaSheet.getDataRange().getValues();
  let userQuota = { sick: 0, personal: 0, vacation: 0 };
  let foundUser = false;
  
  for (let i = 1; i < quotaData.length; i++) {
    if (matchEmployeeNameFlexible_(quotaData[i][0], user, targetName)) {
      userQuota.sick = Number(quotaData[i][1]) || 0;
      userQuota.personal = Number(quotaData[i][2]) || 0;
      userQuota.vacation = Number(quotaData[i][3]) || 0;
      foundUser = true;
      break;
    }
  }
  if (!foundUser) {
    const newName = targetName || empName;
    userQuota = { sick: 30, personal: 6, vacation: 10 };
    if (newName && newName !== 'พนักงาน' && String(newName).trim() !== '' && String(newName).trim() !== 'undefined') {
      try {
        if (isAdminOrManager_(user) || String(empName).toLowerCase() === String(targetName).toLowerCase()) {
          sheets.quotaSheet.appendRow([String(newName).trim(), 30, 6, 10]);
          foundUser = true;
        }
      } catch (e) {}
    }
  }

  const leaveData = sheets.leaveSheet.getDataRange().getValues(); 
  let used = { "ลาป่วย": 0, "ลากิจ": 0, "ลาพักร้อน": 0 };
  let history = [];
  let pendingList = [];
  let allHistory = [];
  const tz = spreadsheet.getSpreadsheetTimeZone() || "Asia/Bangkok";
  const isAdmOrMgr = isAdminOrManager_(user);
  
  for (let j = leaveData.length - 1; j >= 1; j--) {
    const eName = String(leaveData[j][1]).trim();
    const type = leaveData[j][3];
    const days = Number(leaveData[j][6]) || 0;
    const status = String(leaveData[j][9]).trim() || "รออนุมัติ";
    const d1 = formatDate_(leaveData[j][4], tz);
    const d2 = formatDate_(leaveData[j][5], tz);
    const dateStr = (d1 === d2 || !d2) ? d1 : (d1 + " ถึง " + d2);

    if (matchEmployeeNameFlexible_(leaveData[j][1], user, targetName)) {
      history.push({
        id: j + 1,
        date: dateStr,
        type: type,
        reason: leaveData[j][7] ? String(leaveData[j][7]) : "",
        status: status,
        days: days,
        fileUrl: leaveData[j][8]
      });

      if (status === "อนุมัติ" && used[type] !== undefined) {
        used[type] += days;
      }
    }

    if (isAdmOrMgr) {
      const item = {
        id: j + 1,
        name: eName,
        empName: eName,
        department: String(leaveData[j][2]).trim(),
        date: dateStr,
        startDate: d1,
        endDate: d2,
        type: type,
        days: days,
        reason: leaveData[j][7] ? String(leaveData[j][7]) : "",
        status: status,
        fileUrl: leaveData[j][8] || ""
      };
      allHistory.push(item);
      if (status === "รออนุมัติ") {
        pendingList.push(item);
      }
    }
  }

  return {
    quota: userQuota,
    used: used,
    remaining: {
      sick: userQuota.sick - used["ลาป่วย"],
      personal: userQuota.personal - used["ลากิจ"],
      vacation: userQuota.vacation - used["ลาพักร้อน"]
    },
    history: history,
    pendingList: pendingList,
    allHistory: allHistory
  };
}

function getPendingLeaves(token) {
  const user = validateSession(token);
  if (!isAdminOrManager_(user)) throw new Error('Unauthorized');

  const spreadsheet = getOrCreateSpreadsheet_();
  const tz = spreadsheet.getSpreadsheetTimeZone() || "Asia/Bangkok";
  const sheets = ensureLeaveSheets_(spreadsheet);
  const data = sheets.leaveSheet.getDataRange().getValues();
  
  let pendingList = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][9]).trim() === "รออนุมัติ") { 
      const d1 = formatDate_(data[i][4], tz);
      const d2 = formatDate_(data[i][5], tz);
      const dateStr = (d1 === d2 || !d2) ? d1 : (d1 + " ถึง " + d2);

      pendingList.push({
        id: i + 1, 
        name: String(data[i][1]).trim(),
        empName: String(data[i][1]).trim(),
        department: String(data[i][2]).trim(),
        type: data[i][3],
        startDate: d1,
        endDate: d2,
        date: dateStr,
        days: data[i][6],
        reason: data[i][7],
        fileUrl: data[i][8]
      });
    }
  }
  return pendingList.reverse();
}

function updateLeaveStatus(token, leaveId, status) {
  const user = validateSession(token);
  if (!isAdminOrManager_(user)) throw new Error('Unauthorized');

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const spreadsheet = getOrCreateSpreadsheet_();
    const sheets = ensureLeaveSheets_(spreadsheet);
    
    const rowNumber = Number(leaveId);
    const rowData = sheets.leaveSheet.getRange(rowNumber, 1, 1, 10).getValues()[0];
    const currentStatus = String(rowData[9]).trim();
    
    if (currentStatus === "อนุมัติ" || currentStatus === "ไม่อนุมัติ") {
      throw new Error("รายการนี้ได้รับการตรวจสอบไปแล้ว");
    }

    const empName = String(rowData[1]).trim();
    const leaveType = String(rowData[3]).trim();
    const leaveDays = Number(rowData[6]) || 0;
    const startDateStr = String(rowData[4]).trim();
    const endDateStr = String(rowData[5]).trim();

    if (String(user.Role||'').toLowerCase().includes('manager') || String(user.Role||'').includes('หัวหน้า')) {
      const today = new Date();
      today.setHours(0,0,0,0);
      const startDt = new Date(startDateStr);
      startDt.setHours(0,0,0,0);
      if (startDt < today) throw new Error('Manager ไม่สามารถอนุมัติลาย้อนหลังได้');
    }

    if (status === "อนุมัติ") {
      const quotaCheck = verifyQuotaAvailability_(spreadsheet, empName, leaveType, leaveDays);
      if (!quotaCheck.allowed) throw new Error("อนุมัติไม่สำเร็จ: " + quotaCheck.message);
      
      const start = new Date(startDateStr);
      const end = new Date(endDateStr);
      
      // Group dates by month for batch updating
      const datesByMonth = {};
      for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
        const year = d.getFullYear();
        const month = String(d.getMonth() + 1).padStart(2, '0');
        const dateStr = `${year}-${month}-${String(d.getDate()).padStart(2, '0')}`;
        const key = `${year}_${month}`;
        if (!datesByMonth[key]) datesByMonth[key] = [];
        datesByMonth[key].push(dateStr);
      }
      
      // Update each month sheet in batch
      for (const [monthKey, dates] of Object.entries(datesByMonth)) {
        const sheetName = `Schedule_${monthKey}`;
        const sheet = spreadsheet.getSheetByName(sheetName);
        if (!sheet) continue;
        
        const data = sheet.getDataRange().getValues();
        if (data.length < 3) continue;
        const headerDateRow = data[1];
        
        // Find target employee row
        let empRowIdx = -1;
        for (let r = 3; r < data.length; r++) {
          if (String(data[r][0]).trim().toLowerCase() === empName.toLowerCase()) {
            empRowIdx = r;
            break;
          }
        }
        
        if (empRowIdx !== -1) {
          let updated = false;
          for (const dateStr of dates) {
            let targetColIdx = -1;
            for (let c = 1; c < headerDateRow.length; c++) {
              if (String(headerDateRow[c]).trim() === dateStr) {
                targetColIdx = c;
                break;
              }
            }
            if (targetColIdx !== -1) {
              data[empRowIdx][targetColIdx] = 'AL';
              updated = true;
            }
          }
          if (updated) {
            sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
          }
        }
      }
    }

    sheets.leaveSheet.getRange(rowNumber, 10).setValue(status);

    const statusIcon = (status === "อนุมัติ") ? "✅ อนุมัติแล้ว" : "❌ ไม่อนุมัติ";
    const props = PropertiesService.getScriptProperties();
    const rawTemplate = props.getProperty('LINE_TEMPLATE_UPDATE_LEAVE');
    let updateMsg = "";
    const placeholders = {
      Name: empName,
      Type: leaveType,
      Days: leaveDays,
      Status: statusIcon
    };
    
    if (rawTemplate) {
      updateMsg = formatTemplate_(rawTemplate, placeholders);
    } else {
      updateMsg = "📢 [อัปเดตสถานะใบลาจากระบบ]\n" +
                  "--------------------------------\n" +
                  "👤 พนักงาน: " + empName + "\n" +
                  "📋 ประเภท: " + leaveType + " (" + leaveDays + " วัน)\n" +
                  "🔄 ผลการตรวจรับ: " + statusIcon;
    }
    try { sendLineToGroup_(updateMsg); } catch(e) {}

    return { success: true };
  } finally {
    lock.releaseLock();
  }
}


function updateEmployeeShiftInternal_(spreadsheet, dateStr, employeeName, shiftCode) {
  const [year, month] = dateStr.split('-');
  const sheetName = `Schedule_${year}_${month}`;
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) return;
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 3) return;

  const headerDateRow = data[1];

  let targetColIdx = -1;
  for (let c = 1; c < headerDateRow.length; c++) {
    if (String(headerDateRow[c]).trim() === dateStr) {
      targetColIdx = c;
      break;
    }
  }
  if (targetColIdx === -1) return;

  for (let r = 3; r < data.length; r++) {
    if (String(data[r][0]).trim().toLowerCase() === String(employeeName).trim().toLowerCase()) {
      sheet.getRange(r + 1, targetColIdx + 1).setValue(shiftCode);
      return;
    }
  }
}


function getDatabaseInfo(token) {
  const user = validateSession(token);
  if (!isAdminOrManager_(user)) throw new Error('Unauthorized');

  const spreadsheet = getOrCreateSpreadsheet_();
  const sheets = ensureLeaveSheets_(spreadsheet);
  
  const quotaData = sheets.quotaSheet.getDataRange().getValues();
  const leaveData = sheets.leaveSheet.getDataRange().getValues();
  
  return {
    spreadsheetName: spreadsheet.getName(),
    spreadsheetUrl: spreadsheet.getUrl(),
    spreadsheetId: spreadsheet.getId(),
    quotaRowsCount: quotaData.length - 1,
    quotaEmployees: quotaData.slice(1).map(r => ({ name: r[0], sick: r[1], personal: r[2], vacation: r[3] })),
    leaveRequestsCount: leaveData.length - 1,
    lastLeaveRequest: leaveData.length > 1 ? { name: leaveData[leaveData.length-1][1], type: leaveData[leaveData.length-1][3], status: leaveData[leaveData.length-1][9] } : null
  };
}

function testLineMessage(token) {
  const user = validateSession(token);
  if (!isAdminOrManager_(user)) throw new Error('Unauthorized');
  
  const empName = user.FullName || user.Name || user.Email || 'Admin';
  const testMsg = "🔔 [ทดสอบระบบ SMS] ข้อความนี้ถูกส่งจากหน้าจัดการการลาโดย: " + empName + " (" + new Date().toLocaleString('th-TH') + ")";
  const result = sendLineToGroup_(testMsg);
  return result;
}
