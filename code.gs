// ================================================================
//  ENDERUN ACADEMIC MANAGEMENT SYSTEM  |  Code.gs  (Backend)
//  Enhanced & Production-Ready — Full RBAC + End-to-End
// ================================================================

// ----------------------------------------------------------------
//  ROUTING  —  doGet
// ----------------------------------------------------------------
function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};

  // ── One-Click Quick Approve (email link) ──────────────────────
  if (params.action === 'quickApprove' && params.reqId && params.sheet && params.token) {
    try {
      var ss    = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(params.sheet);
      if (!sheet) return _html('Sheet not found.');
      var data  = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]) === params.reqId) {
          var dbToken = String(data[i][11] || '').trim(); // Column L (12th column) for AuthToken
          if (dbToken !== params.token) return _html('Invalid or expired approval token.');
          
          var cur    = String(data[i][7]).trim();
          var next   = (cur === 'Pending Program Head') ? 'Pending Dean' : 'Approved';
          
          updateRequestDetails({
            reqId: params.reqId,
            sheetName: params.sheet,
            status: next,
            remarks: 'Approved via Quick Link',
            logUser: 'Quick Approver'
          });

          return HtmlService.createHtmlOutput(
            '<body style="font-family:sans-serif;display:flex;align-items:center;justify-content:center;min-height:100vh;background:#0f0e0c;color:#f0ead8;">'
            + '<div style="text-align:center;padding:40px;background:#1c1a16;border:1px solid #2e2a22;border-radius:16px;max-width:440px;">'
            + '<div style="font-size:3rem;margin-bottom:1rem;">✅</div>'
            + '<h2 style="color:#c9a96e;margin-bottom:0.5rem;">Successfully Processed</h2>'
            + '<p>Request <strong>' + params.reqId + '</strong> has been updated to <strong>' + next + '</strong>.</p>'
            + '<p style="color:#5c5244;font-size:0.85rem;margin-top:1rem;">You may close this tab.</p>'
            + '</div></body>');
        }
      }
      return _html('Request ID not found.');
    } catch (err) { return _html('Error: ' + err); }
  }

  // ── Account Activation (email link) ──────────────────────────
  if (params.action === 'activate' && params.id) {
    try {
      var ss    = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName('Users');
      if (!sheet) return _html('Users sheet not found.');
      var data  = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (data[i][0] === params.id) {
          sheet.getRange(i + 1, 6).setValue('Active');
          var appUrl = ScriptApp.getService().getUrl() + '?view=admin';
          return HtmlService.createHtmlOutput(
            '<body style="font-family:sans-serif;display:flex;align-items:center;justify-content:center;min-height:100vh;background:#0f0e0c;color:#f0ead8;">'
            + '<div style="text-align:center;padding:40px;background:#1c1a16;border:1px solid #2e2a22;border-radius:16px;max-width:440px;">'
            + '<div style="font-size:3rem;margin-bottom:1rem;">🎓</div>'
            + '<h2 style="color:#c9a96e;margin-bottom:0.5rem;">Account Activated!</h2>'
            + '<p>Your AMS Portal account is now active. You can sign in using your email and temporary password.</p>'
            + '<a href="' + appUrl + '" style="display:inline-block;margin-top:1.5rem;background:#c9a96e;color:#0f0e0c;padding:12px 28px;border-radius:10px;font-weight:700;text-decoration:none;">Go to Login Portal</a>'
            + '</div></body>');
        }
      }
      return _html('User ID not found or activation link is invalid.');
    } catch (err) { return _html('Error: ' + err); }
  }

  // ── Default: serve the app ────────────────────────────────────
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Enderun AMS Portal')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function _html(msg) {
  return HtmlService.createHtmlOutput(
    '<body style="font-family:sans-serif;padding:2rem;background:#0f0e0c;color:#f0ead8;">' + msg + '</body>');
}


// ================================================================
//  AUTHENTICATION
// ================================================================
// Set your super-admin credentials here before deploying.
var ADMIN_EMAIL    = 'YOUR_ADMIN_EMAIL_HERE';
var ADMIN_PASSWORD = 'YOUR_ADMIN_PASSWORD_HERE';

function authenticateUser(email, pass) {
  // Hardcoded super-admin
  if (email === ADMIN_EMAIL && pass === ADMIN_PASSWORD) {
    return {
      success: true,
      user: { id: 'ADMIN', name: 'System Admin', email: email, role: 'Admin', form: 'N/A' },
      isFirstLogin: false
    };
  }

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Users');
  if (!sheet) return { success: false, message: 'System error: Users database not found.' };

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][2]).toLowerCase() === String(email).toLowerCase()) {
      if (String(data[i][5]) !== 'Active') {
        return { success: false, message: 'Account inactive. Please click the activation link sent to your email.' };
      }
      if (String(data[i][7]) === _hashPassword(String(pass))) {
        return {
          success: true,
          user: {
            id: String(data[i][0]), name: String(data[i][1]),
            email: String(data[i][2]), role: String(data[i][3]),
            form: String(data[i][4])
          },
          isFirstLogin: (String(data[i][8]) === 'Yes')
        };
      }
      return { success: false, message: 'Incorrect password. Please try again.' };
    }
  }
  return { success: false, message: 'Email address not found in the system.' };
}

function updateFirstLoginPassword(userId, newPass) {
  if (userId === 'ADMIN') return { success: true };
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === userId) {
      sheet.getRange(i + 1, 8).setValue(_hashPassword(newPass));
      sheet.getRange(i + 1, 9).setValue('No');
      return { success: true };
    }
  }
  return { success: false, message: 'User not found.' };
}

function processForgotPassword(email) {
  if (!email) return { success: false, message: 'Please provide an email address.' };
  if (email === ADMIN_EMAIL) return { success: false, message: 'Admin password cannot be reset here. Please check Code.gs.' };

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  var data  = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][2]).toLowerCase() === String(email).toLowerCase()) {
      if (String(data[i][5]) !== 'Active') {
        return { success: false, message: 'Account is inactive. Please contact the administrator.' };
      }
      
      var newTempPass = Math.random().toString(36).slice(-8);
      sheet.getRange(i + 1, 8).setValue(_hashPassword(newTempPass));
      sheet.getRange(i + 1, 9).setValue('Yes'); // Force user to change password on next login
      
      var subject = 'AMS Portal: Password Reset Request';
      var body = _emailTemplate('Password Reset', 
        '<p>Hello <strong>' + data[i][1] + '</strong>,</p>'
        + '<p>We received a request to reset your password for the AMS Portal.</p>'
        + '<div style="background:#161410;padding:20px;border-radius:10px;border:1px solid #2e2a22;margin:20px 0;text-align:center;">'
        + '<p style="margin:0;color:#7a6540;font-size:12px;text-transform:uppercase;letter-spacing:1px;">Your New Temporary Password</p>'
        + '<p style="margin:10px 0 0;font-size:26px;font-family:monospace;color:#c9a96e;font-weight:700;letter-spacing:3px;">' + newTempPass + '</p>'
        + '</div>'
        + '<p>Please log in using this temporary password. You will be prompted to create a new secure password immediately.</p>'
      );
      
      try {
        GmailApp.sendEmail(email, subject, '', { htmlBody: body });
        _log('System', 'Authentication', 'Password reset generated for ' + email);
        return { success: true, message: 'A temporary password has been sent to your email.' };
      } catch(e) {
        return { success: false, message: 'Error sending email. Please try again.' };
      }
    }
  }
  return { success: false, message: 'Email address not found in our system.' };
}


// ================================================================
//  USER MANAGEMENT
// ================================================================
function manageUserAccount(userData) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = _ensureSheet(ss, 'Users', ['User_ID','Name','Email','Role','Assigned_Form','Status','Date_Added','Password','IsFirstLogin']);

    // Check duplicate email
    var existing = sheet.getDataRange().getValues();
    for (var e = 1; e < existing.length; e++) {
      if (String(existing[e][2]).toLowerCase() === String(userData.email).toLowerCase()) {
        return 'Error: A user with this email already exists.';
      }
    }

    var userId    = 'USR-' + new Date().getTime().toString().slice(-6);
    var defaultPw = Math.random().toString(36).slice(-8);
    var date      = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy');

    sheet.appendRow([userId, userData.name, userData.email, userData.role,
                     userData.assignedForm, 'Inactive', date, _hashPassword(defaultPw), 'Yes']);

    var activationLink = ScriptApp.getService().getUrl() + '?action=activate&id=' + userId;
    var subject = 'AMS Portal: Activate Your Account';
    var body = _emailTemplate('Welcome to AMS Portal', 
      '<p>Hello <strong>' + userData.name + '</strong>,</p>'
      + '<p>An account has been created for you as <strong>' + userData.role + '</strong>.</p>'
      + (userData.role !== 'Admin' ? '<p>Assigned Form: <strong style="color:#c9a96e;">' + userData.assignedForm + '</strong></p>' : '')
      + '<div style="background:#161410;padding:20px;border-radius:10px;border:1px solid #2e2a22;margin:20px 0;text-align:center;">'
      + '<p style="margin:0;color:#7a6540;font-size:12px;text-transform:uppercase;letter-spacing:1px;">Your Temporary Password</p>'
      + '<p style="margin:10px 0 0;font-size:26px;font-family:monospace;color:#c9a96e;font-weight:700;letter-spacing:3px;">' + defaultPw + '</p>'
      + '</div>'
      + '<p>Please click the button below to activate your account. You will be prompted to change your password upon first login.</p>'
      + '<div style="text-align:center;margin-top:25px;">'
      + '<a href="' + activationLink + '" style="background:#c9a96e;color:#0f0e0c;padding:14px 28px;text-decoration:none;border-radius:10px;font-weight:700;font-size:15px;display:inline-block;">Activate My Account</a>'
      + '</div>'
    );
    GmailApp.sendEmail(userData.email, subject, '', { htmlBody: body });
    _log('Admin', 'User Management', 'Added new user: ' + userData.name + ' (' + userData.role + ', ' + userData.assignedForm + ')');
    return 'User added successfully. Activation email sent to ' + userData.email;
  } catch (err) {
    return 'Error adding user: ' + err.toString();
  }
}

function updateUserInSheet(userData) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
    var data  = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === userData.id) {
        sheet.getRange(i + 1, 2).setValue(userData.name);
        sheet.getRange(i + 1, 3).setValue(userData.email);
        sheet.getRange(i + 1, 4).setValue(userData.role);
        sheet.getRange(i + 1, 5).setValue(userData.assignedForm);
        _log('Admin', 'User Management', 'Updated user: ' + userData.name);
        return 'User updated successfully.';
      }
    }
    return 'Error: User not found.';
  } catch (err) { return 'Error: ' + err.toString(); }
}

function deleteUserInSheet(userId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === userId) {
      var userName = data[i][1];
      sheet.deleteRow(i + 1);
      _log('Admin', 'User Management', 'Deleted user: ' + userName + ' (ID: ' + userId + ')');
      return 'User deleted successfully.';
    }
  }
  return 'Error: User not found.';
}

function getUsersList() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (!sheet) return [];
  var data  = sheet.getDataRange().getValues();
  var users = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    users.push({
      id: String(data[i][0]), name: String(data[i][1]),
      email: String(data[i][2]), role: String(data[i][3]),
      assignedForm: String(data[i][4]), status: String(data[i][5])
    });
  }
  return users;
}

function updateUserProfile(profileData) {
  if (profileData.id === 'ADMIN') {
    return 'Notice: System admin profile cannot be edited here. Update credentials directly in Code.gs.';
  }
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === profileData.id) {
      sheet.getRange(i + 1, 2).setValue(profileData.name);
      sheet.getRange(i + 1, 3).setValue(profileData.email);
      if (profileData.pass && profileData.pass.length >= 6) {
        sheet.getRange(i + 1, 8).setValue(_hashPassword(profileData.pass));
      }
      
      if (profileData.role === 'Dept Head' && profileData.formName && profileData.formName !== 'N/A') {
        updateFormStatus(profileData.formName, profileData.formStatus);
      }
      
      _log(profileData.name, 'Profile Settings', 'User updated their profile information.');
      return 'Profile updated successfully!';
    }
  }
  return 'Error: Profile not found.';
}

function getFormStatuses() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Form_Settings');
  if (!sheet) return {};
  var data = sheet.getDataRange().getValues();
  var statuses = {};
  for (var i = 1; i < data.length; i++) {
    statuses[data[i][0]] = data[i][1];
  }
  return statuses;
}

function updateFormStatus(formName, status) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Form_Settings');
  if (!sheet) {
    sheet = ss.insertSheet('Form_Settings');
    sheet.appendRow(['Form_Name', 'Status']);
    sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#161410').setFontColor('#f0ead8');
  }
  var data = sheet.getDataRange().getValues();
  var found = false;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === formName) {
      sheet.getRange(i + 1, 2).setValue(status);
      found = true;
      break;
    }
  }
  if (!found) sheet.appendRow([formName, status]);
  return true;
}


// ================================================================
//  HELPER: Return all sheets (Row-level filtering handles RBAC now)
// ================================================================
function _getAssignedSheets(userRole, userForm) {
  return ['Form1_Data','Form2_Data','Form3_Data','Form4_Data','Form5_Data'];
}


// ================================================================
//  DASHBOARD STATS  (RBAC-filtered)
// ================================================================
function getDashboardStats(userRole, userForm) {
  var ss         = SpreadsheetApp.getActiveSpreadsheet();
  var allSheets  = ['Form1_Data','Form2_Data','Form3_Data','Form4_Data','Form5_Data'];
  var formLabels = ['DTR','Online Mod','Late Grades','Evaluation','Make-up'];
  var allowed    = _getAssignedSheets(userRole, userForm);
  
  var stats = { totalRequests: 0, inProgress: 0, progressed: 0, rejected: 0 };
  var chartData = { barLabels: [], barData: [], lineLabels: [], lineData: [] };
  var dateCounts = {};

  // Build filtered bar chart labels/data
  var barLabels = [];
  var barData   = [];

  allSheets.forEach(function(sName, idx) {
    if (allowed.indexOf(sName) === -1) return; // Skip sheets not assigned
    
    barLabels.push(formLabels[idx]);
    var approvedCount = 0;
    
    var s = ss.getSheetByName(sName);
    if (!s) { barData.push(0); return; }
    
    var rows = s.getDataRange().getValues();
    var role = String(userRole || '').trim();
    for (var i = 1; i < rows.length; i++) {
      if (!rows[i][0] || String(rows[i][0]).toLowerCase().includes('request_id')) continue;
      
      // Row-level RBAC: Dept Head only sees their department
      if (role === 'Dept Head' && String(rows[i][2]).trim() !== String(userForm).trim()) continue;
      
      stats.totalRequests++;
      var status = String(rows[i][7]).trim();
      if (status === 'Approved') { stats.progressed++; approvedCount++; }
      else if (status === 'Rejected') stats.rejected++;
      else stats.inProgress++;

      var raw = rows[i][8];
      var ds  = (raw instanceof Date)
        ? Utilities.formatDate(raw, Session.getScriptTimeZone(), 'MMM dd')
        : String(raw).substring(0, 6);
      dateCounts[ds] = (dateCounts[ds] || 0) + 1;
    }
    barData.push(approvedCount);
  });

  chartData.barLabels = barLabels;
  chartData.barData   = barData;

  var sortedDates = Object.keys(dateCounts).slice(-7);
  chartData.lineLabels = sortedDates;
  chartData.lineData   = sortedDates.map(function(d) { return dateCounts[d]; });
  stats.chartData = chartData;
  return stats;
}


// ================================================================
//  OTP VERIFICATION (Anti-Impersonation)
// ================================================================
function sendFormOtp(email, name) {
  try {
    var otp = Math.floor(100000 + Math.random() * 900000).toString();
    var cache = CacheService.getScriptCache();
    
    // Store OTP in cache for 10 minutes (600 seconds)
    cache.put('OTP_' + String(email).toLowerCase(), otp, 600);
    
    var subject = 'AMS Portal: Verification Code';
    var body = _emailTemplate('Identity Verification', 
      '<p>Hello <strong>' + name + '</strong>,</p>'
      + '<p>You are attempting to submit an application to the AMS Portal. Please use the verification code below to confirm your identity.</p>'
      + '<div style="background:#161410;padding:20px;border-radius:10px;border:1px solid #2e2a22;margin:20px 0;text-align:center;">'
      + '<p style="margin:0;color:#7a6540;font-size:12px;text-transform:uppercase;letter-spacing:1px;">Your 6-Digit Code</p>'
      + '<p style="margin:10px 0 0;font-size:32px;font-family:monospace;color:#c9a96e;font-weight:700;letter-spacing:5px;">' + otp + '</p>'
      + '</div>'
      + '<p>This code will expire in 10 minutes. If you did not request this, please ignore this email.</p>'
    );
    
    GmailApp.sendEmail(email, subject, '', { htmlBody: body });
    return { success: true };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function verifyFormOtp(email, inputOtp) {
  var cache = CacheService.getScriptCache();
  var savedOtp = cache.get('OTP_' + String(email).toLowerCase());
  
  if (!savedOtp) return { success: false, message: 'OTP expired or not found. Please resend the code.' };
  if (String(savedOtp) !== String(inputOtp).trim()) return { success: false, message: 'Incorrect verification code.' };
  
  // Clear OTP after successful use
  cache.remove('OTP_' + String(email).toLowerCase());
  return { success: true };
}

// ================================================================
//  REQUEST SUBMISSION
// ================================================================
function submitApplicationForm(formData) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = _ensureFormSheet(ss, formData.sheetName);

    var year         = new Date().getFullYear();
    var lastRow      = Math.max(sheet.getLastRow(), 1);
    var randomSuffix = Math.random().toString(36).substring(2, 6).toUpperCase();
    var reqId        = formData.prefix + '-' + year + '-' + String(lastRow).padStart(4,'0') + '-' + randomSuffix;
    var dateStr      = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy');
    var authToken    = Utilities.getUuid().replace(/-/g, '').substring(0, 12);

    // Handle file attachment
    var fileUrl = 'No Attachment';
    if (formData.fileData && formData.fileData !== '') {
      try {
        var blob   = Utilities.newBlob(Utilities.base64Decode(formData.fileData), formData.fileMime, formData.fileName);
        var iter   = DriveApp.getFoldersByName('AMS_Attachments');
        var folder = iter.hasNext() ? iter.next() : DriveApp.createFolder('AMS_Attachments');
        var file   = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        fileUrl = file.getUrl();
      } catch(fe) { fileUrl = 'Attachment error: ' + fe.toString(); }
    }

    sheet.appendRow([
      reqId, formData.name, formData.department,
      formData.field1 || '', formData.field2 || '', formData.field3 || '',
      formData.details || '', formData.status || 'Pending Program Head',
      dateStr, '', fileUrl, authToken
    ]);

    _log(formData.logUser || 'Requestor', 'Form Application', 'Submitted: ' + reqId + ' (' + formData.formName + ')');

    // Confirmation email to requestor
    var reqContent = 
      '<p>Hello <strong>' + formData.name + '</strong>,</p>'
      + '<p>Your application for <strong>' + formData.formName + '</strong> has been successfully submitted.</p>'
      + '<div style="background:#161410;padding:16px;border-radius:10px;border:1px solid #2e2a22;margin:20px 0;">'
      + '<p style="margin:4px 0;"><strong style="color:#7a6540;">Request ID:</strong> <span style="font-family:monospace;color:#c9a96e;font-size:16px;font-weight:700;">' + reqId + '</span></p>'
      + '<p style="margin:4px 0;"><strong style="color:#7a6540;">Status:</strong> Pending Program Head Review</p>'
      + '<p style="margin:4px 0;"><strong style="color:#7a6540;">Submitted:</strong> ' + dateStr + '</p>'
      + '</div>'
      + '<p>Save your Request ID to track the status of your application through the AMS Portal.</p>'
      + '<p>You will receive an email notification once a decision is made.</p>';
    
    try {
      GmailApp.sendEmail(formData.email, 'AMS Portal: Application Submitted (' + reqId + ')', '',
                         { htmlBody: _emailTemplate('Application Submitted', reqContent) });
    } catch(mailErr) {
      _log('System', 'Email', 'Requestor confirmation email failed for ' + reqId + ': ' + mailErr);
    }

    // Notify the assigned department head(s)
    var usersSheet = ss.getSheetByName('Users');
    if (usersSheet) {
      var users = usersSheet.getDataRange().getValues();
      for (var u = 1; u < users.length; u++) {
        if (users[u][5] !== 'Active') continue;
        
        // Route by Department match instead of Form Number
        var isAssignedHead = (users[u][3] === 'Dept Head' && String(users[u][4]).trim() === String(formData.department).trim());
        
        if (isAssignedHead) {
          var quickLink = ScriptApp.getService().getUrl() + '?action=quickApprove&reqId=' + reqId + '&sheet=' + formData.sheetName + '&token=' + authToken;
          var notifContent =
            '<p>Hello <strong>' + users[u][1] + '</strong>,</p>'
            + '<p>A new request has been submitted for <strong>' + formData.formName + '</strong> and requires your attention.</p>'
            + '<div style="background:#161410;padding:16px;border-radius:10px;border:1px solid #2e2a22;margin:20px 0;">'
            + '<p style="margin:4px 0;"><strong style="color:#7a6540;">Request ID:</strong> <span style="font-family:monospace;color:#c9a96e;">' + reqId + '</span></p>'
            + '<p style="margin:4px 0;"><strong style="color:#7a6540;">Submitted By:</strong> ' + formData.name + '</p>'
            + '<p style="margin:4px 0;"><strong style="color:#7a6540;">Department:</strong> ' + formData.department + '</p>'
            + '</div>'
            + '<p>Please log in to the AMS Portal to review this request, or use the quick action below:</p>'
            + '<div style="text-align:center;margin-top:20px;">'
            + '<a href="' + quickLink + '" style="background:#4caf7d;color:#fff;padding:12px 24px;text-decoration:none;border-radius:10px;font-weight:700;font-size:14px;display:inline-block;">Quick Approve →</a>'
            + '</div>';
          try {
            GmailApp.sendEmail(users[u][2], 'AMS Portal: New Request (' + reqId + ')', '',
                               { htmlBody: _emailTemplate('New Request Alert', notifContent) });
          } catch(me) { _log('System', 'Email', 'Notification email failed for ' + users[u][1]); }
        }
      }
    }

    return 'Application submitted successfully! Your Request ID is ' + reqId;
  } catch (err) {
    return 'Error submitting request: ' + err.toString();
  }
}


// ================================================================
//  FETCHING REQUESTS  (RBAC-filtered)
// ================================================================
function getInProgressRequests(userRole, userForm) {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var allowed = _getAssignedSheets(userRole, userForm);
  var role    = String(userRole || '').trim();
  var result  = [];

  allowed.forEach(function(sName) {
    var s = ss.getSheetByName(sName);
    if (!s) return;
    var rows = s.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (!rows[i][0] || String(rows[i][0]).toLowerCase().includes('request_id')) continue;
      
      // Row-level RBAC: Dept Head only sees their department
      if (role === 'Dept Head' && String(rows[i][2]).trim() !== String(userForm).trim()) continue;

      var status = String(rows[i][7]).trim();
      
      // Dept Head: only sees "Pending Program Head"
      // Admin: sees all non-final statuses
      var relevant;
      if (role === 'Dept Head') {
        relevant = (status === 'Pending Program Head');
      } else {
        relevant = (status === 'Pending Program Head' || status === 'Pending Dean' || status === 'In Progress');
      }
      if (!relevant) continue;
      
      result.push({
        id: String(rows[i][0]), 
        name: String(rows[i][1]),
        email: String(rows[i][2] || ''), 
        department: String(rows[i][2] || ''),
        status: status,
        date: _fmtDate(rows[i][8]),
        sheetName: sName
      });
    }
  });
  
  // Sort newest first
  result.sort(function(a, b) { return b.date.localeCompare(a.date); });
  return result;
}

function getProgressedRequests(userRole, userForm) {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var allowed = _getAssignedSheets(userRole, userForm);
  var result  = [];

  allowed.forEach(function(sName) {
    var s = ss.getSheetByName(sName);
    if (!s) return;
    var rows = s.getDataRange().getValues();
    var role = String(userRole || '').trim();
    for (var i = 1; i < rows.length; i++) {
      if (!rows[i][0] || String(rows[i][0]).toLowerCase().includes('request_id')) continue;
      
      // Row-level RBAC: Dept Head only sees their department
      if (role === 'Dept Head' && String(rows[i][2]).trim() !== String(userForm).trim()) continue;

      var status = String(rows[i][7]).trim();
      if (status !== 'Approved' && status !== 'Rejected') continue;
      result.push({
        id: String(rows[i][0]), name: String(rows[i][1]),
        status: status,
        date: _fmtDate(rows[i][8]),
        sheetName: sName
      });
    }
  });
  
  result.sort(function(a, b) { return b.date.localeCompare(a.date); });
  return result;
}

function getSingleRequestData(reqId, sheetName) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(reqId).trim()) {
      var row = data[i].map(function(c) {
        if (c instanceof Date) return Utilities.formatDate(c, Session.getScriptTimeZone(), 'MM/dd/yyyy');
        return String(c);
      });
      return { headers: headers, data: row };
    }
  }
  return null;
}

function publicTrackRequest(reqId) {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ['Form1_Data','Form2_Data','Form3_Data','Form4_Data','Form5_Data'];
  for (var s = 0; s < sheets.length; s++) {
    var sheet = ss.getSheetByName(sheets[s]);
    if (!sheet) continue;
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === String(reqId).trim()) {
        return {
          id:      String(rows[i][0]),
          name:    String(rows[i][1]),
          status:  String(rows[i][7]),
          date:    _fmtDate(rows[i][8]),
          remarks: String(rows[i][9] || '')
        };
      }
    }
  }
  return null;
}


// ================================================================
//  DECISION / APPROVAL
// ================================================================
function updateRequestDetails(reqData) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(reqData.sheetName);
    if (!sheet) return 'Error: Sheet not found.';
    var rows = sheet.getDataRange().getValues();

    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) !== reqData.reqId) continue;

      sheet.getRange(i + 1, 8).setValue(reqData.status);
      sheet.getRange(i + 1, 10).setValue(reqData.remarks || '');
      _log(reqData.logUser, 'Approval Action', 'Marked ' + reqData.reqId + ' as ' + reqData.status);

      var reqName  = String(rows[i][1]);
      var reqEmail = String(rows[i][2] || '');

      // Notify requestor on final decision
      if (reqData.status === 'Approved' || reqData.status === 'Rejected') {
        var color   = reqData.status === 'Approved' ? '#4caf7d' : '#e05b5b';
        var icon    = reqData.status === 'Approved' ? '✅' : '❌';
        var content = 
          '<p>Hello <strong>' + reqName + '</strong>,</p>'
          + '<p>Your request has received a final decision from the Academic Office.</p>'
          + '<div style="background:#161410;padding:20px;border-radius:10px;border:1px solid #2e2a22;margin:20px 0;text-align:center;">'
          + '<p style="margin:0;font-size:2rem;">' + icon + '</p>'
          + '<p style="margin:8px 0 0;font-size:22px;font-weight:700;color:' + color + ';">' + reqData.status.toUpperCase() + '</p>'
          + '<p style="margin:8px 0 0;color:#7a6540;font-size:13px;font-family:monospace;">' + reqData.reqId + '</p>'
          + '</div>'
          + (reqData.remarks ? '<p><strong>Remarks from Approver:</strong></p><p style="font-style:italic;color:#a89880;">"' + reqData.remarks + '"</p>' : '')
          + '<p>Please log in to the AMS Portal to download your official application record.</p>';
        try {
          if (reqEmail && reqEmail.includes('@')) {
            GmailApp.sendEmail(reqEmail, 'AMS Portal: Request ' + reqData.status + ' (' + reqData.reqId + ')', '',
                               { htmlBody: _emailTemplate('Request ' + reqData.status, content) });
          }
        } catch(me) { _log('System', 'Email', 'Email failed for ' + reqData.reqId + ': ' + me); }
        return 'Request ' + reqData.reqId + ' has been ' + reqData.status + '. Notification sent.';
      }

      // Endorsed to Dean
      if (reqData.status === 'Pending Dean') {
        var usersSheet = ss.getSheetByName('Users');
        if (usersSheet) {
          var users = usersSheet.getDataRange().getValues();
          for (var u = 1; u < users.length; u++) {
            if (users[u][3] === 'Admin' && users[u][5] === 'Active') {
              var quickLink = ScriptApp.getService().getUrl() + '?action=quickApprove&reqId=' + reqData.reqId + '&sheet=' + reqData.sheetName;
              var deanContent = 
                '<p>Hello <strong>' + users[u][1] + '</strong>,</p>'
                + '<p>A request has been endorsed by the Program Head and awaits your final decision.</p>'
                + '<div style="background:#161410;padding:16px;border-radius:10px;border:1px solid #2e2a22;margin:20px 0;">'
                + '<p style="margin:4px 0;"><strong style="color:#7a6540;">Request ID:</strong> <span style="font-family:monospace;color:#c9a96e;">' + reqData.reqId + '</span></p>'
                + '<p style="margin:4px 0;"><strong style="color:#7a6540;">Requestor:</strong> ' + reqName + '</p>'
                + '<p style="margin:4px 0;"><strong style="color:#7a6540;">Head Remarks:</strong> ' + (reqData.remarks || 'None') + '</p>'
                + '</div>'
                + '<p>Please log in to the AMS Portal to give your final decision, or use the quick action below:</p>'
                + '<div style="text-align:center;margin-top:20px;">'
                + '<a href="' + quickLink + '" style="background:#4caf7d;color:#fff;padding:12px 24px;text-decoration:none;border-radius:10px;font-weight:700;font-size:14px;display:inline-block;">Quick Approve →</a>'
                + '</div>';
              try {
                GmailApp.sendEmail(users[u][2], 'AMS Portal: Request Endorsed (' + reqData.reqId + ')', '',
                                   { htmlBody: _emailTemplate('Endorsed Request', deanContent) });
              } catch(me2) { _log('System', 'Email', 'Dean notification failed for ' + reqData.reqId); }
            }
          }
        }
        return 'Request ' + reqData.reqId + ' has been endorsed to the Dean.';
      }

      return 'Changes saved for request ' + reqData.reqId + '.';
    }
    return 'Error: Request ID not found.';
  } catch (err) {
    return 'Error processing request: ' + err.toString();
  }
}

function processBatchApprovals(batchData) {
  try {
    var count = 0;
    batchData.forEach(function(req) {
      updateRequestDetails({
        reqId: req.reqId,
        sheetName: req.sheetName,
        status: req.status,
        remarks: 'Batch processed by ' + req.logUser,
        logUser: req.logUser
      });
      count++;
    });
    return 'Successfully processed ' + count + ' request(s).';
  } catch (err) {
    return 'Error in batch processing: ' + err.toString();
  }
}


// ================================================================
//  PDF GENERATION
// ================================================================
function generateRequestPDF(reqId, sheetName) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return 'Error: Sheet not found.';
    var data    = sheet.getDataRange().getValues();
    var headers = data[0];
    var reqData = null;

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === reqId) {
        reqData = data[i].map(function(c) {
          if (c instanceof Date) return Utilities.formatDate(c, Session.getScriptTimeZone(), 'MM/dd/yyyy');
          return String(c);
        });
        break;
      }
    }
    if (!reqData) return 'Error: Request not found.';

    var formNames = {
      'Form1_Data': 'Daily Time Record Discrepancy',
      'Form2_Data': 'Shift to Online Modality',
      'Form3_Data': 'Late Submission of Grade',
      'Form4_Data': 'Academic Evaluation',
      'Form5_Data': 'Make-up Class Schedule'
    };
    var formTitle = formNames[sheetName] || sheetName;

    var statusColor = reqData[7] === 'Approved' ? '#4caf7d' : (reqData[7] === 'Rejected' ? '#e05b5b' : '#c9a96e');
    var skip = ['N_A','Status','Date_Submitted','Remarks','Attachment','Request_ID','Name','Department'];
    var fieldsHtml = '';
    for (var j = 3; j <= 6; j++) {
      var h = String(headers[j] || '');
      if (!h || skip.includes(h)) continue;
      var val = (reqData[j] || '—');
      fieldsHtml += '<tr><td style="padding:10px 12px;width:35%;color:#8d6e63;font-weight:600;border-bottom:1px solid #efebe9;">' 
        + h.replace(/_/g,' ') + '</td><td style="padding:10px 12px;border-bottom:1px solid #efebe9;">' 
        + val.replace(/\n/g,'<br>') + '</td></tr>';
    }

    var html = '<!DOCTYPE html>'
    + '<html><head><meta charset="UTF-8">'
    + '<style>'
    + 'body { font-family: Georgia, serif; padding: 40px 50px; background: #fff; color: #333; }'
    + '.header { text-align: center; border-bottom: 3px solid #c9a96e; padding-bottom: 20px; margin-bottom: 30px; }'
    + '.header h1 { color: #3e2723; font-size: 24px; margin: 0; letter-spacing: 2px; }'
    + '.header h2 { color: #c9a96e; font-size: 14px; margin: 8px 0 0; font-weight: 600; text-transform: uppercase; letter-spacing: 1px; }'
    + '.header p  { color: #8d6e63; margin: 5px 0 0; font-size: 13px; }'
    + 'table { width: 100%; border-collapse: collapse; font-size: 14px; }'
    + '.section { background: #faf7f5; padding: 18px; border-left: 4px solid #c9a96e; margin: 20px 0; border-radius: 4px; }'
    + '.section h3 { margin: 0 0 12px; color: #5d4037; font-size: 12px; text-transform: uppercase; letter-spacing: 1.5px; font-weight: 700; }'
    + '.status { color: ' + statusColor + '; font-weight: bold; font-size: 18px; text-transform: uppercase; }'
    + '.footer { text-align: center; margin-top: 50px; border-top: 2px solid #efebe9; padding-top: 15px; font-size: 11px; color: #a1887f; }'
    + '.remarks-box { border: 1px dashed #d7ccc8; padding: 15px; border-radius: 6px; min-height: 50px; font-style: italic; color: #6d4c41; margin-top: 10px; }'
    + '.meta-table td { padding: 8px 0; }'
    + '</style></head><body>'
    + '<div class="header">'
    + '<h1>ENDERUN COLLEGES</h1>'
    + '<h2>' + formTitle + '</h2>'
    + '<p>Official Academic Application Record</p>'
    + '</div>'
    + '<table class="meta-table">'
    + '<tr><td style="width:35%;color:#8d6e63;font-weight:bold;">Request ID:</td><td style="font-weight:bold;font-family:monospace;font-size:16px;">' + reqData[0] + '</td></tr>'
    + '<tr><td style="color:#8d6e63;font-weight:bold;">Date Submitted:</td><td>' + reqData[8] + '</td></tr>'
    + '<tr><td style="color:#8d6e63;font-weight:bold;">Final Status:</td><td class="status">' + reqData[7] + '</td></tr>'
    + '</table>'
    + '<div class="section">'
    + '<h3>Requestor Information</h3>'
    + '<table><tr><td style="width:35%;padding:6px 0;"><strong>Name:</strong></td><td>' + reqData[1] + '</td></tr>'
    + '<tr><td style="padding:6px 0;"><strong>Department:</strong></td><td>' + reqData[2] + '</td></tr></table>'
    + '</div>'
    + '<div class="section">'
    + '<h3>Application Details</h3>'
    + '<table>' + fieldsHtml + '</table>'
    + '</div>'
    + '<div class="section">'
    + '<h3>Approver Remarks</h3>'
    + '<div class="remarks-box">' + (reqData[9] || 'No remarks provided.') + '</div>'
    + '</div>'
    + '<div class="footer">'
    + '<p>This is a system-generated document from Enderun Colleges Academic Management System.</p>'
    + '<p>Generated on ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM dd, yyyy \'at\' hh:mm a') + '</p>'
    + '</div>'
    + '</body></html>';

    var blob = Utilities.newBlob(html, MimeType.HTML).getAs(MimeType.PDF);
    blob.setName(reqId + '_Application.pdf');
    var iter   = DriveApp.getFoldersByName('AMS_PDFs');
    var folder = iter.hasNext() ? iter.next() : DriveApp.createFolder('AMS_PDFs');
    var file   = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getDownloadUrl();
  } catch (err) {
    return 'Error generating PDF: ' + err.toString();
  }
}


// ================================================================
//  NOTIFICATIONS  (RBAC-filtered)
// ================================================================
function getNotifications(userRole, userForm) {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var allowed = _getAssignedSheets(userRole, userForm);
  var role    = String(userRole || '').trim();
  var notifs  = [];

  allowed.forEach(function(sName) {
    if (notifs.length >= 10) return;
    var s = ss.getSheetByName(sName);
    if (!s) return;
    var rows = s.getDataRange().getValues();
    for (var i = rows.length - 1; i >= 1; i--) {
      if (notifs.length >= 10) break;
      if (!rows[i][0] || String(rows[i][0]).toLowerCase().includes('request_id')) continue;
      
      // Row-level RBAC: Dept Head only sees their department
      if (role === 'Dept Head' && String(rows[i][2]).trim() !== String(userForm).trim()) continue;

      var status = String(rows[i][7]).trim();
      var relevant;
      if (role === 'Dept Head') {
        relevant = (status === 'Pending Program Head');
      } else {
        relevant = (status === 'Pending Program Head' || status === 'Pending Dean');
      }
      if (!relevant) continue;
      notifs.push({ id: String(rows[i][0]), name: String(rows[i][1]), date: _fmtDate(rows[i][8]) });
    }
  });
  return notifs;
}


// ================================================================
//  SYSTEM LOGS
// ================================================================
function addSystemLog(user, module, desc) {
  _log(user, module, desc);
}

function _log(user, module, desc) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s  = ss.getSheetByName('System_Logs');
  if (!s) {
    s = ss.insertSheet('System_Logs');
    s.appendRow(['Timestamp','User','Module','Description']);
    s.getRange(1,1,1,4).setFontWeight('bold').setBackground('#161410');
  }
  var ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm:ss');
  s.appendRow([ts, user || 'System', module || '—', desc || '—']);
}

function getSystemLogs() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('System_Logs');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var logs = [];
  for (var i = data.length - 1; i >= 1; i--) {
    if (!data[i][0] || String(data[i][0]).toLowerCase().includes('timestamp')) continue;
    logs.push({
      timestamp: _fmtDateFull(data[i][0]),
      user:   String(data[i][1]),
      module: String(data[i][2]),
      desc:   String(data[i][3])
    });
  }
  return logs;
}


// ================================================================
//  AUTOMATED WEEKLY REPORT  (run createWeeklyTrigger() once)
// ================================================================
function createWeeklyTrigger() {
  ScriptApp.newTrigger('sendAutomatedWeeklyReport')
    .timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(17).create();
}

function sendAutomatedWeeklyReport() {
  var stats      = getDashboardStats('Admin', null);
  var usersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (!usersSheet) return;

  var content = 
    '<h2 style="color:#c9a96e;text-align:center;">Weekly Summary Report</h2>'
    + '<div style="background:#161410;padding:20px;border-radius:10px;border:1px solid #2e2a22;margin:20px 0;">'
    + '<p style="margin:6px 0;"><strong style="color:#7a6540;">Pending / In Progress:</strong> <span style="color:#e0a85b;font-size:18px;font-weight:700;">' + stats.inProgress + '</span></p>'
    + '<p style="margin:6px 0;"><strong style="color:#7a6540;">Approved this period:</strong> <span style="color:#4caf7d;font-size:18px;font-weight:700;">' + stats.progressed + '</span></p>'
    + '<p style="margin:6px 0;"><strong style="color:#7a6540;">Rejected:</strong> <span style="color:#e05b5b;font-size:18px;font-weight:700;">' + stats.rejected + '</span></p>'
    + '<p style="margin:6px 0;"><strong style="color:#7a6540;">Total Historical Requests:</strong> ' + stats.totalRequests + '</p>'
    + '</div>'
    + '<p style="text-align:center;">Log in on Monday to review and clear the pending queue.</p>';

  var users = usersSheet.getDataRange().getValues();
  for (var i = 1; i < users.length; i++) {
    if (users[i][3] === 'Admin' && users[i][5] === 'Active') {
      try {
        GmailApp.sendEmail(users[i][2], 'AMS Portal: Weekly End-of-Week Report', '',
                           { htmlBody: _emailTemplate('Weekly Report', content) });
      } catch(me) {}
    }
  }
}


// ================================================================
//  HELPERS
// ================================================================
function _hashPassword(password) {
  if (!password) return '';
  var rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password, Utilities.Charset.UTF_8);
  var hexHash = '';
  for (var i = 0; i < rawHash.length; i++) {
    var hashVal = rawHash[i];
    if (hashVal < 0) hashVal += 256;
    if (hashVal.toString(16).length == 1) hexHash += '0';
    hexHash += hashVal.toString(16);
  }
  return hexHash;
}

function _fmtDate(raw) {
  if (raw instanceof Date) return Utilities.formatDate(raw, Session.getScriptTimeZone(), 'MM/dd/yyyy');
  return String(raw).substring(0, 10);
}

function _fmtDateFull(raw) {
  if (raw instanceof Date) return Utilities.formatDate(raw, Session.getScriptTimeZone(), 'MM/dd/yyyy HH:mm:ss');
  return String(raw);
}

function _ensureSheet(ss, name, headers) {
  var s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.appendRow(headers);
    s.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#161410');
  } else {
    var first = s.getRange('A1').getValue();
    if (!first || first === '') {
      s.getRange(1, 1, 1, headers.length).setValues([headers]);
      s.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#161410');
    }
  }
  return s;
}

function _ensureFormSheet(ss, sheetName) {
  var cols = ['Request_ID','Name','Department','Field1','Field2','Field3','Details','Status','Date_Submitted','Remarks','Attachment','Auth_Token'];
  return _ensureSheet(ss, sheetName, cols);
}

function _emailTemplate(title, content) {
  return '<!DOCTYPE html>'
  + '<html><head><meta charset="UTF-8">'
  + '<style>'
  + 'body { background:#0f0e0c; margin:0; padding:30px 15px; font-family:"Segoe UI",Tahoma,sans-serif; }'
  + '.wrap { max-width:580px; margin:0 auto; background:#1c1a16; border-radius:14px; overflow:hidden; border:1px solid #2e2a22; }'
  + '.hdr  { background:linear-gradient(135deg,#c9a96e,#e8c98a); padding:28px 30px; text-align:center; }'
  + '.hdr h1 { color:#0f0e0c; margin:0; font-size:20px; letter-spacing:2px; text-transform:uppercase; font-weight:800; }'
  + '.bdy  { padding:30px; color:#a89880; line-height:1.7; font-size:15px; }'
  + '.bdy strong { color:#f0ead8; }'
  + '.ftr  { background:#161410; padding:20px 30px; text-align:center; border-top:1px solid #2e2a22; }'
  + '.ftr p { margin:0; color:#5c5244; font-size:11px; text-transform:uppercase; letter-spacing:1px; }'
  + '</style></head>'
  + '<body><div class="wrap">'
  + '<div class="hdr"><h1>' + title + '</h1></div>'
  + '<div class="bdy">' + content + '</div>'
  + '<div class="ftr"><p>Enderun Colleges &middot; Academic Management System</p><p style="margin-top:6px;">Automated notification — please do not reply directly.</p></div>'
  + '</div></body></html>';
}


// ================================================================
//  ONE-TIME SETUP  (run once to initialize all sheets)
// ================================================================
function setupAllSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  _ensureSheet(ss, 'Users',       ['User_ID','Name','Email','Role','Assigned_Form','Status','Date_Added','Password','IsFirstLogin']);
  _ensureSheet(ss, 'System_Logs', ['Timestamp','User','Module','Description']);
  _ensureSheet(ss, 'Form1_Data',  ['Request_ID','Name','Department','Date_Of_Discrepancy','Time_In_Out','N_A','Reason','Status','Date_Submitted','Remarks','Attachment','Auth_Token']);
  _ensureSheet(ss, 'Form2_Data',  ['Request_ID','Name','Department','Course_Details','Class_Sched','Proposed_Date','Workplan','Status','Date_Submitted','Remarks','Attachment','Auth_Token']);
  _ensureSheet(ss, 'Form3_Data',  ['Request_ID','Name','Department','Semester','Term_Grade','N_A','Class_List','Status','Date_Submitted','Remarks','Attachment','Auth_Token']);
  _ensureSheet(ss, 'Form4_Data',  ['Request_ID','Name','Department','Prof_Name','Subject','Rating','Feedback','Status','Date_Submitted','Remarks','Attachment','Auth_Token']);
  _ensureSheet(ss, 'Form5_Data',  ['Request_ID','Name','Department','Subject','Missed_Class','Makeup_Class','Reason','Status','Date_Submitted','Remarks','Attachment','Auth_Token']);
  Logger.log('All sheets initialized successfully!');
  return 'All sheets initialized successfully!';
}