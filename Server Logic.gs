/**
 * @OnlyCurrentDoc
 *
 * The above comment directs App Script to limit the scope of file authorization now that script is container-bound.
 */

// --- CONFIGURATION ---
const REPORT_LIST_SHEET_NAME = 'Report List';
const USERS_SHEET_NAME = 'Users';
const DEFAULT_USERS_SHEET_NAME = 'Default Users';
// Schedule Sheet (Primary Data Source)
const COVER_TAB_SHEET_NAME = 'Cover Tab';       // Override Data Source
const ACTIVE_USERS_SHEET_NAME = 'Active Users';
// Sheet to track active viewers
const FORM_RESPONSES_SHEET_NAME = 'Form responses 1';
const NEW_REPORTS_SHEET_NAME = 'New Reports';
const USEFUL_LINKS_SHEET_NAME = 'Useful Links';
const LOG_SHEET_NAME = 'Log';
const WEB_APP_URL = "https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec";

// --- Column Indices (1-based) for "Report List" ---
const COL = {
  REPORT_NAME: 1,      // A
  ALLOCATED_TO: 2,     // B
  COMMENTS: 3,         // C  <-- NOTE: Your HTML maps this index to comments
  FREQUENCY: 4,        // D
  REPORT_TYPE: 5,      // E (Not explicitly used in UI logic, but present)
  NEXT_RUN_DATE: 6,    // F
  DUE_TIME: 7,         // G
  COMPLETED_DATE: 8,   // H
  COMPLETION_TIME: 9,  // I (Not explicitly used in UI logic, but present)
  PROCESS_NOTES: 10,   // J
  REPORT_LINK: 11,     // K
  CREATED_DATE: 13     // M <-- Index 13
};

const REPORT_LIST_HEADER_ROWS = 1;
const ALLOCATED_TO_COLUMN = 2; // Column index for 'Allocated To' in REPORT_LIST_SHEET_NAME (Column B)

/**
 * Parses a date string from the sheet, handling UK and ISO formats.
 * @param {string|Date} dateString The date string or Date object from the sheet.
 * @returns {Date|null} A valid Date object with time zeroed out (UTC midnight) or null.
 */
function parseDateString(dateString) {
  if (dateString instanceof Date && !isNaN(dateString)) {
    // If it's already a valid Date object, zero out time in UTC
    dateString.setUTCHours(0, 0, 0, 0);
    return dateString;
  }
  if (!dateString || typeof dateString !== 'string') {
    return null;
  }

  const s = dateString.trim();
  let date = null;

  try {
    // Try parsing ISO format (YYYY-MM-DD) or common variations first
    const isoMatch = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
    if (isoMatch) {
      const [, year, month, day] = isoMatch.map(Number);
      if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
        date = new Date(Date.UTC(year, month - 1, day));
        if (!isNaN(date)) return date; // Return early if successful
      }
    }

    // Try parsing UK format (DD-MM-YYYY)
    const ukMatch = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})/);
    if (ukMatch) {
      const [, day, month, year] = ukMatch.map(Number);
      if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
        date = new Date(Date.UTC(year, month - 1, day));
        if (!isNaN(date)) return date; // Return early if successful
      }
    }

     // Try parsing US format (MM-DD-YYYY) as a fallback
    const usMatch = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})/);
    if (usMatch) {
      const [, month, day, year] = usMatch.map(Number);
      if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
        date = new Date(Date.UTC(year, month - 1, day));
        if (!isNaN(date)) return date;
      }
    }

    // If specific formats fail, try the generic Date constructor as a last resort
    if (!date || isNaN(date.getTime())) {
      let genericDate = new Date(s);
      if (!isNaN(genericDate.getTime())) {
        genericDate.setUTCHours(0, 0, 0, 0);
        date = genericDate;
      }
    }

    // Final check
    if (date && !isNaN(date.getTime())) {
      return date;
    } else {
        throw new Error(`Date string "${s}" did not result in a valid Date object.`);
    }

  } catch (e) {
    Logger.log(`Could not parse date: "${dateString}". Cleaned: "${s}". Error: ${e.message}`);
    return null;
  }
}

/**
 * Appends a row to the 'Log' sheet.
 * @param {string} action - The name of the action being performed (e.g., "deleteReport").
 * @param {string} details - A string with specific details (e.g., "Report: Project X").
 * @param {string} userEmail - The email of the user performing the action.
 */
function logAction(action, details, userEmail) {
  try {
    const user = userEmail || Session.getActiveUser().getEmail();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
    if (sheet) {
      sheet.appendRow([new Date(), user, action, details]);
    }
  } catch (e) {
    Logger.log(`CRITICAL: Failed to write to Log sheet. Error: ${e}`);
  }
}

/**
 * Serves the HTML file for the web app after checking for authorization.
 * @returns {HtmlOutput} The HTML service output or an access denied message.
 */
function doGet() {
  Logger.log("--- START of doGet execution --- Timestamp: " + new Date().toISOString());

  try {
    if (!isCurrentUserAuthorized()) {
      Logger.log("User is NOT authorized. Serving 'unauthorized' page.");
      return HtmlService.createHtmlOutputFromFile('unauthorized')
        .setTitle('Access Denied');
    }
    Logger.log("User is authorized. Serving 'index' page.");
    return HtmlService.createHtmlOutputFromFile('index') // Make sure 'index.html' is your file name
      .setTitle('Company Data & MI Workflow Management')
      .setFaviconUrl('https://www.example.com/favicon.ico')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1') // Add viewport tag
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (e) {
    Logger.log("!!! CRITICAL ERROR in doGet: " + e.toString() + " Stack: " + e.stack);
    return HtmlService.createHtmlOutput("<h1>An unexpected error occurred.</h1><p>The application could not start. Please check the script logs for details.</p>");
  }
}

// -------------------------------------------------------------------------------------------------
// --- DAILY ALLOCATION LOGIC ---
// -------------------------------------------------------------------------------------------------

/**
 * Runs daily via a time-driven trigger.
 * 1. Checks 'Cover Tab' for any date/user overrides.
 * 2. If no override, uses 'Default Users' for today's tasks.
 * 3. Updates Column B (Allocated To) of the 'Report List' sheet.
 * 4. NEW: Sends a summary email to each user with their assigned tasks.
 */
function updateDailyAllocation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reportListSheet = ss.getSheetByName(REPORT_LIST_SHEET_NAME);
  const defaultUsersSheet = ss.getSheetByName(DEFAULT_USERS_SHEET_NAME);
  const coverTabSheet = ss.getSheetByName(COVER_TAB_SHEET_NAME);

  if (!reportListSheet || !defaultUsersSheet || !coverTabSheet) {
    Logger.log("Update failed: One or more required sheets not found.");
    return;
  }

  const today = new Date();
  const todayTime = Date.UTC(today.getUTCFullYear(), today.getUTCMonth(), today.getUTCDate());

  Logger.log(`--- Starting Daily Allocation for: ${today.toUTCString()} (UTC Check: ${new Date(todayTime).toISOString()}) ---`);

  // 1. Get Schedule Data (Report Name, Default User) from Default Users
  const scheduleDataRange = defaultUsersSheet.getRange('A2:B' + defaultUsersSheet.getLastRow());
  const scheduleData = scheduleDataRange.getValues();

  const scheduleMap = scheduleData.reduce((map, row) => {
    const reportName = (row[0] || '').toString().trim();
    if (reportName) {
      map[reportName] = {
        defaultUser: row[1] || '', // Default User in Col B
      };
    }
    return map;
  }, {});

  // 2. Get Override Data from Cover Tab
  const overrideDataRange = coverTabSheet.getRange('A2:J' + coverTabSheet.getLastRow());
  const overrideData = overrideDataRange.getValues();

  const overrideMap = overrideData.reduce((map, row) => {
    const reportName = (row[0] || '').toString().trim();
    if (reportName) {
      map[reportName] = {
        overrideUser: row[7] || '', // Col H
        overrideStart: row[8],     // Col I
        overrideEnd: row[9],       // Col J
      };
    }
    return map;
  }, {});

  // 3. Get main Report List names and current allocations
  const reportListDataRange = reportListSheet.getRange('A2:F' + reportListSheet.getLastRow());
  const reportListValues = reportListDataRange.getValues(); 
  const allocatedUserUpdates = [];
  const dailyTaskMap = {};

  // 4. Determine the assigned user for each report and build the update queue
  reportListValues.forEach((row, index) => {
    const reportName = (row[0] || '').toString().trim();
    if (!reportName) return;

    const currentAllocatedUser = (row[1] || '').toString().trim();

    Logger.log(`Processing report: "${reportName}"`);

    const scheduleEntry = scheduleMap[reportName];
    const overrideEntry = overrideMap[reportName];

    if (!scheduleEntry) {
      Logger.log(`-> SKIPPED: Not found in 'Default Users' sheet.`);
      return;
    }

    let newAssignedUser = '';
    let isOverridden = false;

    // STEP 1: Check for an active override.
    if (overrideEntry && overrideEntry.overrideStart && overrideEntry.overrideEnd) {
        Logger.log(`-> Found override entry. Start: ${overrideEntry.overrideStart}, End: ${overrideEntry.overrideEnd}`);

        const startDateObj = parseDateString(overrideEntry.overrideStart);
        const endDateObj = parseDateString(overrideEntry.overrideEnd);

        if (startDateObj && endDateObj) {
            Logger.log(`-> Dates parsed successfully. Start: ${startDateObj.toDateString()}, End: ${endDateObj.toDateString()}`);

            const startTime = startDateObj.getTime();
            const endTime = endDateObj.getTime();

            const isAfterStart = todayTime >= startTime;
            const isBeforeEnd = todayTime <= endTime;

            Logger.log(`-> Checking date range: IsToday (${todayTime}) >= Start (${startTime})? ${isAfterStart}. IsToday (${todayTime}) <= End (${endTime})? ${isBeforeEnd}.`);

            if (isAfterStart && isBeforeEnd) {
                isOverridden = true;
                Logger.log(`-> SUCCESS: Override is active.`);

                if (overrideEntry.overrideUser) {
                    newAssignedUser = overrideEntry.overrideUser;
                    Logger.log(`-> Set user to Override User: "${newAssignedUser}"`);
                } else {
                    newAssignedUser = scheduleEntry.defaultUser;
                    Logger.log(`-> Override User (Col H) is blank. Set to Default User: "${newAssignedUser}"`);
                }
            } else {
                Logger.log(`-> FAILED: Today is outside the override date range.`);
            }
        } else {
          Logger.log(`-> FAILED: Could not parse override dates: S='${overrideEntry.overrideStart}', E='${overrideEntry.overrideEnd}'`);
        }
    }

    // STEP 2: "otherwise" (No override is active)
    if (!isOverridden) {
        newAssignedUser = scheduleEntry.defaultUser;
        Logger.log(`-> No active override. Set to Default User: "${newAssignedUser}"`);
    }

    // Check if the report is due today by comparing its date to todayTime
    const nextRunDateValue = row[5];
    const runDateObj = parseDateString(nextRunDateValue);
    
    let isDueToday = false;

    if (runDateObj && runDateObj.getTime() === todayTime) {
      isDueToday = true;
    }

    // Only add to the email list if a user is assigned AND it's due today
    if (newAssignedUser && isDueToday) {
      const email = newAssignedUser.trim().toLowerCase();
      if (!dailyTaskMap[email]) {
        dailyTaskMap[email] = [];
      }
      dailyTaskMap[email].push(reportName);
    }

    // Check if the assignment needs to change
    if (currentAllocatedUser !== newAssignedUser.trim()) {
        Logger.log(`-> CHANGE DETECTED: Current ("${currentAllocatedUser}") != New ("${newAssignedUser.trim()}"). Adding to update queue.`);
        allocatedUserUpdates.push({
            row: index + 2, // Sheet row index (1-based + header offset)
            user: newAssignedUser.trim()
        });
    } else {
       Logger.log(`-> No change needed. Current user "${currentAllocatedUser}" is correct.`);
    }
  });


  // 5. Apply all updates to the Report List sheet
  if (allocatedUserUpdates.length > 0) {
    allocatedUserUpdates.forEach(update => {
        reportListSheet.getRange(update.row, ALLOCATED_TO_COLUMN).setValue(update.user);
    });
    Logger.log(`--- Successfully updated allocation for ${allocatedUserUpdates.length} reports. ---`);
  } else {
    Logger.log("--- No allocation updates required for today's date. ---");
  }

  // 6. Send summary emails to all users who were assigned tasks
  Logger.log(`Sending daily task emails...`);
  
  const appUrl = "https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec"; // Get deployed URL dynamically
  const todayString = today.toLocaleDateString('en-GB', { timeZone: 'Europe/London' });

  for (const userEmail in dailyTaskMap) {
    if (Object.prototype.hasOwnProperty.call(dailyTaskMap, userEmail)) {
      const tasks = dailyTaskMap[userEmail];
      if (tasks && Array.isArray(tasks) && tasks.length > 0) {
        sendDailyTaskEmail(userEmail, tasks, appUrl, todayString);
      } else if (tasks && !Array.isArray(tasks)) {
        Logger.log(`Skipping email for ${userEmail}: tasks variable was not an array.`);
      }
    }
  }
  Logger.log(`--- Finished sending daily task emails. ---`);
}

// -------------------------------------------------------------------------------------------------
// --- GENERAL APP FUNCTIONS ---
// -------------------------------------------------------------------------------------------------

/**
 * Checks if the current active user is listed in the 'Users' sheet.
 * @returns {boolean} True if the user is authorized, otherwise false.
 */
function isCurrentUserAuthorized() {
  try {
    const authorizedUserEmails = getUsers().map(email => email.toLowerCase());
    const currentUserEmail = Session.getActiveUser().getEmail().toLowerCase();

    if (!currentUserEmail) return false; // Not a valid user session

    return authorizedUserEmails.includes(currentUserEmail);
  } catch (e) {
    console.error('Authorization check failed:', e);
    return false;
  }
}

/**
 * Logs the current user's email and a timestamp to the 'Active Users' sheet.
 */
function logUserActivity() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ACTIVE_USERS_SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${ACTIVE_USERS_SHEET_NAME}" not found. Please create it.`);

    const data = sheet.getDataRange().getValues();
    let userFound = false;

    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === userEmail) {
        sheet.getRange(i + 1, 2).setValue(new Date());
        userFound = true;
        break;
      }
    }

    if (!userFound) {
      sheet.appendRow([userEmail, new Date()]);
    }
  } catch(e) {
    console.error('Error in logUserActivity:', e);
  }
}

/**
 * Removes the current user from the 'Active Users' sheet when they exit the app.
 */
function removeUserActivity() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) return;

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ACTIVE_USERS_SHEET_NAME);
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 0; i--) {
      if (data[i][0] === userEmail) {
        sheet.deleteRow(i + 1);
        break; 
      }
    }
  } catch(e) {
    console.error('Error in removeUserActivity:', e);
  }
}

/**
 * Cleans up inactive users and returns a list of active viewers.
 * An active viewer is someone who has sent a heartbeat in the last 150 seconds.
 * @returns {string[]} An array of email addresses for active users.
 */
function getActiveViewers() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ACTIVE_USERS_SHEET_NAME);
    if (!sheet) return [];

    const lastRow = sheet.getLastRow();
    if (lastRow < 1) return [];

    const data = sheet.getDataRange().getValues();
    const now = new Date().getTime();
    const activeThreshold = 150 * 1000; // 150 seconds (2.5 minutes)
    const activeEmails = [];
    const rowsToDelete = [];

    for (let i = data.length - 1; i >= 0; i--) {
      const row = data[i];
      const userEmail = row[0];
      const timestampValue = row[1]; // Get the timestamp value

      let timestamp = 0;
      if (timestampValue instanceof Date) {
        timestamp = timestampValue.getTime();
      } else if (timestampValue) {
        try {
          timestamp = new Date(timestampValue).getTime();
        } catch (parseError) {
          Logger.log(`Could not parse timestamp in getActiveViewers for row ${i + 1}: ${timestampValue}`);
          timestamp = 0; // Treat unparseable as inactive
        }
      }

      if (timestamp > 0 && (now - timestamp) < activeThreshold) {
        activeEmails.push(userEmail);
      } else {
        rowsToDelete.push(i + 1);
      }
    }

    rowsToDelete.sort((a, b) => b - a);
    for (const rowIndex of rowsToDelete) {
      sheet.deleteRow(rowIndex);
    }

    return activeEmails.reverse(); // Return in original order (oldest active first)
  } catch(e) {
    console.error('Error in getActiveViewers:', e);
    return []; // Return empty list on error
  }
}


/**
 * Gets a list of user emails from the 'Users' sheet.
 * @returns {string[]} An array of user email addresses.
 */
function getUsers() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USERS_SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${USERS_SHEET_NAME}" not found.`);

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    return sheet.getRange('A2:A' + lastRow).getValues().flat().filter(String);
  } catch (e) {
    console.error('Error in getUsers:', e);
    throw new Error('Could not retrieve user list. Please check sheet names and data.');
  }
}

/**
 * Gets a sorted list of user objects (email and role) from the "Users" sheet. Admin only.
 * Sorted primarily by Role (Admin first), then by Email.
 * @returns {object[]} An array of objects, e.g., [{email: 'a@b.com', role: 'Admin'}].
 */
function getUserData() {
  if (!getUserInfo().isAdmin) {
    throw new Error("Permission denied. Only admins can view user data.");
  }
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USERS_SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${USERS_SHEET_NAME}" not found.`);
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    
    const userData = sheet.getRange('A2:B' + lastRow).getValues().map(row => ({
      email: row[0],
      role: row[1] || '' // Ensure role is always a string, default to User ('')
    })).filter(user => user.email);

    // --- Sort the data ---
    userData.sort((a, b) => {
      const roleA = (a.role || 'User').toLowerCase();
      const roleB = (b.role || 'User').toLowerCase();
      const emailA = a.email.toLowerCase();
      const emailB = b.email.toLowerCase();

      if (roleA === 'admin' && roleB !== 'admin') return -1;
      if (roleA !== 'admin' && roleB === 'admin') return 1;

      if (emailA < emailB) return -1;
      if (emailA > emailB) return 1;

      return 0;
    });

    return userData;

  } catch (e) {
    console.error('Error in getUserData:', e);
    throw new Error('Could not retrieve user data.');
  }
}

/**
 * Updates a user's role in the "Users" sheet. Admin only.
 * Throws an error if the user already has the specified role.
 * @param {object} userData - An object with {email: '...', newRole: '...'}.
 */
function updateUser(userData) {
  if (!getUserInfo().isAdmin) {
    throw new Error("Permission denied. Only admins can update users.");
  }
  const newRole = userData.newRole || 'User';
  logAction("updateUserRole", `User: ${userData.email}, New Role: ${newRole}`, getUserInfo().email);
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(USERS_SHEET_NAME);
    if (!sheet) {
      throw new Error(`Sheet "${USERS_SHEET_NAME}" not found.`);
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
        throw new Error('No user data found in the sheet.');
    }

    const dataRange = sheet.getRange('A2:B' + lastRow);
    const data = dataRange.getValues();

    const rowIndex = data.findIndex(row => row[0] && row[0].toLowerCase() === userData.email.toLowerCase());
    
    if (rowIndex === -1) {
      throw new Error('User not found.');
    }

    const currentRole = data[rowIndex][1] || '';
    if (currentRole.toLowerCase() === newRole.toLowerCase()) {
      throw new Error(`User '${userData.email}' already has the role '${newRole || 'User'}'. No update needed.`);
    }

    sheet.getRange(rowIndex + 2, 2).setValue(newRole);
  } catch (e) {
    console.error('Error in updateUser:', e);
    throw new Error(e.message || 'Failed to update user.');
  }
}

/**
 * Deletes a user from the "Users" sheet. Admin only.
 * Clears their email from "Default Users" (Col B).
 * Clears their assigned tasks from "Report List" (Col B).
 * Updates the Google Form dropdown.
 * @param {string} emailToDelete - The email of the user to delete.
 */
function deleteUser(emailToDelete) {

  const userInfo = getUserInfo();
  if (!getUserInfo().isAdmin) {
      throw new Error("Permission denied. Only admins can delete users.");
  }
  logAction("deleteUser", `User: ${emailToDelete}`, userInfo.email);
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const emailToDeleteLower = emailToDelete.toLowerCase();
    
    // --- Part 1: Delete from "Users" sheet (Deletes entire row) ---
    const sheet = ss.getSheetByName(USERS_SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${USERS_SHEET_NAME}" not found.`);

    const emails = sheet.getRange('A2:A' + sheet.getLastRow()).getValues().flat();
    const rowIndex = emails.findIndex(email => email.toLowerCase() === emailToDeleteLower);

    if (rowIndex === -1) throw new Error('User not found in "Users" sheet.');
    sheet.deleteRow(rowIndex + 2); // Delete the user's row
    Logger.log(`Deleted user ${emailToDelete} from Users sheet.`);
    
    // --- Part 2: Clear from "Default Users" sheet (Clears Col B content only) ---
    const defaultUsersSheet = ss.getSheetByName(DEFAULT_USERS_SHEET_NAME);
    if (defaultUsersSheet) {
      const lastRow = defaultUsersSheet.getLastRow();
      if (lastRow > 1) { 
        const bColumnRange = defaultUsersSheet.getRange('B2:B' + lastRow);
        const bColumnValues = bColumnRange.getValues();
        let hasChanges = false;

        for (let i = 0; i < bColumnValues.length; i++) {
          const email = bColumnValues[i][0];
          if (email && email.toLowerCase() === emailToDeleteLower) {
            bColumnValues[i][0] = '';
            hasChanges = true;
          }
        }

        if (hasChanges) {
          bColumnRange.setValues(bColumnValues);
          Logger.log(`Cleared user ${emailToDelete} from Default Users sheet.`);
        }
      }
    } else {
      console.warn(`Sheet "${DEFAULT_USERS_SHEET_NAME}" not found. Could not clear email from default list.`);
    }

    // --- Part 3: Clear from "Report List" sheet (NEW) ---
    const reportListSheet = ss.getSheetByName(REPORT_LIST_SHEET_NAME);
    if (reportListSheet) {
      const lastRow = reportListSheet.getLastRow();
      if (lastRow > 1) { 
        const allocatedRange = reportListSheet.getRange('B2:B' + lastRow);
        const allocatedValues = allocatedRange.getValues();
        let hasChanges = false;

        for (let i = 0; i < allocatedValues.length; i++) {
          if (allocatedValues[i][0] && allocatedValues[i][0].toLowerCase() === emailToDeleteLower) {
            allocatedValues[i][0] = '';
            hasChanges = true;
          }
        }

        if (hasChanges) {
          allocatedRange.setValues(allocatedValues);
          Logger.log(`Cleared tasks from Report List for ${emailToDelete}.`);
        }
      }
    } else {
      console.warn('Report List sheet not found. Could not clear tasks.');
    }

    // --- Part 4: Update the Google Form Dropdown (NEW) ---
    updateDropdownFromSheet();
    Logger.log('Triggered Google Form dropdown update.');

  } catch (e) {
    console.error('Error in deleteUser:', e);
    throw new Error(e.message || 'Failed to delete user.');
  }
}

/**
 * Adds a new user to the "Users" sheet. Admin only.
 * Sends welcome email to new user and notification email to admins.
 * @param {object} newUser - An object with {email: '...', role: '...'}.
 */
function addUser(newUser) {
  const adminUserInfo = getUserInfo();
  
  if (!adminUserInfo.isAdmin) {
    throw new Error("Permission denied. Only admins can add users.");
  }
  
  logAction("addUser", `New User: ${newUser.email}, Role: ${newUser.role || 'User'}`, adminUserInfo.email);
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(USERS_SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${USERS_SHEET_NAME}" not found.`);

    const emails = sheet.getRange('A2:A' + sheet.getLastRow()).getValues().flat();
    const userExists = emails.some(email => email.toLowerCase() === newUser.email.toLowerCase());

    if (userExists) throw new Error('User with this email already exists.');
    if (!newUser.email || !newUser.email.includes('@')) throw new Error('Invalid email address provided.');
    
    sheet.appendRow([newUser.email, newUser.role || '']);
    
    // --- NEW: Send Emails After Adding User ---
    const appUrl = ScriptApp.getService().getUrl();
    
    // 1. Send Welcome Email to New User
    try {
      sendNewUserWelcomeEmail(newUser.email, appUrl, adminUserInfo.email);
      Logger.log(`Sent welcome email to new user: ${newUser.email}`);
    } catch (e) {
      Logger.log(`Failed to send welcome email to ${newUser.email}. Error: ${e}`);
    }

    // 2. Send Notification Email to All Admins
    try {
      const allUserData = sheet.getRange('A2:B' + sheet.getLastRow()).getValues();
      const adminEmails = allUserData
        .filter(row => row[1] && row[1].toLowerCase() === 'admin' && row[0])
        .map(row => row[0].toLowerCase());
        
      if (adminEmails.length > 0) {
        sendAdminNewUserNotification(adminEmails, newUser.email, newUser.role || 'User', appUrl, adminUserInfo.email);
        Logger.log(`Sent new user notification to admins: ${adminEmails.join(', ')}`);
      } else {
         Logger.log(`No admins found to notify about new user ${newUser.email}.`);
      }
    } catch (e) {
      Logger.log(`Failed to send admin notification email about ${newUser.email}. Error: ${e}`);
    }

    updateDropdownFromSheet();

  } catch (e) {
    console.error('Error in addUser:', e);
    throw new Error(e.message || 'Failed to add new user.');
  }
}

/**
 * Gets the current user's email and checks if they are an admin.
 * @returns {object} An object containing the user's email and a boolean isAdmin status.
 */
function getUserInfo() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USERS_SHEET_NAME);
    
    if (!sheet) throw new Error(`Sheet "${USERS_SHEET_NAME}" not found.`);

    const data = sheet.getRange('A2:B' + sheet.getLastRow()).getValues();
    const adminEmails = data
      .filter(row => row[1] && row[1].toLowerCase() === 'admin')
      .map(row => row[0].toLowerCase());
      
    const isAdmin = adminEmails.includes(userEmail.toLowerCase());
    return { email: userEmail, isAdmin: isAdmin };
  } catch (e) {
    console.error('Error in getUserInfo:', e);
    throw new Error('Could not verify user information.');
  }
}

/**
 * ====================================================================
 * CRITICAL MODIFICATION: Updated getReportData
 * Fetches required columns (up to M) and formats dates/times correctly.
 * ====================================================================
 */
function getReportData() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REPORT_LIST_SHEET_NAME);
    if (!sheet) {
      Logger.log(`Sheet "${REPORT_LIST_SHEET_NAME}" not found.`);
      throw new Error(`Sheet "${REPORT_LIST_SHEET_NAME}" not found.`);
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < REPORT_LIST_HEADER_ROWS + 1) return [];

    const dataRange = sheet.getRange(REPORT_LIST_HEADER_ROWS + 1, 1, lastRow - REPORT_LIST_HEADER_ROWS, 13);
    const data = dataRange.getValues();
    Logger.log(`Retrieved ${data.length} rows of raw data up to column M.`);

    const timeZone = Session.getScriptTimeZone();
    
    const formatDate = (dateInput) => {
      if (dateInput instanceof Date && !isNaN(dateInput)) {
          return Utilities.formatDate(dateInput, timeZone, "yyyy-MM-dd");
      }
      return '';
    };

    const formatDateTime = (datetimeInput) => {
      if (datetimeInput instanceof Date && !isNaN(datetimeInput)) {
        try {
          return Utilities.formatDate(datetimeInput, timeZone, "yyyy-MM-dd HH:mm:ss");
        } catch (e) {
          Logger.log(`Error formatting datetime: ${datetimeInput} - ${e}`);
          return '';
        }
      }
      return '';
    };
    
    const formatTime = (timeInput) => {
       if (timeInput instanceof Date && !isNaN(timeInput)) {
        try {
          return Utilities.formatDate(timeInput, timeZone, "HH:mm");
        } catch(e) {
          Logger.log(`Error formatting time: ${timeInput} - ${e}`);
          return '';
        }
      }
      if (typeof timeInput === 'string') {
        const trimmedTime = timeInput.trim();
        if (/^\d{1,2}:\d{2}$/.test(trimmedTime)) {
          const parts = trimmedTime.split(':');
          const hour = parts[0].padStart(2, '0');
          const minute = parts[1].padStart(2, '0');
          return `${hour}:${minute}`;
        }
      }
      return '';
    };
    
    const reportObjects = data.map(row => {
      if (!row[COL.REPORT_NAME - 1]) return null;

      return {
        reportName:       row[COL.REPORT_NAME - 1],
        allocatedTo:      row[COL.ALLOCATED_TO - 1] || '',
        comments:         row[COL.COMMENTS - 1] || '',
        frequency:        row[COL.FREQUENCY - 1] || '',
        nextRunDate:      formatDate(row[COL.NEXT_RUN_DATE - 1]),
        dueTime:          formatTime(row[COL.DUE_TIME - 1]),
        completedDate:    formatDateTime(row[COL.COMPLETED_DATE - 1]),
        processNotesLink: row[COL.PROCESS_NOTES - 1] || '',
        reportLink:       row[COL.REPORT_LINK - 1] || '',
        createdDate:      formatDateTime(row[COL.CREATED_DATE - 1]) 
      };
    }).filter(report => report !== null);

    reportObjects.sort((a, b) => {
      const compare = (valA, valB) => {
        const strA = (valA || '').toString().toLowerCase();
        const strB = (valB || '').toString().toLowerCase();
        const isEmptyA = !strA;
        const isEmptyB = !strB;

        if (isEmptyA && isEmptyB) return 0;
        if (isEmptyA) return 1;
        if (isEmptyB) return -1;

        if (strA < strB) return -1;
        if (strA > strB) return 1;
        return 0;
      };

      const dateCompare = compare(a.nextRunDate, b.nextRunDate);
      if (dateCompare !== 0) return dateCompare;

      const timeCompare = compare(a.dueTime, b.dueTime);
      if (timeCompare !== 0) return timeCompare;

      const nameCompare = compare(a.reportName, b.reportName);
      return nameCompare;
    });

    Logger.log(`Successfully mapped and sorted ${reportObjects.length} report objects.`);
    return reportObjects;
  } catch (e) {
    console.error('Error in getReportData:', e);
    Logger.log(`Error caught in getReportData: ${e.toString()}\nStack: ${e.stack}`);
    throw new Error('Could not retrieve report data.');
  }
}

/**
 * Gets report names, default users, frequency, run day, and next run date.
 * Merges with current override data from 'Cover Tab'.
 * Admin only.
 */
function getScheduleData() {
  if (!getUserInfo().isAdmin) {
    throw new Error("Permission denied. Only admins can view schedule data.");
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const defaultUsersSheet = ss.getSheetByName(DEFAULT_USERS_SHEET_NAME);
    const coverTabSheet = ss.getSheetByName(COVER_TAB_SHEET_NAME);
    const reportListSheet = ss.getSheetByName(REPORT_LIST_SHEET_NAME); 

    if (!defaultUsersSheet || !coverTabSheet || !reportListSheet) {
      throw new Error("One or more required sheets not found.");
    }

    // 1. Get Report Details from Report List
    const lastRowReport = reportListSheet.getLastRow();
    const reportDetailsMap = {};
    
    if (lastRowReport >= 2) {
      const reportData = reportListSheet.getRange(2, 1, lastRowReport - 1, 6).getValues();
      const timeZone = Session.getScriptTimeZone();

      reportData.forEach(row => {
        const name = (row[0] || '').toString().trim();
        if (name) {
          let nextRun = '';
          if (row[5] instanceof Date) {
             nextRun = Utilities.formatDate(row[5], timeZone, "yyyy-MM-dd");
          }
          
          reportDetailsMap[name] = {
            frequency: row[3] || '-',        
            runDay: row[4] || '-',           
            nextRunDate: nextRun || '-'      
          };
        }
      });
    }

    // 2. Get Default Users
    const defaultUsersLastRow = defaultUsersSheet.getLastRow();
    const defaultUserMap = {};
    
    if (defaultUsersLastRow >= 2) {
      const defaultUsersData = defaultUsersSheet.getRange('A2:B' + defaultUsersLastRow).getValues();
      defaultUsersData.forEach(row => {
          const name = (row[0] || '').toString().trim();
          if (name) defaultUserMap[name] = row[1] || '';
      });
    }

    // 3. Get Overrides
    const coverTabLastRow = coverTabSheet.getLastRow();
    const coverTabMap = {};
    
    if (coverTabLastRow >= 2) {
      const coverTabData = coverTabSheet.getRange('A2:J' + coverTabLastRow).getValues();
      coverTabData.forEach((row) => {
        const name = (row[0] || '').toString().trim();
        if (name) {
           let start = '', end = '';
           if (row[8] instanceof Date) start = Utilities.formatDate(row[8], Session.getScriptTimeZone(), "yyyy-MM-dd");
           if (row[9] instanceof Date) end = Utilities.formatDate(row[9], Session.getScriptTimeZone(), "yyyy-MM-dd");
           
           coverTabMap[name] = {
             overrideUser: row[7] || '',
             overrideStart: start,
             overrideEnd: end
           };
        }
      });
    }

    // 4. Merge Data
    const finalScheduleData = Object.keys(reportDetailsMap).map(reportName => {
      const details = reportDetailsMap[reportName];
      const defaultUser = defaultUserMap[reportName] || '';
      const override = coverTabMap[reportName] || {};

      return {
        reportName: reportName,
        frequency: details.frequency,
        runDay: details.runDay,             
        nextRunDate: details.nextRunDate,
        defaultUser: defaultUser,
        overrideUser: override.overrideUser || '',
        overrideStart: override.overrideStart || '',
        overrideEnd: override.overrideEnd || ''
      };
    });
    
    finalScheduleData.sort((a, b) => {
       const dateA = a.nextRunDate === '-' ? '9999-99-99' : a.nextRunDate;
       const dateB = b.nextRunDate === '-' ? '9999-99-99' : b.nextRunDate;
       if (dateA !== dateB) return dateA.localeCompare(dateB);
       return a.reportName.localeCompare(b.reportName);
    });
    
    return finalScheduleData;

  } catch (e) {
    console.error('Error in getScheduleData:', e);
    throw new Error('Could not retrieve schedule data.');
  }
}

// --- AMENDED FUNCTION: saveReportData ---
/**
 * Saves allocation, comments, and due time, with locking and robust row finding.
 * @param {object} reportData - { reportName, allocatedTo, comments, dueTime, setAsDefault, originalAllocatedTo }
 */
function saveReportData(reportData) {
  if (!reportData || !reportData.reportName) {
     Logger.log("saveReportData: Called with empty report name.");
     throw new Error("Report name cannot be empty.");
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    throw new Error("Another user is currently saving changes. Please wait a moment and try again.");
  }
  
  const userInfo = getUserInfo(); 

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const reportListSheet = ss.getSheetByName(REPORT_LIST_SHEET_NAME);
    if (!reportListSheet) throw new Error(`Sheet "${REPORT_LIST_SHEET_NAME}" not found.`);
    
    const lastReportRow = reportListSheet.getLastRow();
    if (lastReportRow < REPORT_LIST_HEADER_ROWS + 1) throw new Error("Report List sheet is empty.");
    
    const reportNamesRange = reportListSheet.getRange(REPORT_LIST_HEADER_ROWS + 1, COL.REPORT_NAME, lastReportRow - REPORT_LIST_HEADER_ROWS, 1);
    const reportNames = reportNamesRange.getDisplayValues().flat();
    const searchReportNameLowerTrimmed = reportData.reportName.trim().toLowerCase();
    
    Logger.log(`saveReportData: Searching for report name from HTML: "${searchReportNameLowerTrimmed}"`);
    
    const rowIndex = reportNames.findIndex((name, index) => {
        const sheetNameCleaned = (name || '').toString().trim().toLowerCase();
        return sheetNameCleaned === searchReportNameLowerTrimmed;
    });

    Logger.log(`saveReportData: findIndex result (0-based): ${rowIndex}`);
    
    if (rowIndex === -1) {
        Logger.log(`saveReportData: Report "${reportData.reportName}" (trimmed: "${searchReportNameLowerTrimmed}") not found.`);
        throw new Error(`Report "${reportData.reportName}" not found in "Report List". Check for exact name match.`);
    }
    const reportRow = rowIndex + REPORT_LIST_HEADER_ROWS + 1;

    // --- OPTIMISTIC LOCKING CHECK ---
    const actualCurrentUserInSheet = reportListSheet.getRange(reportRow, COL.ALLOCATED_TO).getDisplayValue();
    const actualCurrentUserNormalized = (actualCurrentUserInSheet || '').trim().toLowerCase();
    const originalUserFromClient = reportData.originalAllocatedTo;
    const originalUserNormalized = (originalUserFromClient || '').trim().toLowerCase();
    const newUserFromClient = reportData.allocatedTo;
    const newUserNormalized = (newUserFromClient || '').trim().toLowerCase();
    
    const rawComment = reportData.comments || '';
    let processedComment = String(rawComment);
    if (processedComment.startsWith('=')) {
      processedComment = "'" + processedComment;
    }

    if (actualCurrentUserNormalized !== originalUserNormalized) {
      throw new Error("This report was modified by another user. Please refresh the app and try again.");
    }

    // --- NEW: LOG THE ACTION ---
    const logDetails = [`Report: ${reportData.reportName}`];
    
    if (originalUserNormalized !== newUserNormalized) {
        logDetails.push(`Allocated: ${reportData.allocatedTo || 'Unassigned'}`);
    }
    if (reportData.comments) {
        logDetails.push(`Comment: "${reportData.comments}"`);
    }
    if (reportData.dueTime) {
        logDetails.push(`Due Time: ${reportData.dueTime}`);
    }
    if (reportData.setAsDefault) {
        logDetails.push("Flag: Set as Default User");
    }

    if (logDetails.length > 1) {
        logAction("saveReportData", logDetails.join(', '), userInfo.email);
    }

    // --- Helper function for saving default user ---
    const saveDefaultUser = () => { 
      const defaultUsersSheet = ss.getSheetByName(DEFAULT_USERS_SHEET_NAME);
      if (!defaultUsersSheet) {
        console.warn(`Sheet "${DEFAULT_USERS_SHEET_NAME}" not found. Could not set default user.`);
        return;
      }
      const lastRowDef = defaultUsersSheet.getLastRow();
      if (lastRowDef >= 2) {
        const defaultReportData = defaultUsersSheet.getRange('A2:B' + lastRowDef).getValues();
        const rowIndexDef = defaultReportData.findIndex(row =>
          row[0] && row[0].toString().toLowerCase() === reportData.reportName.toLowerCase()
        );
        if (rowIndexDef !== -1) {
          defaultUsersSheet.getRange(rowIndexDef + 2, 2).setValue(reportData.allocatedTo);
          Logger.log(`Updated default user for "${reportData.reportName}" in ${DEFAULT_USERS_SHEET_NAME}`);
        } else {
           console.warn(`Report "${reportData.reportName}" not found in "Default Users" sheet.`);
        }
      }
    };

    // Update Sheet
    reportListSheet.getRange(reportRow, COL.ALLOCATED_TO).setValue(reportData.allocatedTo || '');
    reportListSheet.getRange(reportRow, COL.COMMENTS).setValue(processedComment);
    reportListSheet.getRange(reportRow, COL.DUE_TIME).setValue(reportData.dueTime || null);

    if (reportData.setAsDefault) {
      saveDefaultUser();
    }

    if (actualCurrentUserNormalized !== newUserNormalized && reportData.allocatedTo) {
      try {
        sendAssignmentEmail(reportData.allocatedTo, reportData.reportName, userInfo.email);
        Logger.log(`Sent assignment email to ${reportData.allocatedTo} for ${reportData.reportName}`);
      } catch (emailError) {
         Logger.log(`Failed to send assignment email, but changes were saved. Error: ${emailError}`);
      }
    }
    Logger.log(`Saved data successfully for report "${reportData.reportName}" (Sheet Row ${reportRow})`);
    return "Changes saved.";
  } catch (e) {
    Logger.log(`Error in saveReportData: ${e.message}\n${e.stack}`);
    throw new Error(e.message || 'Failed to save changes.');
  } finally {
    lock.releaseLock();
  }
}

// --- AMENDED FUNCTION: saveCommentData ---
/**
 * Saves only the comments for a specific report, with permission checks and robust row finding.
 * @param {object} reportData - { reportName, comments }.
 */
function saveCommentData(reportData) {
    const userInfo = getUserInfo();
    let allocatedUser = '';
    let reportRow = -1; 

    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(REPORT_LIST_SHEET_NAME);
        if (!sheet) throw new Error(`Sheet "${REPORT_LIST_SHEET_NAME}" not found.`);
        
        const lastReportRow = sheet.getLastRow();
        if (lastReportRow < REPORT_LIST_HEADER_ROWS + 1) throw new Error("Report List sheet is empty.");
        
        const reportDataRange = sheet.getRange(REPORT_LIST_HEADER_ROWS + 1, 1, lastReportRow - REPORT_LIST_HEADER_ROWS, COL.ALLOCATED_TO);
        const reportSheetData = reportDataRange.getDisplayValues();
        const searchReportNameLowerTrimmed = reportData.reportName.trim().toLowerCase();
        
        const rowIndex = reportSheetData.findIndex(row =>
            (row[COL.REPORT_NAME - 1] || '').toString().trim().toLowerCase() === searchReportNameLowerTrimmed
        );
        
        if (rowIndex === -1) {
            Logger.log(`saveCommentData: Report "${reportData.reportName}" not found.`);
            throw new Error('Report not found.');
        }
        reportRow = rowIndex + REPORT_LIST_HEADER_ROWS + 1;
        allocatedUser = (reportSheetData[rowIndex][COL.ALLOCATED_TO - 1] || '').trim();

        if (!userInfo.isAdmin && (!allocatedUser || allocatedUser.toLowerCase() !== userInfo.email.toLowerCase())) {
            throw new Error("Permission denied. You can only save comments for reports allocated to you.");
        }

        const rawComment = reportData.comments || '';
        let processedComment = String(rawComment);
        if (processedComment.startsWith('=')) {
            processedComment = "'" + processedComment;
        }

        sheet.getRange(reportRow, COL.COMMENTS).setValue(processedComment);
        Logger.log(`Saved comment for "${reportData.reportName}" (Sheet Row ${reportRow})`);
        return "Comment saved.";
    } catch (e) {
        console.error('Error in saveCommentData:', e);
        Logger.log(`Error in saveCommentData for ${reportData.reportName}: ${e.message}\n${e.stack}`);
        throw new Error(e.message || 'Failed to save comment.');
    }
}


// --- AMENDED FUNCTION: markReportComplete ---
/**
 * Marks a report as complete, with permission checks and robust row finding.
 * @param {string} reportName - The name of the report to mark as complete.
 */
function markReportComplete(reportName) {
    const userInfo = getUserInfo();
    logAction("markReportComplete", `Report: ${reportName}`, userInfo.email);
    let allocatedUser = '';
    let reportRow = -1;

    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(REPORT_LIST_SHEET_NAME);
        if (!sheet) throw new Error(`Sheet "${REPORT_LIST_SHEET_NAME}" not found.`);
        
        const lastReportRow = sheet.getLastRow();
        if (lastReportRow < REPORT_LIST_HEADER_ROWS + 1) throw new Error("Report List sheet is empty.");
        
        const reportDataRange = sheet.getRange(REPORT_LIST_HEADER_ROWS + 1, 1, lastReportRow - REPORT_LIST_HEADER_ROWS, COL.ALLOCATED_TO);
        const reportSheetData = reportDataRange.getDisplayValues();
        const searchReportNameLowerTrimmed = reportName.trim().toLowerCase();
        
        const rowIndex = reportSheetData.findIndex(row =>
            (row[COL.REPORT_NAME - 1] || '').toString().trim().toLowerCase() === searchReportNameLowerTrimmed
        );
        
        if (rowIndex === -1) {
            Logger.log(`markReportComplete: Report "${reportName}" not found.`);
            throw new Error('Report not found.');
        }
        reportRow = rowIndex + REPORT_LIST_HEADER_ROWS + 1;
        allocatedUser = (reportSheetData[rowIndex][COL.ALLOCATED_TO - 1] || '').trim();

        if (!userInfo.isAdmin && (!allocatedUser || allocatedUser.toLowerCase() !== userInfo.email.toLowerCase())) {
            throw new Error("Permission denied. You can only complete reports allocated to you.");
        }

        sheet.getRange(reportRow, COL.COMPLETED_DATE).setValue(new Date());
        Logger.log(`Marked "${reportName}" as complete (Sheet Row ${reportRow}).`);
        return "Report marked complete.";
    } catch (e) {
        console.error('Error in markReportComplete:', e);
        Logger.log(`Error in markReportComplete for ${reportName}: ${e.message}\n${e.stack}`);
        throw new Error(e.message || 'Failed to mark report as complete.');
    }
}

/**
 * Clears the completion date (undoes completion) for a specific report.
 * This action is only allowed if the report is allocated to the current user.
 * @param {string} reportName - The name of the report to undo.
 */
function undoReportComplete(reportName) {
  try {
    const currentUserEmail = Session.getActiveUser().getEmail().toLowerCase();
    if (!currentUserEmail) {
      throw new Error("Could not identify current user.");
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(REPORT_LIST_SHEET_NAME);
    if (!sheet) {
      throw new Error(`Sheet "${REPORT_LIST_SHEET_NAME}" not found.`);
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < REPORT_LIST_HEADER_ROWS + 1) {
      throw new Error('No report data found.');
    }

    const dataRange = sheet.getRange(REPORT_LIST_HEADER_ROWS + 1, 1, lastRow - REPORT_LIST_HEADER_ROWS, COL.ALLOCATED_TO);
    const data = dataRange.getDisplayValues(); 

    const searchReportNameLowerTrimmed = reportName.trim().toLowerCase();
    const rowIndex = data.findIndex(row =>
        (row[COL.REPORT_NAME - 1] || '').toString().trim().toLowerCase() === searchReportNameLowerTrimmed
    );
    
    if (rowIndex === -1) {
      throw new Error('Report not found.');
    }

    const allocatedEmail = (data[rowIndex][COL.ALLOCATED_TO - 1] || '').trim().toLowerCase();
    
    if (allocatedEmail !== currentUserEmail) {
      throw new Error('Permission denied. You can only undo completion for reports allocated to you.');
    }

    const reportRow = rowIndex + REPORT_LIST_HEADER_ROWS + 1;
    sheet.getRange(reportRow, COL.COMPLETED_DATE).clearContent(); 
    
    Logger.log(`Undo complete successful for "${reportName}"`);
    return "Completion undone.";

  } catch (e) {
    console.error('Error in undoReportComplete:', e);
    throw new Error(e.message || 'Failed to undo report completion.');
  }
}

/**
 * Updates Column K (Report URL) and logs the change (including OLD vs NEW) to the Log sheet.
 * UNLOCKED: Open to all users.
 */
function saveReportUrl(reportName, newUrl) {
  const userInfo = getUserInfo();
  const currentUserEmail = userInfo.email || Session.getActiveUser().getEmail();

  if (!reportName) {
     throw new Error("Report name cannot be empty.");
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    throw new Error("Could not acquire lock. Please try again.");
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(REPORT_LIST_SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${REPORT_LIST_SHEET_NAME}" not found.`);

    const lastReportRow = sheet.getLastRow();
    if (lastReportRow < REPORT_LIST_HEADER_ROWS + 1) throw new Error("Report List sheet is empty.");
    
    const reportNamesRange = sheet.getRange(REPORT_LIST_HEADER_ROWS + 1, COL.REPORT_NAME, lastReportRow - REPORT_LIST_HEADER_ROWS, 1);
    const reportNames = reportNamesRange.getDisplayValues().flat();
    const searchReportNameLowerTrimmed = reportName.trim().toLowerCase();
    
    const rowIndex = reportNames.findIndex(name => 
        (name || '').toString().trim().toLowerCase() === searchReportNameLowerTrimmed
    );
    
    if (rowIndex === -1) throw new Error(`Report "${reportName}" not found.`);

    const rowToUpdate = rowIndex + REPORT_LIST_HEADER_ROWS + 1;
    
    const urlCell = sheet.getRange(rowToUpdate, COL.REPORT_LINK);
    const oldUrl = urlCell.getValue() || "(Empty)";

    urlCell.setValue(newUrl || '');
    
    const logDetails = `Report: ${reportName} | Old URL: ${oldUrl} | New URL: ${newUrl || '(Empty)'}`;
    logAction("Update Report URL", logDetails, currentUserEmail);

    return "Report URL saved successfully.";
  } catch (e) {
    console.error(`Error in saveReportUrl:`, e);
    throw new Error(e.message || 'Failed to save report URL.');
  } finally {
    lock.releaseLock(); 
  }
}

/**
 * Saves a single row of override data back to the "Cover Tab" sheet.
 * Admin only.
 * @param {object} scheduleRow - The override data object { reportName, overrideUser, overrideStart, overrideEnd }.
 */
function saveScheduleData(scheduleRow) {
  if (!getUserInfo().isAdmin) {
    throw new Error("Permission denied. Only admins can save schedule data.");
  }

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(COVER_TAB_SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${COVER_TAB_SHEET_NAME}" not found.`);

    const reportNames = sheet.getRange('A2:A' + sheet.getLastRow()).getDisplayValues().flat();
    const searchReportName = (scheduleRow.reportName || '').toString().trim();
    
    const rowIndex = reportNames.findIndex(name => (name || '').toString().trim() === searchReportName);
    
    if (rowIndex === -1) {
        throw new Error(`Report "${searchReportName}" not found in "${COVER_TAB_SHEET_NAME}" sheet.`);
    }

    const rowToUpdate = rowIndex + 2;

   const valuesToSet = [
      scheduleRow.overrideUser,
      scheduleRow.overrideStart || null, 
      scheduleRow.overrideEnd || null   
    ];
    
    sheet.getRange(rowToUpdate, 8, 1, 3).setValues([valuesToSet]);
  } catch (e) {
    console.error('Error in saveScheduleData:', e);
    throw new Error(e.message || 'Failed to save schedule override data.');
  }
}

/**
 * Sends a reminder email to the assigned user using a rich HTML template.
 * @param {string} recipientEmail - The email address of the user to remind.
 * @param {string} reportName - The name of the report.
 * @param {string} senderEmail - The email of the admin sending the reminder.
 */
function sendReminderEmail(recipientEmail, reportName, senderEmail) {
  if (!recipientEmail || !reportName || !senderEmail) {
    throw new Error("Recipient, report name, and sender email are required.");
  }

  const appUrl = ScriptApp.getService().getUrl();
  const subject = `Reminder: Task Completion Required for "${reportName}"`;

  const logoFileId = 'YOUR_LOGO_DRIVE_FILE_ID';
  const logoBlob = DriveApp.getFileById(logoFileId).getBlob().setName('logo.png');

  const htmlBody = `
    <div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
      <div style="max-width: 600px; margin: 20px auto; border: 1px solid #ddd; border-radius: 8px; overflow: hidden;">
        
        <div style="background-color: #2B5797; color: #fff; padding: 20px; text-align: center;">
          <img src="cid:logo" alt="Company Logo" style="max-height: 45px; margin-bottom: 10px;">
          <h2 style="margin: 0; font-size: 22px;">Task Reminder</h2>
        </div>
        
        <div style="padding: 25px;">
          <p>Hello,</p>
          <p>This is a friendly reminder that the report "<b>${reportName}</b>" is assigned to you and has not yet been marked as completed.</p>
          <p>Please review it in the workflow management tool as soon as possible.</p>
        
          <div style="text-align: center; margin: 30px 0;">
            <a href="${appUrl}" style="background-color: #6366f1; color: white; text-decoration: none; padding: 12px 24px; border-radius: 6px; font-weight: bold; display: inline-block;">
              Access Workflow Tool
            </a>
          </div>

          <p style="font-size:0.9em; color:#888; border-top:1px solid #eee; padding-top:15px;">
            This reminder was sent by: ${senderEmail}
          </p>
        </div>
        
        <div style="background-color: #f9f9f9; padding: 15px 25px; font-size: 0.9em; color: #555; border-top: 1px solid #eee;">
          <p>Thank you,<br>Company MI Workflow Management</p>
        </div>
      </div>
    </div>
  `;

  try {
    MailApp.sendEmail({
      to: recipientEmail,
      subject: subject,
      htmlBody: htmlBody,
      inlineImages: { logo: logoBlob }, 
      noReply: true,
      name: 'Company Workflow (No Reply)',
      replyTo: senderEmail 
    });
    Logger.log(`Reminder email successfully sent to ${recipientEmail}`);
  } catch (e) {
    Logger.log(`Failed to send reminder email to ${recipientEmail}: ${e}`);
    throw new Error('Failed to send reminder email. Please ensure script has Gmail permissions.');
  }
}

/**
 * Deletes a report row from "Cover Tab" first, then from
 * "Report List", "Default Users", and "Form responses 1".
 * Admin only.
 * @param {string} reportName The name of the report to delete.
 */
function deleteReport(reportName) {
  const userInfo = getUserInfo();
  
  if (!userInfo.isAdmin) {
    Logger.log(`deleteReport: Permission denied for user ${userInfo.email} trying to delete ${reportName}`);
    throw new Error("Permission denied. Only admins can delete reports.");
  }

  logAction("deleteReport", `Report: ${reportName}`, userInfo.email);
  
  if (!reportName) {
     Logger.log("deleteReport: Called with empty report name.");
     throw new Error("Report name cannot be empty.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reportNameToDelete = reportName.trim();

  // 2. Delete from "Cover Tab" first
  try {
    const coverSheet = ss.getSheetByName(COVER_TAB_SHEET_NAME);
    if (coverSheet) {
      const data = coverSheet.getDataRange().getDisplayValues();
      for (let i = data.length - 1; i >= 0; i--) {
        const currentReportName = (data[i][0] || '').toString().trim();
        if (currentReportName === reportNameToDelete) {
          coverSheet.deleteRow(i + 1);
          Logger.log(`deleteReport: Deleted row ${i + 1} for "${reportNameToDelete}" from ${COVER_TAB_SHEET_NAME}.`);
          break;
        }
      }
    } else {
        Logger.log(`deleteReport: Sheet "${COVER_TAB_SHEET_NAME}" not found. Skipping deletion for this sheet.`);
    }
  } catch(e) {
      Logger.log(`deleteReport: Error deleting from ${COVER_TAB_SHEET_NAME}: ${e}`);
      console.error(`Error deleting from ${COVER_TAB_SHEET_NAME}:`, e);
  }

  // 3. Delete from other specified sheets
  const otherSheetNames = [REPORT_LIST_SHEET_NAME, DEFAULT_USERS_SHEET_NAME, FORM_RESPONSES_SHEET_NAME];
  otherSheetNames.forEach(sheetName => {
    try {
        const sheet = ss.getSheetByName(sheetName);
        if (sheet) { 
          const data = sheet.getDataRange().getDisplayValues(); 
          const columnIndex = (sheetName === FORM_RESPONSES_SHEET_NAME) ? 1 : 0; 
          
          for (let i = data.length - 1; i >= 0; i--) {
            const currentReportName = (data[i][columnIndex] || '').toString().trim();
            if (currentReportName === reportNameToDelete) {
              sheet.deleteRow(i + 1);
              Logger.log(`deleteReport: Deleted row ${i + 1} for "${reportNameToDelete}" from ${sheetName}.`);
              
              if (sheetName !== FORM_RESPONSES_SHEET_NAME) {
                  break;
              }
            }
          }
        } else {
            Logger.log(`deleteReport: Sheet "${sheetName}" not found. Skipping deletion for this sheet.`);
        }
    } catch(e) {
        Logger.log(`deleteReport: Error deleting from ${sheetName}: ${e}`);
        console.error(`Error deleting from ${sheetName}:`, e);
    }
  });
  Logger.log(`deleteReport: Finished deletion process for report "${reportNameToDelete}".`);
}

/**
 * Sends a notification email to a user when they are assigned a new task.
 * @param {string} recipientEmail - The email address of the user being assigned.
 * @param {string} reportName - The name of the report.
 * @param {string} adminEmail - The email of the admin who assigned the task.
 */
function sendAssignmentEmail(recipientEmail, reportName, adminEmail) {
  if (!recipientEmail || !reportName || !adminEmail) {
    Logger.log("Missing data for assignment email. Skipping.");
    return;
  }

  const appUrl = ScriptApp.getService().getUrl();
  const subject = `New Task Assignment: "${reportName}"`;

  const logoFileId = 'YOUR_LOGO_DRIVE_FILE_ID';
  const logoBlob = DriveApp.getFileById(logoFileId).getBlob().setName('logo.png');

  const htmlBody = `
    <div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
      <div style="max-width: 600px; margin: 20px auto; border: 1px solid #ddd; border-radius: 8px; overflow: hidden;">
        
        <div style="background-color: #2B5797; color: #fff; padding: 20px; text-align: center;">
          <img src="cid:logo" alt="Company Logo" style="max-height: 45px; margin-bottom: 10px;">
          <h2 style="margin: 0; font-size: 22px;">New Task Assigned</h2>
        </div>
        
        <div style="padding: 25px;">
          <p>Hello,</p>
          <p>You have been assigned a new task:</p>
          <p style="font-size:1.2em; font-weight:bold; margin: 10px 0;">${reportName}</p>
          <p>Please review it in the workflow management tool.</p>
          
          <div style="text-align: center; margin: 30px 0;">
            <a href="${appUrl}" style="background-color: #6366f1; color: white; text-decoration: none; padding: 12px 24px; border-radius: 6px; font-weight: bold; display: inline-block;">
              Access Workflow Tool
            </a>
          </div>

          <p style="font-size:0.9em; color:#888; border-top:1px solid #eee; padding-top:15px;">
            This task was assigned to you by: ${adminEmail}
          </p>
        </div>
        
        <div style="background-color: #f9f9f9; padding: 15px 25px; font-size: 0.9em; color: #555; border-top: 1px solid #eee;">
          <p>Thank you,<br>Company MI Workflow Management</p>
        </div>
      </div>
    </div>
  `;

  try {
    MailApp.sendEmail({
      to: recipientEmail,
      subject: subject,
      htmlBody: htmlBody,
      inlineImages: { logo: logoBlob }, 
      noReply: true,
      name: 'Company Workflow (No Reply)',
      replyTo: adminEmail 
    });
    Logger.log(`Assignment email successfully sent to ${recipientEmail}`);
  } catch (e) {
    Logger.log(`Failed to send assignment email to ${recipientEmail} for report ${reportName}. Error: ${e}`);
  }
}

/**
 * Sends a daily task summary email with embedded image and HTML formatting.
 * Uses MailApp with noReply:true so it sends from the domain's NoReply address.
 *
 * @param {string} recipientEmail - The recipient’s email address.
 * @param {string[]} tasks - Array of task names.
 * @param {string} appUrl - Link to the workflow tool.
 * @param {string} dateString - e.g. "21/10/2025".
 */
function sendDailyTaskEmail(recipientEmail, tasks, appUrl, dateString) {
  const subject = `Your Daily Workflow Tasks - ${dateString}`;
  
  const plainTextBody =
    `Hello,\n\nHere are your assigned workflow tasks for today (${dateString}):\n` +
    tasks.map(t => ` • ${t}`).join('\n') +
    `\n\nPlease access the workflow tool to complete them:\n${appUrl}\n\nThank you,\nCompany MI Workflow Management`;
    
  const htmlTaskList = tasks.map(t => `<li style="margin-bottom:6px;"><b>${t}</b></li>`).join('');
  const logoFileId = 'YOUR_LOGO_DRIVE_FILE_ID';
  const logoBlob = DriveApp.getFileById(logoFileId).getBlob().setName('logo.png');
  
  const htmlBody = `
  <div style="font-family:Arial,sans-serif;line-height:1.6;margin:0;padding:0;color:#333;">
    <div style="width:90%;max-width:600px;margin:20px auto;border:1px solid #ddd;border-radius:8px;overflow:hidden;box-shadow:0 2px 5px rgba(0,0,0,0.05);">
      
      <div style="background-color:#2B5797;color:#fff;padding:20px;text-align:center;">
        <img src="cid:logo" alt="Company Logo" style="max-height:40px;margin-bottom:10px;"><br>
        <h2 style="margin:0;font-size:22px;">Your Daily Tasks</h2>
      </div>

      <div style="padding:25px;">
        <p>Hello,</p>
        <p>Here are your assigned workflow tasks for today (${dateString}):</p>
        <ul style="padding-left:20px;margin-bottom:20px;">${htmlTaskList}</ul>
        <p>Please access the workflow tool to complete them.</p>
        <div style="text-align:center;margin:25px 0 15px;">
          <a href="${appUrl}" style="background-color:#6366f1;color:#fff;text-decoration:none;padding:12px 25px;border-radius:5px;font-weight:bold;display:inline-block;">Go to Workflow Tool</a>
        </div>
      </div>

      <div style="font-size:0.9em;color:#888;padding:15px 25px;background-color:#f9f9f9;border-top:1px solid #eee;">
        <p>Thank you,<br>Company MI Workflow Management</p>
        <p style="font-size:0.8em;color:#aaa;margin-top:10px;border-top:1px solid #eee;padding-top:10px;">
          Please do not reply to this email — this inbox is not monitored.
        </p>
      </div>

    </div>
  </div>`.trim();
  
  try {
    MailApp.sendEmail({
      to: recipientEmail,
      subject: subject,
      body: plainTextBody, 
      htmlBody: htmlBody,
      inlineImages: { logo: logoBlob },
      name: 'Company Workflow (No Reply)',
      noReply: true // ensures it sends from noreply@company.com
    });
    Logger.log(`✅ Sent daily task email to ${recipientEmail} (${tasks.length} tasks).`);
  } catch (e) {
    Logger.log(`❌ Failed to send email to ${recipientEmail}: ${e}`);
  }
}

/**
 * Sends a notification email to admins about a newly added user.
 * @param {string[]} adminEmails - Array of admin email addresses.
 * @param {string} newUserEmail - The email address of the new user.
 * @param {string} newUserRole - The role assigned to the new user ('Admin' or 'User').
 * @param {string} appUrl - The URL of the web app.
 * @param {string} addedByAdminEmail - The email of the admin who added the user.
 */
function sendAdminNewUserNotification(adminEmails, newUserEmail, newUserRole, appUrl, addedByAdminEmail) {
  if (!adminEmails || adminEmails.length === 0 || !newUserEmail || !addedByAdminEmail) {
    Logger.log("Skipping admin notification email due to missing info.");
    return;
  }

  const subject = `New User Added to Workflow Tool: ${newUserEmail}`;
  const recipients = adminEmails.join(',');
  
  const logoFileId = 'YOUR_LOGO_DRIVE_FILE_ID';
  const logoBlob = DriveApp.getFileById(logoFileId).getBlob().setName('logo.png');

  const htmlBody = `
    <div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
      <div style="max-width: 600px; margin: 20px auto; border: 1px solid #ddd; border-radius: 8px; overflow: hidden;">
        
        <div style="background-color: #2B5797; color: #fff; padding: 20px; text-align: center;">
          <img src="cid:logo" alt="Company Logo" style="max-height: 45px; margin-bottom: 10px;">
          <h2 style="margin: 0; font-size: 22px;">New User Notification</h2>
        </div>
        
        <div style="padding: 25px;">
          <p>Hello Admins,</p>
          <p>A new user has been added to the workflow tool by <b>${addedByAdminEmail}</b>.</p>
          
          <div style="background-color:#f4f4f4; padding: 15px; border-radius: 5px; margin: 20px 0;">
            <p style="margin: 5px 0; font-size: 1.1em;"><b>Email:</b> ${newUserEmail}</p>
            <p style="margin: 5px 0; font-size: 1.1em;"><b>Role:</b> ${newUserRole}</p>
          </div>

          <p>You can manage users or review this change in the tool.</p>
          
          <div style="text-align: center; margin: 30px 0;">
            <a href="${appUrl}" style="background-color: #6366f1; color: white; text-decoration: none; padding: 12px 24px; border-radius: 6px; font-weight: bold; display: inline-block;">
              Access Workflow Tool
            </a>
          </div>
        </div>
        
        <div style="background-color: #f9f9f9; padding: 15px 25px; font-size: 0.9em; color: #555; border-top: 1px solid #eee;">
          <p>Thank you,<br>Company MI Workflow Management</p>
        </div>
      </div>
    </div>
  `;

  try {
    MailApp.sendEmail({
      to: recipients,
      subject: subject,
      htmlBody: htmlBody,
      inlineImages: { logo: logoBlob },
      noReply: true,
      name: 'Company Workflow (No Reply)'
    });
    Logger.log(`Admin notification for new user ${newUserEmail} sent to ${recipients}`);
  } catch (e) {
    Logger.log(`Failed to send admin notification for ${newUserEmail}: ${e}`);
  }
}

/**
 * Triggered automatically when the "New Reports" sheet is changed.
 * Sends a notification email to all Admin users about the new report row.
 */
function onNewReportSheetChange(e) {
  Logger.log("onNewReportSheetChange triggered.");
  
  const range = e.range;
  const sheet = range.getSheet();
  const editedRow = range.getRow();
  const sheetName = sheet.getName();

  const columnIndices = {
    reportName: 1,  
    assignedUser: 2, 
    reportType: 3,     
    dueDate: 4     
  };
  
  const firstDataRow = 2; 

  if (sheetName !== NEW_REPORTS_SHEET_NAME || editedRow < firstDataRow || range.getWidth() < Object.keys(columnIndices).length || !e.value ) { 
     if (sheetName === NEW_REPORTS_SHEET_NAME) {
         Logger.log(`Ignoring edit on sheet "${sheetName}" row ${editedRow}. Might be header, deletion, or insufficient data.`);
     }
    return; 
  }
  
  Logger.log(`Processing edit on sheet "${sheetName}" row ${editedRow}.`);
  let adminEmails = [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    const rowValues = sheet.getRange(editedRow, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    Logger.log("Edited row values: " + JSON.stringify(rowValues));

    const reportName = rowValues[columnIndices.reportName - 1] || 'N/A';
    const assignedUser = rowValues[columnIndices.assignedUser - 1] || 'N/A';
    const reportType = rowValues[columnIndices.reportType - 1] || 'N/A';
    let dueDate = rowValues[columnIndices.dueDate - 1] || 'N/A';

     if (reportName === 'N/A' || !reportName.trim()) {
        Logger.log("Ignoring row - Report Name is missing or empty.");
        return;
     }

    if (dueDate !== 'N/A') {
        try {
              let dateObj;
              if (typeof parseDateString === 'function') {
                  dateObj = parseDateString(dueDate);
              } else {
                  dateObj = new Date(dueDate);
              }
            if (dateObj && !isNaN(dateObj.getTime())) {
                dueDate = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "dd/MM/yyyy");
            }
        } catch (dateErr) {
            Logger.log("Could not format due date: " + dueDate);
        }
    }

    Logger.log(`New/Edited Report: Name="${reportName}", Assigned="${assignedUser}", Due="${dueDate}"`);
    
    const userSheet = ss.getSheetByName(USERS_SHEET_NAME);
    if (!userSheet) {
      Logger.log(`onNewReportSheetChange Error: Sheet "${USERS_SHEET_NAME}" not found.`);
      return;
    }
    
    const userData = userSheet.getRange('A2:B' + userSheet.getLastRow()).getValues();
    adminEmails = userData
      .filter(row => row[1] && row[1].toLowerCase() === 'admin' && row[0])
      .map(row => row[0].trim());
      
    if (adminEmails.length === 0) {
      Logger.log("onNewReportSheetChange Warning: No 'Admin' users found to email.");
      return;
    }
    
    Logger.log("Admins to notify: " + adminEmails.join(', '));
    
    const appUrl = ScriptApp.getService().getUrl();
    const subject = `New Report Added/Edited in '${NEW_REPORTS_SHEET_NAME}': ${reportName}`;
    const body = `
      <p>Hello Admins,</p>
      <p>A report has been added or edited in the "${NEW_REPORTS_SHEET_NAME}" sheet:</p>
      <ul>
        <li><b>Report Name:</b> ${reportName}</li>
        <li><b>Assigned User:</b> ${assignedUser}</li>
        <li><b>Report Type:</b> ${reportType}</li>
        <li><b>Due Date:</b> ${dueDate}</li>
      </ul>
      <p>This may require review or further action in the workflow tool.</p>
      
      <p>You can view and manage reports in the tool:</p>
      <p><a href="${appUrl}">Workflow Management Tool</a></p>
      <br>
      <p>Thank you,</p>
      <p>Company MI Workflow Management</p>
    `;
    
    MailApp.sendEmail(adminEmails.join(','), subject, "", {
      noReply: true,
      htmlBody: body,
      name: 'Company Workflow (No Reply)'
    });
    Logger.log("Notification email sent successfully for row " + editedRow);

  } catch (error) {
    Logger.log(`onNewReportSheetChange Error processing row ${editedRow}: ${error.toString()} Stack: ${error.stack}`);
  }
}

/**
 * Sends a notification email to Admin users about a newly routed report.
 * @param {string} reportName The name of the new report.
 * @param {string} assignedUser The user assigned in the 'New Reports' sheet.
 * @param {string} dueDate The due date string from the 'New Reports' sheet.
 */
function sendNewReportAdminNotification(reportName, assignedUser, dueDate) {
  Logger.log(`Attempting to send notification for new report: ${reportName}`);
  let adminEmails = [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    const userSheet = ss.getSheetByName(USERS_SHEET_NAME);
    if (!userSheet) {
      Logger.log(`sendNewReportAdminNotification Error: Sheet "${USERS_SHEET_NAME}" not found.`);
      return;
    }
    const userData = userSheet.getRange('A2:B' + userSheet.getLastRow()).getValues();
    adminEmails = userData
      .filter(row => row[1] && row[1].toLowerCase() === 'admin' && row[0])
      .map(row => row[0].trim());
      
    if (adminEmails.length === 0) {
      Logger.log("sendNewReportAdminNotification Warning: No 'Admin' users found to email.");
      return;
    }
    Logger.log("Admins to notify: " + adminEmails.join(', '));
    
    const appUrl = ScriptApp.getService().getUrl();
    reportName = reportName || 'N/A';
    assignedUser = assignedUser || 'N/A';
    dueDate = dueDate || 'N/A';
    
    if (dueDate !== 'N/A') {
        try {
            let dateObj;
              if (typeof parseDateString === 'function') { 
                  dateObj = parseDateString(dueDate);
              } else {
                  dateObj = new Date(dueDate);
              }
            if (dateObj && !isNaN(dateObj.getTime())) {
                dueDate = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "dd/MM/yyyy");
            }
        } catch (dateErr) {
            Logger.log("Could not format due date for email: " + dueDate);
        }
    }

    const subject = `New Workflow Report Created & Routed: ${reportName}`;
    const body = `
      <p>Hello Admins,</p>
      <p>A new report submitted via the form has been automatically routed to 'Report List' and 'Default Users':</p>
      <ul>
        <li><b>Report Name:</b> ${reportName}</li>
        <li><b>Default Assigned User:</b> ${assignedUser}</li>
        <li><b>Due Date:</b> ${dueDate}</li>
      </ul>
      <p>You can view and manage reports in the tool:</p>
      <p><a href="${appUrl}">Workflow Management Tool</a></p>
      <br>
      <p>Thank you,</p>
      <p>Company MI Workflow Management</p>
    `;
    
    MailApp.sendEmail(adminEmails.join(','), subject, "", {
      noReply: true,
      htmlBody: body,
      name: 'Company Workflow (No Reply)'
    });
    Logger.log("New report notification email sent successfully.");

  } catch (error) {
    Logger.log(`sendNewReportAdminNotification Error: ${error.toString()} Stack: ${error.stack}`);
  }
}

// --- CONFIGURATION FOR DROPDOWN ---
const FORM_ID = 'YOUR_FORM_ID_HERE';
const DROPDOWN_TITLE = 'By default who is assigned this work? (email address)';
const USER_SHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';
// --- END CONFIGURATION ---

/**
 * Fetches users from the Sheet and updates a dropdown in a specific Google Form.
 * This is called by other functions (like addUser) when the user list changes.
 */
function updateDropdownFromSheet() {
  try {
    const sheet = SpreadsheetApp.openById(USER_SHEET_ID).getSheetByName('Users');
    const values = sheet.getRange('A2:A').getValues();

    const userList = values
      .map(row => row[0])        
      .filter(item => item !== "");

    if (userList.length === 0) {
      Logger.log('updateDropdownFromSheet: No users found in range.');
      return; 
    }

    const form = FormApp.openById(FORM_ID);
    let dropdownItem = null;
    const items = form.getItems(FormApp.ItemType.LIST);

    for (let i = 0; i < items.length; i++) {
      if (items[i].getTitle() === DROPDOWN_TITLE) {
        dropdownItem = items[i].asListItem();
        break; 
      }
    }

   if (dropdownItem) {
      dropdownItem.setChoiceValues(userList);
      Logger.log('Dropdown successfully updated with ' + userList.length + ' users.');
    } else {
      Logger.log('Error: Could not find dropdown question with title: ' + DROPDOWN_TITLE);
    }

  } catch (e) {
    Logger.log('Error in updateDropdownFromSheet: ' + e.message);
  }
}

// -------------------------------------------------------------------------------------------------
// --- NEW: WORKFLOW REPORT MANAGEMENT (adds functionality for new HTML tab) ---
// -------------------------------------------------------------------------------------------------

/**
 * Returns report names (Col A) and URLs (Col K) from the 'Report List' sheet.
 */
function getWorkflowReportList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(REPORT_LIST_SHEET_NAME);
  if (!sheet) throw new Error(`Sheet "${REPORT_LIST_SHEET_NAME}" not found.`);
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange('A2:K' + lastRow).getValues();
  return data.map(row => ({
    reportName: row[0],
    reportLink: row[10] || ''
  })).filter(r => r.reportName);
}

/**
 * Updates only Column K (the report URL) for a given report name in Column A.
 */
function updateReportLink({ reportName, newUrl }) {
  const userInfo = getUserInfo();
  if (!userInfo.isAdmin) throw new Error("Permission denied. Only admins can update report URLs.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(REPORT_LIST_SHEET_NAME);
  if (!sheet) throw new Error(`Sheet "${REPORT_LIST_SHEET_NAME}" not found.`);

  const names = sheet.getRange('A2:A' + sheet.getLastRow()).getValues().flat();
  const idx = names.findIndex(n => n && n.toString().trim() === reportName.trim());
  
  if (idx === -1) throw new Error('Report not found: ' + reportName);

  sheet.getRange(idx + 2, 11).setValue(newUrl);
  Logger.log(`Updated report URL for "${reportName}" to "${newUrl}"`);
  return true;
}

/**
 * Gets all links from the "Useful Links" sheet.
 * @returns {object[]} An array of objects, e.g., [{name: 'Google', url: 'https://google.com'}].
 */
function getUsefulLinks() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USEFUL_LINKS_SHEET_NAME);
    if (!sheet) {
      const newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(USEFUL_LINKS_SHEET_NAME);
      newSheet.appendRow(['Link Name', 'URL']);
      return []; 
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return []; 

    return sheet.getRange('A2:B' + lastRow).getValues().map(row => ({
      name: row[0],
      url: row[1] || ''
    })).filter(link => link.name);
    
  } catch (e) {
    console.error('Error in getUsefulLinks:', e);
    throw new Error('Could not retrieve useful links. Please ensure a "Useful Links" sheet exists.');
  }
}

/**
 * Updates the URL for a specific link name in "Useful Links".
 * @param {object} linkData - {name: 'Link Name', newUrl: 'https://newurl.com'}.
 */
function updateUsefulLink(linkData) {
  if (!linkData || !linkData.name) {
    throw new Error("Invalid link data provided.");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USEFUL_LINKS_SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${USEFUL_LINKS_SHEET_NAME}" not found.`);

    const names = sheet.getRange('A2:A' + sheet.getLastRow()).getDisplayValues().flat();
    const rowIndex = names.findIndex(name => name.trim().toLowerCase() === linkData.name.trim().toLowerCase());

    if (rowIndex === -1) {
      throw new Error(`Link "${linkData.name}" not found.`);
    }

    sheet.getRange(rowIndex + 2, 2).setValue(linkData.newUrl || '');
    return "Link updated.";
    
  } catch (e) {
    console.error('Error in updateUsefulLink:', e);
    throw new Error(e.message || 'Failed to update link.');
  } finally {
    lock.releaseLock();
  }
}

/**
 * Adds a new link to the "Useful Links" sheet.
 * @param {object} linkData - {name: 'New Link', url: 'https://url.com'}.
 */
function addUsefulLink(linkData) {
  if (!linkData || !linkData.name || !linkData.url) {
    throw new Error("Both link name and URL are required.");
  }
  
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USEFUL_LINKS_SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${USEFUL_LINKS_SHEET_NAME}" not found.`);
    
    const names = sheet.getRange('A2:A' + sheet.getLastRow()).getDisplayValues().flat();
    const exists = names.some(name => name.trim().toLowerCase() === linkData.name.trim().toLowerCase());
    
    if (exists) {
      throw new Error(`A link with the name "${linkData.name}" already exists.`);
    }

    sheet.appendRow([linkData.name, linkData.url]);
    return "Link added successfully.";
  } catch (e) {
    console.error('Error in addUsefulLink:', e);
    throw new Error(e.message || 'Failed to add link.');
  } finally {
    lock.releaseLock();
  }
}

/**
 * Deletes a link from the "Useful Links" sheet.
 * @param {string} linkName - The name of the link to delete.
 */
function deleteUsefulLink(linkName) {
  if (!linkName) {
    throw new Error("Link name is required.");
  }
  
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USEFUL_LINKS_SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${USEFUL_LINKS_SHEET_NAME}" not found.`);

    const names = sheet.getRange('A2:A' + sheet.getLastRow()).getDisplayValues().flat();
    const rowIndex = names.findIndex(name => name.trim().toLowerCase() === linkName.trim().toLowerCase());

    if (rowIndex === -1) {
      throw new Error(`Link "${linkName}" not found.`);
    }

    sheet.deleteRow(rowIndex + 2); 
    return "Link deleted.";
  } catch (e) {
    console.error('Error in deleteUsefulLink:', e);
    throw new Error(e.message || 'Failed to delete link.');
  } finally {
    lock.releaseLock();
  }
}

/**
 * Gets a list of user emails from the 'Users' sheet, EXCLUDING admins.
 * @returns {string[]} An array of non-admin user email addresses.
 */
function getNonAdminUsers() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USERS_SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${USERS_SHEET_NAME}" not found.`);

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    return sheet.getRange('A2:B' + lastRow).getValues()
      .filter(row => row[1] !== 'Admin' && row[0]) 
      .map(row => row[0]);
      
  } catch (e) {
    console.error('Error in getNonAdminUsers:', e);
    throw new Error('Could not retrieve non-admin user list.');
  }
}

/**
 * Automatically cleans log entries older than 6 months from the 'Log' sheet.
 * This should be run on a time-driven trigger (e.g., monthly).
 */
function cleanOldLogs() {
  const LOG_RETENTION_MONTHS = 6;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
  
  if (!sheet) {
    Logger.log("cleanOldLogs: Log sheet not found. Exiting.");
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { 
    Logger.log("cleanOldLogs: Log sheet is empty. Exiting.");
    return;
  }

  const now = new Date();
  const cutoffDate = new Date(now.getFullYear(), now.getMonth() - LOG_RETENTION_MONTHS, now.getDate());
  const cutoffTime = cutoffDate.getTime();
  Logger.log(`cleanOldLogs: Deleting log entries older than ${cutoffDate.toISOString()}`);
  
  const timestamps = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  let rowsDeleted = 0;
  
  for (let i = timestamps.length - 1; i >= 0; i--) {
    try {
      const timestamp = new Date(timestamps[i][0]);
      if (!isNaN(timestamp.getTime()) && timestamp.getTime() < cutoffTime) {
        sheet.deleteRow(i + 2);
        rowsDeleted++;
      }
    } catch (e) {
      Logger.log(`cleanOldLogs: Error parsing date in row ${i + 2}. Error: ${e}`);
    }
  }
  Logger.log(`cleanOldLogs: Completed. Deleted ${rowsDeleted} old log entries.`);
}

/**
 * Sends a daily summary email at 3 PM to all Admins.
 * UPDATED: 'Outstanding' items are now at the top for better visibility.
 * 'Completed' items are at the bottom and visually de-emphasized.
 */
function sendDailyReportSummary() {
  // --- 1. Weekend check ---
  const now = new Date();
  const dayOfWeek = now.getDay();
  if (dayOfWeek === 0 || dayOfWeek === 6) {
    Logger.log("Skipping daily summary: It's the weekend.");
    return;
  }

  const timeZone = Session.getScriptTimeZone();
  const today = Utilities.formatDate(now, timeZone, "yyyy-MM-dd");
  const todayDisplay = Utilities.formatDate(now, timeZone, "dd/MM/yyyy");
  
  // --- 2. Get Admin Emails ---
  const userSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USERS_SHEET_NAME);
  if (!userSheet) return;
  
  const userData = userSheet.getRange('A2:B' + userSheet.getLastRow()).getValues();
  const adminEmails = userData
    .filter(row => row[1] && row[1].toLowerCase() === 'admin' && row[0])
    .map(row => row[0]);
    
  if (adminEmails.length === 0) return;

  // --- 3. Get Report Data ---
  let allReports;
  try {
    allReports = getReportData(); 
  } catch (e) {
    MailApp.sendEmail(adminEmails.join(','), `FAILED: Daily Report Summary`, `Error: ${e}`);
    return;
  }

  // --- 4. Sort Reports ---
  const completedReports = [];
  const outstandingReports = [];
  
  allReports.forEach(report => {
    if (report.nextRunDate === today) {
      const assignedUser = report.allocatedTo || 'Unassigned';
      const comments = report.comments || '';
      const dueTime = report.dueTime || '';

      if (report.completedDate) {
        const timeMatch = report.completedDate.match(/(\d{2}:\d{2}):\d{2}$/);
        const timePart = timeMatch ? timeMatch[1] : 'Unknown Time';
        completedReports.push({ name: report.reportName, completedTime: timePart, assignedTo: assignedUser, comments });
      } else {
        outstandingReports.push({ name: report.reportName, assignedTo: assignedUser, comments, dueTime });
      }
    }
  });
  
  // --- 5. Assets ---
  const publicLogoUrl = 'https://via.placeholder.com/150x50?text=Company+Logo';
  let logoBlob = null;
  try {
      logoBlob = UrlFetchApp.fetch(publicLogoUrl).getBlob().setName('logo.png');
  } catch (e) { logoBlob = null; }

  // --- 6. Build HTML (LAYOUT UPDATED) ---
  
  // Outstanding List (Actionable items first)
  const outstandingList = outstandingReports.length > 0
      ? `<ul style="padding-left:20px; margin-top:5px;">${outstandingReports
          .map(r => `<li style="margin-bottom:8px;">
                      <strong>${r.name}</strong><br>
                      <span style="font-size:0.9em; color:#d32f2f;">Assigned: ${r.assignedTo} ${r.dueTime ? '(Due: '+r.dueTime+')' : ''}</span>
                      ${r.comments ? `<br><span style="font-size:0.85em; color:#555; font-style:italic;">Note: ${r.comments}</span>` : ''}
                   </li>`)
          .join('')}</ul>`
      : `<p style="color:#2e7d32; font-style:italic;">✅ All reports due today are complete. Great job!</p>`;

  // Completed List (De-emphasized style)
  const completedList = completedReports.length > 0
      ? `<ul style="padding-left:20px; color:#666; font-size:0.9em;">${completedReports
          .map(r => `<li style="margin-bottom:4px;">${r.name} - ${r.assignedTo} <span style="font-size:0.8em; color:#999;">(${r.completedTime})</span></li>`)
          .join('')}</ul>`
      : `<p style="color:#666; font-style:italic;">No reports completed yet today.</p>`;
      
  const htmlBody = `
  <div style="font-family:Segoe UI, Arial, sans-serif; line-height:1.5; color:#333;">
    <div style="max-width:650px; margin:0 auto; border:1px solid #e0e0e0; border-radius:8px; overflow:hidden;">
      
      <div style="background-color:#004B87; color:#fff; padding:20px; text-align:center;">
        ${logoBlob ? '<img src="cid:logo" alt="Company" style="height:35px; margin-bottom:10px;"><br>' : ''}
        <h2 style="margin:0; font-size:20px; font-weight:600;">Daily Status Report</h2>
        <p style="margin:5px 0 0; font-size:14px; opacity:0.9;">${todayDisplay}</p>
      </div>

      <div style="padding:25px;">
        
        <h3 style="color:#D83B01; border-bottom:2px solid #D83B01; padding-bottom:5px; margin-top:0;">
          ⚠️ Outstanding Reports (${outstandingReports.length})
        </h3>
        ${outstandingList}

        <br>

        <div style="text-align:center; margin:20px 0;">
          <a href="${WEB_APP_URL}" style="background-color:#004B87; color:#fff; text-decoration:none; padding:12px 30px; border-radius:4px; font-weight:bold; font-size:14px;">
            Open Workflow Tool
          </a>
        </div>

        <br>

        <h3 style="color:#2e7d32; border-bottom:1px solid #e0e0e0; padding-bottom:5px; font-size:16px;">
          ✅ Completed Reports (${completedReports.length})
        </h3>
        <div style="background-color:#f9f9f9; padding:10px; border-radius:4px;">
          ${completedList}
        </div>

      </div>

      <div style="background-color:#f5f5f5; color:#888; padding:15px; text-align:center; font-size:12px;">
        Company MI Workflow Automation
      </div>
    </div>
  </div>`;

  // --- 7. Send ---
  MailApp.sendEmail({
    to: adminEmails.join(','),
    subject: `Daily Report Summary - ${todayDisplay}`,
    htmlBody: htmlBody,
    inlineImages: logoBlob ? { logo: logoBlob } : {},
    name: 'Company Workflow',
    noReply: true
  });
}

/**
 * Sends a welcome email to a newly added user with an inline logo.
 * @param {string} newUserEmail - The email address of the new user.
 * @param {string} appUrl - The URL of the web app.
 * @param {string} adminEmail - The email of the admin who added the user.
 */
function sendNewUserWelcomeEmail(newUserEmail, appUrl, adminEmail) {
  if (!newUserEmail || !appUrl || !adminEmail) {
    Logger.log("Skipping welcome email due to missing info.");
    return;
  }

  const subject = "Welcome to the Company Workflow Tool";

  const logoFileId = 'YOUR_LOGO_DRIVE_FILE_ID'; 
  const logoBlob = DriveApp.getFileById(logoFileId).getBlob().setName('logo.png');

  const htmlBody = `
    <div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
      <div style="max-width: 600px; margin: 20px auto; border: 1px solid #ddd; border-radius: 8px; overflow: hidden;">
        
        <div style="background-color: #2B5797; color: #fff; padding: 20px; text-align: center;">
          <img src="cid:logo" alt="Company Logo" style="max-height: 45px; margin-bottom: 10px;">
          <h2 style="margin: 0; font-size: 22px;">Welcome to the Workflow Tool</h2>
        </div>
        
        <div style="padding: 25px;">
          <p>Hello,</p>
          <p>You have been added to the <b>Company MI Workflow Management Tool</b> by <b>${adminEmail}</b>.</p>
          <p>You can use this tool to view, manage, and complete your assigned workflow tasks.</p>
          
          <div style="text-align: center; margin: 30px 0;">
            <a href="${appUrl}" style="background-color: #6366f1; color: white; text-decoration: none; padding: 12px 24px; border-radius: 6px; font-weight: bold; display: inline-block;">
              Access Workflow Tool
            </a>
          </div>
        </div>
        
        <div style="background-color: #f9f9f9; padding: 15px 25px; font-size: 0.9em; color: #555; border-top: 1px solid #eee;">
          <p>Thank you,<br>Company MI Workflow Management</p>
        </div>
      </div>
    </div>
  `;

  try {
    MailApp.sendEmail({
      to: newUserEmail,
      subject: subject,
      htmlBody: htmlBody,
      inlineImages: { logo: logoBlob },
      noReply: true,
      name: 'Company Workflow (No Reply)',
    });
    Logger.log(`Welcome email successfully sent to ${newUserEmail}`);
  } catch (e) {
    Logger.log(`Failed to send welcome email to ${newUserEmail}: ${e}`);
  }
}

/**
 * Directly updates the 'Allocated To' column (B) in the 'Report List' sheet for a specific report.
 * Admin only.
 * @param {string|object} reportName The name of the report to update (or an object containing it).
 * @param {string} userEmail The email address to set in Column B. Can be empty to clear.
 */
function updateReportListAllocation(reportName, userEmail) {
  let reportNameString = reportName;
  if (typeof reportName === 'object' && reportName !== null && reportName.reportName) {
    Logger.log('updateReportListAllocation: Received an object, extracting .reportName property.');
    reportNameString = reportName.reportName;
  } else if (typeof reportName === 'object') {
    Logger.log(`updateReportListAllocation: Received an invalid object: ${JSON.stringify(reportName)}`);
    throw new Error("Invalid data received. Expected a report name string.");
  }

  if (!getUserInfo().isAdmin) {
    Logger.log(`updateReportListAllocation: Permission denied for user ${Session.getActiveUser().getEmail()} trying to update ${reportNameString}`);
    throw new Error("Permission denied. Only admins can directly update allocations.");
  }

  if (!reportNameString) {
     Logger.log("updateReportListAllocation: Called with empty report name.");
     throw new Error("Report name cannot be empty.");
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const reportListSheet = ss.getSheetByName(REPORT_LIST_SHEET_NAME);
    if (!reportListSheet) {
      Logger.log(`updateReportListAllocation: Sheet "${REPORT_LIST_SHEET_NAME}" not found.`);
      throw new Error(`Sheet "${REPORT_LIST_SHEET_NAME}" not found.`);
    }

    const reportNames = reportListSheet.getRange('A2:A' + reportListSheet.getLastRow()).getDisplayValues().flat();
    const searchReportName = reportNameString.trim();
    
    const rowIndex = reportNames.findIndex(name => (name || '').toString().trim() === searchReportName);
    
    if (rowIndex === -1) {
       Logger.log(`updateReportListAllocation: Report "${searchReportName}" not found in "Report List".`);
       throw new Error(`Report "${searchReportName}" not found in "Report List".`);
    }

    const rowToUpdate = rowIndex + 2;

    reportListSheet.getRange(rowToUpdate, ALLOCATED_TO_COLUMN).setValue(userEmail || '');
    Logger.log(`updateReportListAllocation: Successfully updated "${searchReportName}" (Row ${rowToUpdate}) to "${userEmail || ''}" in Report List.`);
  } catch (e) {
    Logger.log(`updateReportListAllocation: ERROR updating report "${reportNameString}" - ${e.toString()}`);
    console.error(`Error in updateReportListAllocation for report ${reportNameString}:`, e);
    throw new Error(e.message || 'Failed to update allocation in Report List.');
  }
}

/**
 * Runs overnight via a time-driven trigger.
 * 1. Calls resetDailyReports() to clear comments and completion data.
 * 2. Calls updateDailyAllocation() to run the full, smart allocation.
 */
function runOvernightMasterReset() {
  Logger.log("--- STARTING OVERNIGHT MASTER RESET ---");
  
  try {
    resetDailyReports();
    Logger.log("Step 1/2: resetDailyReports() completed. Comments and timestamps cleared.");
  } catch (e) {
    Logger.log(`CRITICAL: Step 1 (resetDailyReports) FAILED: ${e}. Aborting master reset.`);
    return; 
  }

  try {
    updateDailyAllocation();
    Logger.log("Step 2/2: updateDailyAllocation() completed. Allocations are set.");
  } catch (e) {
    Logger.log(`CRITICAL: Step 2 (updateDailyAllocation) FAILED: ${e}.`);
  }
  
  Logger.log("--- FINISHED OVERNIGHT MASTER RESET ---");
}

/**
 * Fetches data for the Key Info Dashboard from the "Dashboards" tab.
 * Returns 3 separate arrays for Demand, DSV, and Certificates.
 */
function getKeyInfoData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Dashboards");
    
    if (!sheet) {
      console.error("Sheet 'Dashboards' not found.");
      return null;
    }

    return {
      demand: sheet.getRange("A3:G40").getDisplayValues(),
      dsv: sheet.getRange("A44:F48").getDisplayValues(),
      certs: sheet.getRange("A52:F56").getDisplayValues()
    };
  } catch (e) {
    console.error("Error fetching Key Info Data:", e);
    throw new Error("Could not load dashboard data.");
  }
}
