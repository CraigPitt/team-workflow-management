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
    const ss = SpreadsheetApp.
