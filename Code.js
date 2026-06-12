/************ SPREADSHEET CONSTANTS ************/
const SPREADSHEET_ID = "1_W-kKZ6JBltDdwEOVhcq1PZwTPApShCQVnFTku8nf3k"; 
const LOG_SHEET_NAME = "Gemini_Logs";
const DATA_SHEET_NAME = "Sheet1"; 

/************ GEMINI MODELS ************/
const GEMINI_MODEL_PRIMARY = "models/gemini-2.5-flash"; 
const GEMINI_MODEL_FALLBACK1 = "models/gemini-3.5-flash";
const GEMINI_MODEL_FALLBACK2 = "models/gemini-2.0-flash";

var GLOBAL_AI_RAW_RESPONSE = "";

function setupAutoTriggers() {
  console.log("=== [SETUP] Initializing Automatic Triggers ===");
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger("processLinkedInJobsWithGemini").timeBased().everyDays(1).atHour(6).create();
  ScriptApp.newTrigger("processLinkedInJobsWithGemini").timeBased().everyDays(1).atHour(18).create();
  console.log("=== [SUCCESS] Triggers set for 6 AM and 6 PM ===");
}

function processLinkedInJobsWithGemini() {
  console.log("=== [START] Advanced Autopilot Script Started ===");
  cleanUpTemporaryHooks();

  var ss;
  try { ss = SpreadsheetApp.openById(SPREADSHEET_ID); } catch(err) { return; }
  var dataSheet = ss.getSheetByName(DATA_SHEET_NAME) || ss.getSheets()[0];
  
  if (dataSheet.getLastRow() === 0) {
    dataSheet.appendRow([
      "Timestamp", "Job Title (Clickable)", "Company", "Location", 
      "Workplace Type (Remote/Hybrid)", "Employment Type", "Experience Level", 
      "Key Skills Required", "Job Description Summary", "Gmail Source (IST Time)"
    ]);
    dataSheet.getRange("A1:J1").setFontWeight("bold").setBackground("#e6effa");
  }
  
  var apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!apiKey) return;

  var afterDateStr = getYesterdayFormattedDate();
  var searchQuery = 'label:linkdin-jobs is:unread after:' + afterDateStr;
  console.log("=== Running Query: " + searchQuery);
  
  var threads = GmailApp.search(searchQuery);
  console.log("=== Total Pending Unread Threads Found: " + threads.length + " ===");

  if (threads.length === 0) {
    logStep("QUEUE_FINISHED", "Aaj ke saare pending unread emails process ho chuke hain.");
    return; 
  }
  
  var targetThread = threads[0];
  var messages = targetThread.getMessages();

  for (var j = 0; j < messages.length; j++) {
    var message = messages[j];
    if (!message.isUnread()) continue;
    
    var mailDate = message.getDate();
    var emailBodyHtml = message.getBody(); 
    
    // FIX/UPDATE: Email ka exact time IST format me convert karna (e.g., "12-Jun-2026 06:18 AM")
    var istFormattedTime = Utilities.formatDate(mailDate, "GMT+5:30", "dd-MMM-yyyy hh:mm a");
    
    // Gmail Thread Link Setup
    var threadId = targetThread.getId();
    var gmailLinkUrl = "https://mail.google.com/mail/u/0/#inbox/" + threadId;
    
    // Naya Hyperlink Formula: Ab text ki jagah IST Time dikhega jo clickable hoga
    var gmailFormula = '=HYPERLINK("' + gmailLinkUrl + '", "' + istFormattedTime + ' ✉️")';

    var prompt = "Extract all job openings listed in this LinkedIn HTML email. For each job, extract the following fields:\n" +
                 "1. 'jobTitle'\n" +
                 "2. 'company'\n" +
                 "3. 'location'\n" +
                 "4. 'jobUrl' (The exact LinkedIn view or apply HTTP link for this specific job)\n" +
                 "5. 'workplaceType' (Remote, Hybrid, Onsite, or N/A)\n" +
                 "6. 'employmentType' (Full-time, Part-time, Contract, Internship, or N/A)\n" +
                 "7. 'experienceLevel' (Entry level, Associate, Mid-Senior, Director, or N/A)\n" +
                 "8. 'keySkills' (Provide a single plain text string of 3-4 key skills separated by commas, NOT an array list)\n" +
                 "9. 'summary' (A brief 1-2 sentence description or summary of the role)\n\n" +
                 "Return the result ONLY as a valid JSON array of objects. Do not wrap in markdown blocks. HTML Content:\n" + emailBodyHtml;
    
    var payload = {
      "contents": [{ "parts": [{ "text": prompt }] }],
      "generationConfig": { "responseMimeType": "application/json" }
    };
    
    var jsonResponseText = callGeminiWithModels(payload, apiKey);
    
    if (jsonResponseText) {
      try {
        var jobsArray = JSON.parse(jsonResponseText);
        
        for (var k = 0; k < jobsArray.length; k++) {
          var job = jobsArray[k];
          var jobTitleValue = job.jobTitle || "View Job";
          var cellFormula = job.jobUrl ? '=HYPERLINK("' + job.jobUrl + '", "' + jobTitleValue.replace(/"/g, '""') + '")' : jobTitleValue;
          
          var cleanSkills = "";
          if (job.keySkills) {
            if (Array.isArray(job.keySkills)) {
              cleanSkills = job.keySkills.join(", ");
            } else {
              cleanSkills = String(job.keySkills);
            }
          } else {
            cleanSkills = "N/A";
          }
          
          dataSheet.appendRow([
            mailDate, 
            cellFormula, 
            job.company || "N/A", 
            job.location || "N/A",
            job.workplaceType || "N/A", 
            job.employmentType || "N/A", 
            job.experienceLevel || "N/A",
            cleanSkills, 
            job.summary || "N/A",
            gmailFormula // Column J me ab clickable IST timestamp jayega
          ]);
        }
        logStep("SUCCESS_EMAIL", jobsArray.length + " jobs extracted.");
      } catch(e) {
        logStep("JSON_PARSE_ERROR", e.message);
      }
    }
  }

  targetThread.markRead();

  // Next Hook Check
  var remainingThreads = GmailApp.search(searchQuery);
  if (remainingThreads.length > 0) {
    ScriptApp.newTrigger("processLinkedInJobsWithGemini_Hook").timeBased().after(60 * 1000).create();
    logStep("CHAIN_HOOK", remainingThreads.length + " threads remaining. Temporary hook set.");
  } else {
    logStep("CHAIN_END", "Saara queue clear ho gaya.");
  }
}

function processLinkedInJobsWithGemini_Hook() {
  processLinkedInJobsWithGemini();
}

function cleanUpTemporaryHooks() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "processLinkedInJobsWithGemini_Hook") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/************ GEMINI API HELPERS ************/
function callGeminiWithModels(payload, apiKey) {
  var models = [GEMINI_MODEL_PRIMARY, GEMINI_MODEL_FALLBACK1, GEMINI_MODEL_FALLBACK2];
  for (var m = 0; m < models.length; m++) {
    var modelName = models[m];
    var url = "https://generativelanguage.googleapis.com/v1beta/" + modelName + ":generateContent?key=" + apiKey;
    var response = callGeminiWithRetry(url, {
      method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true
    }, 3); 
    if (response) {
      var statusCode = response.getResponseCode();
      var rawText = response.getContentText();
      if (statusCode === 200) {
        try {
          var parsedRes = JSON.parse(rawText);
          if (parsedRes.candidates && parsedRes.candidates[0].content.parts[0].text) {
            return parsedRes.candidates[0].content.parts[0].text.trim();
          }
        } catch(err) { logStep("PARSE_ATTEMPT_ERR", err.toString()); }
      }
    }
  }
  return null;
}

function callGeminiWithRetry(url, options, maxRetries) {
  var attempts = 0;
  while (attempts < maxRetries) {
    try {
      var res = UrlFetchApp.fetch(url, options);
      if (res.getResponseCode() === 200 || attempts === maxRetries - 1) return res;
    } catch(e) { logStep("FETCH_EXCEPTION", e.toString()); }
    attempts++;
    Utilities.sleep(2000); 
  }
  return null;
}

function logStep(tag, message) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var logSheet = ss.getSheetByName(LOG_SHEET_NAME);
    if (!logSheet) {
      logSheet = ss.insertSheet(LOG_SHEET_NAME);
      logSheet.appendRow(["Timestamp", "Tag/Step Name", "Details / Payload"]);
    }
    var displayMessage = (typeof message === 'object') ? JSON.stringify(message, null, 2) : message;
    logSheet.getRange(logSheet.getLastRow() + 1, 1, 1, 3).setValues([[new Date(), tag, displayMessage]]);
  } catch(e) { console.error(e.toString()); }
}

function getYesterdayFormattedDate() {
  var today = new Date();
  today.setDate(today.getDate() - 1); 
  var yyyy = today.getFullYear();
  var mm = String(today.getMonth() + 1).padStart(2, '0');
  var dd = String(today.getDate()).padStart(2, '0');
  return yyyy + '/' + mm + '/' + dd;
}