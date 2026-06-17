/**
 * 1. ROUTING: Decides which page to load and sets Browser Tab Branding.
 */
function doGet(e) {
  try {
    var iconUrl = 'https://cdn.jsdelivr.net/gh/twitter/twemoji@14.0.2/assets/72x72/1f393.png'; 
    
    if (e && e.parameter && e.parameter.page === 'admin') {
      return HtmlService.createHtmlOutputFromFile('Admin')
          .setTitle('Student Learning Hub Admin')
          .setFaviconUrl(iconUrl)
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
          .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    } else {
      return HtmlService.createHtmlOutputFromFile('Index')
          .setTitle('Student Learning Hub')
          .setFaviconUrl(iconUrl)
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
          .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }
  } catch (error) {
    return HtmlService.createHtmlOutput("<h2>Deployment Error</h2><p>" + error.message + "</p>");
  }
}

function getAdminHtml() {
  try { return HtmlService.createHtmlOutputFromFile('Admin').getContent(); } 
  catch (error) { throw new Error("Could not load Admin.html. Make sure the file exists."); }
}

function getLogsSheet(ss) {
  return ss.getSheetByName("Logs") || ss.getSheetByName("data") || ss.getSheetByName("Data");
}

// ==========================================
// STUDENT HUB FUNCTIONS
// ==========================================

function getInitialData() {
  try { 
    var email = (Session.getActiveUser().getEmail() || "anonymous").toLowerCase().trim();
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("modules") || ss.getSheetByName("Modules");
    var data = sheet ? sheet.getDataRange().getValues() : [];
    
    var setQuizzes = {};
    for (var i = 1; i < data.length; i++) {
       if (data[i][1] && data[i][1].toString().trim() === '__ACTIVE_SET_QUIZ__') {
           try {
              var sq = JSON.parse(data[i][2]);
              setQuizzes[data[i][0].toString().trim()] = { name: sq.quizName, qCount: (sq.questions || []).length };
           } catch(e) {}
       }
    }

    return { 
      hubData: getHubData(), 
      userProgress: getUserProgress(email),
      userPreferences: getUserPreferences(email),
      setQuizzes: setQuizzes
    }; 
  } 
  catch (e) { throw new Error("getInitialData Error: " + e.message); }
}

function getHubData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("modules") || ss.getSheetByName("Modules");
  if (!sheet) throw new Error("Could not find a sheet named 'modules'.");
  var data = sheet.getDataRange().getValues();
  var hubList = [];
  for (var i = 1; i < data.length; i++) {
    var modName = data[i][1] ? data[i][1].toString().trim() : "";
    if(data[i][0] && modName && modName !== '__ACTIVE_SET_QUIZ__') {
      var modTheme = "#3498db"; 
      try {
        var jData = JSON.parse(data[i][2]);
        if (jData.theme) modTheme = jData.theme;
      } catch(e) {}
      
      hubList.push({ 
        course: data[i][0].toString().trim(), 
        module: modName,
        theme: modTheme
      });
    }
  }
  return hubList;
}

function getUserProgress(email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getLogsSheet(ss);
  var progress = { lastAssessment: {}, moduleScores: {}, setQuizScores: {}, misconceptions: {}, moduleStats: {} };
  if (!sheet) return progress; 

  var data = sheet.getDataRange().getValues();
  var activeMisconceptions = {}; 

  var formatDate = function(dateObj) {
    if (!dateObj || isNaN(new Date(dateObj))) return "";
    return new Date(dateObj).toLocaleDateString();
  };

  for (var i = 1; i < data.length; i++) {
    var rowEmail = data[i][1] ? data[i][1].toString().toLowerCase().trim() : "";
    if (rowEmail !== email) continue; 
    
    var timestamp = data[i][0];
    var rawDateMs = new Date(timestamp).getTime();
    var course = data[i][2] ? data[i][2].toString().trim() : "";
    var module = data[i][3] ? data[i][3].toString().trim() : "";
    var phase = data[i][4] ? data[i][4].toString().trim() : "";
    var score = parseInt(data[i][5], 10) || 0;
    
    var miscStr = data[i][6] ? data[i][6].toString().trim() : "";
    var miscItems = [];
    if (miscStr && !miscStr.match(/^none$/i) && !miscStr.match(/^n\/a$/i) && miscStr !== "-") {
       miscItems = miscStr.split("|").map(function(s){ return s.trim(); }).filter(Boolean);
    }

    if (!progress.moduleScores[course]) progress.moduleScores[course] = {};
    if (!progress.setQuizScores[course]) progress.setQuizScores[course] = {};
    if (!activeMisconceptions[course]) activeMisconceptions[course] = {};
    if (!progress.moduleStats[course]) progress.moduleStats[course] = {};
    
    var isSetQuiz = module.indexOf("Set Quiz:") === 0;

    // FIX: EXCLUDES RANDOM QUIZ FROM PERSISTENT TRACKING
    if (module !== "Misconception Buster" && module !== "Custom Revision" && module !== "Random Quiz") {
      if (!activeMisconceptions[course][module]) activeMisconceptions[course][module] = {};
      
      if (!isSetQuiz) {
         if (!progress.moduleStats[course][module]) {
            progress.moduleStats[course][module] = { readingTime: 0, compCount: 0, hangmanCount: 0, testCount: 0, testScoreTotal: 0 };
         }
         var pLower = phase.toLowerCase();
         var ms = progress.moduleStats[course][module];

         if (pLower === "reading") { ms.readingTime += score; }
         else if (pLower.indexOf("comp:") === 0) { ms.compCount++; }
         else if (pLower === "keyword") { ms.hangmanCount++; }
         else if (pLower === "final quiz") { ms.testCount++; ms.testScoreTotal += score; }
      }

      if (phase === "Final Quiz") {
        if (isSetQuiz) {
           var sqName = module.replace("Set Quiz:", "").trim();
           var existing = progress.setQuizScores[course][sqName];
           if (!existing || score > parseInt(existing.score)) {
               progress.setQuizScores[course][sqName] = { score: score + "%", date: formatDate(timestamp) };
           }
        } else {
           progress.moduleScores[course][module] = { score: score + "%", date: formatDate(timestamp), rawDate: rawDateMs };
           progress.lastAssessment[course] = { module: module, score: score + "%", date: formatDate(timestamp) };
        }
        activeMisconceptions[course][module] = {};
        miscItems.forEach(function(item) { activeMisconceptions[course][module][item.toLowerCase()] = item; });
      } 
      else if (phase.indexOf("Comp:") === 0) {
        var chunkName = phase.replace("Comp:", "").trim();
        if (score === 100) { delete activeMisconceptions[course][module][chunkName.toLowerCase()]; } 
        else {
           activeMisconceptions[course][module][chunkName.toLowerCase()] = chunkName;
           miscItems.forEach(function(item) { activeMisconceptions[course][module][item.toLowerCase()] = item; });
        }
      } 
      else if (phase === "Keyword" && score === 0) {
         miscItems.forEach(function(item) { activeMisconceptions[course][module][item.toLowerCase()] = item; });
      }
    }

    if (phase === "Buster Clear" && miscItems.length > 0) {
      miscItems.forEach(function(clearedItem) {
        var ct = clearedItem.toLowerCase();
        for (var mod in activeMisconceptions[course]) {
           if (activeMisconceptions[course][mod][ct]) {
              delete activeMisconceptions[course][mod][ct];
           }
        }
      });
    }
  }

  for (var c in activeMisconceptions) {
     progress.misconceptions[c] = {};
     for (var m in activeMisconceptions[c]) {
        var items = Object.values(activeMisconceptions[c][m]);
        if (items.length > 0) progress.misconceptions[c][m] = items;
     }
  }

  return progress;
}

function getUserPreferences(email) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Users");
    if (!sheet) return []; 
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().toLowerCase().trim() === email) {
        return data[i][1] ? data[i][1].toString().split('|') : [];
      }
    }
    return [];
  } catch(e) { return []; }
}

function saveUserPreferences(selectedCoursesArray) {
  try {
    var email = (Session.getActiveUser().getEmail() || "anonymous").toLowerCase().trim();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Users");
    if (!sheet) {
      sheet = ss.insertSheet("Users");
      sheet.appendRow(["Email", "Active Courses"]);
      sheet.setFrozenRows(1);
    }
    var data = sheet.getDataRange().getValues();
    var found = false;
    var courseStr = selectedCoursesArray.join('|');
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().toLowerCase().trim() === email) {
        sheet.getRange(i + 1, 2).setValue(courseStr);
        found = true; break;
      }
    }
    if (!found) sheet.appendRow([email, courseStr]);
    return true;
  } catch(e) { throw new Error("Could not save preferences: " + e.message); }
}

function getModuleData(course, moduleName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("modules") || ss.getSheetByName("Modules");
    if (!sheet) throw new Error("Could not find 'modules' sheet.");
    
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === course && data[i][1].toString().trim() === moduleName) {
        return JSON.parse(data[i][2]);
      }
    }
    throw new Error("Module '" + moduleName + "' not found in database.");
  } catch (e) { throw new Error("getModuleData Error: " + e.message); }
}

function getSetQuizData(courseName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("modules") || ss.getSheetByName("Modules");
    if (!sheet) throw new Error("Could not find 'modules' sheet.");
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === courseName && data[i][1].toString().trim() === '__ACTIVE_SET_QUIZ__') {
        return JSON.parse(data[i][2]);
      }
    }
    throw new Error("No active set quiz found.");
  } catch (e) { throw new Error("getSetQuizData Error: " + e.message); }
}

function getCustomRevisionData(courseName, selectedModules) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("modules") || ss.getSheetByName("Modules");
    if (!sheet) throw new Error("Could not find 'modules' sheet.");

    var data = sheet.getDataRange().getValues();
    var aggData = { readings: [], keywords: [], questions: [] };

    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === courseName) {
        var modName = data[i][1].toString().trim();
        if (selectedModules.indexOf(modName) > -1 && modName !== '__ACTIVE_SET_QUIZ__') {
          try {
            var modJson = JSON.parse(data[i][2]);
            if (modJson.p1) {
               var bestText = modJson.p1.deep || modJson.p1.standard || modJson.p1.core || null;
               if (bestText) aggData.readings.push({ module: modName, text: bestText });
            }
            if (modJson.p2 && modJson.p2.length > 0) {
               aggData.keywords = aggData.keywords.concat(modJson.p2);
            }
            if (modJson.p4 && modJson.p4.length > 0) {
               var modQs = modJson.p4.map(function(q) {
                  q.sourceModule = modName;
                  return q;
               });
               aggData.questions = aggData.questions.concat(modQs);
            }
          } catch(e) {}
        }
      }
    }
    return aggData;
  } catch (e) { throw new Error("getCustomRevisionData Error: " + e.message); }
}

function logInteraction(course, module, phase, score, misconceptions) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getLogsSheet(ss);
    if (!sheet) {
      sheet = ss.insertSheet("Logs");
      sheet.appendRow(["Timestamp", "Email", "Course", "Module", "Phase", "Score", "Misconceptions"]);
      sheet.setFrozenRows(1);
    }
    var email = Session.getActiveUser().getEmail() || "anonymous";
    var timestamp = new Date();
    var miscFormatted = "None";
    if (Array.isArray(misconceptions) && misconceptions.length > 0) {
      miscFormatted = misconceptions.join(" | ");
    } else if (typeof misconceptions === 'string' && misconceptions.trim() !== "") {
      miscFormatted = misconceptions.trim();
    }
    sheet.appendRow([timestamp, email, course, module, phase, score, miscFormatted]);
    return true;
  } catch (e) { throw new Error("logInteraction Error: " + e.message); }
}

function getUserEmail() { return Session.getActiveUser().getEmail() || "Guest Mode"; }

// ==========================================
// TEACHER ADMIN FUNCTIONS
// ==========================================

function getAdminLogs() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = getLogsSheet(ss);
    if (!logSheet) throw new Error("Could not find a student tracking sheet.");
    var data = logSheet.getDataRange().getDisplayValues();
    if (data.length < 2) return [];
    var headers = data[0].map(function(h) { return h.toString().toLowerCase().trim(); });
    var timeCol = headers.indexOf("timestamp");
    var emailCol = headers.findIndex(function(h) { return h.includes("email") || h.includes("student"); });
    var courseCol = headers.indexOf("course");
    var modCol = headers.indexOf("module");
    var phaseCol = headers.indexOf("phase");
    var scoreCol = headers.indexOf("score");
    var misCol = headers.findIndex(function(h) { return h.includes("misconception"); });

    var logs = [];
    for (var i = 1; i < data.length; i++) {
      if (emailCol > -1 && data[i][emailCol]) {
        logs.push({
          timestamp: timeCol > -1 ? data[i][timeCol] : "",
          email: emailCol > -1 ? data[i][emailCol] : "Unknown",
          course: courseCol > -1 && data[i][courseCol] ? data[i][courseCol] : "General",
          module: modCol > -1 ? data[i][modCol] : "General",
          phase: phaseCol > -1 ? data[i][phaseCol] : "Task",
          score: scoreCol > -1 ? data[i][scoreCol] : "0",
          misconceptions: misCol > -1 ? data[i][misCol] : ""
        });
      }
    }
    return logs;
  } catch (error) { throw new Error("getAdminLogs Error: " + error.message); }
}

function generateClassSpreadsheet(courseName, headers, rows) {
  try {
    var ssNew = SpreadsheetApp.create(courseName + " - Performance Overview Matrix");
    var sheet = ssNew.getActiveSheet();
    sheet.setName("Overview Matrix");
    
    sheet.appendRow(headers);
    if (rows && rows.length > 0) {
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }
    
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#4f46e5");
    headerRange.setFontColor("#ffffff");
    headerRange.setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
    
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(2);
    
    for (var i = 1; i <= headers.length; i++) {
      sheet.autoResizeColumn(i);
    }
    
    return ssNew.getUrl();
  } catch (e) {
    throw new Error("Spreadsheet generation failed: " + e.message);
  }
}

function getAdminModuleList(courseName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("modules") || ss.getSheetByName("Modules");
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    var mods = [];
    for (var r = 1; r < data.length; r++) { if (data[r][0].toString().trim() === courseName) { mods.push({ module: data[r][1].toString().trim(), json: data[r][2] }); } }
    return mods;
  } catch (e) { throw new Error("Error fetching modules: " + e.message); }
}

function saveModuleData(courseName, moduleName, jsonData) {
  try {
    JSON.parse(jsonData);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("modules") || ss.getSheetByName("Modules");
    if (!sheet) { sheet = ss.insertSheet("modules"); sheet.appendRow(["Course", "Module", "JSON Data"]); }
    var data = sheet.getDataRange().getValues();
    var found = false;
    for (var r = 1; r < data.length; r++) {
      if (data[r][0].toString().trim() === courseName && data[r][1].toString().trim() === moduleName) {
        sheet.getRange(r + 1, 3).setValue(jsonData);
        found = true; break;
      }
    }
    if (!found) sheet.appendRow([courseName, moduleName, jsonData]);
    return true;
  } catch (e) { throw new Error("Error saving module: " + e.message); }
}

function saveSetQuiz(courseName, quizName, questionsData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("modules") || ss.getSheetByName("Modules");
    if (!sheet) { sheet = ss.insertSheet("modules"); sheet.appendRow(["Course", "Module", "JSON Data"]); }
    
    var jsonStr = JSON.stringify({ quizName: quizName, questions: questionsData });
    var data = sheet.getDataRange().getValues();
    var found = false;
    for (var r = 1; r < data.length; r++) {
      if (data[r][0].toString().trim() === courseName && data[r][1].toString().trim() === '__ACTIVE_SET_QUIZ__') {
        sheet.getRange(r + 1, 3).setValue(jsonStr);
        found = true; break;
      }
    }
    if (!found) sheet.appendRow([courseName, '__ACTIVE_SET_QUIZ__', jsonStr]);
    return true;
  } catch(e) { throw new Error("Error saving Set Quiz: " + e.message); }
}
