// משתנה גלובלי לשמירת כתובת המייל ליעד עבור דיווח שגיאות
var globalDestEmail = "";

/**
 * הפונקציה הראשית שמבצעת את כל העיבודים:
 * - בודקת אם ההרצה מותרת (תא A2 בגיליון)
 * - טוענת את מפתחות ה-API, משפטי הסיכום וכתובת המייל ליעד
 * - סורקת את האימיילים המסומנים בתווית "לתמלול"
 * - עבור כל הודעה עם קובץ שמע, מבצעת:
 *    1. שליחת קובץ השמע ל-Groq לקבלת תמלול
 *    2. קריאה ל-Gemini עבור כל משפט סיכום מהגיליון (עמודות H ו-I)
 *    3. הרכבת מייל עם: תאריך+שעה, טלפון מתקשר, טלפון נענה, הטקסט שהתקבל מ-Gemini ולאחריו התמלול המלא
 *    4. שליחת המייל לכתובת שמוגדרת בתא L2, עם המצורפים: ההקלטה, קובץ התמלול וקובץ הסיכום
 *    5. הסרת התווית "לתמלול" מהשיחה והוספת התווית "תמלולV"
 * בכל שלב, במידה ויש שגיאה, היא תדווח ב-Logger וגם תישלח במייל לכתובת היעד.
 */
function runTranscriptionAndEmail() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("גיליון1");
    var status = ss.getRange("A2").getValue();
    if (status !== "Enabled") {
      Logger.log("Script execution is disabled.");
      return;
    }
    
    // טוען את מפתחות ה-API:
    // מפתח Groq בתא G1, מפתח Google (ל-Gemini) בתא G2
    var groqKey = ss.getRange("G1").getValue();
    var googleKey = ss.getRange("G2").getValue();
    
    // טוען את משפטי הסיכום מהגיליון (עמודות H ו-I, החל משורה 2)
    var lastRow = ss.getLastRow();
    var summaryData = ss.getRange("H2:I" + lastRow).getValues();
    var summaryMapping = [];
    for (var i = 0; i < summaryData.length; i++) {
      var num = summaryData[i][0];
      var sentence = summaryData[i][1];
      if (num && sentence) {
        summaryMapping.push({ number: num, sentence: sentence });
      }
    }
    
    // טוען את כתובת המייל ליעד מתא L2
    var destEmail = ss.getRange("L2").getValue();
    globalDestEmail = destEmail; // לשימוש בדיווח שגיאות
    
    // סורק את כל השיחות עם התווית "לתמלול"
    var labelToProcess = GmailApp.getUserLabelByName("לתמלול");
    if (!labelToProcess) {
      Logger.log("אין תווית 'לתמלול' - אין מה לעבד.");
      return;
    }
    var threads = labelToProcess.getThreads();
    
    for (var i = 0; i < threads.length; i++) {
      var thread = threads[i];
      var messages = thread.getMessages();
      for (var j = 0; j < messages.length; j++) {
        try {
          var message = messages[j];
          var attachments = message.getAttachments();
          if (attachments.length === 0) continue;
          
          // בוחר את קובץ השמע מתוך הקבצים המצורפים (לפי MIME או סיומת)
          var audioAttachment = null;
          for (var k = 0; k < attachments.length; k++) {
            var att = attachments[k];
            var mime = att.getContentType();
            if (mime.indexOf("audio") !== -1 || att.getName().match(/\.(m4a|mp3|wav)$/i)) {
              audioAttachment = att;
              break;
            }
          }
          if (!audioAttachment) continue;
          
          // קריאה ל-Groq לקבלת תמלול
          var transcriptionText = callGroqTranscription(audioAttachment, groqKey);
          if (!transcriptionText) {
            Logger.log("תמלול נכשל עבור הודעה: " + message.getId());
            reportError("תמלול נכשל עבור הודעה: " + message.getId());
            continue;
          }
          
          // עבור כל משפט סיכום מהגיליון, שולח את התמלול ל-Gemini ומאגד את התוצאות
          var aggregatedSummary = "";
          for (var s = 0; s < summaryMapping.length; s++) {
            var prompt = summaryMapping[s].sentence + "\n\n" + transcriptionText;
            var summaryResult = callGeminiSummary(prompt, googleKey);
            if (summaryResult) {
              aggregatedSummary += summaryResult + "\n\n";
            }
          }
          
          // חילוץ פרטים מהאימייל המקורי: תאריך, שעה, טלפון מתקשר וטלפון נענה
          var subject = message.getSubject();
          var body = message.getPlainBody();
          var dateMatch = subject.match(/(\d{4}-\d{2}-\d{2})/);
          var dateStr = dateMatch ? dateMatch[1] : "";
          var timeMatch = body.match(/בשעה[:\s]*([\d]{2}:[\d]{2}:[\d]{2})/);
          var timeStr = timeMatch ? timeMatch[1] : "";
          var callerMatch = subject.match(/טלפון\s*(\d+)/);
          var callerPhone = callerMatch ? callerMatch[1] : "";
          var calledMatch = subject.match(/למספר\s*(\d+)/);
          var calledPhone = calledMatch ? calledMatch[1] : "";
          
          // הרכבת נושא וגוף המייל
          var finalSubject = subject;
          var finalBody = "תאריך: " + dateStr + " שעה: " + timeStr + "\n" +
                           "טלפון מתקשר: " + callerPhone + "\n" +
                           "טלפון נענה: " + calledPhone + "\n\n" +
                           aggregatedSummary.trim() + "\n\n" +
                           "תמלול:\n" + transcriptionText;
          
          // הכנת קבצי מצורף עבור התמלול והסיכום
          var transcriptionBlob = Utilities.newBlob(transcriptionText, "text/plain", audioAttachment.getName() + ".txt");
          var summaryBlob = Utilities.newBlob(aggregatedSummary.trim(), "text/plain", "Summary-" + audioAttachment.getName() + ".txt");
          
          // שליחת המייל לכתובת היעד עם המצורפים: ההקלטה, התמלול והסיכום
          GmailApp.sendEmail(destEmail, finalSubject, finalBody, {
            attachments: [audioAttachment, transcriptionBlob, summaryBlob]
          });
          
          // סימון ההודעה כמעובדת: הסרת התווית "לתמלול" והוספת "תמלולV"
          var processedLabel = GmailApp.getUserLabelByName("תמלולV");
          if (!processedLabel) {
            processedLabel = GmailApp.createLabel("תמלולV");
          }
          thread.removeLabel(labelToProcess);
          thread.addLabel(processedLabel);
          
          Logger.log("הודעה עובדה: " + message.getId());
        } catch (innerError) {
          Logger.log("שגיאה בעיבוד הודעה: " + innerError);
          reportError("שגיאה בעיבוד הודעה: " + innerError + " בהודעה: " + message.getId());
        }
      }
    }
  } catch (e) {
    Logger.log("שגיאה כללית בסקריפט: " + e);
    reportError("שגיאה כללית בסקריפט: " + e);
  }
}

/**
 * פונקציה לדיווח שגיאות: מדווחת דרך Logger וגם שולחת מייל לכתובת היעד אם מוגדרת.
 */
function reportError(errorMsg) {
  Logger.log("דיווח שגיאה: " + errorMsg);
  if (globalDestEmail) {
    try {
      GmailApp.sendEmail(globalDestEmail, "שגיאת סקריפט", errorMsg);
    } catch (emailError) {
      Logger.log("לא ניתן לשלוח מייל דיווח שגיאה: " + emailError);
    }
  }
}

/**
 * שולחת את קובץ השמע ל-Groq ומחזירה את הטקסט של התמלול.
 */
function callGroqTranscription(audioBlob, groqKey) {
  try {
    var url = "https://api.groq.com/openai/v1/audio/transcriptions";
    var fields = {
      "model": "whisper-large-v3",
      "response_format": "verbose_json"
    };
    var multipartData = buildMultipartPayload(fields, "file", audioBlob);
    var options = {
      method: "post",
      contentType: "multipart/form-data; boundary=" + multipartData.boundary,
      payload: multipartData.blob.getBytes(),
      headers: {
        "Authorization": "Bearer " + groqKey
      },
      muteHttpExceptions: true
    };
    
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    if (responseCode >= 200 && responseCode < 300) {
      var json = JSON.parse(response.getContentText());
      return json.text;
    } else {
      Logger.log("שגיאה ב-Groq API: " + response.getContentText());
      reportError("שגיאה ב-Groq API: " + response.getContentText());
      return null;
    }
  } catch (e) {
    Logger.log("Exception ב-callGroqTranscription: " + e);
    reportError("Exception ב-callGroqTranscription: " + e);
    return null;
  }
}

/**
 * בונה את גוף הבקשה מסוג multipart/form-data.
 */
function buildMultipartPayload(fields, fileFieldName, fileBlob) {
  try {
    var boundary = "------WebKitFormBoundary" + new Date().getTime();
    var delimiter = "--" + boundary;
    var closeDelimiter = delimiter + "--";
    var payload = "";
    for (var key in fields) {
      payload += delimiter + "\r\n";
      payload += 'Content-Disposition: form-data; name="' + key + '"\r\n\r\n';
      payload += fields[key] + "\r\n";
    }
    payload += delimiter + "\r\n";
    payload += 'Content-Disposition: form-data; name="' + fileFieldName + '"; filename="' + fileBlob.getName() + '"\r\n';
    payload += 'Content-Type: ' + fileBlob.getContentType() + '\r\n\r\n';
    
    var payloadBytes = Utilities.newBlob(payload).getBytes();
    var fileBytes = fileBlob.getBytes();
    var closingBytes = Utilities.newBlob("\r\n" + closeDelimiter).getBytes();
    
    var combined = payloadBytes.concat(fileBytes).concat(closingBytes);
    var multipartBlob = Utilities.newBlob(combined, 'multipart/form-data; boundary=' + boundary);
    return { blob: multipartBlob, boundary: boundary };
  } catch (e) {
    Logger.log("Exception ב-buildMultipartPayload: " + e);
    reportError("Exception ב-buildMultipartPayload: " + e);
    throw e;
  }
}

/**
 * שולחת את הטקסט המבוקש ל-Gemini ומחזירה את תוצאת הסיכום.
 */
function callGeminiSummary(promptText, googleKey) {
  try {
    var url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash-exp:generateContent?key=" + googleKey;
    var payload = {
      "contents": [
        {
          "role": "user",
          "parts": [
            { "text": promptText }
          ]
        }
      ],
      "generationConfig": {
        "temperature": 0.7,
        "topP": 0.8,
        "topK": 40,
        "maxOutputTokens": 2048
      }
    };
    
    var options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    if (responseCode >= 200 && responseCode < 300) {
      var json = JSON.parse(response.getContentText());
      if (json.candidates && json.candidates.length > 0 && json.candidates[0].content && json.candidates[0].content.parts) {
        var parts = json.candidates[0].content.parts;
        var summaryText = "";
        for (var i = 0; i < parts.length; i++) {
          if (parts[i].text) {
            summaryText += parts[i].text;
          }
        }
        return summaryText.trim();
      } else {
        Logger.log("תשובה לא תקינה מ-Gemini API: " + response.getContentText());
        reportError("תשובה לא תקינה מ-Gemini API: " + response.getContentText());
        return "";
      }
    } else {
      Logger.log("שגיאה ב-Gemini API: " + response.getContentText());
      reportError("שגיאה ב-Gemini API: " + response.getContentText());
      return "";
    }
  } catch (e) {
    Logger.log("Exception ב-callGeminiSummary: " + e);
    reportError("Exception ב-callGeminiSummary: " + e);
    return "";
  }
}
