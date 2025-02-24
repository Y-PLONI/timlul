function processTranscriptionEmails() {
    var labelName = "תמלול";
    var processedLabelName = "מעובד";
    var label = GmailApp.getUserLabelByName(labelName);
    if (!label) {
      Logger.log("אין תווית בשם " + labelName);
      return;
    }
    
    var processedLabel = GmailApp.getUserLabelByName(processedLabelName);
    if (!processedLabel) {
      processedLabel = GmailApp.createLabel(processedLabelName);
    }
    
    var threads = label.getThreads();
    
    for (var i = 0; i < threads.length; i++) {
      var thread = threads[i];
      var threadLabels = thread.getLabels();
      var alreadyProcessed = false;
      for (var j = 0; j < threadLabels.length; j++) {
        if (threadLabels[j].getName() === processedLabelName) {
          alreadyProcessed = true;
          break;
        }
      }
      if (alreadyProcessed) {
        continue;
      }
      
      var messages = thread.getMessages();
      var message = messages[0];
      var attachments = message.getAttachments();
      var audioFile = null;
      
      for (var k = 0; k < attachments.length; k++) {
        var file = attachments[k];
        if (/\.(mp3|m4a|wav)$/i.test(file.getName())) {
          audioFile = file;
          break;
        }
      }
      
      if (!audioFile) {
        Logger.log("אין קובץ שמע במייל: " + message.getSubject());
        continue;
      }
      
      // קריאת מפתח Groq מהגיליון
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName("גיליון1");
      var data = sheet.getDataRange().getValues();
      var groqApiKey = "";
      for (var m = 1; m < data.length; m++) {
        if (data[m][5] && data[m][5].toString().toLowerCase() === "groq") {
          groqApiKey = data[m][6];
          break;
        }
      }
      
      if (!groqApiKey) {
        Logger.log("לא נמצא מפתח Groq בגיליון");
        continue;
      }
      
      // קריאת תמלול מגרוק
      var transcription = callGroqTranscriptionAPI(audioFile, groqApiKey);
      
      // קריאת כתובת המייל למשלוח מתא L2
      var sendToEmail = sheet.getRange("L2").getValue();
      if (!sendToEmail) {
        Logger.log("כתובת המייל לשליחה לא מוגדרת בתא L2");
        continue;
      }
      
      // שליחת מייל עם התמלול
      MailApp.sendEmail({
        to: sendToEmail,
        subject: "תמלול " + message.getSubject(),
        body: transcription
      });
      
      // סימון התכתובת כמעובדת והסרת התווית "תמלול"
      thread.addLabel(processedLabel);
      thread.removeLabel(label);
    }
  }
  
  function callGroqTranscriptionAPI(audioFile, apiKey) {
    var url = "https://api.groq.example.com/audio/transcriptions"; // יש לעדכן לפי התיעוד
    var options = {
      method: "post",
      headers: {
        "Authorization": "Bearer " + apiKey
      },
      payload: {
        file: audioFile.getBytes(),
        model: "whisper-large-v3",
        response_format: "verbose_json"
      },
      muteHttpExceptions: true
    };
    
    var response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) {
      throw "שגיאה בגרוק API, קוד: " + response.getResponseCode();
    }
    
    var result = JSON.parse(response.getContentText());
    return (result.transcription && result.transcription.text) ? result.transcription.text : "לא התקבל תמלול";
  }
  