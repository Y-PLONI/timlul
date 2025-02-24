function processTranscriptionEmails() {
    var label = GmailApp.getUserLabelByName("תמלול");
    if (!label) {
      Logger.log("אין תווית בשם 'תמלול'");
      return;
    }
    
    var processedLabel = GmailApp.getUserLabelByName("מעובד");
    if (!processedLabel) {
      processedLabel = GmailApp.createLabel("מעובד");
    }
    
    var threads = label.getThreads();
    for (var i = 0; i < threads.length; i++) {
      var thread = threads[i];
      var msgs = thread.getMessages();
      if (msgs.length === 0) continue;
      var msg = msgs[0];
      var attachments = msg.getAttachments();
      var audioFile = null;
      
      for (var j = 0; j < attachments.length; j++) {
        var file = attachments[j];
        if (/\.(mp3|m4a|wav)$/i.test(file.getName())) {
          audioFile = file;
          break;
        }
      }
      
      if (!audioFile) {
        Logger.log("אין קובץ שמע במייל: " + msg.getSubject());
        continue;
      }
      
      // קריאת מפתח Groq מתא G1 בגיליון1
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName("גיליון1");
      var groqApiKey = sheet.getRange("G1").getValue();
      if (!groqApiKey) {
        Logger.log("המפתח של Groq לא נמצא בתא G1 בגיליון");
        continue;
      }
      
      // קריאה לשירות Groq לקבלת תמלול
      var transcription = callGroqTranscriptionAPI(audioFile, groqApiKey);
      
      // קריאת כתובת המייל למשלוח מתא L2 בגיליון1
      var sendToEmail = sheet.getRange("L2").getValue();
      if (!sendToEmail) {
        Logger.log("כתובת המייל לשליחה לא מוגדרת בתא L2");
        continue;
      }
      
      // שליחת מייל עם התמלול
      MailApp.sendEmail({
        to: sendToEmail,
        subject: "תמלול " + msg.getSubject(),
        body: transcription
      });
      
      // סימון התכתובת כמעובדת והסרת תווית "תמלול"
      thread.addLabel(processedLabel);
      thread.removeLabel(label);
    }
  }
  
  function callGroqTranscriptionAPI(audioFile, apiKey) {
    // URL מעודכן לפי ההטמעה של Groq
    var url = "https://api.groq.ai/v1/audio/transcriptions";
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
      throw "שגיאה ב-Groq API, קוד: " + response.getResponseCode();
    }
    
    var result = JSON.parse(response.getContentText());
    if (result.transcription && result.transcription.text) {
      return result.transcription.text;
    }
    return "לא התקבל תמלול";
  }
  