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
      
      // קריאת מפתח Groq מהגיליון (עמודה F = "Groq", עמודה G = המפתח)
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName("גיליון1");
      var data = sheet.getDataRange().getValues();
      var groqApiKey = "";
      for (var k = 1; k < data.length; k++) {
        if (data[k][5] && data[k][5].toString().toLowerCase() === "groq") {
          groqApiKey = data[k][6];
          break;
        }
      }
      if (!groqApiKey) {
        Logger.log("לא נמצא מפתח Groq בגיליון");
        continue;
      }
      
      // קריאה לשירות Groq לקבלת תמלול
      var transcription = callGroqTranscriptionAPI(audioFile, groqApiKey);
      
      // קריאת כתובת המייל למשלוח מתא L2 בגיליון
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
    var url = "https://api.groq.example.com/audio/transcriptions"; // עדכן לכתובת הנכונה
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
    if (result.transcription && result.transcription.text) {
      return result.transcription.text;
    }
    return "לא התקבל תמלול";
  }
  