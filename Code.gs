function doGet(e) { return handleRequest(e); }

function doPost(e) {
  var p = JSON.parse(e.postData.contents);
  return handleAction(p);
}

function handleRequest(e) {
  var p = e.parameter;
  return handleAction(p);
}

function handleAction(p) {
  var out = {};

  if (p.action === "headers") {

    out = { headers: ["ID","Date","Type","Audience","Title",
      "Content","Status","Tags","AddedAt","Notes"] };

  } else if (p.action === "read") {

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Log");
    var rows = sheet.getDataRange().getValues();
    out = { rows: rows };

  } else if (p.action === "append") {

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Log");
    var row = JSON.parse(p.row);
    sheet.appendRow(row);
    out = { success: true };

  } else if (p.action === "update") {

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Log");
    sheet.getRange(parseInt(p.row), parseInt(p.col)).setValue(p.value);
    out = { success: true };

  } else if (p.action === "clear") {

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Log");
    sheet.getRange(parseInt(p.row), 1, 1, 10).clearContent();
    out = { success: true };

  } else if (p.action === "ai_suggest") {

    var apiKey = p.apiKey;
    var prompt = p.prompt;

    if (!apiKey || !prompt) {
      out = { error: "Missing apiKey or prompt" };
    } else {
      var payload = JSON.stringify({
        model: "claude-sonnet-5",
        max_tokens: 8000,
        messages: [{ role: "user", content: prompt }]
      });
      var options = {
        method: "post",
        contentType: "application/json",
        headers: { "x-api-key": apiKey,
          "anthropic-version": "2023-06-01" },
        payload: payload,
        muteHttpExceptions: true
      };
      try {
        var resp = UrlFetchApp.fetch(
          "https://api.anthropic.com/v1/messages", options);
        var result = JSON.parse(resp.getContentText());
        if (result.content && result.content[0]) {
          out = { suggestion: result.content[0].text };
        } else if (result.error) {
          out = { error: result.error.message || "API error" };
        } else {
          out = { error: "No response from AI" };
        }
      } catch(err) {
        out = { error: err.message };
      }
    }

  } else {
    out = { error: "unknown action" };
  }

  return ContentService
    .createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}
