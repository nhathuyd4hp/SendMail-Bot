// -- GET --
function doGet(e) {
  const gid = e.parameter.gid;
  if (!gid) {
    return ContentService.createTextOutput(
      JSON.stringify({ error: "Missing or invalid gid parameter" })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetById(gid);
  if (!sheet) {
    return ContentService.createTextOutput(
      JSON.stringify({ error: "Sheet not found" })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  const data = sheet.getDataRange().getValues();
  const backgrounds = sheet.getDataRange().getBackgrounds();

  const response = {
    sheet_name: sheet.getName(),
    data: data,
    backgrounds: backgrounds,
  };

  return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(
    ContentService.MimeType.JSON
  );
}

// -- POST --
function doPost(e) {
  try {
    // -- Check content
    const contents =
      e.postData && e.postData.contents ? e.postData.contents : "";
    if (!contents) {
      return ContentService.createTextOutput(
        JSON.stringify({ error: "Body is empty" })
      ).setMimeType(ContentService.MimeType.JSON);
    }
    const jsonContents = JSON.parse(contents);
    // -- Check method
    const method = jsonContents.method;
    if (!method) {
      return ContentService.createTextOutput(
        JSON.stringify({ error: "method is missing" })
      ).setMimeType(ContentService.MimeType.JSON);
    }
    // -- Check Sheet
    const sheet_id = jsonContents.sheet_id;
    if (!sheet_id) {
      return ContentService.createTextOutput(
        JSON.stringify({ error: "sheet_id is missing" })
      ).setMimeType(ContentService.MimeType.JSON);
    }
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetById(sheet_id);
    if (!sheet) {
      return ContentService.createTextOutput(
        JSON.stringify({ error: "Sheet not found" })
      ).setMimeType(ContentService.MimeType.JSON);
    }
    // -- Action
    if (method === "POST") {
      const data = jsonContents.data;
      if (!data || !Array.isArray(data)) {
        return ContentService.createTextOutput(
          JSON.stringify({ error: "data is missing or invalid" })
        ).setMimeType(ContentService.MimeType.JSON);
      }
      sheet.appendRow(data);
      return ContentService.createTextOutput(
        JSON.stringify({ success: "ok" })
      ).setMimeType(ContentService.MimeType.JSON);
    } else if (method === "DELETE") {
      const row = jsonContents.row;
      if (!row) {
        return ContentService.createTextOutput(
          JSON.stringify({ error: "row is missing" })
        ).setMimeType(ContentService.MimeType.JSON);
      }
      sheet.deleteRow(row);
      return ContentService.createTextOutput(
        JSON.stringify({ success: "ok" })
      ).setMimeType(ContentService.MimeType.JSON);
    } else if (method === "PUT") {
      const row = jsonContents.row;
      const data = jsonContents.data;
      const color = jsonContents.color;

      if (!row || row < 1) {
        return ContentService.createTextOutput(
          JSON.stringify({ error: "row is missing or invalid" })
        ).setMimeType(ContentService.MimeType.JSON);
      }

      const lastColumn = sheet.getLastColumn(); // Lấy số cột cuối cùng hiện có

      if (data && Array.isArray(data)) {
        sheet.getRange(row, 1, 1, lastColumn).clearContent();
        sheet.getRange(row, 1, 1, data.length).setValues([data]);
      }

      if (color) {
        sheet.getRange(row, 1, 1, lastColumn).setBackground(color);
      }

      return ContentService.createTextOutput(
        JSON.stringify({ success: "ok" })
      ).setMimeType(ContentService.MimeType.JSON);
    } else {
      return ContentService.createTextOutput(
        JSON.stringify({ error: "method must be in ['POST','PUT','DELETE']" })
      ).setMimeType(ContentService.MimeType.JSON);
    }
  } catch (error) {
    return ContentService.createTextOutput(
      JSON.stringify({ error: error.message })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}
