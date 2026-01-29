function simplifyText() {
  const sheet = SpreadsheetApp.openById("GOOGLE_SHEET_URL").getSheets()[0];
  const lastRow = sheet.getLastRow();
const alreadyProcessed = sheet.getRange(lastRow, 2).getValue();
if (alreadyProcessed) return;


  const originalText = sheet.getRange(lastRow, 1).getValue();
  if (!originalText) return;

  const apiKey = "YOUR_GROQ_API_KEY";

  const prompt = `
Rewrite the following text in simple English.

Rules:
- Short sentences
- Easy words
- Explain like to a non-native English speaker

Output EXACTLY in this format:

Simple explanation:
<text>

Key points:
- point 1
- point 2
- point 3
- point 4
- point 5

Keywords:
word1, word2, word3, word4, word5

Text:
${originalText}
`;

  const response = UrlFetchApp.fetch(
    "https://api.groq.com/openai/v1/chat/completions",
    {
      method: "post",
      contentType: "application/json",
      headers: {
        Authorization: "Bearer " + apiKey
      },
      payload: JSON.stringify({
        model: "llama-3.1-8b-instant",
        messages: [{ role: "user", content: prompt }],
        temperature: 0.3
      }),
      muteHttpExceptions: true
    }
  );

  const result = JSON.parse(response.getContentText());

  if (result.error) {
    throw new Error(result.error.message);
  }

  const output = result.choices[0].message.content;

// Split sections
const simple = output.match(/Simple explanation:\s*([\s\S]*?)\n\n/i);
const points = output.match(/Key points:\s*([\s\S]*?)\n\n/i);
const keywords = output.match(/Keywords:\s*([\s\S]*)/i);

// Write to columns
if (simple) sheet.getRange(lastRow, 2).setValue(simple[1].trim());
if (points) sheet.getRange(lastRow, 3).setValue(points[1].trim());
if (keywords) sheet.getRange(lastRow, 4).setValue(keywords[1].trim());

sheet.getRange(lastRow, 5).setValue(new Date());

}

function processLatestRow() {
  simplifyText();
}

function createOnEditTrigger() {
  const sheetId = "GOOGLE_SHEET_URL";

  // Delete existing triggers (important)
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));

  // Create installable on-edit trigger
  ScriptApp.newTrigger("processLatestRow")
    .forSpreadsheet(sheetId)
    .onEdit()
    .create();
}
function deleteAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => ScriptApp.deleteTrigger(t));
}



