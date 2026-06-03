const TEMPLATE_ID = "1qtT0HPTkoVY9lJ17vXXcpy0xTuFerfKewgy8BcnJDSg";
const CERT_FOLDER_ID = "1Gf0RVV2tP2fcSHe-BQoWdNLHq0Ww3CuW";

function processLatestQuiz() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quizSheet = ss.getSheetByName("Form Responses 2");
  const lastRow = quizSheet.getLastRow();
  const rowData = quizSheet.getRange(lastRow,1,1,quizSheet.getLastColumn()).getValues()[0];

  const score = parseInt(rowData[1]);
  const fullName = rowData[7];
  const email = rowData[9];
  const department = rowData[10];

  if (score < 7) return;

  generateCertificate(fullName,email,department);
}

function generateCertificate(name,email,department){
  const folder = DriveApp.getFolderById(CERT_FOLDER_ID);
  const copy = DriveApp.getFileById(TEMPLATE_ID).makeCopy(name + "_Certificate", folder);

  const doc = DocumentApp.openById(copy.getId());
  const body = doc.getBody();

  body.replaceText("{{NAME}}", name);
  body.replaceText("{{DEPARTMENT}}", department);
  body.replaceText("{{DATE}}", Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy"));

  doc.saveAndClose();

  const pdf = copy.getBlob().getAs(MimeType.PDF);

  GmailApp.sendEmail(
    email,
    "Training Completion Certificate",
    "Congratulations! You successfully completed the training. Your certificate is attached.",
    {attachments:[pdf]}
  );
}