# Ben
const EMAIL_ADMIN = "coulibalybendjafar@gmail.com";

const SHEET_ID = "1o79b1N3QSWcx0YsjPd1VoL12Z4pHxs2q-PYA731Acxs";

function doPost(e) {

  const sheet = SpreadsheetApp
    .openById(SHEET_ID)
    .getSheets()[0];

  const p = e.parameter || {};

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "Date",
      "Nom et prénom",
      "Téléphone",
      "WhatsApp",
      "Genre",
      "Fier de ZAP à l'ATF ?",
      "Moyenne",
      "Série"
    ]);
  }

  sheet.appendRow([
    new Date(),
    p.nom || "",
    p.telephone || "",
    p.whatsapp || "",
    p.genre || "",
    p.atf || "",
    p.moyenne || "",
    p.serie || ""
  ]);

  MailApp.sendEmail({
    to: EMAIL_ADMIN,
    subject: "Nouvelle réponse - Ben Adaptivité",
    body:
      "Une nouvelle réponse a été reçue.\n\n" +
      "Nom et prénom : " + (p.nom || "") + "\n" +
      "Téléphone : " + (p.telephone || "") + "\n" +
      "WhatsApp : " + (p.whatsapp || "") + "\n" +
      "Genre : " + (p.genre || "") + "\n" +
      "Fier de ZAP à l'ATF ? : " + (p.atf || "") + "\n" +
      "Moyenne : " + (p.moyenne || "") + "\n" +
      "Série : " + (p.serie || "")
  });

  return ContentService
    .createTextOutput("OK")
    .setMimeType(ContentService.MimeType.TEXT);
}