/************************************************
 * WEB APP ENTRYPOINT
 * Renders the scorecard page for a given transporter ID.
 * Deploy this script as a Web App (Anyone with the link)
 * and use: https://script.google.com/macros/s/<DEPLOYMENT_ID>/exec?id=TRANSPORTER_ID
 ************************************************/

function doGet(e) {
  var id =
    e && e.parameter && e.parameter.id
      ? sanitizeTransporterId_(e.parameter.id)
      : "";
  var name =
    e && e.parameter && e.parameter.name ? String(e.parameter.name).trim() : "";

  if (!id) {
    return HtmlService.createHtmlOutput("Missing transporter id.");
  }

  var tpl = HtmlService.createTemplateFromFile("scorecard_page");
  tpl.transporterId = id;
  tpl.driverName = name;

  return tpl
    .evaluate()
    .setTitle("Driver scorecard")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
