function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle(APP_CONFIG.APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSystemStatus() {
  const user = assertAuthorized_();
  const checks = {};
  Object.keys(APP_CONFIG.SHEETS).forEach(function(key) {
    const sheet = getSheet_(key, false);
    checks[key] = sheet ? {
      ok: true,
      name: sheet.getName(),
      rows: sheet.getLastRow()
    } : {
      ok: false,
      expected: APP_CONFIG.SHEETS[key].name
    };
  });
  return serializeValue_({
    ok: Object.keys(checks).every(function(key) { return checks[key].ok; }),
    user: user,
    version: APP_CONFIG.APP_VERSION,
    checks: checks
  });
}
