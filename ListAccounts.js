function getGTMData() {
  // Obtener la lista de cuentas
  var accounts = TagManager.Accounts.list({ fields: 'account(name,accountId)' }).account;

  // Crear una hoja de cálculo o obtener la existente
  var sheetName = 'GTM Hierarchy';
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    var headers = ['Account Name', 'Container Name', 'Container ID', 'Latest Version ID', 'Tag #', 'Triggers #', 'Variables #'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    sheet.clear();
  }

  // Iterar sobre cada cuenta
  accounts.forEach(function(account) {
    var containers = getContainers(account.accountId);

    // Iterar sobre cada contenedor
    containers.forEach(function(container) {
      var latestVersion = getLatestVersion(account.accountId, container.containerId);

      // Agregar datos a la hoja de cálculo
      sheet.appendRow([
        account.name,
        container.name,
        container.publicId,
        latestVersion.containerVersionId,
        latestVersion.numTags,
        latestVersion.numTriggers,
        latestVersion.numVariables
      ]);
    });
  });

  Logger.log('Datos actualizados en la hoja de cálculo.');
}

function getContainers(accountId) {
  var containers = TagManager.Accounts.Containers.list(
    'accounts/' + accountId,
    { fields: 'container(name,publicId,containerId)' }
  ).container;

  return containers || [];
}

function getLatestVersion(accountId, containerId) {
  var latestVersion = TagManager.Accounts.Containers.Version_headers.latest(
    'accounts/' + accountId + '/containers/' + containerId
  );

  return latestVersion;
}

function onOpen() {
  var menu = SpreadsheetApp.getUi().createAddonMenu();
  menu.addItem('Actualizar GTM Hierarchy', 'getGTMData');
  menu.addToUi();
}
