const getSettingsData = () => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Setting');
  const data = sheet.getRange('B1:B14').getDisplayValues().flat();

  return {
    idfolder:   data[0] || '',
    idapi:     data[1] || '',
    idsheet:   data[2] || '',
    logoSystem: data[3] || '',
    nameEng:    data[4] || '',
    version:   data[5] || ''
  };
};

const doGet = () => {
  const settings = getSettingsData();
  const template = HtmlService.createTemplateFromFile('Index');

  template.apiConfig = JSON.stringify({
    idfolder : settings.idfolder,
    idapi : settings.idapi,
    idsheet : settings.idsheet,
    logoSystem : settings.logoSystem,
    nameEng : settings.nameEng,
    version : settings.version
  });

  const page = template.evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setTitle(settings.nameEng)
    .setFaviconUrl(settings.logoSystem)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  return page;
};