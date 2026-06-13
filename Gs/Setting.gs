const getSet = () => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Setting');
  const data = sheet.getRange('B1:B6').getDisplayValues().flat();
  return data;
}

const settingGS = (data) => {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Setting');
    const valuesToSet = [];
    for (let i = 1; i <= 6; i++) {
      valuesToSet.push([data[`set${i}`] || '']);
    }
    sheet.getRange(1, 2, valuesToSet.length, 1).setValues(valuesToSet);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
};