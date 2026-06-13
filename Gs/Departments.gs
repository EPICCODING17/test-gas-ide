const generateCodeDept = (sheet) => {
  const lastRow = sheet.getLastRow();
  let existingIDs = [];
  if (lastRow > 1) {
    existingIDs = sheet.getRange(2, 1, lastRow - 1).getValues().flat().map(v => v.toString());
  }

  let num = 1;
  let newID = '';
  while (true) {
    newID = 'DPM-' + String(num).padStart(3, '0');
    if (!existingIDs.includes(newID)) break;
    num++;
  }
  return newID;
};

const addDept = (obj) => {
  const settings = getSettingsData();
  const sheet = SpreadsheetApp.openById(settings.idsheet).getSheetByName("Departments");
  const data  = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0]?.toString().toUpperCase() === obj.deptId.toUpperCase()) {
      return false;
    }
  }

  const codeID = obj.deptId || generateCodeDept(sheet);
  sheet.appendRow([codeID, obj.name]);
  return true;
};

const updateDept = (obj) => {
  const settings = getSettingsData();
  const sheet = SpreadsheetApp.openById(settings.idsheet).getSheetByName("Departments");
  const data  = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0]?.toString() === obj.key) {
      sheet.getRange(i + 1, 2).setValue(obj.name);
      return { success: true };
    }
  }

  return { success: false, message: 'ไม่พบแผนกนี้ในระบบ' };
};

const deleteDept = (deptId) => {
  const settings = getSettingsData();
  const ss = SpreadsheetApp.openById(settings.idsheet);

  const sheetD = ss.getSheetByName("Departments");
  const dataD  = sheetD.getDataRange().getValues();
  for (let i = dataD.length - 1; i >= 1; i--) {
    if (dataD[i][0]?.toString() === deptId) {
      sheetD.deleteRow(i + 1);
      break;
    }
  }

  const sheetU = ss.getSheetByName("Users");
  const dataU  = sheetU.getDataRange().getValues();
  for (let i = 1; i < dataU.length; i++) {
    if (dataU[i][6]?.toString() === deptId) {
      sheetU.getRange(i + 1, 7).setValue('');
    }
  }

  return { success: true };
};