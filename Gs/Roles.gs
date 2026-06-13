const generateCodeRole = (sheet) => {
  const lastRow = sheet.getLastRow();
  let existingIDs = [];
  if (lastRow > 1) {
    existingIDs = sheet.getRange(2, 1, lastRow - 1).getValues().flat().map(v => v.toString());
  }

  let num = 1;
  let newID = '';
  while (true) {
    newID = 'R' + String(num).padStart(2, '0');
    if (!existingIDs.includes(newID)) break;
    num++;
  }
  return newID;
};

const addRole = (obj) => {
  const settings = getSettingsData();
  const ss    = SpreadsheetApp.openById(settings.idsheet);
  const sheet = ss.getSheetByName("Roles");
  const data  = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0]?.toString().toUpperCase() === obj.roleId.toUpperCase()) {
      return false;
    }
  }

  const codeID = obj.roleId || generateCodeRole(sheet);
  sheet.appendRow([codeID, obj.name, obj.desc]);

  const sheetRA = ss.getSheetByName("Role_Actions");
  const defaultActions = [
    'approve',
    'edit',
    'delete',
    'copy',
    'view',
    'in',
    'out',
    'return',
    'transfer'
  ];
  defaultActions.forEach(action => {
    sheetRA.appendRow([codeID, action, 'true']);
  });

  return true;
};

const updateRole = (obj) => {
  const settings = getSettingsData();
  const sheet = SpreadsheetApp.openById(settings.idsheet).getSheetByName("Roles");
  const data  = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0]?.toString() === obj.key) {
      sheet.getRange(i + 1, 2).setValue(obj.name);
      sheet.getRange(i + 1, 3).setValue(obj.desc);
      return { success: true };
    }
  }

  return { success: false, message: 'ไม่พบ Role นี้ในระบบ' };
};

const deleteRole = (roleId) => {
  const settings = getSettingsData();
  const ss = SpreadsheetApp.openById(settings.idsheet);

  const sheetR = ss.getSheetByName("Roles");
  const dataR  = sheetR.getDataRange().getValues();
  for (let i = dataR.length - 1; i >= 1; i--) {
    if (dataR[i][0]?.toString() === roleId) {
      sheetR.deleteRow(i + 1);
      break;
    }
  }

  const sheetUR = ss.getSheetByName("User_Roles");
  const dataUR  = sheetUR.getDataRange().getValues();
  for (let i = dataUR.length - 1; i >= 1; i--) {
    if (dataUR[i][1]?.toString() === roleId) sheetUR.deleteRow(i + 1);
  }

  const sheetRP = ss.getSheetByName("Role_Permissions");
  const dataRP  = sheetRP.getDataRange().getValues();
  for (let i = dataRP.length - 1; i >= 1; i--) {
    if (dataRP[i][1]?.toString() === roleId) sheetRP.deleteRow(i + 1);
  }

  const sheetRA = ss.getSheetByName("Role_Actions");
  const dataRA  = sheetRA.getDataRange().getValues();
  for (let i = dataRA.length - 1; i >= 1; i--) {
    if (dataRA[i][0]?.toString() === roleId) sheetRA.deleteRow(i + 1);
  }

  return { success: true };
};

const updateRoleAction = (obj) => {
  const settings = getSettingsData();
  try {
    const sheet = SpreadsheetApp.openById(settings.idsheet).getSheetByName('Role_Actions');
    const data  = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0]?.toString() === obj.roleId &&
          data[i][1]?.toString() === obj.action) {
        sheet.getRange(i + 1, 3).setValue(obj.enabled);
        return { success: true };
      }
    }

    sheet.appendRow([obj.roleId, obj.action, obj.enabled]);
    return { success: true };

  } catch (e) {
    console.log('Error updateRoleAction:', e);
    return { success: false, message: e.message };
  }
};