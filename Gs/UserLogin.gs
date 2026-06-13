const checkLogin = (username, password) => {
  const settings = getSettingsData();
  const ssA     = SpreadsheetApp.openById(settings.idsheet);
  const users   = ssA.getSheetByName("Users").getDataRange().getValues();
  const uroles  = ssA.getSheetByName("User_Roles").getDataRange().getValues();
  const rperms  = ssA.getSheetByName("Role_Permissions").getDataRange().getValues();
  const umenus  = ssA.getSheetByName("User_Menu").getDataRange().getValues();
  const ractions = ssA.getSheetByName("Role_Actions").getDataRange().getValues();

  for (let i = 1; i < users.length; i++) {
    const row = users[i];
    if (
      row[1].toString().toLowerCase() === username.toLowerCase() &&
      row[2].toString()               === password &&
      row[10].toString()               === "active"
    ) {
      ssA.getSheetByName("Users").getRange(i + 1, 13).setValue("online");

      const userId = row[0].toString();

      const userRoles = uroles
        .slice(1)
        .filter(r => r[0].toString() === userId)
        .map(r => r[1].toString());

      const allowedMenuIds = new Set(
        rperms.slice(1)
          .filter(r =>
            r[0].toString() === userId &&
            userRoles.includes(r[1].toString())
          )
          .map(r => r[2].toString())
      );

      const allMenuCodes     = [];
      const allowedMenuCodes = [];
      umenus.slice(1).forEach(r => {
        const menuId   = r[0].toString();
        const menuCode = r[1].toString();
        if (menuCode) {
          allMenuCodes.push(menuCode);
          if (allowedMenuIds.has(menuId)) allowedMenuCodes.push(menuCode);
        }
      });

      const roleActions = ractions.slice(1).map(r => ({
        roleId  : r[0].toString(),
        action  : r[1].toString(),
        enabled : r[2]?.toString().toLowerCase() || 'false',
      }));

      return {
        uiduser:        userId,
        username:       row[1],
        password:       row[2],
        email:          row[3],
        firstname:      row[4],
        lastname:       row[5],
        department:     row[6],
        positions:     row[7],
        imgUser:        row[8],
        sigUser:        row[9],
        status:         row[10],
        action:         "online",
        allMenuCodes,
        allowedMenuCodes,
        userRoles,
        roleActions,
      };
    }
  }

  return 'ชื่อผู้ใช้งานหรือรหัสผ่านไม่ถูกต้อง';
};

const updateActionInSheet = (uid, status) => {
  const settings = getSettingsData();
  const sheet = SpreadsheetApp.openById(settings.idsheet).getSheetByName("Users");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === uid) {
      sheet.getRange(i + 1, 13).setValue(status); 
      return "success";
    }
  }
  return "error";
};