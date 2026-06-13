const sendForgotPassword = (username) => {
  const settings = getSettingsData();
  const sheet = SpreadsheetApp.openById(settings.idsheet).getSheetByName("Users");
  const data = sheet.getDataRange().getValues();
  
  const user = data.find(row => row[1] && row[1].toString().toLowerCase() === username.toLowerCase());

  if (user) {
    const password = user[2];
    const email = user[3];
    const fName = user[4];
    const lName = user[5];
    
    if (!email || email.toString().trim() === "") {
      return { success: false, message: "ไม่พบข้อมูลอีเมลที่เชื่อมโยงกับบัญชีนี้" };
    }

    const subject = `🔐 คำขอรหัสผ่านจากระบบ ${settings.nameEng}`;
    const htmlBody = `
      <html>
        <body style="font-family: Arial, sans-serif; margin: 0; padding: 20px;">
          <div style="max-width: 600px; margin: auto; background-color: var(--bg-secondary, #ffffff); border: 2px solid var(--border-color, #e9ecef); border-radius: 1rem; overflow: hidden;">
            
            <div style="color: #333; padding: 25px 20px; text-align: center;">
              <img src="${settings.logoSystem}" 
                   alt="${settings.nameEng} Logo" style="width: 80px; height: 80px; object-fit: contain; margin-bottom: 10px;">
              <h2 style="margin: 0; font-size: 24px;">แจ้งข้อมูลการขอรหัสผ่านใหม่</h2>
              <p style="margin: 10px 0 0; font-size: 15px; color: #333;">เรียนคุณ ${fName} ${lName}</p>
            </div>

            <div style="padding: 30px 20px;">
              <p style="font-size: 16px; color: #333; margin-bottom: 10px; text-align: center;">
                คุณได้ทำการขอรหัสผ่านจากระบบ <strong>${settings.nameEng}</strong>
              </p>
              
              <div style="background-color: #f8f9fa; border: 1px solid #e9ecef; border-radius: 8px; padding: 20px; margin-bottom: 20px; text-align: center;">
                <p style="font-size: 18px; margin: 10px 0; color: #333;"><strong>👨‍💻 Username:</strong> ${user[1]}</p>
                <p style="font-size: 18px; margin: 10px 0; color: #333;"><strong>🔐 Password:</strong> ${password}</p>
              </div>

              <div style="text-align: center;">
                <p style="font-size: 14px; color: #dc3545; font-weight: bold; margin-top: 20px;">
                  ⚠️ กรุณาเก็บรหัสผ่านนี้ไว้เป็นความลับ
                </p>
              </div>
            </div>

            <div style="background-color: #f1f1f1; padding: 15px; border-top: 1px solid #dee2e6; text-align: center;">
              <p style="color: #666; font-size: 14px; margin: 0 0 5px 0;">หากคุณไม่ได้เป็นผู้ร้องขอ กรุณาเพิกเฉยต่ออีเมลนี้</p>
              <p style="color: #888; font-size: 12px; margin: 0;">* อีเมลนี้ถูกส่งโดยระบบอัตโนมัติ กรุณาอย่าตอบกลับ</p>
            </div>
          </div>
        </body>
      </html>
    `;

    try {
      MailApp.sendEmail({
        to: email,
        subject: subject,
        htmlBody: htmlBody
      });
      return { success: true };
    } catch (e) {
      console.log("Error sending email: " + e);
      return { success: false, message: "ระบบส่งอีเมลขัดข้อง กรุณาลองใหม่ภายหลัง" };
    }
    
  } else {
    return { success: false, message: "ไม่พบชื่อผู้ใช้งานนี้ในระบบ" };
  }
};

const generateCodeUser = (sheet) => {
  let isUnique = false;
  let newID = '';

  const lastRow = sheet.getLastRow();
  let existingIDs = [];
  if (lastRow > 1) {
    existingIDs = sheet.getRange(2, 1, lastRow - 1).getValues().flat();
  }

  while (!isUnique) {
    const randomHex = Utilities.getUuid().replace(/-/g, '').toLowerCase().substring(0, 12);
    newID = 'usr_' + randomHex;

    if (!existingIDs.includes(newID)) {
      isUnique = true;
    }
  }

  return newID;
};

const registers = (obj) => {
  const settings = getSettingsData();
  const sheet = SpreadsheetApp.openById(settings.idsheet).getSheetByName("Users");
  const codeID = generateCodeUser(sheet);
  const folder = DriveApp.getFolderById(settings.idfolder);
  const currentTime = new Date();
  const formattedDate = Utilities.formatDate(currentTime, Session.getScriptTimeZone(), 'd/M/yyyy');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] && data[i][1].toString().toLowerCase() === obj.username.toLowerCase()) {
      return false;
    }
  }

  let profileUrl = "";
  let signatureUrl = "";

  if (obj.profileCheck === "upload" && obj.profileImageDataUrl) {
    try {
      let profileData = Utilities.base64Decode(obj.profileImageDataUrl.split(',')[1]);
      let profileBlob = Utilities.newBlob(profileData, obj.profileFiletype, obj.profileFilename);
      let profileFile = folder.createFile(profileBlob);
      let profileFileId = profileFile.getId();
      profileUrl = "https://lh3.googleusercontent.com/d/" + profileFileId;
    } catch (error) {
      console.log("Error uploading profile image: " + error);
      profileUrl = "";
    }
  } else {
    profileUrl = obj.profileImageDataUrl || "";
  }

  if (obj.signatureCheck === "upload" && obj.signatureImageDataUrl) {
    try {
      let signatureData = Utilities.base64Decode(obj.signatureImageDataUrl.split(',')[1]);
      let signatureBlob = Utilities.newBlob(signatureData, obj.signatureFiletype, obj.signatureFilename);
      let signatureFile = folder.createFile(signatureBlob);
      let signatureFileId = signatureFile.getId();
      signatureUrl = "https://lh3.googleusercontent.com/d/" + signatureFileId;
    } catch (error) {
      console.log("Error uploading signature image: " + error);
      signatureUrl = "";
    }
  }

  const fName = obj.firstName || (obj.fullname ? obj.fullname.split(" ")[0] : "");
  const lName = obj.lastName || (obj.fullname ? obj.fullname.split(" ").slice(1).join(" ") : "");

  sheet.appendRow([codeID, "'" + obj.username, "'" + obj.password, obj.email, "'" + fName, "'" + lName, "", "", profileUrl, signatureUrl, "", formattedDate]);

  if (obj.email && obj.email.trim() !== "") {
    const subject = `🎉 ยินดีต้อนรับเข้าสู่ระบบ ${settings.nameEng}`;
    const htmlBody = `
      <html>
        <body style="font-family: Arial, sans-serif; margin: 0; padding: 20px;">
          <div style="max-width: 600px; margin: auto; background-color: var(--bg-secondary, #ffffff); border: 2px solid var(--border-color, #e9ecef); border-radius: 1rem; overflow: hidden;">
            
            <div style="color: #333; padding: 25px 20px; text-align: center;">
              <img src="${settings.logoSystem}" 
                   alt="${settings.nameEng} Logo" style="width: 80px; height: 80px; object-fit: contain; margin-bottom: 10px;">
              <h2 style="margin: 0; font-size: 24px;">ยินดีต้อนรับคุณ ${fName} ${lName}</h2>
              <p style="margin: 10px 0 0; font-size: 15px; color: #333;">การสมัครสมาชิกบนระบบ ${settings.nameEng} เสร็จสมบูรณ์</p>
            </div>

            <div style="padding: 30px 20px;">
              <p style="font-size: 16px; color: #333; margin-bottom: 20px; text-align: center; line-height: 1.5;">
                ขอบคุณที่สมัครสมาชิก นี่คือข้อมูลบัญชีของคุณสำหรับใช้เข้าสู่ระบบ
              </p>
              
              <div style="background-color: #f8f9fa; border: 1px solid #e9ecef; border-radius: 8px; padding: 20px; margin-bottom: 20px;">
                <p style="font-size: 16px; margin: 10px 0; color: #333;"><strong>👤 Username:</strong> ${obj.username}</p>
                <p style="font-size: 16px; margin: 10px 0; color: #333;"><strong>🔐 Password:</strong> ${obj.password}</p>
                <p style="font-size: 16px; margin: 10px 0; color: #333;"><strong>📧 Email:</strong> ${obj.email}</p>
              </div>

              <div style="text-align: center; margin-top: 25px;">
                <p style="font-size: 14px; color: #dc3545; font-weight: bold; margin: 0;">
                  ⚠️ กรุณาเก็บรหัสผ่านนี้ไว้เป็นความลับและห้ามเปิดเผยให้ผู้อื่นทราบ
                </p>
              </div>
            </div>

            <div style="background-color: #f1f1f1; padding: 15px; text-align: center; border-top: 1px solid #dee2e6;">
              <p style="color: #666; font-size: 14px; margin: 0 0 5px 0;">หากคุณมีข้อสงสัย สามารถติดต่อผู้ดูแลระบบ ${settings.nameEng} ได้ทันที</p>
              <p style="color: #888; font-size: 12px; margin: 0;">* อีเมลนี้ถูกส่งโดยระบบอัตโนมัติ กรุณาอย่าตอบกลับ</p>
            </div>

          </div>
        </body>
      </html>
    `;

    try {
      MailApp.sendEmail({
        to: obj.email,
        subject: subject,
        htmlBody: htmlBody
      });
    } catch (e) {
      console.log("ไม่สามารถส่งอีเมลได้: " + e);
    }
  }

  return true;
};

const addUsers = (obj) => {
  const settings = getSettingsData();
  const sheet = SpreadsheetApp.openById(settings.idsheet).getSheetByName("Users");
  const codeID = generateCodeUser(sheet);
  const currentTime = new Date();
  const formattedDate = Utilities.formatDate(currentTime, Session.getScriptTimeZone(), 'd/M/yyyy');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] && data[i][1].toString().toLowerCase() === obj.username.toLowerCase()) {
      return false;
    }
  }

  sheet.appendRow([codeID, "'" + obj.username, "'" + obj.password, obj.email, "'" + obj.firstName, "'" + obj.lastName, obj.dept, obj.pos, "", "", "active", formattedDate, "offline"]);

  if (obj.email && obj.email.trim() !== "") {
    const subject = `🎉 ยินดีต้อนรับเข้าสู่ระบบ ${nameEng}`;
    const htmlBody = `
      <html>
        <body style="font-family: Arial, sans-serif; margin: 0; padding: 20px;">
          <div style="max-width: 600px; margin: auto; border: 2px solid #e9ecef; border-radius: 1rem; overflow: hidden;">
            <div style="padding: 25px 20px; text-align: center;">
              <img src="${settings.logoSystem}" alt="${settings.nameEng}" style="width:80px;height:80px;object-fit:contain;margin-bottom:10px;">
              <h2 style="margin:0;font-size:24px;">ยินดีต้อนรับคุณ ${obj.firstName} ${obj.lastName}</h2>
              <p style="margin:10px 0 0;font-size:15px;color:#333;">บัญชีของคุณถูกสร้างโดยผู้ดูแลระบบ ${settings.nameEng}</p>
            </div>
            <div style="padding:30px 20px;">
              <div style="background:#f8f9fa;border:1px solid #e9ecef;border-radius:8px;padding:20px;margin-bottom:20px;">
                <p style="font-size:16px;margin:10px 0;color:#333;"><strong>👤 Username:</strong> ${obj.username}</p>
                <p style="font-size:16px;margin:10px 0;color:#333;"><strong>🔐 Password:</strong> ${obj.password}</p>
                <p style="font-size:16px;margin:10px 0;color:#333;"><strong>📧 Email:</strong> ${obj.email}</p>
              </div>
              <p style="font-size:14px;color:#dc3545;font-weight:bold;text-align:center;margin:0;">
                ⚠️ กรุณาเก็บรหัสผ่านนี้ไว้เป็นความลับ
              </p>
            </div>
            <div style="background:#f1f1f1;padding:15px;text-align:center;border-top:1px solid #dee2e6;">
              <p style="color:#888;font-size:12px;margin:0;">* อีเมลนี้ถูกส่งโดยระบบอัตโนมัติ กรุณาอย่าตอบกลับ</p>
            </div>
          </div>
        </body>
      </html>`;

    try {
      MailApp.sendEmail({ to: obj.email, subject, htmlBody });
    } catch (e) {
      console.log("ไม่สามารถส่งอีเมลได้: " + e);
    }
  }

  return { success: true, userId: codeID };
};

const updateUsers = (obj) => {
  const settings = getSettingsData();
  const sheet = SpreadsheetApp.openById(settings.idsheet).getSheetByName("Users");
  const data  = sheet.getDataRange().getValues();

  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]?.toString() === obj.userId) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) return { success: false, message: "ไม่พบผู้ใช้งาน" };

  sheet.getRange(rowIndex, 2).setValue("'" + obj.username);
  sheet.getRange(rowIndex, 3).setValue("'" + obj.password);
  sheet.getRange(rowIndex, 4).setValue(obj.email);
  sheet.getRange(rowIndex, 5).setValue("'" + obj.firstName);
  sheet.getRange(rowIndex, 6).setValue("'" + obj.lastName);
  sheet.getRange(rowIndex, 7).setValue(obj.dept);
  sheet.getRange(rowIndex, 8).setValue(obj.pos);
  sheet.getRange(rowIndex, 11).setValue(obj.status);

  return { success: true };
};

const delUsers = (userId) => {
  const settings = getSettingsData();
  const ss = SpreadsheetApp.openById(settings.idsheet);

  const sheetU = ss.getSheetByName("Users");
  const dataU  = sheetU.getDataRange().getValues();
  for (let i = dataU.length - 1; i >= 1; i--) {
    if (dataU[i][0]?.toString() === userId) {
      sheetU.deleteRow(i + 1);
      break;
    }
  }

  const sheetUR = ss.getSheetByName("User_Roles");
  const dataUR  = sheetUR.getDataRange().getValues();
  for (let i = dataUR.length - 1; i >= 1; i--) {
    if (dataUR[i][0]?.toString() === userId) sheetUR.deleteRow(i + 1);
  }

  const sheetRP = ss.getSheetByName("Role_Permissions");
  const dataRP  = sheetRP.getDataRange().getValues();
  for (let i = dataRP.length - 1; i >= 1; i--) {
    if (dataRP[i][0]?.toString() === userId) sheetRP.deleteRow(i + 1);
  }

  const sheetED = ss.getSheetByName("Employee_Details");
  const dataED  = sheetED.getDataRange().getValues();
  for (let i = dataED.length - 1; i >= 1; i--) {
    if (dataED[i][0]?.toString() === userId) {
      sheetED.deleteRow(i + 1);
      break;
    }
  }

  return { success: true };
};

const addEmployeeDetails = (obj) => {
  try {
    const settings = getSettingsData();
    const sheet = SpreadsheetApp.openById(settings.idsheet).getSheetByName("Employee_Details");

    sheet.appendRow([
      obj.empId,
      "'" + (obj.taxId       || ''),
      obj.employmentType      || 'full-time',
      obj.bankName            || '',
      "'" + (obj.bankAccount  || ''),
      obj.baseSalary          || 0,
      obj.sso                 || 0,
      obj.providentFund       || 0,
      obj.remark              || '',
    ]);
    return { success: true };
  } catch (e) {
    console.log('addEmployeeDetails error: ' + e);
    return { success: false, message: e.toString() };
  }
};

const updateEmployeeDetails = (obj) => {
  try {
    const settings = getSettingsData();
    const sheet = SpreadsheetApp.openById(settings.idsheet).getSheetByName("Employee_Details");
    const data  = sheet.getDataRange().getValues();

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]?.toString() === obj.empId?.toString()) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) return addEmployeeDetails(obj);

    sheet.getRange(rowIndex, 2).setValue("'" + (obj.taxId      || ''));
    sheet.getRange(rowIndex, 3).setValue(obj.employmentType     || 'full-time');
    sheet.getRange(rowIndex, 4).setValue(obj.bankName           || '');
    sheet.getRange(rowIndex, 5).setValue("'" + (obj.bankAccount || ''));
    sheet.getRange(rowIndex, 6).setValue(obj.baseSalary         || 0);
    sheet.getRange(rowIndex, 7).setValue(obj.sso                || 0);
    sheet.getRange(rowIndex, 8).setValue(obj.providentFund      || 0);
    sheet.getRange(rowIndex, 9).setValue(obj.remark             || '');

    return { success: true };
  } catch (e) {
    console.log('updateEmployeeDetails error: ' + e);
    return { success: false, message: e.toString() };
  }
};

const approveUser = (obj) => {
  const settings = getSettingsData();
  const sheet = SpreadsheetApp.openById(settings.idsheet).getSheetByName("Users");
  const data  = sheet.getDataRange().getValues();

  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]?.toString() === obj.userId) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) return { success: false, message: "ไม่พบผู้ใช้งาน" };

  sheet.getRange(rowIndex, 7).setValue(obj.dept);
  sheet.getRange(rowIndex, 8).setValue(obj.pos);
  sheet.getRange(rowIndex, 11).setValue(obj.status);
  sheet.getRange(rowIndex, 13).setValue("offline");

  return { success: true };
};

const saveUserRoles = (obj) => {
  const settings = getSettingsData();
  const ss = SpreadsheetApp.openById(settings.idsheet);

  const sheetUR = ss.getSheetByName("User_Roles");
  const dataUR  = sheetUR.getDataRange().getValues();

  for (let i = dataUR.length - 1; i >= 1; i--) {
    if (dataUR[i][0]?.toString() === obj.userId) {
      sheetUR.deleteRow(i + 1);
    }
  }
  obj.roles.forEach(roleId => sheetUR.appendRow([obj.userId, roleId]));

  if (obj.roleMenuMap !== null && obj.roleMenuMap !== undefined) {
    const sheetRP = ss.getSheetByName("Role_Permissions");
    const dataRP  = sheetRP.getDataRange().getValues();

    for (let i = dataRP.length - 1; i >= 1; i--) {
      if (dataRP[i][0]?.toString() === obj.userId) {
        sheetRP.deleteRow(i + 1);
      }
    }

    Object.entries(obj.roleMenuMap).forEach(([roleId, menuIds]) => {
      menuIds.forEach(menuId => {
        sheetRP.appendRow([obj.userId, roleId, menuId]);
      });
    });
  }

  return { success: true };
};

const updateUserProfile = (obj) => {
  const settings = getSettingsData();
  const sheet = SpreadsheetApp.openById(settings.idsheet).getSheetByName("Users");
  const data = sheet.getDataRange().getValues();
  
  let rowIndex = -1;
  let existingProfileUrl = "";
  let existingSignatureUrl = "";
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString() === obj.uid) {
      rowIndex = i + 1;
      existingProfileUrl = data[i][8];
      existingSignatureUrl = data[i][9];
      break;
    }
  }
  
  if (rowIndex === -1) {
    return { success: false, message: "ไม่พบรหัสผู้ใช้นี้ในระบบ" };
  }

  let finalProfileUrl = existingProfileUrl;
  if (obj.profileCheck === "upload" && obj.profileImageDataUrl.startsWith('data:image')) {
    try {
      const folder = DriveApp.getFolderById(settings.idfolder);
      let profileData = Utilities.base64Decode(obj.profileImageDataUrl.split(',')[1]);
      let profileBlob = Utilities.newBlob(profileData, obj.profileFiletype, obj.profileFilename);
      let profileFile = folder.createFile(profileBlob);
      finalProfileUrl = "https://lh3.googleusercontent.com/d/" + profileFile.getId(); 
    } catch (error) {
      console.log("Error uploading updated profile image: " + error);
    }
  } else if (obj.profileCheck === "url" && obj.profileImageDataUrl !== existingProfileUrl) {
    finalProfileUrl = obj.profileImageDataUrl;
  }

  let finalSignatureUrl = existingSignatureUrl;
  if (obj.signatureCheck === "upload" && obj.signatureImageDataUrl.startsWith('data:image')) {
    try {
      const folder = DriveApp.getFolderById(settings.idfolder);
      let signatureData = Utilities.base64Decode(obj.signatureImageDataUrl.split(',')[1]);
      let signatureBlob = Utilities.newBlob(signatureData, obj.signatureFiletype, obj.signatureFilename);
      let signatureFile = folder.createFile(signatureBlob);
      finalSignatureUrl = "https://lh3.googleusercontent.com/d/" + signatureFile.getId(); 
    } catch (error) {
      console.log("Error uploading updated signature image: " + error);
    }
  }

  sheet.getRange(rowIndex, 2).setValue("'" + obj.username);
  sheet.getRange(rowIndex, 3).setValue("'" + obj.password);
  sheet.getRange(rowIndex, 4).setValue(obj.email);
  sheet.getRange(rowIndex, 5).setValue("'" + obj.firstname);
  sheet.getRange(rowIndex, 6).setValue("'" + obj.lastname);
  sheet.getRange(rowIndex, 9).setValue(finalProfileUrl);
  sheet.getRange(rowIndex, 10).setValue(finalSignatureUrl);

  const updatedUser = {
    username: obj.username,
    password: obj.password,
    email: obj.email,
    firstname: obj.firstname,
    lastname: obj.lastname,
    imgUser: finalProfileUrl,
    sigUser: finalSignatureUrl
  };

  return { success: true, data: updatedUser };
};