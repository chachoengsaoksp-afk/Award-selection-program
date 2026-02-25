const SS_ID = '1zMTM-sxOXeevx-bRJR7FudqFdmEEC1vyesm4tXG9nHM';

// กำหนด Username และ Password ของกรรมการ
const USERS = {
  'admin': '1234',
  'กรรมการ01': 'ksp01',
  'กรรมการ02': 'ksp02'
};

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('ระบบบันทึกคะแนน - คุรุสภาฉะเชิงเทรา')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ตรวจสอบการเข้าสู่ระบบ
function checkLogin(username, password) {
  if (USERS[username] && USERS[username] === password) {
    return { success: true, user: username };
  }
  return { success: false, message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
}

// ดึงรายชื่อผู้สมัครจากชีต "ผู้ส่ง"
function getApplicantList() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName('ผู้ส่ง');
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    return sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  } catch (e) { return []; }
}

// บันทึกคะแนนและจัดลำดับ
function processForm(formData) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const name = formData.applicantName;
    const workName = formData.workName;
    const type = formData.professionType;
    const judgeName = formData.judgeName;
    
    const scores = [
      Number(formData.score1) || 0,
      Number(formData.score2) || 0,
      Number(formData.score3) || 0,
      Number(formData.score4) || 0,
      Number(formData.score5) || 0
    ];
    const total = scores.reduce((a, b) => a + b, 0);
    const timestamp = new Date();

    // เลือกชีตเป้าหมายตามวิชาชีพ
    let targetName = type;
    if (type === 'ศึกษานิเทศก์') targetName = 'ศน.';
    
    const targetSheet = ss.getSheetByName(targetName);
    if (!targetSheet) return { success: false, message: 'ไม่พบชีต: ' + targetName };
    
    // บันทึกข้อมูลลงชีตวิชาชีพนั้นๆ
    targetSheet.appendRow([timestamp, name, workName, type, ...scores, total, judgeName]);

    // อัปเดตการจัดลำดับลง "ชีต6"
    updateSummarySheet(ss);

    return { success: true, message: 'บันทึกคะแนนเรียบร้อยแล้ว' };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

// ฟังก์ชันดึงข้อมูลทุกวิชาชีพมาเรียงลำดับใหม่ลง "ชีต6"
function updateSummarySheet(ss) {
  const categories = ['ครู', 'ผู้บริหารสถานศึกษา', 'ผู้บริหารการศึกษา', 'ศน.'];
  let allSortedData = [];

  categories.forEach(cat => {
    const sheet = ss.getSheetByName(cat);
    if (sheet && sheet.getLastRow() > 1) {
      let data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).getValues();
      // เรียงตามคะแนนรวม (index 9) จากมากไปน้อย
      data.sort((a, b) => b[9] - a[9]);
      
      data.forEach((row, idx) => {
        allSortedData.push([idx + 1, row[1], row[3], row[2], row[9]]);
      });
    }
  });

  const summarySheet = ss.getSheetByName('ชีต6');
  summarySheet.clearContents();
  summarySheet.getRange(1, 1, 1, 5).setValues([['ลำดับ', 'ชื่อ-สกุล', 'ประเภท', 'ผลงาน', 'คะแนนรวม']]);
  if (allSortedData.length > 0) {
    summarySheet.getRange(2, 1, allSortedData.length, 5).setValues(allSortedData);
  }
}

function getSummaryData() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName('ชีต6');
    if (!sheet || sheet.getLastRow() < 2) return [];
    return sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  } catch (e) { return []; }
}
