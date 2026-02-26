const SPREADSHEET_ID = '1rymhMcFuDRQIHO1KQ4tyqfEc5nXYdT_sIlTNqAhqTp8';
const MIN_JUDGES = 3;//แก้จำนวนกรรมการ

/* ================= HELPERS ================= */
function getTypesFromCandidates(ss){
  const s = ss.getSheetByName('ผู้ส่ง');
  if(!s) return [];
  const data = s.getDataRange().getValues();
  if(data.length < 2) return [];
  data.shift();
  const types = Array.from(new Set(data.map(r => (r[2] || '').toString().trim()).filter(x => x)));
  return types;
}

/* ================= LOGIN / ENTRY ================= */
function doGet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  try { updateRanking(ss); } catch(e){ /* ignore errors */ }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ระบบให้คะแนน')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function checkLogin(username, password) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('กรรมการ');
  if (!sheet) return {success:false,message:"ไม่พบชีท 'กรรมการ' กรุณาสร้างชีทและกรอกบัญชีกรรมการ"};
  const data = sheet.getDataRange().getValues();
  if(data.length < 2) return {success:false,message:"ไม่มีบัญชีกรรมการในชีท"};
  data.shift();
  const found = data.find(r =>
    r[0].toString().trim() === (username || '').toString().trim() &&
    r[1].toString().trim() === (password || '').toString().trim()
  );
  return found ? {success:true,name:username} : {success:false,message:"Username หรือ Password ไม่ถูกต้อง"};
}

/* ================= ผู้ส่ง ================= */
function getCandidates(){
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('ผู้ส่ง');
  if(!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if(data.length < 2) return [];
  data.shift();
  return data.map(r => ({
    name: r[0],
    work: r[1],
    type: r[2]
  }));
}

/* ================= ตรวจซ้ำ + บันทึกคะแนน ================= */
function submitScore(data){
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(data.type);
  if(!sheet){
    sheet = ss.insertSheet(data.type);
    sheet.appendRow(['วันที่','กรรมการ','ชื่อ','ผลงาน','คะแนน']);
  }
  const allData = sheet.getDataRange().getValues();
  const rows = allData.length > 1 ? allData.slice(1) : [];
  const duplicate = rows.some(r => (r[1] === data.judge) && (r[2] === data.name));
  if(duplicate){
    return "ท่านได้ให้คะแนนผู้สมัครรายนี้แล้ว ❌";
  }
  sheet.appendRow([
    new Date(),
    data.judge,
    data.name,
    data.work,
    Number(data.score)
  ]);
  try { updateRanking(ss); } catch(e){ console.error('updateRanking error: ' + e); }
  return "บันทึกสำเร็จ ✅";
}

/* ================= คำนวณอันดับ ================= */
function updateRanking(ss){
  const types = getTypesFromCandidates(ss);
  if(types.length === 0) return;
  let rankingSheet = ss.getSheetByName('อันดับ');
  if(!rankingSheet) rankingSheet = ss.insertSheet('อันดับ');
  rankingSheet.clear();
  rankingSheet.appendRow(['ประเภท','ชื่อ','เฉลี่ย','กรรมการ','อันดับ','รางวัล']);
  types.forEach(type => {
    const sheet = ss.getSheetByName(type);
    if(!sheet) return;
    const data = sheet.getDataRange().getValues();
    if(data.length < 2) return;
    data.shift();
    const scoreMap = {};
    data.forEach(r => {
      const name = r[2];
      const score = parseFloat(r[4]);
      if(name && !isNaN(score)){
        if(!scoreMap[name]) scoreMap[name] = [];
        scoreMap[name].push(score);
      }
    });
    const results = Object.keys(scoreMap).map(name => {
      const scores = scoreMap[name];
      return {
        type: type,
        name: name,
        avg: parseFloat((scores.reduce((a,b) => a+b,0) / scores.length).toFixed(2)),
        count: scores.length
      };
    })
    .filter(x => x.count >= MIN_JUDGES)
    .sort((a,b) => b.avg - a.avg);
    let rank = 0, prev = null, index = 0;
    results.forEach(item => {
      index++;
      if(item.avg !== prev) rank = index;
      let medal = "";
      if(rank === 1) medal = "🥇";
      else if(rank === 2) medal = "🥈";
      else if(rank === 3) medal = "🥉";
      rankingSheet.appendRow([item.type, item.name, item.avg, item.count, rank, medal]);
      prev = item.avg;
    });
  });
}

/* ================= ดึงอันดับ (เฉพาะผู้ที่ผ่านเงื่อนไข) ================= */
function getRanking(){
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('อันดับ');
  if(!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if(data.length < 2) return [];
  data.shift();
  return data;
}

/* ================= ดึงรายการครบทุกคน (รวมผู้ที่ยังไม่ได้รับการให้คะแนน) ================= */
function getFullRanking(){
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sendSheet = ss.getSheetByName('ผู้ส่ง');
  let candidates = [];
  if (sendSheet) {
    const data = sendSheet.getDataRange().getValues();
    if (data.length > 1) {
      data.shift();
      candidates = data.map(r => ({
        name: (r[0] || '').toString(),
        work: (r[1] || '').toString(),
        type: (r[2] || '').toString()
      }));
    }
  }
  const rankSheet = ss.getSheetByName('อันดับ');
  let rankMap = {};
  if (rankSheet) {
    const rdata = rankSheet.getDataRange().getValues();
    if (rdata.length > 1) {
      rdata.shift();
      rdata.forEach(function(row){
        const type = (row[0] || '').toString();
        const name = (row[1] || '').toString();
        const avg = row[2] === undefined || row[2] === null ? '' : row[2];
        const count = row[3] === undefined || row[3] === null ? '' : row[3];
        const rank = row[4] === undefined || row[4] === null ? '' : row[4];
        const medal = row[5] === undefined || row[5] === null ? '' : row[5];
        const key = type + '|' + name;
        rankMap[key] = { avg: avg, count: count, rank: rank, medal: medal };
      });
    }
  }
  const result = candidates.map(function(c){
    const key = (c.type || '') + '|' + (c.name || '');
    const info = rankMap[key] || { avg: '', count: '', rank: '', medal: '' };
    const avgVal = (info.avg === '' || info.avg === null) ? "ยังไม่ได้รับการให้คะแนน" : info.avg;
    const rankVal = (info.rank === '' || info.rank === null) ? "ยังไม่ได้รับการให้คะแนน" : info.rank;
    const countVal = (info.count === '' || info.count === null) ? "" : info.count;
    const medalVal = (info.medal === '' || info.medal === null) ? "" : info.medal;
    return [ c.type || '', c.name || '', avgVal, countVal, rankVal, medalVal ];
  });
  return result;
}

/* ================= PDF ================= */
function exportPDF(){
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('อันดับ');
  if(!sheet) throw new Error("ไม่พบชีท 'อันดับ'");
  const url = ss.getUrl().replace(/edit$/,'') +
    'export?format=pdf&gid=' + sheet.getSheetId() +
    '&size=A4&portrait=true&fitw=true&gridlines=false';
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + token }
  });
  const blob = response.getBlob().setName("รายงานผลการจัดอันดับ.pdf");
  const file = DriveApp.createFile(blob);
  return file.getUrl();
}
