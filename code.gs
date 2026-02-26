const SPREADSHEET_ID = '1rymhMcFuDRQIHO1KQ4tyqfEc5nXYdT_sIlTNqAhqTp8';
const MIN_JUDGES = 3;

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

/* ================= ENTRY / doGet ================= */
function doGet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  try { updateRanking(ss); } catch(e){ /* ignore errors */ }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ระบบให้คะแนน')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ================= GATE: CHECK JUDGE LOGIN ================= */
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

/* ================= ADMIN LOGIN FOR CRITERIA MANAGER ================= */
function adminLogin(user, pass){
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName('ADMIN');
  if(!sh) return {status:false, message: "ไม่พบชีท ADMIN"};
  const data = sh.getDataRange().getValues();
  if(data.length < 1) return {status:false, message: "ไม่มีผู้ดูแลในชีท ADMIN"};
  // iterate through rows (support header)
  for(let i=0;i<data.length;i++){
    const u = (data[i][0] || '').toString().trim();
    const p = (data[i][1] || '').toString().trim();
    if(u === user && p === pass) return {status:true};
  }
  return {status:false, message: "รหัสผู้จัดการเกณฑ์ไม่ถูกต้อง"};
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

/* ================= CRITERIA (เกณฑ์) =================
   Sheet name: 'เกณฑ์'
   Columns: ประเภท | ตัวชี้วัด | max_score | weight | status
*/
function getCriteria(){
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('เกณฑ์');
  if(!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if(data.length < 2) return [];
  data.shift();
  return data.map(r => ({
    type: (r[0] || '').toString(),
    name: (r[1] || '').toString(),
    max: Number(r[2]) || 0,
    weight: Number(r[3]) || 0,
    status: (r[4] === undefined ? '' : r[4])
  }));
}

/*
 saveCriteria(type, criteriaArray)
 criteriaArray = [{name, max, weight, status}, ...]
 Validate: sum(weights) must be 100
*/
function saveCriteria(type, criteriaArray){
  // validate weights
  let totalWeight = 0;
  for(let i=0;i<criteriaArray.length;i++){
    totalWeight += Number(criteriaArray[i].weight || 0);
  }
  // allow small float tolerance
  if(Math.abs(totalWeight - 100) > 0.01){
    return { success:false, message: "น้ำหนักรวมต้องเท่ากับ 100% (ตอนนี้รวม = " + totalWeight + "%)" };
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('เกณฑ์');
  if(!sheet) sheet = ss.insertSheet('เกณฑ์');
  // read existing and keep other types
  const all = sheet.getDataRange().getValues();
  const header = ['ประเภท','ตัวชี้วัด','max_score','weight','status'];
  let rows = [];
  if(all.length > 0){
    for(let i=1;i<all.length;i++){
      const r = all[i];
      if((r[0]||'') !== type){
        rows.push([r[0]||'', r[1]||'', r[2]||'', r[3]||'', r[4]||'']);
      }
    }
  }
  // append new rows for this type (with weight)
  criteriaArray.forEach(function(c){
    rows.push([ type, c.name || '', c.max || 0, c.weight || 0, (c.status===undefined?'เปิด':c.status) ]);
  });
  // rewrite sheet
  sheet.clear();
  sheet.appendRow(header);
  if(rows.length > 0) sheet.getRange(2,1,rows.length,5).setValues(rows);
  return { success:true, message: "บันทึกเกณฑ์เรียบร้อย" };
}

/* Optional: admin can upload whole table (2D array) to overwrite sheet entirely */
function saveCriteriaTable(table2D){
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('เกณฑ์');
  if(!sheet) sheet = ss.insertSheet('เกณฑ์');
  sheet.clear();
  if(!table2D || table2D.length === 0) return { success:false, message:"ไม่มีข้อมูลให้บันทึก" };
  sheet.getRange(1,1,table2D.length, table2D[0].length).setValues(table2D);
  return { success:true, message:"บันทึกตารางเกณฑ์เรียบร้อย" };
}

/* ================= ตรวจซ้ำ + บันทึกคะแนน =================
   Supports per-criterion scores: data.scores = {criterionName: value, ...}
   Stored row format (per-type sheet): วันที่ | กรรมการ | ชื่อ | ผลงาน | details(JSON) | total
*/
function submitScore(data){
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(data.type);
  if(!sheet){
    sheet = ss.insertSheet(data.type);
    sheet.appendRow(['วันที่','กรรมการ','ชื่อ','ผลงาน','รายละเอียดคะแนน','คะแนนรวม']);
  }
  const allData = sheet.getDataRange().getValues();
  const rows = allData.length > 1 ? allData.slice(1) : [];
  const duplicate = rows.some(r => (r[1] === data.judge) && (r[2] === data.name));
  if(duplicate){
    return "ท่านได้ให้คะแนนผู้สมัครรายนี้แล้ว ❌";
  }
  // if client computed total (weighted), prefer that
  let total = (data.total !== undefined && data.total !== null) ? Number(data.total) : 0;
  if(!(data.total !== undefined && data.total !== null)){
    // compute fallback: sum of scores (legacy)
    if(data.scores && typeof data.scores === 'object'){
      total = 0;
      Object.keys(data.scores).forEach(function(k){
        let v = Number(data.scores[k]);
        if(!isNaN(v)) total += v;
      });
    } else {
      total = Number(data.score) || 0;
    }
  }
  const details = JSON.stringify(data.scores || {});
  sheet.appendRow([ new Date(), data.judge, data.name, data.work, details, Number(total) ]);
  try { updateRanking(ss); } catch(e){ console.error('updateRanking error: ' + e); }
  return "บันทึกสำเร็จ ✅";
}

/* ================= คำนวณอันดับ =================
   Uses stored total (column 6) if present, otherwise attempts to parse JSON details.
*/
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
    data.shift(); // remove header
    const scoreMap = {};
    data.forEach(r => {
      const name = r[2];
      let score = null;
      if(r.length >= 6 && typeof r[5] === 'number') score = r[5];
      else if(typeof r[4] === 'number') score = r[4];
      else {
        try {
          const obj = JSON.parse(r[4] || '{}');
          let ssum = 0; Object.keys(obj).forEach(k => { let v = Number(obj[k]); if(!isNaN(v)) ssum += v; });
          score = ssum;
        } catch(e){
          score = NaN;
        }
      }
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

/* ================= ดึงอันดับ ================= */
function getRanking(){
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('อันดับ');
  if(!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if(data.length < 2) return [];
  data.shift();
  return data;
}

/* ================= ดึงรายการครบทุกคน (รวมผู้ยังไม่ได้รับคะแนน) ================= */
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
