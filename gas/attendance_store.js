/** 點名紀錄共用存取。由呼叫端持有 ScriptLock 後進行寫入。 */
var ATTENDANCE_FIELDS = [
  ['timestamp','填寫日期時間','時間戳記','填寫時間','填答時間'],
  ['fillerInfo','填答人學號資訊','填寫人資訊','填答人資訊'],
  ['club','社團名稱','社團官方名稱','社團'],
  ['date','點名社課日期','社課日期','點名日期'],
  ['fillerName','填寫人姓名','填寫人','填答人姓名'],
  ['desc','社課內容簡述','社課內容'],
  ['teacherPresent','指導老師出席與否','指導老師是否出席','老師出席'],
  ['subTeacher','代課老師姓名','代課老師'],
  ['presentCount','出席總人數','出席人數','實到人數'],
  ['absentList','缺席名單','換行缺席名單'],
  ['signatureUrl','老師簽名'], ['recordType','紀錄類別']
];

function attendanceHeaderMap(headers, allowMissingType) {
  var names=Object.create(null), map=Object.create(null);
  headers.forEach(function(value,index) {
    var name=String(value || '').trim();
    if (!name) return;
    if (names[name] !== undefined) throw new Error('重複欄名：'+name);
    names[name]=index;
  });
  ATTENDANCE_FIELDS.forEach(function(field) {
    var matches=field.slice(1).filter(function(name) { return names[name] !== undefined; });
    if (matches.length > 1) throw new Error('欄位語意重複：'+field[1]);
    if (!matches.length) {
      if (field[0] === 'recordType' && allowMissingType) return;
      throw new Error('缺少必要欄位：'+field[1]+'；請先檢查欄名，不會猜測欄位位置');
    }
    map[field[0]]=names[matches[0]];
  });
  return map;
}

function getAttendanceSchema(addType) {
  var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('點名紀錄');
  if (!sheet) throw new Error('找不到點名紀錄工作表');
  var width=sheet.getLastColumn();
  var headers=width ? sheet.getRange(1,1,1,width).getValues()[0] : [];
  var map=attendanceHeaderMap(headers,true);
  if (addType && map.recordType === undefined) {
    if (width === sheet.getMaxColumns()) sheet.insertColumnsAfter(width,1);
    sheet.getRange(1,width+1).setValue('紀錄類別');
    headers.push('紀錄類別');
    map.recordType=width++;
  }
  return {sheet:sheet,headers:headers,map:map,width:width};
}

function initializeAttendanceSchema() {
  var lock=LockService.getScriptLock();
  lock.waitLock(10000);
  try { getAttendanceSchema(true); } finally { lock.releaseLock(); }
}

function attendanceColumnName(index) {
  var text='';
  for (var n=index+1;n>0;n=Math.floor((n-1)/26)) text=String.fromCharCode(65+(n-1)%26)+text;
  return text;
}

function signatureFileId(value) {
  if (typeof value !== 'string') return null;
  var text=value.trim();
  if (/^[A-Za-z0-9_-]{10,200}$/.test(text)) return text;
  var match=text.match(/^https:\/\/drive\.google\.com\/file\/d\/([A-Za-z0-9_-]{10,200})(?:[/?#][^\s]*)?$/);
  if (!match) match=text.match(/^https:\/\/drive\.google\.com\/(?:open|uc)\?(?:[^\s#]*&)?id=([A-Za-z0-9_-]{10,200})(?:[&#][^\s]*)?$/);
  return match ? match[1] : null;
}

function signatureUrl(value) {
  var id=signatureFileId(value);
  return id ? 'https://drive.google.com/file/d/'+id+'/view' : '';
}

function signatureChip(value) {
  var url=signatureUrl(value);
  if (!url) throw new Error('簽名檔案 ID 格式不正確');
  return {userEnteredValue:{stringValue:'@'},chipRuns:[{startIndex:0,chip:{richLinkProperties:{uri:url}}}]};
}

function readAttendanceRecords(includeSignatures) {
  var schema=getAttendanceSchema(false), data=schema.sheet.getDataRange().getValues();
  var urls=Object.create(null);
  if (includeSignatures && data.length>1) {
    var col=attendanceColumnName(schema.map.signatureUrl);
    var response=Sheets.Spreadsheets.get(SpreadsheetApp.getActiveSpreadsheet().getId(),{
      ranges:["'點名紀錄'!"+col+'2:'+col+data.length],
      fields:'sheets(data(startRow,rowData(values(chipRuns))))'
    });
    (response.sheets || []).forEach(function(sheet) {
      (sheet.data || []).forEach(function(grid) {
        (grid.rowData || []).forEach(function(row,index) {
          var runs=row.values && row.values[0] && row.values[0].chipRuns || [];
          var links=runs.filter(function(run) { return run.chip && run.chip.richLinkProperties; });
          if (links.length===1) urls[(grid.startRow || 0)+index]=signatureUrl(links[0].chip.richLinkProperties.uri);
        });
      });
    });
  }
  return data.slice(1).map(function(row,index) {
    var record={_row:index+2};
    Object.keys(schema.map).forEach(function(key) { record[key]=row[schema.map[key]]; });
    if (includeSignatures) record.signatureUrl=urls[index+1] || signatureUrl(record.signatureUrl);
    record._ts=record.timestamp instanceof Date ? record.timestamp.getTime() : 0;
    record.date=record.date instanceof Date ? formatDateLocal(record.date) : String(record.date || '');
    record.signatureRole=record.recordType==='編輯' ? '修改者簽名' : record.recordType==='新增' ? '老師簽名' : '簽名（歷史角色未分類）';
    return record;
  });
}

function selectLatestAttendance(records) {
  var latest=Object.create(null);
  records.forEach(function(record) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(record.date)) return;
    var old=latest[record.date];
    if (!old || record._ts>old._ts || (record._ts===old._ts && record._row>old._row)) latest[record.date]=record;
  });
  return latest;
}

function validateSignaturePng(value) {
  if (typeof value!=='string' || !value || value.length>200000) throw new Error('請提供本次有效簽名，且不得超過大小限制');
  var prefix='data:image/png;base64,';
  if (value.indexOf(prefix)!==0) throw new Error('簽名必須為 PNG，不可沿用簽名或提供網址');
  var encoded=value.slice(prefix.length);
  if (!encoded || !/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(encoded)) throw new Error('簽名 Base64 格式不正確');
  var bytes=Utilities.base64Decode(encoded);
  if (bytes.length<57 || bytes.length>150000 || Utilities.base64Encode(bytes)!==encoded) throw new Error('簽名圖片大小或編碼不正確');
  var data=bytes.map(function(b) { return b & 255; });
  if (![137,80,78,71,13,10,26,10].every(function(b,i) { return data[i]===b; })) throw new Error('簽名不是 PNG');
  function uint32(at) { return data[at]*16777216+data[at+1]*65536+data[at+2]*256+data[at+3]; }
  var pos=8, chunks=0, idat=false, finished=false;
  while (pos+12<=data.length) {
    var length=uint32(pos);
    if (length>data.length-pos-12) throw new Error('PNG 區塊長度不正確');
    var kind=String.fromCharCode.apply(null,data.slice(pos+4,pos+8));
    if (!chunks) {
      if (kind!=='IHDR' || length!==13) throw new Error('PNG 缺少 IHDR');
      var width=uint32(pos+8), height=uint32(pos+12);
      if (!width || !height || width>4096 || height>4096 || width*height>4194304) throw new Error('簽名尺寸超出限制');
      if (data[pos+16]!==8 || [2,6].indexOf(data[pos+17])<0 || data[pos+18]!==0 || data[pos+19]!==0 || data[pos+20]!==0) throw new Error('請使用簽名板重新產生 PNG');
    } else if (kind==='IHDR') throw new Error('PNG 重複 IHDR');
    var crc=-1;
    for (var i=pos+4;i<pos+8+length;i++) {
      crc^=data[i];
      for (var bit=0;bit<8;bit++) crc=(crc>>>1)^((crc&1) ? 0xEDB88320 : 0);
    }
    if (((crc^-1)>>>0)!==uint32(pos+8+length)) throw new Error('PNG 檢查碼不正確');
    if (kind==='IDAT' && length>0) idat=true;
    pos+=length+12;chunks++;
    if (kind==='IEND') { finished=length===0 && pos===data.length;break; }
  }
  if (!idat || !finished) throw new Error('PNG 圖片不完整');
  return bytes;
}

function attendanceCell(value) {
  // 專案時區固定 Asia/Taipei；Sheets 日期序號為當地牆上時間。
  if (value instanceof Date) return {userEnteredValue:{numberValue:(value.getTime()+8*3600000)/86400000+25569},userEnteredFormat:{numberFormat:{type:'DATE_TIME',pattern:'yyyy/mm/dd hh:mm:ss'}}};
  return {userEnteredValue:typeof value==='number' ? {numberValue:value} : {stringValue:String(value===undefined ? '' : value)}};
}

function appendAttendanceRecord(record,bytes) {
  var schema=getAttendanceSchema(true), sheet=schema.sheet, rowIndex=sheet.getLastRow();
  var folderId=getSignFolderId();
  if (!folderId) throw new Error('尚未設定簽名資料夾');
  var role=record.recordType==='編輯' ? '修改者簽名' : '老師簽名';
  var name=formatDateLocal(record.date)+'_'+record.club+'_'+role+'_'+Utilities.getUuid()+'.png';
  var file=DriveApp.getFolderById(folderId).createFile(Utilities.newBlob(bytes,'image/png',name));
  var sent=false;
  try {
    var values=schema.headers.map(function() { return attendanceCell(''); });
    Object.keys(schema.map).forEach(function(key) { values[schema.map[key]]=key==='signatureUrl' ? signatureChip(file.getId()) : attendanceCell(record[key]); });
    var requests=[];
    if (rowIndex>=sheet.getMaxRows()) requests.push({appendDimension:{sheetId:sheet.getSheetId(),dimension:'ROWS',length:1}});
    requests.push({updateCells:{start:{sheetId:sheet.getSheetId(),rowIndex:rowIndex,columnIndex:0},rows:[{values:values}],fields:'userEnteredValue,userEnteredFormat.numberFormat,chipRuns'}});
    sent=true;
    Sheets.Spreadsheets.batchUpdate({requests:requests},SpreadsheetApp.getActiveSpreadsheet().getId());
  } catch (error) {
    if (!sent) { file.setTrashed(true);throw error; }
    // Timeout 可能已成功寫入。確認檔案 ID 後才清理，不刪除已提交簽名。
    try {
      var committed=readAttendanceRecords(true).some(function(r) { return signatureFileId(r.signatureUrl)===file.getId(); });
      if (committed) return;
      // 尚未讀到不代表遠端永遠不會完成；保留檔案，避免延遲提交成為失效連結。
      throw new Error('提交結果待確認');
    } catch (reconcileError) {
      throw new Error('無法確認提交結果，請先查詢歷史紀錄，勿立即重複提交');
    }
    throw new Error('簽名與紀錄寫入失敗，請稍後重試');
  }
}

function attendanceLegacyRows(includeSignatures) {
  var rows=[ATTENDANCE_FIELDS.map(function(f) { return f[1]; })];
  readAttendanceRecords(includeSignatures).forEach(function(record) {
    rows.push(ATTENDANCE_FIELDS.map(function(f) { return f[0]==='date' ? new Date(record.date.replace(/-/g,'/')) : record[f[0]]; }));
  });
  return rows;
}

function summaryHeaderColumn(headers,aliases) {
  var found=[];
  headers.forEach(function(name,i) { if (aliases.indexOf(String(name).trim())>=0) found.push(i); });
  if (found.length!==1) throw new Error('統計總表缺少或重複欄位：'+aliases[0]);
  return found[0];
}
