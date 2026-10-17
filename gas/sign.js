/** 最新簽名報表：編輯紀錄顯示修改者簽名，不冒充老師重新確認。 */
function renderClubReportDoc() {
  var ui=SpreadsheetApp.getUi();
  var response=ui.prompt('產生簽名報表','請輸入社課日期（yyyy/MM/dd）',ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton()!==ui.Button.OK) return;
  var target=normalizeDate(response.getResponseText().trim());
  if (!target) { ui.alert('請輸入有效日期');return; }
  var ss=SpreadsheetApp.getActiveSpreadsheet(), summary=ss.getSheetByName('統計總表');
  if (!summary) throw new Error('找不到統計總表');
  var data=summary.getDataRange().getValues();
  var idCol=summaryHeaderColumn(data[0],['社團編號','社團代碼']);
  var clubCol=summaryHeaderColumn(data[0],['社團名稱','社團']);
  var teacherCol=summaryHeaderColumn(data[0],['指導老師姓名','指導老師','老師姓名']);
  var clubs=Object.create(null), latest=Object.create(null);
  data.slice(1).forEach(function(row) {
    if (row[clubCol]) clubs[String(row[clubCol])]={id:String(row[idCol]),teacher:String(row[teacherCol] || '未填寫')};
  });
  readAttendanceRecords(true).forEach(function(record) {
    if (normalizeDate(record.date)!==target || !record.club) return;
    var old=latest[record.club];
    if (!old || record._ts>old._ts || (record._ts===old._ts && record._row>old._row)) latest[record.club]=record;
  });
  var names=Object.keys(latest).sort(function(a,b) {
    return String(clubs[a] && clubs[a].id || '').localeCompare(String(clubs[b] && clubs[b].id || ''));
  });
  if (!names.length) { ui.alert('找不到此日期的點名紀錄');return; }
  var doc=DocumentApp.create('社團點名簽名報表 - '+target.replace(/\//g,'-'));
  var body=doc.getBody();
  body.appendParagraph('社團點名簽名報表').setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph('社課日期：'+target);
  body.appendParagraph('編輯紀錄所附為修改者簽名，不代表原指導老師重新確認修改內容。');
  var table=body.appendTable([['社團編號','社團名稱','指導老師','紀錄類別','簽名角色／姓名','簽名圖片']]);
  names.forEach(function(name) {
    var record=latest[name], info=clubs[name] || {id:'未對應',teacher:'未對應'};
    var row=table.appendTableRow();
    [info.id,name,info.teacher,record.recordType || '歷史未分類',record.signatureRole+(record.recordType==='編輯' ? '／'+String(record.fillerName || '') : '')].forEach(function(text) { row.appendTableCell(String(text)); });
    var cell=row.appendTableCell(), id=signatureFileId(record.signatureUrl);
    if (!id) { cell.setText('無可讀取簽名');return; }
    try {
      var file=DriveApp.getFileById(id), parents=file.getParents(), inFolder=false;
      var folderId=getSignFolderId();
      while (parents.hasNext()) if (parents.next().getId()===folderId) inFolder=true;
      if (!inFolder || file.getMimeType()!=='image/png' || file.getSize()>150000) throw new Error('不允許的簽名檔案');
      var image=cell.appendImage(file.getBlob());
      var ratio=Math.min(120/image.getWidth(),60/image.getHeight(),1);
      image.setWidth(Math.max(1,Math.round(image.getWidth()*ratio)));
      image.setHeight(Math.max(1,Math.round(image.getHeight()*ratio)));
    } catch (error) { cell.setText('無法載入簽名（權限、格式或資料夾不符）'); }
  });
  doc.saveAndClose();
  var reportFolderId=getReportFolderId();
  if (reportFolderId) DriveApp.getFileById(doc.getId()).moveTo(DriveApp.getFolderById(reportFolderId));
  // 不刪除同名舊報表，保留歷史稽核資料。
  ui.alert('報表已建立：\n'+doc.getUrl());
}

function normalizeDate(value) {
  if (value instanceof Date) return formatDateLocal(value).replace(/-/g,'/');
  var match=String(value || '').trim().match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
  if (!match) return '';
  var date=new Date(Number(match[1]),Number(match[2])-1,Number(match[3]));
  if (date.getFullYear()!==Number(match[1]) || date.getMonth()+1!==Number(match[2]) || date.getDate()!==Number(match[3])) return '';
  return formatDateLocal(date).replace(/-/g,'/');
}
