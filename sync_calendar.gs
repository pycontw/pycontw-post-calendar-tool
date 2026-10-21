function onEdit(e) {
  Logger.log("onedit");

  let range = e.range;          // 從事件物件 e 中取出了被編輯的單元格範圍（Range），並將它存放在變數 range 中
  let sheet = range.getSheet(); // 使用 getSheet 方法，取得了被編輯的單元格所在的工作表（Sheet），並將它存放在變數 sheet 中。
  let row = range.getRow();     // 使用 getRow 方法，取得了被編輯的單元格的「行數」（水平），並將它存放在變數 row 中。

  console.log("Sheet: " + sheet.getName());
  if (sheet.getName() != "貼文表單") {
    return;
  }

  _updateScheduledPosts(range, sheet, row);
  _insertDashboardLink(sheet, row, "成效報表");
  _insertEnhancedDashboardLink(sheet, row, "範圍內的成效");
}

function checkCellPlace(c) {
  if ((c >= 65 && c <= 68) || (c == 79))
    return true;
  else
    return false;
}

function getRawDatas(sheet, start_row, numRows) {
  let row_data = sheet.getRange(start_row, 1, numRows, 15).getValues();
  return row_data;
}

function getDocLinks(sheet, start_row, numRows) {
  let row_data = sheet.getRange(start_row, 12, numRows).getRichTextValues();
  return row_data;
}

function getRowData(row_data, row) {
  let team = row_data[0];
  let client = row_data[1];
  let postDate = row_data[2];
  let title = row_data[3];
  let index = row_data[4];
  let status = row_data[14];

  if (
    typeof title === 'undefined' || title.length === 0 ||
    typeof postDate === 'undefined' || postDate.length === 0 ||
    typeof client === 'undefined' || client.length === 0 ||
    typeof team === 'undefined' || team.length === 0
  ) {
    return null;
  }

  let calPlace = matchDate(postDate);
  let color = matchColor(status);

  return {
    'team': team,
    'client': client,
    'row': calPlace[0],
    'col': calPlace[1],
    'title': title,
    'index': index,
    'status': status,
    'color': color
  }
}

function reflashCalendar() {
  const START_ROW = 3; // 起始行數
  const POST_FORM_SHEET_NAME = '貼文表單';
  const CALENDAR_SHEET_NAME = '貼文日曆';
  const MAX_COLUMNS = 10; // 需要處理的最大列數

  // 獲取工作表
  const spreadsheet = SpreadsheetApp.getActive();
  const postFormSheet = spreadsheet.getSheetByName(POST_FORM_SHEET_NAME);
  const calendarSheet = spreadsheet.getSheetByName(CALENDAR_SHEET_NAME);
  
  const lastRowPostForm = postFormSheet.getLastRow();
  const lastRowCalendar = calendarSheet.getLastRow();

  // 取出 貼文表單 的數據
  const postFormData = getRawDatas(postFormSheet, START_ROW, lastRowPostForm - START_ROW + 1);
  const docLinks = getDocLinks(postFormSheet, START_ROW, lastRowPostForm - START_ROW + 1);

  let reqDatas = [];

  // 處理每一行數據
  for (let i = 0; i < lastRowPostForm - START_ROW + 1; i++) {
    let data = getRowData(postFormData[i], i);
    if (data) {
      let url = docLinks[i][0].getLinkUrl();
      if (url) data.doc_link = url;
      reqDatas.push(data);
    }
  }

  // 清理 貼文日曆
  if(reqDatas.length > 0)
    cleanCalender();

  // 讀取日曆工作表的現有數據
  const range = calendarSheet.getRange(1, 1, lastRowCalendar, MAX_COLUMNS);
  let calendarValues = range.getRichTextValues();
  let calendarBackgrounds = range.getBackgrounds();

  // 新增 貼文事件 並更新 當日順序 欄位
  reqDatas.forEach((data, index) => {
    let [eventIndex, richText, color] = convertReqToEvent(calendarValues, data);

    // 設置日曆中的富文本和背景色
    calendarValues[data.row - 1][data.col - 1] = richText;
    calendarBackgrounds[data.row - 1][data.col - 1] = color;

    // 更新 貼文表單 中的 當日順序
    postFormSheet.getRange(index + START_ROW, 5).setValue(eventIndex);
  });

  // 清理 貼文表單 中的 當日順序 欄位
  for (let i = reqDatas.length + START_ROW; i <= lastRowPostForm; i++) {
    postFormSheet.getRange(i, 5).setValue("");
  }

  const updateRange = calendarSheet.getRange(1, 1, lastRowCalendar, 3);
  const displayValues = updateRange.getDisplayValues();

  
  displayValues.forEach((row, rowIndex) => {
    row.forEach((value, colIndex) => {
      let richText = SpreadsheetApp.newRichTextValue().setText(value).build();
      calendarValues[rowIndex][colIndex] = richText;
    });
  });

  try {
    // 將更新的數據批量寫回
    range.setRichTextValues(calendarValues);
    range.setBackgrounds(calendarBackgrounds);
  } catch (error) {
    console.log(error);
  }
}

function cleanCalender() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('貼文日曆');
  let lastRow = sheet.getLastRow();

  var range = sheet.getRange('D4:J'+lastRow);
  range.clear();
}

function convertReqToEvent(calenderArray, event) {
  // team, person, row, col, title, color, url
  const team = event.team;
  const person = event.client;
  const row = event.row;
  const col = event.col;
  const title = event.title;
  const color = event.color;
  const url = event.doc_link;
  
  var text  =  title + "\n@" + team + " " + person;

  // 製作 貼文日曆 事件
  text = addEventContent(text, row, col, calenderArray);

  var richText = SpreadsheetApp.newRichTextValue()
      .setText(text)
      .setLinkUrl(url)
      .build();

  var count = getCountEventOfDay(text);

  return [count, richText, color]
}

function addEventContent(value, row, col, calenderArray) {
  var cur_data = calenderArray[row-1][col-1].getText();

  if (cur_data.length > 0)
    cur_data = cur_data + "\n" + value;
  else {
    cur_data = value;
  }

  return cur_data;
}

function matchDate(postDate) {
  let sheet = SpreadsheetApp.getActive().getSheetByName('貼文日曆');

  let month = Utilities.formatDate(postDate, "GMT+8", "M");
  let date = Utilities.formatDate(postDate, "GMT+8", "d");

  var row = 1;
  var real_row, real_col = 0;

  for (row = 1; row < 200; row++) {
    if (+sheet.getRange(row, 2).getValue() == month) {
      break; 
    }
  }

  row_limit = row + 5;
  var cur_date, tmp_date;
  while (row < row_limit) {
    cur_date = sheet.getRange(row, 3).getValue();
    if (+cur_date > date) {
      real_row = row - 1;
      real_col = 11 + (date - cur_date);
      break;
    } else if (+cur_date < tmp_date) {
      real_row = row - 1;
      real_col = 4 + (date - tmp_date);
      break;
    }
    row += 1;
    tmp_date = cur_date;
  }

  return [real_row, real_col];
}

function matchColor(status) {
  if (status == "已審閱")
    color = "#a4c2f4";
  else if (status == "待審閱")
    color = "#ffcfc8";
  else if (status == "已排程")
    color = "#b7d7a8";
  else if (status == "已發布")
    color = "#cccccc";
  else 
    color = "transparent";
  return color;
}

function getCountEventOfDay(cur_data) {
  let count = cur_data.split('').filter(char => char === '@').length;
  return count;
}

const _insertDashboardLink = (sheet, editedRow, editedValue) => {
  // Specify the column number where you want to add the link
  const linkColumn = 18; // For example, column G

  // Get the cell to which you want to add the link
  const linkCell = sheet.getRange(editedRow, linkColumn);
  const postDate = sheet.getRange(editedRow, 3).getValue();
  const year = postDate.getFullYear();
  const month = String(postDate.getMonth() + 1).padStart(2, '0'); // Months are zero-indexed
  const day = String(postDate.getDate()).padStart(2, '0');

  const fb = sheet.getRange(editedRow, 7).getValue();
  const x = sheet.getRange(editedRow, 8).getValue();
  const ig = sheet.getRange(editedRow, 9).getValue();
  const linkedin = sheet.getRange(editedRow, 10).getValue();
  const platform = x ? 'x' : '';
  // Construct the link URL based on the edited value
  // it will be replaced with the actual logic to construct the link URL
  const linkUrl = `https://metabase.pycon.tw/question/214-social-media-marketing-metrics?date=${year}-${month}-${day}&platform=${platform}`

  // Create a rich text value with the link
  const richTextValue = SpreadsheetApp.newRichTextValue()
    .setText(editedValue)
    .setLinkUrl(linkUrl)
    .build();

  // Set the rich text value to the link cell
  linkCell.setRichTextValue(richTextValue);
}

const _insertEnhancedDashboardLink = (sheet, editedRow, editedValue) => {
  const linkColumn = 19;
  const linkCell = sheet.getRange(editedRow, linkColumn);
  const linkUrl = "https://metabase.pycon.tw/question/215?date1=2024-06-01&date2=2024-06-20"
  const richTextValue = SpreadsheetApp.newRichTextValue()
    .setText(editedValue)
    .setLinkUrl(linkUrl)
    .build();
  linkCell.setRichTextValue(richTextValue);
}

const _updateScheduledPosts = (range, sheet, row) => {
  let editRange = range.getA1Notation().split(":");

  // 根據 編輯範圍 採取對應處理方式
  if (editRange.length > 1) {
    console.log("Start: " + editRange[0] + ",End: " + editRange[1]);
    // 編輯 多欄位，且欄位範圍符合條件則採取 刷新日曆
    reflashCalendar();
  } else {
    console.log("Place: " + editRange[0].charCodeAt(0));
    // 編輯 單一欄位，且欄位範圍符合條件則採取 刷新日曆
    if (checkCellPlace(editRange[0].charCodeAt(0))) {
      reflashCalendar();
    } else {
      console.log("Edit is outside of the target range.");
    }
  }
}