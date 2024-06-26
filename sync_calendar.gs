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
  let startTime = new Date();
  var reqDatas = [];
  let sheet = SpreadsheetApp.getActive().getSheetByName('貼文表單');
  let lastRow = sheet.getLastRow();

  // 取出 貼文表單 所有資料，並完成格式前處理
  let row_data = getRawDatas(sheet, 3, lastRow - 3 + 1);
  let doc_links = getDocLinks(sheet, 3, lastRow - 3 + 1);

  for (var i = 0; i < lastRow - 2; i++) {
    let data = getRowData(row_data[i], i);
    let url = doc_links[i][0].getLinkUrl();
    if (url != null)
      data.doc_link = url;

    if (data == null)
      break;
    else
      reqDatas.push(data);
  }

  // 清理 貼文日曆
  if (reqDatas.length > 0)
    cleanCalender();

  // 新增 貼文事件 並更新 當日順序 欄位
  for (var i = 0; i < reqDatas.length; i++) {
    if (reqDatas[i].row <= 26) {
      continue;
    }

    let tarIndex = addPostEvent(reqDatas[i].team, reqDatas[i].client, reqDatas[i].row, reqDatas[i].col, reqDatas[i].title, reqDatas[i].color, reqDatas[i].doc_link);
    sheet.getRange(i + 3, 5).setValue(tarIndex);

    if (isTimeOut(startTime)) {
      console.log("Time Out");
      return;
    }
  }

  // 清理 當日順序 欄位
  for (var i = reqDatas.length + 3; i < lastRow + 1; i++) {
    sheet.getRange(i, 5).setValue("");
  }
}

function cleanCalender() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('貼文日曆');
  let lastRow = sheet.getLastRow();

  var range = sheet.getRange('D27:J' + lastRow);
  range.clear();
}

function addPostEvent(team, person, row, col, title, color, url) {
  var sheet = SpreadsheetApp.getActive().getSheetByName('貼文日曆');
  var text = title + "\n@" + team + " " + person;

  // 製作 貼文日曆 事件
  text = addEventContent(text, row, col);

  const richText = SpreadsheetApp.newRichTextValue()
    .setText(text)
    .setLinkUrl(url)
    .build();

  // 設定 貼文日曆 事件
  var cell = sheet.getRange(row, col);
  cell.setRichTextValue(richText);
  // cell.setValue(text);

  // 設定 貼文日曆 顏色
  if (color != "")
    cell.setBackground(color);

  return getCountEventOfDay(text);
}

function editPostEvent(index, team, person, row, col, title, color) {
  var sheet = SpreadsheetApp.getActive().getSheetByName('貼文日曆');

  // 製作 貼文日曆 事件
  let [text, newIndex] = modifyEvenContent(index, team, person, title, row, col);

  // 設定 貼文日曆 事件
  var cell = sheet.getRange(row, col);
  cell.setValue(text);

  // 設定 貼文日曆 顏色
  console.log(color);
  if (color != "")
    cell.setBackground(color);

  return newIndex;
}

function addEventContent(value, row, col) {
  let sheet = SpreadsheetApp.getActive().getSheetByName('貼文日曆');
  let cell = sheet.getRange(row, col);

  let cur_data = cell.getValue();

  if (cur_data.length > 0)
    cur_data = cur_data + "\n" + value;
  else {
    cur_data = value;
  }

  return cur_data;
}

function modifyEvenContent(index, team, person, title, row, col) {
  let sheet = SpreadsheetApp.getActive().getSheetByName('貼文日曆');
  let old_cell = sheet.getRange(row, col);
  let data = old_cell.getValue();

  let arr = data.split("\n");
  var title_arr = [];
  var info_arr = [];
  var team_arr = [];
  var person_arr = [];

  // 分類欄位內各項內容
  for (var i = 0; i < arr.length; i++) {
    if (i % 2 == 0) {
      title_arr.push(arr[i]);
    } else {
      info_arr.push(arr[i]);
      let arr2 = arr[i].split(" ");
      team_arr.push(arr2[0]);
      person_arr.push(arr2[1]);
    }
  }

  // 更新正確內容
  team_arr[index - 1] = "@" + team;
  person_arr[index - 1] = person;
  title_arr[index - 1] = title;

  // 重設貼文日曆對應欄位內容
  var result = "";
  for (var i = 0; i < title_arr.length; i++) {
    if (result.length === 0) {
      result = title_arr[i] + "\n" + team_arr[i] + " " + person_arr[i];
    } else {
      result = result + "\n" + title_arr[i] + "\n" + team_arr[i] + " " + person_arr[i];
    }
  }

  old_cell.setValue(result);
  return [result, index];
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

function isTimeOut(today) {
  var now = new Date();
  return now.getTime() - today.getTime() > 28000;
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

const _updateScheduledPosts = (range, sheet, row) => {
  let startTime = new Date();
  let editRange = range.getA1Notation().split(":");

  var editState = sheet.getRange(1, 23).getValue();
  while (editState == "更新中") {
    editState = sheet.getRange(1, 23).getValue();

    if (isTimeOut(startTime)) {
      console.log("Time Out");
      sheet.getRange(1, 23).setValue("--");
      return;
    }
  }

  // 根據 編輯範圍 採取對應處理方式
  if (editRange.length > 1) {
    console.log("Start: " + editRange[0] + ",End: " + editRange[1]);

    // 編輯 多欄位，且欄位範圍符合條件則採取 刷新日曆
    sheet.getRange(1, 23).setValue("更新中");
    reflashCalendar();
    sheet.getRange(1, 23).setValue("--");
  } else {
    console.log("Place: " + editRange[0].charCodeAt(0));

    // 編輯 單一欄位，且欄位範圍符合條件則採取 刷新日曆
    if (checkCellPlace(editRange[0].charCodeAt(0))) {
      sheet.getRange(1, 23).setValue("更新中");
      let row_data = getRawDatas(sheet, row, 1);
      let doc_links = getDocLinks(sheet, row, 1);
      let data = getRowData(row_data[0], row);
      data['doc_link'] = doc_links[0][0].getLinkUrl();

      if (data == null) {
        console.log("Element is empty.");
        sheet.getRange(1, 23).setValue("--");
        return;
      }

      console.log(data.index);
      if (data.index === 'undefined' || data.index.length === 0) {
        // 新增 貼文請求
        let tarIndex = addPostEvent(data.team, data.client, data.row, data.col, data.title, data.color, data.doc_link);
        sheet.getRange(row, 5).setValue(tarIndex);
      } else {
        // 修改 貼文請求
        if (editRange[0].charCodeAt(0) == 67) {
          reflashCalendar();
        } else {
          let tarIndex = editPostEvent(data.index, data.team, data.client, data.row, data.col, data.title, data.color);
          sheet.getRange(row, 5).setValue(tarIndex);
        }
      }
      sheet.getRange(1, 23).setValue("--");
    }
  }
}