function onEdit(e) { 
  Logger.log("onedit");
  
  let range = e.range;          // 從事件物件 e 中取出了被編輯的單元格範圍（Range），並將它存放在變數 range 中
  let sheet = range.getSheet(); // 使用 getSheet 方法，取得了被編輯的單元格所在的工作表（Sheet），並將它存放在變數 sheet 中。
  let row = range.getRow();     // 使用 getRow 方法，取得了被編輯的單元格的「行數」（水平），並將它存放在變數 row 中。

  let editRange = range.getA1Notation().split(":");
  
  // 根據 編輯範圍 採取對應處理方式
  if(editRange.length > 1) {
    console.log("Start: "+ editRange[0] + ",End: "+ editRange[1]);

    // 編輯 多欄位，且欄位範圍符合條件則採取 刷新日曆
    if(checkCellPlace(editRange[0].charCodeAt(0)) && checkCellPlace(editRange[1].charCodeAt(0))) {
      reflashCalendar(sheet);
    }
  } else {
    console.log("Place: "+ editRange[0].charCodeAt(0));

    // 編輯 單一欄位，且欄位範圍符合條件則採取 刷新日曆
    if(checkCellPlace(editRange[0].charCodeAt(0))) {
      let data = getRowData(sheet, row);

      if(data == null) {
        console.log("Element is empty.");
        return;
      }

      if(data.index === 'undefined' || data.index.length === 0) {
        // 新增 貼文請求
        let tarIndex = addPostEvent(data.team, data.client, data.row, data.col, data.title, data.status);
        sheet.getRange(row, 5).setValue(tarIndex);
      } else {
        // 修改 貼文請求
        if(editRange[0].charCodeAt(0) == 67) {
          reflashCalendar(sheet);
        } else {
          let tarIndex = editPostEvent(data.index, data.team, data.client, data.row, data.col, data.title, data.status);
          sheet.getRange(row, 5).setValue(tarIndex);
        }
      }
    }
  }
}

function checkCellPlace(c) {
  if((c >= 65 && c <= 68) || (c == 78))
    return true;
  else
    return false;
}

function getRowData(sheet, row){
  let team = sheet.getRange(row, 1).getValue();
  let client = sheet.getRange(row, 2).getValue();
  let postDate = sheet.getRange(row, 3).getValue();
  let title = sheet.getRange(row, 4).getValue();
  let index = sheet.getRange(row, 5).getValue();
  let status = sheet.getRange(row, 15).getValue();

  if(
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

function reflashCalendar(sheet) {
  var reqDatas = [];
  let lastRow = sheet.getLastRow();

  // 取出 貼文表單 所有資料，並完成格式前處理
  for(var i=3; i<lastRow; i++) {
    let data = getRowData(sheet, i);

    if(data == null)
      break;
    else
      reqDatas.push(data);
  }

  // 清理 貼文日曆
  if(reqDatas.length > 0)
    cleanCalender();

  // 新增 貼文事件 並更新 當日順序 欄位
  for(var i=0; i<reqDatas.length; i++) {
    let tarIndex = addPostEvent(reqDatas[i].team, reqDatas[i].client, reqDatas[i].row, reqDatas[i].col, reqDatas[i].title, reqDatas[i].color);
    sheet.getRange(i+3, 5).setValue(tarIndex);
  }

  // 清理 當日順序 欄位
  for(var i=reqDatas.length+3; i<lastRow+1; i++) {
    sheet.getRange(i, 5).setValue("");
  }
}

function cleanCalender() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('貼文日曆');
  let lastRow = sheet.getLastRow();

  var range = sheet.getRange('D4:J'+lastRow);
  range.clear();
}

function addPostEvent(team, person, row, col, title, color) {
  var sheet = SpreadsheetApp.getActive().getSheetByName('貼文日曆');
  var text  =  title + "\n@" + team + " " + person;

  // 製作 貼文日曆 事件
  text = addEventContent(text, row, col);

  // 設定 貼文日曆 事件
  var cell = sheet.getRange(row, col);
  cell.setValue(text);
  
  // 設定 貼文日曆 顏色
  if (color!="")
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
  if (color!="")
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
  for(var i=0; i < arr.length; i++) {
    if(i % 2 == 0) {
      title_arr.push(arr[i]);
    } else {
      info_arr.push(arr[i]);
      let arr2 = arr[i].split(" ");
      team_arr.push(arr2[0]);
      person_arr.push(arr2[1]);
    }
  }

  // 更新正確內容
  team_arr[index-1] = "@" + team;
  person_arr[index-1] = person;
  title_arr[index-1] = title;

  // 重設貼文日曆對應欄位內容
  var result = "";
  for(var i=0; i < title_arr.length; i++) {    
    if(result.length === 0) {
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
  let date  = Utilities.formatDate(postDate, "GMT+8", "d");

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
      real_row = row-1;
      real_col = 11 + (date - cur_date);
      break;
    } else if (+cur_date < tmp_date) {
      real_row = row-1;
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