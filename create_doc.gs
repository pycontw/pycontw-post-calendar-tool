/* Create menu */
function onOpen() {
  SpreadsheetApp
    .getUi()
    .createMenu('貼文小工具')
    .addItem('新增貼文 doc (請至貼文表單先選取要新增的文件列)', 'createPostDoc')
    .addItem('test prompt', 'showPrompt')
    .addToUi()
}

function createDocFromTemplate(docTitle, inputDict, templateId){
  var templateFile = DriveApp.getFileById(templateId);
  file = templateFile.makeCopy();
  file.setName(docTitle)
  doc_id = file.getId();

  var body = DocumentApp.openById(doc_id).getBody();
  for (const [key, value] of Object.entries(inputDict)) {
    body.replaceText("{{" + key + "}}", value);
  }

  // (special hanle) link to folder
  var folder_link_ele = body.findText(inputDict['folder_link']).getElement();
  folder_link_ele.asText().setLinkUrl(inputDict['folder_link']);

  return doc_id
}

function getRowInfo(sheet, row){
  let team = sheet.getRange(row, 1).getValue();
  let client = sheet.getRange(row, 2).getValue();
  let postDate = sheet.getRange(row, 3).getValue();
  var date = Utilities.formatDate(postDate, "GMT+8", "MMdd");
  let title = sheet.getRange(row, 4).getValue();
  let category = sheet.getRange(row, 5).getValue();
  let fb = sheet.getRange(row, 6).getValue();
  let x = sheet.getRange(row, 7).getValue();
  let ig = sheet.getRange(row, 8).getValue();
  let linkedin = sheet.getRange(row, 9).getValue();
  let other_social = sheet.getRange(row, 10).getValue();

  let pendding_text = sheet.getRange(row, 13).getValue();

  return {
    'team': team,
    'client': client,
    'date': date,
    'title': title,
    'category': category,
    'fb': fb,
    'x': x,
    'ig': ig,
    'linkedin': linkedin,
    'other_social': other_social,
    'pendding_text': pendding_text
  }
}

/* Create google doc and make sure if user would create a new folder and file */
function createPostDoc(){
  var templateId = '1sI3hSEdEpwNF2QWv3XXBLToXyl-JefLd10Tozve0Pnc';

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("貼文表單");
  var row = sheet.getActiveRange().getRow();
  Logger.log('row: %d', row);
  let info = getRowInfo(sheet, row);
  
  docTitle =  info['date'] + " - " + info['team'] + " - " + info['title'];
  var text = "是否要新增 " + docTitle + " 的 google doc 文件";

  var response = Browser.msgBox('Greetings', text, Browser.Buttons.YES_NO);
  if (response == "yes") {
    Logger.log('The user clicked "Yes."');
    var parentFolder = DriveApp.getFolderById('1mI4n-gXhVSmvDqHhzJGfPIT0C7uFEhwA');
    var target_folder_ID = parentFolder.createFolder(docTitle).getId();
    var inputDict = {
      'date': info['date'],
      'team': info['team'],
      'title': info['title'],
      'pendding_text': info['pendding_text'],
      'folder_link': 'https://drive.google.com/drive/folders/' + target_folder_ID
    };
    var doc_id = createDocFromTemplate(docTitle, inputDict, templateId);
    moveFile(doc_id, target_folder_ID);
    var url = "https://docs.google.com/document/d/" + doc_id;
    openUrl(url);
  } else {
    Logger.log('The user clicked "No" or the dialog\'s close button.');
  } 

  const rangeToAddLink = sheet.getRange(row, 11)
  const richText = SpreadsheetApp.newRichTextValue()
      .setText(docTitle)
      .setLinkUrl(url)
      .build();
  rangeToAddLink.setRichTextValue(richText);
}

function openUrl( url ){
  var html = HtmlService.createHtmlOutput('<html><script>'
  +'window.close = function(){window.setTimeout(function(){google.script.host.close()},9)};'
  +'var a = document.createElement("a"); a.href="'+url+'"; a.target="_blank";'
  +'if(document.createEvent){'
  +'  var event=document.createEvent("MouseEvents");'
  +'  if(navigator.userAgent.toLowerCase().indexOf("firefox")>-1){window.document.body.append(a)}'                          
  +'  event.initEvent("click",true,true); a.dispatchEvent(event);'
  +'}else{ a.click() }'
  +'close();'
  +'</script>'
  // Offer URL as clickable link in case above code fails.
  +'<body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="'+url+'" target="_blank" onclick="window.close()">Click here to proceed</a>.</body>'
  +'<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script>'
  +'</html>')
  .setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog( html, "Opening ..." );
}

function moveFile(fileId, destinationFolderId) {
  let destinationFolder = DriveApp.getFolderById(destinationFolderId);
  DriveApp.getFileById(fileId).moveTo(destinationFolder);
}

/* test prompt feature*/
function showPrompt() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.prompt(
      'Let\'s get to know each other!',
      'Please enter your name:',
      ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  var button = result.getSelectedButton();
  var text = result.getResponseText();
  if (button == ui.Button.OK) {
    // User clicked "OK".
    ui.alert('Your name is ' + text + '.');
  } else if (button == ui.Button.CANCEL) {
    // User clicked "Cancel".
    ui.alert('I didn\'t get your name.');
  } else if (button == ui.Button.CLOSE) {
    // User clicked X in the title bar.
    ui.alert('You closed the dialog.');
  }
}
