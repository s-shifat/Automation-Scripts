let PRESENTATION_ID;
let BAN_FONT;
let ENG_FONT;
let OUTPUT_FOLDER_ID;
let OUTPUT_FOLDER;
let SUBJECT;
let STARTING_OF_OUTPUT_COLUMN;
let PIC_CONTATINING_COLUMN_NAME;
let EMAIL_TEMPLATE_ID;
let EMAIL_COLUMN_NAME;
const WIDTH = 5000;

/**
 * Log every usefull detatils of each element from a presentation.
 * Usefull for going through a presentaion real fast.
 * 
 * @param {string} PRESENTATION_ID ID of template presentation
*/
function seeSlideElements(PRESENTATION_ID){
  var presentation = SlidesApp.openById(PRESENTATION_ID);
  var slides = presentation.getSlides();
  slides.map(slide => slide.getPageElements().map(element => Logger.log("slide: "+slide.getObjectId(),"Element Name: "+element.getTitle(), "Type: "+element.getPageElementType())));
  presentation.saveAndClose();
}

/**
  * Downloads a slide of a presentation as jpeg to a specified folder in google drive.
  *
  * @param {String} name The name you want to give to the exported image
  * @param {String} presentationId ID of target presentation
  * @param {String} slideId ID of the slide to be exported
  * @param {Folder} folder DriveApp's Folder type object
  * @return {File}  DriveApp's file type object
*/
function downloadOneSlide(name, presentationId, slideId,folder) {
  var url = 'https://docs.google.com/presentation/d/' + presentationId +
    '/export/png?id=' + presentationId + '&pageid=' + slideId; 
  var options = {
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    }
  };
  var response = UrlFetchApp.fetch(url, options);
  var image = response.getAs(MimeType.PNG);
  image.setName(name);
  var file = folder.createFile(image);
  return file
}

/**
  * Downloads every slide of a presentation as jpeg to a specified folder in the google drive.
  * 
  * @param {String} presentationId ID of target presentation
  * @param {Array} slideIds An array slide Ids of the presentation to be exported, `replaceImages(raw,presentationId, detecter)` function might come in handy here.
  * @param {Folder} folder DriveApp's Folder type object
  * @return {File}  DriveApp's file type object
*/
function downloadAllSlides(presentationId, slideIds,folderId,destroy=false){
  let outputFolder = DriveApp.getFolderById(folderId);
  let name;
  let niceDate = dateToNiceString(new Date);
  let downloadedFiles = [];
  slideIds.forEach(slideId => {
    name = slideId+'-s-'+niceDate+'.jpg';
    downloadedFiles.push(
      downloadOneSlide(name, presentationId, slideId,outputFolder));
  });
  DriveApp.getFileById(presentationId).setTrashed(destroy);
  return downloadedFiles;
}

/**
  * Creates a copy of main template presentation.
  * 
  * @param {String} presentationId ID of template presentation
  * @return {String} ID of a copy of given presentation
*/
function makeTemplatePresentation(presentationId){
  return DriveApp.getFileById(presentationId).makeCopy('temp').getId();
}

/**
  * Extracts data from a spreadsheet in a convenient format.
  * Returns an object having two attributes.---
  * Object.head --> Array containing name of the columns
  * Object.data --> Array containing values of the column names, from last row of spreadsheet
  *
  * @param {String} ssId ID of the spreadsheet
  * @sheetName {String} name of the sheet to be considered
  * @param {Boolean} latest if you only need the latest value of data range then true. If false, then reaturns every data, except first row
  * @return {Object}  An object containing two properties.
*/
function extractData(ssId, sheetName ,latest=true){
  let spreadsheet = SpreadsheetApp.openById(ssId);
  let sheet;
  try{
    sheet = spreadsheet.getSheetByName(sheetName);
    
  } catch (e) {
    Logger.log("Could Not find the sheet Name specified. Taking Active Sheet...");
    sheet = spreadsheet.getActiveSheet();
  }
  
  let [head, ...data] = sheet.getDataRange().getDisplayValues();
  if (latest == true){
    data = data[data.length-1];
  }
  
  return {
    head: head,
    data: data,
  }  
}

/**
  * This one will be one of your favourites. This function replaces the texts in the defined variables of
  * template presentation with respective column values of spreadsheet.
  * The defined variables are the name of the spreadsheet's column.
  * In the template presentation the varibale format shoulb be like
  * {{column name}}
  * in a text box.
  *
  * @param {Object} raw the object returned by function `extractData(ssId)`
  * @param {String} presentationId ID of template presentation
  * @return {String}  Returns the Id of a presentation that has all the variables replaced
*/
function replaceTheTexts(raw, presentationId){
  let templateId = makeTemplatePresentation(presentationId);
  let presentation = SlidesApp.openById(templateId);
  let tobeRaplaced = '';
  let i=0;
  for (i=0; i<raw.head.length; i++){
    tobeRaplaced = '{{' + raw.head[i] + '}}';
    presentation.replaceAllText(tobeRaplaced, raw.data[i]);
  }
  presentation.saveAndClose();
  return templateId;
}

/**
  * This one will be one of your favourites. This funciton reads in the text of a 
  * text box or shape, If the text is a valid link of google drive photo, then
  * renders the photo IE, converts the shape to image while going through each slide of given presentation
  * In the template presentation the varibale format shoulb be like
  * {{column name}}
  * in a text box.
  * It is recommended that you should run the `replaceTheTexts(raw, presentationId)` function first,
  * before using this function.
  *
  * @param {Object} raw the object returned by function `extractData(ssId)`
  * @param {String} presentationId ID of target presentation
  * @return {Array}  Returns an array of slide Ids of all the slides of the processed presentation
*/
function replaceImages(raw,presentationId, detecter){
  let presentation = SlidesApp.openById(presentationId);
  let slides = presentation.getSlides();
  // if code blows up look here. check if the column name of pic is correct
  let idx = raw.head.indexOf(detecter);
  //'{{' + raw.head[idx] + '}}';
  let sheetUrls = raw.data[idx];
  // let slideIds = [];
  let blobsrc,crop, slideUrls, imageUrl, leftDelta, topDelta;
  leftDelta = topDelta = 0; 
  slides.forEach(slide => {
    slide.getShapes().forEach(shape => {
      crop = slides.indexOf(slide) != slides.length-1;
      slideUrls = shape.getText().asString().trim();
      if (slideUrls.startsWith('https://') && slideUrls == sheetUrls){
        slideUrls.split(',').forEach(slideUrl => {
          imageUrl = slideUrl.trim();
          _processOneShape(slide, shape, imageUrl, leftDelta, topDelta);
          leftDelta += 10;
          topDelta += 10;
        });
        shape.getText().clear();
        shape.remove();
      }
    });
      
  });
  
  presentation.saveAndClose();
  return presentation.getId(); 
}

function _processOneShape(slide, motherShape, link, leftDelta=0, topDelta=0){
  let shape = motherShape.duplicate().asShape();
  shape.getText().clear();
  let left = motherShape.getLeft();
  let top = motherShape.getTop();
  let blobsrc,imgShape;
  try{
    blobsrc = DriveApp.getFileById(extracId(link));
    imgShape = shape.replaceWithImage(blobsrc);
    imgShape.setTop(top - topDelta);
    imgShape.setTop(left - leftDelta);
    } catch(e){
        try {
          blobsrc = resizing(extracId(imageUrl), WIDTH);
          imgShape = shape.replaceWithImage(blobsrc);
          imgShape.setTop(top - topDelta);
          imgShape.setTop(left - leftDelta);
        } catch (e) {
          Logger.log(slide.getObjectId());
        } 
  }
}

/**
  * Total work flow. Takes the latest row of a spreadsheet, then generates a presentation, then downloads it. Also sends email.
  * You may use global constants as the arguments.
  *
  * @param {String} PRESENTATION_ID ID of target presentation
  * @param {String} PIC_CONTATINING_COLUMN_NAME the name of the spreadsheet column containing images
  * @param {String} OUTPUT_FOLDER_ID ID of output folder
  * @param {String} SUBJECT ID subject of email
  * @param {String} EMAIL_COLUMN_NAME name of the column holding the emails in the spreadsheet
  * @param {String} EMAIL_TEMPLATE_ID ID of the email template doc file
  * @return {Array} Returns an array of slide Ids of all the slides of the processed presentation
*/
function automation(PRESENTATION_ID,PIC_CONTATINING_COLUMN_NAME,OUTPUT_FOLDER_ID,SUBJECT,EMAIL_COLUMN_NAME,EMAIL_TEMPLATE_ID){
  let spreadsheet = SpreadsheetApp.getActive();
  let ssId = spreadsheet.getId();
  let raw = extractData(ssId);
  let presentationId = replaceTheTexts(raw,PRESENTATION_ID);
  let slideIds;
  if (PIC_CONTATINING_COLUMN_NAME != ''){
    slideIds = replaceImages(raw,presentationId, PIC_CONTATINING_COLUMN_NAME);
  }else{
    slideIds = [];
    let presentation = SlidesApp.openById(presentationId);
    let slides = presentation.getSlides();
    slides.forEach(slide => slideIds.push(slide.getObjectId()));
    presentation.saveAndClose();
  }
  let downloadedFiles = downloadAllSlides(presentationId, slideIds,OUTPUT_FOLDER_ID,true);
  let sheet = spreadsheet.getActiveSheet();
  generateOutputLinksToSheet(sheet,downloadedFiles);
  let html = prepareEmail(ssId,downloadedFiles,EMAIL_TEMPLATE_ID);
  let email = raw.data[raw.head.indexOf(EMAIL_COLUMN_NAME)];
  fireEmail(email,SUBJECT,downloadedFiles,html);
}



/**
  * simplified version of GmailApp.sendEmail
  *
  * @param {String} email Recievers email address
  * @param {String} SUBJECT the name of the spreadsheet column containing images
  * @param {Array} downloadedFiles Array of downloaded files to be attached as attachment. Canbe generated by `downloadAllSlides(presentationId, slideIds,OUTPUT_FOLDER_ID,true)`
  * @param {String} html pure html text. Can be generated by `prepareEmail(ssId,files,EMAIL_TEMPLATE_ID)`
  * @param {String} name You Can set name of the sender. Default is "Team Automation"
*/
function fireEmail(email,SUBJECT,downloadedFiles,html,name="Team Automation"){
  if (MailApp.getRemainingDailyQuota() > 0){
    GmailApp.sendEmail(email, SUBJECT, "sssss",{
      htmlBody: html,
      name: name,
      attachments: downloadedFiles
  });
  } 
}


/**
  * Generates the data to a doc template then converts it to html string
  *
  * @param {Object} raw the object returned by function `extractData(ssId)
  * @param {String} email Recievers email address
  * @param {String} SUBJECT the name of the spreadsheet column containing images
  * @param {Array} downloadedFiles Array of downloaded files to be attached as attachment. Canbe generated by `downloadAllSlides(presentationId, slideIds,OUTPUT_FOLDER_ID,true)`
  * @param {String} html pure html text. Can be generated by `prepareEmail(ssId,files,EMAIL_TEMPLATE_ID)`
  * @param {String} name You Can set name of the sender. Default is "Team Automation"
  * @return {Array} Returns an array of slide Ids of all the slides of the processed presentation
*/
function prepareEmail(raw,files,EMAIL_TEMPLATE_ID){
  let file = DriveApp.getFileById(EMAIL_TEMPLATE_ID).makeCopy('tempEmail');
  let docId = file.getId();
  let doc = DocumentApp.openById(docId);
  let body = doc.getBody();
  // let text = body.getText();
//  let toKill = [...text.matchAll(/\{\{.*\}\}/g)][0];
  // let raw = extractData(ssId,sheetName);
  let matchThis,tracker;
  raw.head.forEach(variable =>{
    matchThis = '{{' + variable + '}}';
    body = body.replaceText(matchThis, raw.data[raw.head.indexOf(variable)]);
  });
  
  Logger.log("Body Text\n"+body.getText());
//  let files = [raw.data[raw.head.length-1], raw.data[raw.head.length-2]];
  
  /* Image Inserter-- doc approach*/
  tracker = 0;
  files.forEach(image =>{
    tracker++;
    body.appendParagraph('Layout :'+ tracker + " Download").setLinkUrl(image.getDownloadUrl());
    var smile = body.appendImage(image);
    smile.setLinkUrl(image.getDownloadUrl());
    var height = smile.getHeight();
    var width = smile.getWidth();
    if (height > width){
      height = 550;
      width = 500;
    } else if (height < width){
      height = 500;
      width = 550;
    } else {
      height = 550;
      width = 550;
    } 
    smile.setHeight(0.5*height).setWidth(0.5*width);
  });
  
  /* html approach
  let html = docToHtml(docId);
  let tracker,varName;
  html = HtmlService.createHtmlOutput(html);
  files.forEach(fileLink => {
    tracker = files.indexOf(fileLink) + 1;
    varName = 'image'+tracker;
    html.append('<p>Layout '+ tracker +':<br>'+'<img src="cid:' + varName + '" /></p>');
  });
  */
  doc.saveAndClose();
  let html = docToHtml(docId);
  file.setTrashed(true);
  return html     
}


/**
  * Generates the output file links to sheet.
  * Sets column name As Output1, Output2..etc if not already exists.
  * 
  * @param {Sheet} sheet Current sheet of a spreadsheet
  * @param {Array} downloadedFiles Array of downloaded files to be attached as attachment. Canbe generated by `downloadAllSlides(presentationId, slideIds,OUTPUT_FOLDER_ID,true)`
  * @return {Array} Returns an array of slide Ids of all the slides of the processed presentation
*/
function generateOutputLinksToSheet(sheet,downloadedFiles){
  for (var i=0;i<downloadedFiles.length;i++){
    var startCol = sheet.getLastColumn() + 1;
    var outNo = i+1;
    var colName = 'Output'+ outNo;
    var curHead = sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0];
    if (curHead.indexOf(colName) == -1) {
      sheet.getRange(1, sheet.getLastColumn()+1).setValue(colName);
      curHead = sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0];
    }
    sheet.getRange(sheet.getLastRow(), curHead.indexOf(colName)+1).setValue(downloadedFiles[i].getUrl());
    startCol++;
    
  }
}


function renameVariables(presentationId, oldVar, newVar){
  let presentation = SlidesApp.openById(presentationId);
  newVar = '{{' + newVar + '}}';
  presentation.replaceAllText(oldVar, newVar);
  presentation.saveAndClose(); 
}

function extracId(string){
    // Logger.log(string)
   let matched = string.match(/[-\w]{25,}/);
   return matched[0] ;
}

function whatsTheFont(string,BAN_FONT,ENG_FONT){
  var code = string.charCodeAt(0);
  if (code > 20 && code < 127){
    return ENG_FONT
  }else{
    return BAN_FONT
  }
}

function dateToNiceString(myDate){
/*Reference:
https://www.codegrepper.com/code-examples/delphi/javascript+format+date+to+dd-mm-yyyy*/
  var month=new Array();
  month[0]="Jan";
  month[1]="Feb";
  month[2]="Mar";
  month[3]="Apr";
  month[4]="May";
  month[5]="Jun";
  month[6]="Jul";
  month[7]="Aug";
  month[8]="Sep";
  month[9]="Oct";
  month[10]="Nov";
  month[11]="Dec";
  var hours = myDate.getHours();
  var minutes = myDate.getMinutes();
  var ampm = hours >= 12 ? 'pm' : 'am';
  hours = hours % 12;
  hours = hours ? hours : 12;
  minutes = minutes < 10 ? '0'+minutes : minutes;
  var strTime = hours + '_' + minutes + ampm;
  return month[myDate.getMonth()]+"-"+myDate.getDate()+"-"+myDate.getFullYear()+"-"+strTime;
}

function resizing(id, width){
  // https://github.com/tanaikech/ImgApp
  var res = ImgApp.doResize(id, width);
  var width = res.resizedwidth;
  var height = res.resizedheight;
  var file = DriveApp.createFile(res.blob.setName("filename"));
  return file
}



function docToHtml(tempId){
  var theApi = "https://docs.google.com/feeds/download/documents/export/Export?id="+tempId+"&exportFormat=html";
  var param = {
    method      : "get",
    headers     : {"Authorization": "Bearer " + ScriptApp.getOAuthToken()}
  };
  var sing = UrlFetchApp.fetch(theApi,param).getContentText().trim();
  return sing
}


function play(){
  Logger.log(MailApp.getRemainingDailyQuota());
}
