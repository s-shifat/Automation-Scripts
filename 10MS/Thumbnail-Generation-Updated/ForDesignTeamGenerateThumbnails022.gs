/**
*  This script Creates presentation file using data from google spread sheet services and template from
*  google slides services.
*  Script Id: 1EpDLOyBvmjRtIlf6_7oygPVaYbxfGXsCvu4HycAiL2erQfF9AFa2cVYDgoo
*  @author: Shams-E-Shifat
*  @email: sshifat022@gmail.com
*/

// const DETECTOR = ''; // Keep it empty if current sheet doesn't have any column that contatins photo link

function houseKeeping(){
  return{
    'TEAM_NAME':'Team Automation',
    'ITEM_1': 'Generate Thumbnail Slide',
    'ITEM_2': 'Set Output Folder'
  }
  
}

function onOpen(){
  let ui = SpreadsheetApp.getUi();
  ui.createMenu("Team Automation")
    .addItem("Generate Thumbnail Slide", "generateThumbnail")
    .addItem("Set Output Folder", "configuration")
    .addToUi();
}

function configuration(){
  let ui = SpreadsheetApp.getUi();
  let response = ui.prompt("Enter The Output Folder Url:\nYou only have to run this option the first time you use this spreadsheet.").getResponseText();
  let folder = extracId(response);
  folder = DriveApp.getFolderById(folder);
  let spreadsheet = SpreadsheetApp.getActive();
  let outsheet;
  try {
    outsheet = spreadsheet.insertSheet("Outputs");
    //1
    outsheet.getRange("A1").setValue("OutPut Folder Name:").setFontWeight('bold').setFontColor('#2d9719');
    outsheet.getRange("B1").setValue(folder.getName()).setFontWeight('bold').setFontColor('#2d9719').setFontStyle('italic');
    outsheet.getRange("1:1").setBackground("#fcfab4");
    //2
    outsheet.appendRow(['OutPut Folder Link: ',folder.getUrl()]);
    outsheet.getRange("A2").setFontColor("#2d9719").setFontWeight('bold');
    outsheet.getRange("B2").setFontColor("#1155cc").setFontStyle('italic');
    outsheet.getRange("2:2").setBackground("#fcfab4");
    //3
    outsheet.appendRow([' ',' ']);
    outsheet.getRange("3:3").setBackground("#b4b6a8");
    //4
    outsheet.appendRow(['Course Name','Slide Link']);
    outsheet.getRange("4:4").setBackground("#148cce").setFontColor("#fcfab4");
    outsheet.getRange('A4:B4').setFontWeight('bold');
  } catch (err) {
    let sheet = spreadsheet.getSheetByName("Outputs");
    sheet.getRange("B1").setValue(folder);
  }
}

function generateThumbnail(DETECTOR){
  let ui = SpreadsheetApp.getUi();
  let spreadsheet = SpreadsheetApp.getActive();
  let sheet = spreadsheet.getActiveSheet();
  let sheetName = sheet.getName();
  let rawData = sheet.getDataRange().getDisplayValues();
  let head = rawData[1];
  let data = rawData.slice(2, rawData.length);
  const OUTPUT_FOLDER_ID = spreadsheet.getSheetByName("Outputs").getRange("B2").getValue();
  let folder = DriveApp.getFolderById(extracId(OUTPUT_FOLDER_ID));
  let outsheet = spreadsheet.getSheetByName("Outputs");
  let templatePresso;
  try{
    let response = ui.prompt("Enter the Id or Url of the template slide please: ");
    const TEMPLATE_ID = extracId(response.getResponseText());
//    ui.alert("Extracted Id: "+TEMPLATE_ID); //Debuger
    templatePresso = DriveApp.getFileById(TEMPLATE_ID).makeCopy();
  }catch(err){
    let revisedResponse = ui.prompt("Sorry! That didn't work. Try again!\nMake sure you have pasted the correct slide url/id of the template.");
    const TEMPLATE_ID = extracId(revisedResponse.getResponseText());
//    ui.alert("Extracted Id: "+TEMPLATE_ID); //Debuger
    templatePresso = DriveApp.getFileById(TEMPLATE_ID).makeCopy();
    
  }
  
  let templatePressoId = templatePresso.getId();
  let midpresso = SlidesApp.openById(templatePressoId);
  let slides = midpresso.getSlides();
  let slideLength = slides.length;

  let outputPressoId = DriveApp.getFileById(templatePressoId).makeCopy(sheetName).getId();
  let outputPresso = SlidesApp.openById(outputPressoId);

  let i = 0;
  let midpressoSlideNo = 0;
  let repSlide;
  for(i = 0; i< data.length; i++){
    var raw = {
      'head': head,
      'data': data[i]
    }; 
    if (i >= slideLength){
      if(midpressoSlideNo == slideLength){
        midpressoSlideNo = 0;
      }
        repSlide = outputPresso.appendSlide(slides[midpressoSlideNo]);
        midpressoSlideNo++;
    }
    else if (i<slideLength){
      repSlide = outputPresso.getSlides()[i];
    }
    replaceTheTexts(raw, repSlide);
  }
  outputPresso.saveAndClose();

  if (DETECTOR != ''){
      var raw = {
        'head': head,
        'data': data
      }
      SeaShellshifat.replaceImages(raw, outputPressoId, DETECTOR);
  }

  DriveApp.getFileById(midpresso.getId()).setTrashed(true);
  // outputPressoId = outputTempOpenPresso.getId();
  // let outputPresso  = SlidesApp.openById(outputPressoId);

  spreadsheet.getSheetByName('Outputs').appendRow([sheetName, outputPresso.getUrl()]);
  // outputPresso.saveAndClose();
  SpreadsheetApp.flush();
}

function extracId(string){
  let matched = string.match(/[-\w]{25,}/);
 return matched[0] ;
}

function onlyUnique(value, index, self) {
  return self.indexOf(value) === index;
}


function playGround(){
  let stuff = ui.prompt("Enter Template Slide Id: ");
  let text = stuff.getResponseText();
  // Logger.log(text, typeof(text));
}

function replaceTheTexts(raw, slide){
  // let templateId = makeTemplatePresentation(presentationId);
  // let presentation = SlidesApp.openById(presentationId);
  let tobeRaplaced = '';
  let i=0;
  for (i=0; i<raw.head.length; i++){
    tobeRaplaced = '{{' + raw.head[i] + '}}';
    slide.replaceAllText(tobeRaplaced, raw.data[i]);
  }
  // presentation.saveAndClose();
  return slide.getObjectId();
}


/*
  @author: Shams-E-Shifat
  @email: sshifat022@gmail.com
*/
