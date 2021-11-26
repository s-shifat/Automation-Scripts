// ########################################################################
// Library Used in this script: SeaShell
// Version: 22
// Script ID: `1RHAcPDkl_4XOwezi1YLMsJCOQJBv6ClsGngRoZ3WqETQLgWPMAk_mLYt`
// Reference to Library: SeaShellshifat
// ########################################################################


// const SPREADSHEET_ID = '';
// const PRESENTATION_ID = '';
// const OUTPUT_FOLDER_ID = '';
// const ROOT_FOLDER_ID = '';

// const PHOTOLINK_CONATINING_COLUMN_NAME = 'Question (As Picture)';
// const OUTPUT_FILE_NAME_DETERMINING_COLUMN_NAME = 'Challenge No';
// const OUTPUT_SHEET_NAME = 'OutPuts';
// const OUTPUT_HEADER_ROW = ['Sheet Name','Challange No','Presentation Name', 'Presentation Link'];


// Takes data from google sheets, parses it, then replaces the data in a google slide tamplate
function robustQnA(PRESENTATION_ID, OUTPUT_FOLDER_ID, PHOTOLINK_CONATINING_COLUMN_NAME,
OUTPUT_FILE_NAME_DETERMINING_COLUMN_NAME, OUTPUT_SHEET_NAME, OUTPUT_HEADER_ROW){
  let spreadsheet = SpreadsheetApp.getActive();
  let sheet = spreadsheet.getActiveSheet();
  let selection = sheet.getSelection();
  let ranges = selection.getActiveRangeList().getRanges();
  let head = sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0];
  let indexOf, currentData, currentDataRows, currentDataColumns, raw, rawData, presoId, namePrefix;
  let presoIds = [];
  let finalRow = [];
  finalRow.push(sheet.getName());
  ranges.forEach(range => {
    currentData = range.getDisplayValues();
    currentDataRows = currentData.length;
    currentDataColumns = currentData[0].length;
    indexOf = indexs(range, currentDataRows, currentDataColumns);
    rawData = sheet.getRange(indexOf.starting.row,1, currentDataRows, sheet.getLastColumn());
    rawData.getDisplayValues().forEach(row => {
      raw = {
        'head': head,
        'data': row
      };
      if(ranges.indexOf(range) == 0){
        namePrefix = raw.data[raw.head.indexOf(OUTPUT_FILE_NAME_DETERMINING_COLUMN_NAME)];
      }
      presoId = SeaShellshifat.replaceTheTexts(raw, PRESENTATION_ID);
      presoId = SeaShellshifat.replaceImages(raw, presoId, PHOTOLINK_CONATINING_COLUMN_NAME);
      presoIds.push(presoId);
    });
  });
  Logger.log(presoIds);
  presoId = combinePresos(presoIds);
  let finalPresoFile = DriveApp.getFileById(presoId);
  let outputFolder = DriveApp.getFolderById(OUTPUT_FOLDER_ID);
  finalPresoFile.moveTo(outputFolder);
  finalPresoFile.setName(OUTPUT_FILE_NAME_DETERMINING_COLUMN_NAME + '-' +namePrefix + "-" + SeaShellshifat.dateToNiceString(new Date()));
  finalRow.push(namePrefix);
  finalRow.push(finalPresoFile.getName());
  finalRow.push(finalPresoFile.getUrl());
  Logger.log(OUTPUT_HEADER_ROW);
  Logger.log(finalRow);

  sheet = dumpOutPutToSheet(spreadsheet, OUTPUT_SHEET_NAME, OUTPUT_HEADER_ROW, finalRow);
  sheet.getRange(sheet.getLastRow(),1,1,sheet.getLastColumn()).activate();

}

function indexs(range, rows, cols){
  return {
    'starting': {
      'row': range.getRow(),
      'col': range.getColumn(),
    },
    'ending': {
      'row': rows + range.getRow() - 1,
      'col': cols + range.getColumn() - 1,
    }
  }
}

function combinePresos(presoIds, trashed=true){
  let preso, mainPreso;
  presoIds.forEach(preosId =>{
    if (presoIds.indexOf(preosId) == 0){
      mainPreso = SlidesApp.openById(preosId);
    } else{
      preso = SlidesApp.openById(preosId);
      _combiner(mainPreso, preso, trashed);
    }
  });
  return mainPreso.getId()
}

function _combiner(mainPreso, preso, trashed){
  preso.getSlides().forEach(slide => {
        mainPreso.appendSlide(slide);
      });
  DriveApp.getFileById(preso.getId()).setTrashed(trashed);
}

function dumpOutPutToSheet(spreadsheet, sheetName, headerRow, dataRow){
  let sheet, headerRange, fullRange;
  try{
    sheet = spreadsheet.insertSheet(sheetName);
    sheet = sheet.appendRow(headerRow);  
    Logger.log('created new output sheet');
  } catch (e) {
    Logger.log('output sheet exists');
    sheet = spreadsheet.getSheetByName(sheetName);
  }

  headerRange = sheet.getRange(1,1,1, sheet.getLastColumn());
  fullRange = sheet.getRange(1,1,sheet.getMaxRows(), sheet.getMaxColumns());
  fullRange.clearFormat().applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREEN, true, false);
  headerRange.setFontWeight('bold').setFontColor("#4b089d");
  sheet.setFrozenRows(1);
  sheet.appendRow(dataRow);
  return sheet
}

// function onOpen(){
//   const ui = SpreadsheetApp.getUi();
//   ui.createMenu('Team Automation').addItem(
//     'Convert to Presentation',
//     'robustQnA'
//   ).addToUi();
// }

/*
  Author: Shams-E-Shifat
  Email: sshifat022@gmail.com
*/


