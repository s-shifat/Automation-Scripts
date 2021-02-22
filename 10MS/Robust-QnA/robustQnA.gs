/*
    #############################################################################
    #  Library Information ######################################################
    #  Library Name: RobustQnA_ses                                              #
    #  Script Id: 1mHFjh7sH5TrU3oxy1jMliC85O06zrK04AGXxwXRzUUeetgYtq-HZjgWh     #
    #  Version: 2                                                               #
    #############################################################################
 */

// Editable Part Starts-----------------------------------------------------------------------------------

// Put Relevent Ids Here, you will find this from the url of the docs.
// Sample Ids are given here, replace them with yours
const SPREADSHEET_ID = '1ko87VatucjmY93_bDmdtmEB2y22BFNpVSpVOWlxjS58'; 
const PRESENTATION_ID = '15zuwSDcwWhQZNYmA_N81r2uaNqSELy2bz10Z7oC2EEA';
const OUTPUT_FOLDER_ID = '10CRsAuvyASt4w5JOKkpqmw_rp-WVAWhK';

// Fill Up these variables from your Spreadsheet
const PHOTOLINK_CONATINING_COLUMN_NAME = 'Question (As Picture)';
const OUTPUT_FILE_NAME_DETERMINING_COLUMN_NAME = 'Challenge No';
const OUTPUT_SHEET_NAME = 'OutPuts';
const OUTPUT_HEADER_ROW = ['Sheet Name','Challange No','Presentation Name', 'Presentation Link'];

// Editable Part Ends-------------------------------------------------------------------------------------

function onOpen(){
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Team Automation').addItem(
    'Convert to Presentation',
    'automation'
  ).addToUi();
}

function automation(){
  RobustQnA_ses.robustQnA(
    PRESENTATION_ID,
    OUTPUT_FOLDER_ID,PHOTOLINK_CONATINING_COLUMN_NAME,
    OUTPUT_FILE_NAME_DETERMINING_COLUMN_NAME,
    OUTPUT_SHEET_NAME,
    OUTPUT_HEADER_ROW);
}

/*
  Author: Shams-E-Shifat
  Email: sshifat022@gmail.com
*/