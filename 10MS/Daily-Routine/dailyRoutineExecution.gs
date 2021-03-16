/*
    #############################################################################
    #  Library Information ######################################################
    #  Library Name: DailyRoutineshifat10MS                                     #
    #  Script Id: 1Xh71ZbqgMKOJmLPiXwaO6VUydGRE8ODor7oBxSgQe4vmLf8MdoGHD8An     #
    #  Version: 8                                                               #
    #  Author: Shams-E-Shifat                                                   #
    #  Email: sshifat022@gmail.com                                              #
    #############################################################################
*/

// Editable Part Starts-----------------------------------------------------------------------------------

const SSC_STORY_ID = '1o97xjvHK7-KFhThv9K1ZhzJhTrmdKaWMm2MGQ8kzGF0';
const SSC_POSTER_ID = '1In4FwrQ3mVZpaT2N1AEAJ6L7VkBz10wDpWfWhftJE_M';
const HSC_STORY_ID = '1AqAmPXi4-f7j54Ww2lhcJ_aIfvq7rZwt53Au_EUp9x0';
const HSC_POSTER_ID = '16dSPrSObXlc4yGkbh18dDogbhDY4lOjN6PGFoQxQQgY';
const SPREAD_SHEET_ID = '1j21EFQHgMKuhTX6JxDYi0Vv4b5Z_HiqwRP-Mf7p-i7k';

const ROOT_FOLDER_ID = '1AEmxZ38SS2Zx_Jz_bOey4e68hNxOZuRN';
const OUTPUT_FOLDER_ID = '1ALAVh0TrcvRqyrFRtN3e0tT4IPfd6J4Q';

const DATA_CONTAINING_SHEET_NAME = 'Data';
const OUTPUT_CONTAINING_SHEET_NAME = 'Output';
const CLASS_CONTAINING_COLUMN_NAME = 'Class:'; // Check if there is/are any whitespace
const PICTURE_TYPE_CONTAINING_COLUMN_NAME = 'Picture Type:'; // Check if there is/are any whitespace
const DATE_N_DATE_CONTAINING_COLUMN_NAME = 'Date & Day'; // Check if there is/are any whitespace

// Editable Part Ends-------------------------------------------------------------------------------------


function onOpen(){
  let ui = SpreadsheetApp.getUi();
  ui
    .createMenu('Team Automation')
    .addItem('Make Routine', 'execute')
    .addToUi();
}

function execute(){
  DailyRoutineshifat10MS.dailyRoutine(
    SSC_STORY_ID,SSC_POSTER_ID,
    HSC_STORY_ID,
    HSC_POSTER_ID,
    SPREAD_SHEET_ID,
    ROOT_FOLDER_ID,
    OUTPUT_FOLDER_ID,
    DATA_CONTAINING_SHEET_NAME,
    OUTPUT_CONTAINING_SHEET_NAME,
    CLASS_CONTAINING_COLUMN_NAME,
    PICTURE_TYPE_CONTAINING_COLUMN_NAME,
    DATE_N_DATE_CONTAINING_COLUMN_NAME
  );
}

