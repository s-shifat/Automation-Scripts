/*
    #############################################################################
    #  Library Information ######################################################
    #  Library Name: ForDesignTeamGenerateThumbnails022                         #
    #  Script Id: 1EpDLOyBvmjRtIlf6_7oygPVaYbxfGXsCvu4HycAiL2erQfF9AFa2cVYD     #
    #  Version: 9                                                               #
    #  Author: Shams-E-Shifat                                                   #
    #  Email: sshifat022@gmail.com                                              #
    #############################################################################
*/

// Editable Part Starts-----------------------------------------------------------------------------------

const PHOTO_CONTAINING_COLUMN_NAME = ''; // Keep it empty if current sheet doesn't have any column that contatins photo link

// Editable Part Ends--------------------------------------------------------------------------------------

// Don't Modify anything below this line~~
function onOpen(){
  let ui = SpreadsheetApp.getUi();
  ui
    .createMenu(ForDesignTeamGenerateThumbnails022.houseKeeping().TEAM_NAME)
    .addItem(ForDesignTeamGenerateThumbnails022.houseKeeping().ITEM_1, 'processThumbnail')
    .addItem(ForDesignTeamGenerateThumbnails022.houseKeeping().ITEM_2, 'config')
    .addToUi();
}

function config(){
  ForDesignTeamGenerateThumbnails022
    .configuration();
}

function processThumbnail(){
  ForDesignTeamGenerateThumbnails022
    .generateThumbnail(PHOTO_CONTAINING_COLUMN_NAME);
}
