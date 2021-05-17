# About [thumbnailGeneration.gs](https://github.com/s-shifat/Automation-Scripts/blob/main/10MS/Thumbnail-Generation-Updated/thumbnailGeneration.gs)

 Made solely for a specific automation of Ten Minute School.

### What will you need?

  1. A ***Google Slides*** file as a template. You can add *place holders* following this format: `{{PlaceHolder_Name}}`.

  2. A ***Google Sheets*** file that has column names of the *place holders* you used in your template slide file.

  3. A ***Folder*** in your google drive to store the outputs.

### How to activate?

  1. Open the spreadsheet I mention earlier.

  2. Now from the ribbon go to `Tools` > `Script editor`. This will open an editor in a new tab.

  3. Delete what ever is written in the editor and paste the code from [`thumbnailGeneration.gs`](https://github.com/s-shifat/Automation-Scripts/blob/main/10MS/Thumbnail-Generation-Updated/thumbnailGeneration.gs) file.

  4. Now from the editor click the '+' icon beside `Libraries` (should be at the left).

  5. A window will pop up with a title **Add a Library**. In the prompt you have to paste the script id.
    The current <br> 
    **script id:** `1EpDLOyBvmjRtIlf6_7oygPVaYbxfGXsCvu4HycAiL2erQfF9AFa2cVYD` <br>
    Now click on **Look up**.  This should find the library.
    From the drop down menu under *Version* select `14` or whichever is the latest.
    Keep the **Identifier** as it is. and click on *Add* button.

  6. Now hit `ctrl + s` from your keyboard. That's it!
  7. Now go to your google spreadsheet file and refresh. Look at the ribbon, there should be an extra menu labeled as **Team Automation**.

### How to use?
  After activating the library, you can do *two* things:
  
  1. To setup Output Folder you  can do any of the below:
     * From the ribbon click on `Team Automation` > `Set Output Folder`. This will ask for a folder link, paste your folder link there and hit ok. This will create a sheet named `Outputs` if it does not exists. <br>**OR**<br>
     * Paste your output folder link in `Cell B2` of `Outputs` sheet.
  
  2. The main interest of this project is creating the presentation. To do so,
     * From the ribbon click on `Team Automation` > `Generate Thumbnail Slide`. This will ask for a presentation link, paste the **link of the google presentation file** that contatins the variables (the ones conatining *{{}}*) and click on OK.
     * That is pretty much it! After the execution has finished go to `Outputs` sheet, the latest row contains the newly processed output presentation sotred in your folder.

##### If you find/face any problem feel free to let me know.
