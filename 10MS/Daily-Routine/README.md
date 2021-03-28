# About dailyRoutineExecution.gs

 Made solely for a specific automation of Ten Minute School.

### What will you need?

  1. Four ***Google Slides*** files as templates. You can add *place holders* following this format: `{{PlaceHolder_Name}}`.
    Two of these are HSC story and poster templates and rest two are SSC story and poster.

  2. A ***Google Sheets*** file that has column names of the *place holders* you used in your template slide file. You will also need another sheet in the
     file to dump the outputs. This should contain five columns: Date & Day, Class, Poster pptx Link, Poster JPG Link, Story pptx Link, Story JPG Link.

  3. A ***Folder*** in your google drive to store the outputs.

### How to activate?

  1. Open the spreadsheet I mention earlier.

  2. Now from the ribbon go to `Tools` > `Script editor`. This will open an editor in a new tab.

  3. Delete what ever is written in the editor and paste the code from [`dailyRoutineExecution.gs`](https://github.com/s-shifat/Automation-Scripts/blob/main/10MS/Daily-Routine/dailyRoutineExecution.gs) file.

  4. Now from the editor click the '+' icon beside `Libraries` (should be at the left).

  5. A window will pop up with a title **Add a Library**. In the prompt you have to paste the script id.
    The current <br> 
    **script id:** `1Xh71ZbqgMKOJmLPiXwaO6VUydGRE8ODor7oBxSgQe4vmLf8MdoGHD8An` <br>
    Now click on **Look up**.  This should find the library.
    From the drop down menu under *Version* select `8` (or whichever is the latest).
    Keep the **Identifier** as it is. and click on *Add* button.

  6. Now hit `ctrl + s` from your keyboard. Thats it.
  7. Now go to your google spreadsheet file and refresh. Look at the ribbon, there should be an extra menu labeled as **Team Automation**.

### How to use?

  After activating the library, go to your google spreadsheet select a range. Now from the ribbon click `Team Automation` > `Make Routine`. And that is pretty much it!

##### If you find/face any problem feel free to let me know.
