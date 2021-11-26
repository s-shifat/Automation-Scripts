# SeaShell Library

This is a library I wrote to help me complete the automation tasks in 10 Minute School. While completing the tasks, I found myself writting similar functions
for some particular tasks like converting presentation slides to png/jpg formats, sending automated emails to clients etc. So I noted all those tasks and ended
up writing up this library. This made my life a lot easier, and increased the productivity exponentially.

This library is well documented. So if you write google apps script to combine services provided by Google, then this library will reduce your pain.

### Different Function Features

 * Generates the output file links to sheet.
 * Downloads every slide of a presentation as jpeg to a specified folder in the google drive.
 * Replaces the texts in the defined variables of template presentation with respective column values of spreadsheet. The defined variables are the name of the spreadsheet's column. In the template presentation the varibale format shoulb be like {{column name}} in a text box.
 * Generates the data to a doc template then converts it to html string
 * Downloads a slide of a presentation as jpeg to a specified folder in google drive.
 * Log every usefull detatils of each element from a presentation. Usefull for going through a presentaion real fast.
 * Easy function to fire email to clients
 *  reads in the text of a text box or shape, If the text is a valid link of google drive photo, then renders the photo IE, converts the shape to image while going through each slide of given presentation In the template presentation the varibale format shoulb be like {{column name}} in a text box. It is recommended that you should run the `replaceTheTexts(raw, presentationId)` function first, before using this function.
 *  Extracts data from a spreadsheet in a convenient format. Returns an object having two attributes.--- Object.head --> Array containing name of the columns Object.data --> Array containing values of the column names, from last row of spreadsheet @sheetName {String} name of the sheet to be considered
 *  Convert Document to Html
 *  A single function doing a sample total work flow. Takes the latest row of a spreadsheet, then generates a presentation, then downloads it. Also sends email. You may use global constants as the arguments.

### Library Information

  * Version: 22
  * Script ID: `1RHAcPDkl_4XOwezi1YLMsJCOQJBv6ClsGngRoZ3WqETQLgWPMAk_mLYt`
  * Project Link: [`click here`](https://script.google.com/d/1RHAcPDkl_4XOwezi1YLMsJCOQJBv6ClsGngRoZ3WqETQLgWPMAk_mLYt/edit?usp=sharing)
  * Documentation: [`click here`](https://script.google.com/macros/library/d/1RHAcPDkl_4XOwezi1YLMsJCOQJBv6ClsGngRoZ3WqETQLgWPMAk_mLYt/22)

