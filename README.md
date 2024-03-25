
# Flooring Cost Estimator for Centris Matrix
This Google Sheet is designed to receive a copy-pasted list of rooms from Centris Matrix. Then, you run the macro that comes with it - that'll generate a new list that allows you select the rooms whose floors you want to replace. Those selections then lead to an estimated cost range (low to high), based on the values entered in the options. Something that could be done by hand... somewhat simplified.

## Instructions

 1. Download the Excel workbook.
 2. Upload it into your Google Drive.
 3. Open it. 
 4. In the Google Sheet, click "Extensions", then "Macros", then "Record Macro".
 5. Press "Save" immediately.
 6. Enter "Process" as the name, give it a shortcut, and press "Save". 
 7. Click "Extensions", "Macros", then "Manage Macros".
 8. Click the "..." icon next to the "Process" macro and click "Edit script". 
 9. Replace the contents of the "macro.gs" file with the code I provided. 
 10. Save by pressing the diskette icon. 
 11. Press "Run".
 12. You'll be asked to provide authorization - click "Review permissions".
 13. Select your Google account. 
 14. Read about the risks and click "Allow" if you consent to them. (Spoiler: there's nothing nefarious in the code.)
 15. Return to the Google Sheet.

You can then paste the room info, excluding the headers, in the B column. Then simply press your shortcut to trigger the macro.

See? Wasn't it easy? Only 15 steps! When I decided between going with a Microsoft Excel solution and a Google Sheets solution, I thought I made the right choice with Google, writing macros in JavaScript, awesome! Apps Script has a great IDE, a pleasure to work with. Well... I can tell you one thing, Microsoft Excel wouldn't have taken 15 steps to set up 😂. Anyway, I almost didn't share this for that reason, but figured if it helps someone, why not.

