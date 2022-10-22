# spreadsheet-isbn-to-library
Creates a google spreadsheet that fills the information of a book provided the ISBN.

With this project you can automatically fetch the information of any book to your Google spreadsheet and then see it in an app. Just write a valid ISBN to the specified column and the title, author, publisher, image, link columns will be filled. After that you can use this spreadsheet as a database to your Appsheet application. With this application you can scan a book barcode with your phone. The relative information will be saved to database. You can also edit, delete, filter the books inside of the application.

You can reach the codes code.gs

To create this project in your own enviroment:
1. Go to your Google Spreadsheets page and create a blank sheet.
2. Go to Extensions tab inside of your sheet and choose "Apps Script".
3. Copy the codes inside 'code.gs' and paste into your Apps Script page. Save the project.
4. Run the "onInstall" function inside of this document manually. You will have your header row and a sample book.
5. If you get (most probably you will) "Google hasn’t verified this app" error, verify it using your Google account. This is necessary because the program calls OpenBooksAPI and GoogleBooksAPI.
6. Go to "Triggers" in the left pane. (The clock symbol on the left in apps script) Add these three triggers:
   - fillTheRange_onEdit (event type: on edit)
   - fillTheRange_onEdit (event type: on change)
   - fillTheRange_onEdit (event type: on form submit) (This one requires another verification. Go ahead and verify.)

Now in your spreadsheet when you add an ISBN in the specified column, the related information will be added. You can also add one with a form or even the app that you composed via Appsheet.




