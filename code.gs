/** @OnlyCurrentDoc */
var mySS = SpreadsheetApp.getActiveSpreadsheet();
var mySheet = mySS.getActiveSheet();
var dataRange = mySheet.getDataRange();
var numberOfColumns = dataRange.getNumColumns();
var numberOfRows = dataRange.getNumRows();
var headerRange = mySheet.getRange(1, 1, 1, numberOfColumns);

// Constants that identify the index of the columns. 
var TIMESTAMP_COLUMN = 1;
var ISBN_COLUMN = 2;
var TITLE_COLUMN = 3;
var AUTHOR_COLUMN = 4;
var PUBLISHER_COLUMN = 5;
var PAGE_COLUMN = 6;
var DATE_COLUMN = 7;
var LANGUAGE_COLUMN = 8;
var SUBTITLE_COLUMN = 9;
var LABEL_COLUMN = 10;
var IMAGE_COLUMN = 11;
var LINK1_COLUMN = 12;

/**
 * The event handler triggered when installing the add-on.
 * @param {Event} e The onInstall event.
 * @see https://developers.google.com/apps-script/guides/triggers#oninstalle
 */
function onInstall(e) {
  var headers = [
    ['Timestamp', 'Barcode', 'Title', 'Author', 'Publisher', 'Page',
      'Date', 'Lang.', 'Subtitle', 'Label', 'Picture', 'Link'
    ]
  ];
  var headerRange = mySheet.getRange(1, 1, 1, headers[0].length);
  headerRange.setValues(headers);
  // set sample book
  var sampleBook = [
    ['9780451527745',
      "Alice's Adventures in Wonderland & Through the Looking-Glass",
      'Lewis Carroll', 'Signet Classics',  239, '12/12/2000', 'EN', '', '',
      'https://covers.openlibrary.org/b/id/6449369-L.jpg',
      'https://libgen.is/search.php?req=9780451527745&open=0&res=25&view=simple&phrase=1&column=def'
    ]
  ];
  mySheet.getRange(2, 2, 1, headers[0].length-1).setValues(sampleBook);
  format(headers[0].length);
  onOpen(e);
}

function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Edit Books List')
    .addItem('Reload Selected Books', 'fillTheRange')
    .addSeparator()
    .addItem('Reset Format', 'format')
    .addToUi();
}

function fillTheRange_onEdit() {
  fillTheRange('onEdit');
}


function fillTheRange(command = 'manual') {

  var activeRange = mySheet.getActiveRange();
  var ISBNRange = formatRange_(activeRange, command);
  if (ISBNRange) {
    var numberOfISBNRows = ISBNRange.getNumRows();
    var ISBNValues = ISBNRange.getValues();

    var books = [numberOfISBNRows];
    for (var i = 0; i < numberOfISBNRows; i++) {
      books[i] = new Array(numberOfColumns - 2);
      //books[i][ISBN_COLUMN] = ISBNValues[i];
    }

    for (var row = 0; row < numberOfISBNRows; row++) {
      var title = books[row][TITLE_COLUMN - 3];
      var author = books[row][AUTHOR_COLUMN - 3];
      var page = books[row][PAGE_COLUMN - 3];
      var date = books[row][DATE_COLUMN - 3];
      var language = books[row][LANGUAGE_COLUMN - 3];
      var publisher = books[row][PUBLISHER_COLUMN - 3];
      var subtitle = books[row][SUBTITLE_COLUMN - 3];
      var image = books[row][IMAGE_COLUMN - 3];

      if (ISBNValues[row] != null && ISBNValues[row] != '' && ISBNValues[
          row] != ' ') {
        // Call open books api and fill in the blanks
        var bookData = fetchBookData_(ISBNValues[row], 'openbooksapi');

        // Call Google books API and fill in the blanks.
        var bookData2 = fetchBookData_(ISBNValues[row],
          'googlebooksapi');
        // Sometimes the API doesn't return the information needed.
        // In those cases, don't attempt to update the row.
        if ((!bookData || !bookData.details) && (!bookData2 || !
            bookData2.volumeInfo)) {
          continue;
        }
        if (bookData) {
          // The API might not return a title, so only fill it in
          // if the response has one and if the title is blank in
          // the sheet.
          if ((title == null || title == "") && bookData.details
            .title) {
            title = bookData.details.title;

          }
          // The API might not return an author name, so only fill it in
          // if the response has one and if the author is blank in
          // the sheet.
          if ((author == null || author == "") && bookData.details
            .authors &&
            bookData.details.authors[0].name) {
            author = bookData.details.authors[0].name;
          }
          // The API might not return a page number, so only fill it in
          // if the response has one and if the page number is blank in
          // the sheet.
          if ((page == null || page == "") && bookData.details
            .number_of_pages) {
            page = bookData.details.number_of_pages;
          }
          // The API might not return a publish date, so only fill it in
          // if the response has one and if the date is blank in the sheet.
          if ((date == null || date == "") && bookData.details
            .publish_date) {
            date = bookData.details.publish_date;
          }
          // The API might not return a language, so only fill it in
          // if the response has one and if the lang is blank in the sheet.
          if ((language == null || language == "") && bookData.details
            .languages &&
            bookData.details.languages[0].key) {
            //language = bookData.details.languages[0].key; 
            if (bookData.details.languages[0].key ==
              '/languages/tur') {
              language = 'TR';
            } else if (bookData.details.languages[0].key ==
              '/languages/eng') {
              language = 'EN';
            } else {
              language = bookData.details.languages[0].key;
            }
          }
          // The API might not return a publisher, so only fill it in
          // if the response has one and if the publisher is blank in the sheet.
          if ((publisher == null || publisher == "") && bookData
            .details.publishers &&
            bookData.details.publishers[0]) {
            publisher = bookData.details.publishers[0];
          }
          // The API might not return a subtitle, so only fill it in
          // if the response has one and if the subtitle is blank in the sheet.
          if ((subtitle == null || subtitle == "") && bookData.details
            .other_titles &&
            bookData.details.other_titles[0]) {
            subtitle = bookData.details.other_titles[0];
          }
          if ((image == null || image == "") && bookData
            .thumbnail_url) {
            image = bookData.thumbnail_url;
            image = image.replace("S.jpg", 'L.jpg');
          }
        }
        if (bookData2) {
          // The API might not return a title, so only fill it in
          // if the response has one and if the title is blank in
          // the sheet.
          if ((title == null || title == "") && bookData2.volumeInfo
            .title) {
            title = bookData2.volumeInfo.title;
          }
          // The API might not return an author name, so only fill it in
          // if the response has one and if the author is blank in
          // the sheet.
          if ((author == null || author == "") && bookData2.volumeInfo
            .authors) {
            author = bookData2.volumeInfo.authors;
          }
          // The API might not return a page number, so only fill it in
          // if the response has one and if the page number is blank in
          // the sheet.
          if ((page == null || page == "") && bookData2.volumeInfo
            .pageCount) {
            page = bookData2.volumeInfo.pageCount;
          }
          // The API might not return a publish date, so only fill it in
          // if the response has one and if the date is blank in the sheet.
          if ((date == null || date == "") && bookData2.volumeInfo
            .publishedDate) {
            date = bookData2.volumeInfo.publishedDate;
          }
          // The API might not return a language, so only fill it in
          // if the response has one and if the lang is blank in the sheet.
          if ((language == null || language == "") && bookData2
            .volumeInfo.language) {
            if (bookData2.volumeInfo.language == 'tr') {
              language = 'TR';
            } else if (bookData2.volumeInfo.language == 'en') {
              language = 'EN';
            } else {
              language = bookData2.volumeInfo.language;
            }
          }
          // The API might not return a publisher, so only fill it in
          // if the response has one and if the publisher is blank in the sheet.
          if ((publisher == null || publisher == "") && bookData2
            .volumeInfo.publisher) {
            publisher = bookData2.volumeInfo.publisher;
          }
          // The API might not return a subtitle, so only fill it in
          // if the response has one and if the subtitle is blank in the sheet.
          if ((subtitle == null || subtitle == "") && bookData2
            .volumeInfo.subtitle) {
            subtitle = bookData2.volumeInfo.subtitle;
          }
        }
        books[row][TITLE_COLUMN - 3] = title;
        books[row][AUTHOR_COLUMN - 3] = author;
        books[row][PAGE_COLUMN - 3] = page;
        books[row][DATE_COLUMN - 3] = date;
        books[row][LANGUAGE_COLUMN - 3] = language;
        books[row][PUBLISHER_COLUMN - 3] = publisher;
        books[row][SUBTITLE_COLUMN - 3] = subtitle;
        books[row][IMAGE_COLUMN - 3] = image;
        books[row][LINK1_COLUMN - 3] =
          "https://libgen.is/search.php?req=" +
          ISBNValues[row] +
          "&open=0&res=25&view=simple&phrase=1&column=def";
      }
    }

    // Insert the updated book data values into the spreadsheet.
    mySheet.getRange(ISBNRange.getRow(), 3, numberOfISBNRows, books[0]
      .length).setValues(books);
  }
}

function formatRange_(activeRange, command) {
  if (command == 'onEdit') {
    if (activeRange.getColumn() <= ISBN_COLUMN && activeRange
      .getLastColumn() >= ISBN_COLUMN) {
      var startRow = ((activeRange.getRow() < 2) ? 2 : activeRange
        .getRow());
      var endRow = ((activeRange.getLastRow() > numberOfRows) ?
        numberOfRows : activeRange.getLastRow());
      if (startRow > endRow) {
        startRow = endRow;
      }
      activeRange = mySheet.getRange(startRow, ISBN_COLUMN, endRow -
        startRow + 1, 1);
      return activeRange;
    } else {
      return null;
    }
  } else {
    var startRow = ((activeRange.getRow() < 2) ? 2 : activeRange.getRow());
    var endRow = ((activeRange.getLastRow() > numberOfRows) ? numberOfRows :
      activeRange.getLastRow());
    if (startRow > endRow) {
      startRow = endRow;
    }
    activeRange = mySheet.getRange(startRow, ISBN_COLUMN, endRow -
      startRow + 1, 1);
    return activeRange;
  }
}

/**
 * Helper function to retrieve book data from the Open Library
 * public API.
 *
 * @param {number} ISBN - The ISBN number of the book to find.
 * @return {object} The book's data, in JSON format.
 */
function fetchBookData_(ISBN, api) {
  // Connect to the public API.

  if (api == 'googlebooksapi') {
    url = "https://www.googleapis.com/books/v1/volumes?q=isbn:" + ISBN +
      "&country=US";

    var response = UrlFetchApp.fetch(
      url, {
        'muteHttpExceptions': true
      });

    // Make request to API and get response before this point.
    json = response.getContentText();
    var bookData = JSON.parse(json);
    if (bookData.totalItems) {
      return bookData.items[0];
    }

  } else {
    var url = "https://openlibrary.org/api/books?bibkeys=ISBN:" +
      ISBN + "&jscmd=details&format=json";
    var response = UrlFetchApp.fetch(
      url, {
        'muteHttpExceptions': true
      });

    // Make request to API and get response before this point.
    var json = response.getContentText();
    var bookData = JSON.parse(json);
    return bookData['ISBN:' + ISBN];
  }
}

function format(length = 0) {
  if (length) {
    dataRange = mySheet.getRange(1, 1, 3, length);
    headerRange = mySheet.getRange(1, 1, 1, length);
  }
  // Lock header
  mySheet.setFrozenRows(1);

  // reset alternating background
  var banding = dataRange.getBandings()[0];
  if (banding != null) {
    banding.remove();
  }
  // set alternating background
  dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.GREY);

  // Format header row
  banding = dataRange.getBandings()[0]
    .setHeaderRowColor('#980000');
  headerRange.setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBorder(false, false, true, false, false, false, null,
      SpreadsheetApp.BorderStyle.SOLID_MEDIUM);


  // set number format of ISBN column as plain text
  var totalNumRows = mySheet.getRange('A:A').getNumRows() - 1;
  mySheet.getRange(2, ISBN_COLUMN, totalNumRows).setNumberFormat('@');
  mySheet.getRange(2, PAGE_COLUMN, totalNumRows).setNumberFormat('#,##0');

  // set wrap strategies
  setWrapAndColumnSize_();

  // hide columns
  mySheet.hideColumns(1, 2);
  mySheet.getRange('C' + numberOfRows).activate();

}

function setWrapAndColumnSize_() {

  var maxColumnWidth = 250;
  var maxLinkColWidth = 100;
  var dataRange = mySheet.getDataRange();
  var numberOfColumns = (dataRange.getNumColumns() > 2) ? dataRange
    .getNumColumns() : 3;
  var numberOfRows = dataRange.getNumRows();
  var booksRange = mySheet.getRange(2, 1, numberOfRows, numberOfColumns - 2);
  var linkRange = mySheet.getRange(1, numberOfColumns - 2, numberOfRows,
    numberOfColumns);

  booksRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  headerRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
  linkRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  //Resize all the columns
  mySheet.autoResizeColumns(1, numberOfColumns);
  // Limit the column size
  for (let i = 0; i < numberOfColumns - 3; i++) {
    var columnWidth = mySheet.getColumnWidth(i + 1);
    if (columnWidth > maxColumnWidth)
      mySheet.setColumnWidth(i + 1, maxColumnWidth);
  }
  for (let i = numberOfColumns - 2; i < numberOfColumns; i++) {
    var columnWidth = mySheet.getColumnWidth(i + 1);
    if (columnWidth > maxLinkColWidth)
      mySheet.setColumnWidth(i + 1, maxLinkColWidth);
  }
}
