// ====================================================================================================================
//
// Name: GameTweetWatcher
//
// Desc:
//  Search Tweet based on a keyword which is specified in each Google Spreadsheet, and store the results into
//  the each sheet.
//
// Author:  Mune
//
// History:
//  2020-10-19 : Initial version
//
// ====================================================================================================================

// ID of the Target Google Spreadsheet (Book)
let VAL_ID_TARGET_BOOK       = '1zjxr1BO7fQmO-TDP-HYOL3uh-1g1zKOCwEogX0h3uew';
// ID of the Google Drive where the images will be placed
let VAL_ID_GDRIVE_FOLDER     = '1ttFnVjcZJNJdloaT4ni99ZbEYuEj-WAA';
// Key and Secret to access Twitter APIs
let VAL_CONSUMER_API_KEY     = 'lyhQXM3aZP5HHLqO6jjcEwzux';
let VAL_CONSUMER_API_SECRET  = 'ODBAEfj4VlgVf1BKHJyoz6QwryiaNtcchkr4PikzjSHmct0hjr';

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// DEFINES
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
let VERSION                  = 0.1;
let TIME_LOCALE              = "JST";
let FORMAT_DATETIME          = "yyyy-MM-dd (HH:mm:ss)";
let FORMAT_TIMESTAMP         = "yyyyMMddHHmmss";
let NAME_SHEET_USAGE         = "!USAGE";
let NAME_SHEET_LOG           = "!LOG";
let NAME_SHEET_ERROR         = "!ERROR";
let CELL_HEADER_LAST_KEYWORD = "Last Keyword :";
let CELL_HEADER_LAST_UPDATED = "Last Updated :";
let CELL_ROW_KEYWORD         = 1;
let CELL_ROW_LAST_KEYWORD    = 2;
let CELL_ROW_LAST_UPDATED    = 3;
let CELL_ROW_HEADER          = 5;
let CELL_ROW_DATA_START      = 6;

let CELL_HEADER_TITLES       = ["Tweet Id", "Created at", "Tweet", "User Id", "User name", "Media"];

let DEFAULT_MAX_NUM_TWEETS   = 10;

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// GLOBALS
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
let g_isDebugMode            = true;
let g_isEnabledLogging       = true;
let g_version                = 0.1;
let g_datetime               = TIME_LOCALE + ": " + Utilities.formatDate(new Date(), TIME_LOCALE, FORMAT_DATETIME);
let g_book                   = null;
let g_folder                 = null;

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// OBJECTS
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// --------------------------------------------------------
// object : GlobalSettings
let Tweet = function (id_str, created_at, text, user_id_str, user_name, list_media) {
  this.id_str         = id_str;
  this.created_at_str = Utilities.formatDate(new Date(created_at), TIME_LOCALE, FORMAT_DATETIME);
  this.text           = text;
  this.user_id_str    = user_id_str;
  this.user_name      = user_name;
  this.list_media     = list_media;
};

let Media = function (media_url) {
  this.media_url      = media_url;
};

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
//  OAuth Routines
//
// REFERENCE:
//    https://tech-cci.io/archives/4228
//    https://kazyblog.com/try-google-app-script-for-tweet
//    https://developer.twitter.com/en/docs/twitter-api/v1/tweets/search/api-reference/get-search-tweets
//
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function getOAuthURL() {
  Logger.log(getService().authorize());
}

function getService() {
  //let VAL_CONSUMER_API_KEY     = PropertiesService.getScriptProperties().getProperty("CONSUMER_API_KEY");
  //let VAL_CONSUMER_API_SECRET  = PropertiesService.getScriptProperties().getProperty("CONSUMER_API_SECRET");
  return (
    OAuth1.createService("Twitter")
      .setAccessTokenUrl("https://api.twitter.com/oauth/access_token")
      .setRequestTokenUrl("https://api.twitter.com/oauth/request_token")
      .setAuthorizationUrl("https://api.twitter.com/oauth/authorize")
      .setConsumerKey( VAL_CONSUMER_API_KEY )
      .setConsumerSecret( VAL_CONSUMER_API_SECRET )
      .setCallbackFunction("authCallback")
      .setPropertyStore(PropertiesService.getUserProperties())
  );
}

function resetService() {
  var service = getService();
  service.reset();
}

function authCallback(request) {
  let service = getService();
  let authorized = service.handleCallback(request);
  if (authorized) {
    return HtmlService.createHtmlOutput("success!!");
  } else {
    return HtmlService.createHtmlOutput("failed");
  }
}

// ====================================================================================================================
//
// Debug / Log
//
// ====================================================================================================================

//
// Name: dbgOut
// Desc:
//
function dbgOut(text) {
  if (!g_isDebugMode) {
    return;
  }
  Logger.log(text);
}

//
// Name: logOut
// Desc:
//
function logOut(text) {
  if (!g_isEnabledLogging) {
    return;
  }
  gsAddLineAtBottom(NAME_SHEET_LOG, text);
}

//
// Name: errOut
// Desc:
//
function errOut(text) {
  gsAddLineAtBottom(NAME_SHEET_ERROR, text);
}

// ====================================================================================================================
//
// Basic Twitter Functions
//
// ====================================================================================================================

//
// Func: twGetTimeLine
// Desc:
//    sample func to see how Twitter API works
// Return:
//    none
//
function twGetTimeLine() {
  let twitterService = getService();
  if (twitterService.hasAccess()) {
    let twMethod = { method: "GET" };
    let json = twitterService.fetch(
      "https://api.twitter.com/1.1/statuses/home_timeline.json?count=5",
      twMethod
    );
    let array = JSON.parse(json);
    let tweets = array.map(function (tweet) {
      return tweet.text;
    });
    Logger.log(tweets);
  } else {
    Logger.log(service.getLastError());
  }
}

//
// Func: twSearchTweet
// Desc:
//   search tweet based on the specified keyword
// Return:
//   Tweet object
//
function twSearchTweet(keywords, maxCount = DEFAULT_MAX_NUM_TWEETS) {
  let encodedKeyword = keywords.trim().replace( /\s+/g, '+AND+' );
  encodedKeyword = encodeURIComponent(encodedKeyword);
  try {
    let twitterService = getService();
    if ( ! twitterService.hasAccess()) {
      errOut(twitterService.getLastError());
      return null;
    }
    let url = 'https://api.twitter.com/1.1/search/tweets.json?q='
        + encodedKeyword
        + '&result_type=recent&lang=ja&locale=ja&count='
        + maxCount;
    let response = twitterService.fetch( url, {method:"GET"});
    let json = JSON.parse(response);
    let tweets = json.statuses.map(function (tweet) {

      // INFO: How to get images via Twitter API
      // https://qiita.com/w_cota/items/a87b421ba8bc2b90a938
      list_media = [];
      if (tweet.entities.media != undefined && tweet.entities.media[0].type == 'photo') {
        for(let i=0; i < tweet.extended_entities.media.length; i++) {
          let media_url = tweet.extended_entities.media[i].media_url;
          list_media.push( new Media(media_url));
        }
      }

      return new Tweet(
        tweet.id_str
        , tweet.created_at
        , tweet.text
        , tweet.user.id_str
        , tweet.user.name
        , list_media
      );
    });
    return tweets;
  } catch ( ex ) {
    errOut( ex );
    errOut( "keyword [" + keywords + "] encoded keyword = [" + encodedKeyword + "]" );
  }
  return null;
}

// ====================================================================================================================
//
// Utilities for Misc
//
// ====================================================================================================================

//
// Name: utIsValid
// Desc:
//
function utIsValid(val) {
  return val == null || val == undefined;
}

//
// Name: utGetDateTime
// Desc:
//
function utGetDateTime() {
  return (
    TIME_LOCALE +
    ": " +
    Utilities.formatDate(new Date(), TIME_LOCALE, FORMAT_DATETIME)
  );
}

// ====================================================================================================================
//
// Utilities for Google Spreadsheet
//
// ====================================================================================================================

//
// Name: gsSetCellVal
// Desc:
// NOTE: MUST NOT CALL to update many cells!
//
function gsSetCellVal(sheet, row, col, val) {
  try {
    sheet.getRange(row, col, 1, 1).setValue(val);
  } catch (e) {
    errOut(
      "EXCEPTION: gsSetCellVal: " + e + " - " + sheet.getName() + ":[" + row + ", " + col + "]=[" + val + "]"
    );
  }
  return null;
}

//
// Name: gsGetCellVal
// Desc:
// NOTE: MUST NOT CALL to access many cells!
//
function gsGetCellVal(sheet, row, col) {
  try {
    let range = sheet.getRange(row, col, 1, 1);
    if (range != null && range != undefined) {
      return range.getValue();
    }
  } catch (e) {
    errOut(
      "EXCEPTION: gsGetCellVal: " + e + " - " + sheet.getName() + ":[" + row + ", " + col + "]"
    );
  }
  return null;
}

//
// Name: gsAddLineAtLast
// Desc:
//  Add the specified text at the bootom of the specified sheet.
//
function gsAddLineAtBottom(sheetName, text) {
  try {
    let sheet = g_book.getSheetByName(sheetName);
    if (sheet == null) {
      sheet = g_book.insertSheet(sheetName, g_book.getNumSheets());
    }
    let lastRow = sheet.getLastRow();
    if (lastRow == sheet.getMaxRows()) {
      sheet.insertRowsAfter(lastRow, 1);
    }
    gsSetCellVal(sheet, lastRow + 1, 1, g_datetime);
    gsSetCellVal(sheet, lastRow + 1, 2, String(text));
  } catch (e) {
    Logger.log("EXCEPTION: addLine: " + e.message);
  }
}

//
// Name: gsGetRngVals
// Desc:
//  Get array matrix to cover all values of the specified sheet.
//
function gsGetRngVals(sheet) {
  try {
    let lastRow = sheet.getLastRow();
    let lastCol = sheet.getLastColumn();
    return sheet.getRange(1, 1, lastRow, lastCol).getValues();
  } catch (e) {
    errOut("EXCEPTION: getRngVals: " + e);
  }
  return null;
}

//
// Name: gsClearData
// Desc:
//  Remove all rows from ROW_DATA_START.
//
function gsClearData(sheet) {
  //var lastRow = sheet.getLastRow();
  let lastRow = sheet.getRange("A:A").getLastRow();
  if (CELL_ROW_DATA_START - 1 < lastRow) {
    sheet.deleteRows(CELL_ROW_DATA_START, lastRow - CELL_ROW_DATA_START + 1);
  }
}

// ====================================================================================================================
//
// Google Spreadsheet Routines
//
// ====================================================================================================================

//
// Name: downloadMedia
// Desc:
//  Download media files (image) according to the list of media (list_media).
//
function downloadMedia(folder, list_media) {
  list_media.forEach ( media => {
    console.log( media );
    let imageBlob = UrlFetchApp.fetch(media.media_url).getBlob();
    folder.createFile(imageBlob);
  } );
}

//
// Name: gsAddTweetDataAtBottom
// Desc:
//  Add the specified text at the bootom of the specified sheet.
//
function gsAddTweetDataAtBottom(sheet, tweet) {
  try {
    let lastRow = sheet.getLastRow();
    if (lastRow == sheet.getMaxRows()) {
      sheet.insertRowsAfter(lastRow, 1);
    }
    if (lastRow < CELL_ROW_DATA_START - 1) {
      sheet.insertRowsAfter(lastRow, CELL_ROW_DATA_START - 1 - lastRow);
      lastRow = CELL_ROW_DATA_START - 1;
    }
    let rowToWrite = lastRow + 1;
    let folder = null;
    if ( 0 < tweet.list_media.length ) {
      folder = g_folder.createFolder( tweet.id_str );
      downloadMedia( folder, tweet.list_media );
    }

    gsSetCellVal(sheet, rowToWrite, 1, tweet.id_str);
    gsSetCellVal(sheet, rowToWrite, 2, tweet.created_at_str);
    gsSetCellVal(sheet, rowToWrite, 3, tweet.user_id_str);
    gsSetCellVal(sheet, rowToWrite, 4, tweet.user_name);
    gsSetCellVal(sheet, rowToWrite, 5, tweet.text);
    if ( null != folder ) {
      gsSetCellVal(sheet, rowToWrite, 6, folder.getUrl());
    }

  } catch (e) {
    Logger.log("EXCEPTION: addLine: " + e.message);
  }
}

//
// Name: hasTweet
// Desc:
//  Check if the specified Tweet has already been recorded (true) or not (false).
//
function hasTweet(range, tweet) {
  for (var r = 0; r < range.length; r++) {
    let row = range[r];
    if (0 == row[0].length) {
      break;
    }
    if (row[0] == tweet.id_str) {
      return true;
    }
  }
  return false;
}

//
// Name: updateSheet
// Desc:
//  Add a new row and write the new Tweet data.
//
function updateSheet(sheet, tweets) {
  let lastRow = sheet.getLastRow();
  let lastCol = sheet.getLastColumn();
  let range = sheet
    .getRange(CELL_ROW_DATA_START, 1, lastRow, lastCol)
    .getValues();
  for ( let i = tweets.length - 1; i >= 0; i--) {
    if ( ! hasTweet(range, tweets[i])) {
      gsAddTweetDataAtBottom(sheet, tweets[i]);
    }
  }
}

function main() {
  g_folder = DriveApp.getFolderById(VAL_ID_GDRIVE_FOLDER);
  g_book = SpreadsheetApp.openById(VAL_ID_TARGET_BOOK);
  let sheets = g_book.getSheets();

  sheets.forEach((sheet) => {
    let sheetName = sheet.getName();
    if (sheetName.startsWith("!")) {
      return;
    }
    let rngVals = gsGetRngVals(sheet);
    if (rngVals == null) {
      return;
    }
    let currentKeyword = gsGetCellVal(sheet, CELL_ROW_KEYWORD, 1);
    currentKeyword = currentKeyword == null ? "" : currentKeyword.trim();
    let lastKeyword = gsGetCellVal(sheet, CELL_ROW_LAST_KEYWORD, 2);
    lastKeyword = lastKeyword == null ? "" : lastKeyword.trim();
    gsSetCellVal(sheet, CELL_ROW_KEYWORD, 1, currentKeyword);
    gsSetCellVal(sheet, CELL_ROW_LAST_KEYWORD, 1, CELL_HEADER_LAST_KEYWORD);
    gsSetCellVal(sheet, CELL_ROW_LAST_KEYWORD, 2, currentKeyword);
    gsSetCellVal(sheet, CELL_ROW_LAST_UPDATED, 1, CELL_HEADER_LAST_UPDATED);
    gsSetCellVal(sheet, CELL_ROW_LAST_UPDATED, 2, g_datetime);

    CELL_HEADER_TITLES.forEach( function (value, index) {
      gsSetCellVal(sheet, CELL_ROW_HEADER , index + 1, value );
    });
    if ("" == currentKeyword || lastKeyword != currentKeyword) {
      gsClearData(sheet);
    }
    if ("" == currentKeyword) {
      return;
    }
    var tweets = twSearchTweet(currentKeyword);
    if (null != tweets && 0 < tweets.length) {
      updateSheet(sheet, tweets);
    }
  });
}