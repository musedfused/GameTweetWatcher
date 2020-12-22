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
//  2020-11-15 : Supported the following features
//                * downloading images
//                * ban words, ban users (screen_names)
//                * create link to each tweet
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
let VERSION                       = 1.0;
let TIME_LOCALE                   = "JST";
let FORMAT_DATETIME               = "yyyy-MM-dd (HH:mm:ss)";
let FORMAT_TIMESTAMP              = "yyyyMMddHHmmss";
let NAME_SHEET_USAGE              = "!USAGE";
let NAME_SHEET_LOG                = "!LOG";
let NAME_SHEET_ERROR              = "!ERROR";

let SHEET_NAME_COMMON_SETTINGS    = "%settings";

let CELL_HEADER_LIMIT             = "[limitcounttweets]";
let CELL_HEADER_EMAIL             = "[email]";

let CELL_HEADER_KEYWORD           = "[keyword]";
let CELL_HEADER_BAN_WORDS         = "[banwords]";
let CELL_HEADER_BAN_USERS         = "[banusers]";

let CELL_HEADER_POINT_TWEET       = "[point/tweet]";
let CELL_HEADER_POINT_RETWEET     = "[point/retweet]";
let CELL_HEADER_POINT_REPLY      = "[point/reply]";
let CELL_HEADER_POINT_MEDIA       = "[point/media]";
let CELL_HEADER_POINT_NICE        = "[point/nice]";
let CELL_HEADER_ALERT_THRESHOLD   = "[alertthresholdpoints]";

let CELL_HEADER_LAST_KEYWORD      = "[[lastkeyword]]";
let CELL_HEADER_LAST_UPDATED      = "[[lastupdated]]";

let CELL_HEADER_TITLES            = ["tweetid", "createdat", "userid", "username", "tweet", "screenname", "media"];

let DEFAULT_LIMIT_TWEETS          = 10;
let DEFAULT_POINT_TWEET           = 10;
let DEFAULT_POINT_RETWEET         = 10;
let DEFAULT_POINT_REPLY           = 10;
let DEFAULT_POINT_MEDIA           = 10;
let DEFAULT_POINT_NICE            = 10;
let DEFAULT_ALERT_THRESHOLD       = 100;

let MAX_ROW_RANGE_SETTINGS        = 20;
let MAX_COLUMN_RANGE_SETTINGS     = 30;

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// GLOBALS
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
let g_isDebugMode                 = true;
let g_isEnabledLogging            = true;
let g_datetime                    = null;
let g_book                        = null;
let g_folder                      = null;
let g_settingsCommon              = null;

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// OBJECTS
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// --------------------------------------------------------
// object : GlobalSettings
let Tweet = function (id_str, created_at, text, user_id_str, user_name, screen_name, list_media) {
  this.id_str         = id_str;
  this.created_at_str = Utilities.formatDate(new Date(created_at), TIME_LOCALE, FORMAT_DATETIME);
  this.text           = text;
  this.user_id_str    = user_id_str;
  this.user_name      = user_name;
  this.screen_name    = screen_name;
  this.list_media     = list_media;
};

let Media = function (media_url) {
  this.media_url      = media_url;
};

//
// Settings - This object will be used for as Common settings and as Particular query settings
//
let Settings = function (currentKeyword, lastKeyword, limitTweets, emails, banWords, banUsers, ptTweet, ptRetweet, ptReply, ptMedia, ptNice, ptAlertThreshold, rowDataStart) {
  this.currentKeyword   = currentKeyword;
  this.lastKeyword      = lastKeyword;
  this.limitTweets      = limitTweets;
  this.emails           = emails;
  this.banWords         = banWords;
  this.banUsers         = banUsers;
  this.ptTweet          = ptTweet;
  this.ptRetweet        = ptRetweet;
  this.ptReply          = ptReply;
  this.ptMedia          = ptMedia;
  this.ptNice           = ptNice;
  this.ptAlertThreshold = ptAlertThreshold;
  this.rowDataStart     = rowDataStart;
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
//   'Tweet' object
//
function twSearchTweet(keywords, maxCount = DEFAULT_LIMIT_TWEETS) {
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
      let list_media = [];
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
        , tweet.user.screen_name
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
    let range = sheet.getRange(lastRow+1, 1, 1, 2);
    let rngVals = range.getValues();
    let row = rngVals[0];
    row[0] = g_datetime;
    row[1] = String(text);
    range.setValues( rngVals );
  } catch (e) {
    Logger.log("EXCEPTION: addLine: " + e.message);
  }
}

//
// Name: gsGetRange
// Desc:
//  Get array matrix to cover all values of the specified sheet.
//
function gsGetRange(sheet, row = 1, column = 1, numRows = 0, numColumns = 0) {
  try {
    let numRowsTotal = (0 < numRows )? numRows : (sheet.getLastRow() - (row-1));
    let numColumnsTotal = (0 < numColumns)? numColumns : (sheet.getLastColumn() -(column-1));
    return sheet.getRange(row, column, numRowsTotal, numColumnsTotal);
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
function gsClearData(sheet, rowDataStart) {
  let lastRow = sheet.getRange("A:A").getLastRow();
  if (rowDataStart - 1 < lastRow) {
    sheet.deleteRows(rowDataStart, lastRow - rowDataStart + 1);
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
    // console.log( media );
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
    let rowToWrite = lastRow + 1;
    let folder = null;
    if ( 0 < tweet.list_media.length ) {
      folder = g_folder.createFolder( tweet.id_str );
      downloadMedia( folder, tweet.list_media );
    }
    let range = gsGetRange( sheet, rowToWrite, 1, 1, 8 );
    let rngVals = range.getValues();
    let row = rngVals[0];
    row[0] = '=HYPERLINK("https://twitter.com/' + tweet.screen_name + '/status/' + tweet.id_str + '", "' + tweet.id_str + '")';
    row[1] = tweet.created_at_str;
    row[2] = tweet.user_id_str;
    row[3] = tweet.user_name;
    row[4] = tweet.text;
    row[5] = tweet.screen_name;
    if ( null != folder ) {
      row[6] = folder.getUrl();
    }
    range.setValues( rngVals );
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
// Name: hasBanWords
// Desc:
//  Check if the specified tweet contains any ban words in the specified array of the ban words.
//
function hasBanWords ( tweet, banWords ){
  for ( let i=0; i< banWords.length; i++ ) {
    //let regex = new RegExp(banWords[i]);
    //if ( -1 != tweet.text.search(regex) ) {
    if ( tweet.text.includes(banWords[i]) ) {
      return true;
    }
  }
  return false;
}

//
// Name: isFromBanUsers
// Desc:
//  Check if the specified tweet contains any ban words in the specified array of the ban words.
//
function isFromBanUsers ( tweet, banUsers ){
  for ( let i=0; i< banUsers.length; i++ ) {
    //if ( tweet.screen_name == banUsers[i] ) {
    if ( tweet.user_id_str == banUsers[i] ) {
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
function updateSheet(sheet, settings, tweets) {
  let lastRow = sheet.getLastRow();
  let lastCol = sheet.getLastColumn();
  let range = sheet.getRange(settings.rowDataStart, 1, lastRow, lastCol).getValues();
  for ( let i = tweets.length - 1; i >= 0; i--) {
    if ( ! hasTweet(range, tweets[i])) {
      if ( ! hasBanWords ( tweets[i], settings.banWords ) && ! isFromBanUsers ( tweets[i], settings.banUsers ) ) {
      gsAddTweetDataAtBottom(sheet, tweets[i]);
      }
    }
  }
}

//
// Name: getSettings
// Desc:
//    Get settings from the common setting sheet and each query sheet.
//    This function supports both the Common Setting Sheet and settings for each query.
//
function getSettings (sheet) {
  let range = gsGetRange(sheet, 1, 1, MAX_ROW_RANGE_SETTINGS, MAX_COLUMN_RANGE_SETTINGS);
  if (range == null) {
    return;
  }

  let currentKeyword   = "";
  let lastKeyword      = "";
  let limitTweets      = DEFAULT_LIMIT_TWEETS;
  let emails           = [];
  let banWords         = [];
  let banUsers         = [];
  let ptTweet          = 0;
  let ptRetweet        = 0;
  let ptReply          = 0;
  let ptMedia          = 0;
  let ptNice           = 0;
  let ptAlertThreshold = 0;
  let rowDataStart     = 0;

  let rngVals = range.getValues();
  rngVals.forEach( (row, idxRow) => {
    let title:string = String(row[0]);
    if( !title ) {

    } else {
      title = title.replace(/\s+/g, '').toLowerCase().trim();
      switch (title) {
      case CELL_HEADER_LIMIT:
        if ( row[1] ) {
          limitTweets = Number( row[1] );
        }
        break;

      case CELL_HEADER_KEYWORD:
        if ( row[1] ) {
          currentKeyword = String(row[1]).trim();
          row[1] = currentKeyword;
        }
        break;

      case CELL_HEADER_EMAIL:
        for(let i=1; i<row.length; i++) {
          let val:string = String(row[i]).trim();
          if ( !val ) break;
          emails.push( val );
        }
        break;

      case CELL_HEADER_BAN_WORDS:
        for(let i=1; i<row.length; i++) {
          let val:string = String(row[i]).trim();
          if ( !val ) break;
          banWords.push( val );
        }
        break;

      case CELL_HEADER_BAN_USERS:
        for(let i=1; i<row.length; i++) {
          let val:string = String(row[i]).trim();
          if ( !val ) break;
          banUsers.push( val );
        }
        break;

      case CELL_HEADER_POINT_TWEET:
        {
          let val:string = String(row[1]).trim();
            if ( !val ) break;
          ptTweet = Number( val );
        }
        break;

      case CELL_HEADER_POINT_RETWEET:
        {
          let val:string = String(row[1]).trim();
          if ( !val ) break;
          ptRetweet = Number( val );
        }
        break;

      case CELL_HEADER_POINT_TWEET:
        {
          let val:string = String(row[1]).trim();
          if ( !val ) break;
          ptReply = Number( val );
        }
        break;

      case CELL_HEADER_POINT_MEDIA:
        {
          let val:string = String(row[1]).trim();
          if ( !val ) break;
          ptMedia = Number( val );
        }
        break;

      case CELL_HEADER_POINT_NICE:
        {
          let val:string = String(row[1]).trim();
          if ( !val ) break;
          ptNice = Number( val );
        }
        break;

      case CELL_HEADER_ALERT_THRESHOLD:
        {
          let val:string = String(row[1]).trim();
          if ( !val ) break;
          ptAlertThreshold = Number( val );
        }
        break;

      case CELL_HEADER_LAST_KEYWORD:
        {
          let val:string = String(row[1]);
          if ( !val ) break;
          lastKeyword = val;
          row[1] = currentKeyword;
        }
        break;

      case CELL_HEADER_LAST_UPDATED:
        {
          row[1] = g_datetime;
        }
        break;

      case CELL_HEADER_TITLES[0]:
        rowDataStart = idxRow + 2;
        break;
      }
    }
  });
  range.setValues( rngVals );
  return new Settings (currentKeyword, lastKeyword, limitTweets, emails, banWords, banUsers, ptTweet, ptRetweet, ptReply, ptMedia, ptNice, ptAlertThreshold, rowDataStart);
}

//
// Name: getSettingsActual
// Desc:
//  Returns the actual settings which cover the common settings and local settings for each sheet.
//
function getSettingsActual( settingsCommon, settingsLocal ) {
  return new Settings (
    settingsLocal.currentKeyword,
    settingsLocal.lastKeyword,
    settingsCommon.limitTweets,
    settingsCommon.emails.concat(settingsLocal.emails)     ,
    settingsCommon.banWords.concat(settingsLocal.banWords) ,
    settingsCommon.banUsers.concat(settingsLocal.banUsers) ,
    (0<settingsLocal.ptTweet)          ? settingsLocal.ptTweet          : settingsCommon.ptTweet          ,
    (0<settingsLocal.ptRetweet)        ? settingsLocal.ptRetweet        : settingsCommon.ptRetweet        ,
    (0<settingsLocal.ptReply)          ? settingsLocal.ptReply          : settingsCommon.ptReply          ,
    (0<settingsLocal.ptMedia)          ? settingsLocal.ptMedia          : settingsCommon.ptMedia          ,
    (0<settingsLocal.ptNice)           ? settingsLocal.ptNice           : settingsCommon.ptNice           ,
    (0<settingsLocal.ptAlertThreshold) ? settingsLocal.ptAlertThreshold : settingsCommon.ptAlertThreshold ,
    settingsLocal.rowDataStart);
}

//
// Name: main
// Desc:
//  Entry point of this program.
//
function main() {
  g_datetime = TIME_LOCALE + ": " + Utilities.formatDate(new Date(), TIME_LOCALE, FORMAT_DATETIME);
  g_folder = DriveApp.getFolderById(VAL_ID_GDRIVE_FOLDER);
  g_book = SpreadsheetApp.openById(VAL_ID_TARGET_BOOK);

  let sheets = g_book.getSheets();

  sheets.forEach((sheet) => {
    let sheetName:string = sheet.getName();
    if (sheetName.toLocaleLowerCase().trim().startsWith('!')) {
      return;
    }
    if (sheetName.toLocaleLowerCase().trim() === SHEET_NAME_COMMON_SETTINGS) {
      g_settingsCommon = getSettings( sheet );
      return;
    }
    let settingsLocal = getSettings( sheet );
    if ( 0 == settingsLocal.rowDataStart ) {
      errOut("Unexpected sheet was found: " + sheetName)
      return;
    }
    if ( !settingsLocal.currentKeyword || settingsLocal.lastKeyword != settingsLocal.currentKeyword ) {
      gsClearData(sheet, settingsLocal.rowDataStart);
    }
    if ( !settingsLocal.currentKeyword ) {
      return;
    }
    let settingsActual = getSettingsActual( g_settingsCommon, settingsLocal);
    var tweets = twSearchTweet(settingsActual.currentKeyword);
    if (null != tweets && 0 < tweets.length) {
      updateSheet(sheet, settingsActual, tweets);
    }
  });
}