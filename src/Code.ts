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
//  2021-02-18 : Supported storing history data
//  2021-02-17 : Supported backup feature as expected by the team
//  2021-01-31 : Supported backup feature
//  2021-01-24 : Supported the following features
//                * loading common & local settings
//                * point calculation
//                * sending notification email
//                * ISO-8601 folder for placing images
//                * backup function
//  2020-11-15 : Supported the following features
//                * downloading images
//                * ban words, ban users
//                * create link to each tweet
//  2020-10-19 : Initial version
//
// ====================================================================================================================

// ID of the Target Google Spreadsheet (Book)
const VAL_ID_TARGET_BOOK              = '18PpIEIWru0_z46FkWRt9Ac3-IEGcEhqh8BWTEC5U6i4';
// ID of the Google Drive where the images will be placed
const VAL_ID_GDRIVE_FOLDER_MEDIA      = '1ttFnVjcZJNJdloaT4ni99ZbEYuEj-WAA';
// ID of the Google Drive where backup files will be placed
const VAL_ID_GDRIVE_FOLDER_BACKUP     = '1BVPS0fXT0UzkvuIe0gDESkhHApcyxZsR';
// ID of the Google Drive where history files will be placed
const VAL_ID_GDRIVE_FOLDER_HISTORY    = '18quvxDih5xm59htcaPRtuzeDAu5fPmt0';
// Key and Secret to access Twitter APIs
const VAL_CONSUMER_API_KEY            = 'lyhQXM3aZP5HHLqO6jjcEwzux';
const VAL_CONSUMER_API_SECRET         = 'ODBAEfj4VlgVf1BKHJyoz6QwryiaNtcchkr4PikzjSHmct0hjr';

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// DEFINES
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
const VERSION                               = 1.1;
const DURATION_MILLISEC_NOT_FOR_BACKUP      = 7 * 24 * 60 * 60 * 1000;
const TIME_LOCALE                           = "JST";
const FORMAT_DATETIME_DATE                  = "yyyy-MM-dd";
const FORMAT_DATETIME_ISO8601_DATE          = "yyyy-MM-dd";
const FORMAT_DATETIME_ISO8601_TIME          = "HH:mm:ss";
const FORMAT_DATETIME_DATE_NUM              = "yyyyMMdd";
const FORMAT_DATETIME                       = "yyyy-MM-dd (HH:mm:ss)";
const FORMAT_TIMESTAMP                      = "yyyyMMddHHmmss";
const NAME_SHEET_USAGE                      = "!USAGE";
const NAME_SHEET_LOG                        = "!LOG";
const NAME_SHEET_ERROR                      = "!ERROR";

const SHEET_NAME_COMMON_SETTINGS            = "%settings";

const CELL_SETTINGS_TITLE_LIMIT             = "[limitcounttweets]";
const CELL_SETTINGS_TITLE_EMAIL             = "[email]";

const CELL_SETTINGS_TITLE_KEYWORD           = "[keyword]";
const CELL_SETTINGS_TITLE_BAN_WORDS         = "[banwords]";
const CELL_SETTINGS_TITLE_BAN_USERS         = "[banusers]";

const CELL_SETTINGS_TITLE_DOWNLOAD_MEDIA    = "[downloadmedia]";
const CELL_SETTINGS_TITLE_POINT_TWEET       = "[point/tweet]";
const CELL_SETTINGS_TITLE_POINT_RETWEET     = "[point/retweet]";
const CELL_SETTINGS_TITLE_POINT_REPLY       = "[point/reply]";
const CELL_SETTINGS_TITLE_POINT_MEDIA       = "[point/media]";
const CELL_SETTINGS_TITLE_POINT_FAVORITE    = "[point/nice]";
const CELL_SETTINGS_TITLE_ALERT_THRESHOLD   = "[alertthresholdpoints]";

const CELL_SETTINGS_TITLE_LAST_KEYWORD      = "[[lastkeyword]]";
const CELL_SETTINGS_TITLE_LAST_UPDATED      = "[[lastupdated]]";
const CELL_SETTINGS_TITLE_END_SETTINGS      = "//endofsettings";

const CELL_HEADER_TITLES_DATA               = ["Tweet Id", "Created at", "User Id", "User name", "Screen name", "Retweet", "Favorite", "Tweet", "Media"];
const CELL_HEADER_TITLES_HISTORY            = ["date time", "keyword", "pt/Tweet", "pt/Retweet", "pt/Media", "pt/Fav", "Threshold", "", "Total", "Tweets", "Retweets", "Media", "Favs"];

const DEFAULT_LIMIT_TWEETS          = 10;
const DEFAULT_DOWNLOAD_MEDIA        = true;
const DEFAULT_POINT_TWEET           = 10;
const DEFAULT_POINT_RETWEET         = 10;
const DEFAULT_POINT_REPLY           = 10;
const DEFAULT_POINT_MEDIA           = 10;
const DEFAULT_POINT_FAVORITE        = 10;
const DEFAULT_ALERT_THRESHOLD       = 100;

const MAX_ROW_RANGE_SETTINGS        = 20;
const MAX_COLUMN_RANGE_SETTINGS     = 30;

const MAX_ROW_SEEK_HEADER           = 20;

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// GLOBALS
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
let g_isDebugMode                 = true;
let g_isEnabledLogging            = true;
let g_settingsCommon              = null;

let g_folderBackup                = DriveApp.getFolderById(VAL_ID_GDRIVE_FOLDER_BACKUP);
let g_datetime                    = new Date();
let g_timestamp                   = TIME_LOCALE + ": " + Utilities.formatDate(g_datetime, TIME_LOCALE, FORMAT_DATETIME);
let g_folderMedia                 = DriveApp.getFolderById(VAL_ID_GDRIVE_FOLDER_MEDIA);
let g_folderHistory               = DriveApp.getFolderById(VAL_ID_GDRIVE_FOLDER_HISTORY);
let g_book                        = SpreadsheetApp.openById(VAL_ID_TARGET_BOOK);

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// OBJECTS
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

//
// Tweet
//
class Tweet {
  id_str         : string;
  created_at     : string;
  created_at_str : string;
  text           : string;
  user_id_str    : string;
  user_name      : string;
  screen_name    : string;
  list_media     : Media[];
  retweet_count  : number;
  favorite_count : number;

  constructor (
    id_str         : string
    , created_at     : string
    , text           : string
    , user_id_str    : string
    , user_name      : string
    , screen_name    : string
    , list_media     : Media[]
    , retweet_count  : number
    , favorite_count : number) {
      this.id_str         = id_str;
      this.created_at     = created_at;
      this.created_at_str = Utilities.formatDate(new Date(created_at), TIME_LOCALE, FORMAT_DATETIME);
      this.text           = text;
      this.user_id_str    = user_id_str;
      this.user_name      = user_name;
      this.screen_name    = screen_name;
      this.list_media     = list_media;
      this.retweet_count  = retweet_count;
      this.favorite_count = favorite_count;
  }
}

//
// Media
//
class Media {
  media_url : string;

  constructor (media_url) {
    this.media_url      = media_url;
  }
}

//
// Settings - This object will be used for as Common settings and as Particular query settings
//
class Settings {
  currentKeyword   : string   ;
  lastKeyword      : string   ;
  limitTweets      : number   ;
  emails           : string[] ;
  banWords         : string[] ;
  banUsers         : string[] ;
  bDownloadMedia   : boolean  ;
  ptTweet          : number   ;
  ptRetweet        : number   ;
  ptReply          : number   ;
  ptMedia          : number   ;
  ptFavorite       : number   ;
  ptAlertThreshold : number   ;
  rowEndSettings   : number   ;
  headerInfo       : HeaderInfo ;

  constructor (
    currentKeyword   : string
    , lastKeyword      : string
    , limitTweets      : number
    , emails           : string[]
    , banWords         : string[]
    , banUsers         : string[]
    , bDownloadMedia   : boolean
    , ptTweet          : number
    , ptRetweet        : number
    , ptReply          : number
    , ptMedia          : number
    , ptFavorite       : number
    , ptAlertThreshold : number
    , rowEndSettings   : number
    , headerInfo       : HeaderInfo ) {
      this.currentKeyword   = currentKeyword     ;
      this.lastKeyword      = lastKeyword        ;
      this.limitTweets      = limitTweets        ;
      this.emails           = emails             ;
      this.banWords         = banWords           ;
      this.banUsers         = banUsers           ;
      this.bDownloadMedia   = bDownloadMedia     ;
      this.ptTweet          = ptTweet            ;
      this.ptRetweet        = ptRetweet          ;
      this.ptReply          = ptReply            ;
      this.ptMedia          = ptMedia            ;
      this.ptFavorite       = ptFavorite         ;
      this.ptAlertThreshold = ptAlertThreshold   ;
      this.rowEndSettings   = rowEndSettings     ;
      this.headerInfo       = headerInfo         ;
  }
}

//
// Header
//
class HeaderInfo {
  idx_row        : number ;
  idx_tweetId    : number ;
  idx_createdAt  : number ;
  idx_userId     : number ;
  idx_userName   : number ;
  idx_screenName : number ;
  idx_tweet      : number ;
  idx_retweet    : number ;
  idx_favorite   : number ;
  idx_media      : number ;

  constructor () {
    this.idx_row          = null;
    this.idx_tweetId      = null;
    this.idx_createdAt    = null;
    this.idx_userId       = null;
    this.idx_userName     = null;
    this.idx_screenName   = null;
    this.idx_tweet        = null;
    this.idx_retweet      = null;
    this.idx_favorite     = null;
    this.idx_media        = null;
  }
}

class Stats {
  countTweets    : number;
  countRetweets  : number;
  countMedias    : number;
  countFavorites : number;

  constructor() {
    this.countTweets    = 0;
    this.countRetweets  = 0;
    this.countMedias    = 0;
    this.countFavorites = 0;
  }
}

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
function logOut(text:string) {
  text = g_timestamp + "\t" + text;
  if (!g_isEnabledLogging) {
    return;
  }
  gsAddLineAtBottom(NAME_SHEET_LOG, text);
}

//
// Name: errOut
// Desc:
//
function errOut(text:string) {
  text = g_timestamp + "\t" + text;
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
    let tweets = [];
    let json = JSON.parse(response);
    json.statuses.forEach(tweet => {
      if ( ! tweet.retweeted_status ) {
        // INFO: How to get images via Twitter API
        // https://qiita.com/w_cota/items/a87b421ba8bc2b90a938
        let list_media = [];
        if (tweet.entities.media != undefined && tweet.entities.media[0].type == 'photo') {
          for(let i=0; i < tweet.extended_entities.media.length; i++) {
            let media_url = tweet.extended_entities.media[i].media_url;
            list_media.push( new Media(media_url));
          }
        }
        tweets.push( new Tweet(
          tweet.id_str
          , tweet.created_at
          , tweet.text
          , tweet.user.id_str
          , tweet.user.name
          , tweet.user.screen_name
          , list_media
          , tweet.retweet_count
          , tweet.favorite_count
        ) );
      }
    });
    return tweets;
  } catch ( ex ) {
    errOut( ex );
    errOut( "keyword [" + keywords + "] encoded keyword = [" + encodedKeyword + "]" );
    return null;
  }
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
    if ( !sheet ) {
      sheet = g_book.insertSheet(sheetName, g_book.getNumSheets());
    }
    let lastRow = sheet.getLastRow();
    if (lastRow == sheet.getMaxRows()) {
      sheet.insertRowsAfter(lastRow, 1);
    }
    let range = sheet.getRange(lastRow+1, 1, 1, 2);
    if ( range ){
      let valsRange = range.getValues();
      let row = valsRange[0];
      row[0] = g_timestamp;
      row[1] = String(text);
      range.setValues( valsRange );
    }
  } catch (e) {
    Logger.log("EXCEPTION: gsAddLineAtBottom: " + e.message);
  }
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
// Name: sendMail
// Desc:
//
//
function sendMail(sheet, stats:Stats, settings:Settings) {
  let points = getPoint(stats, settings);
  let subject = "⚡AutoMail:TweetWatcher⚡: [" + sheet.getName() + "]:[" + settings.currentKeyword + "]:" + points;
  let body = "Total Points : " + points;
  let htmlBody = "Total Points : "
    + getPoint(stats, settings) + "<br>"
    + "Threshold: " + settings.ptAlertThreshold + "<br>"
    + "<br>"
    + "<a href='https://docs.google.com/spreadsheets/d/" + VAL_ID_TARGET_BOOK + "/edit'>&lt;Link to the Spreadsheet&gt;</a><br>"
    + "Sheet name: " + sheet.getName() + "<br>"
    + "Keyword: " + settings.currentKeyword + "<br>"
    + "<br>"
    + "Stats:<br>"
    + "<table>"
    + "<tr><th>Item</th><th>Count</th><th>Pt/Item</th><th>Points</th></tr>"
    + "<tr><td>Tweet</td><td>"    + stats.countTweets    + "</td><td>" + settings.ptTweet    + "</td><td>" + (stats.countTweets * settings.ptTweet)       + "</td></tr>"
    + "<tr><td>Retweet</td><td>"  + stats.countRetweets  + "</td><td>" + settings.ptRetweet  + "</td><td>" + (stats.countRetweets * settings.ptRetweet)   + "</td></tr>"
    + "<tr><td>Favorite</td><td>" + stats.countFavorites + "</td><td>" + settings.ptFavorite + "</td><td>" + (stats.countFavorites * settings.ptFavorite) + "</td></tr>"
    + "<tr><td>Media</td><td>"    + stats.countMedias    + "</td><td>" + settings.ptMedia    + "</td><td>" + (stats.countMedias * settings.ptMedia)       + "</td></tr>"
    + "</table>";

  settings.emails.forEach(emailAddr => {
    GmailApp.sendEmail(emailAddr, subject, body, {htmlBody: htmlBody} );
  });
}

//
// Name: getPoint
// Desc:
//
//
function getPoint(stats:Stats, settings:Settings):number {
  return settings.ptTweet * stats.countTweets
    + settings.ptRetweet * stats.countRetweets
    + settings.ptMedia * stats.countMedias
    + settings.ptFavorite * stats.countFavorites;
}

//
// Name: insertHeaderData
// Desc:
//  Insert the header row at the specified row (1-based)
//
function inseartHeaderData(sheet, row) {
  let rangeHeader = sheet.getRange(row, 1, 1, CELL_HEADER_TITLES_DATA.length);
  let valsHeader = rangeHeader.getValues();
  CELL_HEADER_TITLES_DATA.forEach( (val, c) =>{
    valsHeader[0][c] = val;
  } );
  rangeHeader.setValues(valsHeader);
}

//
// Name: insertHeaderHistory
// Desc:
//  Insert the header row at the specified row (1-based)
//
function inseartHeaderHistory(sheet, row) {
  let rangeHeader = sheet.getRange(row, 1, 1, CELL_HEADER_TITLES_HISTORY.length);
  let valsHeader = rangeHeader.getValues();
  CELL_HEADER_TITLES_HISTORY.forEach( (val, c) =>{
    valsHeader[0][c] = val;
  } );
  rangeHeader.setValues(valsHeader);
}

//
// Name: getSheet
// Desc:
//  Get access to the specified sheet in the specified book, and if the book and the sheet don't exist, create them accordingly.
//  This function is used for backup feature and storing history data.
//
function getSheet( folderParent, pathTarget:string, nameSheet:string, indexSheet:number, drawHeaderFunc ) {
  let pathSplit = pathTarget.split("/");
  let listPath = [];
  let nameFile = null;
  for( let i=0 ; i<pathSplit.length; i++) {
    if (pathSplit[i].length > 0 ) {
      if ( i == pathSplit.length -1) {
        nameFile = pathSplit[i];
        break;
      }
      listPath.push(pathSplit[i]);
    }
  }
  if ( null == nameFile ) {
    errOut("getFile() - wrong path name - " + pathTarget );
    return null;
  }

  let folderTarget = folderParent;
  for(let i = 0; i<listPath.length; i++){
    if( folderTarget.getFoldersByName(listPath[i]).hasNext()){
      folderTarget = folderTarget.getFoldersByName(listPath[i]).next();
    } else {
      folderTarget = folderTarget.createFolder( listPath[i] );
    }
  }
  let fileTarget = folderTarget.getFilesByName(nameFile)
  let book = null;
  let sheet = null;
  if (fileTarget && fileTarget.hasNext()) {
    let file = fileTarget.next();
    book = SpreadsheetApp.openById(file.getId());
    sheet = book.getSheetByName(nameSheet);
    if ( !sheet ) {
      sheet = book.insertSheet(nameSheet, indexSheet);
      drawHeaderFunc( sheet, 1);
    }
  } else {
    book = SpreadsheetApp.create(nameFile);
    book.getActiveSheet().setName(nameSheet);
    let idBook = DriveApp.getFileById(book.getId());
    folderTarget.addFile(idBook);
    sheet = book.getActiveSheet();
    drawHeaderFunc( sheet, 1);
  }
  return sheet;
}

//
// Name: addHistory
// Desc:
//
//
function addHistory( nameSheet:string, indexSheet:number, stats:Stats, settings:Settings ) {
  let fullYear:number = g_datetime.getFullYear();
  let month:number = 1 + g_datetime.getMonth();
  let nameBook:string = g_book.getName() + "/" + String(fullYear)+ "_" + ('00'+month).slice(-2);
  let sheetHistory = getSheet(g_folderHistory, nameBook, nameSheet, indexSheet, inseartHeaderHistory);
  let lastRow = sheetHistory.getLastRow();
  let rowToWrite = lastRow + 1;
  if ( rowToWrite > sheetHistory.getMaxRows() ) {
    sheetHistory.insertRowsAfter(lastRow, 1);
  }

  let rangeLastRow = sheetHistory.getRange( rowToWrite, 1, 1, CELL_HEADER_TITLES_HISTORY.length);
  if (rangeLastRow) {
    let valsRange = rangeLastRow.getValues();
    valsRange[0][0] = g_timestamp;
    valsRange[0][1] = settings.currentKeyword;
    valsRange[0][2] = settings.ptTweet;
    valsRange[0][3] = settings.ptRetweet;
    valsRange[0][4] = settings.ptMedia;
    valsRange[0][5] = settings.ptFavorite;
    valsRange[0][6] = settings.ptAlertThreshold;
    valsRange[0][7] = "";
    valsRange[0][8] = getPoint(stats, settings);;
    valsRange[0][9] = stats.countTweets;
    valsRange[0][10] = stats.countRetweets;
    valsRange[0][11] = stats.countMedias;
    valsRange[0][12] = stats.countFavorites;
    rangeLastRow.setValues(valsRange);
  }
}

//
// Name: downloadMedia
// Desc:
//  Download media used in a tweet in the date folder.
//
function downloadMedia( tweet:Tweet, dateCreatedAt:Date, settings:Settings) {
  let folderMedia = null;
  if ( 0 < tweet.list_media.length && settings.bDownloadMedia ) {
    let strDate = Utilities.formatDate(dateCreatedAt, TIME_LOCALE, FORMAT_DATETIME_DATE);
    let foldersOfDate = g_folderMedia.getFoldersByName(strDate);
    let folderDate = null;
    if (foldersOfDate.hasNext()) {
      folderDate = foldersOfDate.next();
    }else {
      folderDate = g_folderMedia.createFolder( strDate );
    }
    folderMedia = folderDate.createFolder( tweet.id_str );

    tweet.list_media.forEach ( media => {
      // console.log( media );
      let imageBlob = UrlFetchApp.fetch(media.media_url).getBlob();
      folderMedia.createFile(imageBlob);
    } );
  }
  return folderMedia;
}

//
// Name: addTweetDataAtBottom
// Desc:
//  Add the specified text at the bootom of the specified sheet.
//
function addTweetDataAtBottom(sheet, indexSheet:number, tweet:Tweet, settings:Settings, stats:Stats ) {
  try {
    let lastRow = sheet.getLastRow();
    let rowToWrite = lastRow + 1;
    if ( rowToWrite > sheet.getMaxRows() ) {
      sheet.insertRowsAfter(lastRow, 1);
    }

    let dateCreatedAt = new Date(tweet.created_at);
    let folderMedia = downloadMedia( tweet, dateCreatedAt, settings);

    let rangeLastRow = sheet.getRange( rowToWrite, 1, 1, CELL_HEADER_TITLES_DATA.length);
    if (rangeLastRow) {
      let valsRange = rangeLastRow.getValues();
      let row = valsRange[0];

      row[settings.headerInfo.idx_tweetId    ] = '=HYPERLINK("https://twitter.com/' + tweet.screen_name + '/status/' + tweet.id_str + '", "' + tweet.id_str + '")';
      row[settings.headerInfo.idx_createdAt  ] = tweet.created_at_str;
      row[settings.headerInfo.idx_userId     ] = tweet.user_id_str;
      row[settings.headerInfo.idx_userName   ] = tweet.user_name;
      row[settings.headerInfo.idx_screenName ] = tweet.screen_name;
      row[settings.headerInfo.idx_retweet    ] = tweet.retweet_count;
      row[settings.headerInfo.idx_favorite   ] = tweet.favorite_count;
      row[settings.headerInfo.idx_tweet      ] = tweet.text;

      if ( folderMedia ) {
        row[settings.headerInfo.idx_media] = folderMedia.getUrl();
      }
      rangeLastRow.setValues( valsRange );

      stats.countTweets ++;
      stats.countMedias += (tweet.list_media.length > 0 ) ? 1 : 0;
      stats.countRetweets += tweet.retweet_count;
      stats.countFavorites += tweet.favorite_count;
    }
  } catch (e) {
    Logger.log("EXCEPTION: addTweetDataAtBottom: " + e.message);
  }
}

//
// Name: updateExistingTweet
// Desc:
//  Check if the specified tweet contains any ban words in the specified array of the ban words.
//
function updateExistingTweet(range, valsRange, idxRow:number, tweet:Tweet, settings:Settings, stats:Stats ) {
  try {
    let prevCountRetweets = valsRange[idxRow][settings.headerInfo.idx_retweet];
    let prevCountFavorites = valsRange[idxRow][settings.headerInfo.idx_favorite];
    stats.countRetweets += (tweet.retweet_count - prevCountRetweets);
    stats.countFavorites += (tweet.favorite_count - prevCountFavorites);
    valsRange[idxRow][settings.headerInfo.idx_retweet] = tweet.retweet_count;
    valsRange[idxRow][settings.headerInfo.idx_favorite] = tweet.favorite_count;
    range.setValues( valsRange ); // Update the sheet
    // logOut( "Updated: idx row = [" + idxRow + "], id=[" + tweet.id_str + "], prev # of retweets = " + prevCountRetweets + ", new # of retweets = " + tweet.retweet_count );
  } catch ( e ) {
    Logger.log("EXCEPTION: aupdateExistingTweet: " + e.message);
  }
}

//
// Name: hasBanWords
// Desc:
//  Check if the specified tweet contains any ban words in the specified array of the ban words.
//
function hasBanWords ( tweet:Tweet, banWords:string[] ){
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
function isFromBanUsers ( tweet:Tweet, banUsers:string[] ){
  for ( let i=0; i< banUsers.length; i++ ) {
    //if ( tweet.screen_name == banUsers[i] ) {
    if ( tweet.user_id_str == banUsers[i] ) {
      return true;
    }
  }
  return false;
}

//
// Name: isSafeTweet
// Desc:
//
function isSafeTweet( tweet:Tweet, settings:Settings ) {
  if ( hasBanWords ( tweet, settings.banWords ) || isFromBanUsers ( tweet, settings.banUsers ) ) {
    return false;
  }
  return true;
}

//
// Name: getRowIndexTweets
// Desc:
//  Check if the specified Tweet has already been recorded (true) or not (false).
//
function getRowIndexTweets(valsRange, tweet:Tweet, settings:Settings) {
  let r = 0;
  for (; r < valsRange.length; r++) {
    let row = valsRange[r];
    if (!row[settings.headerInfo.idx_tweetId]) {
      break;
    }
    if (row[settings.headerInfo.idx_tweetId] == tweet.id_str) {
      return r;
    }
  }
  return -1;
}

//
// Name: isNeededToBeAdded
// Desc:
//
function isNeededToBeAdded( sheet, tweet:Tweet, settings:Settings) {
  let lastRow = sheet.getLastRow();
  if ( 0 < lastRow - (settings.headerInfo.idx_row + 1 )) {
    let range = sheet.getRange(settings.headerInfo.idx_row + 2, 1, lastRow - (settings.headerInfo.idx_row + 1), CELL_HEADER_TITLES_DATA.length);
    if (range) {
      let valsRange = range.getValues();
      if ( -1 == getRowIndexTweets(valsRange, tweet, settings)) {
        return false;
      }
    }
  }
  return true;
}

//
// Name: updateSheet
// Desc:
//  Add a new row and write the new Tweet data.
// Return:
//  Total point per update
//
function updateSheet(sheet, indexSheet:number, settings:Settings, tweets:Tweet[]):Stats {
  let stats:Stats = new Stats();
  let lastRow = sheet.getLastRow();
  let isInitial = true;
  let range = null;
  let valsRange = null;
  if ( 0 < lastRow - (settings.headerInfo.idx_row + 1 )) {
    isInitial = false;
    range = sheet.getRange(settings.headerInfo.idx_row + 2, 1, lastRow - (settings.headerInfo.idx_row + 1), CELL_HEADER_TITLES_DATA.length);
    if (range) {
      valsRange = range.getValues();
    }
  }
  for ( let i = tweets.length - 1; i >= 0; i--) {
    if ( ! isSafeTweet( tweets[i], settings ) ) {
      continue;
    }
    if ( isInitial ) {
        // pure new Tweet which needs to be ADDED
        addTweetDataAtBottom(sheet, indexSheet, tweets[i], settings, stats);
    }
    else if ( range && valsRange ) {
      let idxRow = getRowIndexTweets(valsRange, tweets[i], settings);
      if ( -1 == idxRow ) {
        // pure new Tweet which needs to be ADDED
        addTweetDataAtBottom(sheet, indexSheet, tweets[i], settings, stats);
      } else {
        // already recorded Tweet which needs to be UPDATED
        updateExistingTweet(range, valsRange, idxRow, tweets[i], settings, stats);
      }
    }
  }
  return stats;
}

//
// Name: getSettings
// Desc:
//  Get settings from the common setting sheet and each query sheet.
//  This function supports both the Common Setting Sheet and settings for each query.
//
function getSettings (sheet) {
  let range = sheet.getRange(1, 1, MAX_ROW_RANGE_SETTINGS, MAX_COLUMN_RANGE_SETTINGS);
  if (range == null) {
    return;
  }

  let currentKeyword   = "";
  let lastKeyword      = "";
  let limitTweets      = null;
  let emails           = [];
  let banWords         = [];
  let banUsers         = [];
  let bDownloadMedia   = null;
  let ptTweet          = null;
  let ptRetweet        = null;
  let ptReply          = null;
  let ptMedia          = null;
  let ptFavorite       = null;
  let ptAlertThreshold = null;
  let rowEndSettings   = null;

  let valsRange = range.getValues();
  for(let r = 0; r < valsRange.length ; r++) {
    var row = valsRange[r];
    var title:string = String(row[0]);
    if( title ) {
      title = title.replace(/\s+/g, '').toLowerCase().trim();
      switch (title) {
      case CELL_SETTINGS_TITLE_LIMIT:
        if ( row[1] ) {
          limitTweets = Number( row[1] );
        }
        break;

      case CELL_SETTINGS_TITLE_KEYWORD:
        if ( row[1] ) {
          currentKeyword = String(row[1]).trim();
          row[1] = currentKeyword;
        }
        break;

      case CELL_SETTINGS_TITLE_EMAIL:
        for(let c=1; c<row.length; c++) {
          let val:string = String(row[c]).trim();
          if ( !val ) break;
          emails.push( val );
        }
        break;

      case CELL_SETTINGS_TITLE_BAN_WORDS:
        for(let c=1; c<row.length; c++) {
          let val:string = String(row[c]).trim();
          if ( !val ) break;
          banWords.push( val );
        }
        break;

      case CELL_SETTINGS_TITLE_BAN_USERS:
        for(let c=1; c<row.length; c++) {
          let val:string = String(row[c]).trim();
          if ( !val ) break;
          banUsers.push( val );
        }
        break;

      case CELL_SETTINGS_TITLE_DOWNLOAD_MEDIA:
        {
          let val:string = String(row[1]).trim();
            if ( !val ) break;
            bDownloadMedia = (val.toLowerCase() == "yes");
        }
        break;

      case CELL_SETTINGS_TITLE_POINT_TWEET:
        {
          let val:string = String(row[1]).trim();
            if ( !val ) break;
          ptTweet = Number( val );
        }
        break;

      case CELL_SETTINGS_TITLE_POINT_RETWEET:
        {
          let val:string = String(row[1]).trim();
          if ( !val ) break;
          ptRetweet = Number( val );
        }
        break;

      case CELL_SETTINGS_TITLE_POINT_TWEET:
        {
          let val:string = String(row[1]).trim();
          if ( !val ) break;
          ptReply = Number( val );
        }
        break;

      case CELL_SETTINGS_TITLE_POINT_MEDIA:
        {
          let val:string = String(row[1]).trim();
          if ( !val ) break;
          ptMedia = Number( val );
        }
        break;

      case CELL_SETTINGS_TITLE_POINT_FAVORITE:
        {
          let val:string = String(row[1]).trim();
          if ( !val ) break;
          ptFavorite = Number( val );
        }
        break;

      case CELL_SETTINGS_TITLE_ALERT_THRESHOLD:
        {
          let val:string = String(row[1]).trim();
          if ( !val ) break;
          ptAlertThreshold = Number( val );
        }
        break;

      case CELL_SETTINGS_TITLE_LAST_KEYWORD:
        {
          let val:string = String(row[1]);
          if ( val ) {
            lastKeyword = val;
          }
          row[1] = currentKeyword;
        }
        break;

      case CELL_SETTINGS_TITLE_LAST_UPDATED:
        {
          row[1] = g_timestamp;
        }
        break;

      case CELL_SETTINGS_TITLE_END_SETTINGS:
        {
          rowEndSettings = r;
        }
        break;
      }
      if ( rowEndSettings ) {
        break;
      }
    }
  }
  range.setValues( valsRange );
  return new Settings (currentKeyword, lastKeyword, limitTweets, emails, banWords, banUsers, bDownloadMedia, ptTweet, ptRetweet, ptReply, ptMedia, ptFavorite, ptAlertThreshold, rowEndSettings, null);
}

//
// Name: getSettingsActual
// Desc:
//  Returns the actual settings which cover the common settings and local settings for each sheet.
//
function getSettingsActual( settingsCommon:Settings, settingsLocal:Settings ) {
  return new Settings (
    settingsLocal.currentKeyword
    , settingsLocal.lastKeyword
    , settingsLocal.limitTweets      ? settingsLocal.limitTweets      : ( settingsCommon.limitTweets      ? settingsCommon.limitTweets      : DEFAULT_LIMIT_TWEETS    )
    , settingsCommon.emails.concat(settingsLocal.emails)
    , settingsCommon.banWords.concat(settingsLocal.banWords)
    , settingsCommon.banUsers.concat(settingsLocal.banUsers)
    , settingsCommon.bDownloadMedia  ? settingsCommon.bDownloadMedia  : ( settingsLocal.bDownloadMedia    ? settingsLocal.bDownloadMedia    : DEFAULT_DOWNLOAD_MEDIA  )
    , settingsLocal.ptTweet          ? settingsLocal.ptTweet          : (settingsCommon.ptTweet           ? settingsCommon.ptTweet          : DEFAULT_POINT_TWEET     )
    , settingsLocal.ptRetweet        ? settingsLocal.ptRetweet        : ( settingsCommon.ptRetweet        ? settingsCommon.ptRetweet        : DEFAULT_POINT_RETWEET   )
    , settingsLocal.ptReply          ? settingsLocal.ptReply          : ( settingsCommon.ptReply          ? settingsCommon.ptReply          : DEFAULT_POINT_REPLY     )
    , settingsLocal.ptMedia          ? settingsLocal.ptMedia          : ( settingsCommon.ptMedia          ? settingsCommon.ptMedia          : DEFAULT_POINT_MEDIA     )
    , settingsLocal.ptFavorite       ? settingsLocal.ptFavorite       : ( settingsCommon.ptFavorite       ? settingsCommon.ptFavorite       : DEFAULT_POINT_FAVORITE  )
    , settingsLocal.ptAlertThreshold ? settingsLocal.ptAlertThreshold : ( settingsCommon.ptAlertThreshold ? settingsCommon.ptAlertThreshold : DEFAULT_ALERT_THRESHOLD )
    , settingsLocal.rowEndSettings
    , null );
}

//
// Name: getHeaders
// Desc:
//
//
function getHeaderInfo( sheet ) {
  let range = sheet.getRange( 1, 1, MAX_ROW_SEEK_HEADER, CELL_HEADER_TITLES_DATA.length);
  if (range == null) {
    return;
  }
  let valsRange = range.getValues();
  let r = 0, c;
  let row;
  let rowHeader = -1;
  for( ; r < valsRange.length ; r++) {
    row = valsRange[r];
    for (c = 0; c < row.length; c ++ ) {
      var title:string = String(row[c]);
      title = title.replace(/\s+/g, '').toLowerCase().trim();
      if (title == CELL_HEADER_TITLES_DATA[0].replace(/\s+/g,'').toLowerCase().trim()) {
        rowHeader = r;
        break;
      }
    }
    if ( rowHeader >= 0 ) { break; }
  }
  if ( rowHeader == -1 ) {
    return null;
  }

  let headerInfo = new HeaderInfo();
  headerInfo.idx_row = r;
  for( c = 0; c < valsRange.length; c++) {
    var title:string = String(row[c]);
    title = title.replace(/\s+/g, '').toLowerCase().trim();
    switch ( title ) {
      case "tweetid":
        headerInfo.idx_tweetId = c;
        break;
      case "createdat":
        headerInfo.idx_createdAt = c;
        break;
      case "userid":
        headerInfo.idx_userId = c;
        break;
      case "username":
        headerInfo.idx_userName = c;
        break;
      case "screenname":
        headerInfo.idx_screenName = c;
        break;
      case "retweet":
        headerInfo.idx_retweet = c;
        break;
      case "favorite":
        headerInfo.idx_favorite = c;
        break;
      case "tweet":
        headerInfo.idx_tweet = c;
        break;
      case "media":
        headerInfo.idx_media = c;
        break;
    }
  }
  if (
    null == headerInfo.idx_row
    || null == headerInfo.idx_tweetId
    || null == headerInfo.idx_createdAt
    || null == headerInfo.idx_userId
    || null == headerInfo.idx_userName
    || null == headerInfo.idx_screenName
    || null == headerInfo.idx_tweet
    || null == headerInfo.idx_retweet
    || null == headerInfo.idx_favorite
    || null == headerInfo.idx_media ) {
    return null;
  }
  return headerInfo;
}

//
// Name: main
// Desc:
//  Entry point of this program.
//
function main() {
  let sheets = g_book.getSheets();

  let indexSheet = 0;
  sheets.forEach((sheet) => {
    let sheetName:string = sheet.getName();
    if (sheetName.toLocaleLowerCase().trim().startsWith('!')) {
      return;
    }
    //
    // loading common settings
    //
    if (sheetName.toLocaleLowerCase().trim() === SHEET_NAME_COMMON_SETTINGS) {
      g_settingsCommon = getSettings( sheet );
      return;
    }
    //
    // loading local settings
    //
    let settingsLocal = getSettings( sheet );
    if ( !settingsLocal.currentKeyword ) {
      return;
    }

    //
    // create temporary settings which cover common and local settings
    //
    let settingsActual = getSettingsActual( g_settingsCommon, settingsLocal);

    //
    // get the header row info to add data
    //
    let headerInfo = getHeaderInfo( sheet );
    if ( ! headerInfo ) {
      return;
    }
    settingsActual.headerInfo = headerInfo;

    if ( !settingsLocal.currentKeyword || settingsLocal.lastKeyword != settingsLocal.currentKeyword ) {
      gsClearData(sheet, headerInfo.idx_row + 2);
    }

    var tweets = twSearchTweet(settingsActual.currentKeyword);
    if (null != tweets && 0 < tweets.length) {
      let stats:Stats = updateSheet(sheet, indexSheet, settingsActual, tweets);
      let pt = getPoint(stats, settingsActual);
      if ( pt > settingsActual.ptAlertThreshold ) {
        sendMail(sheet, stats, settingsActual);
      }
      addHistory( sheet.getName(), indexSheet, stats, settingsActual );
    }
    indexSheet ++;
  });
}

//
// Name: duplicateBook
// Desc:
//  Duplicate the specified book at the specified drive
//
function duplicateBook( dateNow: Date) {
  try {
    // crete a backup folder of the day
    g_folderBackup = DriveApp.getFolderById(VAL_ID_GDRIVE_FOLDER_BACKUP);
    let strDate = Utilities.formatDate(dateNow, TIME_LOCALE, FORMAT_DATETIME_DATE);
    let foldersOfDate = g_folderBackup.getFoldersByName(strDate);
    let folderDate = null;
    if (foldersOfDate.hasNext()) {
      folderDate = foldersOfDate.next();
    }else {
      folderDate = g_folderBackup.createFolder( strDate );
    }

    // duplicate the working spreadsheet
    let strDateISO8601 = Utilities.formatDate(dateNow, TIME_LOCALE, FORMAT_DATETIME_ISO8601_DATE) + "T" + Utilities.formatDate(dateNow, TIME_LOCALE, FORMAT_DATETIME_ISO8601_TIME) + "+09:00"
    let fileTarget = DriveApp.getFileById(VAL_ID_TARGET_BOOK);
    let nameFileBackup = "BACKUP_" + strDateISO8601 + "_" + fileTarget.getName();
    fileTarget.makeCopy(nameFileBackup, folderDate);
    return true;
  } catch (ex) {
    errOut( ex );
    return false;
  }
}

//
// Name: moveData
// Desc:
//
//
function moveData( nameBookSrc:string, fullYear:number, month:number, nameSheet:string, indexSheet:number, valsRangeSrc:number[][], rowStart:number, rowNum:number ) {
  try {
    let nameBook:string = nameBookSrc + "/" + String(fullYear)+ "_" + ('00'+month).slice(-2);
    let sheetBackup = getSheet(g_folderBackup, nameBook, nameSheet, indexSheet, inseartHeaderData);

    let lastRowBackup = sheetBackup.getLastRow();
    let maxRowsBackup = sheetBackup.getMaxRows();

    if ( maxRowsBackup - lastRowBackup < rowNum ) {
      sheetBackup.insertRowsAfter(sheetBackup.getMaxRows(), rowNum - (maxRowsBackup - lastRowBackup));
    }
    let rangeBackup = sheetBackup.getRange(sheetBackup.getLastRow() + 1, 1, rowNum, CELL_HEADER_TITLES_DATA.length);
    if ( rangeBackup ) {
      let valsRangeBackup = rangeBackup.getValues();
      for ( let r=0; r<rowNum; r++) {
        for (let c=0; c< CELL_HEADER_TITLES_DATA.length; c++){
          valsRangeBackup[r][c] = valsRangeSrc[rowStart + r][c];
        }
        valsRangeBackup[r][0] = '=HYPERLINK("https://twitter.com/' + valsRangeBackup[r][4] + '/status/' + valsRangeBackup[r][0] + '", "' + valsRangeBackup[r][0] + '")';
      }
      rangeBackup.setValues(valsRangeBackup);
    }
    return true;
  } catch (ex) {
    errOut( "moveData: " + ex );
    return false;
  }
}

//
// Name: backup
// Desc:
//  Entry point for backup
//
function backup() {
  let sheets = g_book.getSheets();
  let nameBook = g_book.getName();

  let indexSheet = 0;
  sheets.forEach((sheet) => {
    let sheetName:string = sheet.getName();
    if (sheetName.toLocaleLowerCase().trim().startsWith('!')) {
      return;
    }
    if (sheetName.toLocaleLowerCase().trim() === SHEET_NAME_COMMON_SETTINGS) {
      return;
    }

    let headerInfo = getHeaderInfo( sheet );
    if ( ! headerInfo ) {
      return;
    }

    let lastFullYear:number = -1;
    let lastMonth:number = -1;
    let lastRow = sheet.getLastRow();
    if ( 0 < lastRow - (headerInfo.idx_row + 1 )) {
      let range = sheet.getRange(headerInfo.idx_row + 2, 1, lastRow - (headerInfo.idx_row + 1), CELL_HEADER_TITLES_DATA.length);
      if (range) {
        let valsRange = range.getValues();
        let rowNumBackuped = 0;
        let r = 0;
        for ( ; r < valsRange.length; r++) {
          let row = valsRange[r];

          let strCreatedAt:string = String(row[headerInfo.idx_createdAt]);
          strCreatedAt = strCreatedAt.replace('(','').replace(')','').replace(' ','T') + "+09:00";
          let dateCreatedAt = new Date( strCreatedAt );

          if ( g_datetime.getTime() - dateCreatedAt.getTime() <= DURATION_MILLISEC_NOT_FOR_BACKUP ) {
            break;
          }

          let year:number = dateCreatedAt.getFullYear();
          let month:number = 1 + dateCreatedAt.getMonth();

          if ( lastMonth == -1 ) {
            lastFullYear = year;
            lastMonth = month;
          } else if ( month != lastMonth ) {
            if ( ! moveData( nameBook, lastFullYear, lastMonth, sheetName, indexSheet, valsRange, rowNumBackuped, r - rowNumBackuped) ) {
              errOut("BACKUP PROCESS WAS TERMINATED.");
              return;
            }
            lastFullYear = year;
            lastMonth = month;
            rowNumBackuped = r;
          }
        }

        if ( r - rowNumBackuped > 0 ) {
            if ( ! moveData( nameBook, lastFullYear, lastMonth, sheetName, indexSheet, valsRange, rowNumBackuped, r - rowNumBackuped) ) {
              errOut("BACKUP PROCESS WAS TERMINATED.");
              return;
            }
            rowNumBackuped = r;
        }
        if ( rowNumBackuped > 0 ) {
          sheet.deleteRows(headerInfo.idx_row + 2, rowNumBackuped);
        }
      }
    }
    indexSheet ++;
  });
}