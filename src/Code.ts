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
const VAL_ID_TARGET_BOOK       = '1zjxr1BO7fQmO-TDP-HYOL3uh-1g1zKOCwEogX0h3uew';
// ID of the Google Drive where the images will be placed
const VAL_ID_GDRIVE_FOLDER     = '1ttFnVjcZJNJdloaT4ni99ZbEYuEj-WAA';
// ID of the Google Drive where backup files will be placed
const VAL_ID_GDRIVE_FOLDER_BACKUP     = '1BVPS0fXT0UzkvuIe0gDESkhHApcyxZsR';
// Key and Secret to access Twitter APIs
const VAL_CONSUMER_API_KEY     = 'lyhQXM3aZP5HHLqO6jjcEwzux';
const VAL_CONSUMER_API_SECRET  = 'ODBAEfj4VlgVf1BKHJyoz6QwryiaNtcchkr4PikzjSHmct0hjr';

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// DEFINES
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
const VERSION                               = 1.1;
const DURATION_MONTH_BACKUP                 = 3;
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

const CELL_HEADER_TITLES            = ["tweetid", "createdat", "userid", "username", "screenname", "retweet", "favorite", "tweet", "media"];

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
let g_datetime                    = null;
let g_book                        = null;
let g_folder                      = null;
let g_folderBackup                = null;
let g_settingsCommon              = null;

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
      let rngVals = range.getValues();
      let row = rngVals[0];
      row[0] = g_datetime;
      row[1] = String(text);
      range.setValues( rngVals );
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
// Name: addTweetDataAtBottom
// Desc:
//  Add the specified text at the bootom of the specified sheet.
//
function addTweetDataAtBottom(sheet, tweet:Tweet, settings:Settings, stats:Stats ) {
  try {
    let lastRow = sheet.getLastRow();
    let rowToWrite = lastRow + 1;
    let folder = null;
    if ( 0 < tweet.list_media.length && settings.bDownloadMedia ) {
      let strDate = Utilities.formatDate(new Date(tweet.created_at), TIME_LOCALE, FORMAT_DATETIME_DATE);
      let foldersOfDate = g_folder.getFoldersByName(strDate);
      let folderDate = null;
      if (foldersOfDate.hasNext()) {
        folderDate = foldersOfDate.next();
      }else {
        folderDate = g_folder.createFolder( strDate );
      }
      folder = folderDate.createFolder( tweet.id_str );
      downloadMedia( folder, tweet.list_media );
    }
    let rangeLastRow = sheet.getRange( rowToWrite, 1, 1, CELL_HEADER_TITLES.length);
    if (rangeLastRow) {
      let rngVals = rangeLastRow.getValues();
      let row = rngVals[0];

      row[settings.headerInfo.idx_tweetId    ] = '=HYPERLINK("https://twitter.com/' + tweet.screen_name + '/status/' + tweet.id_str + '", "' + tweet.id_str + '")';
      row[settings.headerInfo.idx_createdAt  ] = tweet.created_at_str;
      row[settings.headerInfo.idx_userId     ] = tweet.user_id_str;
      row[settings.headerInfo.idx_userName   ] = tweet.user_name;
      row[settings.headerInfo.idx_screenName ] = tweet.screen_name;
      row[settings.headerInfo.idx_retweet    ] = tweet.retweet_count;
      row[settings.headerInfo.idx_favorite   ] = tweet.favorite_count;
      row[settings.headerInfo.idx_tweet      ] = tweet.text;

      if ( folder ) {
        row[settings.headerInfo.idx_media] = folder.getUrl();
      }
      rangeLastRow.setValues( rngVals );

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
function updateExistingTweet(range, rngVals, idxRow:number, tweet:Tweet, settings:Settings, stats:Stats ) {
  try {
    let prevCountRetweets = rngVals[idxRow][settings.headerInfo.idx_retweet];
    let prevCountFavorites = rngVals[idxRow][settings.headerInfo.idx_favorite];
    stats.countRetweets += (tweet.retweet_count - prevCountRetweets);
    stats.countFavorites += (tweet.favorite_count - prevCountFavorites);
    rngVals[idxRow][settings.headerInfo.idx_retweet] = tweet.retweet_count;
    rngVals[idxRow][settings.headerInfo.idx_favorite] = tweet.favorite_count;
    range.setValues( rngVals ); // Update the sheet
    logOut( "Updated: idx row = [" + idxRow + "], id=[" + tweet.id_str + "], prev # of retweets = " + prevCountRetweets + ", new # of retweets = " + tweet.retweet_count );
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
function getRowIndexTweets(rngVals, tweet:Tweet, settings:Settings) {
  let r = 0;
  for (; r < rngVals.length; r++) {
    let row = rngVals[r];
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
    let range = sheet.getRange(settings.headerInfo.idx_row + 2, 1, lastRow - (settings.headerInfo.idx_row + 1), CELL_HEADER_TITLES.length);
    if (range) {
      let rngVals = range.getValues();
      if ( -1 == getRowIndexTweets(rngVals, tweet, settings)) {
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
function updateSheet(sheet, settings:Settings, tweets:Tweet[]):Stats {
  let stats:Stats = new Stats();
  let lastRow = sheet.getLastRow();
  let isInitial = true;
  let range = null;
  let rngVals = null;
  if ( 0 < lastRow - (settings.headerInfo.idx_row + 1 )) {
    isInitial = false;
    range = sheet.getRange(settings.headerInfo.idx_row + 2, 1, lastRow - (settings.headerInfo.idx_row + 1), CELL_HEADER_TITLES.length);
    if (range) {
      rngVals = range.getValues();
    }
  }
  for ( let i = tweets.length - 1; i >= 0; i--) {
    if ( ! isSafeTweet( tweets[i], settings ) ) {
      continue;
    }
    if ( isInitial ) {
        // pure new Tweet which needs to be ADDED
        addTweetDataAtBottom(sheet, tweets[i], settings, stats);
    }
    else if ( range && rngVals ) {
      let idxRow = getRowIndexTweets(rngVals, tweets[i], settings);
      if ( -1 == idxRow ) {
        // pure new Tweet which needs to be ADDED
        addTweetDataAtBottom(sheet, tweets[i], settings, stats);
      } else {
        // already recorded Tweet which needs to be UPDATED
        updateExistingTweet(range, rngVals, idxRow, tweets[i], settings, stats);
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

  let rngVals = range.getValues();
  for(let r = 0; r < rngVals.length ; r++) {
    var row = rngVals[r];
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
          if ( !val ) break;
          lastKeyword = val;
          row[1] = currentKeyword;
        }
        break;

      case CELL_SETTINGS_TITLE_LAST_UPDATED:
        {
          row[1] = g_datetime;
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
  range.setValues( rngVals );
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
  let range = sheet.getRange( 1, 1, MAX_ROW_SEEK_HEADER, CELL_HEADER_TITLES.length);
  if (range == null) {
    return;
  }
  let rngVals = range.getValues();
  let r = 0, c;
  let row;
  let rowHeader = -1;
  for( ; r < rngVals.length ; r++) {
    row = rngVals[r];
    for (c = 0; c < row.length; c ++ ) {
      var title:string = String(row[c]);
      title = title.replace(/\s+/g, '').toLowerCase().trim();
      if (title == CELL_HEADER_TITLES[0]) {
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
  for( c = 0; c < rngVals.length; c++) {
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
  g_datetime = TIME_LOCALE + ": " + Utilities.formatDate(new Date(), TIME_LOCALE, FORMAT_DATETIME);
  g_folder = DriveApp.getFolderById(VAL_ID_GDRIVE_FOLDER);
  g_book = SpreadsheetApp.openById(VAL_ID_TARGET_BOOK);
  let sheets = g_book.getSheets();

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

    if ( !settingsLocal.currentKeyword || settingsLocal.lastKeyword != settingsLocal.currentKeyword ) {
      gsClearData(sheet, settingsLocal.rowEndSettings + 1);
    }
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

    var tweets = twSearchTweet(settingsActual.currentKeyword);
    if (null != tweets && 0 < tweets.length) {
      let stats:Stats = updateSheet(sheet, settingsActual, tweets);
      let pt = getPoint(stats, settingsActual);
      if ( pt > settingsActual.ptAlertThreshold ) {
        sendMail(sheet, stats, settingsActual);
      }
    }
  });
}

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

function backup() {

  g_book = SpreadsheetApp.openById(VAL_ID_TARGET_BOOK);

  let dateNow:Date = new Date();

  if ( ! duplicateBook(dateNow) ) {
    errOut( "BACKUP: Cannot duplicate the book" );
    return;
  }

  // remove unnecessary data from the target sheet
  let sheets = g_book.getSheets();

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

    let dateBaseBackup:Date = new Date(dateNow); // duplicate the current DateNow
    dateBaseBackup= new Date(dateBaseBackup.setMonth(dateNow.getMonth()-DURATION_MONTH_BACKUP));
    let lastRow = sheet.getLastRow();
    if ( 0 < lastRow - (headerInfo.idx_row + 1 )) {
      let range = sheet.getRange(headerInfo.idx_row + 2, 1, lastRow - (headerInfo.idx_row + 1), CELL_HEADER_TITLES.length);
      if (range) {
        let rngVals = range.getValues();
        let r = 0;
        for ( ; r<rngVals.length; r++) {
          let row = rngVals[r];

          let strCreatedAt:string = String(row[headerInfo.idx_createdAt]);
          strCreatedAt = strCreatedAt.replace('(','').replace(')','').replace(' ','T') + "+09:00";
          let dateCreatedAt = new Date( strCreatedAt );
          if ( dateCreatedAt.getTime() > dateBaseBackup.getTime() )
          {
            break;
          }
        }
        if ( r > 0 ) {
          sheet.deleteRows( headerInfo.idx_row + 2, r );
        }
      }
    }
  });
}