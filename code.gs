/**
 * Code from: https://github.com/ClaudiaJ/gas-calendar-accept
 * Modified by Torbjørn Stensland, 10.2019
 * Other links: 
 * https://developers.google.com/apps-script/reference/calendar/
 * https://developers.google.com/adwords/api/docs/appendix/codes-formats#timezone-ids
 
 Tester
 1.  +Send en ny møteinnkalling (ikke møtekonflikt)
 2.  +Send en ny møteinnkalling (møtekonflikt)
 3.  +Oppdater en ny møteinnkalling (ikke møtekonflikt)
 4.  +Oppdater en ny møteinnkalling (fra ikke møtekonflikt til møtekonflikt)
 5.  +Oppdater en ny møteinnkalling (fra  møtekonflikt til ikke møtekonflikt)
 6.  +Slett møteinnkalling
 7.  +Opprett en ny møteserie (ikke møtekonflikt)
 8.  +Send en ny møteserie (møtekonflikt)
 9.  +Send en ny møteserie (delvis møtekonflikt)
 10. +Oppdater en ny møteserie (ikke møtekonflikt)
 11. +Oppdater en ny møteserie (fra ikke møtekonflikt til møtekonflikt)
 12. +Oppdater en ny møteserie (fra  møtekonflikt til ikke møtekonflikt)
 13. +Slett møteserie 
 14. +Send en ny møteinnkalling (møtekonflikt med flere møter)

 
 
 * Install-info:
 * 1. Add permissions: Resources - Advanced Google Service. Enable: Calendar API
 * 2. Define triggers: Edit - Current project triggers. 
 *      Function to run: ProcessIncites
 *      Select event source: From Calendar
 *      Enter calendar details: Calendar updated
 *      Calendard owner email: Same as calendarId below
 */
var doEmail = true;
var deleteContent = true;
var contentOverride = "Innholdet er skjult";
var emailHeaderPrefixAccepted = "Godtatt: ";
var emailHeaderPrefixDenied = "Avslått: ";
var calendarId = "calendar email@gmail.com";
var meetingRoomName = "My meeting room name";
/**
 * Workaround to [Issue 5323](https://code.google.com/p/google-apps-script-issues/issues/detail?id=5323)
 * statusFilters parameter is not working; returns 0 events.
 * @param {string|array} status - GuestStatus or array of GuestStatus to match
 * @returns {function} - Callback parameter to Array.prototype.filter
 * Torbjørn: The function statusFilters does not work if invite is received from non-gmail accounts (eg Outlook). 
 * The status will always be "OWNER". Instead, we have to chec for the guests guestStatus on the guest email like calendarId.
 */
/*function statusFilters(status) {
  if (status instanceof Array) {
    return function(invite) {
      return status.includes(invite.getMyStatus());
    }
  } else {
    return function(invite) {
      return invite.getMyStatus() === status;
    }
  }
}*/

/**
 * Import an HTML template from file
 * @param {string} file - File to import
 * @param {boolean} [template=false] - If true, evaluate imported file as a template
 * @returns {HtmlOutput} Inline, rendered content
 */
function importTemplate(file, template) {
  if (template) {
    return HtmlService.createTemplateFromFile(file).evaluate().getContent();
  } else {
    return HtmlService.createHtmlOutputFromFile(file).getContent();
  }
}

/**
 * Remove duplicates from array
 * @param {array} array to sort
 * @returns {array} array without duplicates
 */
function RemoveDuplicates(array) {
  var outArray = [];
  array.sort(lowerCase);
  function lowerCase(a,b){
    return a.toLowerCase()>b.toLowerCase() ? 1 : -1;// sort function that does not "see" letter case
  }
  outArray.push(array[0]);
  for(var n in array){
    //Logger.log(outArray[outArray.length-1]+'  =  '+array[n]+' ?');
    if(outArray[outArray.length-1].toLowerCase()!=array[n].toLowerCase()){
      outArray.push(array[n]);
    }
  }
  return outArray;
}



function ProcessInvites() {
  //check if script is already running, if so, lock the script. 
  var lock = LockService.getScriptLock();
  //wait for up to 20 sek for other processes to finish
  lock.waitLock(10000); // in milliseconds
  if (!lock.hasLock()) {
    Logger.log('Could not obtain lock after 20 seconds.');
  }
  Logger.log("ProcessInvites trigged.");
  //var calendarId = PropertiesService.getScriptProperties().getProperty('Id');
  
  var calendar = CalendarApp.getCalendarById(calendarId);

  // Auto-accept any invite between last day and x week from now.
  var weeks = 52*7;
  var timeNow = new Date();
  var start = new Date(timeNow.getTime() - (1000 * 60 * 60 * 24));
  var end = new Date(timeNow.getTime() + (1000 * 60 * 60 * 24 * weeks));
  var invites = [];
  var unfilteredInvites = calendar.getEvents(start, end);
  
  //email contentarrays: 
  var inviteEmailContent = [];
  var conflictEmailContent = [];
  // get status-object
  var calendarInvited = CalendarApp.GuestStatus.INVITED;
  //var calendarMaybe = CalendarApp.GuestStatus.MAYBE;
  //var calendarNo = CalendarApp.GuestStatus.NO;
  //var calendarOwner = CalendarApp.GuestStatus.OWNER;
  var calendarYes = CalendarApp.GuestStatus.YES;
  
  //filter the invites
  for (var i = 0; i<unfilteredInvites.length; i++){
    var title = unfilteredInvites[i].getTitle();
    //check for guests invited to the event. If calendaremail status is "Invited", add this event
    var guests = unfilteredInvites[i].getGuestList(true);
    for (var j = 0; j<guests.length; j++){
       //Get the email of the gmail-calendar and check if the status is "Invited", if so, add to list.      
       if(guests[j].getEmail()==calendarId){
         if (guests[j].getGuestStatus() == calendarInvited){
           invites.push(unfilteredInvites[i]);
         }
       }
     }    
  }

  //Check for conflicts with existing accepted meetings
  for (var i = 0; i < invites.length; i++) {
    //get the original calendar id, which is the same as the email who send the invite
    var inviteEmail = invites[i].getOriginalCalendarId();
    var unfilteredConflicts = calendar.getEvents(invites[i].getStartTime(), invites[i].getEndTime());
    //check if existing event at the same time is accepted. 
    var conflicts =[];
    if (unfilteredConflicts.length>0){
      for (var j = 0; j<unfilteredConflicts.length; j++){
        var guests = unfilteredConflicts[j].getGuestList(true);
        for (var k = 0; k<guests.length; k++){
          //Get the email of the gmail-calendar and check if the status is "Yes", if so, add to list.    
          if(guests[k].getEmail()==calendarId){
            if (guests[k].getGuestStatus() == calendarYes){
              conflicts.push(unfilteredConflicts[j]);
            }
          }
        }  
      }
    }
    //loop through all conflicts
   
    for (var ci = 0; ci < conflicts.length; ci++) {
      Logger.log("Found a potential conflict to: " + invites[i].getTitle());
      Logger.log("Creator is: " + inviteEmail);
      var emailContent = {
        "inviteCreators": inviteEmail,
        "inviteTitle": invites[i].getTitle(),
        "inviteStart": Utilities.formatDate(invites[i].getStartTime(), "Europe/Oslo", "yyyy-MM-dd HH:mm"),
        "inviteEnd": Utilities.formatDate(invites[i].getEndTime(), "Europe/Oslo", "yyyy-MM-dd HH:mm"),
        "conflictStart": Utilities.formatDate(conflicts[ci].getStartTime(), "Europe/Oslo", "yyyy-MM-dd HH:mm"),
        "conflictEnd": Utilities.formatDate(conflicts[ci].getEndTime(), "Europe/Oslo", "yyyy-MM-dd HH:mm"),
        "conflictTitle": conflicts[ci].getTitle(),
        "conflictCreators": conflicts[ci].getOriginalCalendarId()
      };
      
      //add all conflicting meetings to array, sorted by inviters.
      if(conflictEmailContent.length>0){
        var found = false;
        for (var k = 0; k < conflictEmailContent.length; k++) {
          if(conflictEmailContent[k][0].inviteCreators == inviteEmail){            
            conflictEmailContent[k].push(emailContent);
            found = true;
            break;
          }          
        }
        if (!found){
          conflictEmailContent.push(new Array(emailContent));         
        }
      }
      else {
        conflictEmailContent.push(new Array(emailContent)); 
      }

      //Delete the event
      try{
        invites[i].deleteEvent();
      }
      catch (e){
        Logger.log("Error during deletion:");
        Logger.log(e);
      }
    }
    
 

    if (conflicts.length === 0) {
      Logger.log("No conflict, accepting: " + invites[i].getTitle());
      var emailContent = {
        "inviteCreators": inviteEmail,
        "inviteTitle": invites[i].getTitle(),
        "inviteStart": Utilities.formatDate(invites[i].getStartTime(), "Europe/Oslo", "yyyy-MM-dd HH:mm"),
      };
      //add all unaccepted meetings to array, sorted by inviters.
      if(inviteEmailContent.length>0){
        var found = false;
        for (var k = 0; k < inviteEmailContent.length; k++) {
          if(inviteEmailContent[k][0].inviteCreators == inviteEmail){
            inviteEmailContent[k].push(emailContent);
            found = true;
            break;
          }          
        }
        if (!found){
          inviteEmailContent.push(new Array(emailContent));         
        }
      }
      else {
        inviteEmailContent.push(new Array(emailContent)); 
      }
      //modify the event (title, content and attachment)
      invites[i].setTitle(emailHeaderPrefixAccepted + invites[i].getTitle());
      if(deleteContent){
        invites[i].setDescription(contentOverride);
      }
      invites[i].setMyStatus(CalendarApp.GuestStatus.YES);
     
    }
  }
  
  //Define variables used in header/footer HTML-templates
    var footerContent = {
    "meetingRoomName": meetingRoomName
  };
  
  //send email for accept
  if (inviteEmailContent.length!=0){
    //send one email pr creator with all new accepted meetings.
    for (var i = 0; i < inviteEmailContent.length; i++) {
      var body =importTemplate('AutoResponse-accept-header');
      var inviteEmail = inviteEmailContent[i][0].inviteCreators;
      var title = inviteEmailContent[i][0].inviteTitle;
      for (var j = 0; j < inviteEmailContent[i].length; j++) {
        body = body + importTemplate('AutoResponse-accept-content').replace(/{{([a-zA-Z\.]+)}}/g, function(match, p1, offset, string) {
          return inviteEmailContent[i][j][p1];
        });
      }
      body = body + importTemplate('AutoResponse-footer').replace(/{{([a-zA-Z\.]+)}}/g, function(match, p1, offset, string) {
          return footerContent[p1];
        });
      if (doEmail) {
        GmailApp.sendEmail(inviteEmail, emailHeaderPrefixAccepted + title, "",
                           { 
                           name: meetingRoomName,
                           htmlBody: body 
                           });
      }
    }
  }
  //send email for conflicts
  if (conflictEmailContent.length!=0){
    //send one email pr creator with all new conflicts meetings.
    for (var i = 0; i < conflictEmailContent.length; i++) {
      var body =importTemplate('AutoResponse-conflict-header');
      var inviteEmail = conflictEmailContent[i][0].inviteCreators;
      var title = conflictEmailContent[i][0].inviteTitle;
      for (var j = 0; j < conflictEmailContent[i].length; j++) {
        body = body + importTemplate('AutoResponse-conflict-content').replace(/{{([a-zA-Z\.]+)}}/g, function(match, p1, offset, string) {
          return conflictEmailContent[i][j][p1];
        });
      }
      body = body + importTemplate('AutoResponse-footer').replace(/{{([a-zA-Z\.]+)}}/g, function(match, p1, offset, string) {
          return footerContent[p1];
        });
      if (doEmail) {
        GmailApp.sendEmail(inviteEmail, emailHeaderPrefixDenied + title, "",
                           { 
                           name: meetingRoomName,
                           htmlBody: body 
                           });
      }
    }
  }
      
      
  lock.releaseLock();
}
