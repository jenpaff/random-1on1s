function run() {
  var peopleSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("People");
  var range = peopleSheet.getDataRange();
  var values = range.getValues();
  var namesAndEmails = values.slice(3);
  namesAndEmails = namesAndEmails.filter(function (arr) {return arr[0] !== ""}); 
  Logger.log("Participants: " + namesAndEmails.length);
  shuffle(namesAndEmails);
  for (var i = 0; i < namesAndEmails.length - 1; i += 2) {
    sendMailsToPair(namesAndEmails[i], namesAndEmails[i+1], "");
  }
  if (namesAndEmails.length % 2 != 0) {
    Logger.log("odd number of participants, choosing someone for second pair");
    var idx = Math.floor(Math.random() * (namesAndEmails.length - 1));
    sendMailsToPair(namesAndEmails[idx],
      namesAndEmails[namesAndEmails.length - 1],
      "@" + namesAndEmails[idx][0] + ": If you're wondering why this is your second mail: This is because an odd" +
      " number of people signed up and we needed someone to fill that gap.");
  }
  Logger.log("done");
}

function sendMailsToPair(pair1, pair2, sentence) {
  Logger.log(pair1[0] + " " + pair1[1]);
  Logger.log(pair2[0] + " " + pair2[1]);
  Logger.log("============");

  var body = [
    "Hi there,",
    "You've signed up to the RANDOM 1on1 !",
    "Now we shuffled all people that signed up and chose this pair: " + pair1[0] + " and " + pair2[0],
    "Please contact your assigned partner to find a time slot for having your Random 1on1 or just send them an encouraging message.", 
    sentence,
    "This email was generated automatically, if you have any doubts or feedback or want to be removed from the list please contact [INSERT CONTACT].",
  ];

  Logger.log("sending to " + pair1[1] + ',' + pair2[1]);
  Logger.log("[Random 1on1s] " + pair1[0] + " & " + pair2[0]);
  try {
    GmailApp.sendEmail(pair1[1] + ',' + pair2[1],
                       "[INSERT SUBJECT] " + pair1[0] + " & " + pair2[0],
                       body.join('\n\n'),
      {name: '[INSERT NAME]'});
  } catch (e) {
    Logger.log("Exception while sending mail " + e);
  }
}

function shuffle(a) {
    var j, x, i;
    for (i = a.length - 1; i > 0; i--) {
        j = Math.floor(Math.random() * (i + 1));
        x = a[i];
        a[i] = a[j];
        a[j] = x;
    }
    return a;
}

