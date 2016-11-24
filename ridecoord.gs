/*
 * @author Matthew Sit (2016), msit@berkeley.edu
 *
 * This script automates coordinating rides by automatically:
 *    - Clearing user responses every Sunday at 3pm to prepare for the following week.
 *    - Sending an email reminder every Wednesday at 7pm to remind people to sign up for rides.
 *    - Organizing rides based upon responses, sending a Provisional email to coordinators Saturday at 12pm
 *      and the Final email to attendees Saturday at 8pm.
 *
 * Note that all names, locations, phone numbers, and email addresses have been redacted/replaced in this public version.
 * Triggers can be scheduled from within the Google Scripts IDE.
 */

/*
 * Deletes responses in columns B and C. Assigns a new tiebreaking number for the week to each person.
 * Trigger: Sundays @ 3pm
 */
function clearResponses() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Rides");

  // Clear column B
  var range = sheet.getRange("B2:B1000");
  range.setValue("");

  // Clear column C
  var friends = sheet.getRange("C2:C1000");
  friends.setValue("");

  // Clear column D
  var friends = sheet.getRange("D2:D1000");
  friends.setValue("");

  // Sort by Name
  sheet.sort(1, true);

  // Assign new tiebreaking number for the week to each person
  for (var i = 2; i <= 1000; i++) {
    var tiebreaker = sheet.getRange("K" + i);
    tiebreaker.setValue(Math.random());
  }
}

/*
 * Sends email reminder to all emails listed in column F.
 * Emails will not be sent to any row where column G is marked with anything.
 * Email pausing should be used for temporary pausing for summer break, study abroad, etc.
 * To unsubscribe from this list, users should directly delete their row of information from the spreadsheet.
 * Trigger: Wednesdays @ 7pm
 */
function sendSignupEmails() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Rides");
  var emailAddresses = sheet.getRange("F2:G1000").getValues();

  // Gets the date of Sunday, which is 4 days after the send date, Wednesday.
  var sunday = new Date();
  // setDate() automatically accounts for overflow into the next month situations.
  sunday.setDate(sunday.getDate() + 4);
  // Must add 1 to getMonth() since by default, Jan = 0, Feb = 1, etc.
  sunday = sunday.getMonth()+1 + "/" + sunday.getDate();

  //Need multiple recipient lists because Google will not send emails to more than 50 addresses in one message.
  var recipients1 = "";
  var recipients2 = "";
  var recipients3 = "";
  var recipients4 = "";
  var recipients5 = ""; //This will support up to 5 * 50 subscribers.
  for (i in emailAddresses) {
    if (emailAddresses[i][0] != "" & emailAddresses[i][1] == "") {
      if (i < 50) { recipients1 += emailAddresses[i][0] + ","; }
      else if (i < 100) { recipients2 += emailAddresses[i][0] + ","; }
      else if (i < 150) { recipients3 += emailAddresses[i][0] + ","; }
      else if (i < 200) { recipients4 += emailAddresses[i][0] + ","; }
      else { recipients5 += emailAddresses[i][0] + ","; }
    }
  }
  recipients1 = recipients1.substring(0, recipients1.length-1); //Remove tailing comma
  if (recipients2.length > 0) { recipients2 = recipients2.substring(0, recipients2.length-1); }
  if (recipients3.length > 0) { recipients3 = recipients3.substring(0, recipients3.length-1); }
  if (recipients4.length > 0) { recipients4 = recipients4.substring(0, recipients4.length-1); }
  if (recipients5.length > 0) { recipients5 = recipients5.substring(0, recipients5.length-1); }

  var subject = "[Rides] Signups for " + sunday;
  var message = "Hi!\n\nPlease sign up by Friday evening at 8 pm if you're coming to church this Sunday! ";
  message += "If anything changes, please text Matt at 123-456-7890 or Caitlyn 123-456-7890.\n\n";
  message += "Sign up here:\nhttps://docs.google.com/spreadsheets/abcd\n\n";
  message += "(If you would like to be removed from this email list </3, you can do so by deleting your information from the spreadsheet.)\n\n";
  message += "Best,\nMatt";
  MailApp.sendEmail("a@email.com, b@email.com",subject, message, {name: "Matthew Sit", bcc: recipients1});
  if (recipients2.length > 0) { MailApp.sendEmail("a@email.com, b@email.com",subject, message, {name: "Matthew Sit", bcc: recipients2}); }
  if (recipients3.length > 0) { MailApp.sendEmail("a@email.com, b@email.com",subject, message, {name: "Matthew Sit", bcc: recipients3}); }
  if (recipients4.length > 0) { MailApp.sendEmail("a@email.com, b@email.com",subject, message, {name: "Matthew Sit", bcc: recipients4}); }
  if (recipients5.length > 0) { MailApp.sendEmail("a@email.com, b@email.com",subject, message, {name: "Matthew Sit", bcc: recipients5}); }
}

/*
 * Sends Test Email to Matt only, for debugging purposes.
 * Trigger: None
 */
function sendTestLogisticsEmails() {
  sendLogisticsEmails(-1);
}

/*
 * Sends Provisional Logistics Email to coords only
 * Trigger: Saturday @ 12pm
 */
function sendProvisionalLogisticsEmails() {
  sendLogisticsEmails(0);
}

/*
 * Sends Final Logistics Email to all attendees
 * Trigger: Saturday @ 8pm
 */
function sendFinalLogisticsEmails() {
  sendLogisticsEmails(1);
}

/*
 * Processes signup data, arranges rides, sends logistics email
 * -1 = test email (for debugging purposes to Matt only)
 *  0 = provisional (send assignments to coords only)
 *  1 = final (send email to all attendees)
 */
function sendLogisticsEmails(send_option) {

  // ***********************
  // Step 1: GRAB THE DATA
  // ***********************

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Rides");
  // Sorts the spreadsheet based on who's coming and takes a headcount
  sheet.sort(2, false);
  //Get headcounts of yes, no, early
  var num_yes = sheet.getRange("L1").getValue();
  var num_no = sheet.getRange("M1").getValue();
  var num_early = sheet.getRange("N1").getValue();
  // Email Addresses of Yes respondents
  var yes_people = sheet.getRange("A2:K" + (2 + num_yes - 1)).getValues();
  // Email Addresses of Early respondents
  if (num_early > 0) {
    var early_people = sheet.getRange("A" + (2 + num_yes + num_no) + ":K" + (2 + num_yes + num_no + num_early - 1)).getValues();
  }

  // ***********************
  // Step 2: ARRANGE RIDES
  // ***********************

  // Create 5 dictionaries that track the computational scores of each person.
  var a_score = {};
  var b_score = {};
  var c_score = {};
  var tiebreaker = {};
  var can_drive = {};
  var bringing_friends = {}; // For use if drivers are bringing friends.

  for (i in yes_people) {
    if (yes_people[i][3] == "Yes") { // Bringing a friend
      // CAUTION: If these scores change, you must update the condition in getFriendOrGuest().
      a_score[yes_people[i][0]] = -1;
      b_score[yes_people[i][0]] = -1;
      c_score[yes_people[i][0]] = 5;
      // Will end up containing all hosts, but hosts who aren't driving will be ignored.
      bringing_friends[yes_people[i][0]] = true;
    } else {
      a_score[yes_people[i][0]] = yes_people[i][7]; // 0
      b_score[yes_people[i][0]] = yes_people[i][8]; // 1
      c_score[yes_people[i][0]] = yes_people[i][9]; // 2
    }
    tiebreaker[yes_people[i][0]] = yes_people[i][10];
    can_drive[yes_people[i][0]] = yes_people[i][2];

    // In case a new row was created with no default values.
    if (yes_people[i][7] == "" && yes_people[i][8] == "" && yes_people[i][9] == "" && yes_people[i][10] == "") {
      a_score[yes_people[i][0]] = -1;
      b_score[yes_people[i][0]] = -1;
      c_score[yes_people[i][0]] = 5;
      tiebreaker[yes_people[i][0]] = 0.5; // Fake a random value to keep it deterministic for the week.
    }
  }

  var assignments = {};
  // List of all drivers
  // Name, Additional People, pickup location (0, 1, or 2), Coming this week (true/false), Bringing friends (true/false)
  var drivers = [
    ["Annita Utsey", 4, 2, false, false],
    ["Hyacinth Weekly", 3, 2, false, false],
    ["Jona Dorado", 4, 0, false, false],
    ["Millard Haslett", 4, 1, false, false],
    ["Reynaldo Schweiger", 4, 2, false, false],
    ["Ubers", 100, 2, false, false]
  ];

  // Remove from driver list those who aren't driving this week
  for (d in drivers) {
    if (drivers[d][0] in can_drive && can_drive[drivers[d][0]] != "Yes" && drivers[d][0] != "Ubers") {
      // Change driver's name in list so won't be assigned riders
      drivers[d][0] = "Not Driving this Week";
    }
  }

  // Flag drivers who are bringing friends this week
  for (d in drivers) {
    if (drivers[d][0] != "Ubers" && bringing_friends[drivers[d][0]] == true) {
      drivers[d][4] = true;
    }
  }

  var any_from_a = false;
  var any_from_b = false;
  var any_from_c = false;
  for (d in drivers) {
    if (drivers[d][0] in a_score || drivers[d][0] == "Ubers") {
      if (Object.keys(a_score).length === 0) {
            break;
      }
      assignments[drivers[d][0]] = [];
      delete a_score[drivers[d][0]];
      delete b_score[drivers[d][0]];
      delete c_score[drivers[d][0]];
      drivers[d][3] = true;
      for (var i = 0; i < drivers[d][1]; i++) {
        if (Object.keys(a_score).length === 0) {
          break;
        } else {
          if (drivers[d][2] == 0) { // Location A
            if (drivers[d][4]) { // Bringing friends
              var curr = getFriendOrGuest(a_score, b_score, c_score, tiebreaker);
            } else {
              var curr = getMaxScorer(a_score, b_score, c_score, tiebreaker);
            }
            assignments[drivers[d][0]].push(curr);
            delete a_score[curr];
            delete b_score[curr];
            delete c_score[curr];
            any_from_a = true;
          } else if (drivers[d][2] == 1) { // Location B
            if (drivers[d][4]) { // Bringing friends
              var curr = getFriendOrGuest(b_score, a_score, c_score, tiebreaker);
            } else {
              var curr = getMaxScorer(b_score, a_score, c_score, tiebreaker);
            }
            assignments[drivers[d][0]].push(curr);
            delete a_score[curr];
            delete b_score[curr];
            delete c_score[curr];
            any_from_b = true;
          } else if (drivers[d][2] == 2) { // Location C
            if (drivers[d][4]) { // Bringing friends
              var curr = getFriendOrGuest(c_score, a_score, b_score, tiebreaker);
            } else {
              var curr = getMaxScorer(c_score, a_score, b_score, tiebreaker);
            }
            assignments[drivers[d][0]].push(curr);
            delete a_score[curr];
            delete b_score[curr];
            delete c_score[curr];
            any_from_c = true;
          }
        }
      }
    }
  }

  if (drivers[drivers.length-1][3] == true) {
    assignments[drivers[drivers.length-1][0]].push("<i>(" + assignments["Ubers"].length + " people)</i>");
  }

  // Ctrl + Enter to see println output
  // Logger.log(assignments);

  // ***********************
  // Step 3: SEND THE EMAIL
  // ***********************

  // Gets the date of Sunday, which is 1 day after the send date, Saturday.
  var sunday = new Date();
  // setDate() automatically accounts for overflow into the next month situations.
  sunday.setDate(sunday.getDate() + 1);
  // Must add 1 to getMonth() since by default, Jan = 0, Feb = 1, etc.
  sunday = sunday.getMonth()+1 + "/" + sunday.getDate();

  //Need multiple recipient lists because Google will not send emails to more than 50 addresses in one message.
  var recipients1 = "";
  var recipients2 = "";
  var recipients3 = "";
  var recipients4 = "";
  var recipients5 = ""; //This will support up to 5 * 50 subscribers.
  for (i in yes_people) {
    // Save room to add early people into recipients1
    if (yes_people[i][5] != "") {
      if (i < 50 - num_early) { recipients1 += yes_people[i][5] + ","; }
      else if (i < 100  - num_early) { recipients2 += yes_people[i][5] + ","; }
      else if (i < 150  - num_early) { recipients3 += yes_people[i][5] + ","; }
      else if (i < 200  - num_early) { recipients4 += yes_people[i][5] + ","; }
      else { recipients5 += yes_people[i][5] + ","; }
    }
  }

  // Add any early recipients into recipients1
  if (num_early > 0) {
    for (i in early_people) {
      if (early_people[i][5] != "") {
        recipients1 += early_people[i][5] + ",";
      }
    }
  }

  recipients1 = recipients1.substring(0, recipients1.length-1); //Remove tailing comma
  if (recipients2.length > 0) { recipients2 = recipients2.substring(0, recipients2.length-1); }
  if (recipients3.length > 0) { recipients3 = recipients3.substring(0, recipients3.length-1); }
  if (recipients4.length > 0) { recipients4 = recipients4.substring(0, recipients4.length-1); }
  if (recipients5.length > 0) { recipients5 = recipients5.substring(0, recipients5.length-1); }

  var subject = "[Rides] Logistics for " + sunday;

  // Writing the email message
  var message = "Hi!<br><br>";
  message += "If anything changes, please text Matt at 123-456-7890 or Caitlyn at 123-456-7890.<br><br>"

  if (any_from_b) {
    message += "---------------------------<br><br>";
    message += "<b>Meet at Location B at 10 am.</b><br><br>";
  }

  for (d in drivers) {
    if (drivers[d][2] == 1 && drivers[d][3]) {
      message += "<b>" + drivers[d][0] + ":</b><br>";
      for (i in assignments[drivers[d][0]]) {
        message += assignments[drivers[d][0]][i] + "<br>";
      }
      message += "<br>";
    }
  }

  if (any_from_a) {
    message += "---------------------------<br><br>";
    message += "<b>Meet at Location A at 10 am.</b><br><br>";
  }

  for (d in drivers) {
    if (drivers[d][2] == 0 && drivers[d][3]) {
      message += "<b>" + drivers[d][0] + ":</b><br>";
      for (i in assignments[drivers[d][0]]) {
        message += assignments[drivers[d][0]][i] + "<br>";
      }
      message += "<br>";
    }
  }

  if (any_from_c) {
    message += "---------------------------<br><br>";
    message += "<b>Meet at Location C at 10 am.</b><br><br>";
  }

  for (d in drivers) {
    if (drivers[d][2] == 2 && drivers[d][3]) {
      message += "<b>" + drivers[d][0] + ":</b><br>";
      for (i in assignments[drivers[d][0]]) {
        message += assignments[drivers[d][0]][i] + "<br>";
      }
      message += "<br>";
    }
  }

  if (num_early > 0) {
    message += "---------------------------<br><br>";
    message += "<b>Leaving early on own.</b><br><br>";
  }

  for (i in early_people) {
    message += early_people[i][0] + "<br>";
  }

  if (num_early > 0) {
    message += "<br>";
  }

  message += "---------------------------<br><br>";
  message += "Best,<br>Matt";

  var msgPlain = message.replace(/<br>/g, "\n").replace(/<\/.>/g, ""); // clear html tags for plain mail

  // Sending according to send options
  if (send_option == -1) { // Test Run
    MailApp.sendEmail("a@email.com", "(Test) " + subject, msgPlain, {name: "Matthew Sit", htmlBody: message});
  } else if (send_option == 0) { // Provisional Run
    MailApp.sendEmail("a@email.com, b@email.com", "(Provisional) " + subject, msgPlain, {name: "Matthew Sit", htmlBody: message});
  } else if (send_option == 1) { // Final Run
    MailApp.sendEmail("a@email.com, b@email.com", subject, msgPlain, {name: "Matthew Sit", bcc: recipients1, htmlBody: message});
    if (recipients2.length > 0) { MailApp.sendEmail("a@email.com, b@email.com", subject, msgPlain, {name: "Matthew Sit", bcc: recipients2, htmlBody: message}); }
    if (recipients3.length > 0) { MailApp.sendEmail("a@email.com, b@email.com", subject, msgPlain, {name: "Matthew Sit", bcc: recipients3, htmlBody: message}); }
    if (recipients4.length > 0) { MailApp.sendEmail("a@email.com, b@email.com", subject, msgPlain, {name: "Matthew Sit", bcc: recipients4, htmlBody: message}); }
    if (recipients5.length > 0) { MailApp.sendEmail("a@email.com, b@email.com", subject, msgPlain, {name: "Matthew Sit", bcc: recipients5, htmlBody: message}); }
  }
}

function getMaxScorer(target_score_list, other_score_list1, other_score_list2, tiebreaker) {
  //Logger.log(target_score_list);
  var max_name = "";
  var max_score = -1000;
  for (i in target_score_list) {
    if (target_score_list[i] == 999) { // Driver
      // Skip
    } else if (target_score_list[i] > max_score) {
      max_score = target_score_list[i];
      max_name = i;
    } else if (target_score_list[i] == max_score) {
      var curr_total = other_score_list1[i] + other_score_list2[i];
      var max_total = other_score_list1[max_name] + other_score_list2[max_name];

      if (curr_total < max_total) {
        max_name = i;
      } else if (curr_total == max_total) {
        if (tiebreaker[max_name] < tiebreaker[i]) {
          max_name = i;
        }
      }
    }
  }
  return max_name;
}

function getFriendOrGuest(target_score_list, other_score_list1, other_score_list2, tiebreaker) {
  for (i in target_score_list) {
    if (target_score_list[i] + other_score_list1[i] + other_score_list2[i] == 3) {
      // Scores are -1, -1, 5 = 3.
      return i;
    }
  }
  return getMaxScorer(target_score_list, other_score_list1, other_score_list2, tiebreaker);
}
