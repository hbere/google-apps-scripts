/**
 * "Vote for the best Google Form response(s) in columns D & E"
 *
 * Description:
 * This script helps users vote-up the best responses to an existing Google Form.
 * From a spreadsheet tracking the existing Google Form's responses, the script
 * will generate a new Google Form asking which of the values from columns
 * D and E were most valuable.
 *
 * Requirements:
 * This code is designed for use in a Google spreadsheet like the following:
 * (A) Tracking responses from a Google Form.
 * (B) Form responses tracking sheet is entitled 'Form Responses 1'.
 * (C) Columns D and E include comments.
 *
 * Usage:
 * (A) Add this code to a Google spreadsheet meeting the requirements above.
 * (B) Re-open the spreadsheet.  Make sure you are viewing the sheet about
 * which you wish to re-survey.
 * (C) From the "survey" menu, choose the option "Set up new survey."
 *
 * Acknowledgements:
 * This script was developed starting from the following example:
 * https://developers.google.com/apps-script/quickstart/forms
 * The example code was licensed under the following:
 * http://www.apache.org/licenses/LICENSE-2.0
 * Thank you to the example script's authors!
 *
 * Last updated: 9/8/17
 */

/**
 * A special function to run when the add-on is installed.
 * Based on the following example: https://developers.google.com/apps-script/add-ons/lifecycle
 * Visit that URL for more information
 */
function onInstall(e) {
  onOpen(e);
  // Perform additional setup as needed.
}

/**
 * A special function that inserts a custom menu when the spreadsheet opens.
 */
function onOpen(e) {
  var menu = [{name: 'Set up new survey', functionName: 'setUpSurvey_'}]
  SpreadsheetApp.getActive().addMenu('Survey', menu);
}

/**
 * A set-up function that uses the survey data in the spreadsheet to create
 * a Google Form and a trigger that allows the script to react to form
 * responses.
 */
function setUpSurvey_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Form Responses 1');
  var range = sheet.getDataRange();
  var values = range.getValues();
  setUpForm_(ss, values);
}

/**
 * Creates a Google Form that allows respondents to select which comments
 * they found most useful, out of the comments listed in column C in the
 * source survey data.
 */
function setUpForm_(ss, values) {
  // get spreadsheet name
  var spreadsheetName = ss.getName();

  // get current date and time
  var d = new Date();
  var currentTime = d.toLocaleTimeString();

  // put "like" comments into an array
  // Comments in column C?  Use 'values[i][2].'  D?  Use 'values[i][3].'   Etc.
  var likeComments = [];
  for (var i = 0; i < values.length; i++) {
    if (i > 0) {
      likeComments[i-1] = (values[i][3]);
    }
  }

  // put "dislike" comments into an array
  // Comments in column D?  Use 'values[i][3].'  E?  Use 'values[i][4].'   Etc.
  var dislikeComments = [];
  for (var i = 0; i < values.length; i++) {
    if (i > 0) {
      dislikeComments[i-1] = (values[i][4]);
    }
  }

  // Create the form and add a multiple-choice question
  var form = FormApp.create('Survey Response Voting');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId())
    .setPublishingSummary(true);
  var item = form.addSectionHeaderItem();
  item.setTitle(spreadsheetName + ' @ ' + currentTime);
  var item = form.addCheckboxItem().setTitle("A. Which of the following comments do you find most valuable? Choose three (3).")
    .setChoiceValues(likeComments);
  var item = form.addCheckboxItem().setTitle("B. Which of the following comments do you find most valuable? Choose three (3).")
    .setChoiceValues(dislikeComments);

  // Log the new form's URL
  Logger.log('Published URL: ' + form.getPublishedUrl());
  Logger.log('Editor URL: ' + form.getEditUrl());
}
