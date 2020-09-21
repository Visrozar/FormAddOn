/**
 * @OnlyCurrentDoc
 *
 * The above comment directs Apps Script to limit the scope of file
 * access for this add-on. It specifies that this add-on will only
 * attempt to read or modify the files in which the add-on is used,
 * and not all of the user's files. The authorization request message
 * presented to users will reflect this limited scope.
 */

/**
 * A global constant String holding the title of the add-on. This is
 * used to identify the add-on in multiple places.
 */
var ADDON_TITLE = 'FormConnector';

/**
 * A global constant 'notice' text to include wherever necessary.
 */
var NOTICE = `${ADDON_TITLE} is meant for connecting your form with the ${ADDON_TITLE} application.
  The ${ADDON_TITLE} application allow you to create awesome reports based on the data you receive from Google Forms.
  That is why it is called Forms on Steroids!`;

/**
 * Adds a custom menu to the active form to show the add-on sidebar.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(e) {
  FormApp.getUi()
    .createAddonMenu()
    .addItem(`Interact with ${ADDON_TITLE}`, 'showSidebar')
    .addItem('About', 'showAbout')
    .addToUi();
}

/**
 * Runs when the add-on is installed.
 *
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE).
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Opens a sidebar in the form containing the add-on's user interface for
 * configuring the notifications this add-on will produce.
 */
function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('sidebar')
    .setWidth(420)
    .setHeight(270);
  FormApp.getUi().showModalDialog(ui, `Connect with ${ADDON_TITLE}`);
}

/**
 * Opens a purely-informational dialog in the form explaining details about
 * this add-on.
 */
function showAbout() {
  var ui = HtmlService.createHtmlOutputFromFile('about')
    .setWidth(420)
    .setHeight(270);
  FormApp.getUi().showModalDialog(ui, `About ${ADDON_TITLE}`);
}

/**
 * Save sidebar settings to this form's Properties, and update the onFormSubmit
 * trigger as needed.
 *
 * @param {Object} settings An Object containing key-value
 *      pairs to store.
 */
function saveSettings(property, value) {
  console.log(property, value)
  PropertiesService.getDocumentProperties().setProperty(property, value);
  let actionsNeeded = adjustFormSubmitTrigger();
  if(actionsNeeded.deleteExistingFormData) deleteExistingFormData()
  if(actionsNeeded.recreateExistingFormData) recreateExistingFormData()
}

/**
 * Queries the User Properties and adds additional data required to populate
 * the sidebar UI elements.
 *
 * @return {Object} A collection of Property values and
 *     related data used to fill the configuration sidebar.
 */
function getSettings() {
  var settings = PropertiesService.getDocumentProperties().getProperties();
  return settings;
}

/**
 * Adjust the onFormSubmit trigger based on user's requests.
 */
function adjustFormSubmitTrigger() {
  let form = FormApp.getActiveForm();
  let triggers = ScriptApp.getUserTriggers(form);
  let settings = PropertiesService.getDocumentProperties();
  let triggerNeeded = settings.getProperty('connect');
  let response = {
    deleteExistingFormData: false,
    recreateExistingFormData: false,
  }

  // Create a new trigger if required; delete existing trigger
  //   if it is not needed.
  let existingTrigger = null;
  console.log('Triggers Found ', triggers.length)
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getEventType() == ScriptApp.EventType.ON_FORM_SUBMIT) {
      existingTrigger = triggers[i];
      break;
    }
  }
  if (triggerNeeded === 'true' && !existingTrigger) {
    // if new trigger needed
    // create new trigger
    let trigger = ScriptApp.newTrigger('respondToFormSubmit')
      .forForm(form)
      .onFormSubmit()
      .create();
    // store all the existing formData in DB
    response.deleteExistingFormData = true;
    response.recreateExistingFormData = true;
  } else if (triggerNeeded === 'false' && existingTrigger) {
    // If trigger exists, delete the trigger
    ScriptApp.deleteTrigger(existingTrigger);
    // delete all related formData from DB
    response.deleteExistingFormData = true;
  }
  return response;
}

function recreateExistingFormData() {
  const formData = getEntireFormData();
  storeDataToDB(formData)

}

function getEntireFormData() {
  let form = FormApp.getActiveForm();
  let responses = form.getResponses();
  let dbResponses = [];
  let fbResponse = {
    formId: form.getId(),
    title: form.getTitle(),
    description: form.getDescription()
  };
  responses.forEach(response => {
    let responseItems = response.getItemResponses();
    responseItems.forEach(responseItem => {
      let dbResponse = {
        formId: form.getId(),
        responseId: response.getId(),
        emailId: response.getRespondentEmail(),
        title: responseItem.getItem().getTitle(),
        type: responseItem.getItem().getType().name(),
        value: responseItem.getResponse(),
        itemId: responseItem.getItem().getId(),
      }

      fbResponse[dbResponse.itemId] = {
        title: dbResponse.title,
        labels: [],
        graphType: '',
        availableGraphType: [],
        data: []
      };
      if (dbResponse.type === 'GRID') {
        fbResponse[dbResponse.itemId].labels = responseItem.getItem().asGridItem().getRows();
      } else if (dbResponse.type === 'CHECKBOXGRID') {
        fbResponse[dbResponse.itemId].labels = responseItem.getItem().asCheckboxGridItem().getRows();
      }
      dbResponses.push(dbResponse);
    })
  })
  return {dbResponses, fbResponse}
}

/**
 * Responds to a form submission event if an onFormSubmit trigger has been
 * enabled.
 *
 * @param {Object} e The event parameter created by a form
 *      submission; see
 *      https://developers.google.com/apps-script/understanding_events
 */
function respondToFormSubmit(e) {
  var settings = PropertiesService.getDocumentProperties();
  var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);

  // Check if the actions of the trigger require authorizations that have not
  // been supplied yet -- if so, warn the active user via email (if possible).
  // This check is required when using triggers with add-ons to maintain
  // functional triggers.
  if (authInfo.getAuthorizationStatus() ==
    ScriptApp.AuthorizationStatus.REQUIRED) {
    // Re-authorization is required. In this case, the user needs to be alerted
    // that they need to reauthorize; the normal trigger action is not
    // conducted, since authorization needs to be provided first. Send at
    // most one 'Authorization Required' email a day, to avoid spamming users
    // of the add-on.
    sendReauthorizationRequest();
  } else {
    // All required authorizations have been granted, so continue to respond to
    // the trigger event.

    // Check if the form creator needs to be notified; if so, construct and
    // send the notification.
    if (settings.getProperty('creatorNotify') == 'true') {
      sendCreatorNotification();
    }

    // Check if the form respondent needs to be notified; if so, construct and
    // send the notification. Be sure to respect the remaining email quota.
    if (settings.getProperty('respondentNotify') == 'true' &&
      MailApp.getRemainingDailyQuota() > 0) {
      sendRespondentNotification(e.response);
    }
  }
}


/**
 * Called when the user needs to reauthorize. Sends the user of the
 * add-on an email explaining the need to reauthorize and provides
 * a link for the user to do so. Capped to send at most one email
 * a day to prevent spamming the users of the add-on.
 */
//function sendReauthorizationRequest() {
//  var settings = PropertiesService.getDocumentProperties();
//  var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
//  var lastAuthEmailDate = settings.getProperty('lastAuthEmailDate');
//  var today = new Date().toDateString();
//  if (lastAuthEmailDate != today) {
//    if (MailApp.getRemainingDailyQuota() > 0) {
//      var template =
//          HtmlService.createTemplateFromFile('authorizationEmail');
//      template.url = authInfo.getAuthorizationUrl();
//      template.notice = NOTICE;
//      var message = template.evaluate();
//      MailApp.sendEmail(Session.getEffectiveUser().getEmail(),
//          'Authorization Required',
//          message.getContent(), {
//            name: ADDON_TITLE,
//            htmlBody: message.getContent()
//          });
//    }
//    settings.setProperty('lastAuthEmailDate', today);
//  }
//}

/**
 * Sends out creator notification email(s) if the current number
 * of form responses is an even multiple of the response step
 * setting.
 */
//function sendCreatorNotification() {
//  var form = FormApp.getActiveForm();
//  var settings = PropertiesService.getDocumentProperties();
//  var responseStep = settings.getProperty('responseStep');
//  responseStep = responseStep ? parseInt(responseStep) : 10;
//
//  // If the total number of form responses is an even multiple of the
//  // response step setting, send a notification email(s) to the form
//  // creator(s). For example, if the response step is 10, notifications
//  // will be sent when there are 10, 20, 30, etc. total form responses
//  // received.
//  if (form.getResponses().length % responseStep == 0) {
//    var addresses = settings.getProperty('creatorEmail').split(',');
//    if (MailApp.getRemainingDailyQuota() > addresses.length) {
//      var template =
//          HtmlService.createTemplateFromFile('creatorNotification');
//      template.summary = form.getSummaryUrl();
//      template.responses = form.getResponses().length;
//      template.title = form.getTitle();
//      template.responseStep = responseStep;
//      template.formUrl = form.getEditUrl();
//      template.notice = NOTICE;
//      var message = template.evaluate();
//      MailApp.sendEmail(settings.getProperty('creatorEmail'),
//          form.getTitle() + ': Form submissions detected',
//          message.getContent(), {
//            name: ADDON_TITLE,
//            htmlBody: message.getContent()
//          });
//    }
//  }
//}

/**
 * Sends out respondent notification emails.
 *
 * @param {FormResponse} response FormResponse object of the event
 *      that triggered this notification
 */
//function sendRespondentNotification(response) {
//  var form = FormApp.getActiveForm();
//  var settings = PropertiesService.getDocumentProperties();
//  var emailId = settings.getProperty('respondentEmailItemId');
//  var emailItem = form.getItemById(parseInt(emailId));
//  var respondentEmail = response.getResponseForItem(emailItem)
//      .getResponse();
//  if (respondentEmail) {
//    var template =
//        HtmlService.createTemplateFromFile('respondentNotification');
//    template.paragraphs = settings.getProperty('responseText').split('\n');
//    template.notice = NOTICE;
//    var message = template.evaluate();
//    MailApp.sendEmail(respondentEmail,
//        settings.getProperty('responseSubject'),
//        message.getContent(), {
//          name: form.getTitle(),
//            htmlBody: message.getContent()
//        });
//  }
//}