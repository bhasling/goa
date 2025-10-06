/******************************************
 * GOA - GMAIL Origanize App.
 * 
 * An application to organize your Gmail Inbox and folders.
 * Source code is in Google Drive: https://drive.google.com
 * ****************************************/

/** Main function to organize emails. */
function organizeEmail() {

  // State object holding all the variables in the state for the application
  var state = {
    goodEmailAddresses: [],
    emailAddressToFolderMap: {},
    folderEmailAddressesMap: {},
    badEmailAddresses: [],
    emailAddressMap: {},
    numberOfStatusEmails: 0,
    previousStatistics: null,
    statusMessageNotRead: false,

    // configuration options
    folderLabels: [],
    batchSize: 20,
    wakeWord: "goa",
    unknownSendersFolder: "UnknownSenders",

    // statistics
    stats: {
      numberOfMessagesWithUnknownSender: 0,
      numberNoContactMessages: 0,
      numberContactMessages: 0,
      numberFolderMessagesMap: {},
      numberEmailsProcessed: 0,
      numberOfEmailThreadsProcessed: 0,
      numberOfActionMessages: 0,
    }
  };
  //debugContact();
  //return;

  // Read and organize the first batchSize messages in the InBox
  readInbox(state);

  // Report the results by sending a summary email
  reportResults(state);
}

/** Read the email in the inbox and classify messages into state. */
function readInbox(state) {
  var returnCode = 0;

  // Loop to restart reading the emails after finding a GOA action email
  for (var j = 0; j < 2; j++) {
    returnCode = 0;

    // Read the persisted properties of the GOA application from Google properties
    readUserProperties(state);

    // Get the most recent threads from the InBox for a batch to process
    var threads = GmailApp.getInboxThreads(0, state.batchSize);

    // Loop through each thread in the batch
    for (var i = 0; i < threads.length; i++) {
      state.stats.numberOfEmailThreadsProcessed = state.stats.numberOfEmailThreadsProcessed + 1;
      // Got an email address thread (which may have several messages)
      var emailThread = threads[i];

      // Process all the messages in the thread
      var messagesInThread = emailThread.getMessages();
      for (var j = 0; j < messagesInThread.length; j++) {
        var msg = messagesInThread[j];
        returnCode = processMessage(state, emailThread, msg);

        // Break from the message loop if we found a GOA action
        if (returnCode) {
          break;
        }
      }

      // Break from the thread loop if we found a GOA action
      if (returnCode) {
        break;
      }
    }

  }
}

/** 
 * Read a single message and classify that message into state.
 * 
 * Return non-zero return code if the message with a GOA action.
 * A non-zero return code causes the main loop to restart reading
 * the InBox after making the changes specified in the action.
 */
function processMessage(state, emailThread, msg) {
  state.stats.numberEmailsProcessed = state.stats.numberEmailsProcessed + 1;

  // Check if this message is a GOA action message
  var from = msg.getFrom();
  var subject = msg.getSubject();
  var returnCode = parseGoaMessage(state, subject, emailThread, msg);
  if (returnCode) {
    return returnCode;
  }

  // Parse out the email address string from the "from" of message.
  var emailAddress = null;
  var match = from.match(/<([^>]+)>/);
  if (match) {
    emailAddress = match[1];
  } else {
    emailAddress = from;
  }

  // Find a GOA configured folder label that is label in the contacts of the email Message
  var folderLabel = state.emailAddressToFolderMap[emailAddress];

  // Classify the message based on the email address contacts
  // Update the statistics of the email and move the email
  // to different label folders if necessary

  if (folderLabel) {
    // Email Address contact is associated with a folderLabel
    applyLabelToEmailMessage(state, emailThread, msg, emailAddress, folderLabel);
    var count = state.stats.numberFolderMessagesMap[folderLabel];
    count = count ? count : 0;
    count = count + 1;
    state.stats.numberFolderMessagesMap[folderLabel] = count;

    //Logger.log(`${emailAddress} is in folder ${folderlabel}`);
  } else if (state.goodEmailAddresses.indexOf(emailAddress) >= 0) {
    // Email Address contact is already determined to be good
    state.stats.numberContactMessages = state.stats.numberContactMessages + 1;
    //Logger.log(`${emailAddress} is in contacts`);
  } else if (state.badEmailAddresses.indexOf(emailAddress) >= 0) {
    // Email Address contact is already determined to be bad
    state.stats.numberNoContactMessages = state.stats.numberNoContactMessages + 1;
    applyLabelToEmailMessage(state, emailThread, msg, emailAddress, state.unknownSendersFolder);
    //Logger.log(`${emailAddress} is not in contacts`);
  } else {
    // Email address contact person has not been seen yet
    classifyEmailContact(state, emailThread, msg, emailAddress);
    //Logger.log(`${emailAddress} has no sender email`);
  }
  return 0;
}

/** Apply a folderLabel to the email thread and remove from the InBox. */
function applyLabelToEmailMessage(state, emailThread, msg, emailAddress, folderLabel) {
  var label = GmailApp.getUserLabelByName(folderLabel) || GmailApp.createLabel(folderLabel);
  var existingLabels = emailThread.getLabels();
  var alreadyLabeled = existingLabels.some(function (l) {
    return l.getName() === label.getName();
  });
  if (!alreadyLabeled) {
    emailThread.addLabel(label);
    emailThread.moveToArchive();
  } else {
    emailThread.moveToArchive();
  }
}

/** Classify the emailAddress by the contact information about the emailAddress and store classification in state. */
function classifyEmailContact(state, emailThread, msg, emailAddress) {
  // Get the contact information about the email address
  var contactInfoList = getEmailAddressContactInfo(emailAddress);
  if (contactInfoList) {
    //  If the email addess is in the users contact list find groups of that contact
    groups = getContactGroups(state, contactInfoList);

    // Look and the GOA configured folder labels to see if user is in that group
    for (var i = 0; i < state.folderLabels.length; i++) {
      var folderLabel = state.folderLabels[i];
      if (groups.indexOf(folderLabel) >= 0) {
        // the contact of the email address is in this GOA folder group
        // Save that this email address is in this folder group
        state.folderEmailAddressesMap[emailAddress] = folderLabel;

        // Move the email to the folder group and update the GOA statistics
        var count = state.stats.numberFolderMessagesMap[folderLabel];
        count = count ? count : 0;
        count = count + 1;
        state.stats.numberFolderMessagesMap[folderLabel] = count;
        applyLabelToEmailMessage(state, emailThread, msg, emailAddress, folderLabel);

        // Return because we handled this email message
        return;
      }
    }

    // If the contact is not in any of the folder groups it is considered good
    // because this is an email in the contact list, just not moved to a folder
    state.goodEmailAddresses.push(emailAddress);
    state.stats.numberContactMessages = state.stats.numberContactMessages + 1;

  } else {
    // The email is not associated with a contact
    // Save the email address as not having a contact and update statistics
    state.badEmailAddresses.push(emailAddress);
    state.stats.numberNoContactMessages = state.stats.numberNoContactMessages + 1;
    // Move this email thread to the unknown senders folder
    applyLabelToEmailMessage(state, emailThread, msg, emailAddress, state.unknownSendersFolder);
  }
}

/** 
 * Get the contact information list for the email address.
 * 
 * Look for the contact in the main Google People contact list.
 * If it is not in that contact list look in the OtherContacts list.
 * The OtherContacts list is an old obsolete version of Google Contacts
 * but this may still contain some old contacts that are still valid.
 */
function getEmailAddressContactInfo(emailAddress) {
  // Get a list of contacts associated with the emailAddress
  var result = null;
  var response = People.People.searchContacts({
    query: emailAddress,
    readMask: "emailAddresses,memberships"
  });
  if (response && response.results && response.results.length > 0) {
    result = response.results;
  } else {
    // No contact found. Lets check in OtherContacts before we give up.
    // But Other Contacts returns a single contact and not an array
    response = People.OtherContacts.search({
      query: emailAddress,
      readMask: "emailAddresses,names"
    });
    // If we got a response set it to null if it is empty or change it a list
    if (response && response.person) {
      result = [responce.person];
    }
  }
  return result;

}

/** 
 * Get the list folder label names for any label group in the users contactInfo list.
 * 
 * The contactInfoList is a list that comes from the Google Contact app.
 * However this just contains membership IDs. These are looked up to get the label name.
 */
function getContactGroups(state, contactInfoList) {
  var result = [];
  for (var j = 0; j < contactInfoList.length; j++) {
    contactInfo = contactInfoList[j];
    var memberships = contactInfo.person.memberships;
    for (var i = 0; i < memberships.length; i++) {
      var membership = memberships[i];
      var groupMembership = membership.contactGroupMembership;
      var groupId = groupMembership.contactGroupResourceName;
      var groupName = state.contactLabelMap[groupId];
      result.push(groupName);
    }
  }
  return result;
}

/** Create a map in state from a contact label ID to the name of the contract label. 
 * 
 * This returns a map of contact ID to a contact label name.
 * It looks this up using the Google People.ContactGroups service.
 */
function readUserProperties(state) {
  state.contactLabelMap = {};
  var response = People.ContactGroups.list({
    pageSize: 100
  });

  if (response.contactGroups) {
    response.contactGroups.forEach(function (group) {
      // Add an entry to the map to map the ID to the label name
      state.contactLabelMap[group.resourceName] = group.name;
    })
  }

  const userProps = PropertiesService.getUserProperties();
  const folderJson = userProps.getProperty("gmail_apps.folder_labels");
  if (folderJson) {
    state.folderLabels = JSON.parse(folderJson);
  } else {
    state.folderLables = [];
  }

  var statsJson = userProps.getProperty("gmail_apps.previousStatistics");
  if (statsJson) {
    state.previousStatistics = JSON.parse(statsJson);
  }

  const batchSizeString = userProps.getProperty("gmail_apps.batchSize");
  if (batchSizeString) {
    state.batchSize = parseInt(batchSizeString);
  }
  state["myemail"] = Session.getActiveUser().getEmail();
}

/** Replaces the list of folder labels in the Google configuration. */
function initializeFolderLabels(labels) {
  const userProps = PropertiesService.getUserProperties();
  userProps.setProperty("gmail_apps.folder_labels", JSON.stringify(labels));
}

/** Find and parse a Goa action message in an email message.
 * 
 * This looka at the subject of the emailThread to see if it starts
 * with the wake up word and then parses the action specified in the
 * email message.
 * 
 * Returns 0 if the message is not an action or if the action was handled.
 * Returns non zero if the message was an action and applies the action.
 *
 * Returning non-zero causes the main loop to restart and start processing
 * InBox messages using the new action changes that are now stored in state.
 */
function parseGoaMessage(state, subject, emailThread, msg) {
  if (subject.includes(state.wakeWord.toUpperCase() + " Organize Email")) {
    if (msg.isUnread()) {
      state.statusMessageNotRead = true;
    }
    emailThread.moveToTrash();
    return 0;
  } else {
    if (subject.toLowerCase().includes(state.wakeWord.toLowerCase())) {
      var parts = subject.split(" ");
      if (parts.length > 3) {
        if (parts[0].toLowerCase() == state.wakeWord.toLowerCase()) {
          if (["add", "enable", "insert"].includes(parts[1].toLowerCase())) {
            if (["folder", "label"].includes(parts[2].toLowerCase())) {
              var folder = parts[3];
              if (!state.folderLabels.includes(folder)) {
                state.folderLabels.push(folder);
                initializeFolderLabels(state.folderLabels);
                emailThread.moveToTrash();
              }
              state.stats.numberOfActionMessages += 1;
              return -1;
            }
          } else if (["delete", "remove", "disable"].includes(parts[1].toLowerCase())) {
            if (["folder", "label"].includes(parts[2].toLowerCase())) {
              var folder = parts[3];
              state.folderLabels = state.folderLabels.filter(item => item !== folder);
              initializeFolderLabels(state.folderLabels);
              emailThread.moveToTrash();
              state.stats.numberOfActionMessages += 1;
              return -1;
            }
          } else if (["set", "assign"].includes(parts[1].toLowerCase())) {
            if (["batch", "size"].includes(parts[2].toLowerCase())) {
              var batchSize = parseInt(parts[3]);
              state.batchSize = batchSize;
              const userProps = PropertiesService.getUserProperties();
              userProps.setProperty("gmail_apps.batchSize", batchSize.toString());
              emailThread.moveToTrash();
              state.stats.numberOfActionMessages += 1;
              return -1;
            }
          }
        }
      }
    }
  }
  return 0;
}

function saveStatistics(state) {
  const userProps = PropertiesService.getUserProperties();
  userProps.setProperty("gmail_apps.previousStatistics", JSON.stringify(state.stats));

}

function addPreviousStats(state) {
  if (state.previousStatistics) {
    state.stats.numberOfMessagesWithUnknownSender += state.previousStatistics.numberOfMessagesWithUnknownSender;
    state.stats.numberNoContactMessages += state.previousStatistics.numberNoContactMessages;
    state.stats.numberEmailsProcessed += state.previousStatistics.numberEmailsProcessed;
    state.stats.numberOfEmailThreadsProcessed += state.previousStatistics.numberOfEmailThreadsProcessed;
    state.stats.numberOfActionMessages += state.previousStatistics.numberOfActionMessages;
    if (state.previousStatistics.numberFolderMessagesMap) {
      var keys = Object.keys(state.previousStatistics.numberFolderMessagesMap);
      for (var i = 0; i < keys.length; i++) {
        var key = keys[i];
        var oldCount = state.previousStatistics.numberFolderMessagesMap[key];
        var newCount = state.stats.numberFolderMessagesMap[key];
        newCount = newCount ? newCount : 0;
        newCount += oldCount;
        state.stats.numberFolderMessagesMap[key] = newCount;
      }
    }
  }
}
/** 
 * Report statistics about results.
 * 
 * Format and send an email message about the newly organized email changes
 * and instructions how to send email back to Goa to perform actions.
 */
function reportResults(state) {
  if (state.statusMessageNotRead) {
    addPreviousStats(state);
  } else {
  }
  saveStatistics(state);

  var htmlBody = `The organizeEmail Google App in your account read ${state.stats.numberEmailsProcessed} email messsages (${state.stats.numberOfEmailThreadsProcessed} threads).\n`;
  htmlBody = htmlBody + "<ul>\n";
  htmlBody = htmlBody + `<li>${state.stats.numberContactMessages} messages from your contacts (left in InBox).</li>\n`
  htmlBody = htmlBody + `<li>${state.stats.numberNoContactMessages} messages from no contact senders (moved to ${state.unknownSendersFolder} folder).</li>\n`
  htmlBody = htmlBody + `<li>${state.stats.numberOfActionMessages} action messages.</li>\n`
  htmlBody = htmlBody + `<li>${state.stats.numberOfMessagesWithUnknownSender} messages from senders with no email address (moved to ${state.unknownSendersFolder} folder).</li>\n`;
  const keys = Object.keys(state.stats.numberFolderMessagesMap);
  for (var i = 0; i < keys.length; i++) {
    var key = keys[i];
    var count = state.stats.numberFolderMessagesMap[key]
    htmlBody = htmlBody + `<li>${count} messages moved to ${key} folder.</li>\n`
  }
  var folderLabelsList = JSON.stringify(state.folderLabels);
  var wake = state.wakeWord.toUpperCase();
  htmlBody = htmlBody + "<p></p>";
  htmlBody = htmlBody + `<li>Folder Labels: ${folderLabelsList}</li>\n`;
  htmlBody = htmlBody + `<li>Batch Size: ${state.batchSize}</li>\n`;
  htmlBody = htmlBody + "</ul><p></p>\n";
  htmlBody = htmlBody + `Send email to ${state.myemail} with a configuration action in subject.\n`;
  htmlBody = htmlBody + `<li>${wake} add folder [folder]</li>\n`;
  htmlBody = htmlBody + `<li>${wake} remove folder [folder]</li>\n`;
  htmlBody = htmlBody + `<li>${wake} set batch [#messageReadPerTime]</li>\n`;
  var subject = `${wake} Organize Email`;
  GmailApp.sendEmail(state.myemail, subject, "", { htmlBody: htmlBody });
  Logger.log(htmlBody);
}

/** Debug function (not used) to check the contact information about an email address. */
function debugContact() {
  var emailAddress = "tshannontopspin@verizon.net";
  var contactInfoList = getEmailAddressContactInfo(emailAddress);
  Logger.log(contactInfoList);
}
