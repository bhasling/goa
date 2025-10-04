/******************************************
 * GOA - GMAIL Origanize App.
 * 
 * An application to organize your Gmail Inbox and folders.
 * ****************************************/

/** Main function to origanize email. */
function organizeEmail() {
  var state = {
    goodEmailAddresses: [],
    emailAddressToFolderMap: {},
    folderEmailAddressesMap: {},
    numberNoContactMessages: 0,
    numberContactMessages: 0,
    numberFolderMessagesMap: {},
    numberEmailsProcessed: 0,
    numberofEmailThreadsProcessed: 0,
    badEmailAddresses: [],
    emailAddressMap: {},
    folderLabels: [],
    numberOfStatusEmails: 0,
    batchSize: 2,
    numberOfMessagesWithUnknownSender: 0
  };
  readInbox(state);
  reportResults(state);
}

/** Read the email in the inbox and classify messages into state. */
function readInbox(state) {
  var returnCode = 0;
  for (var j = 0; j < 2; j++) {
    returnCode = 0;
    readUserProperties(state);
    Logger.log(`Read inbox with batch size ${state.batchSize}`);
    var threads = GmailApp.getInboxThreads(0, state.batchSize);

    // Loop through each thread
    for (var i = 0; i < threads.length; i++) {
      state.numberofEmailThreadsProcessed = state.numberofEmailThreadsProcessed + 1;
      // Got an email address thread (which may have several messages)
      var emailThread = threads[i];
      var messagesInThread = emailThread.getMessages();
      for (var j = 0; j < messagesInThread.length; j++) {
        var msg = messagesInThread[j];
        returnCode = processMessage(state, emailThread, msg);
        if (returnCode) {
          Logger.log("Non zero return code so try again");
          break;
        }
      }
      if (returnCode) {
        break;
      }
    }

  }
}

/** Read a single message and classify that message into state.*/
function processMessage(state, emailThread, msg) {
  state.numberEmailsProcessed = state.numberEmailsProcessed + 1;
  var from = msg.getFrom();
  var subject = msg.getSubject();
  var returnCode = parseGoaMessage(state, subject, emailThread);
  if (returnCode) {
    return returnCode;
  }
  var match = from.match(/<([^>]+)>/);
  var emailAddress = null;
  if (match) {
    emailAddress = match[1];
  } else {
    emailAddress = from;
  }
  var folderLabel = state.emailAddressToFolderMap[emailAddress];
  if (folderLabel) {
    // Email Address contact is associated with a folderLabel
    applyLabelToEmailMessage(state, emailThread, msg, emailAddress, folderLabel);
    var count = state.numberFolderMessagesMap[folderLabel];
    count = count ? count : 0;
    count = count + 1;
    state.numberFolderMessagesMap[folderLabel] = count;

    //Logger.log(`${emailAddress} is in folder ${folderlabel}`);
  } else if (state.goodEmailAddresses.indexOf(emailAddress) >= 0) {
    // Email Address contact is already determined to be good
    state.numberContactMessages = state.numberContactMessages + 1;
    //Logger.log(`${emailAddress} is in contacts`);
  } else if (state.badEmailAddresses.indexOf(emailAddress) >= 0) {
    // Email Address contact is already determined to be bad
    state.numberNoContactMessages = state.numberNoContactMessages + 1;
    applyLabelToEmailMessage(state, emailThread, msg, emailAddress, "UnknownSenders");
    //Logger.log(`${emailAddress} is not in contacts`);
  } else {
    // Email address contact person has not been seen yet
    classifyEmailContact(state, emailThread, msg, emailAddress);
    //Logger.log(`${emailAddress} has no sender email`);
  }
  return 0;
}

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
    Logger.log(`${emailAddress} is already labelled with ${folderLabel}`);
    emailThread.moveToArchive();
  }
  Logger.log(`Moved email from ${emailAddress} to ${folderLabel}`);
}

/** Classify the emailAddress by the contact information about the emailAddress and store classification in state. */
function classifyEmailContact(state, emailThread, msg, emailAddress) {
  var contactInfoList = getEmailAddressContactInfo(emailAddress);
  if (contactInfoList) {
    groups = getContactGroups(state, contactInfoList);
    for (var i = 0; i < state.folderLabels.length; i++) {
      var folderLabel = state.folderLabels[i];
      if (groups.indexOf(folderLabel) >= 0) {
        // the contact of the email address is in this folder group
        state.folderEmailAddressesMap[emailAddress] = folderLabel;
        var count = state.numberFolderMessagesMap[folderLabel];
        count = count ? count : 0;
        count = count + 1;
        state.numberFolderMessagesMap[folderLabel] = count;
        applyLabelToEmailMessage(state, emailThread, msg, emailAddress, folderLabel);

        return;
      }
    }
    // If the contact is not in any of the folder groups it is still good
    state.goodEmailAddresses.push(emailAddress);
    state.numberContactMessages = state.numberContactMessages + 1;
  } else {
    state.badEmailAddresses.push(emailAddress);
    state.numberNoContactMessages = state.numberNoContactMessages + 1;
    applyLabelToEmailMessage(state, emailThread, msg, emailAddress, "UnknownSenders");
  }
  return;

}

function getEmailAddressContactInfo(emailAddress) {
  var response = People.People.searchContacts({
    query: emailAddress,
    readMask: "emailAddresses,memberships"
  });
  var result = null;
  if (response.results && response.results.length > 0) {
    result = response.results;
  }
  return result;

}

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

/** Create a map in state from a contract label ID to the name of the contract label. */
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
  Logger.log("Read properties for folderJson");
  Logger.log(folderJson);
  if (folderJson) {
    state.folderLabels = JSON.parse(folderJson);
  } else {
    state.folderLables = [];
  }

  const batchSizeString = userProps.getProperty("gmail_apps.batchSize");
  if (batchSizeString) {
    state.batchSize = parseInt(batchSizeString);
  }
  state["myemail"] = Session.getActiveUser().getEmail();
}

function initializeFolderLabels(labels) {
  const userProps = PropertiesService.getUserProperties();
  userProps.setProperty("gmail_apps.folder_labels", JSON.stringify(labels));
}

function parseGoaMessage(state, subject, emailThread) {
  if (subject.includes("GOA Organize Email")) {
    Logger.log("Found GOA status email");
    state.numberOfStatusEmails = state.numberOfStatusEmails + 1;
    if (state.numberOfStatusEmails > 2) {
      Logger.log("Too many GOA status messages");
      emailThread.moveToTrash();
      return 0;
    }
  } else {
    if (subject.toLowerCase().includes("goa")) {
      var parts = subject.split(" ");
      if (parts.length > 3) {
        Logger.log(parts);
        if (parts[0].toLowerCase() == "goa") {
          Logger.log("Found GOA");
          if (["add", "enable", "insert"].includes(parts[1].toLowerCase())) {
            if (["folder", "label"].includes(parts[2].toLowerCase())) {
              var folder = parts[3];
              if (!state.folderLabels.includes(folder)) {
                state.folderLabels.push(folder);
                initializeFolderLabels(state.folderLabels);
                Logger.log("added folder " + folder);
                emailThread.moveToTrash();
              }
              return -1;
            }
          } else if (["delete", "remove", "disable"].includes(parts[1].toLowerCase())) {
            Logger.log("Found remove");
            if (["folder", "label"].includes(parts[2].toLowerCase())) {
              var folder = parts[3];
              state.folderLabels = state.folderLabels.filter(item => item !== folder);
              initializeFolderLabels(state.folderLabels);
              Logger.log("deleted folder " + folder);
              emailThread.moveToTrash();
              return -1;
            }
          } else if (["set", "assign"].includes(parts[1].toLowerCase())) {
            if (["batch", "size"].includes(parts[2].toLowerCase())) {
              var batchSize = parseInt(parts[3]);
              state.batchSize = batchSize;
              const userProps = PropertiesService.getUserProperties();
              userProps.setProperty("gmail_apps.batchSize", batchSize.toString());
              Logger.log(`Set batch size ${state.batchSize}`);
              emailThread.moveToTrash();
              return -1;
            }
          }
        }
      }
    }
  }
  return 0;
}

/** Report statistics about results */
function reportResults(state) {
  var htmlBody = `The organizeEmail Google App in your account read ${state.numberEmailsProcessed} email messsages (${state.numberofEmailThreadsProcessed} threads).\n`;
  htmlBody = htmlBody + "<ul>\n";
  htmlBody = htmlBody + `<li>${state.numberContactMessages} messages from your contacts (left in InBox).</li>\n`
  htmlBody = htmlBody + `<li>${state.numberNoContactMessages} messages from no contact senders (moved to UnknownSenders folder).</li>\n`
  htmlBody = htmlBody + `<li>${state.numberOfMessagesWithUnknownSender} messages from senders with no email address (moved to UnknownSenders folder).</li>\n`;
  const keys = Object.keys(state.numberFolderMessagesMap);
  for (var i = 0; i < keys.length; i++) {
    var key = keys[i];
    var count = state.numberFolderMessagesMap[key]
    htmlBody = htmlBody + `<li>${count} messages moved to ${key} folder.</li>\n`
  }
  var folderLabelsList = JSON.stringify(state.folderLabels);
  htmlBody = htmlBody + "<p></p>";
  htmlBody = htmlBody + `<li>Folder Labels: ${folderLabelsList}</li>\n`;
  htmlBody = htmlBody + `<li>Batch Size: ${state.batchSize}</li>\n`;
  htmlBody = htmlBody + "</ul><p></p>\n";
  htmlBody = htmlBody + `Send email to ${state.myemail} with a configuration action in subject.\n`;
  htmlBody = htmlBody + `<li>GOA add folder [folder]</li>\n`;
  htmlBody = htmlBody + `<li>GOA remove folder [folder]</li>\n`;
  htmlBody = htmlBody + `<li>GOA set batch [#messageReadPerTime]</li>\n`;
  var subject = "GOA Organize Email";
  GmailApp.sendEmail(state.myemail, subject, "", { htmlBody: htmlBody });
  Logger.log(htmlBody);
}
