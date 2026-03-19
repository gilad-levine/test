function onEditCustom(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  // Configuration
  var MONITORED_COLUMNS = [22, 23, 24]; // Columns V, W, and X
  var SHEET_NAME = 'Main'; // Change to your sheet name
  var HEADER_ROW = 1;
  var props = PropertiesService.getScriptProperties();
  var JOTFORM_API_KEY = props.getProperty('JOTFORM_API_KEY');
  var JOTFORM_TABLE_ID = props.getProperty('JOTFORM_TABLE_ID');
  var EVENT_NAME = props.getProperty('EVENT_NAME') || 'DeveloperWeek 2026';
  var SHORTIO_API_KEY = props.getProperty('SHORTIO_API_KEY');
  var SHORTIO_DOMAIN = props.getProperty('SHORTIO_DOMAIN');
  if (sheet.getName() !== SHEET_NAME) return;
  var editedColumn = range.getColumn();
  var row = range.getRow();
  if (row === HEADER_ROW) return;
  // --- UNIQUE ID LOGIC (Column A → Column AS) ---
  if (editedColumn === 1 && e.value && e.value.toString().trim() !== '') {
    handleUniqueId(sheet, row);
  }
  // --- EXISTING JOTFORM LOGIC (Columns V, W, X) ---
  if (MONITORED_COLUMNS.indexOf(editedColumn) !== -1) {
    handleJotFormSync(sheet, row, EVENT_NAME, JOTFORM_TABLE_ID, JOTFORM_API_KEY, SHORTIO_API_KEY, SHORTIO_DOMAIN);
  }
  // --- HUBSPOT LOGIC (Trigger: Checkbox in Column AQ/43) ---
  if (editedColumn === 43 && e.value === "TRUE") {
    handleHubSpotSync(sheet, row);
  }
}

// Generates and writes a unique 7-digit ID to column AS (45) if one doesn't already exist
function handleUniqueId(sheet, row) {
  var ID_COLUMN = 45; // Column AS
  var existingId = sheet.getRange(row, ID_COLUMN).getValue();
  if (existingId && existingId.toString().trim() !== '') {
    // ID already exists — do nothing
    return;
  }
  // Collect all existing IDs in column AS to ensure uniqueness
  var lastRow = sheet.getLastRow();
  var existingIds = {};
  if (lastRow > 1) {
    var allIds = sheet.getRange(2, ID_COLUMN, lastRow - 1, 1).getValues();
    for (var i = 0; i < allIds.length; i++) {
      var val = allIds[i][0];
      if (val && val.toString().trim() !== '') {
        existingIds[val.toString()] = true;
      }
    }
  }
  // Generate a unique 7-digit number (1000000–9999999)
  var newId;
  do {
    newId = Math.floor(1000000 + Math.random() * 9000000).toString();
  } while (existingIds[newId]);
  sheet.getRange(row, ID_COLUMN).setValue(newId);
  Logger.log('Assigned unique ID ' + newId + ' to row ' + row);
}

// Factor out JotForm logic for clarity
function handleJotFormSync(sheet, row, EVENT_NAME, JOTFORM_TABLE_ID, JOTFORM_API_KEY, SHORTIO_API_KEY, SHORTIO_DOMAIN) {
  // Get values from the edited row
  var groupId = sheet.getRange(row, 45).getValue(); // Column AS - unique ID
  // If no ID exists yet, generate one now before syncing
  if (!groupId || groupId.toString().trim() === '') {
    handleUniqueId(sheet, row);
    groupId = sheet.getRange(row, 45).getValue();
  }
  var groupName = sheet.getRange(row, 1).getValue(); // Column A
  var exhibitorAvailable = sheet.getRange(row, 22).getValue(); // Column V
  var exhibitorProAvailable = sheet.getRange(row, 23).getValue(); // Column W
  var vipPartyAvailable = sheet.getRange(row, 24).getValue(); // Column X

  // Column AS (45) holds either a 7-digit group ID (before first sync) or a JotForm
  // submission ID (written back after creation/first search). Submission IDs are ~19
  // digits, so anything longer than 10 characters is treated as a submission ID.
  var asValue = groupId.toString().trim();
  if (asValue.length > 10) {
    Logger.log('Column AS contains JotForm submission ID for row ' + row + ': ' + asValue);
    updateJotFormRecord(asValue, exhibitorAvailable, exhibitorProAvailable, vipPartyAvailable, JOTFORM_TABLE_ID, JOTFORM_API_KEY);
    Logger.log('Updated record ID: ' + asValue);
    return;
  }

  // AS still holds the group ID — search JotForm for an existing record
  var existingRecord = findJotFormRecordById(groupId, JOTFORM_TABLE_ID, JOTFORM_API_KEY);
  if (existingRecord && existingRecord.fetchError) {
    // API read failed; use short link as duplicate guard before attempting creation
    var existingShortLink = sheet.getRange(row, 42).getValue();
    if (existingShortLink && existingShortLink.toString().trim() !== '') {
      Logger.log('JotForm API error but short link exists for row ' + row + ' — skipping to avoid duplicate.');
      return;
    }
    Logger.log('JotForm API error for row ' + row + ' and no short link found — will attempt creation.');
    existingRecord = null;
  }

  if (existingRecord) {
    // Found in JotForm — update and overwrite AS with submission ID for future syncs
    updateJotFormRecord(existingRecord.id, exhibitorAvailable, exhibitorProAvailable, vipPartyAvailable, JOTFORM_TABLE_ID, JOTFORM_API_KEY);
    sheet.getRange(row, 45).setValue(existingRecord.id); // Overwrite group ID with submission ID
    Logger.log('Updated record ID: ' + existingRecord.id + ' (stored back in column AS)');
  } else {
    // Safety guard: if column AP already has a short link, the record was already created
    var existingShortLink = sheet.getRange(row, 42).getValue();
    if (existingShortLink && existingShortLink.toString().trim() !== '') {
      Logger.log('Row ' + row + ': JotForm record not found by search but short link exists in column AP — cannot update. Paste the JotForm submission ID into column AS to enable future updates.');
      return;
    }
    // Create new record (sets Used fields to 0 initially)
    var result = createJotFormRecord(groupId, groupName, EVENT_NAME, exhibitorAvailable, exhibitorProAvailable, vipPartyAvailable, JOTFORM_TABLE_ID, JOTFORM_API_KEY);
    if (result.responseCode === 200 && result.content) {
      var submissionId = result.content.toString();
      Logger.log('Created new record with ID: ' + submissionId);
      sheet.getRange(row, 45).setValue(submissionId); // Overwrite group ID with submission ID
      // Create short link
      var originalUrl = 'https://registration-router.resources-8c8.workers.dev/?grpid=' + groupId;
      var shortUrl = createShortLink(originalUrl, SHORTIO_API_KEY, SHORTIO_DOMAIN);
      if (shortUrl) {
        // Update column AP (column 42) with the short link
        sheet.getRange(row, 42).setValue(shortUrl);
        Logger.log('Added short link to column AP: ' + shortUrl);
      }
    }
  }
}

// --- HUBSPOT INTEGRATION ---
var HUBSPOT_ACCESS_TOKEN = PropertiesService.getScriptProperties().getProperty('HUBSPOT_ACCESS_TOKEN');
function handleHubSpotSync(sheet, row) {
  var statusCell = sheet.getRange(row, 44); // Column AR
  var checkboxCell = sheet.getRange(row, 43); // Column AQ
  statusCell.setValue("Syncing...");
  SpreadsheetApp.flush(); // Force update UI
  try {
    if (!HUBSPOT_ACCESS_TOKEN) {
      throw new Error('HubSpot Access Token is missing.');
    }
    // Get Deal ID
    var dealId = sheet.getRange(row, 41).getValue(); // Column AO
    if (!dealId) {
      throw new Error('No Deal ID in column AO.');
    }
    // Extract Contacts
    var contacts = [];
    // Primary Contacts: H (First), I (Last), J (Email)
    contacts = contacts.concat(extractContacts(
      sheet.getRange(row, 8).getValue(),
      sheet.getRange(row, 9).getValue(),
      sheet.getRange(row, 10).getValue()
    ));
    // Additional Contacts: L (First), M (Last), N (Email)
    contacts = contacts.concat(extractContacts(
      sheet.getRange(row, 12).getValue(),
      sheet.getRange(row, 13).getValue(),
      sheet.getRange(row, 14).getValue()
    ));
    if (contacts.length === 0) {
      throw new Error('No valid contacts found (missing emails?).');
    }
    Logger.log('Found ' + contacts.length + ' contacts to sync for Deal ' + dealId);
    // Sync to HubSpot
    contacts.forEach(function (contact) {
      if (!contact.email) return;
      var hubSpotContactId = createOrUpdateHubSpotContact(contact);
      if (hubSpotContactId) {
        associateContactToDeal(hubSpotContactId, dealId);
      }
    });
    statusCell.setValue("Success: " + new Date().toLocaleTimeString());
  } catch (e) {
    Logger.log("HubSpot Sync Error: " + e.toString());
    statusCell.setValue("Error: " + e.message);
  } finally {
    // Uncheck the box
    checkboxCell.setValue(false);
  }
}
function extractContacts(firstNamesRaw, lastNamesRaw, emailsRaw) {
  if (!emailsRaw || emailsRaw.toString().trim() === '') return [];
  var firstNames = (firstNamesRaw || '').toString().split(',').map(function (s) { return s.trim(); });
  var lastNames = (lastNamesRaw || '').toString().split(',').map(function (s) { return s.trim(); });
  var emails = (emailsRaw || '').toString().split(',').map(function (s) { return s.trim(); });
  var extracted = [];
  // Use email list as the source of truth for count
  for (var i = 0; i < emails.length; i++) {
    var email = emails[i];
    if (email) {
      extracted.push({
        firstname: firstNames[i] || '', // Default to empty string if missing
        lastname: lastNames[i] || '',   // Default to empty string if missing
        email: email
      });
    }
  }
  return extracted;
}
function createOrUpdateHubSpotContact(contact) {
  var searchResult = searchHubSpotContact(contact.email);
  var contactId;
  var properties = {
    'email': contact.email,
    'firstname': contact.firstname,
    'lastname': contact.lastname,
    'contact_type': 'Sponsor'
  };
  if (searchResult) {
    contactId = searchResult.id;
    Logger.log('Updating existing HubSpot contact: ' + contactId);
    updateHubSpotContact(contactId, properties);
  } else {
    Logger.log('Creating new HubSpot contact: ' + contact.email);
    contactId = createHubSpotContact(properties);
  }
  return contactId;
}
function searchHubSpotContact(email) {
  var url = 'https://api.hubapi.com/crm/v3/objects/contacts/search';
  var payload = {
    filterGroups: [{
      filters: [{
        propertyName: 'email',
        operator: 'EQ',
        value: email
      }]
    }]
  };
  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + HUBSPOT_ACCESS_TOKEN
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  var response = UrlFetchApp.fetch(url, options);
  var data = JSON.parse(response.getContentText());
  if (data.results && data.results.length > 0) {
    return data.results[0];
  }
  return null;
}
function createHubSpotContact(properties) {
  var url = 'https://api.hubapi.com/crm/v3/objects/contacts';
  var payload = {
    properties: properties
  };
  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + HUBSPOT_ACCESS_TOKEN
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  var response = UrlFetchApp.fetch(url, options);
  var data = JSON.parse(response.getContentText());
  if (response.getResponseCode() === 201) {
    return data.id;
  } else {
    Logger.log('Error creating contact: ' + response.getContentText());
    return null;
  }
}
function updateHubSpotContact(contactId, properties) {
  var url = 'https://api.hubapi.com/crm/v3/objects/contacts/' + contactId;
  var payload = {
    properties: properties
  };
  var options = {
    method: 'patch',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + HUBSPOT_ACCESS_TOKEN
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  UrlFetchApp.fetch(url, options);
}
function associateContactToDeal(contactId, dealId) {
  // Check if association already exists
  var checkUrl = 'https://api.hubapi.com/crm/v4/objects/contacts/' + contactId + '/associations/deals';
  var checkOptions = {
    method: 'get',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + HUBSPOT_ACCESS_TOKEN
    },
    muteHttpExceptions: true
  };
  var checkResponse = UrlFetchApp.fetch(checkUrl, checkOptions);
  var data = JSON.parse(checkResponse.getContentText());
  if (data.results) {
    for (var i = 0; i < data.results.length; i++) {
      if (data.results[i].toObjectId.toString() === dealId.toString()) {
        Logger.log('Contact ' + contactId + ' already associated with Deal ' + dealId + '. Skipping label update.');
        return; // Exit without changing anything
      }
    }
  }
  // Use Association ID 6 (Fulfillment Contact) as requested by the user.
  // Using 'USER_DEFINED' category as this is likely a custom label.
  var url = 'https://api.hubapi.com/crm/v4/objects/contacts/' + contactId + '/associations/deals/' + dealId;
  var payload = [
    {
      "associationCategory": "USER_DEFINED",
      "associationTypeId": 6
    }
  ];
  var options = {
    method: 'put',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + HUBSPOT_ACCESS_TOKEN
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  var assocResponse = UrlFetchApp.fetch(url, options);
  Logger.log('Association response: ' + assocResponse.getContentText());
}
// Returns the JotForm question ID (field number) for a given field label, e.g. "Google Sheets ID"
// Looks it up by scanning answers from existing submissions (avoids needing the /questions endpoint)
function getJotFormFieldIdByName(fieldName, tableId, apiKey) {
  var url = 'https://api.jotform.com/form/' + tableId + '/submissions?apiKey=' + apiKey + '&limit=10';
  try {
    var response = UrlFetchApp.fetch(url, { 'muteHttpExceptions': true });
    var data = JSON.parse(response.getContentText());
    if (data.responseCode === 200 && data.content && data.content.length > 0) {
      var answers = data.content[0].answers;
      for (var key in answers) {
        if (answers[key].name === fieldName) {
          Logger.log('Found field "' + fieldName + '" with ID: ' + key);
          return key;
        }
      }
    }
  } catch (e) {
    Logger.log('Error looking up field ID for "' + fieldName + '": ' + e);
  }
  Logger.log('Field "' + fieldName + '" not found in JotForm submissions');
  return null;
}
function findJotFormRecordById(groupId, tableId, apiKey) {
  if (!groupId || groupId.toString().trim() === '') {
    Logger.log('No Group ID provided, skipping JotForm search');
    return null;
  }
  Logger.log('Searching for Google Sheets ID: "' + groupId + '"');
  var offset = 0;
  var limit = 100;
  while (true) {
    var url = 'https://api.jotform.com/form/' + tableId + '/submissions?apiKey=' + apiKey + '&limit=' + limit + '&offset=' + offset;
    var response = UrlFetchApp.fetch(url, { 'muteHttpExceptions': true });
    var responseText = response.getContentText();
    if (responseText.trim().charAt(0) !== '{') {
      Logger.log('JotForm submissions returned non-JSON (HTTP ' + response.getResponseCode() + '): ' + responseText.substring(0, 300));
      return { fetchError: true };
    }
    var data = JSON.parse(responseText);
    if (data.responseCode !== 200 || !data.content || data.content.length === 0) break;
    for (var i = 0; i < data.content.length; i++) {
      var submission = data.content[i];
      var answers = submission.answers;
      if (answers && answers['28'] && answers['28'].answer !== undefined &&
          answers['28'].answer.toString() === groupId.toString()) {
        Logger.log('Found matching record ID: ' + submission.id);
        return { id: submission.id, data: submission };
      }
    }
    if (data.content.length < limit) break;
    offset += limit;
  }
  Logger.log('No matching record found for Google Sheets ID "' + groupId + '"');
  return null;
}
function updateJotFormRecord(submissionId, exhibitorAvailable, exhibitorProAvailable, vipPartyAvailable, tableId, apiKey) {
  var url = 'https://api.jotform.com/submission/' + submissionId + '?apiKey=' + apiKey;
  // IMPORTANT: Only update Available fields, NEVER touch Used fields
  var payload = {
    'submission[13]': exhibitorAvailable,  // Exhibitor Available
    'submission[17]': exhibitorProAvailable, // Exhibitor PRO Available
    'submission[23]': vipPartyAvailable     // VIP Party Available
  };
  Logger.log('Updating submission ' + submissionId + ' with: ' + JSON.stringify(payload));
  var options = {
    'method': 'post',
    'payload': payload,
    'muteHttpExceptions': true
  };
  var response = UrlFetchApp.fetch(url, options);
  var responseText = response.getContentText();
  var responseCode = response.getResponseCode();
  Logger.log('Update Response Code: ' + responseCode);
  Logger.log('Update Response: ' + responseText);
  if (responseText.trim().charAt(0) !== '{') {
    Logger.log('Update returned non-JSON (HTTP ' + responseCode + ') — likely an API key permissions issue');
    return { responseCode: responseCode, error: 'non-JSON response' };
  }
  return JSON.parse(responseText);
}
function createJotFormRecord(groupId, groupName, eventName, exhibitorAvailable, exhibitorProAvailable, vipPartyAvailable, tableId, apiKey) {
  var url = 'https://api.jotform.com/form/' + tableId + '/submissions?apiKey=' + apiKey;
  var googleSheetsIdFieldId = '28'; // QID for "Google Sheets ID" field (name: googleSheets)
  // IMPORTANT: Initialize Used fields to 0 when creating new records
  var payload = {
    'submission[21]': groupName,              // Group Name
    'submission[10]': eventName,              // event
    'submission[13]': exhibitorAvailable,     // Exhibitor Available
    'submission[14]': '0',                    // Exhibitor Used (initialize to 0)
    'submission[17]': exhibitorProAvailable,  // Exhibitor PRO Available
    'submission[18]': '0',                    // Exhibitor PRO Used (initialize to 0)
    'submission[23]': vipPartyAvailable,      // VIP Party Available
    'submission[24]': '0'                     // VIP Party Used (initialize to 0)
  };
  // Add Google Sheets ID to payload using the dynamically resolved field ID
  if (googleSheetsIdFieldId) {
    payload['submission[' + googleSheetsIdFieldId + ']'] = groupId; // Google Sheets ID
  } else {
    Logger.log('Warning: Could not find "Google Sheets ID" field in JotForm — ID will not be stored');
  }
  Logger.log('Creating submission with payload: ' + JSON.stringify(payload));
  var options = {
    'method': 'post',
    'payload': payload,
    'muteHttpExceptions': true
  };
  var response = UrlFetchApp.fetch(url, options);
  var responseText = response.getContentText();
  var responseCode = response.getResponseCode();
  Logger.log('Response Code: ' + responseCode);
  Logger.log('Response: ' + responseText);
  return JSON.parse(responseText);
}
function createShortLink(originalUrl, apiKey, domain) {
  var url = 'https://api.short.io/links';
  var payload = {
    'originalURL': originalUrl,
    'domain': domain
  };
  var options = {
    'method': 'post',
    'contentType': 'application/json',
    'headers': {
      'Authorization': apiKey
    },
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };
  Logger.log('Creating short link for: ' + originalUrl);
  try {
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();
    Logger.log('Short.io Response Code: ' + responseCode);
    Logger.log('Short.io Response: ' + responseText);
    if (responseCode === 200) {
      var data = JSON.parse(responseText);
      return data.shortURL;
    }
  } catch (e) {
    Logger.log('Error creating short link: ' + e);
  }
  return null;
}
