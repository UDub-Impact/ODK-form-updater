/**
 * Create and add a menu to the UI
 */
function onOpen() {
  // Get the UI
  const ui = SpreadsheetApp.getUi();

  // Create a new menu and add items to it
  ui.createMenu("ODK Form Updater")
    .addItem("Update existing form draft", "createDraftForm")
    .addItem("Create new form", "createForm")
    .addItem("Configure", "configure")
    .addToUi();
}


/**
 * Displays a modal dialog box prompting the user to configure the add-on.
 */
function configure() {
  const ui = SpreadsheetApp.getUi();
  const widget = HtmlService.createHtmlOutputFromFile("ConfigurationForm.html");
  widget.setHeight(400);
  ui.showModalDialog(widget, "Configuration");
}


/**
 * Uploads a new draft form to ODK Central using the current configuration. 
 * If the configuration is faulty, user is notified through a toast message.
 * If the response status of the request is not 200, the user is notified through an alert message.
 */
function createDraftForm() {
  const formUrl = getFormUrl();
  const prep = prepareRequest();
  if (prep === null) {
    return;
  }

  // Confirm with user that they want to proceed with the form creation
  const ui = SpreadsheetApp.getUi();
  const confirmation = ui.alert(
    "Are you sure you want to continue?\n\n" +
    "Email: " + prep["email"] + "\n" +
    "Form Url: " + formUrl + "\n",
    ui.ButtonSet.YES_NO);

  // If user cancels the confirmation, stop form creation
  if (confirmation == ui.Button.NO) {
    return;
  }

  postRequest(formUrl + "/draft?ignoreWarnings=false",
    prep["token"], prep["formId"], prep["sheet"], "Create draft form");
}


/**
 * Uploads a new form to ODK Central using the current configuration. 
 * If the configuration is faulty, user is notified through a toast message.
 * If the response status of the request is not 200, the user is notified through an alert message.
 */
function createForm() {
  const projectUrl = getProjectUrl();
  const prep = prepareRequest();
  if (prep === null) {
    return;
  }

  // ask for confirmation to continue
  const ui = SpreadsheetApp.getUi();
  const confirmation = ui.alert(
    "Are you sure you want to continue?\n\n" +
    "Email: " + prep["email"] + "\n" +
    "Project Url: " + projectUrl + "\n" +
    "Form Id: " + prep["formId"] + "\n",
    ui.ButtonSet.YES_NO);

  // if confirmation is NO, exit
  if (confirmation == ui.Button.NO) {
    return;
  }

  postRequest(projectUrl + "/forms?ignoreWarnings=false&publish=false",
    prep["token"], prep["formId"], prep["sheet"], "Create new form");
}


/**
 * Retrieves user configuration properties and prepares a request object with email, token,
 * form ID, and sheet data. If any configuration properties are missing or invalid, it displays
 * an error message and returns null.
 *
 * @return {Object|null} The request object containing email, token, form ID, and sheet data or
 * null if there is an error.
 */
function prepareRequest() {
  // get user configuration properties
  const properties = PropertiesService.getUserProperties();
  const email = properties.getProperty("email");
  const password = properties.getProperty("password");
  const formId = properties.getProperty("formId");
  const sessionUrl = getSessionUrl();

  // check if any configuration properties are missing, if yes, toast and exit
  if (!email || !password || !formId || !sessionUrl) {
    SpreadsheetApp.getActiveSpreadsheet().toast("Get Config error: Reconfigure");
    return null;
  }

  // get authentication token using email, password and session url
  const token = getToken(email, password, sessionUrl);

  // if authentication failed, toast and exit
  if (!token) {
    SpreadsheetApp.getActiveSpreadsheet().toast("Authentication error: Reconfigure");
    return null;
  }

  // get sheet data
  const sheet = getSheet();

  // if sheet data is invalid, toast and exit
  if (!sheet) {
    SpreadsheetApp.getActiveSpreadsheet().toast("Sheet error: Invalid sheet");
    return null;
  }

  return {
    "email": email,
    "token": token,
    "formId": formId,
    "sheet": sheet
  };
}


/**
 * Sends a POST request to a specified URL with a token, form ID, and spreadsheet data.
 * If the response is not 200, it displays an error message with details if available.
 * Otherwise, it displays a success message.
 *
 * @param {string} url - The URL to send the request to.
 * @param {string} token - The token to include in the request header.
 * @param {string} formId - The form ID to include in the request header.
 * @param {object} sheet - The spreadsheet data to include in the request payload.
 * @param {string} successMessage - The message to display on success.
 */
function postRequest(url, token, formId, sheet, successMessage) {
  const response = UrlFetchApp.fetch(
    url, {
    "method": "post",
    "muteHttpExceptions": true,
    "contentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "headers": {
      "Authorization": "Bearer " + token,
      "X-XlsForm-FormId-Fallback": formId,
    },
    "payload": sheet,
  });

  if (response.getResponseCode() !== 200) {
    const error = JSON.parse(response.getContentText());
    let errorMessage = "Error Code: " + error["code"] + "\nMessage: " + error["message"];

    // Add details if it exists
    if ("details" in error) {
      errorMessage += "\nDetails:\n\n";
      for (let key in error["details"]) {
        errorMessage += key + ": " + JSON.stringify(error["details"][key]) + "\n";
      }
    }

    SpreadsheetApp.getUi().alert(errorMessage);
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast("Success: " + successMessage);
  }
}


/**
 * Retrieves the data from the active spreadsheet and returns it as a string.
 * Returns null if there is an error.
 *
 * @return {string|null} The data from the active spreadsheet or null if there is an error.
 */
function getSheet() {
  // Get the active spreadsheet and its ID.
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetId = spreadsheet.getId();

  // Retrieve the file and its export URL.
  const file = Drive.Files.get(spreadsheetId, { supportsAllDrives: true });
  const url = file.exportLinks[MimeType.MICROSOFT_EXCEL];

  // Get the OAuth token and send a GET request to the export URL.
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: {
      "Authorization": "Bearer " + token
    }
  });

  // Check if the GET request is successful and return the data from the active spreadsheet.
  if (response.getResponseCode() !== 200) {
    return null;
  }
  return response.getContent();
}

/**
 * Sets the configuration parameters in the UserProperties and displays a success message.
 * 
 * @param {string} email - The email address
 * @param {string} password - The password
 * @param {string} baseUrl - The base URL
 * @param {string} projectId - The ID of the project
 * @param {string} formId - The ID of the form
 */
function setConfig(email, password, baseUrl, projectId, formId) {
  PropertiesService.getUserProperties().setProperties({
    "email": email,
    "password": password,
    "baseUrl": baseUrl,
    "projectId": projectId,
    "formId": formId
  });
  SpreadsheetApp.getActiveSpreadsheet().toast("Success: Configuration");
}

/**
 * Returns the session URL
 * 
 * @return {string} The session URL.
 */
function getSessionUrl() {
  const properties = PropertiesService.getUserProperties();
  return properties.getProperty("baseUrl") + "/v1/sessions"
}

/**
 * Returns the project URL
 * 
 * @return {string} The project URL.
 */
function getProjectUrl() {
  const properties = PropertiesService.getUserProperties();
  return properties.getProperty("baseUrl") +
    "/v1/projects/" +
    properties.getProperty("projectId")
}

/**
 * Returns the form URL
 * 
 * @return {string} The form URL.
 */
function getFormUrl() {
  const properties = PropertiesService.getUserProperties();
  return getProjectUrl() +
    "/forms/" +
    properties.getProperty("formId");
}

/**
 * Returns the previous configuration parameters, except the password, as an array.
 * 
 * @return {Array} An array containing the email, project URL, and form ID.
 */
function getConfigWithNoPassword() {
  const properties = PropertiesService.getUserProperties();
  const email = properties.getProperty("email");
  const projectUrl = getProjectUrl();
  const formId = properties.getProperty("formId");
  return [email, projectUrl, formId];
}

/**
 * Makes a POST request to the session URL to retrieve an authentication token.
 * 
 * @param {string} email - The email address 
 * @param {string} password - The password 
 * @param {string} sessionUrl - The session URL 
 * @return {string|null} The authentication token or null if the POST request is unsuccessful.
 */
function getToken(email, password, sessionUrl) {
  // Define the request body and parameters.
  const bodyOfRequest = {
    "email": email,
    "password": password
  };
  const parametersOfRequest = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(bodyOfRequest),
    "muteHttpExceptions": true
  };

  // Send the POST request to the session URL.
  const response = UrlFetchApp.fetch(sessionUrl, parametersOfRequest);

  // Check if the POST request is successful and return the authentication token.
  if (response.getResponseCode() !== 200) {
    return null;
  }
  return JSON.parse(response.getContentText())["token"];
}
