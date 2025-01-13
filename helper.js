function ArrayToJSON(array2D) {
  const header = array2D[0];
  const body = array2D.slice(1);
  const arr2 = body.map((el) => {
    let obj = {};
    for (let i = 0; i < el.length; ++i) {
      obj[header[i]] = el[i];
    }
    return obj;
  });
  return arr2;
}
/**
 * Function to trim the rows of a range. The range should contain a header in the first row.
 * @param {Range} range: a range object from Google spreadsheet. First row of range must be the headers.
 * @returns {Range}
 */
function trimRangeRows(range) {
  var values = range.getValues();
  for (var rowIndex = values.length - 1; rowIndex >= 0; rowIndex--) {
    if (values[rowIndex].join("") !== "") {
      break;
    }
  }
  return range.offset(
    (rowOffset = 0),
    (columnOffset = 0),
    (numRows = rowIndex + 1)
  );
}
/**
 * Function to get JSON from named range.
 * @param {Array} optionSources: Array of range name
 * @returns {Object}
 */
function getDropdowns(optionSources) {
  let obj = {};
  try {
    optionSources.forEach((source) => {
      obj[source] = flattenModelInJSON(NamedRangeToJSON(source));
    });
    return obj;
  } catch (error) {}
}

/**
 * Function to get Indexed dropdowns  from named range.
 * @param {Array} optionSources: Array of range name
 * @returns {Object}
 */
function getIndexedDropdowns(optionSources) {
  let obj = {};
  try {
    optionSources.forEach((source) => {
      let o = {};
      const jsons = flattenModelInJSON(NamedRangeToJSON(source));
      jsons.map((row) => {
        o[row["id"]] = row["name"];
      });
      obj[source] = o;
    });
    return obj;
  } catch (error) {}
}

/**
 * Function to get JSON from named range.
 * @param {Array} JSONArray: Array of JSON object with model field
 * @returns {Array}
 */
function flattenModelInJSON(JSONArray) {
  try {
    let arr = JSONArray.map((el) => {
      // Logger.log(el)
      let o = JSON.parse(el.model);
      o.id = el.id;
      Logger.log(o);
      return o;
    });
    return arr;
  } catch (error) {}
}

/**
 * Function to get JSON from named range.
 * @param {Range} rangeName: Range Name
 * @returns {Object}
 */

function NamedRangeToJSON(rangeName) {
  var rng = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(rangeName);
  var vals = trimRangeRows(rng).getValues();
  return ArrayToJSON(vals);
}

function indexedTable(rangeName, index_col_name, value_col_name) {
  var rng = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(rangeName);
  var vals = trimRangeRows(rng).getValues();
  const jsons = ArrayToJSON(vals);
  let o = {};
  jsons.map((row) => {
    o[row[index_col_name]] = row[value_col_name];
  });
  return o;
}

function include(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
}

function getPageUrl(page) {
  let url = ScriptApp.getService().getUrl().replace("/dev", "/exec");
  if (page) {
    url += "?p=" + page;
  }
  return url;
}

function createFile({ data, type, name }, folder) {
  data = Utilities.base64Decode(data);
  const blob = Utilities.newBlob(data, type, name);
  const file = folder.createFile(blob);
  return file;
}

function getFolderByName(name) {
  const id = SpreadsheetApp.getActive().getId();
  const parentFolder = DriveApp.getFileById(id).getParents().next();
  const folders = parentFolder.getFoldersByName(name);
  let folder;
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = parentFolder.createFolder(name);
  }
  return folder;
}

// create token
function createToken(key) {
  let token = "key-" + Utilities.getUuid() + "-afei";
  let cache = CacheService.getScriptCache();
  let expirationInSeconds = 60 * 60 * 24 * 60;
  cache.put(key, token, expirationInSeconds);
  return token;
}

// save password
function savePassword(key, password) {
  let encodedPassword = encodePassword(password);
  let props = PropertiesService.getDocumentProperties();
  props.setProperty(key, encodedPassword);
}

// get password from document properties
function getPassword(key) {
  let props = PropertiesService.getDocumentProperties();
  let password = props.getProperty(key);
  if (password) {
    password = decodePassword(password);
  }
  return password;
}

// encode password
function encodePassword(password) {
  let encoded = Utilities.base64Encode(password);
  return encoded;
}

// decode password
function decodePassword(encoded) {
  let data = Utilities.base64Decode(encoded);
  let decoded = Utilities.newBlob(data).getDataAsString();
  return decoded;
}

//send invitation email
function sendInvitationEmail(email, code) {
  let subject = "Invitation To Submit Your Details";
  let body = "";
  let url = PUBLISHED_URL;
  let htmlBody = `<table width="100%" cellpadding="0" cellspacing="0"><tr>
      <td width="auto"><\/td>
      <td width="600px" style="padding: 24px; background: #eee;">
          <h1>Invitation<\/h1>
          <p>You were invited to submit your details, please <a href="${url}" style="color: #4DB6AC">
              click here<\/a> to open the form with invitation code written below.<\/p>
          
          <p style="margin: 13px 0px; 
              text-align:center; 
              background: #4DB6AC; 
              letter-spacing: 24px;  
              padding: 24px; 
              font-size: 70px;
              font-weight: bold;
              color: #fff;">${code}<\/p>
          
          <p>Thanks & BR,<br>Approval Workflow Team<\/p>
      <\/td>
      <td width="auto"><\/td><\/table>`;
  let options = {
    htmlBody,
  };
  GmailApp.sendEmail(email, subject, body, options);
}

function render(page, title, data) {
  const template = HtmlService.createTemplateFromFile(page);
  template.data = data;
  let html = template
    .evaluate()
    .setTitle(title)
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return html;
}
