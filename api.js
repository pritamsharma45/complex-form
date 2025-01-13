const RESPONSE_SHEET_NAME = "ResponseSheet";
const CLIENT_EMAIL_FIELD = "client_email";
const ADMIN_EMAIL = "sharma.pritam311@gmail.com";
const PUBLISHED_URL =
  "https://script.google.com/macros/s/AKfycbyqC1d8xpSrZYogPeuFA186fzwGlpEnxdKCTkz_coNN7fEX-dV20MIjjxZzBpZHixgh/exec";

const EMAIL_SUBJECT_LINE = "Tax Return Form";

var ss = SpreadsheetApp.getActiveSpreadsheet();
// const ss = SpreadsheetApp.getActive();
const wsInvitations = ss.getSheetByName("Invitations");
const responseSheet = ss.getSheetByName(RESPONSE_SHEET_NAME);

function addEntry(formDataObject, formName, schema = null) {
  var now = JSON.stringify(new Date());
  var Agent;
  Tamotsu.initialize();
  if (formDataObject.ID) {
    formDataObject.ID = Number(formDataObject.ID);
  }
  try {
    Agent = Tamotsu.Table.define({ sheetName: formName, idColumn: "id" });
    formDataObject.created_date = now;
    var srfEntry = new Agent(formDataObject);
    const savedEntry = Agent.createOrUpdate(srfEntry);

    if (schema && savedEntry) {
      const model = JSON.parse(formDataObject.model);
      const clientName = model.s_0_fld_0 || "Client";
      const htmlBody = generateDetailViewHtml(schema, model);

      // Send email with the generated HTML
      const emailData = {
        subject: `${EMAIL_SUBJECT_LINE} - ${clientName} - ${new Date().toLocaleDateString()}`,
        clientName: clientName,
        htmlBody: htmlBody,
        cc: formDataObject.client_email,
      };

      emailPDF(emailData);
    }

    return savedEntry;
  } catch (error) {
    console.log(error);
    return false;
  }
}

function generateDetailViewHtml(formSchema, model) {
  let html =
    '<div style="font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto; padding: 20px;">';

  formSchema.forEach((section, sectionIndex) => {
    // Add section title if exists
    if (section.title) {
      html += `<h3 style="background: #f5f5f5; padding: 10px; margin: 20px 0 10px; border-bottom: 2px solid #ddd;">${section.title}</h3>`;
    }

    // Handle regular fields
    if (section.flds) {
      html += "<div>";
      section.flds.forEach((field, fieldIndex) => {
        const key = `s_${sectionIndex}_fld_${fieldIndex}`;
        const value = model[key] || "";

        if (value) {
          // Only show fields with values
          html += `
            <div style="margin: 8px 0; padding: 5px 0; border-bottom: 1px solid #eee;">
              <strong style="display: inline-block; min-width: 200px; color: #666;">${
                field.label
              }:</strong> 
              <span>${field.prefix ? field.prefix + " " : ""}${value}${
            field.suffix ? " " + field.suffix : ""
          }</span>
            </div>`;
        }
      });
      html += "</div>";
    }

    // Handle grouped fields
    if (section.groups) {
      section.groups.forEach((group, groupIndex) => {
        if (group.title) {
          html += `<h4 style="margin: 15px 0 10px; color: #444;">${group.title}</h4>`;
        }

        if (group.flds) {
          html += "<div>";
          group.flds.forEach((field, fieldIndex) => {
            const key = `s_${sectionIndex}_g_${groupIndex}_fld_${fieldIndex}`;
            const value = model[key] || "";

            if (value) {
              // Only show fields with values
              html += `
                <div style="margin: 8px 0; padding: 5px 0; border-bottom: 1px solid #eee;">
                  <strong style="display: inline-block; min-width: 200px; color: #666;">${
                    field.label
                  }:</strong> 
                  <span>${field.prefix ? field.prefix + " " : ""}${value}${
                field.suffix ? " " + field.suffix : ""
              }</span>
                </div>`;
            }
          });
          html += "</div>";
        }
      });
    }

    // Add section footer notes if exists
    if (section.footerNotes) {
      html += `<p style="font-style: italic; color: #666; margin-top: 10px;"><em>${section.footerNotes}</em></p>`;
    }
  });

  html += "</div>";
  return html;
}

function sendInvitations(emailIds) {
  let a = [];
  emailIds.forEach((email) => {
    let invitation = Math.random().toString(36).slice(5).toUpperCase();
    let status = "new";
    let rowContents = [new Date(), email, invitation, status];
    sendInvitationEmail(email, invitation);
    wsInvitations.appendRow(rowContents);
    a.push({ email, invitation, status });
  });

  return a;
}
function getInvitations() {
  let invitations = {};
  wsInvitations
    .getDataRange()
    .getValues()
    .forEach((value, i) => {
      if (i) {
        let key = value[1].toString().trim().toLowerCase();
        invitations[key] = {
          email: key,
          invitationCode: value[2].toString().trim(),
        };
      }
    });
  return invitations;
}

function getResposeSheet() {
  getResponseSheet(RESPONSE_SHEET_NAME);
}

function getFormData(enteredId) {
  Tamotsu.initialize();
  var tbl;
  tbl = Tamotsu.Table.define({
    sheetName: RESPONSE_SHEET_NAME,
    idColumn: CLIENT_EMAIL_FIELD,
  });
  try {
    Logger.log(tbl.find(enteredId));
    return tbl.find(enteredId);
  } catch (error) {
    Logger.log("Not Found");
    return null;
  }
}

function GetRecordById(id, tableName) {
  Tamotsu.initialize();
  var Agent;
  Agent = Tamotsu.Table.define({ sheetName: tableName, idColumn: "id" });
  return Agent.find(id);
}

function getStringifiedTables(tableNames) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var o = {};
  tableNames.forEach((name) => {
    var rng = ss.getRangeByName(name);
    var vals = trimRangeRows(rng).getValues();
    o[name] = vals;
  });

  return o;
}

function commitFilesToDB(sheetName, id, formDataObject) {
  var Agent;
  Tamotsu.initialize();
  Agent = Tamotsu.Table.define({ sheetName: sheetName, idColumn: "id" });
  var row = Agent.find(id);
  return row.updateAttributes(formDataObject);
}

function deleteEntryByID(id, tableName) {
  var Agent;
  Tamotsu.initialize();

  try {
    Agent = Tamotsu.Table.define({ sheetName: tableName, idColumn: "id" });
    Agent.find(id).destroy();
    return true;
  } catch (error) {
    console.log(error);
    return false;
  }
}

function UpdateAttributes(sheetName, id, updateObject) {
  var Agent;
  Tamotsu.initialize();
  Agent = Tamotsu.Table.define({ sheetName: sheetName, idColumn: "id" });
  var row = Agent.find(id);
  return row.updateAttributes(updateObject);
}

function GetTableDataSource(sheetName) {
  Tamotsu.initialize();
  var Agent;
  var tableDataSource = {};
  Agent = Tamotsu.Table.define({ sheetName: sheetName, idColumn: "id" });
  var dataTable = Agent.all();
  var fields = Object.keys(Agent.first());
  fields.splice(0, 1);

  cols = _.map(fields, function (el) {
    var obj = {};
    obj.field = el;
    obj.label = _.startCase(el);
    return obj;
  });
  tableDataSource.columns = cols;
  tableDataSource.items = dataTable.reverse();
  return tableDataSource;
}

function LogData() {
  Tamotsu.initialize();
  var Agent;
  Agent = Tamotsu.Table.define({ sheetName: "PRF_Active" });
  Logger.log(Agent.first());
}
function getLoggedInUser() {
  return Session.getActiveUser().getEmail();
}

function saveInDrive(f) {
  const blob = Utilities.newBlob(f, "image/jpeg", "some-name");
  const file = DriveApp.getFolderById("root").createFile(blob);
  return file.getUrl();
}
function saveFile(fileObjects) {
  fileObjects.forEach((o) => {
    var blob = Utilities.newBlob(
      Utilities.base64Decode(o.data),
      o.mimeType,
      o.fileName
    );
    DriveApp.createFile(blob).getId();
  });

  // for (const [k, v] of Object.entries(fileObjects)) {
  //   let arr = [];
  //   v.forEach((o) => {
  //     var blob = Utilities.newBlob(
  //       Utilities.base64Decode(o.data),
  //       o.mimeType,
  //       o.fileName
  //     );
  //     DriveApp.createFile(blob).getId();
  //     arr.push(DriveApp.createFile(blob).getId());
  //   });
  //   o[k] = arr;
  // }

  // var blob = Utilities.newBlob(
  //   Utilities.base64Decode(obj.data),
  //   obj.mimeType,
  //   obj.fileName
  // );
  // return o;
}

function saveFile2(obj) {
  // const FOLDER_ID = "138xQI8p87KZk0vDe2qz_ExBiVJe1Dzfp";
  const FOLDER_ID = "1-1NjFcvFik87CiXnkv-M3_sV3ncxAYeE"; //Gurneet ID
  var blob = Utilities.newBlob(
    Utilities.base64Decode(obj.data),
    obj.mimeType,
    obj.fileName
  );
  var folder = DriveApp.getFolderById(FOLDER_ID);
  return folder.createFile(blob).getUrl();
  // return DriveApp.createFile(blob).getUrl();
}
// function getPRFRecords(sheetName) {
//   Tamotsu.initialize();
//   try {
//     var tbl = Tamotsu.Table.define({
//       sheetName: sheetName,
//       idColumn: "id",
//     }).all();
//     return tbl;
//   } catch (error) {
//     return false;
//   }
// }

function getFormNumber(formName) {
  var Agent;
  Tamotsu.initialize();
  Agent = Tamotsu.Table.define({ sheetName: formName, idColumn: "id" });
  var formNumber = Agent.max("Form_Number");
  Logger.log(formNumber); //=> 300
  return formNumber;
}

function LogData() {
  Tamotsu.initialize();
  var Agent1;
  Agent1 = Tamotsu.Table.define({ sheetName: "PRF_Active", idColumn: "id" });
  Logger.log(Agent1.first());
}

function emailPDF(emailData) {
  try {
    // Get admin email from properties or use default
    const adminEmail = ADMIN_EMAIL;
    // Send email with HTML body
    MailApp.sendEmail({
      to: adminEmail,
      cc: emailData.cc,
      subject: emailData.subject,
      htmlBody: emailData.htmlBody,
    });

    return true;
  } catch (error) {
    console.error("Error sending email:", error);
    throw error;
  }
}
