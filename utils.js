var MILLIS_PER_DAY = 1000 * 60 * 60 * 24;
function getDateBefore(days) {
  return new Date(now.getTime() - days * MILLIS_PER_DAY);
}

function dateDiff(dt_f, dt_i) {
  var diff = new Date(dt_f).getTime() - new Date(dt_i).getTime();
  return diff / MILLIS_PER_DAY;
}
function humanize(str) {
  var i,
    frags = str.split("_");
  for (i = 0; i < frags.length; i++) {
    frags[i] = frags[i].charAt(0).toUpperCase() + frags[i].slice(1);
  }
  return frags.join(" ");
}

// function ArrayToJSON(array2D) {
//   const header = array2D[0];
//   const body = array2D.slice(1);
//   const arr2 = body.map((el) => {
//     let obj = {};
//     for (let i = 0; i < el.length; ++i) {
//       obj[header[i]] = el[i];
//     }
//     return obj;
//   });
//   return arr2;
// }

function JSONToArray(jsonArray) {
  var header = [];
  for (const [k, v] of Object.entries(jsonArray[0])) {
    header.push(k);
  }
  return jsonArray.map((r) => {
    const a = [];
    for (let i = 0; i < header.length; ++i) {
      a.push(r[header[i]]);
    }
    return a;
  });
}

function CreateFile() {
  var name = "your-spreadsheet-name";
  var folderId = "your-folder-id";

  var resource = {
    title: name,
    mimeType: MimeType.GOOGLE_SHEETS,
    parents: [{ id: folderId }],
  };

  var fileJson = Drive.Files.insert(resource);
  var fileId = fileJson.id;
}
function createSpreadSheetInFolder(name, folder) {
  var ss = SpreadsheetApp.create(name);
  var id = ss.getId();
  var file = DriveApp.getFileById(id);
  folder.addFile(file);
  return ss;
}

function getPageUrl(page) {
  let url = ScriptApp.getService().getUrl().replace("/dev", "/exec");
  if (page) {
    url += "?p=" + page;
  }
  // Logger.log(url);
  return url;
}

function render(page, title, data) {
  const template = HtmlService.createTemplateFromFile(page);
  // template.data = data;
  let html = template
    .evaluate()
    .setTitle(title)
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return html;
}
// folderId='your_folder_id'
// name='my_new_ss'
// folder=DriveApp.getFolderById(folderId)
// createSpreadSheetInFolder(name,folder)
