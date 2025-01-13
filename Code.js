
function doGet(e) {
  const pages = {
    index: "Client Form",
    404: "Page not found",
    adminview: "Form Admin",
    about: "About",
  };
  let data = {};
  let page = "index";
  if (e.queryString) {
    page = e.parameter.p;
  }
  Logger.log(page);
  return render(page, pages[page], data);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// get token
function getToken(key) {
  let cache = CacheService.getScriptCache();
  let token = cache.get(key);
  return token;
}
// log in
function login(email, passedCode) {
  let token = "invalid";
  let code = getInvitations()[email].invitationCode;
  if (passedCode === code) {
    token = createToken(email);
  } else {
    token = "mismatch";
  }
  return { token };
}
