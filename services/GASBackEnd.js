class GASBackEnd {
  constructor() {}
  sayHello() {
    return new Promise(function (resolve, reject) {
      google.script.run
        .withSuccessHandler(function (res) {
          resolve(res);
        })
        .withFailureHandler(function () {
          reject();
        })
        .sayHello();
    });
  }

  addEntry(formDataObject, formName) {
    return new Promise(function (resolve, reject) {
      google.script.run
        .withSuccessHandler(function (res) {
          resolve(res);
        })
        .withFailureHandler(function () {
          reject();
        })
        .addEntry(formDataObject, formName);
    });
  }

  createEntity(formDataObject, tableName) {
    return new Promise(function (resolve, reject) {
      google.script.run
        .withSuccessHandler(function (res) {
          resolve(res);
        })
        .withFailureHandler(function () {
          reject();
        })
        .createEntity(formDataObject, tableName);
    });
  }

  getRelatedTablesByModel(model) {
    return new Promise(function (resolve, reject) {
      google.script.run
        .withSuccessHandler(function (res) {
          resolve(res);
        })
        .withFailureHandler(function () {
          reject();
        })
        .getRelatedTablesByModel(model);
    });
  }
  GetPopulatedTable(table_name) {
    return new Promise(function (resolve, reject) {
      google.script.run
        .withSuccessHandler(function (res) {
          resolve(res);
        })
        .withFailureHandler(function () {
          reject();
        })
        .GetPopulatedTable(table_name);
    });
  }
  GetFilteredTasks(clientId) {
    return new Promise(function (resolve, reject) {
      google.script.run
        .withSuccessHandler(function (res) {
          resolve(res);
        })
        .withFailureHandler(function () {
          reject();
        })
        .GetFilteredTasks(clientId);
    });
  }
  GetTable(table_name) {
    return new Promise(function (resolve, reject) {
      google.script.run
        .withSuccessHandler(function (res) {
          resolve(res);
        })
        .withFailureHandler(function () {
          reject();
        })
        .GetTable(table_name);
    });
  }
}
