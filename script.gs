function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Update Right Report', functionName: 'scanGoogleDrive'},
  ];
  spreadsheet.addMenu('Action', menuItems);
}

function scanGoogleDrive() {
  var newSheetName = "Rapport Automatique"

  var files = DriveApp.getFiles();
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  createSheetIfNotExists(activeSpreadsheet, newSheetName);
  var sheet = activeSpreadsheet.getSheetByName(newSheetName);
  var file;
  var data = initData();
  var line = 1;
  while (files.hasNext()) {
    file = files.next();
    path = getPath(file);
    data = addOneFileData(file, data, line);
    line += 1;
    
    Logger.log(data)
  }

  Logger.log(data)
}

function addOneFileData(file, data, line) {
  //null = no access
  //1 = can view
  //2 = can edit
  //3 = owner

  data["path"][line] = getPath(file);
  data["name"][line] = file.getName();
  data["warning"][line] = false;

  if (file.getSharingAccess !== DriveApp.Access.PRIVATE) {
      data["warning"][line] = true;
  }

  viewers = file.getViewers();
  editors = file.getEditors();
  owner = file.getOwner();

  data = addOneFilePermissions(data, viewers, line, 1);
  data = addOneFilePermissions(data, editors, line, 2);
  data = addOneFilePermissions(data, owner, line, 3);

  return data;
}

function addOneFilePermissions(data, users, line, permission) {
  if (Array.isArray(users)) {
    for (var i = 0; i < users.length; i++) {
      var user = users[i];
      data = addOneUserPermission(user, data, permission);
    }
  } else {
    var user = users;
    data = addOneUserPermission(user, data, permission, line);
  }
  return data;
}

function addOneUserPermission(user, data, permission, line) {
  user = user.getName() + ";" + user.getEmail();
  if (data[user] == null) {
    data[user] = []
  }
  data[user][line] = permission;

  return data;
}

function initData() {
  data = {};
  data["path"] = [null];
  data["name"] = [null];
  data["warning"] = [false];
  return data;
}

function getPath(file) {
  var folders = [];
  var parent = file.getParents();

  while (parent.hasNext()) {
    parent = parent.next();
    folders.push(parent.getName());
    parent = parent.getParents();
  }
  if (folders.length) {
    return folders.reverse().join("/");
  }
  return null;
}

function createSheetIfNotExists(activeSpreadsheet, newSheetName) {
  var NewSheet = activeSpreadsheet.getSheetByName(newSheetName);
  if (NewSheet != null) {
    activeSpreadsheet.deleteSheet(NewSheet);
  }
  NewSheet = activeSpreadsheet.insertSheet();
  NewSheet.setName(newSheetName);
}
