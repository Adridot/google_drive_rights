function onOpen() {
  const spreadsheet = SpreadsheetApp.getActive();
  const menuItems = [
    {name: 'Update Right Report', functionName: 'scanGoogleDrive'},
  ];
  spreadsheet.addMenu('Action', menuItems);
}

//main function
function scanGoogleDrive() {
  const newSheetName = "Report";

  //classic getting the file, removing the old sheet and creating a new one
  const files = DriveApp.getFiles();
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  createSheetIfNotExists(activeSpreadsheet, newSheetName);
  const sheet = activeSpreadsheet.getSheetByName(newSheetName);

  let data = initData(); //initiating the datatable

  let file;
  let line = 1; //variable to track our position in the code

  //adding information to the datatable for each file
  while (files.hasNext()) {
    file = files.next();
    data = addOneFileData(file, data, line);
    line += 1;
  }
  //last line to make all lines the same length
  data = endFile(data, line);

  //creating a 2D array from a JS object (dictionary)
  data = dataToBigArray(data);

  //getting the name and the email address of the user on two separate lines
  data = splitName_Email(data);

  //getting the range of the datatable to import it into the sheet
  const range = ["A1:", columnToLetter(data[0].length), data.length].join("");
  sheet.getRange(range).setValues(data);

  //Adding display features for the colors and size of the columns
  displayRules(sheet, range, data)

}

function createSheetIfNotExists(activeSpreadsheet, newSheetName) {
  let NewSheet = activeSpreadsheet.getSheetByName(newSheetName);
  if (NewSheet != null) {
    activeSpreadsheet.deleteSheet(NewSheet);
  }
  NewSheet = activeSpreadsheet.insertSheet();
  NewSheet.setName(newSheetName);
}

//Creating a JS object, and writing the first line
function initData() {
  let data = {};
  data["path"] = [null];
  data["name"] = [null];
  data["warning"] = [null];
  return data;
}

function addOneFileData(file, data, line) {
  data["path"][line] = getPath(file);
  data["name"][line] = file.getName();
  data["warning"][line] = false;

  //if the fileSharingAccess is not PRIVATE, this means the file can be accessed via the link.
  if (file.getSharingAccess() !== DriveApp.Access.PRIVATE) {
    data["warning"][line] = true;
  }

  //adding the permissions for each type of user.
  let viewers = file.getViewers();
  let editors = file.getEditors();
  let owner = file.getOwner();

  data = addOneFilePermissions(data, viewers, line, "can read");
  data = addOneFilePermissions(data, editors, line, "can edit");
  data = addOneFilePermissions(data, owner, line, "owner");

  return data;
}

function getPath(file) {
  const folders = [];
  let parent = file.getParents();

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

function addOneFilePermissions(data, users, line, permission) {
  let user;
  //checking if the user is an array or not, in case of an owner, which is never in array.
  if (Array.isArray(users)) {
    for (let i = 0; i < users.length; i++) {
      user = users[i];
      data = addOneUserPermission(user, data, permission, line);
    }
  } else {
  user = users;
  data = addOneUserPermission(user, data, permission, line);
}
  return data;
}

function addOneUserPermission(user, data, permission, line) {
  //we have to concatenate the name and the email, we'll separate them later. 
  user = user.getName() + ";" + user.getEmail();
  //creating a new entry in the object if this user doesn't exist yet
  if (data[user] == null) {
    data[user] = []
  }
  data[user][line] = permission;

  return data;
}

function endFile(data, line) {
  for (let i in data) {
    data[i][line] = null;
  }
  return data
}

function dataToBigArray(data) {
  let bigArray = [[]];

  //firstly the table header
  for (let i in data) {
    bigArray[0].push(i)
  }

  //then the rest of the data
  for (let j = 0; j < data["name"].length; j++) {
    bigArray.push([])
    for (let i in data) {
        bigArray[j + 1].push(data[i][j])
    }
  }
  return bigArray;
}

function splitName_Email(data) {
  let user;
  for (let i = 3; i < data[0].length; i++) {
    user = data[0][i].split(";")
    data[0][i] = user[0];
    data[1][i] = user[1];
  }
  return data
}

//column number to column letters
function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function displayRules(sheet, range, data) {
  //merging the 3 first columns and aligning the texts in the center
  sheet.getRange('A1:A2').activate().mergeVertically();
  sheet.getRange('B1:B2').activate().mergeVertically();
  sheet.getRange('C1:C2').activate().mergeVertically();
  sheet.getRange('A1:C2').setHorizontalAlignment('center').setVerticalAlignment('middle');

  //sizing the columns 
  sheet.autoResizeColumns(1, 3);
  sheet.setColumnWidths(4, data[0].length - 3, 90)

  //setting the colors depending on the user role.
  conditionalRule(sheet, range, "#4285f4", "can read");
  conditionalRule(sheet, range, "#34a853", "can edit");
  conditionalRule(sheet, range, "#ff6d01", "owner");
  conditionalRule(sheet, range, "#ea4335", "true");
}

function conditionalRule(sheet, range, color, text) {
  range = sheet.getRange(range);
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(text)
    .setBackground(color)
    .setRanges([range])
    .build();
  let rules = sheet.getConditionalFormatRules();
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);
}
