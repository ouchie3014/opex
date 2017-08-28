//Global Settings/Variables:


//Sheet Names:
var sheetNameMainDefault = "MainDB";
var sheetNameInventoryDefault = "InventoryDB";
var sheetNamePMDefault = "PM Summary";
var sheetNameSnapshotDefault = "Snapshot";

//Email Settings:
var enableAutoEmailingDefault = true;
var warehouseDefault = 'zMHC600002';
var regionDefault = '09';
var gmailSnapshotLabelDefault = 'Snapshot';
var latestSnapshotSheetDefault = 'Latest Snapshot';
var emailSnapshotRequestHrDefault = 8; //(24hr) Time of day to request snapshot. Default is 8am.
var updateSnapshotHrDefault = 9; //(24hr) Time of day to update snapshot. Recommended to be 1 hour after emailSnapshotRequestHr.

var excelExtensionDefault = '.xls';

//Location ID's:
var mainSSidDefault = '10Mz2u6E9pdrKisA4MlH1PXbVaW2qTy3DoP7boiP_wFk';
var snapshotSSidDefault = '1FcfagPIlAyKJ4IDN9Hd076rXx-_oZGmmEMv06HF4IuE';
var snapshotFolderIdDefault = '0BwtCesct9-1EQk1jMnZsM2h1aXc';


// ----------- DO NOT EDIT BELOW THIS LINE ----------- //





var scriptProperties = PropertiesService.getScriptProperties();

//Location ID's:
var mainSSid = scriptProperties.getProperty('mainSSid') || mainSSidDefault;
var snapshotSSid = scriptProperties.getProperty('snapshotSSid') || snapshotSSidDefault;
var snapshotFolderId = scriptProperties.getProperty('snapshotFolderId') || snapshotFolderIdDefault;


//Sheet Names:
var sheetNameMain = scriptProperties.getProperty('sheetNameMain') || sheetNameMainDefault;
var sheetNameInventory = scriptProperties.getProperty('sheetNameInventory') || sheetNameInventoryDefault;
var sheetNamePM = scriptProperties.getProperty('sheetNamePM') || sheetNamePMDefault;
var sheetNameSnapshot = scriptProperties.getProperty('sheetNameSnapshot') || sheetNameSnapshotDefault;



//Email Settings:
var enableAutoEmailing = scriptProperties.getProperty('enableAutoEmailing') || enableAutoEmailingDefault;
var gmailSnapshotLabel = scriptProperties.getProperty('gmailSnapshotLabel') || gmailSnapshotLabelDefault;
var latestSnapshotSheet = scriptProperties.getProperty('latestSnapshotSheet') || latestSnapshotSheetDefault;
var warehouse = scriptProperties.getProperty('warehouse') || warehouseDefault;
var region = scriptProperties.getProperty('region') || regionDefault;
var excelExtension = '.xls';
var emailSnapshotRequestHr = scriptProperties.getProperty('emailSnapshotRequestHr') || emailSnapshotRequestHrDefault; //(24hr) Time of day to request snapshot. Default is 8am.
var updateSnapshotHr = scriptProperties.getProperty('updateSnapshotHr') || updateSnapshotHrDefault; //(24hr) Time of day to update snapshot. Recommended to be 1 hour after emailSnapshotRequestHr.


//Global variables:
var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetNameMain);
var inventorySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetNameInventory);
var pmDataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetNamePM);
var snapshotSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetNameSnapshot);




function saveSidebarSettings(form) {
  var scriptProperties = PropertiesService.getScriptProperties();

  scriptProperties.setProperty('sheetNameMain', form.sheetNameMain);
  scriptProperties.setProperty('sheetNameInventory', form.sheetNameInventory);
  scriptProperties.setProperty('sheetNamePM', form.sheetNamePM);
  scriptProperties.setProperty('sheetNameSnapshot', form.sheetNameSnapshot);

  scriptProperties.setProperty('enableAutoEmail', form.enableAutoEmail);
  scriptProperties.setProperty('warehouse', form.warehouse);
  scriptProperties.setProperty('region', form.region);
  scriptProperties.setProperty('gmailSnapshotLabel', form.gmailSnapshotLabel);
  scriptProperties.setProperty('latestSnapshotSheet', form.latestSnapshotSheet);
  
  scriptProperties.setProperty('emailSnapshotRequestHr', form.emailSnapshotRequestHr);
  scriptProperties.setProperty('updateSnapshotHr', form.updateSnapshotHr);
  
  scriptProperties.setProperty('mainSSid', form.mainSSid);
  scriptProperties.setProperty('snapshotSSid', form.snapshotSSid);
  scriptProperties.setProperty('snapshotFolderId', form.snapshotFolderId);
  
  console.log("Settings saved.");
  
}


function getPreferences() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var sheetSettings = {
    sheetNameMain: scriptProperties.getProperty('sheetNameMain'),
    sheetNameInventory: scriptProperties.getProperty('sheetNameInventory'),
    sheetNamePM: scriptProperties.getProperty('sheetNamePM'),
    sheetNameSnapshot: scriptProperties.getProperty('sheetNameSnapshot'),
    
    enableAutoEmailing: scriptProperties.getProperty('enableAutoEmailing'),
    warehouse: scriptProperties.getProperty('warehouse'),
    region: scriptProperties.getProperty('region'),
    gmailSnapshotLabel: scriptProperties.getProperty('gmailSnapshotLabel'),
    latestSnapshotSheet: scriptProperties.getProperty('latestSnapshotSheet'),
    
    emailSnapshotRequestHr: scriptProperties.getProperty('emailSnapshotRequestHr'),
    updateSnapshotHr: scriptProperties.getProperty('updateSnapshotHr'),
    
    mainSSid: scriptProperties.getProperty('mainSSid'),
    snapshotSSid: scriptProperties.getProperty('snapshotSSid'),
    snapshotFolderId: scriptProperties.getProperty('snapshotFolderId')
  };
  console.log("Settings loaded.");
  return sheetSettings;
}

function restoreDefaultSettings() {
  var scriptProperties = PropertiesService.getScriptProperties();

  scriptProperties.setProperty('sheetNameMain', sheetNameMainDefault);
  scriptProperties.setProperty('sheetNameInventory', sheetNameInventoryDefault);
  scriptProperties.setProperty('sheetNamePM', sheetNamePMDefault);
  scriptProperties.setProperty('sheetNameSnapshot', sheetNameSnapshotDefault);

  scriptProperties.setProperty('enableAutoEmailing', enableAutoEmailingDefault);
  scriptProperties.setProperty('warehouse', warehouseDefault);
  scriptProperties.setProperty('region', regionDefault);
  scriptProperties.setProperty('gmailSnapshotLabel', gmailSnapshotLabelDefault);
  scriptProperties.setProperty('latestSnapshotSheet', latestSnapshotSheetDefault);
  
  scriptProperties.setProperty('emailSnapshotRequestHr', emailSnapshotRequestHrDefault);
  scriptProperties.setProperty('updateSnapshotHr', updateSnapshotHrDefault);
  
  scriptProperties.setProperty('mainSSid', mainSSidDefault);
  scriptProperties.setProperty('snapshotSSid', snapshotSSidDefault);
  scriptProperties.setProperty('snapshotFolderId', snapshotFolderIdDefault);
  
  console.log("Settings restored to default values.");
  
}

function deleteAllSettings() {
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteAllProperties();
  console.log("All stored settings deleted!.");
}