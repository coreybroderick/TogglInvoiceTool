var TOGGL_API_KEY = ''; // Comes from Toggl
var WORKSPACE_ID = '';  // Comes from Toggl
var USER_AGENT = ''; // your email address you use for Toggl

function onOpen(){
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Toggl')
      .addItem('Get Data', 'getData')
      .addItem('Create Invoices', 'sortAndPrint')
      .addItem('Make PDFs', 'pdfTheSheets')
      .addToUi();
}

function weeklyTimeReport() {
  // lets assume that today is monday, and we need Sunday to Sunday of last week
  var today = new Date();
  
  var rightSunday = new Date();
  rightSunday.setDate(today.getDate()-1);
  
  var leftSunday = new Date();
  leftSunday.setDate(today.getDate()-8);

  var startDate = [leftSunday.getFullYear(), leftSunday.getMonth(), leftSunday.getDate()];
  var endDate = [rightSunday.getFullYear(), rightSunday.getMonth(), rightSunday.getDate()];

  var d1 = new Date(startDate[0],startDate[1],startDate[2], 0, 0, 0, 0);
  var d2 = new Date(endDate[0],endDate[1],endDate[2], 0, 0, 0, 0);

  var dataUrl = getDataURL(WORKSPACE_ID,d1,d2);
  var fetchOptions = getUrlFetchGETOptions(TOGGL_API_KEY);
  
  var dataResponse = UrlFetchApp.fetch(dataUrl,fetchOptions).getContentText();  
  var dataObjects = JSON.parse(dataResponse);

  // push JSON parsed object into growing 2dArray
  var array2dResponse =  [];
  dataObjects.data.forEach(function(dataRow){
    dataRow.items.forEach(function(field){
      var rowArray=[];
      rowArray.push(dataRow.title.project);
      rowArray.push(dataRow.title.client);
      var timeAmount = field.time/1000/60/60;
      rowArray.push(Math.round10(timeAmount,-2)); //  put into hours deciaml
      rowArray.push(field.title.time_entry);
      array2dResponse.push(rowArray);
    });
  });

  // need to sort the 2d-array by client
  array2dResponse.sort(sortFunctionByClient);

  // build array of items by client name
  var workingClientObject = {clientName:'',lineitem:[]};
  var clientObjects = [];
  for(var i=0;i<array2dResponse.length;i++){

    if(workingClientObject.clientName!=array2dResponse[i][1]){
    
      if(workingClientObject.clientName!=""){
        clientObjects.push(workingClientObject);
        workingClientObject = {clientName:'',lineitem:[]};
      }
    
      workingClientObject.clientName = array2dResponse[i][1];
    }
    workingClientObject.lineitem.push({time:array2dResponse[i][2],description:array2dResponse[i][3],project:array2dResponse[i][0]});
  }

  //push that last item in
  clientObjects.push(workingClientObject);

  for(var i=0;i<clientObjects.length;i++) {
    var clientInfo = getClientInfoByClient(clientObjects[i].clientName); 
    var theBody = 'Here is the automated report for all activities completed last week.  As always, if you have any questions, please let me know. <br/><br/>Thanks,<br/>Consultant Name<br/><br/>';
    theBody += '<table style="border:1px solid #cccccc;padding:.2rem;border-collapse: collapse;"><tr style="border:1px solid #cccccc;padding:.2rem;"><th style="border:1px solid #cccccc;padding:.2rem;">Project</th><th style="border:1px solid #cccccc;padding:.2rem;">Time</th><th style="border:1px solid #cccccc;padding:.2rem;">Description</th></tr>';
    for(var k=0;k<clientObjects[i].lineitem.length;k++){
      theBody += '<tr style="border:1px solid #cccccc;padding:.2rem;"><td style="border:1px solid #cccccc;padding:.2rem;">'+clientObjects[i].lineitem[k].project+'</td>' + '<td style="border:1px solid #cccccc;padding:.2rem;">'+clientObjects[i].lineitem[k].time+'</td>' + '<td style="border:1px solid #cccccc;padding:.2rem;">'+clientObjects[i].lineitem[k].description+'</td></tr>';
    }
    theBody += '</table>';

    var anEmail = GmailApp.createDraft(clientInfo.emails,  'Here is the automated report for all activities completed last week.',{
          htmlBody: theBody
      });
  }
}

function pdfTheSheets(){
  var sa = SpreadsheetApp.getActive();
  var ss = sa.getSheets();

  for(var i=0;i<ss.length;i++){
    if(ss[i].getName()=='Report Data' || ss[i].getName()=='Invoice Template' || ss[i].getName()=='Client Information'){
      // don't print these
      continue;
    }
    else{
      // the client name is in A7, we need the destination folderId for this client....
      var clientName = ss[i].getRange('A7').getValue();
      //var clientFolderId = getFolderIdByClientName(clientName);
      var clientInfo = getClientInfoByFullName(clientName);
      
      if(clientInfo == null){
        Browser.msgBox('BOOM.. Client name did not get a folder Id. Who the fuck is :'+clientName);
        continue;
      }
      
      // prep a URL to download in a PDF format
      var thisUrl = sa.getUrl().replace(/edit$/,'');
      var url_ext = 'export?exportFormat=pdf&format=pdf'   //export as pdf
        + '&gid=' + ss[i].getSheetId()   //the sheet's Id
        + '&size=letter'      // paper size
        + '&portrait=true'    // orientation, false for landscape
        + '&fitw=true'        // fit to width, false for actual size
        + '&sheetnames=false&printtitle=false&pagenumbers=false'  //hide optional headers and footers
        + '&gridlines=false'  // hide gridlines
        + '&fzr=false';       // do not repeat row headers (frozen rows) on each page
      var options = {
        headers: {
          'Authorization': 'Bearer ' +  ScriptApp.getOAuthToken()
        }
      }
      // actually do the download
      var response = UrlFetchApp.fetch(thisUrl + url_ext, options);
      // the invoicenumber is in cell F3
      var blob = response.getBlob().setName(ss[i].getRange('F3').getValue() + '.pdf');
      
      var fldr = DriveApp.getFolderById(clientInfo.clientFolderId);
      var thePdf = fldr.createFile(blob);
      
      // and here's the cool part
      var anEmail = GmailApp.createDraft(clientInfo.emails, 'Invoice ' + ss[i].getRange('F3').getValue(),"Hi,  I\'ve attached the invoice for last month\'s work.  As always, if you have any questions, please let me know.  Thanks, Name.",{
          attachments: [thePdf.getAs(MimeType.PDF)],
          htmlBody: '<hmtl><body>Hi,<br/><br/> I\'ve attached the invoice for last month\'s work. As always, if you have any questions, please let me know.<br/><br/>Thanks,<br/>Name</body></html>'
      });
    }
  }
}

function getClientInfoByClient(clientName){
  var resultObj = {};
  var sa = SpreadsheetApp.getActive();
  var clientInfoSheet = sa.getSheetByName("Client Information");
  var clientInfoData = clientInfoSheet.getRange(1,1,clientInfoSheet.getLastRow(),clientInfoSheet.getLastColumn()).getValues();
  for(var c=0;c<clientInfoData.length;c++){
    if(clientInfoData[c][0]==clientName){
      var resultObj = {};
      resultObj.clientFolderId = clientInfoData[c][5];
      resultObj.emails = clientInfoData[c][7];
      return resultObj;
    }
  }
  return null;
}

function getClientInfoByFullName(clientName){
  var resultObj = {};
  var sa = SpreadsheetApp.getActive();
  var clientInfoSheet = sa.getSheetByName("Client Information");
  var clientInfoData = clientInfoSheet.getRange(1,1,clientInfoSheet.getLastRow(),clientInfoSheet.getLastColumn()).getValues();
  for(var c=0;c<clientInfoData.length;c++){
    if(clientInfoData[c][1]==clientName){
      var resultObj = {};
      resultObj.clientFolderId = clientInfoData[c][5];
      resultObj.emails = clientInfoData[c][7];
      return resultObj;
    }
  }
  return null;
}

function sortAndPrint(){
  var invoiceMainNumber = Browser.inputBox("Invoice main number (20180101,beginning of month) ");
                                           
  var sa = SpreadsheetApp.getActive();
  var ss = sa.getSheetByName("Report Data");
  var data = ss.getRange(1, 1, ss.getLastRow(),ss.getLastColumn()).getValues();
  
  // sort by client
  var workingClientObject = {clientName:'',lineitem:[]};
  var clientObjects = [];
  for(var i=0;i<data.length;i++){
    if(workingClientObject.clientName!=data[i][1]){
      if(workingClientObject.clientName!=""){
        clientObjects.push(workingClientObject);
        workingClientObject = {clientName:'',lineitem:[]};
      }
      workingClientObject.clientName = data[i][1];
    }
    workingClientObject.lineitem.push({time:data[i][2],description:data[i][3],project:data[i][0]});
  }
  //push that last item in
   clientObjects.push(workingClientObject);
  
  var clientInfoSheet = sa.getSheetByName("Client Information");
  var clientInfoData = clientInfoSheet.getRange(1,1,clientInfoSheet.getLastRow(),clientInfoSheet.getLastColumn()).getValues();
  
  // should all be in the array of objects now.
  for(var i=0;i<clientObjects.length;i++){
    var newSheet = sa.getSheetByName("Invoice Template").copyTo(sa);
    
    var destinationFolder;
    newSheet.setName(clientObjects[i].clientName);
    for(var c=0;c<clientInfoData.length;c++){
      if(clientInfoData[c][0]==clientObjects[i].clientName){
        newSheet.getRange("A7").setValue(clientInfoData[c][1]);
        newSheet.getRange("A8").setValue(clientInfoData[c][2]);
        newSheet.getRange("A9").setValue(clientInfoData[c][3]);
        newSheet.getRange("F8").setValue(clientInfoData[c][4]);
        newSheet.getRange("E10").setValue('(payment terms:'+clientInfoData[c][6]+')');
        destinationFolder = clientInfoData[c][5];
      }
    }
    var detailArray = [];
    var hourSum = 0;
    for(var c=0;c<clientObjects[i].lineitem.length;c++){
      hourSum = hourSum + clientObjects[i].lineitem[c].time;
      var newRow = [];
      newRow.push(clientObjects[i].lineitem[c].project+'-'+clientObjects[i].lineitem[c].description);
      newRow.push('');
      newRow.push('');
      newRow.push('');
      newRow.push('');
      newRow.push(clientObjects[i].lineitem[c].time);
      detailArray.push(newRow);
    }
    newSheet.getRange(13,1,detailArray.length,6).setValues(detailArray);
    newSheet.getRange(13,1,detailArray.length,6).setBorder(true,true,true,true,true,true);
    newSheet.getRange("F7").setValue(hourSum);
    newSheet.getRange("F9").setFormula("=F7*F8");
    newSheet.getRange("F3").setValue(invoiceMainNumber + "-" + (i+1));
    newSheet.getRange("F2").setValue(formatDate(Date.now(),"EST","dd MMMM, yyyy"));

  }
  
  SpreadsheetApp.flush(); //before print
}

function getData(){

  var startDateS = Browser.inputBox('What is the start date? (2018-06-01)');
  var startDate = startDateS.split("-");
  
  var endDateS = Browser.inputBox('What is the end date? (2018-06-31)');
  var endDate = endDateS.split("-");
  
  var d1 = new Date(startDate[0],startDate[1]-1,startDate[2], 0, 0, 0, 0);
  var d2 = new Date(endDate[0],endDate[1]-1,endDate[2], 0, 0, 0, 0);
  
  var dataUrl = getDataURL(WORKSPACE_ID,d1,d2);
  var fetchOptions = getUrlFetchGETOptions(TOGGL_API_KEY);
  
  var dataResponse = UrlFetchApp.fetch(dataUrl,fetchOptions).getContentText();  
  var dataObjects = JSON.parse(dataResponse);
  
  // push JSON parsed object into growing 2dArray
  var array2dResponse =  [];
  dataObjects.data.forEach(function(dataRow){
    
    dataRow.items.forEach(function(field){
      var rowArray=[];
      rowArray.push(dataRow.title.project);
      rowArray.push(dataRow.title.client);
      var timeAmount = field.time/1000/60/60;
      rowArray.push(Math.round10(timeAmount,-2)); //  put into hours deciaml
      rowArray.push(field.title.time_entry);
      array2dResponse.push(rowArray);
    });
    
  });
  
  var sa = SpreadsheetApp.getActive();
  if(sa.getSheetByName("Report Data")==null){
    sa.insertSheet("Report Data");
  }
  var ss = sa.getSheetByName("Report Data");
  ss.clear();
  ss.getRange(1, 1, array2dResponse.length,array2dResponse[0].length).setValues(array2dResponse);
  
}

function getDataURL(workspaceId, startDate, endDate, apiKey) {
  var startDateString = formatDate(startDate);
  var endDateString = formatDate(endDate);
  console.log(startDateString + ' ' + endDateString);
  
  return 'https://toggl.com/reports/api/v2/summary?workspace_id='+workspaceId+'&since='+startDateString+'&until='+endDateString+'&user_agent='+USER_AGENT;
}

function getUrlFetchGETOptions(apiKey){
  var authString = Utilities.base64Encode(apiKey+":api_token");
  return {
    "method": "get",
    "Content-Type" : "application/json",
    "muteHttpExceptions" : true,
    "headers" : {
      "Authorization" : "Basic " + authString
    }
  }
}
  
function formatDate(date) {
    var d = new Date(date),
        month = '' + (d.getMonth() + 1),
        day = '' + d.getDate(),
        year = d.getFullYear();

    if (month.length < 2) month = '0' + month;
    if (day.length < 2) day = '0' + day;

    return [year, month, day].join('-');
}
  
function sortFunctionByClient(a, b) {
    if (a[1] === b[1]) {
        return 0;
    }
    else {
        return (a[1] < b[1]) ? -1 : 1;
    }
}

function createSheets1stTimeImport(){
  var sa = SpreadsheetApp.getActive();
  var clientSheet = sa.insertSheet("Client Information");

  var headerRange=[];
  var headerValues=['Client','Address Line 1','Address Line 2','Address Line 3','Rate','Invoice Folder Id','Payment Terms','Email'];
  headerRange.push(headerValues)
  clientSheet.getActiveRange("A1:H1").setValues(headerRange);


  var invoiceTemplateSheet = sa.insertSheet('Invoice Template');
  invoiceTemplateSheet.getRange("A1:B1").merge();
  invoiceTemplateSheet.getRange("A1:B1").setValue('Name');
  invoiceTemplateSheet.getRange("A6:B6").merge();
  invoiceTemplateSheet.getRange("A6:B6").setValue('Bill To');
  invoiceTemplateSheet.getRange("A7:B7").merge();
  invoiceTemplateSheet.getRange("A8:B8").merge();
  invoiceTemplateSheet.getRange("A8:B8").merge();
  invoiceTemplateSheet.getRange("A9:B9").merge();
  invoiceTemplateSheet.getRange("A12:E12").merge();
  invoiceTemplateSheet.getRange("A12:E12").setValue('Description');
  invoiceTemplateSheet.getRange("E13").setValue('Hours');
  invoiceTemplateSheet.getRange("A13:E13").merge();
  invoiceTemplateSheet.getRange("A14:E14").merge();
  invoiceTemplateSheet.getRange("A15:E15").merge();
  invoiceTemplateSheet.getRange("A16:E16").merge();
  invoiceTemplateSheet.getRange("A17:E17").merge();
  invoiceTemplateSheet.getRange("A18:E18").merge();
  invoiceTemplateSheet.getRange("A19:E19").merge();
  invoiceTemplateSheet.getRange("A20:E20").merge();
  invoiceTemplateSheet.getRange("A21:E21").merge();
  invoiceTemplateSheet.getRange("A22:E22").merge();
  invoiceTemplateSheet.getRange("A23:E23").merge();
  invoiceTemplateSheet.getRange("A24:E24").merge();
  invoiceTemplateSheet.getRange("A25:E25").merge();
  invoiceTemplateSheet.getRange("A26:E26").merge();
  invoiceTemplateSheet.getRange("A27:E27").merge();
  invoiceTemplateSheet.getRange("A28:E28").merge();
  invoiceTemplateSheet.getRange("A29:E29").merge();
  invoiceTemplateSheet.getRange("A30:E30").merge();
  invoiceTemplateSheet.getRange("A31:E31").merge();
  invoiceTemplateSheet.getRange("A32:E32").merge();
  invoiceTemplateSheet.getRange("A33:E33").merge();
  invoiceTemplateSheet.getRange("A34:E34").merge();
  invoiceTemplateSheet.getRange("A35:E35").merge();
  invoiceTemplateSheet.getRange("A36:E36").merge();
  invoiceTemplateSheet.getRange("A37:E37").merge();
  invoiceTemplateSheet.getRange("A38:E38").merge();
  invoiceTemplateSheet.getRange("A39:E39").merge();
  invoiceTemplateSheet.getRange("A40:E40").merge();
  invoiceTemplateSheet.getRange("A41:E41").merge();
  invoiceTemplateSheet.getRange("A42:E42").merge();
  invoiceTemplateSheet.getRange("A43:E43").merge();
  invoiceTemplateSheet.getRange("A44:E44").merge();

  invoiceTemplateSheet.getRange("E1:F1").merge();
  invoiceTemplateSheet.getRange("E1:F1").setValue('Invoice');
  invoiceTemplateSheet.getRange("E2").setValue('Date');
  invoiceTemplateSheet.getRange("E3").setValue('Invoice #');
  invoiceTemplateSheet.getRange("E6:F6").merge();
  invoiceTemplateSheet.getRange("E6:F6").setValue('Summary');
  invoiceTemplateSheet.getRange("E7").setValue('Hour Total');
  invoiceTemplateSheet.getRange("E8").setValue('Rate/Hr');
  invoiceTemplateSheet.getRange("E9").setValue('Total');

  var dataSheet = sa.insertSheet('Report Data');

}
  
(function() {
  /**
   * Decimal adjustment of a number.
   *
   * @param {String}  type  The type of adjustment.
   * @param {Number}  value The number.
   * @param {Integer} exp   The exponent (the 10 logarithm of the adjustment base).
   * @returns {Number} The adjusted value.
   */
  function decimalAdjust(type, value, exp) {
    // If the exp is undefined or zero...
    if (typeof exp === 'undefined' || +exp === 0) {
      return Math[type](value);
    }
    value = +value;
    exp = +exp;
    // If the value is not a number or the exp is not an integer...
    if (value === null || isNaN(value) || !(typeof exp === 'number' && exp % 1 === 0)) {
      return NaN;
    }
    // Shift
    value = value.toString().split('e');
    value = Math[type](+(value[0] + 'e' + (value[1] ? (+value[1] - exp) : -exp)));
    // Shift back
    value = value.toString().split('e');
    return +(value[0] + 'e' + (value[1] ? (+value[1] + exp) : exp));
  }

  // Decimal round
  if (!Math.round10) {
    Math.round10 = function(value, exp) {
      return decimalAdjust('round', value, exp);
    };
  }
  // Decimal floor
  if (!Math.floor10) {
    Math.floor10 = function(value, exp) {
      return decimalAdjust('floor', value, exp);
    };
  }
  // Decimal ceil
  if (!Math.ceil10) {
    Math.ceil10 = function(value, exp) {
      return decimalAdjust('ceil', value, exp);
    };
  }
})();

  