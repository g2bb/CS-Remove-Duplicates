//Remove duplicate rows based on entry in specific column
//Joe Hays 2018

var ss=SpreadsheetApp.getActiveSpreadsheet();
var sheet=ss.getActiveSheet();
var ui=SpreadsheetApp.getUi();
var data=sheet.getDataRange().getValues();

//This calls the onOpen function when the Add-on is first installed
function onInstall(e){
    onOpen(e);
}

// The onOpen function is executed automatically every time a Spreadsheet is loaded
function onOpen(e){
    var menu=SpreadsheetApp.getUi().createAddonMenu(); // Or DocumentApp or FormApp.
    menu.addItem('Run','Input');
    menu.addToUi();
}

//Prompts user for some initial inputs
function Input(){
    var htmlOutput=HtmlService
        .createHtmlOutputFromFile('sidebar')
        .setTitle('Remove Duplicates')
        .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    ui.showSidebar(htmlOutput);
}

//This will highlight all the offending rows
function highlight(Term){
    if(Term !="--CANCEL"){
        var Col=findCol(Term);
        if(Col !=undefined){
            ui.alert('Searching for duplicates in "' + Term + '"\n\n Feel free to work on other tabs while this runs.\nThis may take a while.');
            var NumRemoved=highlightDuplicates(Col);
            ui.alert('The script has run sucessfully!\n' + NumRemoved + ' rows will be highlighted');
        }else{
            ui.alert('The script failed to run!' + '\n\nFailed to find column titled "' + Term + '"\n Please check your column header and try again.');
        }
    }
}

//This will remove all the offending rows
function remove(Term){
    if(Term !="--CANCEL"){
        var Col=findCol(Term);
        if(Col !=undefined){
            ui.alert('Searching for duplicates in "' + Term + '"\n\n Feel free to work on other tabs while this runs.\nThis may take a while.');
            var NumRemoved=removeDuplicates(Col);
            ui.alert('The script has run sucessfully!\n' + NumRemoved + ' rows will be removed');
        }else{
            ui.alert('The script failed to run!' + '\n\nFailed to find column titled "' + Term + '"\n Please check your column header and try again.');
        }
    }
}

// Scans row one to find user specified column
function findCol(Term){
    var searchString=Term;
    var columnIndex;
    for(var i=0; i<data[0].length; i++){
        if(data[0][i]==searchString){
            columnIndex=i;
            break;
        }
    }
    return columnIndex;
}

//Removes duplicate rows from sheet
function removeDuplicates(Col){
    var newData=[];
    var Num=0;
    for(var i=0; i<data.length; i++){
        var duplicate=false;
        for(var j=0; j<newData.length; j++){
            if(data[i][Col]==newData[j][Col]){
                duplicate=true;
                Num+=1;
                break
            }
        }
        if(!duplicate){
            newData.push(data[i]);
        }
    }
    sheet.clearContents();
    sheet.getRange(1,1,newData.length,newData[0].length)
        .setValues(newData);
    return Num;
}


//Highlights duplicate rows from sheet
function highlightDuplicates(Col){
    var newData=[];
    var Num=0;
    for(var i=0; i<data.length; i++){
        var duplicate=false;
        for(var j=0; j<newData.length; j++){
            if(data[i][Col]==newData[j][Col]){
                duplicate=true;
                var last_col=sheet.getLastColumn(); // last populated column in sheet
                var range=sheet.getRange(i,1,1,last_col); //gets the range corresponding with the last populated row in the sheet
                range.setBackgroundColor("yellow");
                Num+=1;
                break
            }
        }
        if(!duplicate){
            newData.push(data[i]);
        }
    }
    return Num;
}