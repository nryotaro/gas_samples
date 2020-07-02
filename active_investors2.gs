
var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var primarySheetName = "'active investors'"
var activesheet = spreadsheet.getSheetByName('active investors');
var fundingRoundNames = activesheet.getRange('C1:F1').getValues()[0]
var numberOfInvestors = 90

function main() {
  createStageSheets(spreadsheet, fundingRoundNames)
}

function createStageSheets(spreadsheet, fundingRoundNames) {
  for(let i=0;i< fundingRoundNames.length;i++) {
    spreadsheet.insertSheet(fundingRoundNames[i], i+1)
  }
}

function fillProfiles() {
  for(let nameIndex = 0; nameIndex < fundingRoundNames.length; nameIndex++) {
    var stageSheet = spreadsheet.getSheetByName(fundingRoundNames[nameIndex])
    console.log(stageSheet);
    console.log(fundingRoundNames)
    var columnAlphabet = convertColumnNumberToAlphabet(nameIndex+3);
    for(let i=1;i<=numberOfInvestors + 1;i++) {
      stageSheet.getRange('A' + i).setValue('=' + primarySheetName + '!A' + i)
      stageSheet.getRange('B' + i).setValue('=' + primarySheetName + '!B' + i)
      stageSheet.getRange('C' + i).setValue('=' + primarySheetName + '!' + columnAlphabet +  + i)
    }
  }
  
}

function doSort() {
  
 for(let nameIndex = 0; nameIndex < fundingRoundNames.length; nameIndex++) {
   console.log(fundingRoundNames[nameIndex]);
    var typeSheet = spreadsheet.getSheetByName(fundingRoundNames[nameIndex]);
    typeSheet.getRange('A2:C' + (1 + numberOfInvestors)).sort({column: 3, ascending: false});
  }
}

function doChart() {
  for(let nameIndex = 0; nameIndex < fundingRoundNames.length; nameIndex++) {
    var typeSheet = spreadsheet.getSheetByName(fundingRoundNames[nameIndex]);
    var charts = typeSheet.getCharts();
    for (var i in charts) {
	  typeSheet.removeChart(charts[i]);
    }
    var chart = typeSheet.newChart()
    .asColumnChart()
    .addRange(typeSheet.getRange("B1:C" + (numberOfInvestors + 1)))
    .setPosition(5, 5, 25, 25)
    .build();
    typeSheet.insertChart(chart); 
  }
}

function convertColumnNumberToAlphabet(colmun_number) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var result = sheet.getRange(1, colmun_number);
  result = result.getA1Notation();
  result = result.replace(/\d/,'');
  return result;
}