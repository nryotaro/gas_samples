function myFunction() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastColumn();
  var fundingRoundNames = sheet.getRange('C1:AE1').getValues()[0]
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  for(let i=0;i< fundingRoundNames.length;i++) {
    ss.insertSheet(fundingRoundNames[i], i+1)
  }
}
var ss = SpreadsheetApp.getActiveSpreadsheet();
var countSheet = 'count'
var numInvestors = 91
var fundingRoundNames = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(countSheet).getRange('C1:AE1').getValues()[0]
function hoge() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  for(let nameIndex = 0; nameIndex < fundingRoundNames.length; nameIndex++) {
    var typeSheet = ss.getSheetByName(fundingRoundNames[nameIndex]);
      for(let i=1;i<=91;i++) {
        typeSheet.getRange('A' + i).setValue('=count!A' + i);
        typeSheet.getRange('B' + i).setValue('=count!B' + i);
      }
  }
}
function fuge() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  for(let nameIndex = 0; nameIndex < fundingRoundNames.length; nameIndex++) {
    var typeSheet = ss.getSheetByName(fundingRoundNames[nameIndex]);
    var columnAlphabet = wmap_column_convert(nameIndex+3);
    for(let i=1;i<=numInvestors;i++) {
        typeSheet.getRange('C' + i).setValue('=' + countSheet + '!' + columnAlphabet + i);
      }
  }
}
function doSort() {
 for(let nameIndex = 0; nameIndex < fundingRoundNames.length; nameIndex++) {
    var typeSheet = ss.getSheetByName(fundingRoundNames[nameIndex]);
    typeSheet.getRange('A2:C' + numInvestors).sort({column: 3, ascending: false});
  }
}
function doChart() {
  for(let nameIndex = 0; nameIndex < fundingRoundNames.length; nameIndex++) {
    var typeSheet = ss.getSheetByName(fundingRoundNames[nameIndex]);
    var charts = typeSheet.getCharts();
    for (var i in charts) {
	  typeSheet.removeChart(charts[i]);
    }
    var chart = typeSheet.newChart()
    .asColumnChart()
    .addRange(typeSheet.getRange("B1:C" + numInvestors))
    .setPosition(5, 5, 25, 25)
    .build();
    typeSheet.insertChart(chart); 
  }
}
function wmap_column_convert(colmun_number) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var result = sheet.getRange(1, colmun_number);
  result = result.getA1Notation();
  result = result.replace(/\d/,'');
  return result;
}
