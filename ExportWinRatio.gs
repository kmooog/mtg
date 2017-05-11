function ExportWinRatio() {

  var ss = SpreadsheetApp.openById(SpreadsheetApp.getActiveSpreadsheet().getId());
  var sheet = ss.getSheetByName("list"); 
  var sheet2 = ss.getSheetByName("WinRatio"); 
  var scores = {"2-0":4,"1-1":0,"0-2":-4,"3-0":8,"2-1":3,"1-2":-3,"0-3":-8,"4-0":16,"3-1":4,"2-2":0,"1-3":-4,"0-4":-16,"5-0":32,"4-1":6,"3-2":3,"2-3":-3,"1-4":-6,"0-5":-32,"6-0":64,"5-1":11,"4-2":4,"3-3":0,"2-4":-4,"1-5":-11,"0-6":-64}
  var hondawin = 0;
  var cell = sheet.getRange('a1');
  var namelist = [];
  var lastrow = sheet.getLastRow()
  var numList = ["Wnum","Unum","Bnum","Rnum","Gnum"];
var ratioList = ["Wratio","Uratio","Bratio","Rratio","Gratio"];
var data ={};   
var color = "WUBGR";
var scoreList = [];
//それぞれの試合結果
var ScoreObj = {}
var countVS = {}
var win = 0;
var lose = 0;
var WinLose = 0;
var name = "";
  //名前の取得
  for(i=1;i<lastrow;i++){
    if (namelist.indexOf(cell.offset(i,0).getValue()) == -1){
      namelist.push(cell.offset(i,0).getValue());
    } 
    win = parseInt(cell.offset(i,1).getValue());
   lose = parseInt(cell.offset(i,2).getValue());
   name = cell.offset(i,0).getValue();
  var score = scores[WinLose];
  if (namelist.indexOf(name) != -1 && ScoreObj[name]){
    ScoreObj[name] = ScoreObj[name] + win;
    countVS[name] = countVS[name] + win + lose;
    }else if (namelist.indexOf(name) != -1){
    ScoreObj[name] = + win;
    countVS[name] =  win + lose;    
    }
  }
  Logger.log(namelist.length);
  sheet2.clear()
  sheet2.appendRow(["人","勝率 (%)", "対戦回数"]);
  for (i = 0; i < namelist.length; i++) {
  sheet2.appendRow([namelist[i],ScoreObj[namelist[i]]/ countVS[namelist[i]] * 100,countVS[namelist[i]]]);
  }
var range = sheet2.getRange("A2:B");
return;  
}


