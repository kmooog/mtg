

function ExportColorRatio() {
  //Logger.clear();

  Logger.log('aaaa');
  var ss = SpreadsheetApp.openById(SpreadsheetApp.getActiveSpreadsheet().getId());
  var sheet = ss.getSheetByName("list"); 
  var sheet2 = ss.getSheetByName("color"); 
  var scores = {"2-0":4,"1-1":0,"0-2":-4,"3-0":8,"2-1":3,"1-2":-3,"0-3":-8,"4-0":16,"3-1":4,"2-2":0,"1-3":-4,"0-4":-16,"5-0":32,"4-1":6,"3-2":3,"2-3":-3,"1-4":-6,"0-5":-32,"6-0":64,"5-1":11,"4-2":4,"3-3":0,"2-4":-4,"1-5":-11,"0-6":-64}

  Logger.log(scores["3-0"]);
  //return;


  var hondawin = 0;
  var cell = sheet.getRange('a1');
  var namelist = [];
  var lastrow = sheet.getLastRow()
  var numList = ["Wnum","Unum","Bnum","Rnum","Gnum"];
var data ={};   
var color = ["W","U","B","G","R"];
var scoreList = [];
//それぞれの試合結果

var ScoreObj = {}
var ColorRatioObj = {}

var win = 0;
var lose = 0;
var WinLose = 0;
var name = "";
var col = [];
  
  //名前の取得
  for(i=1;i<lastrow;i++){
  name = cell.offset(i,0).getValue();
  
    if (namelist.indexOf(cell.offset(i,0).getValue()) == -1){
      namelist.push(name);
      ColorRatioObj[name] = {"W":0,"U":0,"B":0,"G":0,"R":0};
      
    }
    col = cell.offset(i,3).getValue().split("");
    var tmpcol = "W"
    if (col.indexOf(tmpcol) != -1){ColorRatioObj[name][tmpcol] = ColorRatioObj[name][tmpcol] + 1}
    var tmpcol = "U"
    if (col.indexOf(tmpcol) != -1){ColorRatioObj[name][tmpcol] = ColorRatioObj[name][tmpcol] + 1}
    var tmpcol = "B"
    if (col.indexOf(tmpcol) != -1){ColorRatioObj[name][tmpcol] = ColorRatioObj[name][tmpcol] + 1}
    var tmpcol = "G"
    if (col.indexOf(tmpcol) != -1){ColorRatioObj[name][tmpcol] = ColorRatioObj[name][tmpcol] + 1}
    var tmpcol = "R"
    if (col.indexOf(tmpcol) != -1){ColorRatioObj[name][tmpcol] = ColorRatioObj[name][tmpcol] + 1}

    //Logger.log(ColorRatioObj);
    //return;
    
    win = parseInt(cell.offset(i,1).getValue());
   lose = parseInt(cell.offset(i,2).getValue());
   WinLose = cell.offset(i,1).getValue() + "-" + cell.offset(i,2).getValue();
   
  var score = scores[WinLose];
  //return;
  
  //Logger.log(ScoreObj);
  }
  Logger.log(ColorRatioObj);
  //return;
  sheet2.clear()
  sheet2.appendRow(["人","W","U","B","G","R"]);
  for (i = 0; i < namelist.length; i++) {
  var total = (ColorRatioObj[namelist[i]]["W"]+ColorRatioObj[namelist[i]]["U"]+ColorRatioObj[namelist[i]]["B"]+ColorRatioObj[namelist[i]]["G"]+ColorRatioObj[namelist[i]]["R"] )/100
  //var total = 1;
  sheet2.appendRow([namelist[i],ColorRatioObj[namelist[i]]["W"]/total,ColorRatioObj[namelist[i]]["U"]/total,ColorRatioObj[namelist[i]]["B"]/total,ColorRatioObj[namelist[i]]["G"]/total,ColorRatioObj[namelist[i]]["R"]/total ]);
  }

sheet2.getRange('B1').setBackground('gray');
sheet2.getRange('C1').setBackground('blue');
sheet2.getRange('D1').setBackground('black');
sheet2.getRange('D1').setFontColor('white');

sheet2.getRange('E1').setBackground('green');
sheet2.getRange('F1').setBackground('red');    
var last_row = sheet2.getLastRow();
var columnD = sheet2.getRange(2,4,last_row);
var valuesD = columnD.getValues(); //D列の値が入った配列
 
    
}

