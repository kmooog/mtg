
function ExportColorKing() {
  //Logger.clear();

  Logger.log('aaaa');
  var ss = SpreadsheetApp.openById(SpreadsheetApp.getActiveSpreadsheet().getId());
  var sheet = ss.getSheetByName("list"); 
  var sheet2 = ss.getSheetByName("colorking"); 
  //var sheet3 = ss.getSheetByName("forculculate"); 
  var scores = {"2-0":4,"1-1":0,"0-2":-4,"3-0":8,"2-1":3,"1-2":-3,"0-3":-8,"4-0":16,"3-1":4,"2-2":0,"1-3":-4,"0-4":-16,"5-0":32,"4-1":6,"3-2":3,"2-3":-3,"1-4":-6,"0-5":-32,"6-0":64,"5-1":11,"4-2":4,"3-3":0,"2-4":-4,"1-5":-11,"0-6":-64}

  //Logger.log(scores["3-0"]);
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
    WinLose = cell.offset(i,1).getValue() + "-" + cell.offset(i,2).getValue();

    col = cell.offset(i,3).getValue().split("");
    var score = scores[WinLose];
    var tmpscore = score / col.length

    var tmpcol = "W"
    if (col.indexOf(tmpcol) != -1){ColorRatioObj[name][tmpcol] = ColorRatioObj[name][tmpcol] + tmpscore}
    var tmpcol = "U"
    if (col.indexOf(tmpcol) != -1){ColorRatioObj[name][tmpcol] = ColorRatioObj[name][tmpcol] + tmpscore}
    var tmpcol = "B"
    if (col.indexOf(tmpcol) != -1){ColorRatioObj[name][tmpcol] = ColorRatioObj[name][tmpcol] + tmpscore}
    var tmpcol = "G"
    if (col.indexOf(tmpcol) != -1){ColorRatioObj[name][tmpcol] = ColorRatioObj[name][tmpcol] + tmpscore}
    var tmpcol = "R"
    if (col.indexOf(tmpcol) != -1){ColorRatioObj[name][tmpcol] = ColorRatioObj[name][tmpcol] + tmpscore}

    
    win = parseInt(cell.offset(i,1).getValue());
   lose = parseInt(cell.offset(i,2).getValue());
   
  

  
  //Logger.log(ScoreObj);
  }
  //Logger.log(ColorRatioObj);
  //return;
  sheet2.clear()
  sheet2.appendRow(["人","W","U","B","G","R"]);
  for (i = 0; i < namelist.length; i++) {
  Logger.log(ColorRatioObj);
  var total = (ColorRatioObj[namelist[i]]["W"]+ColorRatioObj[namelist[i]]["U"]+ColorRatioObj[namelist[i]]["B"]+ColorRatioObj[namelist[i]]["G"]+ColorRatioObj[namelist[i]]["R"] )
  total = 1;
  sheet2.appendRow([namelist[i],ColorRatioObj[namelist[i]]["W"]/total,ColorRatioObj[namelist[i]]["U"]/total,ColorRatioObj[namelist[i]]["B"]/total,ColorRatioObj[namelist[i]]["G"]/total,ColorRatioObj[namelist[i]]["R"]/total ]);
  }

    
  sheet2.getRange('B1').setBackground('gray');
  sheet2.getRange('C1').setBackground('blue');
  sheet2.getRange('D1').setBackground('black');
  sheet2.getRange('D1').setFontColor('white');
  
  sheet2.getRange('E1').setBackground('green');
  sheet2.getRange('F1').setBackground('red');
  var lastrow2 = sheet2.getLastRow();
  var cell2 = sheet2.getRange('a1');
  var Colormax = {"W":0,"U":0,"B":0,"G":0,"R":0};
  var ColormaxPerson = {"W":"none","U":"none","B":"none","G":"none","R":"none"};
  
  for(i=1;i<lastrow2;i++){
    var tmpcol = 1;
    var tmpcolor = "W";
    var ratio = parseFloat(cell2.offset(i,tmpcol).getValue());
    Logger.log([tmpcol,tmpcolor,ratio])
    if (ratio > Colormax[tmpcolor]){
      Colormax[tmpcolor] = ratio;
      ColormaxPerson[tmpcolor] = cell2.offset(i,0).getValue();
      }
    var tmpcol = 2;
    var tmpcolor = "U";
    var ratio = parseFloat(cell2.offset(i,tmpcol).getValue());
    Logger.log([tmpcol,tmpcolor,ratio])
    if (ratio > Colormax[tmpcolor]){
      Colormax[tmpcolor] = ratio;
      ColormaxPerson[tmpcolor] = cell2.offset(i,0).getValue();
      }
          var tmpcol = 3;
    var tmpcolor = "B";
    var ratio = parseFloat(cell2.offset(i,tmpcol).getValue());
    Logger.log([tmpcol,tmpcolor,ratio])
    if (ratio > Colormax[tmpcolor]){
      Colormax[tmpcolor] = ratio;
      ColormaxPerson[tmpcolor] = cell2.offset(i,0).getValue();
      }
          var tmpcol = 4;
    var tmpcolor = "G";
    var ratio = parseFloat(cell2.offset(i,tmpcol).getValue());
    Logger.log([tmpcol,tmpcolor,ratio])
    if (ratio > Colormax[tmpcolor]){
      Colormax[tmpcolor] = ratio;
      ColormaxPerson[tmpcolor] = cell2.offset(i,0).getValue();
      }
          var tmpcol = 5;
    var tmpcolor = "R";
    var ratio = parseFloat(cell2.offset(i,tmpcol).getValue());
    Logger.log([tmpcol,tmpcolor,ratio])
    if (ratio > Colormax[tmpcolor]){
      Colormax[tmpcolor] = ratio;
      ColormaxPerson[tmpcolor] = cell2.offset(i,0).getValue();
      }
  
  
  
    }
  sheet2.appendRow(["KING",ColormaxPerson["W"],ColormaxPerson["U"],ColormaxPerson["B"],ColormaxPerson["G"],ColormaxPerson["R"]]);
  Logger.log([ColormaxPerson["W"],ColormaxPerson["U"],ColormaxPerson["B"],ColormaxPerson["G"],ColormaxPerson["R"]]);


}

