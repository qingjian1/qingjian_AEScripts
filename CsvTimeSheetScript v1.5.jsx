/*
=========================================================================================================
CsvTimeSheetScript.jsx

本脚本使用的是retas摄影表文件另存为的Csv表格;
导入Csv表格文件后选择好图层和脚本面板中的选项点击应用即可;
脚本需要识别表格中的“动画”这一格来确定位置，不会显示空图层;
脚本可以使用两种应用方式，一种是关键帧的方式，一种是序列图层的方式;

安装方法：把本文件复制到AE目录的 \Support Files\Scripts\ScriptUI Panels 目录中，启动AE后在“窗口”菜单里打开;

by: 青涧
=========================================================================================================
*/

CsvTimeSheetScript(this)
//以AE本身为参数运行脚本主函数

function CsvTimeSheetScript(thisObj){
  //这是脚本主函数的定义

var scriptName = "CsvTimeSheetScript";
var version = 1.5;


function myScript_buildUI(thisObj){
    //这里是定义建立脚本UI的函数
    var myPanel = (thisObj instanceof Panel)?thisObj:new Window("palette",scriptName + " v" + version,undefined,{resizeable:true,closeButton:true});
    //判断传入的this里有没有panel，有就用，没有就新建一个window
       
var mainGroup = myPanel.add("group");
mainGroup.orientation = 'row';
mainGroup.alignment = ['fill','fill'];
mainGroup.alignChildren =['fill','fill'];
mainGroup.spacing = 20;

var listGroup = mainGroup.add("Panel");
listGroup.orientation = 'column';
listGroup.alignment = ['fill','fill'];
listGroup.alignChildren = ['fill','fill'];
listGroup.spacing =  5;


var text_1 = listGroup.add ("statictext", undefined, "文件名：");
var text_2 = listGroup.add ("edittext", undefined, "(空)");
var text_3 = listGroup.add ("statictext", undefined, "时间长度(帧)：");
var text_4 = listGroup.add ("edittext", undefined, "0");
text_1.preferredSize = [200, 20];
text_1.alignment = ['fill', 'top'];
text_2.preferredSize = [200, 20];
text_2.alignment = ['fill', 'top'];
text_3.preferredSize = [200, 20];
text_3.alignment = ['fill', 'top'];
text_4.preferredSize = [200, 20];
text_4.alignment = ['fill', 'top'];
var myList = listGroup.add ("listbox", undefined,'asm',
{numberOfColumns: 4, showHeaders: true,
columnTitles: ['图层名','起始张','结束张','缺张']});
//myList.bounds = {x:0, y:0, width:200, height:250}; 
myList.preferredSize = [200, 250];
myList.alignment = ['fill', 'fill'];
myList.alignChildren =['fill','fill'];
myList.itemSize = [200,30];

var item1 = myList.add ('item', '空');
item1.subItems[0].text = '';
item1.subItems[1].text = '';

var buttonGroup = mainGroup.add("group");
buttonGroup.orientation = 'column';
buttonGroup.alignment = ['right','fill'];
buttonGroup.alignChildren = 'fill';
buttonGroup.spacing =  12;

var myWayList = buttonGroup.add ("dropdownlist", undefined, ["关键帧","序列图层"]);
myWayList.selection = 0;
var sheetTypeList = buttonGroup.add ("dropdownlist", undefined, ["动画", "原画"]);
sheetTypeList.selection = 0;

if(app.settings.haveSetting(scriptName, "myWay")){myWayList.selection = app.settings.getSetting(scriptName, "myWay")}
if(app.settings.haveSetting(scriptName, "sheetType")){sheetTypeList.selection = app.settings.getSetting(scriptName, "sheetType")}
//有保存的方式就使用保存的方式
var but_1 = buttonGroup.add ("button", undefined, "导入");
var but_ap = buttonGroup.add ("button", undefined, "应用");
but_1.preferredSize  = [60,25];
but_ap.preferredSize  = [60,50];

listGroup.layout.layout(true);
listGroup.layout.resize();
buttonGroup.layout.layout(true);
buttonGroup.layout.resize();
mainGroup.layout.layout(true);
mainGroup.layout.resize();
myPanel.layout.layout(true);
myPanel.layout.resize();

//UI定义部分结束

//UI内全局变量。需要定义在回调函数外，否则其他回调函数无法使用
var mySheetArray; 
var listItemArray = new Array(); //选项数组
var AnimetionlistItemArray = new Array(); 
var keyAnimetionlistItemArray = new Array(); 
var timeSheetFile;

//UI回调函数部分


//更改方式
myWayList.onChange = function () {
    app.settings.saveSetting(scriptName, "myWay", myWayList.selection.index);
    }
//更改表类型
sheetTypeList.onChange = function () {
    app.settings.saveSetting(scriptName, "sheetType", sheetTypeList.selection.index);
    if(AnimetionlistItemArray == undefined){return;}
    
    listItemArray = [];//清空选项数组
    myList.removeAll();//清空UI选项
    
    if(sheetTypeList.selection.index == 0){
        listItemArray = AnimetionlistItemArray}
        else{listItemArray = keyAnimetionlistItemArray}
    
    text_2.text =  decodeURI (timeSheetFile.name);
    text_4.text =  mySheetArray.length - 2;
    for(i = 0; i < listItemArray.length;i++){
    var item1 = myList.add ('item', listItemArray[i][1]);
    item1.subItems[0].text = listItemArray[i][2];
    item1.subItems[1].text = listItemArray[i][3];
    item1.subItems[2].text = listItemArray[i][4];
        }
    
    }

//导入按钮的函数
but_1.onClick = function ()
{
    var csvFile = null;
    importCsv(csvFile);
}
function importCsv(csvFile)
{
        
    ( function(){
        
    if(csvFile == null ){
        var myPath = Folder.desktop.fullName;
        if(app.settings.haveSetting(scriptName, "myPath")){myPath = app.settings.getSetting(scriptName, "myPath")}
        timeSheetFile = new File(myPath).openDlg("选择CSV表格","CSV表格 : *.csv");    
    }else
    {
        timeSheetFile = csvFile;
    }

    if (timeSheetFile == null) {
        return;
    }
   app.settings.saveSetting(scriptName, "myPath", timeSheetFile.fullName);



    //转化为二维数组的方法定义
    function sheetToArray(sheetText){
        
    var rrr = new RegExp('"', "g");
    var arr = new RegExp("\n", "g");
    var a =  sheetText.replace(rrr,"");
    var b = a.split(arr);
    for(var i=0;i<b.length;i++){
        b[i] = b[i].split(","); 
        } 
    return b;
    }
    

        

    AnimetionlistItemArray = [];//清空选项数组
    keyAnimetionlistItemArray = [];
    myList.removeAll();//清空UI选项
        
    var sheetText = new Array();
 
    var AnimetionIndex;
    var keyAnimetionIndex;
    
    
    //一些常见编码库
    var encodingType = new Array("Default","UTF-8","CP932","CP936","CP950","EUC-JP");

    for (var encodingTypeIndex = 0; encodingTypeIndex < encodingType.length; encodingTypeIndex++)
    {
        if (!File.isEncodingAvailable(encodingType[encodingTypeIndex]) && encodingType[encodingTypeIndex] != "Default") continue;//如果不支持这个编码则跳过
        timeSheetFile.encoding = encodingType[encodingTypeIndex];
        var isOpended = timeSheetFile.open("r");
        sheetText = timeSheetFile.read();
        timeSheetFile.close();

        if( !isOpended)continue;//如果文件打开失败则跳过
        //读取文件
        var sheetTextString = File.decode(sheetText);
        mySheetArray = sheetToArray(sheetText);//得到数组
        if (mySheetArray.length == 0) continue;//如果分析数组失败则跳过


        //寻找关键字
        for (var wordIndex = 0; wordIndex < mySheetArray[0].length; wordIndex++)
            {
            if (mySheetArray[0][wordIndex] == "动画" ||
                mySheetArray[0][wordIndex] == "動畫" ||
                mySheetArray[0][wordIndex] == "動画" ||
                mySheetArray[0][wordIndex] == "中割" ||
                mySheetArray[0][wordIndex] == "上色" ||
                mySheetArray[0][wordIndex] == "中割り" ||
                mySheetArray[0][wordIndex] == "仕上げ"
        
            ) { AnimetionIndex = wordIndex; break; }
            }
            
        for(var wordIndex = 0;wordIndex<mySheetArray[0].length;wordIndex++)
        {
            if (mySheetArray[0][wordIndex] == "原画" ||
                mySheetArray[0][wordIndex] == "原畫" ||
                mySheetArray[0][wordIndex] == "原画")
            { keyAnimetionIndex = wordIndex; break; }
        }
        //如果找到其中一个则退出，否则换个编码试着找
        if (AnimetionIndex != undefined || keyAnimetionIndex != undefined) {
            break;
        }

    }


        
    
    
        


    if(AnimetionIndex == undefined){AnimetionIndex = mySheetArray[0].length+1;}
    else
    {
        for(var i = AnimetionIndex;i<mySheetArray[1].length;i++){
            var aaa = new Array();
            aaa.push (i);//选项数组的第0位是原数组的index
            aaa.push (mySheetArray[1][i]);//选项组的第1位是当前列的图层字母
            var min =mySheetArray.length, max=0;
            for(var k = 2;k<mySheetArray.length;k++){
                if(mySheetArray[k][i] == ""){continue}
                if(mySheetArray[k][i] == "×"){continue}
                if(isNaN(mySheetArray[k][i])){continue}
                var nowValue = mySheetArray[k][i] - 0;
                if(min > nowValue){min = nowValue;}
                if(max <nowValue){max = nowValue}
                }
            aaa.push(min,max);//第2,3位是最小，最大值
            

            if( min > max){
                continue;
             }//空图层不会列入选项中
     
           var numExist = new Array(max);
           var numMiss = new Array();
           for(k = 2;k<mySheetArray.length;k++){
                if(mySheetArray[k][i] == ""){continue}
                if(mySheetArray[k][i] == "×"){continue}
                if(isNaN(mySheetArray[k][i])){continue}
                numExist[mySheetArray[k][i]-1] = true;
                }
           
           for(var num_i = 0;num_i < numExist.length;num_i++){
               if(numExist[num_i] != true){numMiss.push(num_i+1);}
               }
           if(numMiss == null){
              aaa.push('');
              }else{
           aaa.push(numMiss);//第4位是缺张 
           }
            AnimetionlistItemArray.push(aaa);//动画选项数组完成
         }
            
    }    
        

    if(keyAnimetionIndex == undefined){keyAnimetionIndex =0;}
    else
    {
        for(var i = keyAnimetionIndex;i<AnimetionIndex;i++){
            //原画范围取到动画格的前一格以前
            var bbb = new Array();
            bbb.push (i);//选项数组的第0位是原数组的index
            bbb.push (mySheetArray[1][i]);//选项组的第1位是当前列的图层字母
            var min =mySheetArray.length, max=0;
            for(var k = 2;k<mySheetArray.length;k++){
                if(mySheetArray[k][i] == ""){continue}
                if(mySheetArray[k][i] == "×"){continue}
                if(mySheetArray[k][i] == "○"){continue}
                if(mySheetArray[k][i] == "●"){continue}
                if(isNaN(mySheetArray[k][i])){continue}
                var nowValue = mySheetArray[k][i] - 0;
                if(min > nowValue){min = nowValue;}
                if(max <nowValue){max = nowValue}
                }
            bbb.push(min,max);//第2,3位是最小，最大值
            

            if( min > max){
                continue;
         }//空图层不会列入选项中
 
       var numExist = new Array(max);
       var numMiss = new Array();
       for(k = 2;k<mySheetArray.length;k++){
            if(mySheetArray[k][i] == ""){continue}
            if(mySheetArray[k][i] == "×"){continue}
            if(mySheetArray[k][i] == "○"){continue}
            if(mySheetArray[k][i] == "●"){continue}
            if(isNaN(mySheetArray[k][i])){continue}
            numExist[mySheetArray[k][i]-1] = true;
            }
       
       for(var num_i = 0;num_i < numExist.length;num_i++){
           if(numExist[num_i] != true){numMiss.push(num_i+1);}
           }
       if(numMiss == null){
          bbb.push('');
          }else{
       bbb.push(numMiss);//第4位是缺张 
       }
        keyAnimetionlistItemArray.push(bbb);//原画选项数组完成
     }
        
    }

  
    //以下是显示到界面中

 
    if(sheetTypeList.selection.index == 0){
        listItemArray = AnimetionlistItemArray}
        else{listItemArray = keyAnimetionlistItemArray}

    text_2.text =  decodeURI (timeSheetFile.name);
    text_4.text =  mySheetArray.length - 2;
    


    for(i = 0; i < listItemArray.length;i++){
    var item1 = myList.add ('item', listItemArray[i][1]);
    item1.subItems[0].text = listItemArray[i][2];
    item1.subItems[1].text = listItemArray[i][3];
    item1.subItems[2].text = listItemArray[i][4];
        }
    
    
    }())
}
//导入按钮函数结束


//开始写应用按钮的函数

but_ap.onClick = function Apply(){
    app.beginUndoGroup("应用");
    if(myWayList.selection.index == 1){layerWayApply()}
        else{keyWayApply()}
    app.endUndoGroup();
    
    function layerWayApply(){    
    var myComp  , myLayers ;
     if(app.project.activeItem != undefined ){
       myComp = app.project.activeItem;
      if(myComp.selectedLayers[0] != null ){
        myLayers = myComp.selectedLayers;
      }}
    if(myComp == undefined){alert ("请打开一个合成",scriptName);return;}
    if(myLayers == undefined){alert ("请选择图层",scriptName);return;}   
    
    if(myList.selection == undefined){alert ("请选择列表中的选项",scriptName);return;}   
    if(myLayers.length != myList.selection.subItems[1]){alert("选择的图层数量不符合",scriptName);return;}
    
    myLayers.sort(compare);//对选择的图层进行排序
    var nowSheetIndex =  listItemArray[myList.selection.index][0] ;
    var newLayers = new Array();
    for(var i=2, j=0;i< mySheetArray.length;i++,j++){

        var nowGrid = mySheetArray[i][nowSheetIndex];
        switch(nowGrid){
            case "○":
            case "●":
            case "":
                if(newLayers == null || newLayers[newLayers.length - 1] == null){break}
                newLayers[newLayers.length - 1].outPoint += myComp.frameDuration;
                break;               
            case "×":
                newLayers.push(null);
                break;                
            default:
                if(isNaN(nowGrid)){break};
                var nowLayer = myLayers[nowGrid-1].duplicate();
                 newLayers.push(nowLayer);
                 nowLayer.moveBefore(myLayers[0]);
                 nowLayer.startTime = j*myComp.frameDuration
                 nowLayer.outPoint = nowLayer.startTime + myComp.frameDuration;
             
            }
        
        }//复制排列完成

    for(var i=0;i<myLayers.length;i++){
        myLayers[i].remove();
        }//删除原有
    
    }//到这里是序列图层方式应用
  
  
    function keyWayApply(){
      var myComp  , myLayers ,myLayer;
     if(app.project.activeItem != undefined ){
       myComp = app.project.activeItem;
      if(myComp.selectedLayers[0] != null ){
        myLayers = myComp.selectedLayers;
      }}
    if(myComp == undefined){alert ("请打开一个合成",scriptName);return;}
    if(myLayers == undefined){alert ("请选择图层",scriptName);return;}

    if(myList.selection == undefined){alert ("请选择列表中的选项",scriptName);return;}   
    if(myLayers.length != 1){alert("只能选择一个图层",scriptName);return;}
    
     myLayer = myLayers[0];
    
     if(myLayer.canSetTimeRemapEnabled){myLayer.timeRemapEnabled = true;}
     else{alert("该图层不能进行时间重映射",scriptName); return}
     myLayer.timeRemap.removeKey(2);
     var checkStart = true; //检查是否开始打表
     var wipeFx;//这里声明擦除的效果
     var nowSheetIndex =  listItemArray[myList.selection.index][0];
     for(var i=2, j=0;i< mySheetArray.length;i++,j++){
        var nowGrid = mySheetArray[i][nowSheetIndex];
        
        //myLayer.timeRemap.setValuesAtTimes(j*myComp.frameDuration,nowGrid);
        switch(nowGrid){
            case "":
            case "○":
            case "●":
                break;               
            case "×":
                if(wipeFx == undefined){
                wipeFx = myLayer.effect.addProperty("ADBE Linear Wipe");
                wipeFx.property(1).setValueAtTime((j-1)*myComp.frameDuration,0);
                }
                wipeFx.property(1).setValueAtTime(j*myComp.frameDuration,100);

                break; 
            default:
                if(isNaN(nowGrid)){break};
                if(wipeFx != undefined && wipeFx.property(1).keyValue(wipeFx.property(1).numKeys) == 100){wipeFx.property(1).setValueAtTime(j*myComp.frameDuration,0);}
                myLayer.timeRemap.setValueAtTime(j*myComp.frameDuration,(nowGrid-1) * myComp.frameDuration);
                if(checkStart){ 
                    if(j != 0){myLayer.timeRemap.removeKey(1);
                        myLayer.inPoint = j*myComp.frameDuration;     
                        };
                    checkStart = false;
                    };//打表开始后裁剪开头
            }        
        
        }//加关键帧完成
     for(i = 1; i<= myLayer.timeRemap.numKeys; i++){
        myLayer.timeRemap.setInterpolationTypeAtKey(i, KeyframeInterpolationType.LINEAR , KeyframeInterpolationType.HOLD);
        }
    if(wipeFx != undefined){
     for(i = 1; i<= wipeFx.property(1).numKeys; i++){
        wipeFx.property(1).setInterpolationTypeAtKey(i, KeyframeInterpolationType.LINEAR , KeyframeInterpolationType.HOLD);
        }
    }
        //都变成定格关键帧
     myLayer.outPoint =  (mySheetArray.length - 2)*myComp.frameDuration;
    }//到这里是关键帧方式应用
}
//应用按钮函数到此


  myPanel.layout.layout(true);
  myPanel.layout.resize();
  myPanel.onResizing = myPanel.onResize = function () {  
    myPanel.layout.resize();
    mainGroup.layout.resize();
    buttonGroup.layout.resize();
    listGroup.layout.resize();
    this.layout.resize();
  }
  return myPanel;
}//UI建立部分结束




  var myScriptPal = myScript_buildUI(thisObj);
  if(myScriptPal!=null&&myScriptPal instanceof Window){
    myScriptPal.show();
  };
  //一切准备就绪，建立UI


}
//主体到这个大括号结束

//这是比较图层index的函数，用于排列选择中的图层
var compare = function (layer1, layer2) {
  var val1 = layer1.index;
  var val2 = layer2.index;
  if (val1 < val2){
    return -1;
  }
  else if (val1 > val2) {
    return 1;
  }
  else {return 0 ;}
}
