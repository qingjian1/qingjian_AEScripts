# qingjian_AEScripts
这个用来放我写的一些ae脚本

## CsvTimeSheetScript v1.5
本脚本使用的是retas摄影表文件另存为的Csv表格;  
导入Csv表格文件后选择好图层和脚本面板中的选项点击应用即可;  
脚本需要识别表格中的“动画”这一格来确定位置，不会显示空图层;  
脚本可以使用两种应用方式，一种是关键帧的方式，一种是序列图层的方式;  
安装方法：把本文件复制到AE目录的 \Support Files\Scripts\ScriptUI Panels 目录中，启动AE后在“窗口”菜单里打开;  

## 一些问题
打开csv文件没有显示内容，有可能是csv文件的文本编码的问题，可以用一些工具转换成系统默认（GBK）或者UTF-8后再进行导入。  
还要注意，用表格软件（wps）打开csv文件，在保存的时候会将其编码更改为ANSI。所以要注意把文件改为UTF-8编码。这里我找到一个仓库,是[一个能够保存Unicode编码的CSV文件的WPS插件]("https://github.com/akof1314/WPS_COOL_CSV")。我只是搜到的，没用过，各位可以试试。  

by: 青涧  
