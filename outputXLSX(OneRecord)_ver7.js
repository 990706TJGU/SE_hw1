/*
 * kintone to Excel sample program
 * Copyright (c) 2018 Cybozu
 *
 * Licensed under the MIT License
*/
(function() {

  'use strict';

  /* global $ */
  /* global kintone */
  /* global XLSX */
  /* global saveAs */
  /* global Blob */
   
  var CYB = {};
  var CYB002 = {};
  var fileld_name ={};
  var jsono = [];

  var fieldValue = ''; 
  var TableValue1 = ''; 
  var TableValue2 = '';
  var TableValue3 = '';

  var rec1 = {};
  var rec2 = {};
  var rec3 = {};
  var rec4 = {};
  var rec5 = {};
  var rec6 = {};

  var Temp_rec = {};

  // xlsx 出力する kintone のフィールド
  CYB.cols = ['title', 'detail', '人數','時間','Table'];//Haven't do Table
  //CYB002.cols = ['日期', '工作地點', '時數'];
 

  document.body.innerHTML += '<a href="" download="这里是下载的文件名.xlsx" id="hf"></a>';//add some contents to html

  

var tmpDown; //导出的二进制对象




function downloadExl(json, type) {
    var tmpdata = json[0];
    json.unshift({});
    var keyMap = []; //获取keys
    //keyMap =Object.keys(json[0]);
    for (var k in tmpdata) {
        keyMap.push(k);
        json[0][k] = k;
        console.log(keyMap);
        console.log(json[0][k]);
    }
  var tmpdata = [];//用来保存转换好的json 
        json.map((v, i) => keyMap.map((k, j) => Object.assign({}, {
            v: v[k],
            position: (j > 25 ? getCharCol(j) : String.fromCharCode(65 + j +2)) + (i + 13)//控制欄位在xlsx的位置 j:x軸 i:y軸
        }))).reduce((prev, next) => prev.concat(next)).forEach((v, i) => tmpdata[v.position] = {
            v: v.v
        });
        var outputPos = Object.keys(tmpdata); //设置区域,比如表格从A1到D10

        tmpdata["C13"].s = { font: { sz: 14, bold: true, color: { rgb: "FFFFAA00" } }, fill: { bgColor: { indexed: 64 }, fgColor: { rgb: "FFFF00" } } };//<====设置xlsx单元格样式
        

        var tmpWB = {
            SheetNames: ['mySheet'], //保存的表标题
            Sheets: {
                'mySheet': Object.assign({},
                    tmpdata, //内容
                    {
                        //'!ref': outputPos[0] + ':' + outputPos[outputPos.length - 1], //设置填充区域
                        '!ref':'B1:F30', //设置填充区域
                        C5: { v: "記錄號碼：  " + Temp_rec["レコード番号"]['value'] },
                        C6: { v: "標題：  " + Temp_rec["title"]['value'] },
                        C7: { v: "人數：  " + Temp_rec["人數"]['value'] },
                        C8: { v: "時間：  " + Temp_rec["時間"]['value'] },
                    })
            }
        };

        tmpWB["Sheets"]["mySheet"]["C5"].s = { font: { sz: 14, bold: true, color: { rgb: "FFFFAA00" } }, fill: { bgColor: { indexed: 64 }, fgColor: { rgb: "FFFF00" } } };//<====设置xlsx单元格样式
 

 
 
         
        /*var a = tmpWB["Sheets"]["mySheet"]["B5"]["v"];
        var b = tmpWB["Sheets"]["mySheet"]["B6"]["v"];
        var c = tmpWB["Sheets"]["mySheet"]["B7"]["v"];
        var d = tmpWB["Sheets"]["mySheet"]["B8"]["v"];*/
       // tmpWB["Sheets"]["mySheet"]["D5"]["v"] = {};
        /*if (tmpWB["Sheets"]["mySheet"]["B5"]["v"] === undefined) {
            tmpWB["Sheets"]["mySheet"]["B5"]["v"] = ("記錄號碼： " + str1);
        };*/

        /*tmpWB["Sheets"]["mySheet"]["B5"]["v"] = (typeof tmpWB["Sheets"]["mySheet"]["B5"]["v"] === 'undefined') ? ("記錄號碼： " + str1) : tmpWB["Sheets"]["mySheet"]["B5"]["v"];
        tmpWB["Sheets"]["mySheet"]["B6"]["v"] = (typeof tmpWB["Sheets"]["mySheet"]["B6"]["v"] === 'undefined') ? ("標題 " + str2) : tmpWB["Sheets"]["mySheet"]["B6"]["v"];
        tmpWB["Sheets"]["mySheet"]["B7"]["v"] = (typeof tmpWB["Sheets"]["mySheet"]["B7"]["v"] === 'undefined') ? ("人數 " + str3) : tmpWB["Sheets"]["mySheet"]["B7"]["v"];
        tmpWB["Sheets"]["mySheet"]["B8"]["v"] = (typeof tmpWB["Sheets"]["mySheet"]["B8"]["v"] === 'undefined') ? ("時間 " + str4) : tmpWB["Sheets"]["mySheet"]["B8"]["v"];*/

        /*tmpWB["Sheets"]["mySheet"]["B6"]["v"] = "記錄號碼： " + str1;
        tmpWB["Sheets"]["mySheet"]["B7"]["v"] = "標題： " + str2;
        tmpWB["Sheets"]["mySheet"]["B8"]["v"] = "人數： " + str3;
        tmpWB["Sheets"]["mySheet"]["B9"]["v"] = "時間： " + str4;*/

        console.log(tmpWB);
        tmpDown = new Blob([s2ab(XLSX.write(tmpWB, 
            {bookType: (type == undefined ? 'xlsx':type),bookSST: false, type: 'binary'}//这里的数据是用来定义导出的格式类型
            ))], {
            type: ""
        }); //创建二进制对象写入转换好的字节流
        console.log(json[0][k]);
    var href = URL.createObjectURL(tmpDown); //创建对象超链接
    document.getElementById("hf").href = href; //绑定a标签
    document.getElementById("hf").click(); //模拟点击实现下载
    setTimeout(function() { //延时释放
        URL.revokeObjectURL(tmpDown); //用URL.revokeObjectURL()来释放这个object URL
    }, 100);
}

function s2ab(s) { //字符串转字符流
    var buf = new ArrayBuffer(s.length);
    var view = new Uint8Array(buf);
    for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
    return buf;
}
 // 将指定的自然数转换为26进制表示。映射关系：[0-25] -> [A-Z]。
function getCharCol(n) {
    let temCol = '',
    s = '',
    m = 0
    while (n > 0) {
        m = n % 26 + 1
        s = String.fromCharCode(m + 64) + s
        n = (n - m) / 26
    }
    return s
}


  // 一覧画面が開かれたとき
  kintone.events.on(['app.record.detail.show'], function(event) {

    var rec = event.record; 
    console.log(rec);  

    Temp_rec = rec;
    console.log(Temp_rec);  


    /*
     
    TableValue1 = rec["Table"]['value'][0]['value']['日期']['value']; 
    TableValue2 = rec["Table"]['value'][0]['value']['工作地點']['value']; 
    TableValue3 = rec["Table"]['value'][0]['value']['時數']['value']; 

    rec1 = {"記錄號碼": rec["レコード番号"]['value']};
    rec2 = {"標題": rec["title"]['value']};
    rec3 = {"人數": rec["人數"]['value']};
    rec4 = {"時間": rec["時間"]['value']};
    rec5 = {"明細": rec["detail"]['value']};
    rec6 = {
        "記錄號碼": rec["レコード番号"]['value'],
        "標題": rec["title"]['value'],
        "人數": rec["人數"]['value'],
        "時間": rec["時間"]['value'],
        "明細": rec["detail"]['value']
    };*/

    //jsono.push(rec6); 

 
    for(var i in rec["Table"]['value']){
        const copyOfMyArray = { //测试数据
            "日期": rec["Table"]['value'][i]['value']['日期']['value'],
            "工作地點": rec["Table"]['value'][i]['value']['工作地點']['value'],
            "時數": rec["Table"]['value'][i]['value']['時數']['value']
        };

        if(rec["Table"]['value'][i]['value']['日期']['value']!= null){ 
            jsono.push(copyOfMyArray);
        }
        
    }

    

    console.log(jsono);   





    var myRecordButton = document.createElement('button');
        myRecordButton.id = 'my_button';
        myRecordButton.innerHTML = 'Click Me!';

        myRecordButton.onclick = function() {  
            downloadExl(jsono);
            
        };

        // Set the button on the header
        kintone.app.record.getHeaderMenuSpaceElement().appendChild(myRecordButton);

  
      return;
  });

  kintone.events.on(['app.record.edit.submit.success'], function(event) {
    //window.location.reload();
    return;
});



})();