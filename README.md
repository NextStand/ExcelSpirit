
> 作者：BLUE

> 日期：2017年9月20日/2018年3月22日

> 描述：Excel导入导出

> 版本：1.0.4


将文件引入到项目中，提供了3种方式

    #一般引用
    <script src="./dist/js/ExcelSpirit.js"></script>
    
    #commonJS规范
    const ExcelSpirit = require("ExcelSpirit");

    #ES6模块化语法
    import ExcelSpirit from 'ExcelSpirit';
    

## 目录
ExcelSpirit中一共提供了6个API

##### 1.Excel导入
- ExcelSpirit.excel2json4ev([callback][,binary])
- ExcelSpirit.excel2obj4ev([callback][,binary])
- ExcelSpirit.excel2json4drop(targetid[,callback][,binary])
- ExcelSpirit.excel2obj4drop(targetid[,callback][,binary])
##### 2.Excel导出
- html2excel(tbid[,name])
- json2excel(json[,titlemap1] [, filename])

## API详解

#### Excel导入
**# ExcelSpirit.excel2json4ev([callback] [,binary])**

基于一般事件的导入,通过解析Excel文件返回JSON格式数据，支持多文件同时上传,当选择多个文件同时上传时，返回一个二维数组组成的JSON ， ==不支持低版本IE==




- callback <Function> 【回调函数】 回调参数为解析的JSON
- binary <Boolean> 【解析方式】 是否解析成base-64编码进行读取，默认为  false，默认为流的方式进行读取
> 该方式适合于单表头，解析会按照第一行的内容作为返回json数据的key进行匹配，忽略内容为空的单元格

**注意：该方法调用必须基于事件**

```
    var d=document.getElementById("fileSelect1")
    d.addEventListener("click", function (e) {
        ExcelSpirit.excel2json4ev(function(json){
            document.getElementById("demo").innerHTML=JSON.stringify(json);
         });
    }, false);
    
    //返回数据格式
    /*
        [{"c_id":"01","c_name":"陆地水系 ","c_pid":"0"},{"c_id":"121","c_name":"河流","c_pid":"01"}]
    */
```

---

**# ExcelSpirit.excel2obj4ev([callback] [,binary])**

基于一般事件的导入,通过解析Excel文件返回对象格式数据， ==不支持低版本IE==

- callback <Function> 【回调函数】 回调参数为解析的JSON
- binary <Boolean> 【解析方式】 是否解析成base-64编码进行读取，默认为  false，默认为流的方式进行读取

> 该方式适合于复合表头，解析会按照行进行解析，每一行数据为一个数组，如果单元格为空则数组用null占位

**注意：该方法调用必须基于事件,暂时只支持单个Excel,如果选择多个，则解析第一个**
```
    var d=document.getElementById("fileSelect1")
    d.addEventListener("click", function (e) {
        ExcelSpirit.excel2obj4ev(function(json){
            document.getElementById("demo").innerHTML=JSON.stringify(json);
         });
    }, false);
    
    //返回数据格式
    /*
        {
            "Sheet1":[["01","陆地水系 ","0"],["121","河流","01"]],
            "Sheet2":[["1211","河源","01"],["1212","峡谷","01"]],
            "Sheet3":[["1213","河滩","01"],["1214","阶地","01"]]
        }
    */
```


---

**# ExcelSpirit.excel2json4drop(targetid [,callback] [,binary])**

将本地Excel文件拖拽到目标区域实现上传,通过解析Excel文件返回JSON格式数据,支持多文件同时上传,当选择多个文件同时上传时，返回一个二维数组组成的JSON ， ==不支持低版本IE==

- targetid <String> 【拖拽放置区域ID】
- callback <Function> 【回调函数】 回调参数为解析的JSON
- binary <Boolean> 【解析方式】 是否解析成base-64编码进行读取，默认为  false，默认为流的方式进行读取

> 该方式适合于单表头，解析会按照第一行的内容作为返回json数据的key进行匹配，忽略内容为空的单元格

```
<body>
    <div id="container">
     
    </div>
    <div id="demo"></div>
</body>
<script>
    ExcelSpirit.excel2json4drop("container",function(json){
        document.getElementById("demo").innerHTML=JSON.stringify(json);
    })
</script>

    //返回数据格式
    /*
        [{"c_id":"01","c_name":"陆地水系 ","c_pid":"0"},{"c_id":"121","c_name":"河流","c_pid":"01"}]
    */
```

---

**# ExcelSpirit.excel2obj4drop(targetid [,callback] [,binary])**

将本地Excel文件拖拽到目标区域实现上传,通过解析Excel文件返回对象格式数据， ==不支持低版本IE==


- targetid <String> 【拖拽放置区域ID】
- callback <Function> 【回调函数】 回调参数为解析的JSON
- binary <Boolean> 【解析方式】 是否解析成base-64编码进行读取，默认为  false，默认为流的方式进行读取
> 该方式适合于复合表头，解析会按照行进行解析，每一行数据为一个数组，如果单元格为空则数组用null占位

**注意：暂时只支持单个Excel,如果选择多个，则解析第一个**
```
<body>
    <div id="container">
     
    </div>
    <div id="demo"></div>
</body>
<script>
    ExcelSpirit.excel2obj4drop("container",function(json){
        document.getElementById("demo").innerHTML=JSON.stringify(json);
    })
</script>

    //返回数据格式
    /*
        {
            "Sheet1":[["01","陆地水系 ","0"],["121","河流","01"]],
            "Sheet2":[["1211","河源","01"],["1212","峡谷","01"]],
            "Sheet3":[["1213","河滩","01"],["1214","阶地","01"]]
        }
    */
```

---

#### Excel导出
**# ExcelSpirit.html2excel(tbid [,name])**

解析htmlDom节点的导出，同时导出表格的行内样式和复合表头（IE不够友好）

- tbid <String> 【数据源table DOM节点的ID】
- name <String> 【导出的Excel的sheet名称】 默认为"Worksheet"


```
<body>
    <div>
        <button type="button" onclick="ExcelSpirit.html2excel('tableExcel','TGD')">导出</button>
    </div>
    <div id="myDiv">
        <table id="tableExcel" width="100%" border="1" cellspacing="0" cellpadding="0">
            <tr>
                <th style="background-color:red;color:white;height:50px;border:1px solid yellow" colspan="5" align="center">html 表格导出到Excel</th>
            </tr>
            <tr>
                <td>列标题1</td>
                <td>列标题2</td>
                <td>类标题3</td>
                <td>列标题4</td>
                <td>列标题5</td>
            </tr>
            <tr>
                <td rowspan="2">aaa</td>
                <td>bbb</td>
                <td>ccc</td>
                <td>ddd</td>
                <td>eee</td>
            </tr>
            <tr>
                <td>BBB</td>
                <td>CCC</td>
                <td>DDD</td>
                <td>EEE</td>
            </tr>
        </table>
    </div>
</body>
```

---

**# ExcelSpirit.json2excel(json [,titlemap1] [, filename])**

基于JSON数据导出，==不支持复合表头和样式==


- json <Array> 【数据源json】
- titlemap1 <Object> 【表头和字段映射关系】 默认表头为json的key
- filename <String> 【导出Excel文件名】  默认为“下载”


```
var jsono = [{"t_id":"20170921","t_title": "橙子","t_price": "20","t_date":"2017-09-21"},
             {"t_id":"20170922","t_title": "苹果","t_price": "30","t_date":"2017-09-20"}
            ];
            
var titlemap={"t_id":"ID","t_title": "标题","t_price": "价格", "t_date":"生产日期"}

ExcelSpirit.json2excel(jsono,titlemap,"测试")
```

