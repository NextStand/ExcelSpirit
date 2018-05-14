/**************************************
-- 功能：Excel导入导出
-- 作者：BLUE
-- 时间：2017年9月20日
-- 版本：1.0.3
**************************************/
const XLSX = require('xlsx')
let _ExcelSpirit = class _ExcelSpirit {
    //检测浏览器
    static _checkBrowser() {
        //获取浏览器代理数据信息
        let ua = window.navigator.userAgent;
        if (ua.indexOf("MSIE") >= 0) {
            return 'ie';
        }
        else if (ua.indexOf("Firefox") >= 0) {
            return 'Firefox';
        }
        else if (ua.indexOf("Chrome") >= 0) {
            return 'Chrome';
        }
        else if (ua.indexOf("Opera") >= 0) {
            return 'Opera';
        }
        else if (ua.indexOf("Safari") >= 0) {
            return 'Safari';
        }
    }
    /// <summary>
    /// 序列化解码之后编码成base-64编码的的ASCII字符串
    /// </summary>
    /// <param name="s">预编码字符串</param>
    /// <returns>base-64编码的的ASCII字符串</returns>
    static getbase64(s) {
        //解码转义序列、进行特殊字符16进制转义
        return window.btoa(unescape(encodeURIComponent(s)))
    }
    /// <summary>
    /// 序列化解码之后编码成base-64编码的的ASCII字符串
    /// </summary>
    /// <param name="template">模板字符串</param>
    /// <param name="dataobj">模板数据对象</param>
    /// <returns>模板被数据替换之后的字符串</returns>
    static format(template, dataobj) {
        return template.replace(/{(\w+)}/g, (value, key) => { return dataobj[key]; });
    }
    /// <summary>
    /// 非IE浏览器导出，不确定IE是否OK，IE应该还是要调ActiveXObject
    /// </summary>
    /// <param name="tbid">table标签的id</param>
    /// <param name="sheetname">excel左下角的sheet名称</param>
    /// <returns>null</returns>
    static _exportExcel(tbid, sheetname) {
        //让浏览器调用excel，以流的方式将传入的base64数据写入excel
        let uri = 'data:application/vnd.ms-excel;base64,',
            //定义模板，以xmlns定义单独命名空间，为编译base64编码做准备
            template = `<html 
            xmlns:o="urn:schemas-microsoft-com:office:office" 
            xmlns:x="urn:schemas-microsoft-com:office:excel" 
            <head>
                <!--[if gte mso 9]>
                <xml>
                <x:ExcelWorkbook>
                    <x:ExcelWorksheets>
                        <x:ExcelWorksheet>
                            <x:Name>{sheetname}</x:Name>
                            <x:WorksheetOptions>
                                <x:DisplayGridlines/>
                            </x:WorksheetOptions>
                        </x:ExcelWorksheet>
                    </x:ExcelWorksheets>
                </x:ExcelWorkbook>
                </xml>
                <![endif]-->
                <meta charset="UTF-8">
            </head>
            <body>
                <table>{tbexcel}</table>
            </body>
            </html>`;
        //获取DOM对象
        if (!tbid.nodeType)
            tbid = document.getElementById(tbid);
        if (!tbid.nodeType) {
            throw new Error("The DOM node is not found");
            return false;
        } else {
            //定义要往模板中替换的数据对象
            let ctx = { sheetname, tbexcel: tbid.innerHTML };
            //将编译的base-64 URI植入地址栏，让浏览器自动执行写入流
            window.location.href = uri + this.getbase64(this.format(template, ctx))
        }
    }
    //为IE准备的，选择保存路径
    static _browseFolder(path) {
        try {
            var Message = "\u8bf7\u9009\u62e9\u6587\u4ef6\u5939"; //选择框提示信息
            var Shell = new ActiveXObject("Shell.Application");
            var Folder = Shell.BrowseForFolder(0, Message, 64, 17); //起始目录为：我的电脑
            //var Folder = Shell.BrowseForFolder(0, Message, 0); //起始目录为：桌面
            if (Folder != null) {
                Folder = Folder.items(); // 返回 FolderItems 对象
                Folder = Folder.item(); // 返回 Folderitem 对象
                Folder = Folder.Path; // 返回路径
                if (Folder.charAt(Folder.length - 1) != "\\") {
                    Folder = Folder + "\\";
                }
                document.getElementById(path).value = Folder;
                return Folder;
            }
        }
        catch (e) {
            alert(e.message);
        }
    }
    /// <summary>
    /// IE利用ActiveXObject调用Excel，很奇怪，跟其他浏览器不一样，
    /// 样式不生效，合并单元格还有问题
    /// </summary>
    /// <param name="tbid">table标签的id</param>
    /// <param name="sheetname">excel左下角的sheet名称</param>
    /// <returns>null</returns>
    static _exportExcel4IE(tbid) {
        //使用微软自己的ActiveX调用excel程序
        let curTb = document.getElementById(tbid),
            excelApp = new ActiveXObject("Excel.Application"),
            xlBook = excelApp.Workbooks.Add(),             //新增工作簿
            ExcelSheet = xlBook.ActiveSheet,              //获取工作sheet
            lenRow = curTb.rows.length;                      //html行数
        if (!curTb) {
            throw new Error("The DOM node is not found");
            return false;
        } else {
            for (i = 0; i < lenRow; i++) {
                let lenCell = curTb.rows(i).cells.length;
                for (j = 0; j < lenCell; j++) {
                    ExcelSheet.Cells(i + 1, j + 1).value = curTb.rows(i).cells(j).innerText;
                }
            }
            excelApp.Visible = true;
            ExcelSheet.SaveAs("E:\\sheet.xls");//保存路径先写死
        }
    }
    /// <summary>
    /// ExcelSpirit.html2excel(tbid)
    /// 公共调用将html表格里面的数据导出到excel，
    /// </summary>
    /// <param name="tbid">table标签的id</param>
    /// <returns>null</returns>
    static html2excel(tbid, name = "Worksheet") {
        let _bowerType = this._checkBrowser();
        if (_bowerType === 'ie') {
            this._exportExcel4IE(tbid);
        } else {
            this._exportExcel(tbid, name)
        }
    }
    //----------------------以下是将JSON数据导出到EXCEL，基于js-xlsx--------------------------------
    /// <summary>
    /// 获取对应数字的字母，重叠字母
    /// </summary>
    /// <param name="n">数字</param>
    /// <returns>单个大写字母或者重叠大写字母</returns>
    static _getdblColater(n) {
        let s = '', m = 0;
        while (n > 0) {
            m = n % 26 + 1
            s = String.fromCharCode(m + 64) + s
            n = (n - m) / 26
        }
        return s
    }
    /// <summary>
    /// 生产预备数据[{v:"1111"},{v:"222"}]
    /// </summary>
    /// <param name="json">json数据[{},{},{}]</param>
    /// <param name="keyArr">键数组</param>
    /// <returns>[]</returns>
    static _dft(json, keyArr) {
        //数据工厂
        let tmpdata = [],
            r = json.map((dataobj, jsonindex) => keyArr.map((value, index) => Object.assign({}, {
                v: dataobj[value],
                position: (index > 25 ? this._getdblColater(index) : String.fromCharCode(65 + index)) + (jsonindex + 1)
                //通过Code号获取字母
            })))
        //将二维数组合并为一维数组[{position:"A1",v:"1111"},{position:"A2",v:"222"}]，
        //生产参数数据[{v:"1111"},{v:"222"}]
        r.reduce((prev, curr) => prev.concat(curr)).forEach(value => tmpdata[value.position] = {
            v: value.v
        });
        return tmpdata;
    }
    /// <summary>
    /// 字符串转字符流
    /// </summary>
    /// <param name="s">预转换字符串</param>
    /// <returns>buffer</returns>
    static _s2ab(s) {
        let buf = new ArrayBuffer(s.length),
            view = new Uint8Array(buf);
        for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
        return buf;
    }
    /// <summary>
    /// ExcelSpirit.json2excel(json,titlemap)
    /// 公共调用将json表格里面的数据导出到excel，
    /// </summary>
    /// <param name="json">json数据[{},{},{}]</param>
    /// <param name="json">titlemap字段与明文映射关系{}</param>
    /// <returns>null</returns>
    static json2excel(json, titlemap1, filename = "下载") {
        if (json.length > 0) {
            let titlemap = titlemap1 || json[0];
            //在json头部插入一个对象，用于存储表头
            json.unshift({});
            let keyArr = [];//存储key
            for (let k in titlemap) {
                keyArr.push(k);
                if (titlemap1) {
                    json[0][k] = titlemap1[k];
                } else {
                    json[0][k] = k;
                }
            }
            //获取数据工厂产生的数据
            let tmpdata = this._dft(json, keyArr),
                //设置Excel单元格区域
                outputPos = Object.keys(tmpdata),
                tmpWB = {
                    SheetNames: ['mySheet'], // 保存的表标题
                    Sheets: {
                        'mySheet': Object.assign({},
                            tmpdata, // 内容
                            {
                                '!ref': outputPos[0] + ':' + outputPos[outputPos.length - 1] // 设置填充区域
                            })
                    }
                },
                tmpDown = new Blob([this._s2ab(XLSX.write(tmpWB,
                    { bookType: 'xlsx', bookSST: false, type: 'binary' } // 这里的数据是用来定义导出的格式类型
                ))], {
                        type: ''
                    }),                                               // 创建二进制对象写入转换好的字节流
                href = URL.createObjectURL(tmpDown);                  // 创建对象超链接
            //创建虚拟a标签实现下载
            var a = document.createElement("a");
            a.download = filename + '.xlsx'  // 下载名称
            a.href = href  // 绑定a标签
            a.click()  // 模拟点击实现下载
            setTimeout(function () {  // 延时释放
                URL.revokeObjectURL(tmpDown) // 用URL.revokeObjectURL()来释放这个object URL
                a = null;
            }, 100)
        } else {
            throw new Error("Please pass in the json object");
        }
    }
    //----------------------导入Excel到返回JSON，基于js-xlsx--------------------------------
    /// <summary>
    /// ExcelSpirit.json2excel(json,titlemap)
    /// 基于一般事件的上传
    /// </summary>
    /// <param name="e">事件对象</param>
    /// <param name="callback(json)">回调函数</param>
    /// <param name="binary" default="true">是否以二进制的形式读取文件</param>
    /// <returns>null</returns>
    static excel2json4ev(callback, binary = false, type) {
        var input = document.createElement('input');   //创建虚拟file
        var inputid = "file_" + Math.ceil(Math.random() * 100);  //创建随机ID
        var body_element = document.getElementsByTagName("body")[0];
        input.setAttribute('type', 'file');
        input.setAttribute('id', inputid);
        input.setAttribute('multiple', "multiple");
        input.style.display = "none";
        input.setAttribute("onchange", "ExcelSpirit.importExcel(this," + callback + "," + binary + ",\"" + type + "\")");
        body_element.appendChild(input);
        var fileElem = document.getElementById(inputid);
            if (fileElem) {
                //触发file伪事件
                fileElem.click();
                //fileElem = null;
            }
        //进行事件伪触发
        /* var evnodeid = e.target.attributes.id.value;
        if (evnodeid) {
            var fileElem = document.getElementById(inputid);
            if (fileElem) {
                //触发file伪事件
                fileElem.click();
                fileElem = null;
            }
        } else {
            throw new Error("Please add an id for the event object!!!");
        } */

    }
    static excel2json4drop(target, callback, binary = false, type) {
        let _self = this;
        document.ondragover = function (e) {
            e.preventDefault();  //阻止浏览器默认的拖拽解析行为
            return false;
        };
        document.ondrop = function (e) {
            e.preventDefault();
            return false; //阻止 document.ondrop的默认行为 
        };
        var container = document.getElementById(target);
        container.ondragover = function (e) {
            e.preventDefault();
            return false;
        };
        container.ondrop = function (e) {
            var filelist = e.dataTransfer.files;
            _self._doimport(filelist, callback, false, type);
        };
    }
    //文件流转BinaryString
    static fixdata(data) {
        var o = "",
            l = 0,
            w = 10240;
        for (; l < data.byteLength / w; ++l) o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l * w, l * w + w)));
        o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l * w)));
        return o;
    }
    /// <summary>
    /// ExcelSpirit.json2excel(json,titlemap)
    /// 解析Excel
    /// </summary>
    /// <param name="ev">file事件对象</param>
    /// <param name="callback(json)">回调函数</param>
    /// <param name="binary" default="true">是否以二进制的形式读取文件</param>
    /// <returns>null</returns>
    static importExcel(ev, callback, binary, type) {
        if (!ev.files) {
            return;
        }
        var filelist = ev.files;
        ev.parentNode.removeChild(ev);
        this._doimport(filelist, callback, binary, type);
    }
    static _doimport(fileList, callback, binary, type) {
        let _self = this,
            JSON2lat = [];
        (function dofileList(index) {
            if (index === fileList.length) {
                if (callback) {
                    callback(JSON.stringify(JSON2lat));
                    return;
                } else {
                    JSON2lat = null;
                }
            }
            let wb;
            var f = fileList[index];
            if (f.name.endsWith(".xls") || f.name.endsWith(".xlsx")) {
                var reader = new FileReader();
                reader.readAsBinaryString(f);
                reader.onload = function (e) {
                    var data = e.target.result;
                    if (binary) {
                        wb = XLSX.read(btoa(_self.fixdata(data)), {
                            type: 'base64'
                        });
                    } else {
                        wb = XLSX.read(data, {
                            type: 'binary'
                        });
                    }
                    if (type === "obj") {
                        let r = _self._read4excel2obj(wb);
                        callback && callback(r);
                        return;
                    }
                    let sheetLength = Object.keys(wb.Sheets).length,
                        json = [];
                    (function getjson(n) {
                        if (n == sheetLength) {
                            if (fileList.length === 1) {
                                //如果一个文件则返回一维数组
                                if (callback) {
                                    callback(json);
                                    return;
                                } else {
                                    json = null;
                                }

                            } else {
                                //如果多个文件则返回二维数组，一个数组包含一个文件的数据JSON
                                JSON2lat.push(json);
                                dofileList(++index);
                            }
                        } else {
                            let currJSON = JSON.stringify(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[n]]));
                            if (currJSON.length > 2) {
                                currJSON = JSON.parse(currJSON);
                                json = json.concat(currJSON);
                            }
                            getjson(++n);
                        }
                    })(0)
                    //下面那种方式可能由于sheet过多，由于异步问题造成返回json不完全
                    /* for (let i = 0; i < sheetLength; i++) {
                        var currJSON = JSON.stringify(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[i]]));
                        if (currJSON.length > 2) {
                            json = json.concat(currJSON);
                        }
                    }
                    if (callback) {
                        callback(json);
                    } */
                }
            } else {
                throw new Error("文件格式错误！请选择后缀为.xls或.xlsx的Excel文件！");
            }
        })(0)
    }
    static _read4excel2obj(wb) {
        let result = {};
        wb.SheetNames.forEach(function (sheetName) {
            var roa = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1 });
            if (roa.length) result[sheetName] = roa;
        });
        for (let key in result) {
            let a = JSON.parse(JSON.stringify(result[key]));
            result[key].splice(0, result[key].length);
            for (let i = 0, len = a.length; i < len; i++) {
                if (a[i].length !== 0) {
                    result[key].push(a[i])
                }
            }
            a = null;

        }
        return JSON.parse(JSON.stringify(result));
    }
}
//保护类中的私有方法
class ExcelSpirit {
    static html2excel(tbid, name = "Worksheet") {
        _ExcelSpirit.html2excel(tbid, name);
    }
    static json2excel(json, titlemap1, filename = "下载") {
        _ExcelSpirit.json2excel(json, titlemap1, filename);
    }
    static excel2json4ev(callback, binary = false) {
        _ExcelSpirit.excel2json4ev(callback, binary);
    }
    static excel2obj4ev(callback, binary = false) {
        _ExcelSpirit.excel2json4ev(callback, binary, "obj");
    }
    static excel2json4drop(target, callback, binary = false) {
        _ExcelSpirit.excel2json4drop(target, callback, binary);
    }
    static excel2obj4drop(target, callback, binary = false) {
        _ExcelSpirit.excel2json4drop(target, callback, binary, "obj");
    }
    static importExcel(ev, callback, binary, type) {
        _ExcelSpirit.importExcel(ev, callback, binary, type);
    }
}
//以三种方式暴露此类
window.ExcelSpirit = ExcelSpirit;     //script引用

module.exports = ExcelSpirit;         //commonjs规范

export default ExcelSpirit;         //ES6暴露