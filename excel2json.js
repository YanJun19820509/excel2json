var xlsx = require('node-xlsx');
var fs = require('fs');
var P = require('path');
var dir = process.cwd();
var paths = fs.readdirSync(dir);

// console.log(dir, paths);

var destBuffer = fs.readFileSync(dir + '/dest.json', 'utf8');
var destConfig = JSON.parse(destBuffer);
var defaultDest, fileDest;
destConfig.forEach(c => {
    if (!c.files) defaultDest = c.dest;
    else {
        fileDest = fileDest || {};
        c.files.forEach(file => {
            fileDest[file] = c.dest;
        });
    }
});

var getVal = function (type, v) {
    // console.log(type, v);
    if (v == 'undefined') return null;
    switch (type) {
        case 'str':
        case 'arr':
            return v != undefined ? String(v) : null;
        case 'float':
            return v != undefined ? String(v) : null;
        case 'num':
            return v != undefined ? Number(v) : null;
        case 'int':
            return v != undefined ? Number(v) : null;
        // case 'arr':
        //     return v ? String(v).split(',') : [];
    }
};

function match(conten, regex) {
    let r = new RegExp(regex);
    return r.test(conten);
}

var parseExcel = function (file, name) {
    if (name.indexOf('!') == 0) return;
    console.log('parseExcel', file);
    let buffer = readFileSync(file);
    var sheets = xlsx.parse(buffer);
    var json = new Object();
    sheets.forEach((sheet) => {
        console.log('parseSheet', sheet.name);
        if (sheet.name.indexOf('!') == 0) return;
        var o;
        var len = sheet.data.length;
        var names = [];
        var types = [];
        var regexs = [];
        var hasId = false;
        var need = [];
        for (var i = 1; i < len; i++) {
            var row = sheet.data[i];
            var l = i > 1 ? names.length : row.length;
            var data = {};
            for (var j = 0; j < l; j++) {
                var cell = String(row[j]).trim();
                if (i == 1) {//第一行是字段名
                    if (cell == '' || cell == null) break;
                    names.push(cell);
                    if (cell.toLowerCase() == 'id') {
                        hasId = true;
                        o = new Object();
                    }
                } else if (i == 2) {//第二行是字段属性
                    types.push(cell);
                } else if (i == 3) {//第三行是0需要导出，1不需要导出
                    need.push(cell)
                } else if (i == 4) {//第四行是正则表达式
                    regexs.push(cell);
                } else {
                    //当id为空时
                    if (j == 0 && cell == 'undefined') continue;

                    if (need[j] == 'both' || need[j] == 'client' || need[j] == '0') {
                        var nn = names[j];
                        var type = types[j];
                        var regex = regexs[j];
                        // console.log(name, type, cell);
                        data[nn] = getVal(type, cell);
                        if (nn.toLowerCase() == 'id') {
                            o[data[nn]] = data;
                        }
                        if (regex != 'undefined' && !match(String(cell), regex)) {
                            outPutInfo.push({ dir: `数据格式不匹配：文件${name} 表${sheet.name} 字段${nn} 第${i + 1}行`, name: name, state: 0 });
                        }
                    }
                }
            }
            if (!hasId && i > 3) {
                o = o || new Array();
                o.push(data);
            }
        }
        json[sheet.name.split('!')[0]] = o;
    });
    name = name.split('!')[0];
    var dest = outputDir + '/' + name + '.json';
    writeFileSync(dest, JSON.stringify(json), 'utf8');
    outPutInfo.push({ dir: dest, name: name, state: 1 });
}

paths.forEach(path => {
    // console.log(path);
    if (P.extname(path) == '.xlsx' && path.indexOf('~$') == -1) {
        parseExcel(dir + '/' + path, path.split('.xlsx')[0]);
    }
});