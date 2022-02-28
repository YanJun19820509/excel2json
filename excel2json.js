var xlsx = require('node-xlsx');
var fs = require('fs');
var P = require('path');
var dir = process.cwd();
var paths = fs.readdirSync(dir);

// console.log(dir, paths);

var getVal = function (type, v) {
    // console.log(type, v);
    switch (type) {
        case 'str':
            return String(v || '');
        case 'num':
            return v != undefined ? Number(v) : null;
        case 'arr':
            return v ? String(v).split(',') : [];
    }
};

var parseExcel = function (file, name) {
    let buffer = fs.readFileSync(file);
    var sheets = xlsx.parse(buffer);
    var json = new Object();
    sheets.forEach(sheet => {
        var o;
        var len = sheet.data.length;
        var names = [];
        var types = [];
        var hasId = false;
        for (var i = 0; i < len; i++) {
            var row = sheet.data[i];
            var l = i > 0 ? names.length : row.length;
            var data = new Object();
            for (var j = 0; j < l; j++) {
                var cell = row[j];
                if (i == 0) {//第一行是字段名
                    names.push(cell);
                    if (cell == 'id') {
                        hasId = true;
                        o = new Object();
                    }
                } else if (i == 1) {//第二行是字段属性
                    types.push(cell);
                } else {
                    var name = names[j];
                    var type = types[j];
                    console.log(name, type, cell);
                    data[name] = getVal(type, cell);
                    if (name == 'id') {
                        o[data[name]] = data;
                    }
                }
            }
            if (!hasId && i > 0) {
                o = o || new Array();
                o.push(data);
            }
        }
        json[sheet.name] = o;
    });

    var to = file.split('/');
    to[to.length - 1] = name + '.json';
    var dest = to.join('/');
    fs.writeFileSync(dest, JSON.stringify(json), 'utf8');
    console.log('创建：', dest);
}

paths.forEach(path => {
    // console.log(path);
    if (P.extname(path) == '.xlsx' && path.indexOf('~$') == -1) {
        parseExcel(dir + '/' + path, path.split('.xlsx')[0]);
    }
});