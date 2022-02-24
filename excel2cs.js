var xlsx = require('node-xlsx');
var fs = require('fs');
var P = require('path');
var dir = P.dirname(process.argv.splice(2)[0]);
var paths = fs.readdirSync(dir);

// console.log(sheets);

var getVal = function (type, v) {
    switch (type) {
        case 'str':
            return 'string';
        case 'num':
            return 'double';
        case 'arr':
            return 'ArrayList';
    }
};

var parseExcel = function (file) {
    let buffer = fs.readFileSync(file);
    var sheets = xlsx.parse(buffer);
    var json = new Object();
    sheets.forEach(sheet => {
        var o;
        var len = sheet.data.length;
        var cellType = new Array();
        var hasId = false;
        for (var i = 0; i < len; i++) {
            var row = sheet.data[i];
            var l = row.length;
            var data = new Object();
            for (var j = 0; j < l; j++) {
                var cell = row[j];
                if (i == 0) {
                    var a = cell.split('_');
                    cellType.push({ name: a[0], type: a[1] });
                    if (a[0] == 'id') {
                        hasId = true;
                        o = new Object();
                    }
                } else {
                    var name = cellType[j].name;
                    data[name] = getVal(cellType[j].type, cell);
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
    to[to.length - 1] = 'allJson.json';
    fs.writeFileSync(to.join('/'), JSON.stringify(json), 'utf8');
};

paths.forEach(path => {
    if (P.extname(path) == '.xlsx') {
        parseExcel(dir + '/' + path);
    }
});