"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.parseExcelToJson = void 0;
var xlsx_1 = __importDefault(require("xlsx"));
function parseExcelToJson(excel, fileName) {
    var data = {};
    var file = xlsx_1.default.read(excel, { type: 'array' });
    for (var _i = 0, _a = file.SheetNames; _i < _a.length; _i++) {
        var sheetName = _a[_i];
        var tokens = tokenize(file.Sheets[sheetName], sheetName, fileName);
        parse(tokens, data, fileName);
    }
    return data;
}
exports.parseExcelToJson = parseExcelToJson;
function riseError(message, token, fileName) {
    throw Error(message + (" at " + fileName + " (" + token.sheetName + ":" + token.cell + ")"));
}
function tokenize(sheet, sheetName, fileName) {
    var _a, _b;
    var keyColumnCellNames = Object.keys(sheet).filter(function (it) { return it.startsWith('A'); }).sort(function (a, b) {
        var aNum = Number.parseInt(a.slice(1));
        var bNum = Number.parseInt(b.slice(1));
        return aNum - bNum;
    });
    var started = false;
    var data = [];
    for (var _i = 0, keyColumnCellNames_1 = keyColumnCellNames; _i < keyColumnCellNames_1.length; _i++) {
        var keyCellName = keyColumnCellNames_1[_i];
        var key = sheet[keyCellName].w;
        if (typeof key !== 'string') {
            throw Error(keyCellName + " (" + key + ") is not a string value.");
        }
        if (key === '__' && !started) {
            started = true;
            continue;
        }
        if (key === '__' && started) {
            return data;
        }
        if (key.startsWith('  '))
            continue;
        var valueCellName = 'B' + keyCellName.slice(1);
        var value = (_b = (_a = sheet[valueCellName]) === null || _a === void 0 ? void 0 : _a.w) !== null && _b !== void 0 ? _b : '';
        if (key.startsWith(' ')) {
            data.push({
                type: 'array',
                key: key.slice(1),
                value: value,
                cell: keyCellName,
                sheetName: sheetName
            });
        }
        else if (key.startsWith('\t\t')) {
            data.push({
                type: 'h2',
                key: key.slice(2),
                value: value,
                cell: keyCellName,
                sheetName: sheetName
            });
        }
        else if (key.startsWith('\t')) {
            data.push({
                type: 'h1',
                key: key.slice(1),
                value: value,
                cell: keyCellName,
                sheetName: sheetName
            });
        }
        else {
            data.push({
                type: 'normal',
                key: key,
                value: value,
                cell: keyCellName,
                sheetName: sheetName
            });
        }
    }
    riseError("\u672B\u5C3E\u306B __ \u304C\u3042\u308A\u307E\u305B\u3093", { cell: '', sheetName: sheetName }, fileName);
}
function parse(tokens, data, fileName) {
    var _a, _b;
    var _c;
    if (data === void 0) { data = {}; }
    if (((_c = tokens[0]) === null || _c === void 0 ? void 0 : _c.type) !== 'h1') {
        riseError('????????????????????????????????????????????????', tokens[0], fileName);
    }
    var currentSheetObj = {};
    var currentH2 = null;
    var firstArrayKey = '';
    for (var _i = 0, _d = tokens.slice(1); _i < _d.length; _i++) {
        var token = _d[_i];
        if (token.type === 'normal') {
            if (currentH2 === null) {
                currentSheetObj[token.key] = token.value;
            }
            else {
                // {} ???????????????????????????????????????????????????????????????
                currentSheetObj[currentH2][token.key] = token.value;
            }
        }
        if (token.type === 'h2') {
            currentH2 = token.key;
            currentSheetObj[currentH2] = {};
        }
        if (token.type === 'array') {
            //?????????????????????????????? Array ???????????????
            if (currentH2 === null) {
                riseError('????????????????????????????????????????????????', token, fileName);
            }
            if (typeof currentSheetObj[currentH2] === 'string') {
                riseError('????????????????????????????????????????????????????????????????????????????????????????????????', token, fileName);
            }
            if (!Array.isArray(currentSheetObj[currentH2]) && Object.keys(currentSheetObj[currentH2]).length > 0) {
                riseError('????????????????????????????????????????????????????????????????????????????????????????????????', token, fileName);
            }
            if (!Array.isArray(currentSheetObj[currentH2])) {
                currentSheetObj[currentH2] = [(_a = {},
                        _a[token.key] = token.value,
                        _a)];
                firstArrayKey = token.key;
            }
            else {
                var self_1 = currentSheetObj[currentH2];
                if (token.key === firstArrayKey) {
                    self_1.push((_b = {},
                        _b[token.key] = token.value,
                        _b));
                }
                else {
                    self_1[self_1.length - 1][token.key] = token.value;
                }
            }
        }
        if (token.type === 'h1') {
            riseError('????????????????????????????????????????????????????????????', token, fileName);
        }
    }
    data[tokens[0].key] = currentSheetObj;
    return data;
}
//# sourceMappingURL=excel.js.map