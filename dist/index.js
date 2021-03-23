"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const fs_1 = __importDefault(require("fs"));
const excel_1 = require("./excel");
const file = fs_1.default.readFileSync('./sample.xlsx');
const data = excel_1.parseExcelToJson(file);
console.log(JSON.stringify(data, null, 2));
//# sourceMappingURL=index.js.map