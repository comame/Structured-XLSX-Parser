import fs from 'fs'
import { parseExcelToJson } from './excel'

const file: Uint8Array = fs.readFileSync('./sample.xlsx')

const data = parseExcelToJson(file)

console.log(JSON.stringify(data, null, 2))
