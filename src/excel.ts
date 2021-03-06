import Xlsx from 'xlsx'

export function parseExcelToJson(excel: Uint8Array, fileName: string): Data {
    const data = {} as Data
    const file = Xlsx.read(excel, { type: 'array' })

    for (const sheetName of file.SheetNames) {
        const tokens = tokenize(file.Sheets[sheetName]!!, sheetName, fileName)
        parse(tokens, data, fileName)
    }

    return data
}

type Token = {
    type: 'h1' | 'h2' | 'array' | 'normal',
    key: string,
    value: string,

    cell: string,
    sheetName: string
}

type KVData = {
    [key: string]: string
}

type SheetData = {
    [key: string]: string | KVData | KVData[]
}

type Data = {
    [tabName: string]: SheetData
}

function riseError(message: string, token: Pick<Token, "cell"|"sheetName">, fileName: string): never {
    throw Error(message + ` at ${fileName} (${token.sheetName}:${token.cell})`)
}

function tokenize(sheet: Xlsx.WorkSheet, sheetName: string, fileName: string): Token[] {
    const keyColumnCellNames = Object.keys(sheet).filter(it => it.startsWith('A')).sort((a, b) => {
        const aNum = Number.parseInt(a.slice(1))
        const bNum = Number.parseInt(b.slice(1))
        return aNum - bNum
    })

    let started = false
    const data: Token[] = []

    for (const keyCellName of keyColumnCellNames) {
        const key: unknown = sheet[keyCellName].w
        if (typeof key !== 'string') {
            throw Error(`${keyCellName} (${key}) is not a string value.`)
        }
        if (key === '__' && !started) {
            started = true
            continue
        }
        if (key === '__' && started) {
            return data
        }
        if (key.startsWith('  ')) continue

        const valueCellName = 'B' + keyCellName.slice(1)
        const value: string = sheet[valueCellName]?.w ?? ''

        if (key.startsWith(' ')) {
            data.push({
                type: 'array',
                key: key.slice(1),
                value,
                cell: keyCellName,
                sheetName
            })
        } else if (key.startsWith('\t\t')) {
            data.push({
                type: 'h2',
                key: key.slice(2),
                value,
                cell: keyCellName,
                sheetName
            })
        } else if (key.startsWith('\t')) {
            data.push({
                type: 'h1',
                key: key.slice(1),
                value,
                cell: keyCellName,
                sheetName
            })
        } else {
            data.push({
                type: 'normal',
                key,
                value,
                cell: keyCellName,
                sheetName
            })
        }
    }

    riseError(`????????? __ ??????????????????`, { cell: '', sheetName }, fileName)
}

function parse(tokens: Token[], data: Data = {}, fileName: string): Data {
    if (tokens[0]?.type !== 'h1') {
        riseError('????????????????????????????????????????????????', tokens[0]!!, fileName)
    }

    const currentSheetObj: SheetData = {}

    let currentH2: string | null = null
    let firstArrayKey: string = ''

    for (const token of tokens.slice(1)) {
        if (token.type === 'normal') {
            if (currentH2 === null) {
                currentSheetObj[token.key] = token.value
            } else {
                // {} ???????????????????????????????????????????????????????????????
                (currentSheetObj[currentH2] as KVData)[token.key] = token.value
            }
        }
        if (token.type === 'h2') {
            currentH2 = token.key
            currentSheetObj[currentH2] = {}
        }
        if (token.type === 'array') {
            //?????????????????????????????? Array ???????????????
            if (currentH2 === null) {
                riseError('????????????????????????????????????????????????', token, fileName)
            }

            if (typeof currentSheetObj[currentH2] === 'string') {
                riseError('????????????????????????????????????????????????????????????????????????????????????????????????', token, fileName)
            }
            if (!Array.isArray(currentSheetObj[currentH2]) && Object.keys(currentSheetObj[currentH2] as any).length > 0) {
                riseError('????????????????????????????????????????????????????????????????????????????????????????????????', token, fileName)
            }

            if (!Array.isArray(currentSheetObj[currentH2])) {
                currentSheetObj[currentH2] = [{
                    [token.key]: token.value
                }]
                firstArrayKey = token.key
            } else {
                const self = currentSheetObj[currentH2] as KVData[]
                if (token.key === firstArrayKey) {
                    self.push({
                        [token.key]: token.value
                    })
                } else {
                    self[self.length - 1]!![token.key] = token.value
                }
            }
        }
        if (token.type === 'h1') {
            riseError('????????????????????????????????????????????????????????????', token, fileName)
        }
    }

    data[tokens[0].key] = currentSheetObj

    return data
}
