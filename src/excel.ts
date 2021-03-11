import Xlsx from 'xlsx'

export function parseExcelToJson(excel: Uint8Array): Data {
    const data = {} as Data
    const file = Xlsx.read(excel)

    for (const sheetName of file.SheetNames) {
        const tokens = parseSheet(file.Sheets[sheetName]!!)
        parse(tokens, data)
    }

    return data
}

type Tokens = {
    type: 'h1' | 'h2' | 'array' | 'normal',
    key: string,
    value: string
}[]

type KVData = {
    [key: string]: string
}

type SheetData = {
    [key: string]: string | KVData | KVData[]
}

type Data = {
    [tabName: string]: SheetData
}

function parseSheet(sheet: Xlsx.WorkSheet): Tokens {
    const keyColumnCellNames = Object.keys(sheet).filter(it => it.startsWith('A')).sort((a, b) => {
        const aNum = Number.parseInt(a.slice(1))
        const bNum = Number.parseInt(b.slice(1))
        return aNum - bNum
    })

    let started = false
    const data: Tokens = []

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
                value
            })
        } else if (key.startsWith('\t\t')) {
            data.push({
                type: 'h2',
                key: key.slice(2),
                value
            })
        } else if (key.startsWith('\t')) {
            data.push({
                type: 'h1',
                key: key.slice(1),
                value
            })
        } else {
            data.push({
                type: 'normal',
                key,
                value
            })
        }
    }

    throw Error(`末尾に __ がありません`)
}

function parse(tokens: Tokens, data: Data = {}): Data {
    if (tokens[0]?.type !== 'h1') {
        throw Error('最初の行がシート名ではありません')
    }

    const currentSheetObj: SheetData = {}

    let currentH2: string | null = null
    let firstArrayKey: string = ''

    for (const token of tokens.slice(1)) {
        if (token.type === 'normal') {
            if (currentH2 === null) {
                currentSheetObj[token.key] = token.value
            } else {
                // {} で初期化しているため、型エラーは無視できる
                (currentSheetObj[currentH2] as KVData)[token.key] = token.value
            }
        }
        if (token.type === 'h2') {
            currentH2 = token.key
            currentSheetObj[currentH2] = {}
        }
        if (token.type === 'array') {
            //シートに小項目なしに Array がある場合
            if (currentH2 === null) {
                throw Error('配列は小項目でのみ使用できます。')
            }

            if (typeof currentSheetObj[currentH2] === 'string') {
                throw Error('小項目に配列を含むとき、その小項目では他の要素は使用できません。')
            }
            if (!Array.isArray(currentSheetObj[currentH2]) && Object.keys(currentSheetObj[currentH2] as any).length > 0) {
                throw Error('小項目に配列を含むとき、その小項目では他の要素は使用できません。')
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
            throw Error('シート名は最初の行にしか記述できません。')
        }
    }

    data[tokens[0].key] = currentSheetObj

    return data
}
