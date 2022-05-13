const xlsx = require('xlsx')
const fs = require('fs')
const excelConfig = [
  {
    entry: './excel/活動資訊.xlsx', // 讀取excel位置
    sheetNo: 0, // 讀取excel的哪張sheet
    outputDir: './json', // 導出位置
    outputName: 'events' // 導出檔案名稱
  },
  {
    entry: './excel/活動影片.xlsx', // 讀取excel 位置
    sheetNo: 0, // 讀取excel的哪張sheet
    outputDir: './json', // 導出位置
    outputName: 'eventsVideo' // 導出檔案名稱
  },
  {
    entry: './excel/新聞稿.xlsx', // 讀取excel位置
    sheetNo: 0, // 讀取excel的哪張sheet
    outputDir: './json', // 導出位置
    outputName: 'news' // 導出檔案名稱
  },
  {
    entry: './excel/響應企業資訊.xlsx', // 讀取excel位置
    sheetNo: 0, // 讀取excel的哪張sheet
    outputDir: './json', // 導出位置
    outputName: 'plans' // 導出檔案名稱
  }
]

parseHandler(excelConfig)

// ----------- func
function parseHandler (excelConfig) {
  excelConfig.forEach((config) => {
    // 讀檔案
    const wb = xlsx.readFile(config.entry)

    // 讀檔案中的sheet轉檔案出來
    const result = parseToJson(wb, config)

    // 導出位置
    fs.writeFileSync(
      config.outputDir + '/' + config.outputName + '.json',
      JSON.stringify(result, null, 2)
    )
  })
}

// 核心邏輯
/**
 * 在這邊客製化你的excel轉出來的檔案即可
 * @param {*} sheetName sheetName
 * @param {*} sheet sheet
 * @param {*} range 獲取有值X,Y的範圍
 * @param {*} startC excel有值的最左邊
 * @param {*} startR excel有值的最上面
 * @param {*} endC excel有值的最右邊
 * @param {*} endR excel有值的最下面
 * @param {*} lengthC excel有值的左到右長度
 * @param {*} lengthR excel有值的上到下長度
 * @returns
 */
function getMultiValueArray (
  sheetName,
  sheet,
  range,
  startC,
  startR,
  endC,
  endR,
  lengthC,
  lengthR,
  config
) {
  if(config.outputName === 'xxx'){
    return []
  }

  const result = []
  // 上到下
  for (let r = 1; r <= endR; r++) {
    const newObject = {}
    // 左到右
    for (let c = startC; c <= endC; c++) {
      if (!getCellValue(sheet, range, c, 0)) continue
      newObject[getCellValue(sheet, range, c, 0)] = getCellValue(
        sheet,
        range,
        c,
        r
      )
    }
    result.push(newObject)
  }
  return result
}

function parseToJson (wb, config) {
  let data = {}
  // 獲取第一個sheet
  const sheetName = wb.SheetNames[config.sheetNo || 0]
  const sheet = wb.Sheets[sheetName]
  // 獲取第一個sheet

  // 獲取sheet資料範圍
  const range = xlsx.utils.decode_range(sheet['!ref'])
  const { c: startC, r: startR } = range.s
  const { c: endC, r: endR } = range.e
  const lengthC = endC - startC
  const lengthR = endR - startR
  // 獲取sheet資料範圍

  // 轉成自己要的格式
  data = getMultiValueArray(
    sheetName,
    sheet,
    range,
    startC,
    startR,
    endC,
    endR,
    lengthC,
    lengthR,
    config
  )
  // 轉成自己要的格式
  return data
}

// 獲取單元格資料
/**
 * 獲取excel指定單元格位置的資料
 * @param {*} sheet 哪張sheet
 * @param {*} range 可以獲取有值的範圍
 * @param {*} x 從0開始
 * @param {*} y 從0開始
 * @returns
 */
function getCellValue (sheet, range, x, y) {
  const position = xlsx.utils.encode_cell({
    c: range.s.c + x,
    r: range.s.r + y
  })
  return sheet[position] ? sheet[position].v : ''
}
// ----------- func
