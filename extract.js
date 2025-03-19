const fs = require('fs')
const path = require('path')
const glob = require('glob')
const ExcelJS = require('exceljs')

//  한글 추출 (유니코드 \uAC00-\uD7A3 범위)
const koreanRegex = /[\uAC00-\uD7A3]+/g

// 프로젝트 내 모든 .js, .jsx, .ts, .tsx 파일을 찾기
const searchPattern = '**/*.{js,jsx,ts,tsx,java}'

const projectRoot = process.argv[2]
  ? path.resolve(process.argv[2])
  : process.cwd()

const projectName = path.basename(projectRoot)
const outputFileName = `${projectName}.xlsx`

const toExcel = async (data) => {
  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet('sheet1')

  ws.columns = [
    { header: 'filename', key: 'filename' },
    { header: 'line', key: 'line' },
    { header: 'message', key: 'message' },
  ]

  ws.addRows(data)

  try {
    await wb.xlsx.writeFile(outputFileName)
    console.log('✅ Excel file created successfully.')
  } catch (err) {
    console.error('❌ Error creating Excel file:', err)
  }
}

const extractKoreanText = (text) => {
  const matches = text.match(/[`'"]([^`'"]+)[`'"]/) // 백틱, 작은 따옴표, 큰 따옴표 포함
  if (matches) {
    return matches[1].match(koreanRegex)?.join(' ') || ''
  }
  return text.match(koreanRegex)?.join('') || ''
}

const extractKoreanLines = (content, filePath) => {
  return content
    .split('\n')
    .map((line, index) => {
      const message = extractKoreanText(line)
      return message
        ? {
            filename: filePath.replace(projectRoot, ''),
            line: index + 1,
            message,
          }
        : null
    })
    .filter(Boolean) // `null` 값 제거
}

const extractKoreanFromFile = async (filePath) => {
  try {
    const content = await fs.promises.readFile(filePath, 'utf8')
    return extractKoreanLines(content, filePath)
  } catch (error) {
    console.error(`❌ Error reading file: ${filePath}`, error)
    return []
  }
}

const run = async () => {
  console.log(`📂 Project Root: ${projectRoot}`)

  const files = glob.sync(searchPattern, {
    cwd: projectRoot,
    ignore: ['node_modules/**'],
    absolute: true,
  })

  console.log(`🔍 검색된 파일 수: ${files.length}`)

  const results = await Promise.all(files.map(extractKoreanFromFile))

  const ws_data = results.flat()

  console.log(`✅ 추출된 한글 데이터 개수: ${ws_data.length}`)

  if (ws_data.length > 0) {
    await toExcel(ws_data)
  } else {
    console.log('⚠️ 추출된 데이터가 없습니다.')
  }
}

run()
