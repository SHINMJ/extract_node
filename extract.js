const fs = require('fs')
const path = require('path')
const glob = require('glob')
const ExcelJS = require('exceljs')

// 한글 포함 여부 확인 (유니코드 \uAC00-\uD7A3 범위)
const koreanRegex = /[\uAC00-\uD7A3]+/

// 프로젝트 내 모든 .js, .jsx, .ts, .tsx, .java 파일 찾기
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
    console.log(`✅ Excel file created: ${outputFileName}`)
  } catch (err) {
    console.error('❌ Error creating Excel file:', err)
  }
}

// 주석을 제거하는 함수
const removeComments = (content) => {
  // 한 줄 주석 (//)과 멀티라인 주석 (/* */)을 제거
  content = content.replace(/\/\/.*$/gm, '') // 한 줄 주석 제거
  content = content.replace(/\/\*[\s\S]*?\*\//g, '') // 멀티라인 주석 제거
  return content
}

// 한글 포함 여부가 있는 라인 추출 함수
const extractKoreanLines = (content, filePath) => {
  const cleanedContent = removeComments(content) // 주석 제거
  const lines = cleanedContent.split('\n')
  const filteredLines = []

  lines.forEach((line, index) => {
    if (koreanRegex.test(line)) {
      // 주석을 제외한 라인에서 한글을 찾음
      filteredLines.push({
        filename: filePath.replace(projectRoot, ''),
        line: index + 1,
        message: line.trim(), // 전체 라인 저장
      })
    }
  })

  return filteredLines
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
    ignore: ['node_modules/**', '**/test/**'],
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
