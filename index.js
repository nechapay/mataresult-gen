const ExcelJS = require('exceljs')
const fs = require('fs').promises

const results = [
  'умение самостоятельно планировать альтернативные пути достижения целей, осознанно выбирать наиболее эффективные способы решения учебных и познавательных задач;',
  'умение осуществлять контроль по образцу и вносить необходимые коррективы;',
  'умение адекватно оценивать правильность или ошибочность выполнения учебной задачи, её объективную трудность и собственные возможности её решения; ',
  'умение устанавливать причинно-следственные связи; строить логические рассуждения, умозаключения (индуктивные, дедуктивные и по аналогии) и выводы; ',
  'умение создавать, применять и преобразовывать знаково-символические средства, модели и схемы для решения учебных и познавательных задач; ',
  'умение организовывать учебное сотрудничество и совместную деятельность с учителем и сверстниками: определять цели, распределять функции и роли участников, взаимодействовать и находить общие способы работы; умение работать в группе: находить общее решение и разрешать конфликты на основе согласования позиций и учёта интересов; слушать партнёра; формулировать, аргументировать и отстаивать своё мнение; ',
  'формирование учебной и общепользовательской компетентности в области использования информационно-коммуникационных технологий (ИКТ-компетентности); ',
  'формирование первоначальных представлений об идеях и о методах информатики как об универсальном языке науки и техники;',
  'развитие умения работать с учебником; с электронным приложением, обобщать и систематизировать представления об информации и способах её получения;',
  'развитие умения формулировать и удерживать учебную задачу, применять установленные правила в работе.'
]

const PROBABILITY = 50

const subjects = {
  '5кл': {
    list: [
      { name: 'Биология', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'В мире информатики', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Второй иностранный язык (немецкий)', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Второй иностранный язык (французский)', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'География', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Изобразительное искусство ', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Иностранный язык  (английский)', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'История', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Литература', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Математика', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Музыка', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Основы архитектурного рисунка', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Первые шаги в военной карьере', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Русский язык', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Труд (технология)', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Физическая культура', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] }
    ]
  },
  '6кл': {
    list: [
      { name: 'Биология', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'В мире информатики', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Второй иностранный язык (немецкий)', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Второй иностранный язык (французский)', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'География', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Изобразительное искусство', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Иностранный язык (английский)', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'История', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Литература', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Математика', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Музыка', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Основы архитектурного рисунка', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Первые шаги в военной карьере', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Русский язык', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Труд (технология)', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Физическая культура', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] }
    ]
  },
  '7кл': {
    list: [
      { name: 'Биология', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Военная топография', letters: ['А', 'Б'] },
      { name: 'Второй иностранный язык (немецкий)', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Второй иностранный язык (французский)', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'География', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Изобразительное искусство', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Иностранный язык (английский)', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Информатика', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'История', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Литература', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Математика', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Музыка', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Основы черчения', letters: ['А', 'Б'] },
      { name: 'Разговорный немецкий', letters: ['А', 'Б'] },
      { name: 'Разговорный французский', letters: ['А', 'Б'] },
      { name: 'Решение задач с параметрами', letters: ['Д'] },
      { name: 'Русский язык', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Труд (технология)', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Физика в задачах и экспериментах', letters: ['Е'] },
      { name: 'Физика', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Физическая культура', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] }
    ]
  },
  '8кл': {
    list: [
      { name: 'Биология', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Второй иностранный язык (немецкий)', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Второй иностранный язык (французский)', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'География', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Иностранный язык (английский)', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Информатика', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'История', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Литература', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Математика', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Музыка', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Обществознание', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Основы безопасности и защиты Родины', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Русский язык', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Труд (технология)', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Физика', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Физическая культура', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Химия', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] }
    ]
  },
  '9кл': {
    list: [
      { name: 'Биология', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Второй иностранный язык (немецкий)', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Второй иностранный язык (французский)', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'География', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Иностранный язык (английский)', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Информатика', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'История', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Лёгкая атлетика', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Литература', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Математика', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Обществознание', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Основы безопасности и защиты Родины', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Русский язык', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Труд (технология)', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Физика', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Физическая культура', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Химия', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] }
    ]
  },
  '10кл': {
    list: [
      { name: 'Базовая физическая подготовка', letters: ['А', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Биология', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Военное страноведение (английский язык)', letters: ['А'] },
      { name: 'География', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Глобальная география', letters: ['Б', 'В'] },
      { name: 'Индивидуальный проект', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Иностранный язык (английский)', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Информатика', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'История', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Литература', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Математика', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Обществознание', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Основы безопасности и защиты Родины', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Основы военной подготовки', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Основы генетики', letters: ['Г'] },
      { name: 'Основы фармакологии', letters: ['Г'] },
      { name: 'Основы экономики', letters: ['А'] },
      { name: 'Русский язык', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Физика', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Физическая культура', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] },
      { name: 'Химия', letters: ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж'] }
    ]
  },
  '11кл': {
    list: []
  }
}

async function createExcel() {
  const data = await readListOfStudents('списки.xlsx')
  data.forEach((item) => {
    Object.keys(item).forEach((key) => {
      createExcelFile(item, key)
    })
  })
}

async function createExcelFile(item, key) {
  const workbook = new ExcelJS.Workbook()
  const worksheet = workbook.addWorksheet(key)
  const offset = 4
  console.log(key)
  worksheet.getCell('B1').value = 'Уровень освоения планируемых метапредметных результатов {предмет}'
  worksheet.getCell('B2').value = 'ФИО преподавателя: {преподаватель}'
  worksheet.getRow(offset).values = ['П/н', 'Фамилия, имя, отчество', 'итог']

  item[key].forEach((student, index) => {
    const sheetName = createResultsSheet(workbook, index + 1, results)
    let formula = `ROUND(AVERAGE('${sheetName}'!C3:'${sheetName}'!C12)*100,0)`
    const link = '#' + sheetName + '!A1'
    worksheet.addRow([
      {
        text: index + 1,
        hyperlink: link
      },
      student,
      { formula: formula }
    ])
    let cell = worksheet.getCell(index + offset + 1, 1)
    cell.font = {
      color: { argb: 'FF00F0FF' },
      underline: true
    }
    cell.numFmt = 0
  })

  const headerRow = worksheet.getRow(offset)
  headerRow.height = 35
  headerRow.font = { bold: false, size: 11 }
  headerRow.alignment = { vertical: 'middle', horizontal: 'center' }
  worksheet.getColumn(2).width = 45
  const firstColumn = worksheet.getColumn(1)
  firstColumn.alignment = { horizontal: 'right' }
  applyGridBorders(worksheet, offset, 1, item[key].length + offset, 3, {
    outerStyle: 'thin',
    innerStyle: 'thin',
    outerColor: '000000',
    innerColor: '000000'
  })
  const defaultFont = {
    name: 'Times New Roman',
    size: 11,
    color: { argb: 'FF000000' } // черный
  }
  worksheet.eachRow((row) => {
    row.font = defaultFont
  })
  await workbook.xlsx.writeFile(key + '.xlsx')
}

function createResultsSheet(wb, name, data) {
  const worksheet = wb.addWorksheet('' + name)

  worksheet.getRow(2).values = ['П/н', 'Метапредметные результаты', 'год']

  data.forEach((item, index) => {
    let value = getRandomInt(0, 100) >= PROBABILITY ? 1 : 0

    worksheet.addRow([index + 1, item, value])
  })

  worksheet.getCell(data.length + 4, 3).value = '0-неверно'
  worksheet.getCell(data.length + 5, 3).value = '1-неверно'

  applyGridBorders(worksheet, 2, 1, data.length + 2, 3, {
    outerStyle: 'thin',
    innerStyle: 'thin',
    outerColor: '000000',
    innerColor: '000000'
  })

  worksheet.getColumn(1).width = 5
  worksheet.getColumn(2).width = 100
  worksheet.getColumn(2).alignment = { wrapText: true, vertical: 'top' }
  worksheet.getColumn(3).width = 5
  worksheet.getColumn(1).alignment = { vertical: 'middle', horizontal: 'center' }
  worksheet.getColumn(3).alignment = { vertical: 'middle', horizontal: 'center' }
  const defaultFont = {
    name: 'Times New Roman',
    size: 11,
    color: { argb: 'FF000000' } // черный
  }
  worksheet.eachRow((row) => {
    row.font = defaultFont
    const cellValue = row.getCell(2).value || ''
    const columnWidth = worksheet.getColumn(2).width || 10

    const approximateLines = Math.ceil(cellValue.toString().length / columnWidth)
    row.height = approximateLines * 14 <= 15 ? 15 : approximateLines * 14
  })
  const headerRow = worksheet.getRow(2)
  headerRow.height = 35
  headerRow.font = { bold: true, size: 11 }
  headerRow.alignment = { vertical: 'middle', horizontal: 'center' }
  worksheet.getCell(data.length + 4, 3).alignment = { vertical: 'middle', horizontal: 'left' }
  worksheet.getCell(data.length + 5, 3).alignment = { vertical: 'middle', horizontal: 'left' }
  return worksheet.name
}

async function readListOfStudents(path) {
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(path)

  const data = []

  workbook.eachSheet(function (worksheet) {
    const course = worksheet.getCell('C1').value

    const obj = {}
    obj[course] = []

    worksheet.eachRow((row) => {
      row.eachCell((cell, colNumber) => {
        if (colNumber === 1) obj[course].push(cell.value)
      })
    })

    data.push(obj)
  })
  return data
}

createExcel()

function applyGridBorders(worksheet, startRow, startCol, endRow, endCol, options = {}) {
  const { outerStyle = 'thick', innerStyle = 'thin', outerColor = '000000', innerColor = '808080' } = options

  for (let row = startRow; row <= endRow; row++) {
    for (let col = startCol; col <= endCol; col++) {
      const cell = worksheet.getCell(row, col)
      const isOuterRow = row === startRow || row === endRow
      const isOuterCol = col === startCol || col === endCol
      const style = isOuterRow || isOuterCol ? outerStyle : innerStyle
      const color = isOuterRow || isOuterCol ? outerColor : innerColor

      cell.border = {
        top: { style, color: { argb: color } },
        left: { style, color: { argb: color } },
        bottom: { style, color: { argb: color } },
        right: { style, color: { argb: color } }
      }
    }
  }
}

function getRandomInt(min, max) {
  return Math.floor(Math.random() * (max - min)) + min
}
