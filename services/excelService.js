const exl = require('exceljs')

async function getNationalCatalog() {

    const names = []
    const gtins = []

    const wb = new exl.Workbook()

    await wb.xlsx.readFile('./public/Краткий отчет.xlsx')

    const ws = wb.getWorksheet('Краткий отчет')

    const [c1, c2] = [ws.getColumn(1), ws.getColumn(2)]

    c1.eachCell(c => {
        gtins.push(`0${c.value}`)
    })

    c2.eachCell(c => {
        names.push(c.value)
    })

    return [names, gtins]

}

async function getMonthlyMarks() {

    const actual_gtins = []
    const actual_marks = []
    const actual_dates = []
    const actual_status = []

    const wb = new exl.Workbook()

    await wb.xlsx.readFile('./public/actual_marks.xlsx')

    const ws = wb.getWorksheet('Sheet0')

    const [c1, c2, c16, c23] = [ws.getColumn(1), ws.getColumn(2), ws.getColumn(16), ws.getColumn(23)]

    c1.eachCell(c => {
        if(c.value.indexOf('01') >= 0) {
            let str = c.value
            if(str.indexOf('<') >= 0) {
                str = str.replace(/</g, '&lt;')
            }
            actual_marks.push(str)
        }
    })

    c2.eachCell(c => {
        if(c.value !== null) {
            if(c.value.indexOf('029') >= 0) {
                actual_gtins.push(c.value)
            }
        }
    })

    c16.eachCell(c => {
        if(c.value != null && c.value != 'Статус кода') {
            actual_status.push(c.value)
        }
    })

    c23.eachCell(c => {
        if(c.value !== null) {
            if(c.value.indexOf('-') >= 0) {
                let str = c.value
                actual_dates.push(str.replace(str.substring(10), ''))
            }
        }
    })

    return [actual_gtins, actual_marks, actual_dates, actual_status]

}

module.exports = { getNationalCatalog, getMonthlyMarks }
