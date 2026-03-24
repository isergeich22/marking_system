const router = require('express').Router()
const { getNationalCatalog, getMonthlyMarks } = require('../services/excelService')
const buttons = require('../config').buttons

const STATUS_LABELS = {
    INTRODUCED: 'В обороте',
    APPLIED: 'Нанесен',
    RETIRED: 'Выбыл'
}

function buildPages(marks, gtins, names, actual_gtins, actual_dates, actual_status) {
    const items = marks.map((mark, i) => ({
        mark,
        gtin: actual_gtins[i],
        date: actual_dates[i],
        status: actual_status[i],
        statusLabel: STATUS_LABELS[actual_status[i]] || actual_status[i],
        name: names[gtins.indexOf(actual_gtins[i])] ?? '-'
    }))
    const pages = []
    for (let i = 0; i < items.length; i += 10) {
        pages.push(items.slice(i, i + 10))
    }
    return pages
}

router.get('/home', async function(req, res) {
    const pageNumber = Math.max(1, parseInt(req.query.page) || 1)
    const [names, gtins] = await getNationalCatalog()
    const [actual_gtins, actual_marks, actual_dates, actual_status] = await getMonthlyMarks()

    const pages = buildPages(actual_marks, gtins, names, actual_gtins, actual_dates, actual_status)
    const lastPage = Math.max(1, pages.length)
    const marks = pages[pageNumber - 1] || []

    res.render('home', { title: 'Главная', marks, pageNumber, lastPage, buttons })
})

router.get('/home/:status/', async function(req, res) {
    const status = req.params.status
    const pageNumber = Math.max(1, parseInt(req.query.page) || 1)
    const [names, gtins] = await getNationalCatalog()
    const [actual_gtins, actual_marks, actual_dates, actual_status] = await getMonthlyMarks()

    const allItems = actual_marks
        .map((mark, i) => ({
            mark,
            gtin: actual_gtins[i],
            date: actual_dates[i],
            status: actual_status[i],
            statusLabel: STATUS_LABELS[actual_status[i]] || actual_status[i],
            name: names[gtins.indexOf(actual_gtins[i])] ?? '-'
        }))
        .filter(item => item.status === status)

    const pages = []
    for (let i = 0; i < allItems.length; i += 10) {
        pages.push(allItems.slice(i, i + 10))
    }
    const lastPage = Math.max(1, pages.length)
    const marks = pages[pageNumber - 1] || []

    res.render('home', { title: 'Главная', marks, pageNumber, lastPage, status, buttons })
})

router.get('/filter', async function(req, res) {
    const [names, gtins] = await getNationalCatalog()
    const [actual_gtins, actual_marks, actual_dates, actual_status] = await getMonthlyMarks()

    let marks = []

    if (req.query.cis && !req.query.gtin) {
        const index = actual_marks.findIndex(m => m.indexOf(req.query.cis) >= 0)
        if (index >= 0) {
            marks = [{
                mark: actual_marks[index],
                gtin: actual_gtins[index],
                date: actual_dates[index],
                statusLabel: STATUS_LABELS[actual_status[index]] || actual_status[index],
                name: names[gtins.indexOf(actual_gtins[index])] ?? '-'
            }]
        }
    }

    if (req.query.gtin && !req.query.cis) {
        marks = actual_marks
            .map((mark, i) => ({ mark, gtin: actual_gtins[i], date: actual_dates[i], status: actual_status[i] }))
            .filter(item => item.gtin === req.query.gtin)
            .map(item => ({
                ...item,
                statusLabel: STATUS_LABELS[item.status] || item.status,
                name: names[gtins.indexOf(item.gtin)] ?? '-'
            }))
    }

    res.render('home', { title: 'Главная', marks, buttons })
})

module.exports = router
