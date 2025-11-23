const express = require('express')
const exl = require('exceljs')
const cio = require('cheerio')
const QuickChart = require('quickchart-js')
const fs = require('fs')
const fetch = require('node-fetch')
const axios = require('axios')
const dotenv = require('dotenv')
const { headerComponent, navComponent, footerComponent } = require('./components/htmlComponents')
const app = express()

// Переменная для формирования html-разметки ответа
let html = ``

dotenv.config({path:__dirname + '/.env'})

const dbsId = process.env.YANDEX_DBS_ID
const fbsId = process.env.YANDEX_FBS_ID

function findMatchesByPostingNumber(arr1, arr2) {
  // создаём множество всех posting_number из второго массива
  const set2 = new Set(arr2.map(item => item.posting_number));

  // возвращаем только те элементы из первого массива, чей posting_number есть во втором
  return arr1.filter(item => set2.has(item.posting_number));
}

function compareStrings(str1, str2) {

    if (str1.length !== str2.length) {
        console.log('Строки разной длины!');
        return;
    }

    for (let i = 0; i < str1.length; i++) {
        if (str1[i] !== str2[i]) {
        console.log(`❌ Различие на позиции ${i}: '${str1[i]}' (код ${str1.charCodeAt(i)}) vs '${str2[i]}' (код ${str2.charCodeAt(i)})`);
        }
    }
    console.log('✅ Если различий нет выше — строки идентичны по символам.');

}

async function renderImportButtons(array) {

    for(let i = 0; i < array.length; i++) {                
        
        if(array[i] === 'stocks') {

            html += `<button class="button-import">
                        <a href="http://localhost:3030/stocks" target="_blank">Создать импорт для остатков</a>
                    </button>`

        }

        if(array[i] === 'wb') {

            html += `<button class="button-import">
                        <a href="http://localhost:3030/wildberries" target="_blank">Создать импорт для ${array[i]}</a>
                    </button>`

        }

        if(array[i] !== 'wb' && array[i] !== 'stocks') {
            html += `<button class="button-import">
                        <a href="http://localhost:3030/${array[i]}" target="_blank">Создать импорт для ${array[i]}</a>
                    </button>`
        }
        
    }

    html += `   </div>`

}

async function renderMarkingButtons() {
    html += `<div class="marking-control">
                <button class="marking-button remarking-button"><a href="http://localhost:3030/input_remarking" target="_blank">Ввод в оборот (Перемаркировка)</a></button>
                <button class="marking-button distance-button"><a href="http://localhost:3030/sale_ozon" target="_blank">Вывод из оборота (Дистанционная продажа)</a></button>
                <button class="marking-button distance-button"><a href="http://localhost:3030/sale_wb" target="_blank">Вывод из оборота (Дистанционная продажа)</a></button>
                <button class="marking-button distance-button"><a href="http://localhost:3030/wildberries/set_marks" target="_blank">Подстановка маркировки (Wildberries)</a></button>
            </div>`
}

async function renderExtraButtons() {

    html += `<div class="other-control">
                <button class="other-button mark-stocks"><a href="http://localhost:3030/test_features" target="_blank">Обновить краткий отчет</a></button>
                <button class="other-button mark-stocks"><a href="http://localhost:3030/personal_orders" target="_blank">Создать персональный заказ</a></button>
            </div>`

}

let buttons = ['ozon', 'wb', 'yandex', 'stocks']

app.use(express.static(__dirname + '/public'))

app.get('/home', async function(req, res){

    html = `${headerComponent}
                    <title>Главная</title>
                </head>
                <body>
                    ${navComponent}
                        <section class="sub-nav import-main">
                            <div class="import-control">`

    await renderImportButtons(buttons)
    await renderMarkingButtons()
    await renderExtraButtons()

    html += `</section>`

    html += `<section class="filter-control">
                <div class="search-field">
                    <input class="search-input" type="text" placeholder="Код или GTIN товара">
                    <button id="search" type="submit">
                        <svg width="20" height="20" fill="none" xmlns="http://www.w3.org/2000/svg" cursor="default" style="color: rgb(122, 129, 155);"><path fill-rule="evenodd" clip-rule="evenodd" d="M10.75 1.739a.75.75 0 00-1.5 0V9.25H1.739a.75.75 0 100 1.5H9.25V18.261H10h-.75a.75.75 0 101.5 0H10h.75V10.75H18.261V10v.75a.75.75 0 000-1.5V10v-.75H10.75V1.739z" fill="currentColor">
                        </path></svg>
                    </button>
                </div>
                <div class="multiple-list">
                    <div class="multiple-status">
                        Статус
                    </div>
                    <div class="status-list">
                        <ul class="list">
                            <li class="list-item">Нанесен</li>
                            <li class="list-item">В обороте</li>
                            <li class="list-item">Выбыл</li>
                        </ul>
                    </div>
                    <svg width="16" height="16" fill="none" xmlns="http://www.w3.org/2000/svg" class="MuiSelect-icon MuiSelect-iconStandard css-1rb0eps"><path d="M12 6H4l4 4 4-4z" fill="currentColor">
                    </path></svg>
                </div>
                <button class="show-button"><a id="show-anchor">Показать</a></button>
             </section>`

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

        const ws = wb.getWorksheet('Worksheet')

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
            if(c.value != null && c.value != 'status') {
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

    async function renderMarksTable() {
        
        const [names, gtins] = await getNationalCatalog()
        const [actual_gtins, actual_marks, actual_dates, actual_status] = await getMonthlyMarks()

        async function createPages(array) {

            let marks_list = []
            let _temp = []

            for(let i = 0; i < array.length; i++) {

                _temp.push({
                    gtin: actual_gtins[i],
                    mark: array[i],
                    date: actual_dates[i],
                    status: actual_status[i],
                    order: ''
                })

                if(_temp.length%10 === 0) {
                    marks_list.push(_temp)
                    _temp = []
                }

            }

            marks_list.push(_temp)
            _temp = []

            return marks_list

        }

        let pageNumber = 0

        if(req.query.page == null || req.query.page == undefined || req.query.page == 0) {

            pageNumber = 1

        } else {

            pageNumber = parseInt(req.query.page)

        }

        // if(req.query.order == null || req.query.order == undefined || req.query.order == 0) {
            
        // } else {

        //     let page = 0
        //     let index = Pages[page].findIndex(el => el.mark == req.query.mark)
            
        //     Pages[page]

        // }

        let Pages = await createPages(actual_marks)

        html += `<section class="table">
                            <div class="marks-table">
                                <div class="marks-table-header">
                                    <div class="header-cell">КИЗ</div>
                                    <div class="header-cell">GTIN</div>
                                    <div class="header-cell">Товар</div>
                                    <div class="header-cell">Дата эмиссии</div>
                                    <div class="header-cell">Статус</div>
                                    <!--<div class="header-cell">Номер заказа</div>-->
                                </div>
                                <div class="header-wrapper"></div>`

        for(let j = 0; j < Pages[pageNumber - 1].length; j++) {

                let status = ''
                if(Pages[pageNumber - 1][j].status == 'INTRODUCED') {
                    status = 'В обороте'
                } else if(Pages[pageNumber - 1][j].status == 'APPLIED') {
                    status = 'Нанесен'
                } else if(Pages[pageNumber - 1][j].status == 'RETIRED') {
                    status = 'Выбыл'
                }
                    
                html+= `<div class="table-row">
                            <span type="text" id="mark">${Pages[pageNumber - 1][j].mark}</span>
                            <span id="gtin">${Pages[pageNumber - 1][j].gtin}</span>
                            <span id="name">${names[gtins.indexOf(Pages[pageNumber - 1][j].gtin)] == undefined ? '-' : names[gtins.indexOf(Pages[pageNumber - 1][j].gtin)]}</span>
                            <span id="date">${Pages[pageNumber - 1][j].date}</span>
                            <span id="status">${status}</span>
                            <!--<div>
                                <input id="order" type="text" placeholder="${Pages[pageNumber - 1][j].order}">
                                <button type="submit"><a class="order-number" href="">Отправить</a></button>
                            </div>-->
                        </div>`
                
        }
        
        return Math.ceil(Pages.length)
    
    }


    let lastPage = await renderMarksTable()

    html += `       </div>
                <div class="pages">
                    <a id="begin" href="http://localhost:3030/home">На первую страницу</a>
                    <div class="pages-prev">
                        <svg id="prev-icon" width="6" height="10" viewBox="0 0 6 10" xmlns="http://www.w3.org/2000/svg" style=""><path fill-rule="evenodd" clip-rule="evenodd" d="M4.113 9.669c.432.441 1.13.441 1.563 0a1.145 1.145 0 0 0 0-1.596L2.668 4.999l3.008-3.072a1.145 1.145 0 0 0 0-1.596 1.087 1.087 0 0 0-1.563 0l-3.79 3.87A1.14 1.14 0 0 0 0 5c0 .29.108.578.324.799l3.79 3.87z">
                        </path></svg>
                        <a id="prev" href="">Предыдущая страница</a>
                    </div>
                    <div class="pages-next">
                        <a id="next" href="">Следующая страница</a>
                        <svg id="next-icon" width="6" height="10" viewBox="0 0 6 10" xmlns="http://www.w3.org/2000/svg" style=""><path fill-rule="evenodd" clip-rule="evenodd" d="M1.887.331a1.087 1.087 0 0 0-1.563 0 1.145 1.145 0 0 0 0 1.596l3.008 3.074L.324 8.073a1.145 1.145 0 0 0 0 1.596c.432.441 1.13.441 1.563 0l3.79-3.87A1.14 1.14 0 0 0 6 5c0-.29-.108-.578-.324-.799L1.886.332z">
                        </path></svg>
                    </div>
                    <a id="last" class="pages-last" href="http://localhost:3030/home?page=${lastPage}">На последнюю страницу</a>                  
                </div>
            </section>
        <div class="body-wrapper"></div>
    ${footerComponent}`

    res.send(html)
})

app.get('/home/:status/', async function(req, res){

    html = `${headerComponent}
                    <title>Главная</title>
                </head>
                <body>
                    ${navComponent}
                        <section class="sub-nav import-main">
                            <div class="import-control">`

    await renderImportButtons(buttons)
    await renderMarkingButtons()
    await renderExtraButtons()

    html += `</section>`

    html += `<section class="filter-control">
                <div class="search-field">
                    <input class="search-input" type="text" placeholder="Код или GTIN товара">
                    <button id="search" type="submit">
                        <svg width="20" height="20" fill="none" xmlns="http://www.w3.org/2000/svg" cursor="default" style="color: rgb(122, 129, 155);"><path fill-rule="evenodd" clip-rule="evenodd" d="M10.75 1.739a.75.75 0 00-1.5 0V9.25H1.739a.75.75 0 100 1.5H9.25V18.261H10h-.75a.75.75 0 101.5 0H10h.75V10.75H18.261V10v.75a.75.75 0 000-1.5V10v-.75H10.75V1.739z" fill="currentColor">
                        </path></svg>
                    </button>
                </div>
                <div class="multiple-list">
                    <div class="multiple-status">
                        Статус
                    </div>
                    <div class="status-list">
                        <ul class="list">
                            <li class="list-item">Нанесен</li>
                            <li class="list-item">В обороте</li>
                            <li class="list-item">Выбыл</li>
                        </ul>
                    </div>
                    <svg width="16" height="16" fill="none" xmlns="http://www.w3.org/2000/svg" class="MuiSelect-icon MuiSelect-iconStandard css-1rb0eps"><path d="M12 6H4l4 4 4-4z" fill="currentColor">
                    </path></svg>
                </div>
                <button class="show-button"><a id="show-anchor">Показать</a></button>
             </section>`

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

        const ws = wb.getWorksheet('Worksheet')

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
            if(c.value != null && c.value != 'status') {
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

    async function renderMarksTable() {

        const [names, gtins] = await getNationalCatalog()
        const [actual_gtins, actual_marks, actual_dates, actual_status] = await getMonthlyMarks()

        async function createFilterPages(array, status) {

            let marks_list = []
    
            let _temp = []
    
            if(status == 'APPLIED') {
    
                for(let i = 0; i < array.length; i++) {
    
                    if(status == actual_status[i]) {
    
                        _temp.push({
                            gtin: actual_gtins[i],
                            mark: array[i],
                            date: actual_dates[i],
                            status: actual_status[i],
                            order: ''
                        })
    
                        if(_temp.length%10 === 0) {
                            marks_list.push(_temp)
                            _temp = []
                        }
    
                    }
    
                }
    
                marks_list.push(_temp)
                _temp = []
    
            }
    
            if(status == 'RETIRED') {
    
                for(let i = 0; i < array.length; i++) {
    
                    if(status == actual_status[i]) {
    
                        _temp.push({
                            gtin: actual_gtins[i],
                            mark: array[i],
                            date: actual_dates[i],
                            status: actual_status[i],
                            order: ''
                        })
    
                        if(_temp.length%10 === 0) {
                            marks_list.push(_temp)
                            _temp = []
                        }
    
                    }
    
                }
    
                marks_list.push(_temp)
                _temp = []
            }
    
            if(status == 'INTRODUCED') {
    
                for(let i = 0; i < array.length; i++) {
    
                    if(status == actual_status[i]) {
    
                        _temp.push({
                            gtin: actual_gtins[i],
                            mark: array[i],
                            date: actual_dates[i],
                            status: actual_status[i],
                            order: ''
                        })
    
                        if(_temp.length%10 === 0) {
                            marks_list.push(_temp)
                            _temp = []
                        }
    
                    }
    
                }
    
                marks_list.push(_temp)
                _temp = []
    
            }
    
            return marks_list
    
        }

        let pageNumber = 0

        if(req.query.page == null || req.query.page == undefined || req.query.page == 0) {

            pageNumber = 1

        } else {

            pageNumber = parseInt(req.query.page)

        }

        // if(req.query.order == null || req.query.order == undefined || req.query.order == 0) {
            
        // } else {

        //     let page = 0
        //     let index = Pages[page].findIndex(el => el.mark == req.query.mark)
            
        //     Pages[page]

        // }

        let Pages = await createFilterPages(actual_marks, req.params.status)

        html += `<section class="table">
                            <div class="marks-table">
                                <div class="marks-table-header">
                                    <div class="header-cell">КИЗ</div>
                                    <div class="header-cell">GTIN</div>
                                    <div class="header-cell">Товар</div>
                                    <div class="header-cell">Дата эмиссии</div>
                                    <div class="header-cell">Статус</div>
                                    <!--<div class="header-cell">Номер заказа</div>-->
                                </div>
                                <div class="header-wrapper"></div>`

        for(let j = 0; j < Pages[pageNumber - 1].length; j++) {

                let status = ''
                if(Pages[pageNumber - 1][j].status == 'INTRODUCED') {
                    status = 'В обороте'
                } else if(Pages[pageNumber - 1][j].status == 'APPLIED') {
                    status = 'Нанесен'
                } else if(Pages[pageNumber - 1][j].status == 'RETIRED') {
                    status = 'Выбыл'
                }
                    
                html+= `<div class="table-row">
                            <span type="text" id="mark">${Pages[pageNumber - 1][j].mark}</span>
                            <span id="gtin">${Pages[pageNumber - 1][j].gtin}</span>
                            <span id="name">${names[gtins.indexOf(Pages[pageNumber - 1][j].gtin)] == undefined ? '-' : names[gtins.indexOf(Pages[pageNumber - 1][j].gtin)]}</span>
                            <span id="date">${Pages[pageNumber - 1][j].date}</span>
                            <span id="status">${status}</span>
                            <!--<div>
                                <input id="order" type="text" placeholder="${Pages[pageNumber - 1][j].order}">
                                <button type="submit"><a class="order-number" href="">Отправить</a></button>
                            </div>-->
                        </div>`
                
            }
        
        return Math.ceil(Pages.length)

    }

    let lastPage = await renderMarksTable()

    html += `       </div>
                <div class="pages">
                    <a id="begin" href="http://localhost:3030/home/${req.params.status}">На первую страницу</a>
                    <div class="pages-prev">
                        <svg id="prev-icon" width="6" height="10" viewBox="0 0 6 10" xmlns="http://www.w3.org/2000/svg" style=""><path fill-rule="evenodd" clip-rule="evenodd" d="M4.113 9.669c.432.441 1.13.441 1.563 0a1.145 1.145 0 0 0 0-1.596L2.668 4.999l3.008-3.072a1.145 1.145 0 0 0 0-1.596 1.087 1.087 0 0 0-1.563 0l-3.79 3.87A1.14 1.14 0 0 0 0 5c0 .29.108.578.324.799l3.79 3.87z">
                        </path></svg>
                        <a id="prev" href="">Предыдущая страница</a>
                    </div>
                    <div class="pages-next">
                        <a id="next" href="">Следующая страница</a>
                        <svg id="next-icon" width="6" height="10" viewBox="0 0 6 10" xmlns="http://www.w3.org/2000/svg" style=""><path fill-rule="evenodd" clip-rule="evenodd" d="M1.887.331a1.087 1.087 0 0 0-1.563 0 1.145 1.145 0 0 0 0 1.596l3.008 3.074L.324 8.073a1.145 1.145 0 0 0 0 1.596c.432.441 1.13.441 1.563 0l3.79-3.87A1.14 1.14 0 0 0 6 5c0-.29-.108-.578-.324-.799L1.886.332z">
                        </path></svg>
                    </div>
                    <a id="last" class="pages-last" href="http://localhost:3030/home/${req.params.status}?page=${lastPage}">На последнюю страницу</a>                  
                </div>
            </section>
        <div class="body-wrapper"></div>
    ${footerComponent}`

    res.send(html)

})

app.get('/filter', async function(req, res) {
    html = `${headerComponent}
                    <title>Главная</title>
                </head>
                <body>
                    ${navComponent}
                        <section class="sub-nav import-main">
                            <div class="import-control">`

    await renderImportButtons(buttons)
    await renderMarkingButtons()
    await renderExtraButtons()

    html += `</section>`

    html += `<section class="filter-control">
                <div class="search-field">
                    <input class="search-input" type="text" placeholder="Код или GTIN товара">
                    <button id="search" type="submit">
                        <svg width="20" height="20" fill="none" xmlns="http://www.w3.org/2000/svg" cursor="default" style="color: rgb(122, 129, 155);"><path fill-rule="evenodd" clip-rule="evenodd" d="M10.75 1.739a.75.75 0 00-1.5 0V9.25H1.739a.75.75 0 100 1.5H9.25V18.261H10h-.75a.75.75 0 101.5 0H10h.75V10.75H18.261V10v.75a.75.75 0 000-1.5V10v-.75H10.75V1.739z" fill="currentColor">
                        </path></svg>
                    </button>
                </div>
                <div class="multiple-list">
                    <div class="multiple-status">
                        Статус
                    </div>
                    <div class="status-list">
                        <ul class="list">
                            <li class="list-item">Нанесен</li>
                            <li class="list-item">В обороте</li>
                            <li class="list-item">Выбыл</li>
                        </ul>
                    </div>
                    <svg width="16" height="16" fill="none" xmlns="http://www.w3.org/2000/svg" class="MuiSelect-icon MuiSelect-iconStandard css-1rb0eps"><path d="M12 6H4l4 4 4-4z" fill="currentColor">
                    </path></svg>
                </div>
                <button class="show-button"><a id="show-anchor">Показать</a></button>
             </section>`

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

        const ws = wb.getWorksheet('Worksheet')

        const [c1, c2, c16, c23] = [ws.getColumn(1), ws.getColumn(2), ws.getColumn(16), ws.getColumn(23)]

        c1.eachCell(c => {
            if(c.value.indexOf('01') >= 0) {
                let str = c.value
                if(str.indexOf('&') >= 0) {
                    str = str.replace(/&/g, '&amp;')
                }
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
            if(c.value != null && c.value != 'status') {
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

    async function renderMarksTable() {

        const [names, gtins] = await getNationalCatalog()
        const [actual_gtins, actual_marks, actual_dates, actual_status] = await getMonthlyMarks()

        html += `<section class="table">
                            <div class="marks-table">
                                <div class="marks-table-header">
                                    <div class="header-cell">КИЗ</div>
                                    <div class="header-cell">GTIN</div>
                                    <div class="header-cell">Товар</div>
                                    <div class="header-cell">Дата эмиссии</div>
                                    <div class="header-cell">Статус</div>
                                    <!--<div class="header-cell">Номер заказа</div>-->
                                </div>
                                <div class="header-wrapper"></div>`

        if(req.query.cis != '' && req.query.gtin == undefined) {

            // console.log(req.query.cis)
            // let str = req.query.cis.replace(/</g, '&lt;')
            // console.log(str)
            // console.log(actual_marks[8])

            let index = 0
        
            for(let i = 0; i < actual_marks.length; i++) {

                if(actual_marks[i].indexOf(req.query.cis) >= 0) {

                    index = i

                }

            }
        
            let status = ''
                    if(actual_status[index] == 'INTRODUCED') {
                        status = 'В обороте'
                    } else if(actual_status[index] == 'APPLIED') {
                        status = 'Нанесен'
                    } else if(actual_status[index] == 'RETIRED') {
                        status = 'Выбыл'
                    }
                        
                    html+= `<div class="table-row">
                                <span type="text" id="mark">${actual_marks[index]}</span>
                                <span id="gtin">${actual_gtins[index]}</span>
                                <span id="name">${names[gtins.indexOf(actual_gtins[index])]}</span>
                                <span id="date">${actual_dates[index]}</span>
                                <span id="status">${status}</span>
                                <!--<div>
                                    <input id="order" type="text" placeholder="">
                                    <button type="submit"><a class="order-number" href="">Отправить</a></button>
                                </div>-->
                            </div>`
        }

        if(req.query.gtin != '' && req.query.cis == undefined) {

            for(let i = 0; i < actual_marks.length; i++) {

                if(actual_gtins[i] == req.query.gtin) {

                    let status = ''
                    if(actual_status[i] == 'INTRODUCED') {
                        status = 'В обороте'
                    } else if(actual_status[i] == 'APPLIED') {
                        status = 'Нанесен'
                    } else if(actual_status[i] == 'RETIRED') {
                        status = 'Выбыл'
                    }
                        
                    html+= `<div class="table-row">
                                <span type="text" id="mark">${actual_marks[i]}</span>
                                <span id="gtin">${actual_gtins[i]}</span>
                                <span id="name">${names[gtins.indexOf(actual_gtins[i])]}</span>
                                <span id="date">${actual_dates[i]}</span>
                                <span id="status">${status}</span>
                                <!--<div>
                                    <input id="order" type="text" placeholder="">
                                    <button type="submit"><a class="order-number" href="">Отправить</a></button>
                                </div>-->
                            </div>`

                }

            }

        }

    }

    await renderMarksTable()

    html += `</section>
        <div class="body-wrapper"></div>
    ${footerComponent}`

    res.send(html)

})

app.get('/ozon', async function(req, res){

    const nat_cat = []
    const nat_catGtins = []
    const nat_catNames = []
    let oz_orders = []
    const new_items = []
    const current_items = []
    const names = []

    html = `${headerComponent}
                    <title>Импорт - OZON</title>
                </head>
                <body>
                        ${navComponent}
                        <section class="sub-nav import-main">
                            <div class="import-control">`

    await renderImportButtons(buttons)
    await renderMarkingButtons()
    await renderExtraButtons()

    html += `</section>`

    const wb = new exl.Workbook()

    await wb.xlsx.readFile('./public/Краткий отчет.xlsx')

    const nc_ws = wb.getWorksheet('Краткий отчет')

    const nc_c1 = nc_ws.getColumn(1)

    nc_c1.eachCell({includeEmpty: false}, (c, rowNumber) => {

        if(rowNumber < 5) return
        nat_catGtins.push(c.value)

    })

    const nc_c2 = nc_ws.getColumn(2)

    nc_c2.eachCell({includeEmpty: false}, (c, rowNumber) => {

        if(rowNumber < 5) return
        nat_catNames.push(c.value.trim())

    })

    for(let i = 0; i < nat_catNames.length; i++) {

        nat_cat.push({
            'gtin': nat_catGtins[i],
            'name': nat_catNames[i]
        })

    }

    let response = await axios.post('https://api-seller.ozon.ru/v3/posting/fbs/list', {
        
        'dir': 'asc',
        'filter': {
            'since':`${new Date().getFullYear()}-01-01T00:00:00.000Z`,
            'status':'awaiting_packaging',
            'to':'2025-12-31T23:59:59.000Z'
        },
        'limit': 1000,
        'offset':0

    }, {

        headers: {
            'Host':'api-seller.ozon.ru',
            'Client-Id':`${process.env.OZON_CLIENT_ID}`,
            'Api-Key':`${process.env.OZON_API_KEY}`,
            'Content-Type':'application/json'
        }

    })

    let result = response.data.result.postings

    result.forEach(el => {

        for(let i = 0; i < el.products.length; i++) {

            // console.log(el.products[i].offer_id)
            if(oz_orders.findIndex(o => o.vendor === el.products[i].offer_id) >= 0) {

                oz_orders.find(o => o.vendor === el.products[i].offer_id).quantity += Number(el.products[i].quantity)

            }

            // console.log(oz_orders.findIndex(o => o.vendor === el.products[i].offer_id))

            if(oz_orders.findIndex(o => o.vendor === el.products[i].offer_id) < 0) {

                oz_orders.push({
                    'name': el.products[i].name,
                    'vendor': el.products[i].offer_id,
                    'quantity': Number(el.products[i].quantity)
                })

            }

        }
    })

    oz_orders = oz_orders.filter(o => o.name.indexOf('Одеяло') < 0 && o.name.indexOf('Подушка') < 0 && o.name.indexOf('Матрас') < 0 && o.name.indexOf('Ветошь') < 0)

    for(let i = 0; i < oz_orders.length; i++) {

        console.log(oz_orders[i].vendor)

        const response = await axios.post('https://api-seller.ozon.ru/v4/product/info/attributes', {
                    
            "filter": {
                "offer_id": [
                    oz_orders[i].vendor
                ],
                "visibility": "ALL"
            },
            "limit": 1000,
            "sort_dir": "ASC"

        }, {
            headers: {
                'Host':'api-seller.ozon.ru',
                'Client-Id':`${process.env.OZON_CLIENT_ID}`,
                'Api-Key':`${process.env.OZON_API_KEY}`,
                'Content-Type':'application/json'
            }
        })

        if(response.data.result[0].name.indexOf('Пододеяльник') >= 0) {

            names.push({            
                'vendor': oz_orders[i].vendor,
                'name': response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                .trim()                  // убрать пробелы по краям
                                                .replace(/\s+/g, ' '),
                'size': response.data.result[0].attributes.find(o => o.id === 6773).values[0].value,
                'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                'productType': 'ПОДОДЕЯЛЬНИК С КЛАПАНОМ'
            })   

        }

        if(response.data.result[0].name.indexOf('Простыня') >= 0 && response.data.result[0].name.indexOf('белье') < 0 && response.data.result[0].name.indexOf('бельё') < 0) {

            if(response.data.result[0].name.indexOf('на резинке') >= 0) {

                names.push({            
                    'vendor': oz_orders[i].vendor,
                    'name': response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                    .trim()                  // убрать пробелы по краям
                                                    .replace(/\s+/g, ' '),
                    'size': `${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}x${response.data.result[0].attributes.find(o => o.id === 8414).values[0].value}`,
                    'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                    'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                    'productType': 'ПРОСТЫНЯ НА РЕЗИНКЕ'
                })

            }

            if(response.data.result[0].name.indexOf('на резинке') < 0) {

                names.push({            
                    'vendor': oz_orders[i].vendor,
                    'name': response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                    .trim()                  // убрать пробелы по краям
                                                    .replace(/\s+/g, ' '),
                    'size': response.data.result[0].attributes.find(o => o.id === 6771).values[0].value,
                    'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                    'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                    'productType': 'ПРОСТЫНЯ'
                })

            }

        }

        if(response.data.result[0].name.indexOf('Наволочка') >= 0 || response.data.result[0].name.indexOf('наволочка') >= 0 && response.data.result[0].name.indexOf('белье') < 0 && response.data.result[0].name.indexOf('бельё') < 0) {

            if(response.data.result[0].name.indexOf('50х70') >= 0 || response.data.result[0].name.indexOf('40х60') >= 0 || response.data.result[0].name.indexOf('50 х 70') >= 0 || response.data.result[0].name.indexOf('40 х 60') >= 0 ) {

                names.push({
                    'vendor': oz_orders[i].vendor,
                    'name': response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                    .trim()                  // убрать пробелы по краям
                                                    .replace(/\s+/g, ' '),
                    'size': response.data.result[0].attributes.find(o => o.id === 6772).values[0].value,
                    'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                    'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                    'productType': 'НАВОЛОЧКА ПРЯМОУГОЛЬНАЯ'
                })

            } else {

                names.push({
                    'vendor': oz_orders[i].vendor,
                    'name': response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                    .trim()                  // убрать пробелы по краям
                                                    .replace(/\s+/g, ' '),
                    'size': response.data.result[0].attributes.find(o => o.id === 6772).values[0].value,
                    'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                    'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                    'productType': 'НАВОЛОЧКА КВАДРАТНАЯ'
                })

            }

        }

        if(response.data.result[0].name.indexOf('белье') >= 0 || response.data.result[0].name.indexOf('бельё') >= 0) {

            if(response.data.result[0].attributes.find(o => o.id === 6772).values.length === 2) {

                if(response.data.result[0].name.indexOf('на резинке') >= 0) {

                    if(response.data.result[0].name.indexOf('х20 -') >= 0 ||response.data.result[0].name.indexOf('х 20 -') >= 0) {

                        names.push({
                            'vendor': oz_orders[i].vendor,
                            'name': `КПБ ${response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                            .trim()                  // убрать пробелы по краям
                                                            .replace(/\s+/g, ' ')}`,
                            'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}x20; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[1].value}`,
                            'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                            'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                            'productType': 'КОМПЛЕКТ'
                        })

                    }

                    if(response.data.result[0].name.indexOf('х30 -') >= 0 ||response.data.result[0].name.indexOf('х 30 -') >= 0) {

                        names.push({
                            'vendor': oz_orders[i].vendor,
                            'name': `КПБ ${response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                            .trim()                  // убрать пробелы по краям
                                                            .replace(/\s+/g, ' ')}`,
                            'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}x30; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[1].value}`,
                            'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                            'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                            'productType': 'КОМПЛЕКТ'
                        })

                    }

                    if(response.data.result[0].name.indexOf('х40') >= 0 ||response.data.result[0].name.indexOf('х 40 -') >= 0) {

                        names.push({
                            'vendor': oz_orders[i].vendor,
                            'name': `КПБ ${response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                            .trim()                  // убрать пробелы по краям
                                                            .replace(/\s+/g, ' ')}`,
                            'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}x40; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[1].value}`,
                            'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                            'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                            'productType': 'КОМПЛЕКТ'
                        })

                    }

                }

                if(response.data.result[0].name.indexOf('на резинке') < 0) {

                    names.push({
                        'vendor': oz_orders[i].vendor,
                        'name': `КПБ ${response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                            .trim()                  // убрать пробелы по краям
                                                            .replace(/\s+/g, ' ')}`,
                        'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[1].value}`,
                        'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                        'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                        'productType': 'КОМПЛЕКТ'
                    })

                }

            }

            if(response.data.result[0].attributes.find(o => o.id === 6772).values.length === 1) {

                if(response.data.result[0].name.indexOf('на резинке') >= 0) {

                    if(response.data.result[0].name.indexOf('х20 -') >= 0 ||response.data.result[0].name.indexOf('х 20 -') >= 0) {

                        names.push({
                            'vendor': oz_orders[i].vendor,
                            'name': `КПБ ${response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                            .trim()                  // убрать пробелы по краям
                                                            .replace(/\s+/g, ' ')}`,
                            'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}x20; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}`,
                            'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                            'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                            'productType': 'КОМПЛЕКТ'
                        })

                    }

                    if(response.data.result[0].name.indexOf('х30 -') >= 0 ||response.data.result[0].name.indexOf('х 30 -') >= 0) {

                        names.push({
                            'vendor': oz_orders[i].vendor,
                            'name': `КПБ ${response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                            .trim()                  // убрать пробелы по краям
                                                            .replace(/\s+/g, ' ')}`,
                            'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}x30; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}`,
                            'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                            'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                            'productType': 'КОМПЛЕКТ'
                        })

                    }

                    if(response.data.result[0].name.indexOf('х40 -') >= 0 ||response.data.result[0].name.indexOf('х 40 -') >= 0) {

                        names.push({
                            'vendor': oz_orders[i].vendor,
                            'name': `КПБ ${response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                            .trim()                  // убрать пробелы по краям
                                                            .replace(/\s+/g, ' ')}`,
                            'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}x40; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}`,
                            'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                            'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                            'productType': 'КОМПЛЕКТ'
                        })

                    }

                }

                if(response.data.result[0].name.indexOf('на резинке') < 0) {

                    names.push({
                        'vendor': oz_orders[i].vendor,
                        'name': `КПБ ${response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                            .trim()                  // убрать пробелы по краям
                                                            .replace(/\s+/g, ' ')}`,
                        'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}`,
                        'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                        'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                        'productType': 'КОМПЛЕКТ'
                    })

                }

            }
            

        }

    }

    names.forEach(el => {

            if(nat_cat.findIndex(o => o.name === el.name) < 0) {
                new_items.push(el.name)
            }

            if(nat_cat.findIndex(o => o.name === el.name) >= 0) {
                current_items.push(el.name)
            }

    })

    html += `<section class="table">
                <div class="marks-table">
                    <div class="marks-table-header">
                        <div class="header-cell">Наименование</div>
                        <div class="header-cell">Статус</div>                            
                    </div>
                <div class="header-wrapper"></div>`

    names.forEach(elem => {
        if(new_items.indexOf(elem.name) >= 0) {
            html += `<div class="table-row">
                        <span id="name">${elem.name}</span>
                        <span id="status-new">Новый товар</span>
                     </div>`
        } else {
            html += `<div class="table-row">
                        <span id="name">${elem.name}</span>
                        <span id="status-current">Актуальный товар</span>
                     </div>`
        }
    })

    html += `</section>
             <section class="action-form">
                <button id="current-order"><a href="http://localhost:3030/ozon_marks_order" target="_blank">Создать заказы маркировки</a></button>
             </section>
             <div class="body-wrapper"></div>                        
             ${footerComponent}`

    async function createImport(array) {

        const fileName = './public/IMPORT_TNVED_6302 (3).xlsx'
        
        const wb = new exl.Workbook()

        await wb.xlsx.readFile(fileName)

        const ws = wb.getWorksheet('IMPORT_TNVED_6302')

        let cellNumber = 5

        for(let i = 0; i < array.length; i++) {

            ws.getCell(`A${cellNumber}`).value = 6302
            ws.getCell(`B${cellNumber}`).value = names.find(o => o.name === array[i]).name
            ws.getCell(`C${cellNumber}`).value = 'Ивановский текстиль'
            ws.getCell(`D${cellNumber}`).value = 'Артикул'
            ws.getCell(`E${cellNumber}`).value = names.find(o => o.name === array[i]).vendor
            ws.getCell(`F${cellNumber}`).value = names.find(o => o.name === array[i]).productType
            ws.getCell(`G${cellNumber}`).value = names.find(o => o.name === array[i]).color
            ws.getCell(`H${cellNumber}`).value = 'ВЗРОСЛЫЙ'

            if(names.find(o => o.name === array[i]).cloth === 'КРЕП-ЖАТКА' || names.find(o => o.name === array[i]).cloth === 'КРЕП ЖАТКА') ws.getCell(`I${cellNumber}`).value = 'КРЕП'
            if(names.find(o => o.name === array[i]).cloth === 'ВАРЕНЫЙ ХЛОПОК') ws.getCell(`I${cellNumber}`).value = 'ХЛОПКОВАЯ ТКАНЬ'
            if(names.find(o => o.name === array[i]).cloth === 'ЛЕН' || names.find(o => o.name === array[i]).cloth === 'ЛЁН') ws.getCell(`I${cellNumber}`).value = 'ЛЬНЯНАЯ ТКАНЬ'
            if(names.find(o => o.name === array[i]).cloth === 'СТРАЙП САТИН') ws.getCell(`I${cellNumber}`).value = 'СТРАЙП-САТИН'
            if(names.find(o => o.name === array[i]).cloth === 'САТИН ЛЮКС') ws.getCell(`I${cellNumber}`).value = 'САТИН'
            if(names.find(o => o.name === array[i]).cloth !== 'САТИН ЛЮКС' && names.find(o => o.name === array[i]).cloth !== 'СТРАЙП САТИН' && names.find(o => o.name === array[i]).cloth !== 'ВАРЕНЫЙ ХЛОПОК' && names.find(o => o.name === array[i]).cloth !== 'ЛЕН' && names.find(o => o.name === array[i]).cloth !== 'ЛЁН') ws.getCell(`I${cellNumber}`).value = names.find(o => o.name === array[i]).cloth
            
            if(names.find(o => o.name === array[i]).cloth === 'ПОЛИСАТИН') ws.getCell(`J${cellNumber}`).value = '100% Полиэстер'

            if(names.find(o => o.name === array[i]).cloth === 'ТЕНСЕЛЬ') ws.getCell(`J${cellNumber}`).value = '100% Лиоцелл'
            if(names.find(o => o.name === array[i]).cloth === 'ЛЕН' || names.find(o => o.name === array[i]).cloth === 'ЛЁН') ws.getCell(`J${cellNumber}`).value = '100% Лен'
            if(names.find(o => o.name === array[i]).cloth !== 'КРЕП-ЖАТКА' && names.find(o => o.name === array[i]).cloth !== 'КРЕП ЖАТКА' && names.find(o => o.name === array[i]).cloth !== 'ПОЛИСАТИН' && names.find(o => o.name === array[i]).cloth !== 'ТЕНСЕЛЬ' && names.find(o => o.name === array[i]).cloth !== 'ЛЕН' && names.find(o => o.name === array[i]).cloth !== 'ЛЁН') ws.getCell(`J${cellNumber}`).value = '100% Хлопок'

            ws.getCell(`K${cellNumber}`).value = names.find(o => o.name === array[i]).size
            ws.getCell(`L${cellNumber}`).value = '6302100001'
            ws.getCell(`M${cellNumber}`).value = 'ТР ТС 017/2011 "О безопасности продукции легкой промышленности'
            ws.getCell(`N${cellNumber}`).value = 'На модерации'                

            cellNumber++

        }

        ws.unMergeCells('D2')

        ws.getCell('E2').value = '13914'

        ws.getCell('E2').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor:{argb:'E3E3E3'}
        }

        ws.getCell('E2').font = {
            size: 10,
            name: 'Arial'
        }

        ws.getCell('E2').alignment = {
            horizontal: 'center',
            vertical: 'bottom'
        }        

        const date_ob = new Date()

        let month = date_ob.getMonth() + 1

        let filePath = ''

        month < 10 ? filePath = `./public/ozon/IMPORT_TNVED_6302_${date_ob.getDate()}_0${month}_ozon` : filePath = `./public/ozon/IMPORT_TNVED_6302_${date_ob.getDate()}_0${month}_ozon`

        fs.access(`${filePath}.xlsx`, fs.constants.R_OK, async (err) => {
            if(err) {
                await wb.xlsx.writeFile(`${filePath}.xlsx`)
            } else {
                let count = 1
                fs.access(`${filePath}_(1).xlsx`, fs.constants.R_OK, async (err) => {
                    if(err) {
                        await wb.xlsx.writeFile(`${filePath}_(1).xlsx`)
                    } else {
                        await wb.xlsx.writeFile(`${filePath}_(2).xlsx`)
                    }
                })
                
            }
        })

    }

    if(new_items.length > 0) await createImport(new_items)

    res.send(html)

})

app.get('/ozon_marks_order', async function(req, res){
    
    let oz_orders = []
    const nat_cat = []
    const gtins = []
    let names = []

    html = `${headerComponent}
                    <title>Заказ маркировки - OZON</title>
                </head>
                <body>
                    ${navComponent}
                        <section class="sub-nav import-main">
                            <div class="import-control">`

    await renderImportButtons(buttons)
    await renderMarkingButtons()
    await renderExtraButtons()

    html += `</section>`

    const wb = new exl.Workbook()
    
    const fileName = './public/Краткий отчет.xlsx'    

    await wb.xlsx.readFile(fileName)
        
    const nc_ws = wb.getWorksheet('Краткий отчет')

    const nc_c1 = nc_ws.getColumn(1)

    nc_c1.eachCell({includeEmpty: false}, (c, rowNumber) => {

        if(rowNumber < 5) return
        gtins.push(c.value)

    })

    const nc_c2 = nc_ws.getColumn(2)

    nc_c2.eachCell({includeEmpty: false}, (c, rowNumber) => {

        if(rowNumber < 5) return
        nat_cat.push(c.value.trim())

    })

    let response = await axios.post('https://api-seller.ozon.ru/v3/posting/fbs/list', {
        
        'dir': 'asc',
        'filter': {
            'since':`2025-01-01T00:00:00.000Z`,
            'status':'awaiting_packaging',
            'to':'2025-12-31T23:59:59.000Z'
        },
        'limit': 1000,
        'offset':0

    }, {

        headers: {
            'Host':'api-seller.ozon.ru',
            'Client-Id':`${process.env.OZON_CLIENT_ID}`,
            'Api-Key':`${process.env.OZON_API_KEY}`,
            'Content-Type':'application/json'
        }

    })

    let result = response.data.result.postings

    result = result.filter(o => {

        if(o.delivery_method.id === 23463726191000) {

            return o

        }

    })

    result.forEach(el => {

        for(let i = 0; i < el.products.length; i++) {

            // console.log(el.products[i].offer_id)
            if(oz_orders.findIndex(o => o.vendor === el.products[i].offer_id) >= 0) {

                oz_orders.find(o => o.vendor === el.products[i].offer_id).quantity += Number(el.products[i].quantity)

            }

            // console.log(oz_orders.findIndex(o => o.vendor === el.products[i].offer_id))

            if(oz_orders.findIndex(o => o.vendor === el.products[i].offer_id) < 0) {

                if(el.products[i].name.indexOf('белье') >= 0 || el.products[i].name.indexOf('бельё') >= 0) {

                    oz_orders.push({
                        'name': `КПБ ${el.products[i].name}`,
                        'vendor': el.products[i].offer_id,
                        'quantity': Number(el.products[i].quantity)
                    })

                }

                if(el.products[i].name.indexOf('белье') < 0 && el.products[i].name.indexOf('бельё') < 0) {

                    oz_orders.push({
                        'name': el.products[i].name,
                        'vendor': el.products[i].offer_id,
                        'quantity': Number(el.products[i].quantity)
                    })

                }

            }

        }
    })

    oz_orders = oz_orders.filter(o => o.name.indexOf('Одеяло') < 0 && o.name.indexOf('Подушка') < 0 && o.name.indexOf('Матрас') < 0 && o.name.indexOf('Ветошь') < 0 && o.name.indexOf('холстопрошивное') < 0)

    html += `<section class="table">
                    <div class="marks-table">
                        <div class="marks-table-header">
                            <div class="header-cell">Наименование</div>
                            <div class="header-cell">Статус</div>                            
                        </div>
                    <div class="header-wrapper"></div>`

    for(let i = 0; i < oz_orders.length; i++) {
        html += `<div class="table-row">
                    <span id="name">${oz_orders[i].name}</span>
                    <span id="status-current">Актуальный товар</span>
                    <span id="quantity">${oz_orders[i].quantity}</span>
                 </div>`
    }

    html += `</section>
             <div class="body-wrapper"></div>                        
             ${footerComponent}`

    function createNameList() {

        let orderList = []
        let _temp = []

        for (let i = 0; i < oz_orders.length; i++) {

            if(nat_cat.indexOf(oz_orders[i].name) >= 0) {

                _temp.push(oz_orders[i].name)

            }            
            
            if(_temp.length%10 === 0) {
                if(_temp.length !== 0) {
                    orderList.push(_temp)
                }
                _temp = []
            }

        }        

        if(_temp.length !== 0) {
            orderList.push(_temp)
        }
        _temp = []

        return orderList

    }

    function createQuantityList() {

        let quantityList = []
        let temp = []

        for(let i = 0; i < oz_orders.length; i++) {

            if(nat_cat.indexOf(oz_orders[i].name) >= 0) {

                temp.push(oz_orders[i].quantity)

            }            

            if(temp.length%10 === 0) {
                if(temp.length !== 0) {
                    quantityList.splice(-1, 0, ...quantityList.splice(-1, 1, temp))
                }
                temp = []
            }

        }

        if(temp.length !== 0) {

            quantityList.splice(-1, 0, ...quantityList.splice(-1, 1, temp))

        }

        return quantityList

    }

    function createOrder() {

        let List = createNameList()
        let Quantity = createQuantityList()
        let content = ``

        for(let i = 0; i < List.length; i++) {
            if(List[i].length > 0) {
                content += `<?xml version="1.0" encoding="utf-8"?>
                            <order xmlns="urn:oms.order" xsi:schemaLocation="urn:oms.order schema.xsd" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
                                <lp>
                                    <productGroup>lp</productGroup>
                                    <contactPerson>333</contactPerson>
                                    <releaseMethodType>REMARK</releaseMethodType>
                                    <createMethodType>SELF_MADE</createMethodType>
                                    <productionOrderId>OZON_${i+1}</productionOrderId>
                                    <products>`
                
                    for(let j = 0; j < List[i].length; j++) {                
                        if(nat_cat.indexOf(List[i][j]) >= 0) {
                            content += `<product>
                                            <gtin>0${gtins[nat_cat.indexOf(List[i][j])]}</gtin>
                                            <quantity>${Quantity[i][j]}</quantity>
                                            <serialNumberType>OPERATOR</serialNumberType>
                                            <cisType>UNIT</cisType>
                                            <templateId>10</templateId>
                                        </product>`
                        }
                    }

                content += `    </products>
                            </lp>
                        </order>`

            }

            const date_ob = new Date()

            let month = date_ob.getMonth() + 1

            let filePath = ''

            month < 10 ? filePath = `./public/orders/lp_ozon_${i}_${date_ob.getDate()}_0${month}.xml` : filePath = `./public/orders/lp_ozon_${i}_${date_ob.getDate()}_${month}.xml`

            if(content !== ``) {
                fs.writeFileSync(filePath, content)
            }

            content = ``

        }   

    }

    createOrder()

    res.send(html)

})

app.get('/ozon/:from', async function(req, res){

    const nat_cat = []
    const nat_catGtins = []
    const nat_catNames = []
    let oz_orders = []
    const new_items = []
    const current_items = []
    const names = []

    html = `${headerComponent}
                    <title>Импорт - OZON</title>
                </head>
                <body>
                        ${navComponent}
                        <section class="sub-nav import-main">
                            <div class="import-control">`

    await renderImportButtons(buttons)
    await renderMarkingButtons()
    await renderExtraButtons()

    html += `</section>`

    const wb = new exl.Workbook()

    await wb.xlsx.readFile('./public/Краткий отчет.xlsx')

    const nc_ws = wb.getWorksheet('Краткий отчет')

    const nc_c1 = nc_ws.getColumn(1)

    nc_c1.eachCell({includeEmpty: false}, (c, rowNumber) => {

        if(rowNumber < 5) return
        nat_catGtins.push(c.value)

    })

    const nc_c2 = nc_ws.getColumn(2)

    nc_c2.eachCell({includeEmpty: false}, (c, rowNumber) => {

        if(rowNumber < 5) return
        nat_catNames.push(c.value.trim())

    })

    for(let i = 0; i < nat_catNames.length; i++) {

        nat_cat.push({
            'gtin': nat_catGtins[i],
            'name': nat_catNames[i]
        })

    }

    let response = await axios.post('https://api-seller.ozon.ru/v3/posting/fbs/list', {
        
        'dir': 'asc',
        'filter': {
            'since':`${req.params.from}:50:00.000Z`,
            'status':'awaiting_deliver',
            'to':'2025-12-31T23:59:59.000Z'
        },
        'limit': 1000,
        'offset':0

    }, {

        headers: {
            'Host':'api-seller.ozon.ru',
            'Client-Id':`${process.env.OZON_CLIENT_ID}`,
            'Api-Key':`${process.env.OZON_API_KEY}`,
            'Content-Type':'application/json'
        }

    })

    let result = response.data.result.postings

    result = result.filter(o => {

        if(o.delivery_method.id === 23463726191000) {

            return o

        }

    })

    result.forEach(el => {

        for(let i = 0; i < el.products.length; i++) {

            // console.log(el.products[i].offer_id)
            if(oz_orders.findIndex(o => o.vendor === el.products[i].offer_id) >= 0) {

                oz_orders.find(o => o.vendor === el.products[i].offer_id).quantity += Number(el.products[i].quantity)

            }

            // console.log(oz_orders.findIndex(o => o.vendor === el.products[i].offer_id))

            if(oz_orders.findIndex(o => o.vendor === el.products[i].offer_id) < 0) {

                oz_orders.push({
                    'name': el.products[i].name,
                    'vendor': el.products[i].offer_id,
                    'quantity': Number(el.products[i].quantity)
                })

            }

        }
    })

    oz_orders = oz_orders.filter(o => o.name.indexOf('Одеяло') < 0 && o.name.indexOf('Подушка') < 0 && o.name.indexOf('Матрас') < 0 && o.name.indexOf('Ветошь') < 0)

    for(let i = 0; i < oz_orders.length; i++) {

        console.log(oz_orders[i].vendor)
        
        const response = await axios.post('https://api-seller.ozon.ru/v4/product/info/attributes', {
                    
            "filter": {
                "offer_id": [
                    oz_orders[i].vendor
                ],
                "visibility": "ALL"
            },
            "limit": 1000,
            "sort_dir": "ASC"

        }, {
            headers: {
                'Host':'api-seller.ozon.ru',
                'Client-Id':`${process.env.OZON_CLIENT_ID}`,
                'Api-Key':`${process.env.OZON_API_KEY}`,
                'Content-Type':'application/json'
            }
        })

        if(response.data.result[0].name.indexOf('Пододеяльник') >= 0) {

            names.push({            
                'vendor': oz_orders[i].vendor,
                'name': response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                .trim()                  // убрать пробелы по краям
                                                .replace(/\s+/g, ' '),
                'size': response.data.result[0].attributes.find(o => o.id === 6773).values[0].value,
                'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                'productType': 'ПОДОДЕЯЛЬНИК С КЛАПАНОМ'
            })   

        }

        if(response.data.result[0].name.indexOf('Простыня') >= 0 && response.data.result[0].name.indexOf('белье') < 0 && response.data.result[0].name.indexOf('бельё') < 0) {

            if(response.data.result[0].name.indexOf('на резинке') >= 0) {

                names.push({            
                    'vendor': oz_orders[i].vendor,
                    'name': response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                    .trim()                  // убрать пробелы по краям
                                                    .replace(/\s+/g, ' '),
                    'size': `${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}x${response.data.result[0].attributes.find(o => o.id === 8414).values[0].value}`,
                    'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                    'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                    'productType': 'ПРОСТЫНЯ НА РЕЗИНКЕ'
                })

            }

            if(response.data.result[0].name.indexOf('на резинке') < 0) {

                names.push({            
                    'vendor': oz_orders[i].vendor,
                    'name': response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                    .trim()                  // убрать пробелы по краям
                                                    .replace(/\s+/g, ' '),
                    'size': response.data.result[0].attributes.find(o => o.id === 6771).values[0].value,
                    'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                    'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                    'productType': 'ПРОСТЫНЯ'
                })

            }

        }

        if(response.data.result[0].name.indexOf('Наволочка') >= 0 || response.data.result[0].name.indexOf('наволочка') >= 0 && response.data.result[0].name.indexOf('белье') < 0 && response.data.result[0].name.indexOf('бельё') < 0) {

            if(response.data.result[0].name.indexOf('50х70') >= 0 || response.data.result[0].name.indexOf('40х60') >= 0 || response.data.result[0].name.indexOf('50 х 70') >= 0 || response.data.result[0].name.indexOf('40 х 60') >= 0 ) {

                names.push({
                    'vendor': oz_orders[i].vendor,
                    'name': response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                    .trim()                  // убрать пробелы по краям
                                                    .replace(/\s+/g, ' '),
                    'size': response.data.result[0].attributes.find(o => o.id === 6772).values[0].value,
                    'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                    'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                    'productType': 'НАВОЛОЧКА ПРЯМОУГОЛЬНАЯ'
                })

            } else {

                names.push({
                    'vendor': oz_orders[i].vendor,
                    'name': response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                    .trim()                  // убрать пробелы по краям
                                                    .replace(/\s+/g, ' '),
                    'size': response.data.result[0].attributes.find(o => o.id === 6772).values[0].value,
                    'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                    'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                    'productType': 'НАВОЛОЧКА КВАДРАТНАЯ'
                })

            }

        }

        if(response.data.result[0].name.indexOf('белье') >= 0 || response.data.result[0].name.indexOf('бельё') >= 0) {

            if(response.data.result[0].attributes.find(o => o.id === 6772).values.length === 2) {

                if(response.data.result[0].name.indexOf('на резинке') >= 0) {

                    if(response.data.result[0].name.indexOf('х20 -') >= 0 ||response.data.result[0].name.indexOf('х 20 -') >= 0) {

                        names.push({
                            'vendor': oz_orders[i].vendor,
                            'name': `КПБ ${response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                            .trim()                  // убрать пробелы по краям
                                                            .replace(/\s+/g, ' ')}`,
                            'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}x20; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[1].value}`,
                            'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                            'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                            'productType': 'КОМПЛЕКТ'
                        })

                    }

                    if(response.data.result[0].name.indexOf('х30 -') >= 0 ||response.data.result[0].name.indexOf('х 30 -') >= 0) {

                        names.push({
                            'vendor': oz_orders[i].vendor,
                            'name': `КПБ ${response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                            .trim()                  // убрать пробелы по краям
                                                            .replace(/\s+/g, ' ')}`,
                            'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}x30; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[1].value}`,
                            'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                            'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                            'productType': 'КОМПЛЕКТ'
                        })

                    }

                    if(response.data.result[0].name.indexOf('х40') >= 0 ||response.data.result[0].name.indexOf('х 40 -') >= 0) {

                        names.push({
                            'vendor': oz_orders[i].vendor,
                            'name': `КПБ ${response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                            .trim()                  // убрать пробелы по краям
                                                            .replace(/\s+/g, ' ')}`,
                            'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}x40; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[1].value}`,
                            'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                            'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                            'productType': 'КОМПЛЕКТ'
                        })

                    }

                }

                if(response.data.result[0].name.indexOf('на резинке') < 0) {

                    names.push({
                        'vendor': oz_orders[i].vendor,
                        'name': `КПБ ${response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                            .trim()                  // убрать пробелы по краям
                                                            .replace(/\s+/g, ' ')}`,
                        'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[1].value}`,
                        'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                        'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                        'productType': 'КОМПЛЕКТ'
                    })

                }

            }

            if(response.data.result[0].attributes.find(o => o.id === 6772).values.length === 1) {

                if(response.data.result[0].name.indexOf('на резинке') >= 0) {

                    if(response.data.result[0].name.indexOf('х20 -') >= 0 ||response.data.result[0].name.indexOf('х 20 -') >= 0) {

                        names.push({
                            'vendor': oz_orders[i].vendor,
                            'name': `КПБ ${response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                            .trim()                  // убрать пробелы по краям
                                                            .replace(/\s+/g, ' ')}`,
                            'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}x20; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}`,
                            'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                            'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                            'productType': 'КОМПЛЕКТ'
                        })

                    }

                    if(response.data.result[0].name.indexOf('х30 -') >= 0 ||response.data.result[0].name.indexOf('х 30 -') >= 0) {

                        names.push({
                            'vendor': oz_orders[i].vendor,
                            'name': `КПБ ${response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                            .trim()                  // убрать пробелы по краям
                                                            .replace(/\s+/g, ' ')}`,
                            'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}x30; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}`,
                            'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                            'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                            'productType': 'КОМПЛЕКТ'
                        })

                    }

                    if(response.data.result[0].name.indexOf('х40 -') >= 0 ||response.data.result[0].name.indexOf('х 40 -') >= 0) {

                        names.push({
                            'vendor': oz_orders[i].vendor,
                            'name': `КПБ ${response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                            .trim()                  // убрать пробелы по краям
                                                            .replace(/\s+/g, ' ')}`,
                            'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}x40; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}`,
                            'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                            'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                            'productType': 'КОМПЛЕКТ'
                        })

                    }

                }

                if(response.data.result[0].name.indexOf('на резинке') < 0) {

                    names.push({
                        'vendor': oz_orders[i].vendor,
                        'name': `КПБ ${response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                            .trim()                  // убрать пробелы по краям
                                                            .replace(/\s+/g, ' ')}`,
                        'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}`,
                        'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                        'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                        'productType': 'КОМПЛЕКТ'
                    })

                }

            }
            

        }

    }

    names.forEach(el => {

            if(nat_cat.findIndex(o => o.name === el.name) < 0) {
                new_items.push(el.name)
            }

            if(nat_cat.findIndex(o => o.name === el.name) >= 0) {
                current_items.push(el.name)
            }

    })

    html += `<section class="table">
                <div class="marks-table">
                    <div class="marks-table-header">
                        <div class="header-cell">Наименование</div>
                        <div class="header-cell">Статус</div>                            
                    </div>
                <div class="header-wrapper"></div>`

    names.forEach(elem => {
        if(new_items.indexOf(elem.name) >= 0) {
            html += `<div class="table-row">
                        <span id="name">${elem.name}</span>
                        <span id="status-new">Новый товар</span>
                     </div>`
        } else {
            html += `<div class="table-row">
                        <span id="name">${elem.name}</span>
                        <span id="status-current">Актуальный товар</span>
                     </div>`
        }
    })

    html += `</section>
             <section class="action-form">
                <button id="current-order"><a href="http://localhost:3030/ozon_marks_order/${req.params.from}" target="_blank">Создать заказы маркировки</a></button>
             </section>
             <div class="body-wrapper"></div>                        
             ${footerComponent}`

    async function createImport(array) {

        const fileName = './public/IMPORT_TNVED_6302 (3).xlsx'
        
        const wb = new exl.Workbook()

        await wb.xlsx.readFile(fileName)

        const ws = wb.getWorksheet('IMPORT_TNVED_6302')

        let cellNumber = 5

        for(let i = 0; i < array.length; i++) {

            ws.getCell(`A${cellNumber}`).value = 6302
            ws.getCell(`B${cellNumber}`).value = names.find(o => o.name === array[i]).name
            ws.getCell(`C${cellNumber}`).value = 'Ивановский текстиль'
            ws.getCell(`D${cellNumber}`).value = 'Артикул'
            ws.getCell(`E${cellNumber}`).value = names.find(o => o.name === array[i]).vendor
            ws.getCell(`F${cellNumber}`).value = names.find(o => o.name === array[i]).productType
            ws.getCell(`G${cellNumber}`).value = names.find(o => o.name === array[i]).color
            ws.getCell(`H${cellNumber}`).value = 'ВЗРОСЛЫЙ'

            if(names.find(o => o.name === array[i]).cloth === 'КРЕП-ЖАТКА' || names.find(o => o.name === array[i]).cloth === 'КРЕП ЖАТКА') ws.getCell(`I${cellNumber}`).value = 'КРЕП'
            if(names.find(o => o.name === array[i]).cloth === 'ВАРЕНЫЙ ХЛОПОК') ws.getCell(`I${cellNumber}`).value = 'ХЛОПКОВАЯ ТКАНЬ'
            if(names.find(o => o.name === array[i]).cloth === 'ЛЕН' || names.find(o => o.name === array[i]).cloth === 'ЛЁН') ws.getCell(`I${cellNumber}`).value = 'ЛЬНЯНАЯ ТКАНЬ'
            if(names.find(o => o.name === array[i]).cloth === 'СТРАЙП САТИН') ws.getCell(`I${cellNumber}`).value = 'СТРАЙП-САТИН'
            if(names.find(o => o.name === array[i]).cloth === 'САТИН ЛЮКС') ws.getCell(`I${cellNumber}`).value = 'САТИН'
            if(names.find(o => o.name === array[i]).cloth !== 'САТИН ЛЮКС' && names.find(o => o.name === array[i]).cloth !== 'СТРАЙП САТИН' && names.find(o => o.name === array[i]).cloth !== 'ВАРЕНЫЙ ХЛОПОК' && names.find(o => o.name === array[i]).cloth !== 'ЛЕН' && names.find(o => o.name === array[i]).cloth !== 'ЛЁН') ws.getCell(`I${cellNumber}`).value = names.find(o => o.name === array[i]).cloth
            
            if(names.find(o => o.name === array[i]).cloth === 'ПОЛИСАТИН') ws.getCell(`J${cellNumber}`).value = '100% Полиэстер'

            if(names.find(o => o.name === array[i]).cloth === 'ТЕНСЕЛЬ') ws.getCell(`J${cellNumber}`).value = '100% Лиоцелл'
            if(names.find(o => o.name === array[i]).cloth === 'ЛЕН' || names.find(o => o.name === array[i]).cloth === 'ЛЁН') ws.getCell(`J${cellNumber}`).value = '100% Лен'
            if(names.find(o => o.name === array[i]).cloth !== 'КРЕП-ЖАТКА' && names.find(o => o.name === array[i]).cloth !== 'КРЕП ЖАТКА' && names.find(o => o.name === array[i]).cloth !== 'ПОЛИСАТИН' && names.find(o => o.name === array[i]).cloth !== 'ТЕНСЕЛЬ' && names.find(o => o.name === array[i]).cloth !== 'ЛЕН' && names.find(o => o.name === array[i]).cloth !== 'ЛЁН') ws.getCell(`J${cellNumber}`).value = '100% Хлопок'

            ws.getCell(`K${cellNumber}`).value = names.find(o => o.name === array[i]).size
            ws.getCell(`L${cellNumber}`).value = '6302100001'
            ws.getCell(`M${cellNumber}`).value = 'ТР ТС 017/2011 "О безопасности продукции легкой промышленности'
            ws.getCell(`N${cellNumber}`).value = 'На модерации'                

            cellNumber++

        }

        ws.unMergeCells('D2')

        ws.getCell('E2').value = '13914'

        ws.getCell('E2').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor:{argb:'E3E3E3'}
        }

        ws.getCell('E2').font = {
            size: 10,
            name: 'Arial'
        }

        ws.getCell('E2').alignment = {
            horizontal: 'center',
            vertical: 'bottom'
        }        

        const date_ob = new Date()

        let month = date_ob.getMonth() + 1

        let filePath = ''

        month < 10 ? filePath = `./public/ozon/IMPORT_TNVED_6302_${date_ob.getDate()}_0${month}_ozon` : filePath = `./public/ozon/IMPORT_TNVED_6302_${date_ob.getDate()}_0${month}_ozon`

        fs.access(`${filePath}.xlsx`, fs.constants.R_OK, async (err) => {
            if(err) {
                await wb.xlsx.writeFile(`${filePath}.xlsx`)
            } else {
                let count = 1
                fs.access(`${filePath}_(1).xlsx`, fs.constants.R_OK, async (err) => {
                    if(err) {
                        await wb.xlsx.writeFile(`${filePath}_(1).xlsx`)
                    } else {
                        await wb.xlsx.writeFile(`${filePath}_(2).xlsx`)
                    }
                })
                
            }
        })

    }

    if(new_items.length > 0) await createImport(new_items)

    res.send(html)

})

app.get('/ozon_marks_order/:from', async function(req, res){
    
    let oz_orders = []
    const nat_cat = []
    const gtins = []
    let names = []

    html = `${headerComponent}
                    <title>Заказ маркировки - OZON</title>
                </head>
                <body>
                    ${navComponent}
                        <section class="sub-nav import-main">
                            <div class="import-control">`

    await renderImportButtons(buttons)
    await renderMarkingButtons()
    await renderExtraButtons()

    html += `</section>`

    const wb = new exl.Workbook()
    
    const fileName = './public/Краткий отчет.xlsx'    

    await wb.xlsx.readFile(fileName)
        
    const nc_ws = wb.getWorksheet('Краткий отчет')

    const nc_c1 = nc_ws.getColumn(1)

    nc_c1.eachCell({includeEmpty: false}, (c, rowNumber) => {

        if(rowNumber < 5) return
        gtins.push(c.value)

    })

    const nc_c2 = nc_ws.getColumn(2)

    nc_c2.eachCell({includeEmpty: false}, (c, rowNumber) => {

        if(rowNumber < 5) return
        nat_cat.push(c.value.trim())

    })

    let response = await axios.post('https://api-seller.ozon.ru/v3/posting/fbs/list', {
        
        'dir': 'asc',
        'filter': {
            'since':`${req.params.from}:00:00.000Z`,
            'status':'awaiting_deliver',
            'to':'2025-12-31T23:59:59.000Z'
        },
        'limit': 1000,
        'offset':0

    }, {

        headers: {
            'Host':'api-seller.ozon.ru',
            'Client-Id':`${process.env.OZON_CLIENT_ID}`,
            'Api-Key':`${process.env.OZON_API_KEY}`,
            'Content-Type':'application/json'
        }

    })

    let result = response.data.result.postings

    result = result.filter(o => {

        if(o.delivery_method.id === 23463726191000) {

            return o

        }

    })

    result.forEach(el => {

        for(let i = 0; i < el.products.length; i++) {

            // console.log(el.products[i].offer_id)
            if(oz_orders.findIndex(o => o.vendor === el.products[i].offer_id) >= 0) {

                oz_orders.find(o => o.vendor === el.products[i].offer_id).quantity += Number(el.products[i].quantity)

            }

            // console.log(oz_orders.findIndex(o => o.vendor === el.products[i].offer_id))

            if(oz_orders.findIndex(o => o.vendor === el.products[i].offer_id) < 0) {

                if(el.products[i].name.indexOf('белье') >= 0 || el.products[i].name.indexOf('бельё') >= 0) {

                    oz_orders.push({
                        'name': `КПБ ${el.products[i].name}`,
                        'vendor': el.products[i].offer_id,
                        'quantity': Number(el.products[i].quantity)
                    })

                }

                if(el.products[i].name.indexOf('белье') < 0 && el.products[i].name.indexOf('бельё') < 0) {

                    oz_orders.push({
                        'name': el.products[i].name,
                        'vendor': el.products[i].offer_id,
                        'quantity': Number(el.products[i].quantity)
                    })

                }

            }

        }
    })

    oz_orders = oz_orders.filter(o => o.name.indexOf('Одеяло') < 0 && o.name.indexOf('Подушка') < 0 && o.name.indexOf('Матрас') < 0 && o.name.indexOf('Ветошь') < 0 && o.name.indexOf('холстопрошивное') < 0)

    html += `<section class="table">
                    <div class="marks-table">
                        <div class="marks-table-header">
                            <div class="header-cell">Наименование</div>
                            <div class="header-cell">Статус</div>                            
                        </div>
                    <div class="header-wrapper"></div>`

    for(let i = 0; i < oz_orders.length; i++) {
        html += `<div class="table-row">
                    <span id="name">${oz_orders[i].name}</span>
                    <span id="status-current">Актуальный товар</span>
                    <span id="quantity">${oz_orders[i].quantity}</span>
                 </div>`
    }

    html += `</section>
             <div class="body-wrapper"></div>                        
             ${footerComponent}`

    function createNameList() {

        let orderList = []
        let _temp = []

        for (let i = 0; i < oz_orders.length; i++) {

            if(nat_cat.indexOf(oz_orders[i].name) >= 0) {

                _temp.push(oz_orders[i].name)

            }            
            
            if(_temp.length%10 === 0) {
                if(_temp.length !== 0) {
                    orderList.push(_temp)
                }
                _temp = []
            }

        }        

        if(_temp.length !== 0) {
            orderList.push(_temp)
        }
        _temp = []

        return orderList

    }

    function createQuantityList() {

        let quantityList = []
        let temp = []

        for(let i = 0; i < oz_orders.length; i++) {

            if(nat_cat.indexOf(oz_orders[i].name) >= 0) {

                temp.push(oz_orders[i].quantity)

            }            

            if(temp.length%10 === 0) {
                if(temp.length !== 0) {
                    quantityList.splice(-1, 0, ...quantityList.splice(-1, 1, temp))
                }
                temp = []
            }

        }

        if(temp.length !== 0) {

            quantityList.splice(-1, 0, ...quantityList.splice(-1, 1, temp))

        }

        return quantityList

    }

    function createOrder() {

        let List = createNameList()
        let Quantity = createQuantityList()
        let content = ``

        for(let i = 0; i < List.length; i++) {
            if(List[i].length > 0) {
                content += `<?xml version="1.0" encoding="utf-8"?>
                            <order xmlns="urn:oms.order" xsi:schemaLocation="urn:oms.order schema.xsd" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
                                <lp>
                                    <productGroup>lp</productGroup>
                                    <contactPerson>333</contactPerson>
                                    <releaseMethodType>REMARK</releaseMethodType>
                                    <createMethodType>SELF_MADE</createMethodType>
                                    <productionOrderId>OZON_${i+1}</productionOrderId>
                                    <products>`
                
                    for(let j = 0; j < List[i].length; j++) {                
                        if(nat_cat.indexOf(List[i][j]) >= 0) {
                            content += `<product>
                                            <gtin>0${gtins[nat_cat.indexOf(List[i][j])]}</gtin>
                                            <quantity>${Quantity[i][j]}</quantity>
                                            <serialNumberType>OPERATOR</serialNumberType>
                                            <cisType>UNIT</cisType>
                                            <templateId>10</templateId>
                                        </product>`
                        }
                    }

                content += `    </products>
                            </lp>
                        </order>`

            }

            const date_ob = new Date()

            let month = date_ob.getMonth() + 1

            let filePath = ''

            month < 10 ? filePath = `./public/orders/lp_ozon_${i}_${date_ob.getDate()}_0${month}.xml` : filePath = `./public/orders/lp_ozon_${i}_${date_ob.getDate()}_${month}.xml`

            if(content !== ``) {
                fs.writeFileSync(filePath, content)
            }

            content = ``

        }   

    }

    createOrder()

    res.send(html)

})

app.get('/wildberries', async function(req, res){
    
    const new_items = []
    const current_items = []
    const moderation_items = []
    const wb_orders = []
    const nat_cat = []
    let names = []

    html = `${headerComponent}
                    <title>Импорт - WILDBERRIES</title>
                </head>
                <body>
                    ${navComponent}
                        <section class="sub-nav import-main">
                            <div class="import-control">`

    await renderImportButtons(buttons)
    await renderMarkingButtons()
    await renderExtraButtons()

    html += `</section>`

    const wb = new exl.Workbook()

    const hsFile = './public/Краткий отчет.xlsx'
    const wbFile = './public/wildberries/new.xlsx'

    await wb.xlsx.readFile(hsFile)
        
    const ws = wb.getWorksheet('Краткий отчет')

    const c2 = ws.getColumn(2)

    c2.eachCell(c => {
        nat_cat.push(c.value)
    })

    await wb.xlsx.readFile(wbFile)

    const _ws = wb.getWorksheet('Сборочные задания')

    const c13 = _ws.getColumn(13)

    c13.eachCell({includeEmpty: false}, (c, rowNumber) => {

        if(rowNumber < 2) return

        if(wb_orders.findIndex(o => o.vendor === c.value) >= 0) {

            wb_orders.find(o => o.vendor === c.value).quantity++

        }

        if(wb_orders.findIndex(o => o.vendor === c.value) < 0) {

            wb_orders.push({
                'vendor': c.value,
                'quantity': 1
            })

        }

    })

    for(let i = 0; i < wb_orders.length; i++) {

        console.log(wb_orders[i].vendor)

        const response = await axios.post('https://api-seller.ozon.ru/v4/product/info/attributes', {
                    
            "filter": {
                "offer_id": [
                    wb_orders[i].vendor
                ],
                "visibility": "ALL"
            },
            "limit": 1000,
            "sort_dir": "ASC"

        }, {
            headers: {
                'Host':'api-seller.ozon.ru',
                'Client-Id':`${process.env.OZON_CLIENT_ID}`,
                'Api-Key':`${process.env.OZON_API_KEY}`,
                'Content-Type':'application/json'
            }
        })

        if(response.data.result[0].name.indexOf('Пододеяльник') >= 0) {

            names.push({            
                'vendor': wb_orders[i].vendor,
                'name': response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                .trim()                  // убрать пробелы по краям
                                                .replace(/\s+/g, ' '),
                'size': response.data.result[0].attributes.find(o => o.id === 6773).values[0].value,
                'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                'productType': 'ПОДОДЕЯЛЬНИК С КЛАПАНОМ'
            })   

        }

        if(response.data.result[0].name.indexOf('Простыня') >= 0 && response.data.result[0].name.indexOf('белье') < 0 && response.data.result[0].name.indexOf('бельё') < 0) {

            if(response.data.result[0].name.indexOf('на резинке') >= 0) {

                names.push({            
                    'vendor': wb_orders[i].vendor,
                    'name': response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                    .trim()                  // убрать пробелы по краям
                                                    .replace(/\s+/g, ' '),
                    'size': `${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}x${response.data.result[0].attributes.find(o => o.id === 8414).values[0].value}`,
                    'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                    'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                    'productType': 'ПРОСТЫНЯ НА РЕЗИНКЕ'
                })

            }

            if(response.data.result[0].name.indexOf('на резинке') < 0) {

                names.push({            
                    'vendor': wb_orders[i].vendor,
                    'name': response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                    .trim()                  // убрать пробелы по краям
                                                    .replace(/\s+/g, ' '),
                    'size': response.data.result[0].attributes.find(o => o.id === 6771).values[0].value,
                    'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                    'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                    'productType': 'ПРОСТЫНЯ'
                })

            }

        }

        if(response.data.result[0].name.indexOf('Наволочка') >= 0 || response.data.result[0].name.indexOf('наволочка') >= 0 && response.data.result[0].name.indexOf('белье') < 0 && response.data.result[0].name.indexOf('бельё') < 0) {

            if(response.data.result[0].name.indexOf('50х70') >= 0 || response.data.result[0].name.indexOf('40х60') >= 0 || response.data.result[0].name.indexOf('50 х 70') >= 0 || response.data.result[0].name.indexOf('40 х 60') >= 0 ) {

                names.push({
                    'vendor': wb_orders[i].vendor,
                    'name': response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                    .trim()                  // убрать пробелы по краям
                                                    .replace(/\s+/g, ' '),
                    'size': response.data.result[0].attributes.find(o => o.id === 6772).values[0].value,
                    'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                    'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                    'productType': 'НАВОЛОЧКА ПРЯМОУГОЛЬНАЯ'
                })

            } else {

                names.push({
                    'vendor': wb_orders[i].vendor,
                    'name': response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                    .trim()                  // убрать пробелы по краям
                                                    .replace(/\s+/g, ' '),
                    'size': response.data.result[0].attributes.find(o => o.id === 6772).values[0].value,
                    'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                    'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                    'productType': 'НАВОЛОЧКА КВАДРАТНАЯ'
                })

            }

        }

        if(response.data.result[0].name.indexOf('белье') >= 0 || response.data.result[0].name.indexOf('бельё') >= 0) {

            if(response.data.result[0].attributes.find(o => o.id === 6772).values.length === 2) {

                if(response.data.result[0].name.indexOf('на резинке') >= 0) {

                    if(response.data.result[0].name.indexOf('х20 -') >= 0 ||response.data.result[0].name.indexOf('х 20 -') >= 0) {

                        names.push({
                            'vendor': wb_orders[i].vendor,
                            'name': `КПБ ${response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                            .trim()                  // убрать пробелы по краям
                                                            .replace(/\s+/g, ' ')}`,
                            'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}x20; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[1].value}`,
                            'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                            'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                            'productType': 'КОМПЛЕКТ'
                        })

                    }

                    if(response.data.result[0].name.indexOf('х30 -') >= 0 ||response.data.result[0].name.indexOf('х 30 -') >= 0) {

                        names.push({
                            'vendor': wb_orders[i].vendor,
                            'name': `КПБ ${response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                            .trim()                  // убрать пробелы по краям
                                                            .replace(/\s+/g, ' ')}`,
                            'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}x30; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[1].value}`,
                            'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                            'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                            'productType': 'КОМПЛЕКТ'
                        })

                    }

                    if(response.data.result[0].name.indexOf('х40') >= 0 ||response.data.result[0].name.indexOf('х 40 -') >= 0) {

                        names.push({
                            'vendor': wb_orders[i].vendor,
                            'name': `КПБ ${response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                            .trim()                  // убрать пробелы по краям
                                                            .replace(/\s+/g, ' ')}`,
                            'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}x40; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[1].value}`,
                            'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                            'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                            'productType': 'КОМПЛЕКТ'
                        })

                    }

                }

                if(response.data.result[0].name.indexOf('на резинке') < 0) {

                    names.push({
                        'vendor': wb_orders[i].vendor,
                        'name': `КПБ ${response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                            .trim()                  // убрать пробелы по краям
                                                            .replace(/\s+/g, ' ')}`,
                        'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[1].value}`,
                        'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                        'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                        'productType': 'КОМПЛЕКТ'
                    })

                }

            }

            if(response.data.result[0].attributes.find(o => o.id === 6772).values.length === 1) {

                if(response.data.result[0].name.indexOf('на резинке') >= 0) {

                    if(response.data.result[0].name.indexOf('х20 -') >= 0 ||response.data.result[0].name.indexOf('х 20 -') >= 0) {

                        names.push({
                            'vendor': wb_orders[i].vendor,
                            'name': `КПБ ${response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                            .trim()                  // убрать пробелы по краям
                                                            .replace(/\s+/g, ' ')}`,
                            'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}x20; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}`,
                            'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                            'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                            'productType': 'КОМПЛЕКТ'
                        })

                    }

                    if(response.data.result[0].name.indexOf('х30 -') >= 0 ||response.data.result[0].name.indexOf('х 30 -') >= 0) {

                        names.push({
                            'vendor': wb_orders[i].vendor,
                            'name': `КПБ ${response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                            .trim()                  // убрать пробелы по краям
                                                            .replace(/\s+/g, ' ')}`,
                            'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}x30; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}`,
                            'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                            'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                            'productType': 'КОМПЛЕКТ'
                        })

                    }

                    if(response.data.result[0].name.indexOf('х40 -') >= 0 ||response.data.result[0].name.indexOf('х 40 -') >= 0) {

                        names.push({
                            'vendor': wb_orders[i].vendor,
                            'name': `КПБ ${response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                            .trim()                  // убрать пробелы по краям
                                                            .replace(/\s+/g, ' ')}`,
                            'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}x40; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}`,
                            'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                            'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                            'productType': 'КОМПЛЕКТ'
                        })

                    }

                }

                if(response.data.result[0].name.indexOf('на резинке') < 0) {

                    names.push({
                        'vendor': wb_orders[i].vendor,
                        'name': `КПБ ${response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                            .trim()                  // убрать пробелы по краям
                                                            .replace(/\s+/g, ' ')}`,
                        'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}`,
                        'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                        'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                        'productType': 'КОМПЛЕКТ'
                    })

                }

            }
            

        }

        names = names.filter(o => o.name.indexOf('Одеяло') < 0 && o.name.indexOf('Подушка') < 0 && o.name.indexOf('Матрас') < 0)

    }

    names.forEach(el => {

            if(nat_cat.indexOf(el.name) < 0) {
                new_items.push(el.name)
            }

            if(nat_cat.indexOf(el.name) >= 0) {
                current_items.push(el.name)
            }

    })

    html += `<section class="table">
                <div class="marks-table">
                    <div class="marks-table-header">
                        <div class="header-cell">Наименование</div>
                        <div class="header-cell">Статус</div>                            
                    </div>
                <div class="header-wrapper"></div>`

    names.forEach(elem => {
        if(new_items.indexOf(elem.name) >= 0) {
            html += `<div class="table-row">
                        <span id="name">${elem.name}</span>
                        <span id="status-new">Новый товар</span>
                     </div>`
        } else if(moderation_items.indexOf(elem.name) >= 0){
            html += `<div class="table-row">
                        <span id="name">${elem.name}</span>
                        <span id="status-moderation">На модерации</span>
                     </div>`        
        } else {
            html += `<div class="table-row">
                        <span id="name">${elem.name}</span>
                        <span id="status-current">Актуальный товар</span>
                     </div>`
        }
    })

    html += `</section>
             <section class="action-form">
                <button id="current-order"><a href="http://localhost:3030/wildberries_marks_order" target="_blank">Создать заказы маркировки</a></button>
             </section>
             <div class="body-wrapper"></div>                        
             ${footerComponent}`

    async function createImport(array) {

        const fileName = './public/IMPORT_TNVED_6302 (3).xlsx'
        
        const wb = new exl.Workbook()

        await wb.xlsx.readFile(fileName)

        const ws = wb.getWorksheet('IMPORT_TNVED_6302')

        let cellNumber = 5

        for(let i = 0; i < array.length; i++) {

            ws.getCell(`A${cellNumber}`).value = 6302
            ws.getCell(`B${cellNumber}`).value = names.find(o => o.name === array[i]).name
            ws.getCell(`C${cellNumber}`).value = 'Ивановский текстиль'
            ws.getCell(`D${cellNumber}`).value = 'Артикул'
            ws.getCell(`E${cellNumber}`).value = names.find(o => o.name === array[i]).vendor
            ws.getCell(`F${cellNumber}`).value = names.find(o => o.name === array[i]).productType
            ws.getCell(`G${cellNumber}`).value = names.find(o => o.name === array[i]).color
            ws.getCell(`H${cellNumber}`).value = 'ВЗРОСЛЫЙ'

            if(names.find(o => o.name === array[i]).cloth === 'КРЕП-ЖАТКА' || names.find(o => o.name === array[i]).cloth === 'КРЕП ЖАТКА') ws.getCell(`I${cellNumber}`).value = 'КРЕП'
            if(names.find(o => o.name === array[i]).cloth === 'ВАРЕНЫЙ ХЛОПОК') ws.getCell(`I${cellNumber}`).value = 'ХЛОПКОВАЯ ТКАНЬ'
            if(names.find(o => o.name === array[i]).cloth === 'ЛЕН' || names.find(o => o.name === array[i]).cloth === 'ЛЁН') ws.getCell(`I${cellNumber}`).value = 'ЛЬНЯНАЯ ТКАНЬ'
            if(names.find(o => o.name === array[i]).cloth === 'СТРАЙП САТИН') ws.getCell(`I${cellNumber}`).value = 'СТРАЙП-САТИН'
            if(names.find(o => o.name === array[i]).cloth === 'САТИН ЛЮКС') ws.getCell(`I${cellNumber}`).value = 'САТИН'
            if(names.find(o => o.name === array[i]).cloth !== 'САТИН ЛЮКС' && names.find(o => o.name === array[i]).cloth !== 'СТРАЙП САТИН' && names.find(o => o.name === array[i]).cloth !== 'ВАРЕНЫЙ ХЛОПОК' && names.find(o => o.name === array[i]).cloth !== 'ЛЕН' && names.find(o => o.name === array[i]).cloth !== 'ЛЁН') ws.getCell(`I${cellNumber}`).value = names.find(o => o.name === array[i]).cloth
            
            if(names.find(o => o.name === array[i]).cloth === 'ПОЛИСАТИН') ws.getCell(`J${cellNumber}`).value = '100% Полиэстер'

            if(names.find(o => o.name === array[i]).cloth === 'ТЕНСЕЛЬ') ws.getCell(`J${cellNumber}`).value = '100% Лиоцелл'
            if(names.find(o => o.name === array[i]).cloth === 'ЛЕН' || names.find(o => o.name === array[i]).cloth === 'ЛЁН') ws.getCell(`J${cellNumber}`).value = '100% Лен'
            if(names.find(o => o.name === array[i]).cloth !== 'КРЕП-ЖАТКА' && names.find(o => o.name === array[i]).cloth !== 'КРЕП ЖАТКА' && names.find(o => o.name === array[i]).cloth !== 'ПОЛИСАТИН' && names.find(o => o.name === array[i]).cloth !== 'ТЕНСЕЛЬ' && names.find(o => o.name === array[i]).cloth !== 'ЛЕН' && names.find(o => o.name === array[i]).cloth !== 'ЛЁН') ws.getCell(`J${cellNumber}`).value = '100% Хлопок'

            ws.getCell(`K${cellNumber}`).value = names.find(o => o.name === array[i]).size
            ws.getCell(`L${cellNumber}`).value = '6302100001'
            ws.getCell(`M${cellNumber}`).value = 'ТР ТС 017/2011 "О безопасности продукции легкой промышленности'
            ws.getCell(`N${cellNumber}`).value = 'На модерации'                

            cellNumber++

        }

        ws.unMergeCells('D2')

        ws.getCell('E2').value = '13914'

        ws.getCell('E2').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor:{argb:'E3E3E3'}
        }

        ws.getCell('E2').font = {
            size: 10,
            name: 'Arial'
        }

        ws.getCell('E2').alignment = {
            horizontal: 'center',
            vertical: 'bottom'
        }        

        const date_ob = new Date()

        let month = date_ob.getMonth() + 1

        let filePath = ''

        month < 10 ? filePath = `./public/wildberries/IMPORT_TNVED_6302_${date_ob.getDate()}_0${month}_wildberries` : filePath = `./public/wildberries/IMPORT_TNVED_6302_${date_ob.getDate()}_0${month}_wildberries`

        fs.access(`${filePath}.xlsx`, fs.constants.R_OK, async (err) => {
            if(err) {
                await wb.xlsx.writeFile(`${filePath}.xlsx`)
            } else {
                let count = 1
                fs.access(`${filePath}_(1).xlsx`, fs.constants.R_OK, async (err) => {
                    if(err) {
                        await wb.xlsx.writeFile(`${filePath}_(1).xlsx`)
                    } else {
                        await wb.xlsx.writeFile(`${filePath}_(2).xlsx`)
                    }
                })
                
            }
        })

    }

    if(new_items.length > 0) await createImport(new_items)

    res.send(html)

})

app.get('/wildberries_marks_order', async function(req, res) {

    const wb_orders = []
    const nat_cat = []
    const gtins = []
    let names = []

    html = `${headerComponent}
                    <title>Заказ маркировки - WILDBERRIES</title>
                </head>
                <body>
                    ${navComponent}
                        <section class="sub-nav import-main">
                            <div class="import-control">`

    await renderImportButtons(buttons)
    await renderMarkingButtons()
    await renderExtraButtons()

    html += `</section>`

    const wb = new exl.Workbook()

    const hsFile = './public/Краткий отчет.xlsx'
    const wbFile = './public/wildberries/new.xlsx'

    await wb.xlsx.readFile(hsFile)
        
    const ws = wb.getWorksheet('Краткий отчет')

    const c1 = ws.getColumn(1)

    c1.eachCell(c => {
        gtins.push(c.value)
    })

    const c2 = ws.getColumn(2)

    c2.eachCell(c => {
        nat_cat.push(c.value)
    })

    await wb.xlsx.readFile(wbFile)

    const _ws = wb.getWorksheet('Сборочные задания')

    const c13 = _ws.getColumn(13)

    c13.eachCell({includeEmpty: false}, (c, rowNumber) => {

        if(rowNumber < 2) return

        if(wb_orders.findIndex(o => o.vendor === c.value) >= 0) {

            wb_orders.find(o => o.vendor === c.value).quantity++

        }

        if(wb_orders.findIndex(o => o.vendor === c.value) < 0) {

            wb_orders.push({
                'vendor': c.value,
                'quantity': 1
            })

        }

    })

    for(let i = 0; i < wb_orders.length; i++) {

        const response = await axios.post('https://api-seller.ozon.ru/v4/product/info/attributes', {
                    
            "filter": {
                "offer_id": [
                    wb_orders[i].vendor
                ],
                "visibility": "ALL"
            },
            "limit": 1000,
            "sort_dir": "ASC"

        }, {
            headers: {
                'Host':'api-seller.ozon.ru',
                'Client-Id':`${process.env.OZON_CLIENT_ID}`,
                'Api-Key':`${process.env.OZON_API_KEY}`,
                'Content-Type':'application/json'
            }
        })

        if(response.data.result[0].name.indexOf('белье') >= 0 || response.data.result[0].name.indexOf('бельё') >= 0) {

            names.push({
                'name': `КПБ ${response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                .trim()                  // убрать пробелы по краям
                                                .replace(/\s+/g, ' ')}`,
                'vendor': wb_orders[i].vendor
            })    

        }

        if(response.data.result[0].name.indexOf('белье') < 0 && response.data.result[0].name.indexOf('бельё') < 0) {
            
            names.push({
                'name': response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                .trim()                  // убрать пробелы по краям
                                                .replace(/\s+/g, ' '),
                'vendor': wb_orders[i].vendor
            })

        }

        names = names.filter(o => o.name.indexOf('Одеяло') < 0 && o.name.indexOf('Подушка') && o.name.indexOf('Матрас') < 0 && o.name.indexOf('Ветошь') < 0 && o.name.indexOf('Наматрас') < 0 && o.name.indexOf('Плед') < 0)

    }

    function createNameList() {

        let orderList = []
        let _temp = []

        for (let i = 0; i < names.length; i++) {

                _temp.push(names[i].name)
            
                if(_temp.length%10 === 0) {
                    orderList.push(_temp)
                    _temp = []
                }
        }        

        if(_temp.length > 0) {
            orderList.push(_temp)
            _temp = []
        }

        return orderList

    }

    function createQuantityList() {

        let quantityList = []
        let temp = []

        for(let i = 0; i < names.length; i++) {
            
            if(nat_cat.indexOf(names[i].name) >= 0) {
                temp.push(wb_orders.find(o => o.vendor === names[i].vendor).quantity)
            }

            if(temp.length%10 === 0) {
                quantityList.splice(-1, 0, ...quantityList.splice(-1, 1, temp))
                temp = []
            }

        }

        quantityList.splice(-1, 0, ...quantityList.splice(-1, 1, temp))

        return quantityList

    }

    let List = createNameList()
    let Quantity = createQuantityList()

    function createOrder() {

        let content = ``

        for(let i = 0; i < List.length; i++) {
            if(List[i].length > 0) {
                content += `<?xml version="1.0" encoding="utf-8"?>
                            <order xmlns="urn:oms.order" xsi:schemaLocation="urn:oms.order schema.xsd" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
                                <lp>
                                    <productGroup>lp</productGroup>
                                    <contactPerson>333</contactPerson>
                                    <releaseMethodType>REMARK</releaseMethodType>
                                    <createMethodType>SELF_MADE</createMethodType>
                                    <productionOrderId>WB_${i+1}</productionOrderId>
                                    <products>`
                
                    for(let j = 0; j < List[i].length; j++) {                
                        if(nat_cat.indexOf(List[i][j]) >= 0) {
                            content += `<product>
                                            <gtin>0${gtins[nat_cat.indexOf(List[i][j])]}</gtin>
                                            <quantity>${Quantity[i][j]}</quantity>
                                            <serialNumberType>OPERATOR</serialNumberType>
                                            <cisType>UNIT</cisType>
                                            <templateId>10</templateId>
                                        </product>`
                        }
                    }
                
                content += `    </products>
                            </lp>
                        </order>`

            }

            const date_ob = new Date()

            let month = date_ob.getMonth() + 1

            let filePath = ''

            month < 10 ? filePath = `./public/orders/lp_wb_${i}_${date_ob.getDate()}_0${month}.xml` : filePath = `./public/orders/lp_wb_${i}_${date_ob.getDate()}_${month}.xml`

            if(content !== ``) {
                fs.writeFileSync(filePath, content)
            }

            content = ``

        }

        html += `<section class="table">
                <div class="marks-table">
                    <div class="marks-table-header">
                        <div class="header-cell">Наименование</div>
                        <div class="header-cell">Статус</div>
                        <div class="header-cell">Кол-во</div>
                    </div>
                <div class="header-wrapper"></div>`

        for(let i = 0; i < List.length; i++) {
            for(let j = 0; j < List[i].length; j++) {
                if(nat_cat.indexOf(List[i][j]) < 0) {
                    html += `<div class="table-row">
                                <span id="name">${List[i][j]}</span>
                                <span id="status-new">Новый товар</span>
                                <span id="quantity">${Quantity[i][j]}</span>
                             </div>`
                } else {
                    html += `<div class="table-row">
                                <span id="name">${List[i][j]}</span>
                                <span id="status-current">Актуальный товар</span>
                                <span id="quantity">${Quantity[i][j]}</span>
                             </div>`
                }
            }
        }

        html += `<section>
                <div class="body-wrapper"></div>
            ${footerComponent}`        

    }

    createOrder()

    res.send(html)

})

app.get('/wildberries/set_marks', async function (req, res){

    let wbOrder = []
    const nat_cat = []
    const ozon_cat = []
    const ozonCodes = []
    const ozonNames = []
    const gtins = []
    const marksGtins = []
    const marksCodes = []
    const marks = []
    const orderNumbers = []
    const orderCodes = []
    const marksOrderNumbers = []
    
    html = `${headerComponent}
                    <title>Подстановка маркировки Wildberries</title>
                </head>
                <body>
                        ${navComponent}
                        <section class="sub-nav import-main">
                            <div class="import-control">`

    await renderImportButtons(buttons)
    await renderMarkingButtons()
    await renderExtraButtons()

    html += `</section>`

    const marksFile = './public/wildberries/marks.xlsx'
    const ozonFile = './public/products.xlsx'
    const hsFile = './public/Краткий отчет.xlsx'
    const wbOrderFile = './public/wildberries/new.xlsx'
    const marksTemplateFile = './public/wildberries/marks_template.xlsx'

    const wb = new exl.Workbook()

    await wb.xlsx.readFile(wbOrderFile)

    const ws_1 = wb.getWorksheet('Сборочные задания')

    const w_c1 = ws_1.getColumn(1)

    const w_c13 = ws_1.getColumn(13)

    w_c1.eachCell({includeEmpty: false}, (c, rowNumber) => {

        if(rowNumber < 2) return

        orderNumbers.push(c.value)

    })

    w_c13.eachCell({includeEmpty: false}, (c, rowNumber) => {

        if(rowNumber < 2) return

        orderCodes.push(c.value)

    })

    for(let i = 0; i < orderNumbers.length; i++) {

        wbOrder.push({
            'orderNumber': orderNumbers[i],
            'orderCode': orderCodes[i]
        })

    }

    for(let i = 0; i < wbOrder.length; i++) {
    
        const response = await axios.post('https://api-seller.ozon.ru/v4/product/info/attributes', {
                    
            "filter": {
                "offer_id": [
                    wbOrder[i].orderCode
                ],
                "visibility": "ALL"
            },
            "limit": 1000,
            "sort_dir": "ASC"

        }, {
            headers: {
                'Host':'api-seller.ozon.ru',
                'Client-Id':'144225',
                'Api-Key':'52bf59da-6c76-4f26-b668-8704dfa71726',
                'Content-Type':'application/json'
            }
        })

        if(response.data.result[0].name.indexOf('белье') >= 0 || response.data.result[0].name.indexOf('бельё') >= 0) {

            wbOrder[i].orderProduct = `КПБ ${response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                              .trim()                  // убрать пробелы по краям
                                                              .replace(/\s+/g, ' ')}`

        }

        if(response.data.result[0].name.indexOf('белье') < 0 && response.data.result[0].name.indexOf('бельё') < 0) {
        
            wbOrder[i].orderProduct = response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                                .trim()                  // убрать пробелы по краям
                                                                .replace(/\s+/g, ' ')

        }

    }

    wbOrder = wbOrder.filter(o => o.orderProduct.indexOf('Матрас') < 0 && o.orderProduct.indexOf('Подушка') < 0 && o.orderProduct.indexOf('Одеяло') < 0 && o.orderProduct.indexOf('Ветошь') < 0 && o.orderProduct.indexOf('Наматра') < 0 && o.orderProduct.indexOf('Плед'))

    await wb.xlsx.readFile(marksTemplateFile)

    const ws_2 = wb.getWorksheet('Сборочные задания')

    const m_c1 = ws_2.getColumn(1)

    m_c1.eachCell({includeEmpty: false}, (c, rowNumber) => {

        if(rowNumber < 2) return
        
        marksOrderNumbers.push({
            'address': c.address.split(''),
            'value': c.value
        })

    })

    await wb.xlsx.readFile(hsFile)
        
    const ws_3 = wb.getWorksheet('Краткий отчет')

    const h_c1 = ws_3.getColumn(1)

    h_c1.eachCell({includeEmpty: false}, (c, rowNumber) => {
        if(rowNumber < 5) return
        gtins.push(c.value)
    })

    const h_c2 = ws_3.getColumn(2)

    h_c2.eachCell({includeEmpty: false}, (c, rowNumber) => {
        if(rowNumber < 5) return
        nat_cat.push(c.value)
    })

    await wb.xlsx.readFile(ozonFile)

    const ws_4 = wb.getWorksheet('products')

    const o_c1 = ws_4.getColumn(1)

    const o_c6 = ws_4.getColumn(6)

    o_c1.eachCell({includeEmpty: false}, (c, rowNumber) => {

        if(rowNumber < 2) return
        ozonCodes.push(c.value.replace("'", ""))

    })

    o_c6.eachCell({includeEmpty: false}, (c, rowNumber) => {

        if(rowNumber < 2) return
        ozonNames.push(c.value)

    })

    for(let i = 0; i < ozonCodes.length; i++) {

        ozon_cat.push({
            'code': ozonCodes[i],
            'name': ozonNames[i]
        })

    }

    let _temp = []

    wbOrder.forEach(el => {

        if(_temp.find(o => o.name === el.orderProduct) >= 0) return

        if(nat_cat.indexOf(el.orderProduct) < 0) {

            const item = ozon_cat.find(o => o.code === el.orderCode)
            _temp.push({

                'name': el.orderProduct,
                'gtin': gtins[nat_cat.indexOf(item.name)]

            })

        }
        
        if(nat_cat.indexOf(el.orderProduct) >= 0) {

            _temp.push({

                'name': el.orderProduct,
                'gtin': gtins[nat_cat.indexOf(el.orderProduct)]
                
            })

        }

    })

    _temp = _temp.filter(o => o.name.indexOf('Матрас') < 0 && o.name.indexOf('Подушка') < 0 && o.name.indexOf('Одеяло') < 0 && o.name.indexOf('Ветошь') < 0)

    await wb.xlsx.readFile(marksFile)

    const ws_5 = wb.getWorksheet('Sheet0')

    const ma_c1 = ws_5.getColumn(1)

    ma_c1.eachCell({includeEmpty: false}, (c, rowNumber) => {

        if(rowNumber < 3) return
        marksCodes.push(c.value.replace(/\u003C/g, '<').replace(/\u003E/g, '>'))

    })

    const ma_c2 = ws_5.getColumn(2)

    ma_c2.eachCell({includeEmpty: false}, (c, rowNumber) => {

        if(rowNumber < 3) return
        marksGtins.push(c.value.slice(1))

    })

    for(let i = 0; i < marksCodes.length; i++) {

        marks.push({
            'mark': marksCodes[i],
            'gtin': marksGtins[i],
            'status': 'not_used'
        })

    }

    for(let i = 0; i < wbOrder.length; i++) {

        const gtin = _temp.find(o => o.name === wbOrder[i].orderProduct).gtin

        const mark = marks.find(o => o.gtin === String(gtin) && o.status === 'not_used').mark

        if(mark) {

            wbOrder[i].mark = mark
            marks.find(o => o.gtin === String(gtin) && o.status === 'not_used').status = 'used'

        } else {
            return
        }

    }

    html += `<section class="table">
                <div class="marks-table">
                    <div class="marks-table-header">
                        <div class="header-cell">Номер заказа WB</div>
                        <div class="header-cell">Наименование</div>
                        <div class="header-cell">Артикул</div>
                        <div class="header-cell">Код маркировки</div>
                    </div>
                <div class="header-wrapper"></div>`

    for(let i = 0; i < wbOrder.length; i++) {

        html += `<div class="table-row">
                    <span id="order">${wbOrder[i].orderNumber}</span>
                    <span id="product">${wbOrder[i].orderProduct}</span>
                    <span id="vendor">${wbOrder[i].orderCode}</span>
                    <span id="mark">${wbOrder[i].mark.replace(/\u003C/g, '&lt').replace(/\u003E/g, '&gt').replace(/\"/g, '&quot')}</span>
                </div>`

    }

    html += `</section>
             <div class="body-wrapper"></div>                        
             ${footerComponent}`

    const wb_1 = new exl.Workbook()

    await wb_1.xlsx.readFile(marksTemplateFile)

    const ws_6 = wb_1.getWorksheet('Сборочные задания')

    for(let i = 0; i < wbOrder.length; i++) {

        const order = marksOrderNumbers.find(o => o.value === wbOrder[i].orderNumber)
        if(order.address.length === 2) {
            ws_6.getCell(`C${order.address[1]}`).value = wbOrder[i].mark
        }

        if(order.address.length === 3) {
            ws_6.getCell(`C${order.address[1]}${order.address[2]}`).value = wbOrder[i].mark
        }

        if(order.address.length === 4) {
            ws_6.getCell(`C${order.address[1]}${order.address[2]}${order.address[3]}`).value = wbOrder[i].mark
        }

    }

    await wb_1.xlsx.writeFile(`marks_template_completed.xlsx`)

    res.send(html)

    // res.json({wbOrder, marks, marksOrderNumbers})
    
})

app.get('/yandex', async function(req, res){

    const nat_cat = []
    const nat_catGtins = []
    const nat_catNames = []
    let ya_orders = []
    const new_items = []
    const current_items = []
    let names = []

    html = `${headerComponent}
                    <title>Импорт - Я.Маркет</title>
                </head>
                <body>
                        ${navComponent}
                        <section class="sub-nav import-main">
                            <div class="import-control">`

    await renderImportButtons(buttons)
    await renderMarkingButtons()
    await renderExtraButtons()

    html += `</section>`

    const wb = new exl.Workbook()

    await wb.xlsx.readFile('./public/Краткий отчет.xlsx')

    const nc_ws = wb.getWorksheet('Краткий отчет')

    const nc_c1 = nc_ws.getColumn(1)

    nc_c1.eachCell({includeEmpty: false}, (c, rowNumber) => {

        if(rowNumber < 5) return
        nat_catGtins.push(c.value)

    })

    const nc_c2 = nc_ws.getColumn(2)

    nc_c2.eachCell({includeEmpty: false}, (c, rowNumber) => {

        if(rowNumber < 5) return
        nat_catNames.push(c.value.trim())

    })

    for(let i = 0; i < nat_catNames.length; i++) {

        nat_cat.push({
            'gtin': nat_catGtins[i],
            'name': nat_catNames[i]
        })

    }

    async function getOrders(clientId) {

        let response = await axios.get(`https://api.partner.market.yandex.ru/campaigns/${clientId}/orders?status=PROCESSING&substatus=STARTED&pageSize=50`, {
            headers: {
                'Authorization': `Bearer ${process.env.YANDEX_API_KEY}`
            }
        })

        let result = response.data

        let currentPage = result.pager.currentPage

        let _response = await axios.get(`https://api.partner.market.yandex.ru/campaigns/${clientId}/orders?status=PROCESSING&substatus=STARTED&page=2`, {
            headers: {
                'Authorization': `Bearer ${process.env.YANDEX_API_KEY}`
            }
        })

        let _result = _response.data

        _result.orders.forEach(elem => {

            elem.items.forEach(el => {
                
                if(el.requiredInstanceTypes) {
                    if(el.requiredInstanceTypes.indexOf('CIS_OPTIONAL') >= 0) {

                        if(el.instances === undefined) {

                            ya_orders.push({

                                'name': el.offerName,
                                'vendor': el.offerId

                            })

                        }

                    }
                }

            })

        })

        let pageTotal = Math.ceil(result.pager.total / 50)

        result.orders.forEach(elem => {

            elem.items.forEach(el => {
                
                if(el.requiredInstanceTypes) {
                    if(el.requiredInstanceTypes.indexOf('CIS_OPTIONAL') >= 0) {

                        if(el.instances === undefined) {

                            ya_orders.push({

                                'name': el.offerName,
                                'vendor': el.offerId

                            })

                        }

                    }
                }

            })
        })

    }

    // await getOrders(fbsId)
    await getOrders(dbsId)

    for(let i = 0; i < ya_orders.length; i++) {

        const response = await axios.post('https://api-seller.ozon.ru/v4/product/info/attributes', {
                    
            "filter": {
                "offer_id": [
                    ya_orders[i].vendor
                ],
                "visibility": "ALL"
            },
            "limit": 1000,
            "sort_dir": "ASC"

        }, {
            headers: {
                'Host':'api-seller.ozon.ru',
                'Client-Id':`${process.env.OZON_CLIENT_ID}`,
                'Api-Key':`${process.env.OZON_API_KEY}`,
                'Content-Type':'application/json'
            }
        })

        if(response.data.result[0].name.indexOf('Пододеяльник') >= 0) {

            names.push({            
                'vendor': ya_orders[i].vendor,
                'name': response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                .trim()                  // убрать пробелы по краям
                                                .replace(/\s+/g, ' '),
                'size': response.data.result[0].attributes.find(o => o.id === 6773).values[0].value,
                'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                'productType': 'ПОДОДЕЯЛЬНИК С КЛАПАНОМ'
            })   

        }

        if(response.data.result[0].name.indexOf('Простыня') >= 0 && response.data.result[0].name.indexOf('белье') < 0 && response.data.result[0].name.indexOf('бельё') < 0) {

            if(response.data.result[0].name.indexOf('на резинке') >= 0) {

                names.push({            
                    'vendor': ya_orders[i].vendor,
                    'name': response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                    .trim()                  // убрать пробелы по краям
                                                    .replace(/\s+/g, ' '),
                    'size': `${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}x${response.data.result[0].attributes.find(o => o.id === 8414).values[0].value}`,
                    'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                    'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                    'productType': 'ПРОСТЫНЯ НА РЕЗИНКЕ'
                })

            }

            if(response.data.result[0].name.indexOf('на резинке') < 0) {

                names.push({            
                    'vendor': ya_orders[i].vendor,
                    'name': response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                    .trim()                  // убрать пробелы по краям
                                                    .replace(/\s+/g, ' '),
                    'size': response.data.result[0].attributes.find(o => o.id === 6771).values[0].value,
                    'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                    'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                    'productType': 'ПРОСТЫНЯ'
                })

            }

        }

        if(response.data.result[0].name.indexOf('Наволочка') >= 0 || response.data.result[0].name.indexOf('наволочка') >= 0 && response.data.result[0].name.indexOf('белье') < 0 && response.data.result[0].name.indexOf('бельё') < 0) {

            if(response.data.result[0].name.indexOf('50х70') >= 0 || response.data.result[0].name.indexOf('40х60') >= 0 || response.data.result[0].name.indexOf('50 х 70') >= 0 || response.data.result[0].name.indexOf('40 х 60') >= 0 ) {

                names.push({
                    'vendor': ya_orders[i].vendor,
                    'name': response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                    .trim()                  // убрать пробелы по краям
                                                    .replace(/\s+/g, ' '),
                    'size': response.data.result[0].attributes.find(o => o.id === 6772).values[0].value,
                    'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                    'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                    'productType': 'НАВОЛОЧКА ПРЯМОУГОЛЬНАЯ'
                })

            } else {

                names.push({
                    'vendor': ya_orders[i].vendor,
                    'name': response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                    .trim()                  // убрать пробелы по краям
                                                    .replace(/\s+/g, ' '),
                    'size': response.data.result[0].attributes.find(o => o.id === 6772).values[0].value,
                    'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                    'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                    'productType': 'НАВОЛОЧКА КВАДРАТНАЯ'
                })

            }

        }

        if(response.data.result[0].name.indexOf('белье') >= 0 || response.data.result[0].name.indexOf('бельё') >= 0) {

            if(response.data.result[0].attributes.find(o => o.id === 6772).values.length === 2) {

                if(response.data.result[0].name.indexOf('на резинке') >= 0) {

                    if(response.data.result[0].name.indexOf('х20 -') >= 0 ||response.data.result[0].name.indexOf('х 20 -') >= 0) {

                        names.push({
                            'vendor': ya_orders[i].vendor,
                            'name': `КПБ ${response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                            .trim()                  // убрать пробелы по краям
                                                            .replace(/\s+/g, ' ')}`,
                            'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}x20; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[1].value}`,
                            'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                            'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                            'productType': 'КОМПЛЕКТ'
                        })

                    }

                    if(response.data.result[0].name.indexOf('х30 -') >= 0 ||response.data.result[0].name.indexOf('х 30 -') >= 0) {

                        names.push({
                            'vendor': ya_orders[i].vendor,
                            'name': `КПБ ${response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                            .trim()                  // убрать пробелы по краям
                                                            .replace(/\s+/g, ' ')}`,
                            'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}x30; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[1].value}`,
                            'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                            'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                            'productType': 'КОМПЛЕКТ'
                        })

                    }

                    if(response.data.result[0].name.indexOf('х40') >= 0 ||response.data.result[0].name.indexOf('х 40 -') >= 0) {

                        names.push({
                            'vendor': ya_orders[i].vendor,
                            'name': `КПБ ${response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                            .trim()                  // убрать пробелы по краям
                                                            .replace(/\s+/g, ' ')}`,
                            'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}x40; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[1].value}`,
                            'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                            'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                            'productType': 'КОМПЛЕКТ'
                        })

                    }

                }

                if(response.data.result[0].name.indexOf('на резинке') < 0) {

                    names.push({
                        'vendor': ya_orders[i].vendor,
                        'name': `КПБ ${response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                            .trim()                  // убрать пробелы по краям
                                                            .replace(/\s+/g, ' ')}`,
                        'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[1].value}`,
                        'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                        'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                        'productType': 'КОМПЛЕКТ'
                    })

                }

            }

            if(response.data.result[0].attributes.find(o => o.id === 6772).values.length === 1) {

                if(response.data.result[0].name.indexOf('на резинке') >= 0) {

                    if(response.data.result[0].name.indexOf('х20 -') >= 0 ||response.data.result[0].name.indexOf('х 20 -') >= 0) {

                        names.push({
                            'vendor': ya_orders[i].vendor,
                            'name': `КПБ ${response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                            .trim()                  // убрать пробелы по краям
                                                            .replace(/\s+/g, ' ')}`,
                            'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}x20; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}`,
                            'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                            'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                            'productType': 'КОМПЛЕКТ'
                        })

                    }

                    if(response.data.result[0].name.indexOf('х30 -') >= 0 ||response.data.result[0].name.indexOf('х 30 -') >= 0) {

                        names.push({
                            'vendor': ya_orders[i].vendor,
                            'name': `КПБ ${response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                            .trim()                  // убрать пробелы по краям
                                                            .replace(/\s+/g, ' ')}`,
                            'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}x30; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}`,
                            'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                            'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                            'productType': 'КОМПЛЕКТ'
                        })

                    }

                    if(response.data.result[0].name.indexOf('х40 -') >= 0 ||response.data.result[0].name.indexOf('х 40 -') >= 0) {

                        names.push({
                            'vendor': ya_orders[i].vendor,
                            'name': `КПБ ${response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                            .trim()                  // убрать пробелы по краям
                                                            .replace(/\s+/g, ' ')}`,
                            'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}x40; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}`,
                            'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                            'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                            'productType': 'КОМПЛЕКТ'
                        })

                    }

                }

                if(response.data.result[0].name.indexOf('на резинке') < 0) {

                    names.push({
                        'vendor': ya_orders[i].vendor,
                        'name': `КПБ ${response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                            .trim()                  // убрать пробелы по краям
                                                            .replace(/\s+/g, ' ')}`,
                        'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}`,
                        'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                        'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                        'productType': 'КОМПЛЕКТ'
                    })

                }

            }
            

        }

    }

    names.forEach(el => {

            if(nat_cat.findIndex(o => o.name === el.name) < 0) {
                new_items.push(el.name)
            }

            if(nat_cat.findIndex(o => o.name === el.name) >= 0) {
                current_items.push(el.name)
            }

    })

    html += `<section class="table">
                <div class="marks-table">
                    <div class="marks-table-header">
                        <div class="header-cell">Наименование</div>
                        <div class="header-cell">Статус</div>                            
                    </div>
                <div class="header-wrapper"></div>`

    names.forEach(elem => {
        if(new_items.indexOf(elem.name) >= 0) {
            html += `<div class="table-row">
                        <span id="name">${elem.name}</span>
                        <span id="status-new">Новый товар</span>
                     </div>`
        } else {
            html += `<div class="table-row">
                        <span id="name">${elem.name}</span>
                        <span id="status-current">Актуальный товар</span>
                     </div>`
        }
    })

    html += `</section>
             <section class="action-form">
                <button id="current-order"><a href="http://localhost:3030/yandex_marks_order" target="_blank">Создать заказы маркировки</a></button>
             </section>
             <div class="body-wrapper"></div>                        
             ${footerComponent}`

    async function createImport(array) {

        const fileName = './public/IMPORT_TNVED_6302 (3).xlsx'
        
        const wb = new exl.Workbook()

        await wb.xlsx.readFile(fileName)

        const ws = wb.getWorksheet('IMPORT_TNVED_6302')

        let cellNumber = 5

        for(let i = 0; i < array.length; i++) {

            ws.getCell(`A${cellNumber}`).value = 6302
            ws.getCell(`B${cellNumber}`).value = names.find(o => o.name === array[i]).name
            ws.getCell(`C${cellNumber}`).value = 'Ивановский текстиль'
            ws.getCell(`D${cellNumber}`).value = 'Артикул'
            ws.getCell(`E${cellNumber}`).value = names.find(o => o.name === array[i]).vendor
            ws.getCell(`F${cellNumber}`).value = names.find(o => o.name === array[i]).productType
            ws.getCell(`G${cellNumber}`).value = names.find(o => o.name === array[i]).color
            ws.getCell(`H${cellNumber}`).value = 'ВЗРОСЛЫЙ'

            if(names.find(o => o.name === array[i]).cloth === 'КРЕП-ЖАТКА' || names.find(o => o.name === array[i]).cloth === 'КРЕП ЖАТКА') ws.getCell(`I${cellNumber}`).value = 'КРЕП'
            if(names.find(o => o.name === array[i]).cloth === 'ВАРЕНЫЙ ХЛОПОК') ws.getCell(`I${cellNumber}`).value = 'ХЛОПКОВАЯ ТКАНЬ'
            if(names.find(o => o.name === array[i]).cloth === 'ЛЕН' || names.find(o => o.name === array[i]).cloth === 'ЛЁН') ws.getCell(`I${cellNumber}`).value = 'ЛЬНЯНАЯ ТКАНЬ'
            if(names.find(o => o.name === array[i]).cloth === 'СТРАЙП САТИН') ws.getCell(`I${cellNumber}`).value = 'СТРАЙП-САТИН'
            if(names.find(o => o.name === array[i]).cloth === 'САТИН ЛЮКС') ws.getCell(`I${cellNumber}`).value = 'САТИН'
            if(names.find(o => o.name === array[i]).cloth !== 'САТИН ЛЮКС' && names.find(o => o.name === array[i]).cloth !== 'СТРАЙП САТИН' && names.find(o => o.name === array[i]).cloth !== 'ВАРЕНЫЙ ХЛОПОК' && names.find(o => o.name === array[i]).cloth !== 'ЛЕН' && names.find(o => o.name === array[i]).cloth !== 'ЛЁН') ws.getCell(`I${cellNumber}`).value = names.find(o => o.name === array[i]).cloth
            
            if(names.find(o => o.name === array[i]).cloth === 'ПОЛИСАТИН') ws.getCell(`J${cellNumber}`).value = '100% Полиэстер'

            if(names.find(o => o.name === array[i]).cloth === 'ТЕНСЕЛЬ') ws.getCell(`J${cellNumber}`).value = '100% Лиоцелл'
            if(names.find(o => o.name === array[i]).cloth === 'ЛЕН' || names.find(o => o.name === array[i]).cloth === 'ЛЁН') ws.getCell(`J${cellNumber}`).value = '100% Лен'
            if(names.find(o => o.name === array[i]).cloth !== 'КРЕП-ЖАТКА' && names.find(o => o.name === array[i]).cloth !== 'КРЕП ЖАТКА' && names.find(o => o.name === array[i]).cloth !== 'ПОЛИСАТИН' && names.find(o => o.name === array[i]).cloth !== 'ТЕНСЕЛЬ' && names.find(o => o.name === array[i]).cloth !== 'ЛЕН' && names.find(o => o.name === array[i]).cloth !== 'ЛЁН') ws.getCell(`J${cellNumber}`).value = '100% Хлопок'

            ws.getCell(`K${cellNumber}`).value = names.find(o => o.name === array[i]).size
            ws.getCell(`L${cellNumber}`).value = '6302100001'
            ws.getCell(`M${cellNumber}`).value = 'ТР ТС 017/2011 "О безопасности продукции легкой промышленности'
            ws.getCell(`N${cellNumber}`).value = 'На модерации'                

            cellNumber++

        }

        ws.unMergeCells('D2')

        ws.getCell('E2').value = '13914'

        ws.getCell('E2').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor:{argb:'E3E3E3'}
        }

        ws.getCell('E2').font = {
            size: 10,
            name: 'Arial'
        }

        ws.getCell('E2').alignment = {
            horizontal: 'center',
            vertical: 'bottom'
        }        

        const date_ob = new Date()

        let month = date_ob.getMonth() + 1

        let filePath = ''

        month < 10 ? filePath = `./public/yandex/IMPORT_TNVED_6302_${date_ob.getDate()}_0${month}_yandex` : filePath = `./public/yandex/IMPORT_TNVED_6302_${date_ob.getDate()}_0${month}_yandex`

        fs.access(`${filePath}.xlsx`, fs.constants.R_OK, async (err) => {
            if(err) {
                await wb.xlsx.writeFile(`${filePath}.xlsx`)
            } else {
                let count = 1
                fs.access(`${filePath}_(1).xlsx`, fs.constants.R_OK, async (err) => {
                    if(err) {
                        await wb.xlsx.writeFile(`${filePath}_(1).xlsx`)
                    } else {
                        await wb.xlsx.writeFile(`${filePath}_(2).xlsx`)
                    }
                })
                
            }
        })

    }

    if(new_items.length > 0) await createImport(new_items)

    res.send(html)

})

app.get('/yandex_marks_order', async function (req, res){
    
    let ya_orders = []
    const nat_cat = []
    const gtins = []
    let names = []

    html = `${headerComponent}
                    <title>Заказ маркировки - Я.Маркет</title>
                </head>
                <body>
                    ${navComponent}
                        <section class="sub-nav import-main">
                            <div class="import-control">`

    await renderImportButtons(buttons)
    await renderMarkingButtons()
    await renderExtraButtons()

    html += `</section>`

    const wb = new exl.Workbook()
    
    const fileName = './public/Краткий отчет.xlsx'    

    await wb.xlsx.readFile(fileName)
        
    const nc_ws = wb.getWorksheet('Краткий отчет')

    const nc_c1 = nc_ws.getColumn(1)

    nc_c1.eachCell({includeEmpty: false}, (c, rowNumber) => {

        if(rowNumber < 5) return
        gtins.push(c.value)

    })

    const nc_c2 = nc_ws.getColumn(2)

    nc_c2.eachCell({includeEmpty: false}, (c, rowNumber) => {

        if(rowNumber < 5) return
        nat_cat.push(c.value.trim())

    })

    async function getOrders(clientId) {

        let response = await axios.get(`https://api.partner.market.yandex.ru/campaigns/${clientId}/orders?status=PROCESSING&substatus=STARTED&pageSize=50`, {
            headers: {
                'Authorization': `Bearer ${process.env.YANDEX_API_KEY}`
            }
        })

        let result = response.data

        let currentPage = result.pager.currentPage

        let _response = await axios.get(`https://api.partner.market.yandex.ru/campaigns/${clientId}/orders?status=PROCESSING&substatus=STARTED&page=2`, {
            headers: {
                'Authorization': `Bearer ${process.env.YANDEX_API_KEY}`
            }
        })

        let _result = _response.data

        _result.orders.forEach(elem => {

            elem.items.forEach(el => {
                
                if(el.requiredInstanceTypes) {
                    if(el.requiredInstanceTypes.indexOf('CIS_OPTIONAL') >= 0) {

                        if(el.instances === undefined) {

                            if(ya_orders.findIndex(o => o.name === el.offerName) >= 0) {

                                ya_orders.find(o => o.name === el.offerName).quantity += Number(el.count)

                            }

                            if(ya_orders.findIndex(o => o.name === el.offerName) < 0) {

                                if(el.offerName.indexOf('белье') >= 0 || el.offerName.indexOf('бельё') >= 0) {

                                    ya_orders.push({

                                        'name': `КПБ ${el.offerName}`,
                                        'vendor': el.offerId,
                                        'quantity': el.count

                                    })

                                }

                                if(el.offerName.indexOf('белье') < 0 && el.offerName.indexOf('бельё') < 0) {

                                    ya_orders.push({

                                        'name': el.offerName,
                                        'vendor': el.offerId,
                                        'quantity': el.count

                                    })

                                }

                            }
                        }

                    }
                }

            })

        })

        let pageTotal = Math.ceil(result.pager.total / 50)

        result.orders.forEach(elem => {

            elem.items.forEach(el => {
                
                if(el.requiredInstanceTypes) {
                    if(el.requiredInstanceTypes.indexOf('CIS_OPTIONAL') >= 0) {

                        if(el.instances === undefined) {

                            if(ya_orders.findIndex(o => o.name === el.offerName) >= 0) {

                                ya_orders.find(o => o.name === el.offerName).quantity += Number(el.count)

                            }

                            if(ya_orders.findIndex(o => o.name === el.offerName) < 0) {

                                if(el.offerName.indexOf('белье') >= 0 || el.offerName.indexOf('бельё') >= 0) {

                                    ya_orders.push({

                                        'name': `КПБ ${el.offerName}`,
                                        'vendor': el.offerId,
                                        'quantity': el.count

                                    })

                                }

                                if(el.offerName.indexOf('белье') < 0 && el.offerName.indexOf('бельё') < 0) {

                                    ya_orders.push({

                                        'name': el.offerName,
                                        'vendor': el.offerId,
                                        'quantity': el.count

                                    })

                                }

                            }
                        }

                    }
                }

            })
        })

    }

    // await getOrders(fbsId)
    await getOrders(dbsId)

    ya_orders = ya_orders.filter(o => o.name.indexOf('Одеяло') < 0 && o.name.indexOf('Подушка') < 0 && o.name.indexOf('Матрас') < 0 && o.name.indexOf('Ветошь') < 0)

    html += `<section class="table">
                    <div class="marks-table">
                        <div class="marks-table-header">
                            <div class="header-cell">Наименование</div>
                            <div class="header-cell">Статус</div>                            
                        </div>
                    <div class="header-wrapper"></div>`

    for(let i = 0; i < ya_orders.length; i++) {
        html += `<div class="table-row">
                    <span id="name">${ya_orders[i].name}</span>
                    <span id="status-current">Актуальный товар</span>
                    <span id="quantity">${ya_orders[i].quantity}</span>
                 </div>`
    }

    html += `</section>
             <div class="body-wrapper"></div>                        
             ${footerComponent}`

    function createNameList() {

        let orderList = []
        let _temp = []

        for (let i = 0; i < ya_orders.length; i++) {

            if(nat_cat.indexOf(ya_orders[i].name) >= 0) {

                _temp.push(ya_orders[i].name)

            }            
            
            if(_temp.length%10 === 0) {
                if(_temp.length !== 0) {
                    orderList.push(_temp)
                }
                _temp = []
            }

        }        

        if(_temp.length !== 0) {
            orderList.push(_temp)
        }
        _temp = []

        return orderList

    }

    function createQuantityList() {

        let quantityList = []
        let temp = []

        for(let i = 0; i < ya_orders.length; i++) {

            if(nat_cat.indexOf(ya_orders[i].name) >= 0) {

                temp.push(ya_orders[i].quantity)

            }            

            if(temp.length%10 === 0) {
                if(temp.length !== 0) {
                    quantityList.splice(-1, 0, ...quantityList.splice(-1, 1, temp))
                }
                temp = []
            }

        }

        if(temp.length !== 0) {

            quantityList.splice(-1, 0, ...quantityList.splice(-1, 1, temp))

        }

        return quantityList

    }

    function createOrder() {

        let List = createNameList()
        let Quantity = createQuantityList()
        let content = ``

        for(let i = 0; i < List.length; i++) {
            if(List[i].length > 0) {
                content += `<?xml version="1.0" encoding="utf-8"?>
                            <order xmlns="urn:oms.order" xsi:schemaLocation="urn:oms.order schema.xsd" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
                                <lp>
                                    <productGroup>lp</productGroup>
                                    <contactPerson>333</contactPerson>
                                    <releaseMethodType>REMARK</releaseMethodType>
                                    <createMethodType>SELF_MADE</createMethodType>
                                    <productionOrderId>YANDEX_${i+1}</productionOrderId>
                                    <products>`
                
                    for(let j = 0; j < List[i].length; j++) {                
                        if(nat_cat.indexOf(List[i][j]) >= 0) {
                            content += `<product>
                                            <gtin>0${gtins[nat_cat.indexOf(List[i][j])]}</gtin>
                                            <quantity>${Quantity[i][j]}</quantity>
                                            <serialNumberType>OPERATOR</serialNumberType>
                                            <cisType>UNIT</cisType>
                                            <templateId>10</templateId>
                                        </product>`
                        }
                    }

                content += `    </products>
                            </lp>
                        </order>`

            }

            const date_ob = new Date()

            let month = date_ob.getMonth() + 1

            let filePath = ''

            month < 10 ? filePath = `./public/orders/lp_yandex_${i}_${date_ob.getDate()}_0${month}.xml` : filePath = `./public/orders/lp_yandex_${i}_${date_ob.getDate()}_${month}.xml`

            if(content !== ``) {
                fs.writeFileSync(filePath, content)
            }

            content = ``

        }

    }

    createOrder()

    res.send(html)

    // res.json(ya_orders)

})

app.get('/stocks', async function(req, res){

    const nat_cat = []
    const ncGtins = []
    let ncNames = []
    let wh_prod = []
    const wh_code = []
    let wh = []
    const fullGtins = []
    const fullNames = []
    const fullVendors = []
    const full_cat = []
    const new_items = []
    const errorCodes = []
    let full_difference = []
    const full_matches = []
    const names = []

    html = `${headerComponent}
                    <title>Маркировка остатков склада</title>
                </head>
                <body>
                    ${navComponent}
                        <section class="sub-nav import-main">
                            <div class="import-control">`

    await renderImportButtons(buttons)
    await renderMarkingButtons()
    await renderExtraButtons()

    html += `</section>`

    const wb = new exl.Workbook()

    const hsFile = './public/Краткий отчет.xlsx'
    const wProductsFile = './public/warehouse_products.xlsx'
    const fCatalogFile = './public/Выгрузка 372900043349.xlsx'

    await wb.xlsx.readFile(hsFile)
        
    const nc_ws1 = wb.getWorksheet('Краткий отчет')

    const [nc_c1, nc_c2] = [nc_ws1.getColumn(1), nc_ws1.getColumn(2)]

    nc_c1.eachCell({includeEmpty: false}, (c, rowNumber) => {

        if(rowNumber < 5) return
        ncGtins.push(c.value.trim())

    })

    nc_c2.eachCell({includeEmpty: false}, (c, rowNumber) => {

        if(rowNumber < 5) return
        ncNames.push(c.value.trim())

    })

    for(let i = 0; i < ncGtins.length; i++) {

        if(nat_cat.findIndex(o => o.name === ncNames[i]) >= 0) {

            nat_cat.find(o => o.name === ncNames[i]).gtin.push(ncGtins[i])

        }

        if(nat_cat.findIndex(o => o.name === ncNames[i]) < 0) {

            nat_cat.push({
                'name': ncNames[i],
                'gtin': [ncGtins[i]]
            })

        }

    }

    await wb.xlsx.readFile(wProductsFile)

    const wh_ws2 = wb.getWorksheet('Лист1')

    const [wh_c1, wh_c2] = [wh_ws2.getColumn(1), wh_ws2.getColumn(2)]

    wh_c1.eachCell({includeEmpty: false}, (c, rowNumber) => {

        if(rowNumber < 2) return
        wh_prod.push(c.value.trim())

    })

    wh_c2.eachCell({includeEmpty: false}, (c, rowNumber) => {

        if(rowNumber < 2) return
        wh_code.push(c.value.trim())

    })

    for(let i = 0; i < wh_code.length; i++) {

        try {

        const response = await axios.post('https://api-seller.ozon.ru/v4/product/info/attributes', {
                        
            "filter": {
                "offer_id": [
                    wh_code[i]
                ],
                "visibility": "ALL"
            },
            "limit": 1000,
            "sort_dir": "ASC"

        }, {

            headers: {
                'Host':'api-seller.ozon.ru',
                'Client-Id':`${process.env.OZON_CLIENT_ID}`,
                'Api-Key':`${process.env.OZON_API_KEY}`,
                'Content-Type':'application/json'
            }

        })

        wh.push({
            'name': response.data.result[0].name,
            'vendor': wh_code[i]
        })

        } catch(err) {

            console.log(err.response.data.message)

            errorCodes.push({
                'name': wh_prod[i],
                'vendor': wh_code[i]
            })

        }

    }

    for(let i = 0; i < nat_cat.length; i++) {

        if(nat_cat[i].name.indexOf('- Р ') >= 0) {

            nat_cat[i].name = nat_cat[i].name.replace('- Р ', '')

        }

    }

    await wb.xlsx.readFile(fCatalogFile)

    const f_ws3 = wb.getWorksheet('result')

    const [f_c4, f_c8, f_c10] = [f_ws3.getColumn(4), f_ws3.getColumn(8), f_ws3.getColumn(10)]

    f_c4.eachCell({includeEmpty: false}, (c, rowNumber) => {

        if(rowNumber < 2) return
        fullGtins.push(c.value.trim())

    })

    f_c8.eachCell({includeEmpty: false}, (c, rowNumber) => {

        if(rowNumber < 2) return
        fullNames.push(c.value.trim())

    })

    f_c10.eachCell({includeEmpty: false}, (c, rowNumber) => {

        if(rowNumber < 2) return
        fullVendors.push(c.value.trim())

    })

    for(let i = 0; i < fullNames.length; i++) {

        if(full_cat.findIndex(o => o.name === fullNames[i]) >= 0) {

            full_cat.find(o => o.name === fullNames[i]).gtin.push(fullGtins[i])

        }

        if(full_cat.findIndex(o => o.name === fullNames[i]) < 0) {

            full_cat.push({

                'name': fullNames[i].replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                    .trim()                  // убрать пробелы по краям
                                    .replace(/\s+/g, ' '),
                'gtin': [fullGtins[i]],
                'vendor': fullVendors[i]

            })

        }

    }

    for(let i = 0; i < wh.length; i++) {

        if(full_cat.findIndex(o => o.vendor === wh[i].vendor ) < 0 && nat_cat.findIndex(o => o.name === wh[i].name) < 0) {

            full_difference.push(wh[i])

        }

        if(full_cat.findIndex(o => o.vendor === wh[i].vendor) >= 0 || nat_cat.findIndex(o => o.name === wh[i].name) >= 0) {

            full_matches.push(wh[i])

        }

    }

    full_difference = full_difference.filter(o => o.name.indexOf('Наматрасник') < 0)

    for(let i = 0; i < full_difference.length; i++) {

        try {

            const response = await axios.post('https://api-seller.ozon.ru/v4/product/info/attributes', {
                        
                "filter": {
                    "offer_id": [
                        full_difference[i].vendor
                    ],
                    "visibility": "ALL"
                },
                "limit": 1000,
                "sort_dir": "ASC"

            }, {
                headers: {
                    'Host':'api-seller.ozon.ru',
                    'Client-Id':`${process.env.OZON_CLIENT_ID}`,
                    'Api-Key':`${process.env.OZON_API_KEY}`,
                    'Content-Type':'application/json'
                }
            })

            if(response.data.result[0].name.indexOf('Пододеяльник') >= 0) {

                names.push({            
                    'vendor': full_difference[i].vendor,
                    'name': response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                    .trim()                  // убрать пробелы по краям
                                                    .replace(/\s+/g, ' ')
                                                    .replace('- Р ', ''),
                    'size': response.data.result[0].attributes.find(o => o.id === 6773).values[0].value,
                    'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                    'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                    'productType': 'ПОДОДЕЯЛЬНИК С КЛАПАНОМ'
                })   

            }

            if(response.data.result[0].name.indexOf('Простыня') >= 0 && response.data.result[0].name.indexOf('белье') < 0 && response.data.result[0].name.indexOf('бельё') < 0) {

                if(response.data.result[0].name.indexOf('на резинке') >= 0) {

                    names.push({            
                        'vendor': full_difference[i].vendor,
                        'name': response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                        .trim()                  // убрать пробелы по краям
                                                        .replace(/\s+/g, ' ')
                                                        .replace('- Р ', ''),
                        'size': `${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}x${response.data.result[0].attributes.find(o => o.id === 8414).values[0].value}`,
                        'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                        'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                        'productType': 'ПРОСТЫНЯ НА РЕЗИНКЕ'
                    })

                }

                if(response.data.result[0].name.indexOf('на резинке') < 0) {

                    names.push({            
                        'vendor': full_difference[i].vendor,
                        'name': response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                        .trim()                  // убрать пробелы по краям
                                                        .replace(/\s+/g, ' ')
                                                        .replace('- Р ', ''),
                        'size': response.data.result[0].attributes.find(o => o.id === 6771).values[0].value,
                        'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                        'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                        'productType': 'ПРОСТЫНЯ'
                    })

                }

            }

            if(response.data.result[0].name.indexOf('Наволочка') >= 0 || response.data.result[0].name.indexOf('наволочка') >= 0 && response.data.result[0].name.indexOf('белье') < 0 && response.data.result[0].name.indexOf('бельё') < 0) {

                if(response.data.result[0].name.indexOf('50х70') >= 0 || response.data.result[0].name.indexOf('40х60') >= 0 || response.data.result[0].name.indexOf('50 х 70') >= 0 || response.data.result[0].name.indexOf('40 х 60') >= 0 ) {

                    names.push({
                        'vendor': full_difference[i].vendor,
                        'name': response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                        .trim()                  // убрать пробелы по краям
                                                        .replace(/\s+/g, ' ')
                                                        .replace('- Р ', ''),
                        'size': response.data.result[0].attributes.find(o => o.id === 6772).values[0].value,
                        'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                        'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                        'productType': 'НАВОЛОЧКА ПРЯМОУГОЛЬНАЯ'
                    })

                } else {

                    names.push({
                        'vendor': full_difference[i].vendor,
                        'name': response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                        .trim()                  // убрать пробелы по краям
                                                        .replace(/\s+/g, ' ')
                                                        .replace('- Р ', ''),
                        'size': response.data.result[0].attributes.find(o => o.id === 6772).values[0].value,
                        'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                        'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                        'productType': 'НАВОЛОЧКА КВАДРАТНАЯ'
                    })

                }

            }

            if(response.data.result[0].name.indexOf('белье') >= 0 || response.data.result[0].name.indexOf('бельё') >= 0) {

                if(response.data.result[0].attributes.find(o => o.id === 6772).values.length === 2) {

                    if(response.data.result[0].name.indexOf('на резинке') >= 0) {

                        if(response.data.result[0].name.indexOf('х20 -') >= 0 ||response.data.result[0].name.indexOf('х 20 -') >= 0) {

                            names.push({
                                'vendor': full_difference[i].vendor,
                                'name': response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                                .trim()                  // убрать пробелы по краям
                                                                .replace(/\s+/g, ' ')
                                                                .replace('- Р ', ''),
                                'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}x20; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[1].value}`,
                                'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                                'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                                'productType': 'КОМПЛЕКТ'
                            })

                        }

                        if(response.data.result[0].name.indexOf('х30 -') >= 0 ||response.data.result[0].name.indexOf('х 30 -') >= 0) {

                            names.push({
                                'vendor': full_difference[i].vendor,
                                'name': response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                                .trim()                  // убрать пробелы по краям
                                                                .replace(/\s+/g, ' ')
                                                                .replace('- Р ', ''),
                                'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}x30; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[1].value}`,
                                'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                                'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                                'productType': 'КОМПЛЕКТ'
                            })

                        }

                        if(response.data.result[0].name.indexOf('х40') >= 0 ||response.data.result[0].name.indexOf('х 40 -') >= 0) {

                            names.push({
                                'vendor': full_difference[i].vendor,
                                'name': response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                                .trim()                  // убрать пробелы по краям
                                                                .replace(/\s+/g, ' ')
                                                                .replace('- Р ', ''),
                                'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}x40; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[1].value}`,
                                'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                                'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                                'productType': 'КОМПЛЕКТ'
                            })

                        }

                    }

                    if(response.data.result[0].name.indexOf('на резинке') < 0) {

                        names.push({
                            'vendor': full_difference[i].vendor,
                            'name': response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                            .trim()                  // убрать пробелы по краям
                                                            .replace(/\s+/g, ' ')
                                                            .replace('- Р ', ''),
                            'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[1].value}`,
                            'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                            'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                            'productType': 'КОМПЛЕКТ'
                        })

                    }

                }

                if(response.data.result[0].attributes.find(o => o.id === 6772).values.length === 1) {

                    if(response.data.result[0].name.indexOf('на резинке') >= 0) {

                        if(response.data.result[0].name.indexOf('х20 -') >= 0 ||response.data.result[0].name.indexOf('х 20 -') >= 0) {

                            names.push({
                                'vendor': full_difference[i].vendor,
                                'name': response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                                .trim()                  // убрать пробелы по краям
                                                                .replace(/\s+/g, ' ')
                                                                .replace('- Р ', ''),
                                'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}x20; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}`,
                                'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                                'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                                'productType': 'КОМПЛЕКТ'
                            })

                        }

                        if(response.data.result[0].name.indexOf('х30 -') >= 0 ||response.data.result[0].name.indexOf('х 30 -') >= 0) {

                            names.push({
                                'vendor': full_difference[i].vendor,
                                'name': response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                                .trim()                  // убрать пробелы по краям
                                                                .replace(/\s+/g, ' ')
                                                                .replace('- Р ', ''),
                                'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}x30; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}`,
                                'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                                'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                                'productType': 'КОМПЛЕКТ'
                            })

                        }

                        if(response.data.result[0].name.indexOf('х40 -') >= 0 ||response.data.result[0].name.indexOf('х 40 -') >= 0) {

                            names.push({
                                'vendor': full_difference[i].vendor,
                                'name': response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                                .trim()                  // убрать пробелы по краям
                                                                .replace(/\s+/g, ' ')
                                                                .replace('- Р ', ''),
                                'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}x40; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}`,
                                'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                                'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                                'productType': 'КОМПЛЕКТ'
                            })

                        }

                    }

                    if(response.data.result[0].name.indexOf('на резинке') < 0) {

                        names.push({
                            'vendor': full_difference[i].vendor,
                            'name': response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                            .trim()                  // убрать пробелы по краям
                                                            .replace(/\s+/g, ' ')
                                                            .replace('- Р ', ''),
                            'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}`,
                            'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                            'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                            'productType': 'КОМПЛЕКТ'
                        })

                    }

                }
                

            }

            new_items.push(response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                       .trim()                  // убрать пробелы по краям
                                                       .replace(/\s+/g, ' ')
                                                       .replace('- Р ', ''),)

        } catch(err) {

            errorCodes.push(full_difference[i])

        }

    }

    async function createImport(array) {

        const fileName = './public/IMPORT_TNVED_6302 (3).xlsx'
        
        const wb = new exl.Workbook()

        await wb.xlsx.readFile(fileName)

        const ws = wb.getWorksheet('IMPORT_TNVED_6302')

        let cellNumber = 5

        for(let i = 0; i < array.length; i++) {

            ws.getCell(`A${cellNumber}`).value = 6302
            ws.getCell(`B${cellNumber}`).value = names.find(o => o.name === array[i]).name
            ws.getCell(`C${cellNumber}`).value = 'Ивановский текстиль'
            ws.getCell(`D${cellNumber}`).value = 'Артикул'
            ws.getCell(`E${cellNumber}`).value = names.find(o => o.name === array[i]).vendor
            ws.getCell(`F${cellNumber}`).value = names.find(o => o.name === array[i]).productType
            ws.getCell(`G${cellNumber}`).value = names.find(o => o.name === array[i]).color
            ws.getCell(`H${cellNumber}`).value = 'ВЗРОСЛЫЙ'

            if(names.find(o => o.name === array[i]).cloth === 'КРЕП-ЖАТКА' || names.find(o => o.name === array[i]).cloth === 'КРЕП ЖАТКА') ws.getCell(`I${cellNumber}`).value = 'КРЕП'
            if(names.find(o => o.name === array[i]).cloth === 'ВАРЕНЫЙ ХЛОПОК') ws.getCell(`I${cellNumber}`).value = 'ХЛОПКОВАЯ ТКАНЬ'
            if(names.find(o => o.name === array[i]).cloth === 'ЛЕН' || names.find(o => o.name === array[i]).cloth === 'ЛЁН') ws.getCell(`I${cellNumber}`).value = 'ЛЬНЯНАЯ ТКАНЬ'
            if(names.find(o => o.name === array[i]).cloth === 'СТРАЙП САТИН') ws.getCell(`I${cellNumber}`).value = 'СТРАЙП-САТИН'
            if(names.find(o => o.name === array[i]).cloth === 'САТИН ЛЮКС') ws.getCell(`I${cellNumber}`).value = 'САТИН'
            if(names.find(o => o.name === array[i]).cloth !== 'САТИН ЛЮКС' && names.find(o => o.name === array[i]).cloth !== 'СТРАЙП САТИН' && names.find(o => o.name === array[i]).cloth !== 'ВАРЕНЫЙ ХЛОПОК' && names.find(o => o.name === array[i]).cloth !== 'ЛЕН' && names.find(o => o.name === array[i]).cloth !== 'ЛЁН') ws.getCell(`I${cellNumber}`).value = names.find(o => o.name === array[i]).cloth
            
            if(names.find(o => o.name === array[i]).cloth === 'ПОЛИСАТИН') ws.getCell(`J${cellNumber}`).value = '100% Полиэстер'

            if(names.find(o => o.name === array[i]).cloth === 'ТЕНСЕЛЬ') ws.getCell(`J${cellNumber}`).value = '100% Лиоцелл'
            if(names.find(o => o.name === array[i]).cloth === 'ЛЕН' || names.find(o => o.name === array[i]).cloth === 'ЛЁН') ws.getCell(`J${cellNumber}`).value = '100% Лен'
            if(names.find(o => o.name === array[i]).cloth !== 'КРЕП-ЖАТКА' && names.find(o => o.name === array[i]).cloth !== 'КРЕП ЖАТКА' && names.find(o => o.name === array[i]).cloth !== 'ПОЛИСАТИН' && names.find(o => o.name === array[i]).cloth !== 'ТЕНСЕЛЬ' && names.find(o => o.name === array[i]).cloth !== 'ЛЕН' && names.find(o => o.name === array[i]).cloth !== 'ЛЁН') ws.getCell(`J${cellNumber}`).value = '100% Хлопок'

            ws.getCell(`K${cellNumber}`).value = names.find(o => o.name === array[i]).size
            ws.getCell(`L${cellNumber}`).value = '6302100001'
            ws.getCell(`M${cellNumber}`).value = 'ТР ТС 017/2011 "О безопасности продукции легкой промышленности'
            ws.getCell(`N${cellNumber}`).value = 'На модерации'                

            cellNumber++

        }

        ws.unMergeCells('D2')

        ws.getCell('E2').value = '13914'

        ws.getCell('E2').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor:{argb:'E3E3E3'}
        }

        ws.getCell('E2').font = {
            size: 10,
            name: 'Arial'
        }

        ws.getCell('E2').alignment = {
            horizontal: 'center',
            vertical: 'bottom'
        }        

        const date_ob = new Date()

        let month = date_ob.getMonth() + 1

        let filePath = ''

        month < 10 ? filePath = `./public/stocks/IMPORT_TNVED_6302_${date_ob.getDate()}_0${month}_stocks` : filePath = `./public/stocks/IMPORT_TNVED_6302_${date_ob.getDate()}_0${month}_stocks`

        fs.access(`${filePath}.xlsx`, fs.constants.R_OK, async (err) => {
            if(err) {
                await wb.xlsx.writeFile(`${filePath}.xlsx`)
            } else {
                let count = 1
                fs.access(`${filePath}_(1).xlsx`, fs.constants.R_OK, async (err) => {
                    if(err) {
                        await wb.xlsx.writeFile(`${filePath}_(1).xlsx`)
                    } else {
                        await wb.xlsx.writeFile(`${filePath}_(2).xlsx`)
                    }
                })
                
            }
        })

    }

    async function createReport(array) {

        const workbook = new exl.Workbook()

        const sheet = workbook.addWorksheet('Отчет')

        sheet.columns = [
            {header: 'Наименование', key: 'title', width: 100},
            {header: 'Артикул', key: 'code', width: 20}
        ]

        for(let i = 0; i < array.length; i++) {

            sheet.addRow({title: array[i].name, code: array[i].vendor})

        }

        await workbook.xlsx.writeFile('./public/Отчет_new.xlsx')

    }

    if(new_items.length > 0) await createImport(new_items)

    if(errorCodes.length > 0) await createReport(errorCodes)

    html += `<section class="table">
                <div class="marks-table">
                    <div class="marks-table-header">
                        <div class="header-cell">Наименование</div>
                        <div class="header-cell">Артикул</div>
                        <div class="header-cell">Статус</div>                            
                    </div>
                <div class="header-wrapper"></div>`

    for(let i = 0; i < full_difference.length; i++) {

        html += `<div class="table-row">
                    <span id="name">${full_difference[i].name}</span>
                    <span id="vendor">${full_difference[i].vendor}</span>
                    <span id="status-new">Новый</span>
                </div>`

    }

    for(let i = 0; i < full_matches.length; i++) {

        html += `<div class="table-row">
                    <span id="name">${full_matches[i].name}</span>
                    <span id="vendor">${full_matches[i].vendor}</span>
                    <span id="status-current">Актуальный</span>
                </div>`

    }

    html += `</section>
             <div class="body-wrapper"></div>                        
             ${footerComponent}`

    console.log(new_items.length)

    res.send(html)

})

app.get('/input_remarking', async function(req, res){

    html = `${headerComponent}
                    <title>Перемаркировка</title>
                </head>
                <body>
                    ${navComponent}
                        <section class="sub-nav import-main">
                            <div class="import-control">`

    await renderImportButtons(buttons)
    await renderMarkingButtons()
    await renderExtraButtons()

    html += `</section>`

    let remark_date = ''

    const date_ob = new Date()

    let year = date_ob.getFullYear()

    let month = date_ob.getMonth()+1

    let day = date_ob.getDate()

    month < 10 ? month = '0' + month : month

    day < 10 ? day = '0' + day : day

    remark_date = year + '-' + month + '-' + day    

    let content = `<?xml version="1.0" encoding="UTF-8"?>
                    <remark version="7">
                        <trade_participant_inn>372900043349</trade_participant_inn>
                        <remark_date>${remark_date}</remark_date>
                        <remark_cause>KM_SPOILED</remark_cause>
                            <products_list>`    

    const marks = []

    const wb = new exl.Workbook()

    await wb.xlsx.readFile('./public/inputinsale/marks.xlsx')

    const ws = wb.getWorksheet(1)

    ws.getColumn(1).eachCell(el => {
        marks.push(el.value.trim())
    })

    marks.forEach(el => {
        if(el.length === 31) {
            content += `<product>
                            <new_ki><![CDATA[${el}]]></new_ki>
                            <tnved_code_10>6302100001</tnved_code_10>
                            <production_country>РОССИЯ</production_country>
                        </product>`
        }
    })

    // console.log(content)   

    content += `    </products_list>
            </remark>`

    fs.writeFileSync('./public/inputinsale/remarking.xml', content)

    html += `<div class="result">Файл remarking.xml успешно сформирован</div>
                <section class="table">
                    <div class="marks-table">
                        <div class="marks-table-header">
                            <div class="header-cell">КИЗ</div>
                            <div class="header-cell">Код ТНВЭД</div>
                            <div class="header-cell">Страна</div>
                        </div>
                        <div class="header-wrapper"></div>`

    marks.forEach(el => {
        if(el.length === 31) {
            html += `<div class="table-row">
                        <span type="text" id="mark">${el.replace(/</g, '&lt;')}</span>
                        <span id="name">6302100001</span>
                        <span id="name">РОССИЯ</span>
                     </div>`
        }
    })
    

    html += `   </div>
            </section>
        ${footerComponent}`

    res.send(html)
    
})

app.get('/sale_ozon', async function(req, res){

    const actualMarksFile = './public/actual_marks.xlsx'

    html = `${headerComponent}
                    <title>Перемаркировка</title>
                </head>
                <body>
                    ${navComponent}
                        <section class="sub-nav import-main">
                            <div class="import-control">`

    await renderImportButtons(buttons)
    await renderMarkingButtons()
    await renderExtraButtons()

    html += `</section>`

    const date_ob = new Date()

    let orders = []
    let consignments = []

    let date_string = ''

    let [year, month, day] = [date_ob.getFullYear(), date_ob.getMonth()+1, date_ob.getDate()]

    month < 10 ? month = '0' + month : month
    day < 10 ? day = '0' + day : day

    date_string = `${year}-${month}-${day}`

    let content = `<?xml version="1.0" encoding="utf-8"?>
                    <withdrawal version="8">
                        <trade_participant_inn>372900043349</trade_participant_inn>
                        <withdrawal_type>DISTANCE</withdrawal_type>
                        <withdrawal_date>${date_string}</withdrawal_date>
                        <primary_document_type>CONSIGNMENT_NOTE</primary_document_type>
                        <products_list>`
    
    const wb = new exl.Workbook()

    async function getActualList() {

        const [marks, status] = [[], []]

        await wb.xlsx.readFile(actualMarksFile)

        const ws = wb.getWorksheet('Worksheet')

        const [c1, c16] = [ws.getColumn(1), ws.getColumn(16)]

        c1.eachCell(c => {
            marks.push(c.value)
        })

        c16.eachCell(c => {
            status.push(c.value)
        })

        const introduced_marks = []
        
        marks.forEach(e => {
            if(status[marks.indexOf(e)] == 'INTRODUCED') {
                introduced_marks.push(e)
            }
        })

    }

    await getActualList()

    //получаем данные из xlsx файла с реализациями и
    //формируем массив объектов реализаций

    async function getConsignments() {

        let consignments = []

        const consignmentDate = []

        const consignmentNumbers = []        

        const consignmentTypes = []

        const filePath = './public/distance/релизации.xlsx'

        await wb.xlsx.readFile(filePath)

        const ws = wb.getWorksheet('TDSheet')

        const [c2, c3, c8] = [ws.getColumn(2), ws.getColumn(3), ws.getColumn(8)]
        
        c2.eachCell(c => {
            let str = c.value
            consignmentDate.push(str.replace(str.substring(10), ''))
        })

        c3.eachCell(c => {
            let str = c.value
            consignmentNumbers.push(str.replace('MT00-', ''))
        })

        c8.eachCell(c => {
            consignmentTypes.push(c.value)
        })

        let noRepeatConsignmentTypes = []

        for(let i = 0; i < consignmentTypes.length; i++) {
            if(consignmentTypes[i] != null && consignmentTypes[i].indexOf('ozon') >= 0 && noRepeatConsignmentTypes.indexOf(consignmentTypes[i]) < 0) {
                noRepeatConsignmentTypes.push(consignmentTypes[i])
            }
        }

        for(let i = 0; i < consignmentDate.length; i++) {
            let _tempArray = consignmentDate[i].split('.')
            let str = `${_tempArray[2]}-${_tempArray[1]}-${_tempArray[0]}`
            consignmentDate[i] = str
        }

        for(let i = 0; i < noRepeatConsignmentTypes.length; i++) {
            consignments.push({
                orderNumber: noRepeatConsignmentTypes[i].substring(5),
                consignmentNumber: consignmentNumbers[consignmentTypes.indexOf(noRepeatConsignmentTypes[i])],
                consignmentDate: consignmentDate[consignmentTypes.indexOf(noRepeatConsignmentTypes[i])]
            })
        }

        return consignments

    }

    consignments = await getConsignments()

    async function getOrders() {

        let orders = []

        let response = await fetch('https://api-seller.ozon.ru/v3/posting/fbs/list', {
            method: 'POST',
            headers: {
                'Host': 'api-seller.ozon.ru',
                'Client-Id': '144225',
                'Api-Key': 'c139ba7b-611a-4447-870c-f85d8e4ad9f8',
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                'dir': 'asc',
                'filter':{
                    'since':'2023-08-01T00:00:00Z',
                    'to':'2023-10-25T00:00:00Z',
                    'status':'cancelled'
                },
                'limit':1000,
                'offset':0
            })
        })
        
        let result = await response.json()
        
        result.result.postings.forEach(e => {
            let orderNumber = e.posting_number
            let products = []
            e.products.forEach(el => {
                let marks = []
                el.mandatory_mark.forEach(elem => {
                    marks.push(elem)
                })
                products.push({
                    name: el.name,
                    marksList: marks,
                    price: el.price
                })
            })

            let obj = {
                orderNumber: orderNumber,
                productsList: products
            }

            orders.push(obj)
            
        })

        return orders

    }

    orders = await getOrders()

    orders.forEach(e => {
        console.log(e.orderNumber)
    })

    let equals = []

    for(let i = 0; i < orders.length; i++) {
        for(let j = 0; j < consignments.length; j++) {
            if(orders[i].orderNumber == consignments[j].orderNumber) {
                equals.push(orders[i])
            }
        }
    }

    equals.forEach(e => {
        e.productsList.forEach(el => {
            if(el.marksList.length > 0) {
                if(el.marksList.indexOf('') < 0) {
                    for(let i = 0; i < el.marksList.length; i++) {
                        content += `<product>
                                        <cis><![CDATA[${el.marksList[i]}]]></cis>
                                        <cost>${(el.price).replace(el.price.substring(el.price.indexOf('.')), '')}00</cost>
                                        <primary_document_type>CONSIGNMENT_NOTE</primary_document_type>
                                        <primary_document_number>${(consignments.find(c => c.orderNumber == e.orderNumber)).consignmentNumber}</primary_document_number>
                                        <primary_document_date>${(consignments.find(c => c.orderNumber == e.orderNumber)).consignmentDate}</primary_document_date>
                                    </product>`
                    }
                }
            }
        })
    })

    content += `</products_list>
            </withdrawal>`

    const fileName = `./public/distance/ozon_distance_${date_string}.xml`
    
    fs.writeFileSync(fileName, content)

    html += `<div class=result>Файл ${fileName.substring(fileName.lastIndexOf('/') + 1)} успешно сформирован</div>
            <section class="table">
                <div class="marks-table">
                    <div class="marks-table-header">
                        <div class="header-cell">КИЗ</div>
                        <div class="header-cell">Цена</div>
                        <div class="header-cell">Тип документа</div>
                        <div class="header-cell">Номер документа</div>
                        <div class="header-cell">Дата документа</div>
                    </div>
                    <div class="header-wrapper"></div>`

    equals.forEach(e => {
        e.productsList.forEach(el => {
            if(el.marksList.length > 0) {
                if(el.marksList.indexOf('') < 0) {
                    for(let i = 0; i < el.marksList.length; i++) {
                        // console.log(el.marksList[i])
                        html += `<div class="table-row">
                                    <span type="text" id="mark">${el.marksList[i].replace(/</g, '&lt;')}</span>
                                    <span id="gtin">${(el.price).replace(el.price.substring(el.price.indexOf('.')), '')}00</span>
                                    <span id="name">CONSIGNMENT_NOTE</span>
                                    <span id="status">${(consignments.find(c => c.orderNumber == e.orderNumber)).consignmentNumber}</span>
                                    <span id="date">${(consignments.find(c => c.orderNumber == e.orderNumber)).consignmentDate}</span>
                                 </div>`
                    }
                }
            }
        })
    })

    html += `       </div>
                </section>
            ${footerComponent}`

    res.send(html)

})

app.get('/sale_wb', async function(req, res){

    html = `${headerComponent}
                    <title>Перемаркировка</title>
                </head>
                <body>
                    ${navComponent}
                        <section class="sub-nav import-main">
                            <div class="import-control">`

    await renderImportButtons(buttons)
    await renderMarkingButtons()
    await renderExtraButtons()

    html += `</section>`

    const wbordersPath = './public/distance/wb_orders.xlsx'
    const consignmentsPath = './public/distance/релизации.xlsx'

    const date_ob = new Date()

    let date_string = ''

    let [year, month, day] = [date_ob.getFullYear(), date_ob.getMonth()+1, date_ob.getDate()]

    month < 10 ? month = '0' + month : month
    day < 10 ? day = '0' + day : day

    date_string = `${year}-${month}-${day}`

    let content = `<?xml version="1.0" encoding="utf-8"?>
                    <withdrawal version="8">
                        <trade_participant_inn>372900043349</trade_participant_inn>
                        <withdrawal_type>DISTANCE</withdrawal_type>
                        <withdrawal_date>${date_string}</withdrawal_date>
                        <primary_document_type>CONSIGNMENT_NOTE</primary_document_type>
                        <products_list>`
    
    // let response = await fetch('https://suppliers-api.wildberries.ru/api/v3/orders?limit=10&next=0&dateFrom=1687755600&dateTo=1688187600',{
    //     method: 'GET',
    //     headers: {
    //         'Authorization':'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJhY2Nlc3NJRCI6IjBhYmMxZWNmLTlmOWEtNDQzNi04YmNiLTM3Mjg1ZDJkYzJlZCJ9.-OGN5Jvwsf9XQHYy7LPPJjATV98xOSBXQMISSkjVNCg'
    //     }
    // })

    // let result = await response.json()

    // console.log(result.orders.forEach(e => {
    //     e.offices.forEach(el => {
    //         console.log(el)
    //     })
    // }))

    const wb = new exl.Workbook()

    let orders = []
    let consignments = []

    async function getOrders() {

        await wb.xlsx.readFile(wbordersPath)

        const ws = wb.getWorksheet('КИЗ')

        const orders = []

        const [orderNumbers, orderCises, orderPrices] = [[], [], []]

        const [c1, c3, c5] = [ws.getColumn(1), ws.getColumn(3), ws.getColumn(5)]
    
        c1.eachCell(c => {
            orderNumbers.push(c.value)
        })
    
        c3.eachCell(c => {
            orderCises.push(c.value)
        })
    
        c5.eachCell(c => {
            orderPrices.push(c.value)
        })

        for(let i = 0; i < orderNumbers.length; i++) {
            let obj = {
                orderNumber: orderNumbers[i],
                orderCis: orderCises[i],
                orderPrice: orderPrices[i]
            }

            orders.push(obj)
        }

        return orders

    }

    async function getConsignments() {

        await wb.xlsx.readFile(consignmentsPath)

        const ws = wb.getWorksheet('TDSheet')

        const [c2, c3, c8] = [ws.getColumn(2), ws.getColumn(3), ws.getColumn(8)]

        const [consDates, consNumbers, orderNumbers, wbNumbers] = [[], [], [], [], []]

        const numbers = []

        const consignments = []

        c2.eachCell(c => {
            let str = c.value.replace(c.value.substring(10), '')
            let date = str.split('.')
            consDates.push(`${date[2]}-${date[1]}-${date[0]}`)
        })

        c3.eachCell(c => {
            consNumbers.push(c.value.slice(c.value.indexOf('-')+1))
        })

        c8.eachCell(c => {
            numbers.push(c.value)
            if(c.value != null) {
                wbNumbers.push(c.value)
                orderNumbers.push(c.value.substring(3))
            }
        })

        

        for(let i = 0; i < orderNumbers.length; i++) {

            let obj = {
                consDate: consDates[numbers.indexOf(wbNumbers[i])],
                consNumber: consNumbers[numbers.indexOf(wbNumbers[i])],
                orderNumber: orderNumbers[i]
            }

            consignments.push(obj)

        }

        // console.log(consignments)
        return consignments

    }

    orders = await getOrders()
    consignments = await getConsignments()

    // console.log(orders)
    // console.log(consignments)

    let equals = []

    for(let i = 0; i < orders.length; i++) {
        let index = consignments.indexOf(consignments.find(c => c.orderNumber == orders[i].orderNumber))
        if(index >= 0) {
            equals.push({
                consignmentNumber: consignments[index].consNumber,
                consignmentDate: consignments[index].consDate,
                consignmentPrice: orders[i].orderPrice,
                consignmentCis: orders[i].orderCis
            })
        }
    }

    for(let i = 0; i < equals.length; i++) {
        let price = ''

        if((equals[i].consignmentPrice.toString()).indexOf('.') >= 0) {
            let arr = (equals[i].consignmentPrice.toString()).split('.')
            price = arr[0]+arr[1]
        } else {
            price = equals[i].consignmentPrice + '00'
        }

        content += `<product>
                        <cis><![CDATA[${equals[i].consignmentCis}]]></cis>
                        <cost>${price}</cost>
                        <primary_document_type>CONSIGNMENT_NOTE</primary_document_type>
                        <primary_document_number>${equals[i].consignmentNumber}</primary_document_number>
                        <primary_document_date>${equals[i].consignmentDate}</primary_document_date>
                    </product>`
        
    }
                            
    content += `</products_list>
            </withdrawal>`
    
    const fileName = `./public/distance/wb_distance_${date_string}.xml`
    
    fs.writeFileSync(fileName, content)

    html += `<div class=result>Файл ${fileName.substring(fileName.lastIndexOf('/') + 1)} успешно сформирован</div>
            <section class="table">
                <div class="marks-table">
                    <div class="marks-table-header">
                        <div class="header-cell">КИЗ</div>
                        <div class="header-cell">Цена</div>
                        <div class="header-cell">Тип документа</div>
                        <div class="header-cell">Номер документа</div>
                        <div class="header-cell">Дата документа</div>
                    </div>
                    <div class="header-wrapper"></div>`

    for(let i = 0; i < equals.length; i++) {

        let price = ''

        if((equals[i].consignmentPrice.toString()).indexOf('.') >= 0) {
            let arr = (equals[i].consignmentPrice.toString()).split('.')
            price = arr[0]+arr[1]
        } else {
            price = equals[i].consignmentPrice + '00'
        }

        html += `<div class="table-row">
                    <span type="text" id="mark">${equals[i].consignmentCis.replace(/</g, '&lt;')}</span>
                    <span id="gtin">${price}</span>
                    <span id="name">CONSIGNMENT_NOTE</span>
                    <span id="status">${equals[i].consignmentNumber}</span>
                    <span id="date">${equals[i].consignmentDate}</span>
                </div>`

    }

    html += `           </div>
                    </section>
                <div class="body-wrapper"></div>
            ${footerComponent}`

    res.send(html)

})

app.get('/test_features', async function(req, res){

    const [products, actual_products, gtins, actual_gtins] = [[], [], [], []]

    const newProducts = []

    const filePath = './public/actual_marks.html'

    const fileContent = fs.readFileSync(filePath, 'utf-8')

    async function getList(fileContent) {

        const content = cio.load(fileContent)

        const spans = content('span')

        const divs = content('.jDMyyj')

        spans.each((i, elem) => {
            products.push(content(elem).text())
        })

        for(let i = 24; i < products.length; i++) {
            if(i%10 === 4 && products[i].indexOf('Готов к вводу в оборот') < 0 && products[i].indexOf('Опубликована') < 0 && products[i] !== '') {
                actual_products.push(products[i])
            }
        }

        console.log(actual_products)

        divs.each((i, elem) => {
            gtins.push(content(elem).text())
        })

        for(let i = 0; i < gtins.length; i++) {
            if(gtins[i].indexOf('029') >= 0) {
                actual_gtins.push(gtins[i].replace('0', ''))
            }
        }

    }

    html = `${headerComponent}
                    <title>Обновление краткого отчета</title>
                </head>
                <body>
                        ${navComponent}
                        <section class="sub-nav import-main">
                            <div class="import-control">`

    await renderImportButtons(buttons)
    await renderMarkingButtons()
    await renderExtraButtons()

    html += `</section>`

    await getList(fileContent)

    for(let i = 0; i < actual_products.length; i++) {

        newProducts.push({

            'name': actual_products[i],
            'gtin': actual_gtins[i]

        })

    }

    async function updateShortReport() {

        const fileName = './public/Краткий отчет.xlsx'

        const wb = new exl.Workbook()

        await wb.xlsx.readFile(fileName)

        const ws = wb.getWorksheet(1)

        const firstColumn = ws.getColumn(1)

        let cellNumber = firstColumn.values.length

        console.log(cellNumber)

        html += `<section class="table">
                <div class="marks-table">
                    <div class="marks-table-header">
                        <div class="header-cell">GTIN</div>
                        <div class="header-cell">Наименование</div>                            
                    </div>
                <div class="header-wrapper"></div>`

        for(let i = 0; i < newProducts.length; i++) {

            html += `<div class="table-row">
                        <span id="gtin">${newProducts[i].gtin}</span>
                        <span id="name">${newProducts[i].name}</span>
                     </div>`

            const row = ws.getRow(cellNumber)
            row.getCell(1).value = newProducts[i].gtin
            row.getCell(2).value = newProducts[i].name
            row.commit()

            cellNumber++

        }

        html += `</section>
             <div class="body-wrapper"></div>                        
             ${footerComponent}`

        await wb.xlsx.writeFile(fileName)

    }

    await updateShortReport()

    res.send(html)

})

app.get('/clear_duplicate', async function(req, res){

    const workbook = new exl.Workbook()

})

app.get('/personal_orders', async function(req, res) {

    html = `${headerComponent}
                    <title>Заказ маркировки - WILDBERRIES</title>
                </head>
                <body>
                    ${navComponent}
                        <section class="sub-nav import-main">
                            <div class="import-control">`

    await renderImportButtons(buttons)
    await renderMarkingButtons()
    await renderExtraButtons()

    html += `</section>`

    const workbook = new exl.Workbook()

    const orderPath = './public/personalOrder.xlsx'
    const shortReport = './public/Краткий отчет.xlsx'
    const unloadFile = './public/Выгрузка 372900043349.xlsx'

    const orderProducts = []
    const full_cat = []
    const nat_cat = []    
    const gtins = []

    await workbook.xlsx.readFile(orderPath)

    const ws_1 = workbook.getWorksheet('Лист1')

    ws_1.eachRow((row, rowNumber) => {

        orderProducts.push({

            name: row.values[2].richText[0].text.trim(),
            quantity: row.values[1],
            vendor: row.values[3].richText[0].text

        })

    })

    await workbook.xlsx.readFile(shortReport)
        
    const ws_2 = workbook.getWorksheet('Краткий отчет')

    const c1 = ws_2.getColumn(1)

    c1.eachCell({includeEmpty: false}, (c, rowNumber) => {
        if(rowNumber < 5) return        
        gtins.push(c.value)
    })

    const c2 = ws_2.getColumn(2)

    c2.eachCell({includeEmpty: false}, (c, rowNumber) => {
        if(rowNumber < 5) return
        nat_cat.push(c.value)
    })

    let difference = []

    for(let i = 0; i < orderProducts.length; i++) {

        if(nat_cat.indexOf(orderProducts[i].name) < 0) {

            difference.push(orderProducts[i])

        }

    }

    await workbook.xlsx.readFile(unloadFile)

    const ws_3 = workbook.getWorksheet('result')

    ws_3.eachRow((row, rowNumber) => {

        if(rowNumber < 2) return
        full_cat.push({
            name: row.values[8],
            vendor: row.values[10]
        })

    })

    console.log(difference)
    console.log(difference.length)

    difference = difference.filter((o) => {

        if(full_cat.findIndex(i => i.vendor === o.vendor) < 0) {

            return o

        }

    })

    console.log(difference.length)

    let names = []

    for(let i = 0; i < difference.length; i++) {

        try {

            const response = await axios.post('https://api-seller.ozon.ru/v4/product/info/attributes', {
                        
                "filter": {
                    "offer_id": [
                        difference[i].vendor
                    ],
                    "visibility": "ALL"
                },
                "limit": 1000,
                "sort_dir": "ASC"

            }, {
                headers: {
                    'Host':'api-seller.ozon.ru',
                    'Client-Id':`${process.env.OZON_CLIENT_ID}`,
                    'Api-Key':`${process.env.OZON_API_KEY}`,
                    'Content-Type':'application/json'
                }
            })

            if(response.data.result[0].name.indexOf('Пододеяльник') >= 0) {

                names.push({            
                    'vendor': difference[i].vendor,
                    'name': response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                    .trim()                  // убрать пробелы по краям
                                                    .replace(/\s+/g, ' '),
                    'size': response.data.result[0].attributes.find(o => o.id === 6773).values[0].value,
                    'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                    'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                    'productType': 'ПОДОДЕЯЛЬНИК С КЛАПАНОМ'
                })   

            }

            if(response.data.result[0].name.indexOf('Простыня') >= 0 && response.data.result[0].name.indexOf('белье') < 0 && response.data.result[0].name.indexOf('бельё') < 0) {

                if(response.data.result[0].name.indexOf('на резинке') >= 0) {

                    names.push({            
                        'vendor': difference[i].vendor,
                        'name': response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                        .trim()                  // убрать пробелы по краям
                                                        .replace(/\s+/g, ' '),
                        'size': `${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}x${response.data.result[0].attributes.find(o => o.id === 8414).values[0].value}`,
                        'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                        'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                        'productType': 'ПРОСТЫНЯ НА РЕЗИНКЕ'
                    })

                }

                if(response.data.result[0].name.indexOf('на резинке') < 0) {

                    names.push({            
                        'vendor': difference[i].vendor,
                        'name': response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                        .trim()                  // убрать пробелы по краям
                                                        .replace(/\s+/g, ' '),
                        'size': response.data.result[0].attributes.find(o => o.id === 6771).values[0].value,
                        'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                        'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                        'productType': 'ПРОСТЫНЯ'
                    })

                }

            }

            if(response.data.result[0].name.indexOf('Наволочка') >= 0 || response.data.result[0].name.indexOf('наволочка') >= 0 && response.data.result[0].name.indexOf('белье') < 0 && response.data.result[0].name.indexOf('бельё') < 0) {

                if(response.data.result[0].name.indexOf('50х70') >= 0 || response.data.result[0].name.indexOf('40х60') >= 0 || response.data.result[0].name.indexOf('50 х 70') >= 0 || response.data.result[0].name.indexOf('40 х 60') >= 0 ) {

                    names.push({
                        'vendor': difference[i].vendor,
                        'name': response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                        .trim()                  // убрать пробелы по краям
                                                        .replace(/\s+/g, ' '),
                        'size': response.data.result[0].attributes.find(o => o.id === 6772).values[0].value,
                        'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                        'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                        'productType': 'НАВОЛОЧКА ПРЯМОУГОЛЬНАЯ'
                    })

                } else {

                    names.push({
                        'vendor': difference[i].vendor,
                        'name': response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                        .trim()                  // убрать пробелы по краям
                                                        .replace(/\s+/g, ' '),
                        'size': response.data.result[0].attributes.find(o => o.id === 6772).values[0].value,
                        'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                        'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                        'productType': 'НАВОЛОЧКА КВАДРАТНАЯ'
                    })

                }

            }

            if(response.data.result[0].name.indexOf('белье') >= 0 || response.data.result[0].name.indexOf('бельё') >= 0) {

                if(response.data.result[0].attributes.find(o => o.id === 6772).values.length === 2) {

                    if(response.data.result[0].name.indexOf('на резинке') >= 0) {

                        if(response.data.result[0].name.indexOf('х20 -') >= 0 ||response.data.result[0].name.indexOf('х 20 -') >= 0) {

                            names.push({
                                'vendor': difference[i].vendor,
                                'name': `КПБ ${response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                                .trim()                  // убрать пробелы по краям
                                                                .replace(/\s+/g, ' ')}`,
                                'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}x20; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[1].value}`,
                                'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                                'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                                'productType': 'КОМПЛЕКТ'
                            })

                        }

                        if(response.data.result[0].name.indexOf('х30 -') >= 0 ||response.data.result[0].name.indexOf('х 30 -') >= 0) {

                            names.push({
                                'vendor': difference[i].vendor,
                                'name': `КПБ ${response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                                .trim()                  // убрать пробелы по краям
                                                                .replace(/\s+/g, ' ')}`,
                                'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}x30; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[1].value}`,
                                'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                                'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                                'productType': 'КОМПЛЕКТ'
                            })

                        }

                        if(response.data.result[0].name.indexOf('х40') >= 0 ||response.data.result[0].name.indexOf('х 40 -') >= 0) {

                            names.push({
                                'vendor': difference[i].vendor,
                                'name': `КПБ ${response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                                .trim()                  // убрать пробелы по краям
                                                                .replace(/\s+/g, ' ')}`,
                                'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}x40; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[1].value}`,
                                'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                                'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                                'productType': 'КОМПЛЕКТ'
                            })

                        }

                    }

                    if(response.data.result[0].name.indexOf('на резинке') < 0) {

                        names.push({
                            'vendor': difference[i].vendor,
                            'name': `КПБ ${response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                                .trim()                  // убрать пробелы по краям
                                                                .replace(/\s+/g, ' ')}`,
                            'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[1].value}`,
                            'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                            'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                            'productType': 'КОМПЛЕКТ'
                        })

                    }

                }

                if(response.data.result[0].attributes.find(o => o.id === 6772).values.length === 1) {

                    if(response.data.result[0].name.indexOf('на резинке') >= 0) {

                        if(response.data.result[0].name.indexOf('х20 -') >= 0 ||response.data.result[0].name.indexOf('х 20 -') >= 0) {

                            names.push({
                                'vendor': difference[i].vendor,
                                'name': `КПБ ${response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                                .trim()                  // убрать пробелы по краям
                                                                .replace(/\s+/g, ' ')}`,
                                'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}x20; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}`,
                                'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                                'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                                'productType': 'КОМПЛЕКТ'
                            })

                        }

                        if(response.data.result[0].name.indexOf('х30 -') >= 0 ||response.data.result[0].name.indexOf('х 30 -') >= 0) {

                            names.push({
                                'vendor': difference[i].vendor,
                                'name': `КПБ ${response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                                .trim()                  // убрать пробелы по краям
                                                                .replace(/\s+/g, ' ')}`,
                                'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}x30; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}`,
                                'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                                'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                                'productType': 'КОМПЛЕКТ'
                            })

                        }

                        if(response.data.result[0].name.indexOf('х40 -') >= 0 ||response.data.result[0].name.indexOf('х 40 -') >= 0) {

                            names.push({
                                'vendor': difference[i].vendor,
                                'name': `КПБ ${response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                                .trim()                  // убрать пробелы по краям
                                                                .replace(/\s+/g, ' ')}`,
                                'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}x40; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}`,
                                'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                                'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                                'productType': 'КОМПЛЕКТ'
                            })

                        }

                    }

                    if(response.data.result[0].name.indexOf('на резинке') < 0) {

                        names.push({
                            'vendor': difference[i].vendor,
                            'name': `КПБ ${response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                                .trim()                  // убрать пробелы по краям
                                                                .replace(/\s+/g, ' ')}`,
                            'size': `Пододеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}`,
                            'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                            'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                            'productType': 'КОМПЛЕКТ'
                        })

                    }

                }
                

            }

            names = names.filter(o => o.name.indexOf('Одеяло') < 0 && o.name.indexOf('Подушка') < 0 && o.name.indexOf('Матрас') < 0 && o.name.indexOf('Наматрас')  < 0 && o.name.indexOf('Ветошь') < 0)

        } catch(err) {

            names.push({            
                    'vendor': difference[i].vendor,
                    'name': difference[i].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                    .trim()                  // убрать пробелы по краям
                                                    .replace(/\s+/g, ' '),
                    'size': '',
                    'color': '',
                    'cloth': '',
                    'productType': ''
                })

        }

    }
    
    const new_items = []

    if(difference.length > 0) {

        difference.forEach(el => {

            new_items.push(el.name)

        })

    }

    async function createImport(array) {

        const fileName = './public/IMPORT_TNVED_6302 (3).xlsx'
        
        const wb = new exl.Workbook()

        await wb.xlsx.readFile(fileName)

        const ws = wb.getWorksheet('IMPORT_TNVED_6302')

        let cellNumber = 5

        for(let i = 0; i < array.length; i++) {

            ws.getCell(`A${cellNumber}`).value = 6302
            ws.getCell(`B${cellNumber}`).value = names.find(o => o.name === array[i]).name
            ws.getCell(`C${cellNumber}`).value = 'Ивановский текстиль'
            ws.getCell(`D${cellNumber}`).value = 'Артикул'
            ws.getCell(`E${cellNumber}`).value = names.find(o => o.name === array[i]).vendor
            ws.getCell(`F${cellNumber}`).value = names.find(o => o.name === array[i]).productType
            ws.getCell(`G${cellNumber}`).value = names.find(o => o.name === array[i]).color
            ws.getCell(`H${cellNumber}`).value = 'ВЗРОСЛЫЙ'

            if(names.find(o => o.name === array[i]).cloth === 'КРЕП-ЖАТКА' || names.find(o => o.name === array[i]).cloth === 'КРЕП ЖАТКА') ws.getCell(`I${cellNumber}`).value = 'КРЕП'
            if(names.find(o => o.name === array[i]).cloth === 'ВАРЕНЫЙ ХЛОПОК') ws.getCell(`I${cellNumber}`).value = 'ХЛОПКОВАЯ ТКАНЬ'
            if(names.find(o => o.name === array[i]).cloth === 'ЛЕН' || names.find(o => o.name === array[i]).cloth === 'ЛЁН') ws.getCell(`I${cellNumber}`).value = 'ЛЬНЯНАЯ ТКАНЬ'
            if(names.find(o => o.name === array[i]).cloth === 'СТРАЙП САТИН') ws.getCell(`I${cellNumber}`).value = 'СТРАЙП-САТИН'
            if(names.find(o => o.name === array[i]).cloth === 'САТИН ЛЮКС') ws.getCell(`I${cellNumber}`).value = 'САТИН'
            if(names.find(o => o.name === array[i]).cloth !== 'САТИН ЛЮКС' && names.find(o => o.name === array[i]).cloth !== 'СТРАЙП САТИН' && names.find(o => o.name === array[i]).cloth !== 'ВАРЕНЫЙ ХЛОПОК' && names.find(o => o.name === array[i]).cloth !== 'ЛЕН' && names.find(o => o.name === array[i]).cloth !== 'ЛЁН') ws.getCell(`I${cellNumber}`).value = names.find(o => o.name === array[i]).cloth
            
            if(names.find(o => o.name === array[i]).cloth === 'ПОЛИСАТИН') ws.getCell(`J${cellNumber}`).value = '100% Полиэстер'

            if(names.find(o => o.name === array[i]).cloth === 'ТЕНСЕЛЬ') ws.getCell(`J${cellNumber}`).value = '100% Лиоцелл'
            if(names.find(o => o.name === array[i]).cloth === 'ЛЕН' || names.find(o => o.name === array[i]).cloth === 'ЛЁН') ws.getCell(`J${cellNumber}`).value = '100% Лен'
            if(names.find(o => o.name === array[i]).cloth !== 'КРЕП-ЖАТКА' && names.find(o => o.name === array[i]).cloth !== 'КРЕП ЖАТКА' && names.find(o => o.name === array[i]).cloth !== 'ПОЛИСАТИН' && names.find(o => o.name === array[i]).cloth !== 'ТЕНСЕЛЬ' && names.find(o => o.name === array[i]).cloth !== 'ЛЕН' && names.find(o => o.name === array[i]).cloth !== 'ЛЁН') ws.getCell(`J${cellNumber}`).value = '100% Хлопок'

            ws.getCell(`K${cellNumber}`).value = names.find(o => o.name === array[i]).size
            ws.getCell(`L${cellNumber}`).value = '6302100001'
            ws.getCell(`M${cellNumber}`).value = 'ТР ТС 017/2011 "О безопасности продукции легкой промышленности'
            ws.getCell(`N${cellNumber}`).value = 'На модерации'                

            cellNumber++

        }

        ws.unMergeCells('D2')

        ws.getCell('E2').value = '13914'

        ws.getCell('E2').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor:{argb:'E3E3E3'}
        }

        ws.getCell('E2').font = {
            size: 10,
            name: 'Arial'
        }

        ws.getCell('E2').alignment = {
            horizontal: 'center',
            vertical: 'bottom'
        }        

        const date_ob = new Date()

        let month = date_ob.getMonth() + 1

        let filePath = ''

        month < 10 ? filePath = `./public/personal/IMPORT_TNVED_6302_${date_ob.getDate()}_0${month}_personal` : filePath = `./public/personal/IMPORT_TNVED_6302_${date_ob.getDate()}_0${month}_personal`

        fs.access(`${filePath}.xlsx`, fs.constants.R_OK, async (err) => {
            if(err) {
                await wb.xlsx.writeFile(`${filePath}.xlsx`)
            } else {
                let count = 1
                fs.access(`${filePath}_(1).xlsx`, fs.constants.R_OK, async (err) => {
                    if(err) {
                        await wb.xlsx.writeFile(`${filePath}_(1).xlsx`)
                    } else {
                        await wb.xlsx.writeFile(`${filePath}_(2).xlsx`)
                    }
                })
                
            }
        })

    }

    function createNameList() {

        let orderList = []
        let _temp = []

        for (let i = 0; i < orderProducts.length; i++) {

                _temp.push(orderProducts[i].name)
            
                if(_temp.length%10 === 0) {
                    orderList.push(_temp)
                    _temp = []
                }
        }        

        if(_temp.length > 0) {
            orderList.push(_temp)
            _temp = []
        }

        return orderList

    }

    function createQuantityList() {

        let quantityList = []
        let temp = []

        for(let i = 0; i < orderProducts.length; i++) {
            
            if(nat_cat.indexOf(orderProducts[i].name) >= 0) {
                temp.push(orderProducts.find(o => o.vendor === orderProducts[i].vendor).quantity)
            }

            if(temp.length%10 === 0) {
                quantityList.splice(-1, 0, ...quantityList.splice(-1, 1, temp))
                temp = []
            }

        }

        quantityList.splice(-1, 0, ...quantityList.splice(-1, 1, temp))

        return quantityList

    }

    let List = createNameList()
    let Quantity = createQuantityList()

    function createOrder() {

        let content = ``

        for(let i = 0; i < List.length; i++) {
            if(List[i].length > 0) {
                content += `<?xml version="1.0" encoding="utf-8"?>
                            <order xmlns="urn:oms.order" xsi:schemaLocation="urn:oms.order schema.xsd" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
                                <lp>
                                    <productGroup>lp</productGroup>
                                    <contactPerson>333</contactPerson>
                                    <releaseMethodType>REMARK</releaseMethodType>
                                    <createMethodType>SELF_MADE</createMethodType>
                                    <productionOrderId>PERSONAL_${i+1}</productionOrderId>
                                    <products>`
                
                    for(let j = 0; j < List[i].length; j++) {                
                        if(nat_cat.indexOf(List[i][j]) >= 0) {
                            content += `<product>
                                            <gtin>0${gtins[nat_cat.indexOf(List[i][j])]}</gtin>
                                            <quantity>${Quantity[i][j]}</quantity>
                                            <serialNumberType>OPERATOR</serialNumberType>
                                            <cisType>UNIT</cisType>
                                            <templateId>10</templateId>
                                        </product>`
                        }
                    }
                
                content += `    </products>
                            </lp>
                        </order>`

            }

            const date_ob = new Date()

            let month = date_ob.getMonth() + 1

            let filePath = ''

            month < 10 ? filePath = `./public/orders/lp_personal_${i}_${date_ob.getDate()}_0${month}.xml` : filePath = `./public/orders/lp_personal_${i}_${date_ob.getDate()}_${month}.xml`

            if(content !== ``) {
                fs.writeFileSync(filePath, content)
            }

            content = ``

        }

        html += `<section class="table">
                <div class="marks-table">
                    <div class="marks-table-header">
                        <div class="header-cell">Наименование</div>
                        <div class="header-cell">Статус</div>
                        <div class="header-cell">Кол-во</div>
                    </div>
                <div class="header-wrapper"></div>`

        for(let i = 0; i < List.length; i++) {
            for(let j = 0; j < List[i].length; j++) {
                if(nat_cat.indexOf(List[i][j]) < 0) {
                    html += `<div class="table-row">
                                <span id="name">${List[i][j]}</span>
                                <span id="status-new">Новый товар</span>
                                <span id="quantity">${Quantity[i][j]}</span>
                             </div>`
                } else {
                    html += `<div class="table-row">
                                <span id="name">${List[i][j]}</span>
                                <span id="status-current">Актуальный товар</span>
                                <span id="quantity">${Quantity[i][j]}</span>
                             </div>`
                }
            }
        }

        html += `<section>
                <div class="body-wrapper"></div>
            ${footerComponent}`        

    }

    if(new_items.length > 0) await createImport(new_items)

    if(difference.length <= 0) createOrder()

    res.send(html)
    
})

app.get('/get_products_analytic/:year', async function (req, res) {

    const year = req.params.year
    const ordersList = []

    const analyticObject = {
        "Одеяло": 0,
        "Подушка": 0,
        "Наперник": 0,
        "Наматрасник": 0,
        "Постельное": 0,
        "Ветошь": 0,
        "Халат": 0,
        "Простыня": 0,
        "Пододеяльник": 0,
        "Покрывало": 0,
        "Наволочка": 0,
        "Матрас": 0
    }

    let count = 0
    let hasNext = true
    let offset = 0

    while(hasNext === true) {
        const response = await axios.post('https://api-seller.ozon.ru/v3/posting/fbs/list', {
            "dir": "asc",
            "filter": {
                "since": `${year}-01-01T00:00:00.000Z`,
                "status": "delivered",
                "to": `${year}-12-31T23:59:59.999Z`
            },
            "limit": 1000,
            "offset": offset
        }, {
            headers: {
                "Client-Id": process.env.OZON_CLIENT_ID,
                "Api-Key": process.env.OZON_API_KEY
            }
        })

        for(let order of response.data.result.postings) {
            ordersList.push(order)
        }

        hasNext = response.data.result.has_next
        count += response.data.result.postings.length
        offset = count

    }

    for(let order of ordersList) {

        if(order.products.find(o => o.name.indexOf('Одеяло') >= 0)) {
            analyticObject["Одеяло"] = analyticObject["Одеяло"] + 1
        }

        if(order.products.find(o => o.name.indexOf('Подушка') >= 0)) {
            analyticObject["Подушка"] = analyticObject["Подушка"] + 1
        }

        if(order.products.find(o => o.name.indexOf('Наперник') >= 0)) {
            analyticObject["Наперник"] = analyticObject["Наперник"] + 1
        }

        if(order.products.find(o => o.name.indexOf('Наматрасник') >= 0)) {
            analyticObject["Наматрасник"] = analyticObject["Наматрасник"] + 1
        }

        if(order.products.find(o => o.name.indexOf('Постельное') >= 0)) {
            analyticObject["Постельное"] = analyticObject["Постельное"] + 1
        }

        if(order.products.find(o => o.name.indexOf('Пеленки') >= 0)) {
            analyticObject["Пеленки"] = analyticObject["Пеленки"] + 1
        }

        if(order.products.find(o => o.name.indexOf('Халат') >= 0)) {
            analyticObject["Халат"] = analyticObject["Халат"] + 1
        }

        if(order.products.find(o => o.name.indexOf('Простынь') >= 0 || o.name.indexOf('Простыня') >= 0)) {
            analyticObject["Простыня"] = analyticObject["Простыня"] + 1
        }

        if(order.products.find(o => o.name.indexOf('Покрывало') >= 0)) {
            analyticObject["Покрывало"] = analyticObject["Покрывало"] + 1
        }

        if(order.products.find(o => o.name.indexOf('Наволочка') >= 0)) {
            analyticObject["Наволочка"] = analyticObject["Наволочка"] + 1
        }

        if(order.products.find(o => o.name.indexOf('Ветошь') >= 0)) {
            analyticObject["Ветошь"] = analyticObject["Ветошь"] + 1
        }

        if(order.products.find(o => o.name.indexOf('Пододеяльник') >= 0)) {
            analyticObject["Пододеяльник"] = analyticObject["Пододеяльник"] + 1
        }

        if(order.products.find(o => o.name.indexOf('Матрас') >= 0)) {
            analyticObject["Матрас"] = analyticObject["Матрас"] + 1
        }

    }

    let html = `${headerComponent}
            <title>Аналитика спроса на товары</title>
        </head>
        <body>
            <div class="columns is-mobile">
                <div class="column is-three-fifths is-offset-one-fifth">
                    <h1 class="title is-4 has-text-centered has-text-black">Аналитика спроса за ${year} год.</h1>
                </div>
            </div>`

    const myChart = new QuickChart()

    myChart.setConfig({
        type: 'bar',
        data: {
            labels: Object.keys(analyticObject),
            datasets: [
                {
                    label: 'Получено, шт.',
                    data: Object.values(analyticObject),
                    fill: false
                }
            ]
        }
    })
    .setWidth(800)
    .setHeight(400)
    .setBackgroundColor('transparent')

    const chartUrl = myChart.getUrl()

    html += `
            <div class="columns is-mobile">
                <div class="column is-three-fifths is-offset-one-fifth">
                    <img src="${chartUrl}">
                </div>
            </div>`

    html += `<div class="columns is-mobile">
                    <div class="column is-three-fifths is-offset-one-fifth">
                        <table class="table is-fullwidth my-table">
                            <thead>
                                <tr>
                                    <th class="has-text-left">Продукт</th>
                                    <th class="has-text-left">Количество</th>
                                </tr>
                            </thead>
                            <tbody>`

    for(let key of Object.keys(analyticObject)) {

        html += `<tr>
                    <td>${key}</td>
                    <td>
                        ${analyticObject[key]} шт.
                    </td>
                </tr>`

    }

    html += `</tbody>
        </table>`    
                    
    html += `</div>
            </div>`

    html += footerComponent

    res.send(html)

    // console.log(ordersList.length)
    // res.json(analyticObject)

})

app.get('/get_products_analytic/:year/:product', async function (req, res) {

    const List = []

    const year = req.params.year
    let ordersList = []

    let analyticObject = {}

    if(req.params.product.toLowerCase().indexOf('простын') >= 0 || req.params.product.toLowerCase().indexOf('пододе') >= 0 || req.params.product.toLowerCase().indexOf('наволочка') >= 0 || req.params.product.toLowerCase().indexOf('постельное') >= 0) {

        analyticObject = {
            "Тенсель": 0,
            "Сатин": 0,
            "Страйп-сатин": 0,
            "Твил-сатин": 0,
            "Полисатин": 0,
            "Бязь": 0,
            "Сатин-жаккард": 0,
            "Вареный хлопок": 0,
            "Мулетон": 0,
            "Микрофибра": 0,
            "Перкаль": 0,
            "Поплин": 0,
            "Ранфорс": 0,
            "Микросатин": 0,
            "Креп-жатка": 0,
            "Жатка": 0
        }

    }

    let count = 0
    let hasNext = true
    let offset = 0

    while(hasNext === true) {
        const response = await axios.post('https://api-seller.ozon.ru/v3/posting/fbs/list', {
            "dir": "asc",
            "filter": {
                "since": `${year}-01-01T00:00:00.000Z`,
                "status": "delivered",
                "to": `${year}-12-31T23:59:59.999Z`
            },
            "limit": 1000,
            "offset": offset
        }, {
            headers: {
                "Client-Id": process.env.OZON_CLIENT_ID,
                "Api-Key": process.env.OZON_API_KEY
            }
        })

        for(let order of response.data.result.postings) {
            ordersList.push(order)
        }

        hasNext = response.data.result.has_next
        count += response.data.result.postings.length
        offset = count

    }

    ordersList = ordersList.filter(o => {
        if(o.products.find(i => i.name.indexOf(req.params.product) >= 0)) {
            return o
        }
    })

    let specialFilter = {}

    if(req.params.product.toLowerCase().indexOf('простын') >= 0) {

        if(req.params.product.toLowerCase().indexOf('простын') >= 0) {

            specialFilter = {
                "На резинке": 0,
                "Стандартная": 0
            }

            ordersList.forEach(o => {

                if(o.products.find(i => i.name.indexOf(req.params.product) >= 0 && i.name.indexOf('на резинке') >= 0)) {

                    specialFilter["На резинке"] = specialFilter["На резинке"] + 1

                } else {

                    specialFilter["Стандартная"] = specialFilter["Стандартная"] + 1

                }

            })

        }

        for(let order of ordersList) {

            if(order.products.find(o => o.name.indexOf('тенсел') >= 0)) {
                analyticObject["Тенсель"] = analyticObject["Тенсель"] + 1
            }

            if(order.products.find(o => o.name.indexOf('сатин') >= 0 && o.name.indexOf('страйп') < 0 && o.name.indexOf('жаккард') < 0 && o.name.indexOf('твил') < 0 && o.name.indexOf('поли') < 0)) {
                analyticObject["Сатин"] = analyticObject["Сатин"] + 1
            }

            if(order.products.find(o => o.name.indexOf('сатин') >= 0 && o.name.indexOf('страйп') >= 0)) {
                analyticObject["Страйп-сатин"] = analyticObject["Страйп-сатин"] + 1
            }

            if(order.products.find(o => o.name.indexOf('сатин') >= 0 && o.name.indexOf('твил') >= 0)) {
                analyticObject["Твил-сатин"] = analyticObject["Твил-сатин"] + 1
            }

            if(order.products.find(o => o.name.indexOf('сатин') >= 0 && o.name.indexOf('поли') >= 0)) {
                analyticObject["Полисатин"] = analyticObject["Полисатин"] + 1
            }

            if(order.products.find(o => o.name.indexOf('сатин') >= 0 && o.name.indexOf('жаккард') >= 0)) {
                analyticObject["Сатин-жаккард"] = analyticObject["Сатин-жаккард"] + 1
            }

            if(order.products.find(o => o.name.indexOf('бяз') >= 0)) {
                analyticObject["Бязь"] = analyticObject["Бязь"] + 1
            }

            if(order.products.find(o => o.name.indexOf('варен') >= 0 || o.name.indexOf('варён') >= 0 || o.name.indexOf('хлоп') >= 0)) {
                analyticObject["Вареный хлопок"] = analyticObject["Вареный хлопок"] + 1
            }

            if(order.products.find(o => o.name.indexOf('микрофибр') >= 0)) {
                analyticObject["Микрофибра"] = analyticObject["Микрофибра"] + 1
            }

            if(order.products.find(o => o.name.indexOf('мулетон') >= 0)) {
                analyticObject["Мулетон"] = analyticObject["Мулетон"] + 1
            }
            
            if(order.products.find(o => o.name.indexOf('поплин') >= 0)) {
                analyticObject["Поплин"] = analyticObject["Поплин"] + 1
            }

            if(order.products.find(o => o.name.indexOf('перкал') >= 0)) {
                analyticObject["Перкаль"] = analyticObject["Перкаль"] + 1
            }

            if(order.products.find(o => o.name.indexOf('ранфор') >= 0)) {
                analyticObject["Ранфорс"] = analyticObject["Ранфорс"] + 1
            }

            if(order.products.find(o => o.name.indexOf('микросатин') >= 0)) {
                analyticObject["Микросатин"] = analyticObject["Микросатин"] + 1
            }

            if(order.products.find(o => o.name.indexOf('креп-ж') >= 0)) {
                analyticObject["Креп-жатка"] = analyticObject["Креп-жатка"] + 1
            }

            if(order.products.find(o => o.name.indexOf('креп-ж') < 0 && order.products.find(o => o.name.indexOf('жатка') >= 0))) {
                analyticObject["Жатка"] = analyticObject["Жатка"] + 1
            }

        }

        let html = `${headerComponent}
                <title>Аналитика спроса на товары</title>
            </head>
            <body>
                <div class="columns is-mobile">
                    <div class="column is-three-fifths is-offset-one-fifth">
                        <h1 class="title is-4 has-text-centered has-text-black">Аналитика спроса на простыни за ${year} год.</h1>
                    </div>
                </div>
                <div class="fixed-grid has-1-cols">
                    <div class="grid">`

        const myChart = new QuickChart()

        myChart.setConfig({
            type: 'bar',
            data: {
                labels: Object.keys(analyticObject),
                datasets: [
                    {
                        label: 'Получено, шт.',
                        data: Object.values(analyticObject),
                        fill: false
                    }
                ]
            }
        })
        .setWidth(800)
        .setHeight(400)
        .setBackgroundColor('transparent')

        const chartUrl = myChart.getUrl()

        const specialChart = new QuickChart()

        specialChart.setConfig({
            type: 'bar',
            data: {
                labels: Object.keys(specialFilter),
                datasets: [
                    {
                        label: 'Получено, шт.',
                        data: Object.values(specialFilter),
                        fill: false
                    }
                ]
            }
        })
        .setWidth(800)
        .setHeight(400)
        .setBackgroundColor('transparent')

        const specialUrl = specialChart.getUrl()

        html += `<div class="cell">
                    <img src="${chartUrl}">`

        html += `
                <table class="table is-fullwidth my-table">
                    <thead>
                        <tr>
                            <th class="has-text-left has-text-black">Продукт</th>
                            <th class="has-text-left has-text-black">Количество</th>
                        </tr>
                    </thead>
                    <tbody>`

        for(let key of Object.keys(analyticObject)) {

            html += `<tr>
                        <td class="has-text-black">${key}</td>
                        <td class="has-text-black">
                            ${analyticObject[key]} шт.
                        </td>
                    </tr>`

        }

        html += `</tbody>
            </table>`    
                        
        html += `</div>`

        html += `<div class="cell">
                    <img src="${specialUrl}">`

        html += `<table class="table is-fullwidth my-table">
                    <thead>
                        <tr>
                            <th class="has-text-left has-text-black">Продукт</th>
                            <th class="has-text-left has-text-black">Количество</th>
                        </tr>
                    </thead>
                    <tbody>`
                                
        for(let key of Object.keys(specialFilter)) {

            html += `<tr>
                        <td class="has-text-black">${key}</td>
                        <td class="has-text-black">
                            ${specialFilter[key]} шт.
                        </td>
                    </tr>`

        }

        html += `</tbody>
            </table>`    
                        
        html += `</div>`

        const rubberOrders = ordersList.filter(o => {
            if(o.products.find(i => i.name.indexOf(req.params.product) >= 0 && i.name.indexOf('на резинке') >= 0)) {
                return o
            }
        })

        let bedsheetHeight = {
            "10": 0,
            "20": 0,
            "30": 0,
            "40": 0,
            "45": 0
        }

        let bedsheetSizes = {

        }

        let rubberOffersCount = {
            
        }

        const rubberOffers = []

        for(order of rubberOrders) {

            order.products.forEach( async (i) => {

                if(i.name.indexOf('на резинке') >= 0 && i.name.toLowerCase().indexOf('простын') >= 0) {

                    rubberOffers.push(i.offer_id)

                }

            })

        }

        for(item of rubberOffers) {

            if(String(item) in rubberOffersCount) {

                rubberOffersCount[String(item)] = rubberOffersCount[String(item)] + 1

            } else {

                rubberOffersCount[String(item)] = 1

            }

        }

        const uniqueOffers = [...new Set(rubberOffers)]

        const response = await axios.post('https://api-seller.ozon.ru/v4/product/info/attributes', {

            "filter": {
                "offer_id": uniqueOffers,
                "visibility": "ALL"
            },
            "limit": 1000,
            "sort_dir": "ASC"

        }, {
            headers: {
                "Client-Id": process.env.OZON_CLIENT_ID,
                "Api-Key": process.env.OZON_API_KEY
            }
        })

        const data = response.data.result

        for(let i of data) {

            console.log(i.attributes.find(o => o.id === 6771).values[0].value)

            if(String(i.attributes.find(o => o.id === 6771).values[0].value) in bedsheetSizes) {

                bedsheetSizes[String(i.attributes.find(o => o.id === 6771).values[0].value)] = bedsheetSizes[String(i.attributes.find(o => o.id === 6771).values[0].value)] + 1

            } else {

                bedsheetSizes[String(i.attributes.find(o => o.id === 6771).values[0].value)] = 1

            }

            if(Number(i.attributes.find(o => o.id === 8414).values[0].value) === 10) {

                bedsheetHeight["10"] = bedsheetHeight["10"] + 1

            }

            if(Number(i.attributes.find(o => o.id === 8414).values[0].value) === 20) {

                bedsheetHeight["20"] = bedsheetHeight["20"] + 1

            }

            if(Number(i.attributes.find(o => o.id === 8414).values[0].value) === 30) {

                bedsheetHeight["30"] = bedsheetHeight["30"] + 1

            }

            if(Number(i.attributes.find(o => o.id === 8414).values[0].value) === 40) {

                bedsheetHeight["40"] = bedsheetHeight["40"] + 1

            }

            if(Number(i.attributes.find(o => o.id === 8414).values[0].value) === 45) {

                bedsheetHeight["45"] = bedsheetHeight["45"] + 1

            }

        }

        const heightChart = new QuickChart()

        heightChart.setConfig({
            type: 'bar',
            data: {
                labels: Object.keys(bedsheetHeight),
                datasets: [
                    {
                        label: 'Получено, шт.',
                        data: Object.values(bedsheetHeight),
                        fill: false
                    }
                ]
            }
        })
        .setWidth(800)
        .setHeight(400)
        .setBackgroundColor('transparent')

        const heightChartUrl = heightChart.getUrl()

        html += `<div class="cell">
                    <img src="${heightChartUrl}">`

        html += `<table class="table is-fullwidth my-table">
                    <thead>
                        <tr>
                            <th class="has-text-left has-text-black">Высота</th>
                            <th class="has-text-left has-text-black">Количество</th>
                        </tr>
                    </thead>
                    <tbody>`

        for(let key of Object.keys(bedsheetHeight)) {

            html += `<tr>
                        <td class="has-text-black">${key}</td>
                        <td class="has-text-black">
                            ${bedsheetHeight[key]} шт.
                        </td>
                    </tr>`

        }

        html += `</tbody>
            </table>`    
                        
        html += `</div>`

        const sizeChart = new QuickChart()

        sizeChart.setConfig({
            type: 'bar',
            data: {
                labels: Object.keys(bedsheetSizes),
                datasets: [
                    {
                        label: 'Получено, шт.',
                        data: Object.values(bedsheetSizes),
                        fill: false
                    }
                ]
            }
        })
        .setWidth(800)
        .setHeight(400)
        .setBackgroundColor('transparent')

        const sizeChartUrl = sizeChart.getUrl()

        html += `<div class="cell">
                        <img src="${sizeChartUrl}">`

        html += `<table class="table is-fullwidth my-table">
                    <thead>
                        <tr>
                            <th class="has-text-left has-text-black">Размер, (Д×Ш)</th>
                            <th class="has-text-left has-text-black">Количество</th>
                        </tr>
                    </thead>
                    <tbody>`

        for(let key of Object.keys(bedsheetSizes)) {

            html += `<tr>
                        <td class="has-text-black">${key}</td>
                        <td class="has-text-black">
                            ${bedsheetSizes[key]} шт.
                        </td>
                    </tr>`

        }

        html += `</tbody>
            </table>`    
                        
        html += `</div>`

        const standartOrders = ordersList.filter(o => {
            if(o.products.find(i => i.name.indexOf(req.params.product) >= 0 && i.name.indexOf('на резинке') < 0)) {
                return o
            }
        })

        let bedsheetSizesStandart = {

        }

        const standartOffers = []

        for(order of standartOrders) {

            order.products.forEach( async (i) => {

                if(i.name.indexOf('на резинке') < 0 && i.name.toLowerCase().indexOf('простын') >= 0) {

                    standartOffers.push(i.offer_id)                    

                }

            })

        }

        const uniqueOffersStandart = [...new Set(standartOffers)]

        const responseStandart = await axios.post('https://api-seller.ozon.ru/v4/product/info/attributes', {

            "filter": {
                "offer_id": uniqueOffersStandart,
                "visibility": "ALL"
            },
            "limit": 1000,
            "sort_dir": "ASC"

        }, {
            headers: {
                "Client-Id": process.env.OZON_CLIENT_ID,
                "Api-Key": process.env.OZON_API_KEY
            }
        })

        const dataStandart = responseStandart.data.result

        for(let i of dataStandart) {

            if(String(i.attributes.find(o => o.id === 6771).values[0].value) in bedsheetSizesStandart) {

                bedsheetSizesStandart[String(i.attributes.find(o => o.id === 6771).values[0].value)] = bedsheetSizesStandart[String(i.attributes.find(o => o.id === 6771).values[0].value)] + 1

            } else {

                bedsheetSizesStandart[String(i.attributes.find(o => o.id === 6771).values[0].value)] = 1

            }            

        }

        const sizeStandartChart = new QuickChart()

        sizeStandartChart.setConfig({
            type: 'bar',
            data: {
                labels: Object.keys(bedsheetSizesStandart),
                datasets: [
                    {
                        label: 'Получено, шт.',
                        data: Object.values(bedsheetSizesStandart),
                        fill: false
                    }
                ]
            }
        })
        .setWidth(800)
        .setHeight(400)
        .setBackgroundColor('transparent')

        const sizeStandartChartUrl = sizeStandartChart.getUrl()

        html += `<div class="cell">
                        <img src="${sizeStandartChartUrl}">`

        html += `<table class="table is-fullwidth my-table">
                    <thead>
                        <tr>
                            <th class="has-text-left has-text-black">Размер, (Д×Ш)</th>
                            <th class="has-text-left has-text-black">Количество</th>
                        </tr>
                    </thead>
                    <tbody>`

        for(let key of Object.keys(bedsheetSizesStandart)) {

            html += `<tr>
                        <td class="has-text-black">${key}</td>
                        <td class="has-text-black">
                            ${bedsheetSizesStandart[key]} шт.
                        </td>
                    </tr>`

        }

        html += `</tbody>
            </table>`    
                        
        html += `</div>`

        let bedsheetColors = {

        }

        const orderOffers = []

        for(order of ordersList) {

            order.products.forEach( async (i) => {

                if(i.name.toLowerCase().indexOf('простын') >= 0) {

                    orderOffers.push(i.offer_id)                    

                }

            })

        }

        const uniqueOffersColor = [...new Set(orderOffers)]

        if(uniqueOffersColor.length > 1000) {

            let temp = []

            for(let i = 0; i < 1000; i++) {

                temp.push(uniqueOffersColor[i])

            }

            List.push(temp)

            temp = []

            for(let i = 1000; i < uniqueOffersColor.length; i++) {

                temp.push(uniqueOffersColor[i])

            }

            List.push(temp)

        }

        const responseColorFirst = await axios.post('https://api-seller.ozon.ru/v4/product/info/attributes', {

            "filter": {
                "offer_id": List[0],
                "visibility": "ALL"
            },
            "limit": 1000,
            "sort_dir": "ASC"

        }, {
            headers: {
                "Client-Id": process.env.OZON_CLIENT_ID,
                "Api-Key": process.env.OZON_API_KEY
            }
        })

        let first = responseColorFirst.data.result

        const responseColorSecond = await axios.post('https://api-seller.ozon.ru/v4/product/info/attributes', {

            "filter": {
                "offer_id": List[1],
                "visibility": "ALL"
            },
            "limit": 1000,
            "sort_dir": "ASC"

        }, {
            headers: {
                "Client-Id": process.env.OZON_CLIENT_ID,
                "Api-Key": process.env.OZON_API_KEY
            }
        })

        let second = responseColorSecond.data.result

        const dataColor = [...first, ...second]

        for(let i of dataColor) {

            console.log(i.attributes.find(o => o.id === 10096).values[0].value)

            if(String(i.attributes.find(o => o.id === 10096).values[0].value) in bedsheetColors) {

                bedsheetColors[String(i.attributes.find(o => o.id === 10096).values[0].value)] = bedsheetColors[String(i.attributes.find(o => o.id === 10096).values[0].value)] + 1

            } else {

                bedsheetColors[String(i.attributes.find(o => o.id === 10096).values[0].value)] = 1

            }            

        }

        const colorChart = new QuickChart()

        colorChart.setConfig({
            type: 'bar',
            data: {
                labels: Object.keys(bedsheetColors),
                datasets: [
                    {
                        label: 'Получено, шт.',
                        data: Object.values(bedsheetColors),
                        fill: false
                    }
                ]
            }
        })
        .setWidth(800)
        .setHeight(400)
        .setBackgroundColor('transparent')

        const colorChartUrl = colorChart.getUrl()

        html += `<div class="cell">
                        <img src="${colorChartUrl}">`

        html += `<table class="table is-fullwidth my-table">
                    <thead>
                        <tr>
                            <th class="has-text-left has-text-black">Размер, (Д×Ш)</th>
                            <th class="has-text-left has-text-black">Количество</th>
                        </tr>
                    </thead>
                    <tbody>`

        for(let key of Object.keys(bedsheetColors)) {

            html += `<tr>
                        <td class="has-text-black">${key}</td>
                        <td class="has-text-black">
                            ${bedsheetColors[key]} шт.
                        </td>
                    </tr>`

        }

        html += `</tbody>
            </table>`    
                        
        html += `</div>`

        html += `
                    </div>
                </div>
            ${footerComponent}`

        // res.send(html)

        res.json({data, rubberOffersCount})
        
    }

    if(req.params.product.toLowerCase().indexOf('пододе') >= 0) {

        for(let order of ordersList) {

            if(order.products.find(o => o.name.indexOf('тенсел') >= 0)) {
                analyticObject["Тенсель"] = analyticObject["Тенсель"] + 1
            }

            if(order.products.find(o => o.name.indexOf('сатин') >= 0 && o.name.indexOf('страйп') < 0 && o.name.indexOf('жаккард') < 0 && o.name.indexOf('твил') < 0 && o.name.indexOf('поли') < 0)) {
                analyticObject["Сатин"] = analyticObject["Сатин"] + 1
            }

            if(order.products.find(o => o.name.indexOf('сатин') >= 0 && o.name.indexOf('страйп') >= 0)) {
                analyticObject["Страйп-сатин"] = analyticObject["Страйп-сатин"] + 1
            }

            if(order.products.find(o => o.name.indexOf('сатин') >= 0 && o.name.indexOf('твил') >= 0)) {
                analyticObject["Твил-сатин"] = analyticObject["Твил-сатин"] + 1
            }

            if(order.products.find(o => o.name.indexOf('сатин') >= 0 && o.name.indexOf('поли') >= 0)) {
                analyticObject["Полисатин"] = analyticObject["Полисатин"] + 1
            }

            if(order.products.find(o => o.name.indexOf('сатин') >= 0 && o.name.indexOf('жаккард') >= 0)) {
                analyticObject["Сатин-жаккард"] = analyticObject["Сатин-жаккард"] + 1
            }

            if(order.products.find(o => o.name.indexOf('бяз') >= 0)) {
                analyticObject["Бязь"] = analyticObject["Бязь"] + 1
            }

            if(order.products.find(o => o.name.indexOf('варен') >= 0 || o.name.indexOf('варён') >= 0 || o.name.indexOf('хлоп') >= 0)) {
                analyticObject["Вареный хлопок"] = analyticObject["Вареный хлопок"] + 1
            }

            if(order.products.find(o => o.name.indexOf('микрофибр') >= 0)) {
                analyticObject["Микрофибра"] = analyticObject["Микрофибра"] + 1
            }

            if(order.products.find(o => o.name.indexOf('мулетон') >= 0)) {
                analyticObject["Мулетон"] = analyticObject["Мулетон"] + 1
            }
            
            if(order.products.find(o => o.name.indexOf('поплин') >= 0)) {
                analyticObject["Поплин"] = analyticObject["Поплин"] + 1
            }

            if(order.products.find(o => o.name.indexOf('перкал') >= 0)) {
                analyticObject["Перкаль"] = analyticObject["Перкаль"] + 1
            }

            if(order.products.find(o => o.name.indexOf('ранфор') >= 0)) {
                analyticObject["Ранфорс"] = analyticObject["Ранфорс"] + 1
            }

            if(order.products.find(o => o.name.indexOf('микросатин') >= 0)) {
                analyticObject["Микросатин"] = analyticObject["Микросатин"] + 1
            }

            if(order.products.find(o => o.name.indexOf('креп-ж') >= 0)) {
                analyticObject["Креп-жатка"] = analyticObject["Креп-жатка"] + 1
            }

            if(order.products.find(o => o.name.indexOf('креп-ж') < 0 && order.products.find(o => o.name.indexOf('жатка') >= 0))) {
                analyticObject["Жатка"] = analyticObject["Жатка"] + 1
            }

        }

        let html = `${headerComponent}
                <title>Аналитика спроса на товары</title>
            </head>
            <body>
                <div class="columns is-mobile">
                    <div class="column is-three-fifths is-offset-one-fifth">
                        <h1 class="title is-4 has-text-centered has-text-black">Аналитика спроса на простыни за ${year} год.</h1>
                    </div>
                </div>
                <div class="fixed-grid has-1-cols">
                    <div class="grid">`

        const myChart = new QuickChart()

        myChart.setConfig({
            type: 'bar',
            data: {
                labels: Object.keys(analyticObject),
                datasets: [
                    {
                        label: 'Получено, шт.',
                        data: Object.values(analyticObject),
                        fill: false
                    }
                ]
            }
        })
        .setWidth(800)
        .setHeight(400)
        .setBackgroundColor('transparent')

        const chartUrl = myChart.getUrl()

        html += `<div class="cell">
                    <img src="${chartUrl}">`

        html += `
                <table class="table is-fullwidth my-table">
                    <thead>
                        <tr>
                            <th class="has-text-left has-text-black">Продукт</th>
                            <th class="has-text-left has-text-black">Количество</th>
                        </tr>
                    </thead>
                    <tbody>`

        for(let key of Object.keys(analyticObject)) {

            html += `<tr>
                        <td class="has-text-black">${key}</td>
                        <td class="has-text-black">
                            ${analyticObject[key]} шт.
                        </td>
                    </tr>`

        }

        html += `</tbody>
            </table>`    
                        
        html += `</div>`

        html += `
                    </div>
                </div>
            ${footerComponent}`

        res.send(html)

    }

})

app.get('/get_year_dynamic/:year', async function (req, res) {

    const monthQuantityOrders = []
    const months = ['Янв', 'Фев', 'Мар', 'Апр', 'Май', 'Июн', 'Июл', 'Авг', 'Сен', 'Окт', 'Ноя', 'Дек']

    const year = req.params.year

    for(let i = 1; i <= 11; i++) {

        let monthFrom = ''
        let monthTo = ''

        if(i < 10 && i+1 < 10) {
            monthFrom = `0${i}`
            monthTo = `0${i+1}`
        } else if (i < 10 && i+1 >= 10) {
            monthFrom = `0${i}`
            monthTo = `${i+1}`
        } else if (i >= 10) {
            monthFrom = `${i}`
            monthTo = `${i+1}`
        }

        const response = await axios.post('https://api-seller.ozon.ru/v3/posting/fbs/list', {
            "dir": "asc",
            "filter": {
                "since": `${year}-${monthFrom}-01T00:00:00.000Z`,
                "status": "delivered",
                "to": `${year}-${monthTo}-01T23:59:59.999Z`,

            },
            "limit": 1000,
            "offset": 0
        }, {
            headers: {
                "Client-Id": process.env.OZON_CLIENT_ID,
                "Api-Key": process.env.OZON_API_KEY
            }
        })

        const filteredArray = response.data.result.postings.filter(o => {
            if(o.shipment_date.indexOf(`${year}-${monthFrom}-`) >= 0) {
                return o
            }
        })

        monthQuantityOrders.push(filteredArray.length)

    }

    const response = await axios.post('https://api-seller.ozon.ru/v3/posting/fbs/list', {
        "dir": "asc",
        "filter": {
            "since": `${year}-12-01T00:00:00.000Z`,
            "status": "delivered",
            "to": `${year}-12-31T23:59:59.999Z`,

        },
        "limit": 1000,
        "offset": 0
    }, {
        headers: {
            "Client-Id": process.env.OZON_CLIENT_ID,
            "Api-Key": process.env.OZON_API_KEY
        }
    })

    const filteredArray = response.data.result.postings.filter(o => {
        if(o.shipment_date.indexOf(`${year}-12-`) >= 0) {
            return o
        }
    })

    monthQuantityOrders.push(filteredArray.length)

    let html = `${headerComponent}
            <title>Динамика заказов в течение года</title>
            </head>
        <body>
            <div class="columns is-mobile">
                <div class="column is-three-fifths is-offset-one-fifth">
                    <h1 class="title is-4 has-text-centered has-text-black">Динамика заказов в течение года за ${year} год.</h1>
                </div>
            </div>`

    const myChart = new QuickChart()

    myChart.setConfig({
        type: 'line',
        data: {
            labels: ['Янв', 'Фев', 'Мар', 'Апр', 'Май', 'Июн', 'Июл', 'Авг', 'Сен', 'Окт', 'Ноя', 'Дек'],
            datasets: [{
                label: 'Количество заказов в месяц, шт.',
                data: monthQuantityOrders,
                fill: '#4e79a6'
            }]
        }
    })

    const chartUrl = myChart.getUrl()

    html += `<div class="columns is-mobile">
                <div class="column is-three-fifths is-offset-one-fifth">
                    <img src="${chartUrl}">
                </div>
            </div>
            <div class="columns is-mobile">
                    <div class="column is-three-fifths is-offset-one-fifth">
                        <table class="table is-fullwidth my-table">
                                        <thead>
                                            <tr>
                                                <th class="has-text-black">Месяц</th>
                                                <th class="has-text-left has-text-black">Количество</th>
                                            </tr>
                                        </thead>
                                        <tbody>`

    for(let i = 0; i < months.length; i++) {

        html += `<tr>
                    <td class="has-text-black">${months[i]}</td>
                    <td class="has-text-black">
                        ${monthQuantityOrders[i]} шт.
                    </td>
                </tr>`

    }

    html += `</tbody>
        </table>`    
                    
    html += `</div>
            </div>`

    html += footerComponent

    // res.json(monthQuantityOrders)

    res.send(html)

})

app.get('/get_income_analytic/:month/:product', async function (req, res) {

    let monthFrom = ''
    let monthTo = ''

    let year = new Date().getFullYear()

    if(Number(req.params.month) < 10 && Number(req.params.month) + 1 < 10) {

        monthFrom = `0${req.params.month}`
        monthTo = `0${Number(req.params.month) + 1}`

    }

    if(Number(req.params.month) < 10 && Number(req.params.month) + 1 === 10) {

        monthFrom = `0${req.params.month}`
        monthTo = `${Number(req.params.month) + 1}`

    }

    if(Number(req.params.month) >= 10 && Number(req.params.month) + 1 < 12) {

        monthFrom = `${req.params.month}`
        monthTo = `${Number(req.params.month) + 1}`

    }

    if(Number(req.params.month) === 12) {

        monthFrom = `${req.params.month}`
        monthTo = `01`
        year = Number(year) + 1

    }

    const response =  await axios.post('https://api-seller.ozon.ru/v3/posting/fbs/list', {
        "dir": "asc",
        "filter": {
            "delivery_method_id": [
                "23463726191000"
            ],
            "since": `${new Date().getFullYear()}-${monthFrom}-01T00:00:00.000Z`,
            "status": "delivered",
            "to": `${year}-${monthTo}-01T23:59:59.999Z`,
        },
        "limit": 1000,
        "offset": 0
    },{
        headers: {
            "Client-Id": process.env.OZON_CLIENT_ID,
            "Api-Key": process.env.OZON_API_KEY
        }
    })

    let ozonDelOrders = response.data.result.postings

    ozonDelOrders = ozonDelOrders.filter(o => {
        if(o.products.find(i => i.name.indexOf(req.params.product) >= 0)) {
            return o
        }
    })

    let orderIncomeInfo = []

    for(let order of ozonDelOrders) {

        const response = await axios.post('https://api-seller.ozon.ru/v3/finance/transaction/list', {

            "filter": {
                "date": {
                    "from": `${new Date().getFullYear()}-${monthFrom}-01T00:00:00.000Z`,
                    "to": `${year}-${monthTo}-01T23:59:59.999Z`
                },
            "operation_type": [ ],
            "posting_number": order.posting_number,
            "transaction_type": "all"
            },
            "page": 1,
            "page_size": 1000

        }, {
            headers: {
                "Client-Id": process.env.OZON_CLIENT_ID,
                "Api-Key": process.env.OZON_API_KEY
            }
        })

        orderIncomeInfo.push(response.data.result)

    }

    let html = `${headerComponent}
            <title>Аналитика дохода с заказов по указанному продукту</title>
        </head>
    <body>`

    html += `<div class="columns is-mobile has-text-black">
                    <div class="column is-three-fifths is-offset-one-fifth">
                        <table class="table is-fullwidth my-table">
                            <thead>
                                <tr>
                                    <th class="has-text-left has-text-black">Номер заказа</th>
                                    <th class="has-text-left has-text-black">Наименование</th>
                                    <th class="has-text-left has-text-black">Сумма к начислению</th>
                                    <th class="has-text-left has-text-black">Сумма заказа</th>
                                    <th class="has-text-left has-text-black">% начисления от суммы заказа</th>
                                </tr>
                            </thead>
                            <tbody>`

    for(let item of orderIncomeInfo) {

        let element = item.operations.find(o => o.operation_type_name === 'Доставка покупателю')

        if(element.accruals_for_sale > 0) {

            html += `<tr>
                        <td class="has-text-black">${element.posting.posting_number}</td>
                        <td class="has-text-black">${element.items[0].name}</td>
                        <td class="has-text-black">
                            ${Math.round(element.amount)} руб.
                        </td>
                        <td class="has-text-black">
                            ${Math.round(element.accruals_for_sale)} руб.
                        </td>
                        <td class="has-text-black">
                            ${Math.round((Math.round(element.amount)/Math.round(element.accruals_for_sale)) * 100)} %
                        </td>
                    </tr>`

        }

    }

    html += `</tbody>
        </table>`    
                    
    html += `</div>
            </div>`

    html += footerComponent

    res.send(html)

    // res.json(orderIncomeInfo)

})

app.get('/api_test', async function (req, res) {

    const ya_response = await axios.post('https://b2b-authproxy.taxi.yandex.net/api/b2b/platform/pickup-points/list', {
        "geo_id": 5,
        "type": "pickup_point",
        "payment_method": "already_paid",
        "available_for_dropoff": false,
        "is_yandex_branded": false,
        "is_not_branded_partner_station": false,
        "is_post_office": false,
        "payment_methods": [
            "already_paid"
        ],
        "pickup_services": {
            "is_fitting_allowed": false,
            "is_partial_refuse_allowed": false,
            "is_paperless_pickup_allowed": false,
            "is_unboxing_allowed": false
        }
    },{
        headers: {
            "Authorization": "Bearer y0__xDt8OShCBix9BwglviD3BQOak0tUo1gQM_yL80qweGmZnEfwg"
        }
    })

    let ozResponse = await axios.post('https://api-seller.ozon.ru/v1/delivery-method/list', {
        "filter": {
            "status": "ACTIVE",
            "warehouse_id": 21474696078000
        },
        "limit": 100,
        "offset": 0
    }, {
        headers: {
            'Host':'api-seller.ozon.ru',
            'Client-Id':`${process.env.OZON_CLIENT_ID}`,
            'Api-Key':`${process.env.OZON_API_KEY}`,
            'Content-Type':'application/json'
        }
    })

    res.json({ozon: ozResponse.data.result, yandex: ya_response.data})
    
})

app.get('/cdek_test/:from/:to', async function (req, res) {

    const authHandle = async () => {

        const response = await axios.post(`https://api.cdek.ru/v2/oauth/token?grant_type=client_credentials&client_id=${process.env.CDEK_ID}&client_secret=${process.env.CDEK_PASS}`)

        return response.data

    }

    const bearerToken = await authHandle()

    // const tariffResponse = await axios.get(`${process.env.CDEK_API_URL}/v2/calculator/alltariffs?x-user-lang=rus`, {

    //     headers: {

    //         Authorization: `Bearer ${bearerToken.access_token}`

    //     }

    //})

    console.log({
        from: req.params.from,
        to: req.params.to
    })

    const calculateResponse = await axios.post('https://api.cdek.ru/v2/calculator/tariff', {

        type: 1,
        tariff_code: 136,
        from_location: {
            code: CITIES.find(o => o.city === req.params.from).code,
            city: req.params.from,
            contragent_type: 'LEGAL_ENTITY',
            longitude: CITIES.find(o => o.city === req.params.from).longitude,
            latitude: CITIES.find(o => o.city === req.params.from).latitude
        },
        to_location: {
            code: CITIES.find(o => o.city === req.params.to).code,
            city: req.params.to,
            contragent_type: 'INDIVIDUAL',
            longitude: CITIES.find(o => o.city === req.params.to).longitude,
            latitude: CITIES.find(o => o.city === req.params.to).latitude
        },
        services: [
            {
                code: 'INSURANCE',
                parameter: 1245
            }
        ],
        packages: {
            weight: 1500,
            length: 77,
            width: 77,
            height: 2
        }

    }, {
        headers: {

            Authorization: `Bearer ${bearerToken.access_token}`

        }
    })

    res.json(calculateResponse.data)
    
})

app.get('/revenue', async (req, res) => {

    let i = 0
    let hasNext = true

    const REPORT_PATH = 'REPORT.xlsx'

    let orders = []

    while(hasNext) {

        const response = await axios.post('https://api-seller.ozon.ru/v3/posting/fbs/list', {

            "dir": "ASC",
            "filter": {
                "since": "2024-01-01T00:00:00.000Z",
                "status": "delivered",
                "to": "2024-12-31T23:59:59.999Z",
            },
            "limit": 1000,
            "offset": i * 1000

        }, {

            headers: {

                'Client-Id': process.env.OZON_CLIENT_ID,
                'Api-Key': process.env.OZON_API_KEY

            }

        })

        response.data.result.postings.forEach(el => {
            orders.push(el)
        })
        hasNext = response.data.result.has_next
        i += 1

    }

    const accrualsList = []

    const wb = new exl.Workbook()

    await wb.xlsx.readFile(REPORT_PATH)

    const ws = wb.getWorksheet('Начисления')

    ws.eachRow((row, rowNumber) => {

        if(rowNumber <= 2) return

        const existingEntry = accrualsList.find(o => o.posting_number === row.getCell(1).value)

        if(existingEntry) {

            existingEntry.accruals.push({
                accrual_name: row.getCell(4).value,
                accrual_value: row.getCell(15).value
            })

        }

        if(!existingEntry) {

            accrualsList.push({
                posting_number: row.getCell(1).value,
                accruals: [
                    {
                        accrual_name: row.getCell(4).value,
                        accrual_value: row.getCell(15).value
                    }
                ]
            })

        }

    })

    console.log(orders.length)

    orders.forEach(o => {

        const rec = accrualsList.find(i => i.posting_number === o.posting_number)

        if(!rec) return

        if(rec) {

            let _temp = 0

            rec.accruals.forEach(a => {

                _temp += a.accrual_value

            })

            o.revenue = Math.round(_temp)

        }

    })

    orders = orders.filter(o => Object.hasOwn(o, 'revenue'))

    console.log(orders.length)

    let revenuesObject = {
        sewingRevenue: 0,
        mattressesRevenue: 0,
        otherRevenue: 0,
        totalRevenue: 0
    }

    let ordersObject = {
        sewingOrders: [],
        mattressesOrders: [],
        otherOrders: []
    }

    orders.forEach(o => {

        if (o.products.find(i => (i.name.toLowerCase().indexOf('постельно') >= 0 || i.name.toLowerCase().indexOf('простын') >= 0 || i.name.toLowerCase().indexOf('пододе') >= 0 || i.name.toLowerCase().indexOf('наволоч') >= 0) && i.name.toLowerCase().indexOf('матрас') < 0)) {

            revenuesObject.sewingRevenue += o.revenue
            ordersObject.sewingOrders.push(o)

        }

        if (o.products.find(i => (i.name.toLowerCase().indexOf('постельно') < 0 && i.name.toLowerCase().indexOf('простын') < 0 && i.name.toLowerCase().indexOf('пододе') < 0 && i.name.toLowerCase().indexOf('наволоч') < 0) && i.name.toLowerCase().indexOf('матрас') >= 0)) {

            revenuesObject.mattressesRevenue += o.revenue
            ordersObject.mattressesOrders.push(o)

        }

        if (o.products.find(i => (i.name.toLowerCase().indexOf('постельно') < 0 && i.name.toLowerCase().indexOf('простын') < 0 && i.name.toLowerCase().indexOf('пододе') < 0 && i.name.toLowerCase().indexOf('наволоч') < 0) && i.name.toLowerCase().indexOf('матрас') < 0)) {

            revenuesObject.otherRevenue += o.revenue
            ordersObject.otherOrders.push(o)

        }

        revenuesObject.totalRevenue += o.revenue

    })

    findMatchesByPostingNumber(ordersObject.sewingOrders, ordersObject.otherOrders).forEach(el => {

        revenuesObject.totalRevenue += el.revenue

    })

    res.json(
        {
            orderTotals: {
                sewing: ordersObject.sewingOrders.length,
                mattresses: ordersObject.mattressesOrders.length,
                other: ordersObject.otherOrders.length
            },
            revenueTotals: revenuesObject,
            sewingProportion: revenuesObject.sewingRevenue / revenuesObject.totalRevenue * 100,
            mattressesProportion: revenuesObject.mattressesRevenue / revenuesObject.totalRevenue * 100,
            otherProportion: revenuesObject.otherRevenue / revenuesObject.totalRevenue * 100
        }
    )

})

app.listen(3030)