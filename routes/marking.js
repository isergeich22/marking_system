const router = require('express').Router()
const exl = require('exceljs')
const fs = require('fs')
const cio = require('cheerio')
const axios = require('axios')
const buttons = require('../config').buttons

router.get('/input_own', async function(req, res){


    let remark_date = ''

    const date_ob = new Date()

    let year = date_ob.getFullYear()

    let month = date_ob.getMonth()+1

    let day = date_ob.getDate()

    month < 10 ? month = '0' + month : month

    day < 10 ? day = '0' + day : day

    const production_date = year + '-' + month + '-' + day

    let content = `<introduce_rf version="9">
                    <trade_participant_inn>${process.env.ORG_INN}</trade_participant_inn>
                    <producer_inn>${process.env.ORG_INN}</producer_inn>
                    <owner_inn>${process.env.ORG_INN}</owner_inn>
                    <production_date>${production_date}</production_date>
                    <production_order>OWN_PRODUCTION</production_order>
                        <products_list>`

    const marks = []

    const wb = new exl.Workbook()

    await wb.xlsx.readFile('./public/inputinsale/marks.xlsx')

    const ws = wb.getWorksheet(1)

    ws.eachRow((row, rowNumber) => {

        if (rowNumber < 3) {

            return

        }

        if (rowNumber >= 3) {

            marks.push({
                mark: row.values[1],
                name: row.values[10]
            })

        }

    })

    // console.log(marks)

    let tnved = ''

    marks.forEach(el => {

        if(el.name.toUpperCase().indexOf('ЖАТКА') < 0 && el.name.toUpperCase().indexOf('КРЕП-ЖАТКА') < 0 && el.name.toUpperCase().indexOf('ПОЛИСАТИН') < 0 && el.name.toUpperCase().indexOf('ТЕНСЕЛ') < 0 && el.name.toUpperCase().indexOf('ЛЕН') < 0 && el.name.toUpperCase().indexOf('ЛЁН') < 0) {

            tnved = '6302310009'

        } else {

            tnved = '6302299000'

        }

        if(el.mark.length === 31) {
            content += `<product>
                            <ki><![CDATA[${el.mark}]]></ki>
                            <production_date>${production_date}</production_date>
                            <tnved_code>${tnved}</tnved_code>
                            <certificate_document_data>
                                <product>
                                    <certificate_type>${process.env.CERT_TYPE}</certificate_type>
                                    <certificate_number>${process.env.CERT_NUMBER}</certificate_number>
                                    <certificate_date>${process.env.CERT_DATE}</certificate_date>
                                </product>
                            </certificate_document_data>
                        </product>`
        }
    })

    // console.log(content)

    content += `    </products_list>
            </introduce_rf>`

    fs.writeFileSync('./public/inputinsale/own.xml', content)

    const viewMarks = marks
        .filter(el => el.mark.length === 31)
        .map(el => {
            const name = el.name.toUpperCase()
            const isNatural = name.indexOf('ЖАТКА') >= 0 || name.indexOf('КРЕП-ЖАТКА') >= 0 ||
                name.indexOf('ПОЛИСАТИН') >= 0 || name.indexOf('ТЕНСЕЛ') >= 0 ||
                name.indexOf('ЛЕН') >= 0 || name.indexOf('ЛЁН') >= 0
            return {
                mark: el.mark.replace(/</g, '&lt;'),
                tnved: isNatural ? '6302299000' : '6302310009'
            }
        })

    res.render('input-own', { title: 'Ввод в оборот. Производство РФ', marks: viewMarks, buttons })

})

router.get('/sale_ozon', async function(req, res){

    const actualMarksFile = './public/actual_marks.xlsx'

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
                        <products_list>`

    const wb = new exl.Workbook()

    async function getActualList() {

        const [marks, status] = [[], []]

        await wb.xlsx.readFile(actualMarksFile)

        const ws = wb.getWorksheet('Sheet0')

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

        const ws = wb.getWorksheet('Лист_1')

        const [c2, c4, c7] = [ws.getColumn(2), ws.getColumn(4), ws.getColumn(7)]

        c2.eachCell(c => {
            let str = c.value
            consignmentDate.push(str.replace(str.substring(10), ''))
        })

        c4.eachCell(c => {
            let str = c.value
            consignmentNumbers.push(str.replace('MT00-', ''))
        })

        c7.eachCell(c => {
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

        let response = await axios.post('https://api-seller.ozon.ru/v3/posting/fbs/list', {            
                'dir': 'asc',
                'filter':{
                    'since':'2023-10-01T00:00:00Z',
                    'to':'2023-12-31T00:00:00Z',
                    'status':'delivered'
                },
                'limit':1000,
                'offset':0
            },
            {
                headers: {
                    'Host': 'api-seller.ozon.ru',
                    'Client-Id':`${process.env.OZON_CLIENT_ID}`,
                    'Api-Key':`${process.env.OZON_API_KEY}`,
                    'Content-Type': 'application/json'
                }
            }
        )

        let result = response.data

        console.log(result)

        result.result.postings.forEach(e => {
            console.log(e.posting_number)
            let orderNumber = e.posting_number
            let products = []
            e.products.forEach(el => {
                let marks = []
                console.log(el)
                el.mandatory_mark.forEach(elem => {
                    if(!elem) marks.push('')
                    marks.push(elem)
                })
                products.push({
                    name: el.name,
                    marksList: marks,
                    price: el.price
                })
            })

            if(products.find(o => o.marksList.indexOf('') < 0)) {
                
                let obj = {
                    orderNumber: orderNumber,
                    productsList: products
                }

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

    const rows = []
    equals.forEach(e => {
        e.productsList.forEach(el => {
            if(el.marksList.length > 0 && el.marksList.indexOf('') < 0) {
                for(let i = 0; i < el.marksList.length; i++) {
                    const cons = consignments.find(c => c.orderNumber == e.orderNumber)
                    rows.push({
                        mark: el.marksList[i].replace(/</g, '&lt;'),
                        price: `${(el.price).replace(el.price.substring(el.price.indexOf('.')), '')}00`,
                        consignmentNumber: cons.consignmentNumber,
                        consignmentDate: cons.consignmentDate
                    })
                }
            }
        })
    })

    res.render('sale', { title: 'Вывод из оборота — Ozon', fileName: fileName.substring(fileName.lastIndexOf('/') + 1), rows, buttons })

})

router.get('/sale_wb', async function(req, res){

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

        const ws = wb.getWorksheet('Лист_1')

        const [c2, c4, c7] = [ws.getColumn(2), ws.getColumn(4), ws.getColumn(7)]

        const [consDates, consNumbers, orderNumbers, wbNumbers] = [[], [], [], [], []]

        const numbers = []

        const consignments = []

        c2.eachCell(c => {
            let str = c.value.replace(c.value.substring(10), '')
            let date = str.split('.')
            consDates.push(`${date[2]}-${date[1]}-${date[0]}`)
        })

        // console.log(consDates)

        c4.eachCell(c => {
            consNumbers.push(c.value.trim().replace('MT00-0', ''))
        })

        c7.eachCell(c => {
            if(c.value == null) {
                numbers.push(c.value)
            }
            // console.log(c.value)
            if(c.value != null) {
                numbers.push(c.value.trim().replace('WB ', '').replace('WB-', '').replace('ozon ', ''))
                wbNumbers.push(c.value.trim().replace('WB ', '').replace('WB-', '').replace('ozon ', ''))
                orderNumbers.push(c.value.trim().replace('WB ', '').replace('WB-', '').replace('ozon ', ''))
            }
        })

        console.log(wbNumbers)

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
    console.log(consignments)

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

    const rows = equals.map(item => {
        let price = ''
        if((item.consignmentPrice.toString()).indexOf('.') >= 0) {
            let arr = (item.consignmentPrice.toString()).split('.')
            price = arr[0] + arr[1]
        } else {
            price = item.consignmentPrice + '00'
        }
        return {
            mark: item.consignmentCis.replace(/</g, '&lt;'),
            price,
            consignmentNumber: item.consignmentNumber,
            consignmentDate: item.consignmentDate
        }
    })

    res.render('sale', { title: 'Вывод из оборота — Wildberries', fileName: fileName.substring(fileName.lastIndexOf('/') + 1), rows, buttons })

})

router.get('/test_features', async function(req, res){

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
            if(i%10 === 5 && products[i].indexOf('Готов к вводу в оборот') < 0 && products[i].indexOf('Опубликована') < 0 && products[i] !== '') {
                actual_products.push(products[i])
            }
        }

        console.log(actual_products)

        divs.each((i, elem) => {
            gtins.push(content(elem).text())
        })

        for(let i = 0; i < gtins.length; i++) {
            if(gtins[i].indexOf('046') >= 0) {
                actual_gtins.push(gtins[i].replace('0', ''))
            }
        }

    }


    await getList(fileContent)

    for(let i = 0; i < actual_products.length; i++) {

        newProducts.push({

            'name': actual_products[i],
            'gtin': actual_gtins[i]

        })

    }

    let dateString = ''

    if(new Date().getDate() <= 9) {

        dateString += `0${new Date().getDate()}-`

    }

    if(new Date().getDate() >= 10) {

        dateString += `${new Date().getDate()}-`

    }

    if(new Date().getMonth() <= 9) {

        dateString += `0${new Date().getMonth() + 1}-`

    }

    if(new Date().getMonth() >= 10) {

        dateString += `${new Date().getMonth() + 1}-`

    }

    dateString += `${new Date().getFullYear()}`

    async function updateShortReport() {

        const fileName = './public/Краткий отчет.xlsx'

        const wb = new exl.Workbook()

        await wb.xlsx.readFile(fileName)

        const ws = wb.getWorksheet(1)

        const firstColumn = ws.getColumn(1)

        let cellNumber = firstColumn.values.length

        console.log(cellNumber)

        for(let i = 0; i < newProducts.length; i++) {

            const row = ws.getRow(cellNumber)
            row.getCell(1).value = newProducts[i].gtin
            row.getCell(2).value = newProducts[i].name
            row.getCell(3).value = dateString
            row.commit()

            cellNumber++

        }

        await wb.xlsx.writeFile(fileName)

    }

    await updateShortReport()

    res.render('test-features', { title: 'Обновление краткого отчёта', products: newProducts, buttons })

})

router.get('/clear_duplicate', async function(req, res){

    const workbook = new exl.Workbook()

})

module.exports = router
