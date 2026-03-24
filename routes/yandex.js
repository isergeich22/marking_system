const router = require('express').Router()
const exl = require('exceljs')
const axios = require('axios')
const fs = require('fs')
const buttons = require('../config').buttons

const dbsId = process.env.YANDEX_DBS_ID
const fbsId = process.env.YANDEX_FBS_ID

router.get('/yandex', async function(req, res){

    const nat_cat = []
    const nat_catGtins = []
    const nat_catNames = []
    let ya_orders = []
    const new_items = []
    const current_items = []
    let names = []

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

        console.log(ya_orders[i].vendor)

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
                'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase() === 'ТЕНСЕЛЬ' ? 'ТЕНСЕЛ' : response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                'productType': 'ПОДОДЕЯЛЬНИК'
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
                    'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase() === 'ТЕНСЕЛЬ' ? 'ТЕНСЕЛ' : response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
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
                    'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase() === 'ТЕНСЕЛЬ' ? 'ТЕНСЕЛ' : response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
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
                    'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase() === 'ТЕНСЕЛЬ' ? 'ТЕНСЕЛ' : response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                    'productType': 'НАВОЛОЧКА'
                })

            } else {

                names.push({
                    'vendor': ya_orders[i].vendor,
                    'name': response.data.result[0].name.replace(/\u00A0/g, ' ') // заменить неразрывные пробелы на обычные
                                                    .trim()                  // убрать пробелы по краям
                                                    .replace(/\s+/g, ' '),
                    'size': response.data.result[0].attributes.find(o => o.id === 6772).values[0].value,
                    'color': response.data.result[0].attributes.find(o => o.id === 10096).values[0].value.toUpperCase(),
                    'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase() === 'ТЕНСЕЛЬ' ? 'ТЕНСЕЛ' : response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                    'productType': 'НАВОЛОЧКА'
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
                            'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase() === 'ТЕНСЕЛЬ' ? 'ТЕНСЕЛ' : response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                            'productType': 'КОМПЛЕКТ ПОСТЕЛЬНОГО БЕЛЬЯ'
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
                            'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase() === 'ТЕНСЕЛЬ' ? 'ТЕНСЕЛ' : response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                            'productType': 'КОМПЛЕКТ ПОСТЕЛЬНОГО БЕЛЬЯ'
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
                            'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase() === 'ТЕНСЕЛЬ' ? 'ТЕНСЕЛ' : response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                            'productType': 'КОМПЛЕКТ ПОСТЕЛЬНОГО БЕЛЬЯ'
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
                        'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase() === 'ТЕНСЕЛЬ' ? 'ТЕНСЕЛ' : response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                        'productType': 'КОМПЛЕКТ ПОСТЕЛЬНОГО БЕЛЬЯ'
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
                            'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase() === 'ТЕНСЕЛЬ' ? 'ТЕНСЕЛ' : response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                            'productType': 'КОМПЛЕКТ ПОСТЕЛЬНОГО БЕЛЬЯ'
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
                            'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase() === 'ТЕНСЕЛЬ' ? 'ТЕНСЕЛ' : response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                            'productType': 'КОМПЛЕКТ ПОСТЕЛЬНОГО БЕЛЬЯ'
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
                            'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase() === 'ТЕНСЕЛЬ' ? 'ТЕНСЕЛ' : response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                            'productType': 'КОМПЛЕКТ ПОСТЕЛЬНОГО БЕЛЬЯ'
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
                        'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase() === 'ТЕНСЕЛЬ' ? 'ТЕНСЕЛ' : response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                        'productType': 'КОМПЛЕКТ ПОСТЕЛЬНОГО БЕЛЬЯ'
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

    async function createImport(array) {

        const fileName = './public/IMPORT_TNVED_6302.xlsx'

        const wb = new exl.Workbook()

        await wb.xlsx.readFile(fileName)

        const ws = wb.getWorksheet('IMPORT_TNVED_6302')

        let cellNumber = 5

        for(let i = 0; i < array.length; i++) {

            ws.getCell(`B${cellNumber}`).value = 6302
            names.find(o => o.name === array[i]).productType === 'КОМПЛЕКТ ПОСТЕЛЬНОГО БЕЛЬЯ' ? ws.getCell(`C${cellNumber}`).value = 'Да' : ws.getCell(`C${cellNumber}`).value = 'Нет'
            ws.getCell(`D${cellNumber}`).value = names.find(o => o.name === array[i]).name
            ws.getCell(`E${cellNumber}`).value = 'Ивановский текстиль'
            ws.getCell(`F${cellNumber}`).value = 'Артикул'
            ws.getCell(`G${cellNumber}`).value = names.find(o => o.name === array[i]).vendor
            ws.getCell(`H${cellNumber}`).value = names.find(o => o.name === array[i]).productType
            ws.getCell(`I${cellNumber}`).value = names.find(o => o.name === array[i]).color
            ws.getCell(`J${cellNumber}`).value = 'ВЗРОСЛЫЙ'

            if(names.find(o => o.name === array[i]).cloth.includes('ЖАТКА')) ws.getCell(`K${cellNumber}`).value = 'КРЕП'
            if(names.find(o => o.name === array[i]).cloth === 'ВАРЕНЫЙ ХЛОПОК') ws.getCell(`K${cellNumber}`).value = 'ХЛОПКОВАЯ ТКАНЬ'
            if(names.find(o => o.name === array[i]).cloth === 'ЛЕН' || names.find(o => o.name === array[i]).cloth === 'ЛЁН') ws.getCell(`K${cellNumber}`).value = 'ЛЬНЯНАЯ ТКАНЬ'
            if(names.find(o => o.name === array[i]).cloth === 'СТРАЙП САТИН') ws.getCell(`K${cellNumber}`).value = 'СТРАЙП-САТИН'
            if(names.find(o => o.name === array[i]).cloth === 'САТИН ЛЮКС') ws.getCell(`K${cellNumber}`).value = 'САТИН'
            if(names.find(o => o.name === array[i]).cloth !== 'ЖАТКА' && names.find(o => o.name === array[i]).cloth !== 'КРЕП-ЖАТКА' && names.find(o => o.name === array[i]).cloth !== 'САТИН ЛЮКС' && names.find(o => o.name === array[i]).cloth !== 'СТРАЙП САТИН' && names.find(o => o.name === array[i]).cloth !== 'ВАРЕНЫЙ ХЛОПОК' && names.find(o => o.name === array[i]).cloth !== 'ЛЕН' && names.find(o => o.name === array[i]).cloth !== 'ЛЁН') ws.getCell(`K${cellNumber}`).value = names.find(o => o.name === array[i]).cloth

            if(names.find(o => o.name === array[i]).cloth === 'ПОЛИСАТИН' || names.find(o => o.name === array[i]).cloth.includes('ЖАТКА')) ws.getCell(`L${cellNumber}`).value = '100% Полиэстер'

            if(names.find(o => o.name === array[i]).cloth.includes('ТЕНСЕЛ')) ws.getCell(`L${cellNumber}`).value = '100% Лиоцелл'
            if(names.find(o => o.name === array[i]).cloth === 'ЛЕН' || names.find(o => o.name === array[i]).cloth === 'ЛЁН') ws.getCell(`L${cellNumber}`).value = '100% Лен'
            if(names.find(o => o.name === array[i]).cloth !== 'ЖАТКА' && names.find(o => o.name === array[i]).cloth !== 'КРЕП-ЖАТКА' && names.find(o => o.name === array[i]).cloth !== 'ПОЛИСАТИН' && names.find(o => o.name === array[i]).cloth !== 'ТЕНСЕЛ' && names.find(o => o.name === array[i]).cloth !== 'ЛЕН' && names.find(o => o.name === array[i]).cloth !== 'ЛЁН') ws.getCell(`L${cellNumber}`).value = '100% Хлопок'

            ws.getCell(`M${cellNumber}`).value = names.find(o => o.name === array[i]).size
            if(names.find(o => o.name === array[i]).cloth !== 'ЖАТКА' && names.find(o => o.name === array[i]).cloth !== 'КРЕП-ЖАТКА' && names.find(o => o.name === array[i]).cloth !== 'ПОЛИСАТИН' && names.find(o => o.name === array[i]).cloth !== 'ТЕНСЕЛ' && names.find(o => o.name === array[i]).cloth !== 'ЛЕН' && names.find(o => o.name === array[i]).cloth !== 'ЛЁН') {

                ws.getCell(`N${cellNumber}`).value = '6302310009'

            } else {

                ws.getCell(`N${cellNumber}`).value = '6302299000'

            }
            ws.getCell(`O${cellNumber}`).value = 'ТР ТС 017/2011 "О безопасности продукции легкой промышленности"'
            ws.getCell(`P${cellNumber}`).value = 'На модерации'

            cellNumber++

        }

        ws.unMergeCells('F2')

        ws.getCell('G2').value = '13914'

        ws.getCell('G2').fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor:{argb:'E3E3E3'}
        }

        ws.getCell('G2').font = {
            size: 10,
            name: 'Arial'
        }

        ws.getCell('G2').alignment = {
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

    const items = names.map(el => ({ name: el.name, isNew: new_items.indexOf(el.name) >= 0 }))
    res.render('import-products', { title: 'Импорт - Яндекс.Маркет', items, marksOrderUrl: '/yandex_marks_order', buttons })

})

router.get('/yandex_marks_order', async function (req, res){

    let ya_orders = []
    const nat_cat = []
    const gtins = []
    let names = []

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

    let answer = null

    async function getOrders(clientId) {

        let response = await axios.get(`https://api.partner.market.yandex.ru/campaigns/${clientId}/orders?status=PROCESSING&substatus=STARTED&pageSize=50`, {
            headers: {
                'Authorization': `Bearer ${process.env.YANDEX_API_KEY}`
            }
        })

        let result = response.data

        console.log(result)

        answer = result

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

                            if(ya_orders.findIndex(o => o.vendor === el.offerId) >= 0) {

                                ya_orders.find(o => o.vendor === el.offerId).quantity += Number(el.count)

                            }

                            if(ya_orders.findIndex(o => o.vendor === el.offerId) < 0) {

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
                                    <releaseMethodType>PRODUCTION</releaseMethodType>
                                    <createMethodType>SELF_MADE</createMethodType>
                                    <productionOrderId>YANDEX_${i+1}</productionOrderId>
                                    <products>`

                    for(let j = 0; j < List[i].length; j++) {
                        if(nat_cat.indexOf(List[i][j]) >= 0) {

                            if(nat_cat[nat_cat.indexOf(List[i][j])].includes('КПБ')) {

                                content += `<product>
                                                <gtin>0${gtins[nat_cat.indexOf(List[i][j])]}</gtin>
                                                <quantity>${Quantity[i][j]}</quantity>
                                                <serialNumberType>OPERATOR</serialNumberType>
                                                <cisType>BUNDLE</cisType>
                                                <templateId>10</templateId>
                                            </product>`

                            }

                            if(nat_cat[nat_cat.indexOf(List[i][j])].indexOf('КПБ') < 0) {

                                content += `<product>
                                                <gtin>0${gtins[nat_cat.indexOf(List[i][j])]}</gtin>
                                                <quantity>${Quantity[i][j]}</quantity>
                                                <serialNumberType>OPERATOR</serialNumberType>
                                                <cisType>UNIT</cisType>
                                                <templateId>10</templateId>
                                            </product>`

                            }

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

    const oz_orders = ya_orders
    const orders = oz_orders.map(o => ({ name: o.name, quantity: o.quantity }))
    res.render('marks-order', { title: 'Заказ маркировки - Яндекс.Маркет', orders, buttons })

})

module.exports = router
