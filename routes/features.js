const router = require('express').Router()
const exl = require('exceljs')
const fs = require('fs')
const axios = require('axios')
const buttons = require('../config').buttons

router.get('/personal_orders', async function(req, res) {

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

            name: row.values[2].trim(),
            quantity: row.values[1],
            vendor: row.values[3]

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

        if(nat_cat.find(o => orderProducts[i].name.indexOf(o) < 0)) {

            difference.push(orderProducts[i])

        }

    }

    console.log(difference)

    await workbook.xlsx.readFile(unloadFile)

    const ws_3 = workbook.getWorksheet('result')

    ws_3.eachRow((row, rowNumber) => {

        if(rowNumber < 2) return
        full_cat.push({
            name: row.values[8],
            vendor: row.values[10]
        })

    })

    // difference = difference.filter((o) => {

    //     if(full_cat.findIndex(i => i.vendor === o.vendor) < 0) {

    //         return o

    //     }

    // })

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

    console.log(new_items)

    async function createImport(array) {

        const fileName = './public/IMPORT_TNVED_6302 (3).xlsx'

        const wb = new exl.Workbook()

        await wb.xlsx.readFile(fileName)

        const ws = wb.getWorksheet('IMPORT_TNVED_6302')

        let cellNumber = 5

        for(let i = 0; i < array.length; i++) {

            ws.getCell(`A${cellNumber}`).value = 6302
            ws.getCell(`B${cellNumber}`).value = names.find(o => o.name.indexOf(array[i]) >= 0).name
            ws.getCell(`C${cellNumber}`).value = 'Ивановский текстиль'
            ws.getCell(`D${cellNumber}`).value = 'Артикул'
            ws.getCell(`E${cellNumber}`).value = names.find(o => o.name.indexOf(array[i]) >= 0).vendor
            ws.getCell(`F${cellNumber}`).value = names.find(o => o.name.indexOf(array[i]) >= 0).productType
            ws.getCell(`G${cellNumber}`).value = names.find(o => o.name.indexOf(array[i]) >= 0).color
            ws.getCell(`H${cellNumber}`).value = 'ВЗРОСЛЫЙ'

            if(names.find(o => o.name.indexOf(array[i]) >= 0).cloth === 'КРЕП-ЖАТКА' || names.find(o => o.name.indexOf(array[i]) >= 0).cloth === 'КРЕП ЖАТКА') ws.getCell(`I${cellNumber}`).value = 'КРЕП'
            if(names.find(o => o.name.indexOf(array[i]) >= 0).cloth === 'ВАРЕНЫЙ ХЛОПОК') ws.getCell(`I${cellNumber}`).value = 'ХЛОПКОВАЯ ТКАНЬ'
            if(names.find(o => o.name.indexOf(array[i]) >= 0).cloth === 'ЛЕН' || names.find(o => o.name.indexOf(array[i]) >= 0).cloth === 'ЛЁН') ws.getCell(`I${cellNumber}`).value = 'ЛЬНЯНАЯ ТКАНЬ'
            if(names.find(o => o.name.indexOf(array[i]) >= 0).cloth === 'СТРАЙП САТИН') ws.getCell(`I${cellNumber}`).value = 'СТРАЙП-САТИН'
            if(names.find(o => o.name.indexOf(array[i]) >= 0).cloth === 'САТИН ЛЮКС') ws.getCell(`I${cellNumber}`).value = 'САТИН'
            if(names.find(o => o.name.indexOf(array[i]) >= 0).cloth !== 'САТИН ЛЮКС' && names.find(o => o.name.indexOf(array[i]) >= 0).cloth !== 'СТРАЙП САТИН' && names.find(o => o.name.indexOf(array[i]) >= 0).cloth !== 'ВАРЕНЫЙ ХЛОПОК' && names.find(o => o.name.indexOf(array[i]) >= 0).cloth !== 'ЛЕН' && names.find(o => o.name.indexOf(array[i]) >= 0).cloth !== 'ЛЁН') ws.getCell(`I${cellNumber}`).value = names.find(o => o.name.indexOf(array[i]) >= 0).cloth

            if(names.find(o => o.name.indexOf(array[i]) >= 0).cloth === 'ПОЛИСАТИН') ws.getCell(`J${cellNumber}`).value = '100% Полиэстер'

            if(names.find(o => o.name.indexOf(array[i]) >= 0).cloth === 'ТЕНСЕЛЬ') ws.getCell(`J${cellNumber}`).value = '100% Лиоцелл'
            if(names.find(o => o.name.indexOf(array[i]) >= 0).cloth === 'ЛЕН' || names.find(o => o.name.indexOf(array[i]) >= 0).cloth === 'ЛЁН') ws.getCell(`J${cellNumber}`).value = '100% Лен'
            if(names.find(o => o.name.indexOf(array[i]) >= 0).cloth !== 'КРЕП-ЖАТКА' && names.find(o => o.name.indexOf(array[i]) >= 0).cloth !== 'КРЕП ЖАТКА' && names.find(o => o.name.indexOf(array[i]) >= 0).cloth !== 'ПОЛИСАТИН' && names.find(o => o.name.indexOf(array[i]) >= 0).cloth !== 'ТЕНСЕЛЬ' && names.find(o => o.name.indexOf(array[i]) >= 0).cloth !== 'ЛЕН' && names.find(o => o.name.indexOf(array[i]) >= 0).cloth !== 'ЛЁН') ws.getCell(`J${cellNumber}`).value = '100% Хлопок'

            ws.getCell(`K${cellNumber}`).value = names.find(o => o.name.indexOf(array[i]) >= 0).size
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

                if(orderProducts[i].name.indexOf('Постельн') >= 0) {

                    _temp.push(`КПБ ${orderProducts[i].name}`)

                }

                if(orderProducts[i].name.indexOf('Постельн') < 0) {

                    _temp.push(orderProducts[i].name)

                }

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

            if(orderProducts[i].name.indexOf('Постельн') >= 0) {

                if(nat_cat.indexOf(`КПБ ${orderProducts[i].name}`) >= 0) {
                    temp.push(orderProducts.find(o => o.vendor === orderProducts[i].vendor).quantity)
                }

                if(temp.length%10 === 0) {
                    quantityList.splice(-1, 0, ...quantityList.splice(-1, 1, temp))
                    temp = []
                }

            }

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

    function createOrder() {

        let List = createNameList()
        let Quantity = createQuantityList()
        let content = ``

        console.log(List)

        for(let i = 0; i < List.length; i++) {
            if(List[i].length > 0) {
                content += `<?xml version="1.0" encoding="utf-8"?>
                            <order xmlns="urn:oms.order" xsi:schemaLocation="urn:oms.order schema.xsd" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
                                <lp>
                                    <productGroup>lp</productGroup>
                                    <contactPerson>333</contactPerson>
                                    <releaseMethodType>PRODUCTION</releaseMethodType>
                                    <createMethodType>SELF_MADE</createMethodType>
                                    <productionOrderId>PERSONAL_${i+1}</productionOrderId>
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

            month < 10 ? filePath = `./public/orders/lp_personal_${i}_${date_ob.getDate()}_0${month}.xml` : filePath = `./public/orders/lp_personal_${i}_${date_ob.getDate()}_${month}.xml`

            if(content !== ``) {
                fs.writeFileSync(filePath, content)
            }

            content = ``

        }

    }

    console.log(difference.length)
    console.log(new_items.length)

    if(new_items.length > 0) await createImport(new_items)

    if(difference.length <= 0) createOrder()

    const List = createNameList()
    const Quantity = createQuantityList()
    const orders = []
    for(let i = 0; i < List.length; i++) {
        for(let j = 0; j < List[i].length; j++) {
            orders.push({ name: List[i][j], quantity: Quantity[i][j], isNew: nat_cat.indexOf(List[i][j]) < 0 })
        }
    }
    res.render('marks-order', { title: 'Персональный заказ', orders, buttons })

})

module.exports = router
