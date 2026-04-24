const router = require('express').Router()
const exl = require('exceljs')
const axios = require('axios')
const fs = require('fs')
const buttons = require('../config').buttons

router.get('/wildberries', async function(req, res){

    const new_items = []
    const current_items = []
    const moderation_items = []
    const wb_orders = []
    const nat_cat = []
    let names = []

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

        try {

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
                    'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase() === 'ТЕНСЕЛЬ' ? 'ТЕНСЕЛ' : response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                    'productType': 'ПОДОДЕЯЛЬНИК'
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
                        'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase() === 'ТЕНСЕЛЬ' ? 'ТЕНСЕЛ' : response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
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
                        'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase() === 'ТЕНСЕЛЬ' ? 'ТЕНСЕЛ' : response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
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
                        'cloth': response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase() === 'ТЕНСЕЛЬ' ? 'ТЕНСЕЛ' : response.data.result[0].attributes.find(o => o.id === 6769).values[0].value.toUpperCase(),
                        'productType': 'НАВОЛОЧКА'
                    })

                } else {

                    names.push({
                        'vendor': wb_orders[i].vendor,
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
                                'vendor': wb_orders[i].vendor,
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
                                'vendor': wb_orders[i].vendor,
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
                                'vendor': wb_orders[i].vendor,
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
                            'vendor': wb_orders[i].vendor,
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
                                'vendor': wb_orders[i].vendor,
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
                                'vendor': wb_orders[i].vendor,
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
                                'vendor': wb_orders[i].vendor,
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
                            'vendor': wb_orders[i].vendor,
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

            names = names.filter(o => o.name.indexOf('Одеяло') < 0 && o.name.indexOf('Подушка') < 0 && o.name.indexOf('Матрас') < 0)

        } catch(err) {

            console.log(`error: ${wb_orders[i].vendor}`)

        }

    }

    names.forEach(el => {

            if(nat_cat.indexOf(el.name) < 0) {
                new_items.push(el.name)
            }

            if(nat_cat.indexOf(el.name) >= 0) {
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
            ws.getCell(`P${cellNumber}`).value = 'Черновик'

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

    if(new_items.length > 0) {
        console.log('worked')
        await createImport(new_items)
    }

    const items = names.map(el => ({ name: el.name, isNew: new_items.indexOf(el.name) >= 0, isModeration: moderation_items.indexOf(el.name) >= 0 }))

    res.render('wildberries-import', { title: 'Импорт - Wildberries', items, marksOrderUrl: '/wildberries_marks_order', buttons })

})

router.get('/wildberries_marks_order', async function(req, res) {

    const wb_orders = []
    const nat_cat = []
    const gtins = []
    let names = []

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
                                    <releaseMethodType>PRODUCTION</releaseMethodType>
                                    <createMethodType>SELF_MADE</createMethodType>
                                    <productionOrderId>WB_${i+1}</productionOrderId>
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

            month < 10 ? filePath = `./public/orders/lp_wb_${i}_${date_ob.getDate()}_0${month}.xml` : filePath = `./public/orders/lp_wb_${i}_${date_ob.getDate()}_${month}.xml`

            if(content !== ``) {
                fs.writeFileSync(filePath, content)
            }

            content = ``

        }

    }

    createOrder()

    const orders = []
    for(let i = 0; i < List.length; i++) {
        for(let j = 0; j < List[i].length; j++) {
            orders.push({ name: List[i][j], quantity: Quantity[i][j], isNew: nat_cat.indexOf(List[i][j]) < 0 })
        }
    }

    res.render('marks-order', { title: 'Заказ маркировки - Wildberries', orders, buttons })

})

router.get('/wildberries/set_marks', async function (req, res){

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

    console.log(_temp)

    await wb.xlsx.readFile(marksFile)

    const ws_5 = wb.getWorksheet(1)

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

        console.log(wbOrder[i].orderProduct)

        const gtin = _temp.find(o => o.name === wbOrder[i].orderProduct).gtin

        if(gtin === undefined) {            
            
            wbOrder[i].mark = ''

        } else {

            console.log(gtin)

            const mark = marks.find(o => o.gtin === String(gtin) && o.status === 'not_used').mark

            // console.log(mark)

            if(mark) {

                wbOrder[i].mark = mark
                marks.find(o => o.gtin === String(gtin) && o.status === 'not_used').status = 'used'

            } else {
                
                wbOrder[i].mark = ''

            }

        }

    }

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

    const orders = wbOrder.map(o => ({ orderNumber: o.orderNumber, orderProduct: o.orderProduct, orderCode: o.orderCode, mark: o.mark.replace(/\u003C/g, '&lt').replace(/\u003E/g, '&gt').replace(/\"/g, '&quot') }))

    res.render('wildberries-set-marks', { title: 'Подстановка маркировки - Wildberries', orders, buttons })

    // res.json({wbOrder, marks, _temp})

})

module.exports = router
