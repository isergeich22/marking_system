const router = require('express').Router()
const exl = require('exceljs')
const fs = require('fs')
const axios = require('axios')
const buttons = require('../config').buttons

router.get('/stocks', async function(req, res){

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
                            'size': `Пododеяльник: ${response.data.result[0].attributes.find(o => o.id === 6773).values[0].value}; Простыня: ${response.data.result[0].attributes.find(o => o.id === 6771).values[0].value}; Наволочка: ${response.data.result[0].attributes.find(o => o.id === 6772).values[0].value}`,
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

    const newItems = full_difference.map(i => ({ name: i.name, vendor: i.vendor }))
    const currentItems = full_matches.map(i => ({ name: i.name, vendor: i.vendor }))

    console.log(new_items.length)

    res.render('stocks', { title: 'Остатки', newItems, currentItems, buttons })

})

module.exports = router
