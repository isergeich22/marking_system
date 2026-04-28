const router = require('express').Router()
const exl = require('exceljs')
const fs = require('fs')
const axios = require('axios')
const QuickChart = require('quickchart-js')
const buttons = require('../config').buttons
const { findMatchesByPostingNumber, splitArrayIntoChunks } = require('../helpers/utils')

const dbsId = process.env.YANDEX_DBS_ID
const fbsId = process.env.YANDEX_FBS_ID

// TODO: CITIES не был определён в исходном app.js — /cdek_test требует заполнения этого массива
const CITIES = []

router.get('/get_products_analytic/:year', async function (req, res) {

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
    const rows = Object.entries(analyticObject).map(([label, value]) => ({ label, value }))

    res.render('analytics-chart', {
        title: `Аналитика спроса за ${year} год`,
        chartUrl,
        col1: 'Продукт',
        col2: 'Количество',
        rows,
        buttons
    })

    // console.log(ordersList.length)
    // res.json(analyticObject)

})

router.get('/get_products_analytic/:year/:product', async function (req, res) {

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

            if(order.products.find(o => o.name.toLowerCase().indexOf('тенсел') >= 0)) {
                analyticObject["Тенсель"] = analyticObject["Тенсель"] + 1
            }

            if(order.products.find(o => o.name.toLowerCase().indexOf('сатин') >= 0 && o.name.toLowerCase().indexOf('страйп') < 0 && o.name.toLowerCase().indexOf('жаккард') < 0 && o.name.toLowerCase().indexOf('твил') < 0 && o.name.toLowerCase().indexOf('поли') < 0)) {
                analyticObject["Сатин"] = analyticObject["Сатин"] + 1
            }

            if(order.products.find(o => o.name.toLowerCase().indexOf('сатин') >= 0 && o.name.toLowerCase().indexOf('страйп') >= 0)) {
                analyticObject["Страйп-сатин"] = analyticObject["Страйп-сатин"] + 1
            }

            if(order.products.find(o => o.name.toLowerCase().indexOf('сатин') >= 0 && o.name.toLowerCase().indexOf('твил') >= 0)) {
                analyticObject["Твил-сатин"] = analyticObject["Твил-сатин"] + 1
            }

            if(order.products.find(o => o.name.toLowerCase().indexOf('сатин') >= 0 && o.name.toLowerCase().indexOf('поли') >= 0)) {
                analyticObject["Полисатин"] = analyticObject["Полисатин"] + 1
            }

            if(order.products.find(o => o.name.toLowerCase().indexOf('сатин') >= 0 && o.name.toLowerCase().indexOf('жаккард') >= 0)) {
                analyticObject["Сатин-жаккард"] = analyticObject["Сатин-жаккард"] + 1
            }

            if(order.products.find(o => o.name.toLowerCase().indexOf('бяз') >= 0)) {
                analyticObject["Бязь"] = analyticObject["Бязь"] + 1
            }

            if(order.products.find(o => o.name.toLowerCase().indexOf('варен') >= 0 || o.name.toLowerCase().indexOf('варён') >= 0 || o.name.toLowerCase().indexOf('хлоп') >= 0)) {
                analyticObject["Вареный хлопок"] = analyticObject["Вареный хлопок"] + 1
            }

            if(order.products.find(o => o.name.toLowerCase().indexOf('микрофибр') >= 0)) {
                analyticObject["Микрофибра"] = analyticObject["Микрофибра"] + 1
            }

            if(order.products.find(o => o.name.toLowerCase().indexOf('мулетон') >= 0)) {
                analyticObject["Мулетон"] = analyticObject["Мулетон"] + 1
            }

            if(order.products.find(o => o.name.toLowerCase().indexOf('поплин') >= 0)) {
                analyticObject["Поплин"] = analyticObject["Поплин"] + 1
            }

            if(order.products.find(o => o.name.toLowerCase().indexOf('перкал') >= 0)) {
                analyticObject["Перкаль"] = analyticObject["Перкаль"] + 1
            }

            if(order.products.find(o => o.name.toLowerCase().indexOf('ранфор') >= 0)) {
                analyticObject["Ранфорс"] = analyticObject["Ранфорс"] + 1
            }

            if(order.products.find(o => o.name.toLowerCase().indexOf('микросатин') >= 0)) {
                analyticObject["Микросатин"] = analyticObject["Микросатин"] + 1
            }

            if(order.products.find(o => o.name.toLowerCase().indexOf('креп-ж') >= 0 || order.products.find(o => o.name.toLowerCase().indexOf('креп') >= 0))) {
                analyticObject["Креп-жатка"] = analyticObject["Креп-жатка"] + 1
            }

            if(order.products.find(o => o.name.toLowerCase().indexOf('жатка') >= 0 && order.products.find(o => o.name.toLowerCase().indexOf('креп') < 0))) {
                analyticObject["Жатка"] = analyticObject["Жатка"] + 1
            }

        }

        let html = `<div class="fixed-grid has-1-cols"><div class="grid">`

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

        let data = []

        let responseCount = 0
        let chuncksArray = []

        try {

            if(uniqueOffers.length > 1000) {

                responseCount = Math.ceil(uniqueOffers.length / 1000)
                chuncksArray = splitArrayIntoChunks(uniqueOffers, 1000)

            }

            for(let i = 0; i < responseCount; i++) {

                const response = await axios.post('https://api-seller.ozon.ru/v4/product/info/attributes', {

                    "filter": {
                        "offer_id": chuncksArray[i],
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

                data = data.concat(response.data.result)

            }

        } catch(err) {
            console.error(err)
        }

        for(let i of data) {

            // console.log(i.attributes.find(o => o.id === 6771).values[0].value)

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

        let dataStandart = []

        if (uniqueOffersStandart.length > 0) {
            const chunksStandart = splitArrayIntoChunks(uniqueOffersStandart, 1000)

            for (let chunk of chunksStandart) {
                const responseStandart = await axios.post('https://api-seller.ozon.ru/v4/product/info/attributes', {
                    "filter": {
                        "offer_id": chunk,
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

                dataStandart = dataStandart.concat(responseStandart.data.result)
            }
        }

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

        let dataColor = []

        if (uniqueOffersColor.length > 0) {
            const chunksColor = splitArrayIntoChunks(uniqueOffersColor, 1000)

            for (let chunk of chunksColor) {
                const responseColor = await axios.post('https://api-seller.ozon.ru/v4/product/info/attributes', {
                    "filter": {
                        "offer_id": chunk,
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

                dataColor = dataColor.concat(responseColor.data.result)
            }
        }

        for(let i of dataColor) {

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

        html += `</div></div>`

        res.render('analytics-raw', { title: `Аналитика спроса за ${year} год`, content: html, buttons })

    }

    if(req.params.product.toLowerCase().indexOf('пододе') >= 0) {

        for(let order of ordersList) {

            if(order.products.find(o => o.name.toLowerCase().indexOf('тенсел') >= 0)) {
                analyticObject["Тенсель"] = analyticObject["Тенсель"] + 1
            }

            if(order.products.find(o => o.name.toLowerCase().indexOf('сатин') >= 0 && o.name.toLowerCase().indexOf('страйп') < 0 && o.name.toLowerCase().indexOf('жаккард') < 0 && o.name.toLowerCase().indexOf('твил') < 0 && o.name.toLowerCase().indexOf('поли') < 0)) {
                analyticObject["Сатин"] = analyticObject["Сатин"] + 1
            }

            if(order.products.find(o => o.name.toLowerCase().indexOf('сатин') >= 0 && o.name.toLowerCase().indexOf('страйп') >= 0)) {
                analyticObject["Страйп-сатин"] = analyticObject["Страйп-сатин"] + 1
            }

            if(order.products.find(o => o.name.toLowerCase().indexOf('сатин') >= 0 && o.name.toLowerCase().indexOf('твил') >= 0)) {
                analyticObject["Твил-сатин"] = analyticObject["Твил-сатин"] + 1
            }

            if(order.products.find(o => o.name.toLowerCase().indexOf('сатин') >= 0 && o.name.toLowerCase().indexOf('поли') >= 0)) {
                analyticObject["Полисатин"] = analyticObject["Полисатин"] + 1
            }

            if(order.products.find(o => o.name.toLowerCase().indexOf('сатин') >= 0 && o.name.toLowerCase().indexOf('жаккард') >= 0)) {
                analyticObject["Сатин-жаккард"] = analyticObject["Сатин-жаккард"] + 1
            }

            if(order.products.find(o => o.name.toLowerCase().indexOf('бяз') >= 0)) {
                analyticObject["Бязь"] = analyticObject["Бязь"] + 1
            }

            if(order.products.find(o => o.name.toLowerCase().indexOf('варен') >= 0 || o.name.toLowerCase().indexOf('варён') >= 0 || o.name.toLowerCase().indexOf('хлоп') >= 0)) {
                analyticObject["Вареный хлопок"] = analyticObject["Вареный хлопок"] + 1
            }

            if(order.products.find(o => o.name.toLowerCase().indexOf('микрофибр') >= 0)) {
                analyticObject["Микрофибра"] = analyticObject["Микрофибра"] + 1
            }

            if(order.products.find(o => o.name.toLowerCase().indexOf('мулетон') >= 0)) {
                analyticObject["Мулетон"] = analyticObject["Мулетон"] + 1
            }

            if(order.products.find(o => o.name.toLowerCase().indexOf('поплин') >= 0)) {
                analyticObject["Поплин"] = analyticObject["Поплин"] + 1
            }

            if(order.products.find(o => o.name.toLowerCase().indexOf('перкал') >= 0)) {
                analyticObject["Перкаль"] = analyticObject["Перкаль"] + 1
            }

            if(order.products.find(o => o.name.toLowerCase().indexOf('ранфор') >= 0)) {
                analyticObject["Ранфорс"] = analyticObject["Ранфорс"] + 1
            }

            if(order.products.find(o => o.name.toLowerCase().indexOf('микросатин') >= 0)) {
                analyticObject["Микросатин"] = analyticObject["Микросатин"] + 1
            }

            if(order.products.find(o => o.name.toLowerCase().indexOf('креп-ж') >= 0 || order.products.find(o => o.name.toLowerCase().indexOf('креп') >= 0))) {
                analyticObject["Креп-жатка"] = analyticObject["Креп-жатка"] + 1
            }

            if(order.products.find(o => o.name.toLowerCase().indexOf('жатка') >= 0 && order.products.find(o => o.name.toLowerCase().indexOf('креп') < 0))) {
                analyticObject["Жатка"] = analyticObject["Жатка"] + 1
            }

        }

        let html = `<div class="fixed-grid has-1-cols"><div class="grid">`

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

        html += `</div></div>`

        res.render('analytics-raw', { title: `Аналитика спроса за ${year} год`, content: html, buttons })

    }

})

router.get('/get_year_dynamic/:year', async function (req, res) {

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

    let html = ''

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

    // res.json(monthQuantityOrders)

    res.render('analytics-raw', { title: `Динамика заказов за ${year} год`, content: html, buttons })

})

router.get('/get_income_analytic/:month/:product', async function (req, res) {

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

    const rows = []

    for(let item of orderIncomeInfo) {

        let element = item.operations.find(o => o.operation_type_name === 'Доставка покупателю')

        if(element.accruals_for_sale > 0) {
            rows.push({
                postingNumber: element.posting.posting_number,
                name: element.items[0].name,
                amount: Math.round(element.amount),
                accruals: Math.round(element.accruals_for_sale),
                percent: Math.round((Math.round(element.amount) / Math.round(element.accruals_for_sale)) * 100)
            })
        }

    }

    res.render('analytics-income', { title: 'Аналитика дохода с заказов', rows, buttons })

    // res.json(orderIncomeInfo)

})

router.get('/api_test', async function (req, res) {

    // let pageSize = 0
    // let token = ''
    // let result = []
    // let itterator = 0

    // const ya_response = await axios.get(`https://api.partner.market.yandex.ru/v2/campaigns/${dbsId}/outlets?limit=50`, {
    //     headers: {
    //         "Authorization": `Bearer ${process.env.YANDEX_API_KEY}`
    //     }
    // })

    // result.push(ya_response.data.outlets)
    // pageSize = ya_response.data.pager.pageSize
    // token = ya_response.data.paging.nextPageToken

    // while (itterator < 68) {

    //     console.log(token)
    //     console.log(pageSize)
    //     console.log(itterator)

    //     const ya_response = await axios.get(`https://api.partner.market.yandex.ru/v2/campaigns/${dbsId}/outlets?limit=50&page_token=${token}`, {
    //         headers: {
    //             "Authorization": `Bearer ${process.env.YANDEX_API_KEY}`
    //         }
    //     })

    //     result.push(ya_response.data.outlets)
    //     pageSize = ya_response.data.pager.pageSize
    //     token = ya_response.data.paging.nextPageToken
    //     itterator++

    // }

    // res.json({ yandex: result })

    let total = 0
    let last_id = ''
    let items = []
    let counter = 0

    const ozonResponse = await axios.post('https://api-seller.ozon.ru/v3/product/list', {
        "filter": {
            "visibility": "MANUAL_ARCHIVED"
        },
        "last_id": "",
        "limit": 200
    }, {
        headers: {
            'Host':'api-seller.ozon.ru',
            'Client-Id':`${process.env.OZON_CLIENT_ID}`,
            'Api-Key':`${process.env.OZON_API_KEY}`,
            'Content-Type':'application/json'
        }
    })

    total = ozonResponse.data.result.total
    last_id = ozonResponse.data.result.last_id

    items.push(ozonResponse.data.result.items)
    counter += ozonResponse.data.result.items.length

    // console.log(total)

    while(counter < total) {

        const ozonResponse = await axios.post('https://api-seller.ozon.ru/v3/product/list', {
            "filter": {
                "visibility": "MANUAL_ARCHIVED"
            },
            "last_id": `${last_id}`,
            "limit": 200
        }, {
            headers: {
                'Host':'api-seller.ozon.ru',
                'Client-Id':`${process.env.OZON_CLIENT_ID}`,
                'Api-Key':`${process.env.OZON_API_KEY}`,
                'Content-Type':'application/json'
            }
        })

        last_id = ozonResponse.data.result.last_id
        items.push(ozonResponse.data.result.items)
        counter += ozonResponse.data.result.items.length

    }

    const skuArray = ['00-00185470']

    // for(let i = 0; i < 5; i++) {

    //     skuArray.push(ozonResponse.data.result.items[i].offer_id)

    // }

    // res.json(items)

    const yaDbsStockResponse = await axios.put(`https://api.partner.market.yandex.ru/v2/campaigns/${dbsId}/offers/stocks`, {
        "skus": [
            {
            "sku": `${skuArray[0]}`,
            "items": [
                {
                    "count": 0,
                    "updatedAt": "2026-02-24T00:00:00Z"
                }
            ]
            }
        ]
    }, {
        headers: {
            'Authorization': `Bearer ${process.env.YANDEX_API_KEY}`
        }
    })

    console.log(yaDbsStockResponse.data)

    const yaFbsStockResponse = await axios.put(`https://api.partner.market.yandex.ru/v2/campaigns/${fbsId}/offers/stocks`, {
        "skus": [
            {
            "sku": `${skuArray[0]}`,
            "items": [
                {
                    "count": 0,
                    "updatedAt": "2026-02-24T00:00:00Z"
                }
            ]
            }
        ]
    }, {
        headers: {
            'Authorization': `Bearer ${process.env.YANDEX_API_KEY}`
        }
    })

    console.log(yaFbsStockResponse.data)

    const yaArchiveResponse = await axios.post(`https://api.partner.market.yandex.ru/v2/businesses/${process.env.YANDEX_BUSINESS_ID}/offer-mappings/archive`, {
        "offerIds": [
            skuArray[0]
        ]
    }, {
        headers: {
            'Authorization': `Bearer ${process.env.YANDEX_API_KEY}`
        }
    })

    res.json(yaArchiveResponse.data)

})

router.get('/cdek_test/:from/:to', async function (req, res) {

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

router.get('/revenue/:year', async (req, res) => {

    let i = 0
    let hasNext = true

    const REPORT_PATH = `REPORT-${req.params.year}.xlsx`

    let orders = []

    while(hasNext) {

        const response = await axios.post('https://api-seller.ozon.ru/v3/posting/fbs/list', {

            "dir": "ASC",
            "filter": {
                "since": `${req.params.year}-01-01T00:00:00.000Z`,
                "status": "delivered",
                "to": `${req.params.year}-12-31T23:59:59.999Z`,
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

            console.log(rec)

            let _temp = 0

            rec.accruals.forEach(a => {

                _temp += a.accrual_value

            })

            o.revenue = Math.round(_temp)

        }

    })

    // res.json(orders)

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

module.exports = router
