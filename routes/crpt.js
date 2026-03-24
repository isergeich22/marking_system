const router = require('express').Router()
const axios = require('axios')
const { cryptoProPage } = require('../components/htmlComponents')

let crptToken = ''

router.get('/crpt_test', async (req, res) => {
    res.send(cryptoProPage)
})

router.get('/crpt_api_test', async (req, res) => {

    console.log(crptToken)

    const response = await axios.get('https://markirovka.crpt.ru/api/v3/true-api/nk/short-product?gtin=04610594532618',
        {
            headers: {
                Accept: 'application/json',
                Authorization: `Bearer ${crptToken}`
            }
        }
    )

    res.json(response.data)

})

router.get('/api/v3/true-api/auth/key', async (req, res) => {
    try {
        const response = await axios.get('https://markirovka.crpt.ru/api/v3/true-api/auth/key', {
            headers: { "Accept": "application/json" }
        })
        res.json(response.data)
    } catch (e) {
        res.status(e.response?.status || 500).json({ error: e.message })
    }
})

router.post('/api/v3/true-api/auth/simpleSignIn', async (req, res) => {
    try {
        const response = await axios.post('https://markirovka.crpt.ru/api/v3/true-api/auth/simpleSignIn', req.body, {
            headers: { "Content-Type": "application/json", "Accept": "application/json" }
        })
        crptToken = response.data.token
        console.log(crptToken)
        res.json(response.data)
    } catch (e) {
        res.status(e.response?.status || 500).json({ error: e.message })
    }
})

module.exports = router
