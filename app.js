const express = require('express')
const dotenv = require('dotenv')
const errorHandler = require('./middleware/errorHandler')

dotenv.config({ path: __dirname + '/.env' })

const app = express()
app.use(express.json())
app.use(express.static(__dirname + '/public'))

app.use('/', require('./routes/home'))
app.use('/', require('./routes/ozon'))
app.use('/', require('./routes/wildberries'))
app.use('/', require('./routes/yandex'))
app.use('/', require('./routes/stocks'))
app.use('/', require('./routes/marking'))
app.use('/', require('./routes/features'))
app.use('/', require('./routes/analytics'))
app.use('/', require('./routes/crpt'))

app.use(errorHandler)

app.listen(3030, () => {
    console.log('Server running on http://localhost:3030')
})
