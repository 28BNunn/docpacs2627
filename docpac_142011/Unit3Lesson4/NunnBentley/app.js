require("dotenv").config()

const express = require("express")
const app = express()
const path = require('path')

app.set('view engine', 'ejs')
app.set('views', path.join(__dirname, 'views'))

const take = require('./utils/requestlogger')
const {page} = require('./routes/pageRoutes')
const {form} = require('./routes/formRoutes')
const {parameter} = require('./routes/parameterRoutes')
const {sessionmiddleware} = require('./utils/sessionConfig')
const {sessionRoutes} = require('./routes/SessionRoutes')
const {auth} = require('./routes/authRoutes')
const {authViewData} = require('./utils/authViewData')

const port = process.env.PORT;
const app_name = process.env.APP_NAME

app.use(take)

app.use(express.static('public'))
app.use(express.urlencoded({
    extended: true,
}))

app.use(sessionmiddleware)
// Must come after the session and before any route that renders a View
app.use(authViewData)

app.use(page)
app.use(form)
app.use(parameter)
app.use(sessionRoutes)
app.use(auth)

app.all('*path', (req, res) => {
    res.status(404).render('error', {
        title: '404: Page not found',
        code: res.statusCode,
        message: 'The page you are looking for does not exist. put in correct url'
    })
})

app.listen(port, 'localhost', () => {
    console.log(`Server running at http://localhost:${port}`)
})
