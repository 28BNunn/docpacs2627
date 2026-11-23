// Puts SAFE login info into res.locals so every EJS View (and the shared
// header) can see it. Never puts the password or the full session here.
const { findUserById } = require('../models/userModel')

function authViewData(req, res, next) {
    res.locals.isLoggedIn = false
    res.locals.currentUsername = null

    if (req.session && req.session.userId) {
        const user = findUserById(req.session.userId)
        if (user) {
            res.locals.isLoggedIn = true
            res.locals.currentUsername = user.username
        }
    }
    next()
}

module.exports = { authViewData }
