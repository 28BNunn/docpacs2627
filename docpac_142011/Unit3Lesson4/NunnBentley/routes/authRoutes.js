// Authentication Controller: register, login, account, logout.
// No SQL here — all database work goes through the User Model.
const express = require('express')
const auth = express.Router()
const {
    findUserByUsername,
    findUserByEmail,
    findUserById,
    createUser
} = require('../models/userModel')

const EMAIL_PATTERN = /^[^\s@]+@[^\s@]+\.[^\s@]+$/

// Turns any form value into a trimmed string ('' if it's missing).
function clean(value) {
    return typeof value === 'string' ? value.trim() : ''
}

// Renders the register page. `values` refills the form (never the passwords).
function renderRegister(res, status, error, values = {}) {
    res.status(status).render('register', {
        title: 'Register',
        heading: 'Create an Account',
        error: error,
        values: values
    })
}

function renderLogin(res, status, error, identifier = '') {
    res.status(status).render('login', {
        title: 'Log In',
        heading: 'Log In',
        error: error,
        identifier: identifier
    })
}

// Starts a fresh session for the user, then redirects once it is saved.
// regenerate() gives the user a new session ID when they log in,
// which stops an old (possibly stolen) session ID from being reused.
function logUserIn(req, res, userId) {
    req.session.regenerate((err) => {
        if (err) return res.status(500).render('error', { title: 'Error', message: 'Could not start session.', code: 500 })
        req.session.userId = userId
        req.session.save((err) => {
            if (err) return res.status(500).render('error', { title: 'Error', message: 'Could not save session.', code: 500 })
            res.redirect('/account')
        })
    })
}

// ---------- Registration ----------

auth.get('/register', (req, res) => {
    if (req.session.userId) return res.redirect('/account')
    renderRegister(res, 200, null)
})

auth.post('/register', (req, res) => {
    const username = clean(req.body.username)
    const email = clean(req.body.email).toLowerCase()
    const displayName = clean(req.body.displayName)
    // Passwords are not trimmed — spaces can be part of a password.
    const password = typeof req.body.password === 'string' ? req.body.password : ''
    const repeatPassword = typeof req.body.repeatPassword === 'string' ? req.body.repeatPassword : ''
    const values = { username, email, displayName }

    // 1. Validation happens BEFORE touching the database.
    if (!username || !email || !password || !repeatPassword) {
        return renderRegister(res, 400, 'All fields except display name are required.', values)
    }
    if (username.length < 3 || username.length > 30) {
        return renderRegister(res, 400, 'Username must be 3–30 characters.', values)
    }
    if (!/^[A-Za-z0-9_]+$/.test(username)) {
        return renderRegister(res, 400, 'Username can only use letters, numbers, and underscores.', values)
    }
    if (email.length > 254 || !EMAIL_PATTERN.test(email)) {
        return renderRegister(res, 400, 'Please enter a valid email address.', values)
    }
    if (displayName.length > 50) {
        return renderRegister(res, 400, 'Display name must be 50 characters or fewer.', values)
    }
    if (password.length < 6) {
        return renderRegister(res, 400, 'Password must be at least 6 characters.', values)
    }
    if (password !== repeatPassword) {
        return renderRegister(res, 400, 'Passwords do not match.', values)
    }

    // 2. Duplicate checks (Read). Only say WHICH value is taken.
    if (findUserByUsername(username)) {
        return renderRegister(res, 409, 'That username is already taken.', values)
    }
    if (findUserByEmail(email)) {
        return renderRegister(res, 409, 'That email address is already registered.', values)
    }

    // 3. Create the user (Create). The UNIQUE constraints are the final safety net.
    let newUserId
    try {
        newUserId = createUser(username, email, password, displayName)
    } catch (err) {
        if (err.code === 'SQLITE_CONSTRAINT_UNIQUE') {
            return renderRegister(res, 409, 'That username or email is already registered.', values)
        }
        console.error(err)
        return renderRegister(res, 500, 'Something went wrong. Please try again.', values)
    }

    // 4. Log them in by storing ONLY the user ID in the session.
    console.log(`Registered new user id=${newUserId} username=${username}`)
    logUserIn(req, res, newUserId)
})

// ---------- Login ----------

auth.get('/login', (req, res) => {
    // Already logged in? Send them to their account instead of changing anything.
    if (req.session.userId) return res.redirect('/account')
    renderLogin(res, 200, null)
})

auth.post('/login', (req, res) => {
    const identifier = clean(req.body.identifier)
    const password = typeof req.body.password === 'string' ? req.body.password : ''

    if (!identifier || !password) {
        return renderLogin(res, 400, 'Enter your username or email and your password.', identifier)
    }

    // Read: an @ means it's an email, otherwise treat it as a username.
    const user = identifier.includes('@')
        ? findUserByEmail(identifier.toLowerCase())
        : findUserByUsername(identifier)

    if (!user) {
        return renderLogin(res, 401, 'No account was found with that username or email.', identifier)
    }

    // Plain comparison for this lesson only — Lesson 5 switches to hashed passwords.
    if (user.password !== password) {
        return renderLogin(res, 401, 'Incorrect password.', identifier)
    }

    console.log(`User id=${user.id} logged in`)
    logUserIn(req, res, user.id)
})

// ---------- Account ----------

auth.get('/account', (req, res) => {
    if (!req.session.userId) return res.redirect('/login')

    // Always get fresh data from the database using the stored ID.
    const user = findUserById(req.session.userId)

    // The ID points to an account that no longer exists: clear the session.
    if (!user) {
        return req.session.destroy(() => {
            res.clearCookie('connect.sid')
            res.redirect('/login')
        })
    }

    // Only pass safe fields to the View — never the password.
    res.render('account', {
        title: 'My Account',
        account: {
            username: user.username,
            email: user.email,
            displayName: user.display_name,
            createdAt: user.created_at
        }
    })
})

// ---------- Logout ----------

auth.post('/logout', (req, res) => {
    req.session.destroy((err) => {
        if (err) {
            console.error(err)
            return res.status(500).render('error', { title: 'Error', message: 'Could not log out.', code: 500 })
        }
        // Redirect only after the session is gone.
        res.clearCookie('connect.sid')
        res.redirect('/login')
    })
})

module.exports = { auth }
