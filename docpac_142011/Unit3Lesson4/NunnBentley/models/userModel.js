// User Model: every SQL statement for the Users table lives here.
// Routes call these functions instead of writing SQL themselves.
// Every user-supplied value goes through a ? placeholder (parameterized).
const { db } = require('./database');

// COLLATE NOCASE makes "Alex" and "alex" count as the same username.
function findUserByUsername(username) {
    return db.prepare('SELECT * FROM Users WHERE username = ? COLLATE NOCASE').get(username);
}

// Emails are saved in lowercase, but NOCASE also catches older mixed-case rows.
function findUserByEmail(email) {
    return db.prepare('SELECT * FROM Users WHERE email = ? COLLATE NOCASE').get(email);
}

function findUserById(id) {
    return db.prepare('SELECT * FROM Users WHERE id = ?').get(id);
}

// Returns the new user's id so the Controller can store it in the session.
// NOTE: the password is stored as plain text for THIS LESSON ONLY.
// Lesson 5 replaces this with hashing.
function createUser(username, email, password, displayName) {
    const result = db
        .prepare('INSERT INTO Users (username, email, password, display_name) VALUES (?, ?, ?, ?)')
        .run(username, email, password, displayName || null);
    return Number(result.lastInsertRowid);
}

module.exports = {
    findUserByUsername,
    findUserByEmail,
    findUserById,
    createUser
};
