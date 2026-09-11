const readline = require('readline');

const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
});

rl.question('What is your name?', (answer) => {
    rl.question('What is your age?', (answer1) => {
        console.log(`Hello, ${answer} who is ${answer1}, and born in ${2026-answer1}!`);
        console.log(`The current year is`,2026)
        rl.close();
})});

