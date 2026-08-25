function calculate() {
    const resultElement = document.getElementById("result");
    const number1Value  = document.getElementById("number1").value;
    const number2Value  = document.getElementById("number2").value;
    const operatorValue = document.getElementById("operation").value;

    const number1 = Number(number1Value);
    const number2 = Number(number2Value);

    let result = "ERROR: Unknown";
    switch (operatorValue) {
        case "addition": 
            result = String(number1 + number2);
            break;
        case "subtraction":
            result = String(number1 - number2);
        case "multiplication":

        case "division":
            if (number2 === 0) {
                result = "ERROR: Divide by zero";
            }
            else {
                result = String(number1 / number2);
            }
            break;
    }
}