function Calculate(){

num1 = Number(document.getElementById("num1").V)
num2 = Number(document.getElementById(num2))
operation =(document.getElementById(operation))

switch(String(operation)){


    case "*":
        output = num1 * num2
        break;
    case "/":
        if(num2 != 0){ output = num1 / num2 } else output = "Error!"
        break;
    case "+":
        output = num1 + num2
        break;
    case "-":
        output = num1 - num2        
}

//document.getElementById("result").textContent = String(output)

console.log("num1=" + num1 + " num2=" + num2 + " operation=" + operation + " output="/* + output*/)

}