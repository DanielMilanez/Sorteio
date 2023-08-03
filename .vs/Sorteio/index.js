var XLSX = require("xlsx")
var workbook = XLSX.readFile("C:\\Teste.xlsx")

let worksheet = workbook.Sheets[workbook.SheetNames[0]]

for (let index = 1; index < 7; index++){
    const numero = worksheet['A${index}'].v
    const name = worksheet['B${index}'].v

console.log({
    id: id, name: name
})
}