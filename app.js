var xlsx = require ("xlsx");
var fs = require("fs");
var path = require("path");

var sourceDir = './LIQUIDACIONES';

function readFileToJson(filename){
    var wb = xlsx.readFile(filename,{cellDates: true});
    var firstTabName = wb.SheetNames[0];
    var ws = wb.Sheets[firstTabName];
    var data = xlsx.utils.sheet_to_json(ws);
    // console.log("1") q q
    return data;
}

var tagedDir = path.join(__dirname, sourceDir);
var files = fs.readdirSync(tagedDir);

// console.log(files);

var combinedData = [];

files.forEach(function(file){
var fileExtenstion = path.parse(file).ext;
console.log(fileExtenstion);
if((fileExtenstion === ".xlsx" || fileExtenstion === ".xls" || fileExtenstion === ".XLS") && file[0] !== "~"){
    var fullFilePath = path.join(__dirname, sourceDir, file);
    var data = readFileToJson(fullFilePath);
    combinedData = combinedData.concat(data);
    console.log(data);
}
});

var newWB = xlsx.utils.book_new();
var newWS = xlsx.utils.json_to_sheet(combinedData);
xlsx.utils.book_append_sheet(newWB, newWS, "Combined Data");

xlsx.writeFile(newWB, "LIQUIDACIONES.xlsx");
console
console.log("done!")