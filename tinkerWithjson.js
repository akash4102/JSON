let fs=require("fs");
const { json } = require("stream/consumers");
const { fileURLToPath } = require("url");
let xlsx=require("xlsx");
// let buffer=fs.readFileSync("./example.json");
// console.log(buffer);
// console.log("........................")
// let data=JSON.parse(buffer);
let data=require("./example.json")
// console.log(data);
// console.log("........................")
// data.push({
//     "name":"thor",
//     "last Name":"rogers",
//     "isAvenger":true,
//     "frinds":["Bruce","peter","natasha"],
//     "age":45,
//     "address":{
//         "city":"new york",
//         "state":"manhaton"
//     }
// });
// let stringdata=JSON.stringify(data);
// fs.writeFileSync("./example.json",stringdata);
// console.log(data);


function excelWriter(filePath,json,sheetName){
    //write
    // wb->filepath,ws,->name,json data
    // new worksheet
    let newWb=xlsx.utils.book_new();
    //json data -> excel format convert
    let newWS=xlsx.utils.json_to_sheet(json);
    // ->new workbook, worksheet,sheet,name    "sheet-1"
    xlsx.utils.book_append_sheet(newWb,newWS,sheetName);
    //file path   "abc.xlsx"
    xlsx.writeFile(newWb,filePath);
}

excelReader("webdevelopmentjsonpqr.xlsx","sheet-5");
excelWriter("webdevelopmentjsonpqr.xlsx",data,"sheet-5");
function excelReader(filePath,sheetName){
    //read
    if(fs.existsSync(filePath)){
        let wb=xlsx.readFile(filePath);
        let exceldata=wb.Sheets[sheetName];
        let ans=xlsx.utils.sheet_to_json(exceldata);
        console.log(ans);
    }
    else{
        console.log("file doesn't exist")
    }
}
