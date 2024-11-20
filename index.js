const ExcelJS = require('exceljs/dist/es5');
const path = require('path');
const fs = require("fs")
const ps = require("process");
let workbook = new ExcelJS.Workbook();
if(ps.argv.length < 3){
    console.error("请指定输入文件地址")
    return
}
const outFile = ps.argv.length == 4 ? ps.argv[3].startsWith("/") ? ps.argv[3] : path.join(__dirname,ps.argv[3]) : path.join(__dirname,"out.xlsx");
const inFile = ps.argv[2].startsWith("/") ? ps.argv[2] : path.join(__dirname,ps.argv[2]);
const jsonFile = fs.readFileSync(inFile).toString();
const jsonFileData = JSON.parse(jsonFile);

function doSheet(jsonData,workbook,strSheetName){
    const sheet = workbook.addWorksheet(strSheetName);
    let keys = Object.keys(jsonData[0]);
    sheet.addRow(keys);
    jsonData.forEach(k => {
        sheet.addRow(Object.values(k))
    });
    
}

if(Array.isArray(jsonFileData)){//检查是否为数组，是数组代表只有一个表
    doSheet(jsonFileData,workbook,"Sheet")
}else{
    const keys = Object.keys(jsonFileData)//不是数组代表有多张表
    keys.forEach(k=>{
        doSheet(jsonFileData[k],workbook,k);
    });
}
workbook.xlsx.writeFile(outFile);


