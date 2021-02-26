const express = require('express')
const app = express()
const Excel = require('exceljs');
const fs = require('fs');
var _ = require("lodash");
const prompts = require('prompts');
function splitColRow(str, num) {
    let arr = [str.slice(0, num), str.slice(num, str.length)]
    if (num < str.length && isNaN(parseInt(arr[1]))) {
        return aa(str, ++num)
    }
    else {
        return arr
    }
}

// console.log('qqqqqq');
// console.log.prototype = (aa) => {
//   console.log("my log:-----", aa)
// }

const getNameFromNumber = function (num) {
    var numeric = num % 26;
    var letter = String.fromCharCode(65 + numeric);
    var num2 = parseInt(num / 26);
    if (num2 > 0) {
        return getNameFromNumber(num2 - 1) + letter;
    } else {
        return letter;
    }
}
let defaultStartText = `const Excel = require("exceljs");
const path = require("path");
var fs = require("fs");
var _ = require("lodash");
const EXCEL_CELL_FONT_FAMILY = "Calibri";
const EXCEL_CELL_FONT_COLOR_NUM = "red";
const EXCEL_CELL_FONT_SIZE = 9;
const getNameFromNumber = function (num) {
    var numeric = num % 26;
    var letter = String.fromCharCode(65 + numeric);
    var num2 = parseInt(num / 26);
    if (num2 > 0) {
        return getNameFromNumber(num2 - 1) + letter;
    } else {
        return letter;
    }
}
module.exports = {
  async generateExcel(data,options={}) {
    try {
    data = [
        {
          "date": "02/01/2021",
          "e-scooter / vehicle id": "S1149 - (0073)",
          "rider name": " NAJUTHA RAMEESH",
          "mobile": "+974 33288451",
          "email": "-",
          "zone name": "West Bay_Onaiza_Al Q",
          "ride number": "CS036995ES",
          "trip start date & time": "02/01/2021 00:00:10",
          "trip end date & time": "02/01/2021 00:05:45",
          "total fare": "5",
          "ride fare": "3",
          "total km": "0.72",
          "trip time": "00:05:35",
          "booking pass type": "-",
          "total cost": "5",
          "status": "Completed",
          "unlock fees": "2",
          "paused time": "-",
          "paused charge": "0",
          "reserved time": "-",
          "reserved charge": "0",
          "cancelled charge": "0"
        },
        {
          "date": "02/01/2021",
          "e-scooter / vehicle id": "S0161 - (0142)",
          "rider name": "Abdulrahman ",
          "mobile": "+974 66330441",
          "email": "a.abdalla1008@gmail.com",
          "zone name": "PEARL QATAR",
          "ride number": "CS036996ES",
          "trip start date & time": "02/01/2021 00:01:30",
          "trip end date & time": "02/01/2021 00:04:40",
          "total fare": "4",
          "ride fare": "2",
          "total km": "0.31",
          "trip time": "00:03:09",
          "booking pass type": "-",
          "total cost": "4",
          "status": "Completed",
          "unlock fees": "2",
          "paused time": "-",
          "paused charge": "0",
          "reserved time": "-",
          "reserved charge": "0",
          "cancelled charge": "0"
        },
        {
          "date": "02/01/2021",
          "e-scooter / vehicle id": "S0112 - (0114)",
          "rider name": "Ahmad",
          "mobile": "+974 33447655",
          "email": "lovely.tomas@hotmail.com",
          "zone name": "PEARL QATAR",
          "ride number": "CS036997ES",
          "trip start date & time": "02/01/2021 00:03:00",
          "trip end date & time": "02/01/2021 00:06:46",
          "total fare": "4",
          "ride fare": "2",
          "total km": "0.09",
          "trip time": "00:03:45",
          "booking pass type": "-",
          "total cost": "4",
          "status": "Completed",
          "unlock fees": "2",
          "paused time": "-",
          "paused charge": "0",
          "reserved time": "-",
          "reserved charge": "0",
          "cancelled charge": "0"
        },
        {
          "date": "02/01/2021",
          "e-scooter / vehicle id": "S0339 - (0053)",
          "rider name": "Hamad",
          "mobile": "+974 66678002",
          "email": "hamad.adammm@gmail.com",
          "zone name": "PEARL QATAR",
          "ride number": "CS036998ES",
          "trip start date & time": "02/01/2021 00:05:08",
          "trip end date & time": "02/01/2021 00:05:12",
          "total fare": "0",
          "ride fare": "0",
          "total km": "0",
          "trip time": "-",
          "booking pass type": "-",
          "total cost": "0",
          "status": "Cancelled",
          "unlock fees": "0",
          "paused time": "-",
          "paused charge": "0",
          "reserved time": "-",
          "reserved charge": "0",
          "cancelled charge": "0"
        },
        {
          "date": "02/01/2021",
          "e-scooter / vehicle id": "S0337 - (0056)",
          "rider name": "Mohammed",
          "mobile": "+974 66165892",
          "email": "mr-dixter1@hotmail.com",
          "zone name": "PEARL QATAR",
          "ride number": "CS036999ES",
          "trip start date & time": "02/01/2021 00:05:04",
          "trip end date & time": "02/01/2021 00:05:08",
          "total fare": "0",
          "ride fare": "0",
          "total km": "0",
          "trip time": "-",
          "booking pass type": "-",
          "total cost": "0",
          "status": "Cancelled",
          "unlock fees": "0",
          "paused time": "-",
          "paused charge": "0",
          "reserved time": "-",
          "reserved charge": "0",
          "cancelled charge": "0"
        }
      ]
    // [{rapaverage:500,shape:"round",carat:2,"rap%":-20,"rr/rt":5000,"rap value":5000,"amount":4000},{rapaverage:500,shape:"round",carat:2,"rap%":-20,"rr/rt":5000,"rap value":5000,"amount":4000},{rapaverage:500,shape:"round",carat:2,"rap%":-20,"rr/rt":5000,"rap value":5000,"amount":4000},{rapaverage:500,shape:"round",carat:2,"rap%":-20,"rr/rt":5000,"rap value":5000,"amount":4000},{rapaverage:500,shape:"round",carat:2,"rap%":-20,"rr/rt":5000,"rap value":5000,"amount":4000}]

      let sheetName = (options.sheetName ? options.sheetName : "Excel Sheet");
      let destPath = path.join("./");
            let excelName = Math.random() * 1000;
            let extension = ".xlsx";
            //check file name already exists
            if (fs.existsSync(destPath + excelName + extension)) {
                excelName = excelName + "-" + Math.random().toFixed(4).toString() + extension;
            }
            else {
                excelName = excelName + extension;
            }
      `

let defaultEndText = ` } catch (err) {
        throw err;
    }
},
}`
let excelData = async () => {

    //input required 
    // for sample data see 
    let headerRow = (await prompts({
        type: 'number',
        name: 'data',
        message: 'What is header row number in given excel?'
    })).data || 6
    let dataLength = (await prompts({
        type: 'number',
        name: 'data',
        message: 'What is data row length in given excel?'
    })).data || 527
    let workbook = new Excel.Workbook();
    await workbook.xlsx.readFile('test-org-5-4.xlsx')
    let worksheet = workbook.worksheets[0];
    let row = worksheet.getRow(1);

    // generate grid columns
    let gridColumn = []
    let excelHeaderRow = worksheet.getRow(headerRow);
    excelHeaderRow.eachCell((cell, cellNo) => {
        let nextRowCell = worksheet.getCell(getNameFromNumber(cellNo - 1) + (headerRow + 1))
        let dt = {
            title: cell.value,
            field: cell.value.toLowerCase(),
            style: cell.model.style
        }
        if (nextRowCell.value && nextRowCell.value.formula) {
            console.log(nextRowCell.value.formula)
            let regexEnd = new RegExp("(?<=[A-Z$])" + (headerRow + 1) + "(?![0-9])", "g");
            dt.formula = nextRowCell.value.formula.replace(regexEnd, "@@")
        }
        gridColumn.push(dt)

    })
    // ---- grid
    let worksheetData = {}
    if (worksheet.views) {
        worksheetData.views = worksheet.views
    }
    if (worksheet.autoFilter) {
        worksheetData.autoFilter = worksheet.autoFilter
    }
    if (worksheet.properties) {
        worksheetData.properties = worksheet.properties
    }
    if (worksheet.name) {
        worksheetData.name = worksheet.name
    }
    let dataSheet = []
    let mergeData = []
    if (worksheet._merges) {
        _.each(worksheet._merges, (val, key) => {
            mergeData.push({
                cell1: key,
                cell2: val.br
            })
        })
    }

    worksheet.eachRow((row, rowNo) => {
        let data = []
        if (rowNo < headerRow || rowNo > headerRow + dataLength) {
            row.eachCell((cell, cellNo) => {
                // console.log(cell.value);
                data.push({
                    row: rowNo,
                    value: cell.value,
                    style: cell.model.style,
                    cell: cellNo,
                    mergeCount: cell._mergeCount,
                    isMerged: cell.isMerged
                })
            })
            dataSheet.push({ data: data, hidden: row.hidden })
        }
    })



    defaultStartText += `
    let worksheetData=${JSON.stringify(worksheetData)}
  let extraDataByCell=${JSON.stringify(dataSheet)}
  let extraDataForMerge=${JSON.stringify(mergeData)}
  options.columns=${JSON.stringify(gridColumn)}

let sampleExcelDataLength = ${dataLength};
  let startOfRows = ${headerRow}
let sampleExcelDataLastRow = startOfRows + sampleExcelDataLength
  let headerRows=[1];
  

let workbook = new Excel.Workbook();
let sheet = workbook.addWorksheet(sheetName);
let excelColumn = [];
let dataFirstRow = startOfRows + 1
let dataLastRow = startOfRows + data.length

// set header
for (let i in options.columns) {
    let column = options.columns[i];
    let hAlignment = "center";
    if (column.cellClass) {
        let splited = column.cellClass.split("-");
        hAlignment = splited[1] ? splited[1].toLowerCase() : splited[0].toLowerCase();
    }
    //console.log("ha", hAlignment)
    let width = _.has(column, 'width') ? parseInt(column.width) : 8;
    // reset numeric format of certificate number
    let obj = {
        header: column.title,
        key: column.field,
        width: width,
        alignment: {
            vertical: "middle",
            horizontal: "center"
        },
    };
    if (column.subTitle) obj.subTitle = column.subTitle
    excelColumn.push(obj);
}
// set excel headers
            sheet.columns = excelColumn;
                // empty first row
                _.each(excelColumn, function (column, index) {
                        let headerIndex = getNameFromNumber(index) + "1";
                        sheet.getCell(headerIndex).value = "";
                });

sheet.getRow(startOfRows).values = _.map(excelColumn, "header");
            sheet.getRow(startOfRows).style = options.columns[0].style||{}


data = _.map(data, function (d,index) {
  let obj = {};
  _.each(options.columns, function (column) {
      if(column.formula){
        obj[column.field]={
            formula:column.formula.replace(/@@/g,(dataFirstRow+index))
        }
      }
      else if (column.field) {
          // seperate column info
          let columnField = column.field.split(".");
          if (columnField.length == 2) {
              if (_.isArray(d[columnField[0]])) {
                  obj[column.field] = _.map(d[columnField[0]], columnField[1]).join(", ");
              }
              else if (d[columnField[0]] != null && d[columnField[0]][columnField[1]] != null) {

                  // is boolean
                  if (_.isBoolean(d[columnField[0]][columnField[1]])) {
                      obj[column.field] = d[columnField[0]][columnField[1]] ? "YES" : "NO";
                  }
                  // is Numeric
                  else if (_.isNumber(d[columnField[0]][columnField[1]])) {

                      obj[column.field] = d[columnField[0]][columnField[1]];
                  }
                  // default String
                  else {

                      obj[column.field] = d[columnField[0]][columnField[1]].toString();
                  }
              }
              // Empty
              else {
                  obj[column.field] = "-";
              }
          }
          else if (d[columnField[0]] != null) {

              // is Bool
              if (_.isBoolean(d[columnField[0]])) {
                  obj[column.field] = d[columnField[0]] ? "YES" : "NO";
              }

              // is numeric
              else if (_.isNumber(d[columnField[0]])) {
                  obj[column.field] = d[columnField[0]];
              }

              // string
              else {
                  obj[column.field] = d[columnField[0]].toString();
              }
          }
          // Empty
          else {
              obj[column.field] = "-";
          }
      }
  });
  startOfRows++;
  return obj;

});


    // append vallue to sheet
    for (let i in data) {
        sheet.addRow(data[i])
    }

    let excludeRowNumbers=[sampleExcelDataLastRow, sampleExcelDataLastRow - sampleExcelDataLength + 1]
    let differenceFromLastRow=dataLastRow-sampleExcelDataLastRow
    for (let row of extraDataForMerge) {
        let formula = row.cell1 + ":" + row.cell2
        let allNumbersInFormula = _.uniq(formula.match(/\\d+/g) || []).map(Number);
        _.each(allNumbersInFormula, (num) => {
            if (/*!_.includes(excludeRowNumbers,num) &&*/ num > dataFirstRow) {
                // if (num > sampleExcelDataLastRow) {
                    // let rowDiffWithLastRow = num - sampleExcelDataLastRow
                    let regexEnd = new RegExp("(?<=[A-Z$])" + num + "(?![0-9])", "g");
                    formula = formula.replace(regexEnd, num+differenceFromLastRow)
                // }
            }
        })
        sheet.mergeCells(formula);
    }
    for (let row of extraDataByCell) {
        let currentRowNo = row.data[0].row
            let rowDiffWithLastRow = 0
        if (row.data[0].row > sampleExcelDataLastRow) {
            rowDiffWithLastRow = row.data[0].row - sampleExcelDataLastRow
            currentRowNo = dataLastRow + rowDiffWithLastRow
        }
        if(row.hidden){
            sheet.getRow(currentRowNo).hidden=true
        }
        for (let cell of row.data) {
            if (!cell.isMerged || cell.mergeCount) {
                // // }
                // // else{
                // let currentCellChar = getNameFromNumber(cell.cell - 1)
                // if (cell.mergeCount) {
                //     sheet.mergeCells(""+currentCellChar+currentRowNo+":"+getNameFromNumber(cell.mergeCount + cell.cell - 1)+currentRowNo);
                // }
            let currentCellChar=getNameFromNumber(cell.cell-1)
            if (_.isObject(cell.value)) {
                if (cell.value.formula) {
                    let allNumbersInFormula= _.uniq(cell.value.formula.match(/\\d+/g)||[]).map(Number);
                    let formula=cell.value.formula
                    _.each(allNumbersInFormula,(num)=>{
                        if(/*!_.includes(excludeRowNumbers,num) &&*/ num > dataFirstRow){
                        let regexEnd = new RegExp("(?<=[A-Z$])" + num + "(?![0-9])", "g");
                        formula=formula.replace(regexEnd,num+differenceFromLastRow)
                        }
                    })
                    // sheet.getCell(currentCellChar + currentRowNo).value = {
                    //         formula: formula
                    //     }
                    //  let regexCurrentRow = new RegExp("(?<![0-9])" + cell.row + "(?![0-9])", "g");
                     
                    // let regexCurrentRow = new RegExp("(?<=[A-Z$])" + cell.row + "(?![0-9])", "g");
                    // let regexEnd = new RegExp("(?<=[A-Z$])" + sampleExcelDataLastRow + "(?![0-9])", "g");
                    // formula=formula.replace(regexEnd, dataLastRow).replace(regexCurrentRow,currentRowNo)
                    sheet.getCell(currentCellChar + currentRowNo).value = {
                        formula: formula
                    }
                }
                else {
                    sheet.getCell(currentCellChar + currentRowNo).value = cell.value
                }
            }
            else {
                sheet.getCell(currentCellChar + currentRowNo).value = cell.value
            }
            sheet.getCell(currentCellChar + currentRowNo).style = cell.style
        }
    }
    }

    _.each(worksheetData,(val,key)=>{
        if(key == "autoFilter"){
            let formula = val
        let allNumbersInFormula = _.uniq(formula.match(/\\d+/g) || []).map(Number);
        _.each(allNumbersInFormula, (num) => {
            if (/*!_.includes(excludeRowNumbers,num) &&*/ num > dataFirstRow) {
                // if (num > sampleExcelDataLastRow) {
                    let regexEnd = new RegExp("(?<=[A-Z$])" + num + "(?![0-9])", "g");
                    formula = formula.replace(regexEnd, num+differenceFromLastRow)
                // }
            }
        }) 
        val=formula
        }
        sheet[key]=val
    })

    if (!fs.existsSync(destPath)) {
      fs.mkdirSync(destPath);
  }
  await new Promise((resolve, reject) => {
      workbook.xlsx.writeFile(destPath + excelName).then(function (err) {
          if (err) {
              console.log("excelll", err);
              reject(err);
          }
          else {
              resolve();
          }
      });
  });
  console.log("/excel-temp/" + excelName)

  `

    let fd = fs.openSync('./generated.js', 'w');
    fs.appendFileSync(fd, defaultStartText);
    fs.appendFileSync(fd, defaultEndText);
    console.log("---------- code generated ------------")
}
(async () => {
    await excelData()
    let gg = require('./generated')
    gg.generateExcel()
})()
app.get('/', function (req, res) {
    res.send('Hello World!')
})

app.listen(3000, function () {
    console.log('Listening on port 3000...')
})
