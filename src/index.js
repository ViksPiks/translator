const path = require("path");
const fs = require("fs");
const ExcelJS = require("exceljs");

const pathToExcelFile = path.join(__dirname, "required-assets/index.xlsx");
const pathToHtmlFile = path.join(__dirname, "required-assets/index.html");
const pathToResultFile = path.join(__dirname, "required-assets/result.html");

async function readHTMLFile() {
  return await new Promise((resolve, reject) => {
    fs.readFile(pathToHtmlFile, "utf8", (err, data) => {
      if (err) {
        reject(err);
      } else {
        resolve(data);
      }
    });
  });
}

async function saveToNewHTMLFile(data) {
  return await new Promise((_, reject) => {
    fs.writeFile(pathToResultFile, data, "utf-8", (err) => {
      if (err) {
        console.error("Error writing to file:", err);
        reject(err);
      }
      console.log("File created and data written successfully!");
    });
  });
}

async function runTranslator() {
  const workbook = new ExcelJS.Workbook();

  await workbook.xlsx.readFile(pathToExcelFile);

  const worksheet = workbook.getWorksheet("e-Cat");

  let htmlFile = await readHTMLFile();

  worksheet.getColumn(1).eachCell((cell, rowNumber) => {
    const { value } = worksheet.getRow(rowNumber).getCell(2);

    htmlFile = htmlFile.replace(cell.value, value);
  });

  await saveToNewHTMLFile(htmlFile);
}

runTranslator();
