// Requiring the module
import { readFile, utils, writeFile } from "xlsx";
import { Workbook } from "excel4node";
import { existsSync } from "fs"; // Or `import fs from "fs";` with ESM

// Reading our test file
const fileName = "list_form_2";

// result file name
const retFileName = fileName + "-Done.xlsx";
var workbook = new Workbook();
var worksheet = workbook.addWorksheet("Sheet1");
workbook.write(retFileName);

const file = readFile("./" + fileName + ".xlsx");
const sheets = file.SheetNames;
const data = [];
const result = [];

for (let i = 0; i < sheets.length; i++) {
  const temp = utils.sheet_to_json(file.Sheets[file.SheetNames[i]]);
  temp.forEach((res) => {
    data.push(res);
  });
}

console.log("Data__length", data.length);
console.log("Sample_first_10:");

for (let i = 0; i < data.length; i++) {
  const person = data[i];

  if (i < 10) console.log("person", person);
  let phone_1 = person.phone_1;
  let mail_1 = person.mail_1;
  let isFound = false;
  for (let j = 0; j < data.length; j++) {
    const temp = data[j];
    if (temp.mail_2 && temp.mail_2 != "" && temp.mail_2 === mail_1) {
      isFound = true;
      break;
    }
    if (temp.phone_2 && temp.phone_2 != "" && temp.phone_2 === phone_1) {
      isFound = true;
      break;
    }
  }
  if (!isFound) {
    result.push(person);
  }
}

function saveFile() {
  let path = "./" + retFileName;
  console.log("Saving_file...");
  if (existsSync(path)) {
    const ws = utils.json_to_sheet(result);
    const fileResult = readFile(path);
    console.log("Result_length_final", result.length);
    utils.book_append_sheet(fileResult, ws, "result");
    // Writing to our file
    writeFile(fileResult, path);
    clearInterval(iid);
  }
}
let iid = setInterval(() => {
  saveFile();
}, 1000 * 0.1);
