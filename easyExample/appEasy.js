const xlsx = require("xlsx");
const fs = require("fs");
// 讀檔案
// const wb = xlsx.readFile("./data.xlsx",{cellDates:true});
// const wb = xlsx.readFile("./data.xlsx", { dateNF: "mm/dd/yyyy" });
const wb = xlsx.readFile("./data.xlsx");

// 讀檔案中的sheet
// console.log(wb.SheetNames);
const ws = wb.Sheets["GOOGLE企業"];

// 讀sheet中的data
// const data = xlsx.utils.sheet_to_json(ws, { raw: false });
const data = xlsx.utils.sheet_to_json(ws);

// 轉換格式
let newData = [];
newData = data.map((d) => {
  return d;
});

// 導出位置
fs.writeFileSync("./datajson.json", JSON.stringify(newData,null,2));
