使用步驟
- npm i
- npm run excel
- 預設使用example.json來示範
- 建議可以安裝vscode Excel Viewer 挺好用的

參數說明
```
app.js

const excelConfig = [{
    entry:'./example.xlsx',  //讀取excel位置
    sheetNo:0, //讀取excel的哪張sheet
    outputDir:'./',  //導出位置
    outputName:'example', //導出檔案名稱
}]
```
核心邏輯 getMultiValueArray 函式
```
 app.js

 可使用參數
 * 在這邊客製化你的excel轉出來的資料
 * @param {*} sheetName sheetName
 * @param {*} sheet sheet
 * @param {*} range 獲取有值X,Y的範圍
 * @param {*} startC excel有值的最左邊
 * @param {*} startR excel有值的最上面
 * @param {*} endC excel有值的最右邊
 * @param {*} endR excel有值的最下面
 * @param {*} lengthC excel有值的左到右長度
 * @param {*} lengthR excel有值的上到下長度
 * returns json資料

```
