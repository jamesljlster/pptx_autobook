# PPTX Autobook

將 PowerPoint (pptx) 投影片，依據標題、章節標題整理成一份 Word (docx) 文件。

### 相依套件

基於 Anaconda3 2019.07：
- [python-pptx](https://pypi.org/project/python-pptx/)
- [python-docx](https://pypi.org/project/python-docx/)

### 使用方式

開啟 Python 執行環境（CMD 或者 PowerShell），將專案資料夾作為模組呼叫：

```bash
python pptx_autobook <參數列表>
```

##### 餐數定義：

- --pptx-src: pptx 檔案路徑列表。
- --docx-out: docx 產生路徑，須注意檔案如果已經存在，程式會覆蓋該檔案。
- --docx-in: docx 輸入路徑，可帶入文件模板，或是半成品檔案。
- -h: 顯示說明，絕對沒有這份 README 詳細。

##### 使用範例

今天我有兩份 pptx 需要整理成一份 Word，分別為 lesson1.pptx 與 lesson2.pptx，並且已經準備一份 docx 模板 template.docx。  
於是依照下面的指令呼叫：

```bash
python pptx_autobook --pptx-src lesson1.pptx lesson2.pptx --docx-in template.docx --docx-out book.docx
```

程式將把 lesson1.pptx 與 lesson2.pptx 的內容依據 template.docx 的格式整理成 book.docx。

### 創作發想

想要我把 PPT 裡面的投影片一張一張貼到 Word 上，還要整理標題跟目錄 :confused:  
Over my dead body :smiley: