# [Anki＃5] 用AutoHotkey由文字檔產生Anki閃卡

* 需求：讀取英譯中設定檔再以AnkiConnect產生Anki閃卡資料

## Anki端的設定
* 執行腳本前必須啟動Anki系統
* 必須安裝附加元件 AnkiConnect，代碼：**2055492159**
* 使用Default牌組
* 產生「基本型（含反向卡片）」、「克漏字」兩種筆記類型

## 設定檔格式
* 預設設定檔名是`C:\temp\anki.txt`
* 以半型分號開頭的是註解
* 必須以半形井號開頭，代表閃卡的標籤(hash tag)，例如：
```
#tag1 #tag1-1 #tag1-1-1
```
* 閃卡資料有三種格式：
	1. 英文單字 : 會自動到雅虎奇摩字典網站取回中文釋義與語音檔
	2. 片語［Tab］說明：以［Tab］為分隔字元
	3. 克漏字句子：必須有`{{c1::xxx}}`的克漏字關鍵字


## 參考
* [\[AHK＃40\] 用AutoHotkey＋AnkiConnect寫閃卡資料到Anki 的簡單範例](https://youtu.be/Pt7XPYzf2gc)
* [\[AHK＃41\] 讀取字典網站、解析網頁與下載MP3檔案的方法](https://youtu.be/LPk4vVnjfHk)

