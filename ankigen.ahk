/*
; 分號開頭表示是註解行。井號開頭是標籤，多個標籤以逗點分隔
#waytogo-6, #unit1
; 一行一個單字，自動到奇摩網站取回翻譯和語音檔
park
; 片語：一行用Tab鍵分隔兩個欄位
watch a baseball game	看棒球賽
; 一行一句，含有 {{c1:: 時是克漏字句子
I love {{c1::summer}} {{c2::vacation}}. 我喜歡暑假
*/
#SingleInstance,Force

Global _EXPL_ := 1  ; 解釋
Global _NOTION_ := 2  ; 詞類
Global _PRONOUNCE_ := 3  ; 發音

; 輸入文字檔有3種格式
;  1=英文單字
;  2=片語 [Tab] 中文解釋
;  3=克漏字
Global _ACTION_WORD_ := 1
Global _ACTION_PHRASE_ := 2
Global _ACTION_CLOZE_ := 3

Global sInputFile := "c:\temp\anki.txt"  ; 輸入的設定檔, 檔案編碼：UTF-8
Global sSeparator := A_Tab  ; 輸入檔的分隔字元

Global action_word, action_phrase, action_cloze

action_word =
(
{
    "action": "addNote",
    "version": 6,
    "params": {
      "note": {
        "deckName": "Default",
        "modelName": "基本型（含反向卡片）",
        "options": {
          "allowDuplicate": true,
          "duplicateScope": "deck",
          "duplicateScopeOptions": {
            "deckName": "Default",
            "checkChildren": false
          }
        },
        "fields": {
          "正面": "{1} {3}",
          "背面": "{2}"
        },
        "tags": [
          "{4}"
        ],
        "audio": {
          "url": "https://s.yimg.com/bg/dict/dreye/live/f/{1}.mp3",
          "filename": "{1}.mp3",
          "skipHash": "{5}",
          "fields": [
            "正面"
          ]
        }
      }
    }
  }
)  ; action_word  

action_phrase =
(
{
    "action": "addNote",
    "version": 6,
    "params": {
      "note": {
        "deckName": "Default",
        "modelName": "基本型（含反向卡片）",
        "options": {
          "allowDuplicate": true,
          "duplicateScope": "deck",
          "duplicateScopeOptions": {
            "deckName": "Default",
            "checkChildren": false
          }
        },
        "fields": {
          "正面": "{1}",
          "背面": "{2}"
        },
        "tags": [
          "{3}"
        ]
      }
    }
  }
)  ; action_word  

action_cloze =
(
{
    "action": "addNote",
    "version": 6,
    "params": {
      "note": {
        "deckName": "Default",
        "modelName": "克漏字",
        "options": {
          "allowDuplicate": true,
          "duplicateScope": "deck",
          "duplicateScopeOptions": {
            "deckName": "Default",
            "checkChildren": false
          }
        },
        "fields": {
          "文字": "{1}",
          "背面額外": "{2}"
        },
        "tags": [
          "{3}"
        ]
      }
    }
  }
)  ; action_cloze

f1::
  ; {1}=English, {2}=中文, {3}=hash

FileEncoding , UTF-8
InputBox, sInputFile, 輸入檔名, 輸入設定檔名：, , 300, 120, , , , , %sInputFile%
if ErrorLevel
	Return

if !FileExist(sInputFile) {
	MsgBox %sInputFile% 必須先建立。
	Return
}
_iCount := 0

; 先算出設定檔有幾行，供進度條使用
FileRead _oFile, %sInputFile%
StringReplace _oFile, _oFile, `n, `n, All UseErrorLevel
_iLines := ErrorLevel
;MsgBox Total number of lines is %_iLines%

_sTags := ""
Progress R0-%_iLines%
Loop, read, %sInputFile% 
{
	Progress, %A_Index%, %A_LoopFileName%, 閃卡處理中..., 產生Anki閃卡
	FileReadLine, _sLine, %sInputFile%, %A_Index%
	;MsgBox line #%A_Index%=%_sLine%.
	_sFirstChar := SubStr(_sLine, 1, 1)
	if (_sFirstChar == "#") {  ; tag必須以井號開始開頭
		; #tag1 => "#tag1",
		_sTags := StrReplace(_sLine, " ", """,""")
		;MsgBox tags=%_sTags%
		continue
	} else if (_sFirstChar == ";") {  ;; 第一個字元分號是註解行
		continue
	}
	_aTokens := StrSplit(_sLine, sSeparator)   ; (string, deliChar, omitChar)
	_iTokens := _aTokens.MaxIndex()
	if (_iTokens > 0) {
		_iCount++
		_sEnglish := _aTokens[1]
		;_iPos := InStr(_sEnghish, "{{c1")
		_iPos := RegExMatch(_sLine, "{{c1::.*}}")
		if (_iPos > 0) {  ; 克漏字
			output(httpClient, _ACTION_CLOZE_, _sEnglish, "", _sTags)
		} else if (_iTokens == 1) {  ; 單字
			output(httpClient, _ACTION_WORD_, _sEnglish, "", _sTags)	
		} else if (_iTokens == 2) {  ; 片語
			_sChinese := _aTokens[2]
			output(httpClient, _ACTION_PHRASE_, _sEnglish, _sChinese, _sTags)
		}
      ;MsgBox, Field number %A_Index% is %_sEnglish%, %_sChinese%
	}
}
Progress, Off
MsgBox 轉換筆數: %_iCount%
Return

;; 寫出單字
output(httpClient, iActionKind, sEnglish, sChinese, sTags) {
	_skipHash := A_TickCount
	if (sTags == "") {
		sTags := "#AHK"
	}
	if (iActionKind == _ACTION_WORD_) {  ; 單字
		action := action_word
		_aTokens := translate(sEnglish)  ; 取翻譯與語音檔
		_sChineseGet := _aTokens[_EXPL_]
    ;; \r與\t會造成寫入失敗，必須先行轉換
		_sChineseGet := StrReplace(_sChineseGet, "`r", " ")
		_sChineseGet := StrReplace(_sChineseGet, "`t", " ")
		_sEnghishGet := " " . _aTokens[_PRONOUNCE_]
		_sEnghishGet := StrReplace(_sEnghishGet, "`r", " ")
		
		;MsgBox output chinese=%_sChineseGet%
		
		action := Format(action, sEnglish, _sChineseGet, _sEnghishGet, sTags, sSkipHash)
	} else if (iActionKind == _ACTION_PHRASE_) {
		action := action_phrase
		action := Format(action, sEnglish, sChinese, sTags)
	} else if (iActionKind == _ACTION_CLOZE_) {
		;MsgBox Cloze english=%sEnglish%, chinese=%sChinese%,tags=%sTags%
		action := Format(action_cloze, sEnglish, sChinese, sTags)
	}
	;MsgBox action=%action%
	
	url := "http://localhost:8765"
	httpClient := ComObjCreate("WinHttp.WinHttpRequest.5.1")
	httpClient.Open("POST", url, false)
	httpClient.SetRequestHeader("Content-Type", "application/json")
	httpClient.Send(action)
	httpClient.WaitForResponse()
	response := httpClient.ResponseText
	
  ;MsgBox result=%response%
	if (response != "null") {
		ankiJson := jsonAHK(response)
		anki_ID := ankiJson.result
		error := ankiJson.error
		if (error != "" && error != "null") {
			MsgBox error=%error%.
		}
	}
	
	return
}

;; 傳入要搜尋的文字，傳回查詢到的結果
;; 再將結果以Markdown格式輸出到檔案
translate(sSearch) {
  ;msgbox sSearch=%sSearch%
	url := "https://tw.dictionary.search.yahoo.com/search?p=" . sSearch . "&fr=sfp&iscqry="
	
	httpClient := ComObjCreate("WinHttp.WinHttpRequest.5.1")
	httpClient.Open("POST", url, false)
	httpClient.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded")
	httpClient.Send()
	httpClient.WaitForResponse()
	Result := httpClient.ResponseText
	
	html := ComObjCreate("HTMLFile")
	html.write(Result)
	
	elements := html.getElementsByTagName("div") 
	_sPronounce := ""  ; 發音
	_sNotion := ""     ; 詞類
	_sDictionaryExplanation := ""  ; 解釋
	
	Loop % elements.length
	{
		ele := elements[A_Index-1] ; zero based collection
		_sClassName := ele.className
     ;;Msgbox %A_Index%=%_sClassName%
		if (InStr(_sClassName, "pos_button") > 0) {
			_sNotion := ele.innerHTML
		} else if (InStr(_sClassName, "compList d-ib") > 0) {
			_sPronounce := ele.innerText
		} else if (InStr(_sClassName, "dictionaryExplanation") > 0) {
			_sDictionaryExplanation .= _sNotion " " ele.innerHTML . "`r"
			_sNotion := ""
		}
	}
  ;MsgBox expl=%_sDictionaryExplanation%, notion=%_sNotion%, pron=%_sPronounce%
	_aTokens := []
	_aTokens.Push(_sDictionaryExplanation, _sNotion, _sPronounce)
	
	return _aTokens
}

jsonAHK(s) {
	static o := GetObjJScript()
	o.language:="jscript"
	return o.eval("(" s ")")
}

GetObjJScript()
{
	if !FileExist(ComObjFile := A_ScriptDir "\JS.wsc")
		FileAppend,
         (LTrim
            <component>
            <public><method name='eval'/></public>
            <script language='JScript'></script>
            </component>
         ), % ComObjFile
	Return ComObjGet("script:" . ComObjFile)
}
