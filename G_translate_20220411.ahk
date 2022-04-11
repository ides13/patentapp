#SingleInstance force

; 預設翻譯語言為正體中文。若要翻譯其他語言時：
; 1、先按左邊的Ctrl及Win，再按左邊的Alt
; 2、出現選擇語言的 Menu ，可以選擇要翻譯的語言。
; 3、完成選擇翻譯的對象語言後，「選取」欲翻譯文字，然後先按左邊的Ctrl，再按左邊的Alt。
; 4、程式會自動將選取的文字，送到google翻譯，翻譯成對象語言，之後再傳送到剪貼簿。
; 5、最後，在希望的位置處，按 Ctrl+V，將翻譯後的文字貼上。

; 其他，要新增語言時，可以參考下方的「LanguageCodeArray」的語言代碼。

;另一個常用功能，「無格式貼上」。
; Ctrl, Win及v鍵，「無格式貼上」並替換字元。


HotstringMenuAutoExecute:
    current_code := "zh-TW" 
    GlobalIndex := 0

	gtranscode :=      {"【正體中文】" : "zh-TW"
					   ,"【日文】" : "ja"
					   ;,"【簡體中文】" : "zh-CN"
					   ;,"【簡體中文】" : "zh-CN"
                       ,"【英文】" : "en"}

	HotstringMenuF("gtransmenu", "change_lang"  , gtranscode*)
Return

;========================================================================================= 
;原始碼源自 「HotstringSubMenu.ahk 」
;https://www.autohotkey.com/boards/viewtopic.php?t=69791
;作者為「Jack's AutoHotkey Blog」的作者
;=========================================================================================
HotstringMenuF(MenuName, HandlerName, MenuArray*)
{
  global GlobalIndex ;初始值為零，每執行 HotstringMenuF 一次，會增加 Index 的值。
  ArrayLength := MenuArray.SetCapacity(0) ; Get array size
  For Each, Item in MenuArray {
        ;If (ArrayLength < 10)
        If ((A_Index+GlobalIndex)<0 )
          Shortcut := "&" . (A_Index+GlobalIndex)
        Else 
          Shortcut := "&" . Chr(A_Index+96+GlobalIndex)
          If (InStr(Item,"Brk"))      ; Add column breaks to long menus
          {
             Item := StrReplace(Item,"Brk")
             Options := "+BarBreak"
          }
          Else    
              Options := ""

		; Bind output data to the DoSearch()
        Handler := Func(HandlerName).Bind(Each,Item)

        ; If (Each = A_index)    ; Simple array
        ;    Menu, % MenuName, Add, % Shortcut " — " Item, % Handler, % Options
        ; Else
           Menu, % MenuName, Add, %  Shortcut "  " Each, % Handler, % Options
  }
  GlobalIndex := ArrayLength + GlobalIndex  ;初始值為零，每執行 HotstringMenuF 一次，會增加 Index 的值。
}

change_lang(Key,Item)
{
    global current_code
    current_code :=  Item      
	MsgBox, , , % "變更翻譯語言為：" . Key . "。", 1  ;MsgBox
}

;先按左邊的Ctrl、Win，再按左邊的Alt
;顯示翻譯語言 Menu 。
^#LAlt:: Menu, gtransmenu, Show

;先按左邊的Ctrl，再按左邊的Alt
;複製到簡貼簿後，用googl翻譯，在傳到簡貼簿。
^LAlt::
    global current_code
    ;MsgBox, %current_code%

	SavedClipTemp := ClipboardAll
	Clipboard =                         ; empty
	SendInput, ^c                       ; copy highlighted text
	ClipWait %timeoutSeconds%           ; Wait for the copied text to arrive at the clipboard.
	Clipboard := RegExReplace(Clipboard, "s)^\s+|\s+$", "")
	Clipboard := RegExReplace(Clipboard, "m)^[ `t]+|[ `t]+$", "")
	;Clipboard := RegExReplace(Clipboard, "\r\n", " ") 

	if ErrorLevel
	{
	  Clipboard := SavedClipTemp
	  SoundBeep
	  return
	}
	Else
	{ 
	  ;日文 -> 中文
	  ;Clipboard := GoogleTranslate(Clipboard, "ja", "zh-TW")	   
	  ;自動識別 -> 日文
	  ;Clipboard := GoogleTranslate(Clipboard, "auto", "ja")
	  ;自動識別 -> 英文
	  ;Clipboard := GoogleTranslate(Clipboard, "auto", "en")	  
	  ;自動識別 -> 簡體中文="zh-CN"      
	  ;Clipboard := GoogleTranslate(Clipboard, from :="auto", to :="zh-CN")	  
	  ;自動識別 -> 正體中文="zh-TW"
	  ;Clipboard := GoogleTranslate(Clipboard, from :="auto", to :="zh-TW")	  

	  ;利用 語言 Menu 進行選擇
	  ;先按左邊的Alt再按左邊的Ctrl
	  ;出現選擇語言的 Menu ，可以選擇要翻譯的語言。
	  Clipboard := GoogleTranslate(Clipboard, from := "auto", to := current_code)	
	  ;FileAppend, %Clipboard%`n, ToTransList.txt
	  
	  ClipWait %timeoutSeconds% 
	}
	;SendInput, !{TAB}
	sleep 300
	SendInput, {Down}
	;SendInput, ^v 
Return


;======================================
;Ctrl, Win及v鍵，「無格式貼上」並替換字元。
^#v::
    MyRadio := 0
    ;讀取預設文件及自編輯文件。
    ;MsgBox You entered:`n是否有執行`nMyRadio=%MyRadio%
    OutputVar=
    if (MyRadio = 0){
        FileRead, OutputVar, %A_ScriptDir%\ReplacePase.txt
        if ErrorLevel {
            OutputVar=
        }
    }  
    OutputVar := % MyEdit . ";" . OutputVar
    OutputVar := RegExReplace(OutputVar, "`r`n", "")
    OutputVar := RegExReplace(OutputVar, "`n", "")
    OutputVar := RegExReplace(OutputVar, ";;", ";")
    ;FileAppend, `n%OutputVar%, %A_ScriptDir%\MyFile.txt
    ;MsgBox, , , 需要替代的文字：`n%OutputVar% , 2.5
    
	ClipSaved := ClipboardAll

	Text = %Clipboard%

	if (MyCheckbox1 = 0){
	    ;刪除格式。
        Text := RegExReplace(Text, "s)^\s+|\s+$", "")
        Text := RegExReplace(Text, "m)^[ `t]+|[ `t]+$", "")		
 	    ;將「換行字元」取代成「空白字元」。
        Text := RegExReplace(Text, "\r\n", " ") 
	}
	if (MyCheckbox2 = 1){
		Text := RegExReplace(Text, "> ", "") 
		Text := RegExReplace(Text, ">" , "") 
 		Text := RegExReplace(Text, "該", "")        		
	}

    replacelist := StrSplit(OutputVar, ";")
    for k, item in replacelist
    {
      Replacetext := StrSplit(item, "|")
      Text := RegExReplace(Text, Replacetext[1], Replacetext[2])
	  ;Text := StrReplace(Text, Replacetext[1], Replacetext[2])
    }

	Clipboard := Text
	Send ^v
	Sleep, 1000  ; 需等剪貼簿的內容送出後才能復原原本的內容
	Clipboard := ClipSaved
	ClipSaved =
	Text=
Return
;======================================
;======================================


;======================================
; Google Translate script
; Take a string in any language and translate to any other language.
;
; Credited to teadrinker: https://www.autohotkey.com/boards/viewtopic.php?f=5&t=40876#p186877
; Slightly modified by Osprey to allow for determining and using the system language.
; Should be run with the Unicode version of AutoHotkey.
;
; Sample usage
; MsgBox, % GoogleTranslate("今日の天気はとても良いです")			; Translate string from auto-detected language to system language
; MsgBox, % GoogleTranslate("今日の天気はとても良いです", "jp", "en")	; Translate string from Japanese to English

GoogleTranslate(str, from := "auto", to := 0) {
   static JS := GetJScripObject(), _ := JS.( GetJScript() ) := JS.("delete ActiveXObject; delete GetObject;")
   
   if(!to)				; If no "to" parameter was passed
      to := GetISOLanguageCode()	; Assign the system (OS) language to "to"

   if(from = to)			; If the "from" and "to" parameters are the same
      Return str			; Abort translation and return the original string

   json := SendRequest(JS, str, to, from, proxy := "")
   if(!json or InStr(json, "document.getElementById('captcha-form')"))	; If no response (ex. internet down) or spam is detetected
     Return str				; Return the original, untranslated string
   oJSON := JS.("(" . json . ")")

   if !IsObject(oJSON[1])  {
      Loop % oJSON[0].length
         trans .= oJSON[0][A_Index - 1][0]
   }
   else  {
      MainTransText := oJSON[0][0][0]
      Loop % oJSON[1].length  {
         trans .= "`n+"
         obj := oJSON[1][A_Index-1][1]
         Loop % obj.length  {
            txt := obj[A_Index - 1]
            trans .= (MainTransText = txt ? "" : "`n" txt)
         }
      }
   }
   if !IsObject(oJSON[1])
      MainTransText := trans := Trim(trans, ",+`n ")
   else
      trans := MainTransText . "`n+`n" . Trim(trans, ",+`n ")

   from := oJSON[2]
   trans := Trim(trans, ",+`n ")
   Return trans
}

; Take a 4-digit language code or (if no parameter) the current language code and return the corresponding 2-digit ISO code
GetISOLanguageCode(lang := 0) {
   LanguageCodeArray := { 0436: "af" ; Afrikaans
			, 041c: "sq" ; Albanian
			, 0401: "ar" ; Arabic_Saudi_Arabia
			, 0801: "ar" ; Arabic_Iraq
			, 0c01: "ar" ; Arabic_Egypt
			, 1001: "ar" ; Arabic_Libya
			, 1401: "ar" ; Arabic_Algeria
			, 1801: "ar" ; Arabic_Morocco
			, 1c01: "ar" ; Arabic_Tunisia
			, 2001: "ar" ; Arabic_Oman
			, 2401: "ar" ; Arabic_Yemen
			, 2801: "ar" ; Arabic_Syria
			, 2c01: "ar" ; Arabic_Jordan
			, 3001: "ar" ; Arabic_Lebanon
			, 3401: "ar" ; Arabic_Kuwait
			, 3801: "ar" ; Arabic_UAE
			, 3c01: "ar" ; Arabic_Bahrain
			, 042c: "az" ; Azeri_Latin
			, 082c: "az" ; Azeri_Cyrillic
			, 042d: "eu" ; Basque
			, 0423: "be" ; Belarusian
			, 0402: "bg" ; Bulgarian
			, 0403: "ca" ; Catalan
			, 0404: "zh-CN" ; Chinese_Taiwan
			, 0804: "zh-CN" ; Chinese_PRC
			, 0c04: "zh-CN" ; Chinese_Hong_Kong
			, 1004: "zh-CN" ; Chinese_Singapore
			, 1404: "zh-CN" ; Chinese_Macau
			, 041a: "hr" ; Croatian
			, 0405: "cs" ; Czech
			, 0406: "da" ; Danish
			, 0413: "nl" ; Dutch_Standard
			, 0813: "nl" ; Dutch_Belgian
			, 0409: "en" ; English_United_States
			, 0809: "en" ; English_United_Kingdom
			, 0c09: "en" ; English_Australian
			, 1009: "en" ; English_Canadian
			, 1409: "en" ; English_New_Zealand
			, 1809: "en" ; English_Irish
			, 1c09: "en" ; English_South_Africa
			, 2009: "en" ; English_Jamaica
			, 2409: "en" ; English_Caribbean
			, 2809: "en" ; English_Belize
			, 2c09: "en" ; English_Trinidad
			, 3009: "en" ; English_Zimbabwe
			, 3409: "en" ; English_Philippines
			, 0425: "et" ; Estonian
			, 040b: "fi" ; Finnish
			, 040c: "fr" ; French_Standard
			, 080c: "fr" ; French_Belgian
			, 0c0c: "fr" ; French_Canadian
			, 100c: "fr" ; French_Swiss
			, 140c: "fr" ; French_Luxembourg
			, 180c: "fr" ; French_Monaco
			, 0437: "ka" ; Georgian
			, 0407: "de" ; German_Standard
			, 0807: "de" ; German_Swiss
			, 0c07: "de" ; German_Austrian
			, 1007: "de" ; German_Luxembourg
			, 1407: "de" ; German_Liechtenstein
			, 0408: "el" ; Greek
			, 040d: "iw" ; Hebrew
			, 0439: "hi" ; Hindi
			, 040e: "hu" ; Hungarian
			, 040f: "is" ; Icelandic
			, 0421: "id" ; Indonesian
			, 0410: "it" ; Italian_Standard
			, 0810: "it" ; Italian_Swiss
			, 0411: "ja" ; Japanese
			, 0412: "ko" ; Korean
			, 0426: "lv" ; Latvian
			, 0427: "lt" ; Lithuanian
			, 042f: "mk" ; Macedonian
			, 043e: "ms" ; Malay_Malaysia
			, 083e: "ms" ; Malay_Brunei_Darussalam
			, 0414: "no" ; Norwegian_Bokmal
			, 0814: "no" ; Norwegian_Nynorsk
			, 0415: "pl" ; Polish
			, 0416: "pt" ; Portuguese_Brazilian
			, 0816: "pt" ; Portuguese_Standard
			, 0418: "ro" ; Romanian
			, 0419: "ru" ; Russian
			, 081a: "sr" ; Serbian_Latin
			, 0c1a: "sr" ; Serbian_Cyrillic
			, 041b: "sk" ; Slovak
			, 0424: "sl" ; Slovenian
			, 040a: "es" ; Spanish_Traditional_Sort
			, 080a: "es" ; Spanish_Mexican
			, 0c0a: "es" ; Spanish_Modern_Sort
			, 100a: "es" ; Spanish_Guatemala
			, 140a: "es" ; Spanish_Costa_Rica
			, 180a: "es" ; Spanish_Panama
			, 1c0a: "es" ; Spanish_Dominican_Republic
			, 200a: "es" ; Spanish_Venezuela
			, 240a: "es" ; Spanish_Colombia
			, 280a: "es" ; Spanish_Peru
			, 2c0a: "es" ; Spanish_Argentina
			, 300a: "es" ; Spanish_Ecuador
			, 340a: "es" ; Spanish_Chile
			, 380a: "es" ; Spanish_Uruguay
			, 3c0a: "es" ; Spanish_Paraguay
			, 400a: "es" ; Spanish_Bolivia
			, 440a: "es" ; Spanish_El_Salvador
			, 480a: "es" ; Spanish_Honduras
			, 4c0a: "es" ; Spanish_Nicaragua
			, 500a: "es" ; Spanish_Puerto_Rico
			, 0441: "sw" ; Swahili
			, 041d: "sv" ; Swedish
			, 081d: "sv" ; Swedish_Finland
			, 0449: "ta" ; Tamil
			, 041e: "th" ; Thai
			, 041f: "tr" ; Turkish
			, 0422: "uk" ; Ukrainian
			, 0420: "ur" ; Urdu
			, 042a: "vi"} ; Vietnamese
   If(lang)
     Return LanguageCodeArray[lang]
   Else Return LanguageCodeArray[A_Language]
}

SendRequest(JS, str, tl, sl, proxy) {
   ;sl := "ja"
   ;tl := "zh-TW"

   ComObjError(false)
   http := ComObjCreate("WinHttp.WinHttpRequest.5.1")
   ( proxy && http.SetProxy(2, proxy) )
   http.open( "POST", "https://translate.google.com/translate_a/single?client=t&sl="
      . sl . "&tl=" . tl . "&hl=" . tl
      . "&dt=at&dt=bd&dt=ex&dt=ld&dt=md&dt=qca&dt=rw&dt=rm&dt=ss&dt=t&ie=UTF-8&oe=UTF-8&otf=1&ssel=3&tsel=3&pc=1&kc=2"
      . "&tk=" . JS.("tk").(str), 1 )

   http.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded;charset=utf-8")
   http.SetRequestHeader("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:47.0) Gecko/20100101 Firefox/47.0")
   http.send("q=" . URIEncode(str))
   http.WaitForResponse(-1)
   Return http.responsetext
}

URIEncode(str, encoding := "UTF-8")  { ;*[my]
   VarSetCapacity(var, StrPut(str, encoding))
   StrPut(str, &var, encoding)

   While code := NumGet(Var, A_Index - 1, "UChar")  {
      bool := (code > 0x7F || code < 0x30 || code = 0x3D)
      UrlStr .= bool ? "%" . Format("{:02X}", code) : Chr(code)
   }
   Return UrlStr
}

GetJScript()
{
   script =
   (
      var TKK = ((function() {
        var a = 561666268;
        var b = 1526272306;
        return 406398 + '.' + (a + b);
      })());

      function b(a, b) {
        for (var d = 0; d < b.length - 2; d += 3) {
            var c = b.charAt(d + 2),
                c = "a" <= c ? c.charCodeAt(0) - 87 : Number(c),
                c = "+" == b.charAt(d + 1) ? a >>> c : a << c;
            a = "+" == b.charAt(d) ? a + c & 4294967295 : a ^ c
        }
        return a
      }

      function tk(a) {
          for (var e = TKK.split("."), h = Number(e[0]) || 0, g = [], d = 0, f = 0; f < a.length; f++) {
              var c = a.charCodeAt(f);
              128 > c ? g[d++] = c : (2048 > c ? g[d++] = c >> 6 | 192 : (55296 == (c & 64512) && f + 1 < a.length && 56320 == (a.charCodeAt(f + 1) & 64512) ?
              (c = 65536 + ((c & 1023) << 10) + (a.charCodeAt(++f) & 1023), g[d++] = c >> 18 | 240,
              g[d++] = c >> 12 & 63 | 128) : g[d++] = c >> 12 | 224, g[d++] = c >> 6 & 63 | 128), g[d++] = c & 63 | 128)
          }
          a = h;
          for (d = 0; d < g.length; d++) a += g[d], a = b(a, "+-a^+6");
          a = b(a, "+-3^+b+-f");
          a ^= Number(e[1]) || 0;
          0 > a && (a = (a & 2147483647) + 2147483648);
          a `%= 1E6;
          return a.toString() + "." + (a ^ h)
      }
   )
   Return script
}

GetJScripObject()  {
   static doc
   doc := ComObjCreate("htmlfile")
   doc.write("<meta http-equiv='X-UA-Compatible' content='IE=9'>")
   Return ObjBindMethod(doc.parentWindow, "eval")
}




