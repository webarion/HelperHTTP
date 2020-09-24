#include-once

;~ #include <dev.au3> ; библиотека разработчика

#Region User variables. Переменные пользователя

Global $igDEBUG_HelperHTTP = 0 ; Allows you to show additional information in the console. Позволяет показывать в консоли дополнительную информацию.
$igTimeout_HelperHTTP = 5000 ;  Request timeout. Таймаут запроса

#EndRegion User variables. Переменные пользователя


#Region Internal variables. Внутренние системные переменные
Global $igInit_HelperHTTP = 0, $hgTimer_HelperHTTP = 0, $sgResponse_Function_HelperHTTP = '', $ogObject_HelperHTTP
Global $ogHeaders_HelperHTTP = ObjCreate('Scripting.Dictionary')
$ogHeaders_HelperHTTP.Item('User-Agent') = 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4250.0 Iron Safari/537.36'
$ogHeaders_HelperHTTP.Item('Content-Type') = 'application/x-www-form-_EncodeURL_HelperHTTPd'
#EndRegion Internal variables. Внутренние системные переменные


#Region Пользовательские функции. User functions

; #ПОЛЬЗОВАТЕЛЬСКАЯ_ФУНКЦИЯ# ====================================================================================================
; Описание ...: Инициализирует библиотеку
; Параметры ..: $sMethod             - [необязательный]  По умолчанию ''.
; Возвращает .: None
; Примечание .:
;             :
; ===============================================================================================================================
Func _Init_HelperHTTP($sMethod = '')
	If $igInit_HelperHTTP Then Return 1
	Local $aMethod_HelperHTTP = ['Msxml2.XMLHTTP.6.0', 'Msxml2.XMLHTTP.5.0', 'Msxml2.XMLHTTP.4.0', 'Microsoft.XMLHTTP']
	If $sMethod Then
		$aMethod_HelperHTTP[0] = $sMethod
		ReDim $aMethod_HelperHTTP[1]
	EndIf
	For $i = 0 To UBound($aMethod_HelperHTTP) - 1
		$ogObject_HelperHTTP = ObjCreate($aMethod_HelperHTTP[$i])
		If Not @error Then ExitLoop
	Next

	If $i = UBound($aMethod_HelperHTTP) Then Return SetError(1, __Debug_HelperHTTP('Not find HTTP object', @ScriptLineNumber), 0)

	__Debug_HelperHTTP('+Ok Init HTTP object (' & $aMethod_HelperHTTP[$i] & ')', @ScriptLineNumber)
	Global $ogError_HelperHTTP = ObjEvent('AutoIt.Error', '__Debug_HelperHTTP')
	Return 1
EndFunc   ;==>_Init_HelperHTTP

Func _Response_Function_HelperHTTP($sRespFunc)
	If $sRespFunc Then $sgResponse_Function_HelperHTTP = $sRespFunc
EndFunc   ;==>_Response_Function_HelperHTTP

Func _AddHeader_HelperHTTP($sHeaderName, $sHeaderValue)
	$ogHeaders_HelperHTTP.Item($sHeaderName) = $sHeaderValue
EndFunc   ;==>_AddHeader_HelperHTTP

Func _DelHeader_HelperHTTP($sHeaderName)
	If $ogHeaders_HelperHTTP.Exists($sHeaderName) Then
		$ogHeaders_HelperHTTP.Remove($sHeaderName)
		Return 1
	EndIf
	Return 0
EndFunc   ;==>_DelHeader_HelperHTTP

Func _Request_HelperHTTP($sURL, $sMethod = 'GET', $sParams = '', $iAsync = 0, $sCallbackFunction = '', $sHeader = '')
	If Not IsObj($ogObject_HelperHTTP) Then _Init_HelperHTTP()
	If @error Then Return SetError(1, __Debug_HelperHTTP('No request object', @ScriptLineNumber - 1), 0)

	If Not $sMethod Or $sMethod = Default Then $sMethod = 'GET'
	$ogObject_HelperHTTP.Open($sMethod, $sURL, $iAsync)
	If @error Then Return SetError(2, __Debug_HelperHTTP('Failed to execute HTTP.Open', @ScriptLineNumber - 1), 0)

	For $sKey In $ogHeaders_HelperHTTP
		$ogObject_HelperHTTP.SetRequestHeader($sKey, $ogHeaders_HelperHTTP.Item($sKey))
	Next

	If $sHeader Then
		Local $aHeades = StringSplit($sHeader, @CRLF, 2)
		If UBound($aHeades) Then
			For $sHeader In $aHeades
				__WriteHeader_HelperHTTP($sHeader)
			Next
		Else
			__WriteHeader_HelperHTTP($sHeader)
		EndIf
	EndIf

	$hgTimer_HelperHTTP = TimerInit()

	If Not $sParams Or $sParams = Default Then
		$ogObject_HelperHTTP.Send()
	Else
		$ogObject_HelperHTTP.Send($sParams)
	EndIf
	If @error Then Return SetError(3, __Debug_HelperHTTP('Failed to execute HTTP.Send', @ScriptLineNumber), 0)

	If $sCallbackFunction Then $sgResponse_Function_HelperHTTP = $sCallbackFunction

	If $iAsync Then
		AdlibRegister('__GetResponse_HelperHTTP', 300)
		Return 1
	Else
		Return __GetResponse_HelperHTTP()
	EndIf

EndFunc   ;==>_Request_HelperHTTP


Func _isTimeoutHTTP()
	If $hgTimer_HelperHTTP And TimerDiff($hgTimer_HelperHTTP) > $igTimeout_HelperHTTP Then Return 1
	Return 0
EndFunc   ;==>_isTimeoutHTTP


Func _EncodeURL_HelperHTTP($sRawStr)
	Local $sUrl = "", $sAscCode
	For $i = 1 To StringLen($sRawStr)
		$sAscCode = Asc(StringMid($sRawStr, $i, 1))
		Select
			Case ($sAscCode >= 48 And $sAscCode <= 57) Or _
					($sAscCode >= 65 And $sAscCode <= 90) Or _
					($sAscCode >= 97 And $sAscCode <= 122)
				$sUrl = $sUrl & StringMid($sRawStr, $i, 1)
			Case $sAscCode = 32
				$sUrl = $sUrl & "+"
			Case Else
				$sUrl = $sUrl & "%" & Hex($sAscCode, 2)
		EndSelect
	Next
	Return $sUrl
EndFunc   ;==>_EncodeURL_HelperHTTP


Func DecodeURL_HelperHTTP($urlText)
	$urlText = StringReplace($urlText, "+", " ")
	Local $matches = StringRegExp($urlText, "\%([abcdefABCDEF0-9]{2})", 3)
	If Not @error Then
		For $match In $matches
			$urlText = StringReplace($urlText, "%" & $match, BinaryToString('0x' & $match))
		Next
	EndIf
	Return $urlText
EndFunc   ;==>DecodeURL_HelperHTTP


; #USER_FUNCTION# ===============================================================================================================
; Description ...: Determines if a URL is available
; Parameters ....: $sURL            - Url
;                  $iTimeout        - Waiting time. Default 1000
;                  $iNumberRequests - [optional] An integer value. Default is 1.
; Return values .: Response time
; ===============================================================================================================================
; #ПОЛЬЗОВАТЕЛЬСКАЯ_ФУНКЦИЯ# ====================================================================================================
; Описание ...: Определяет, доступен ли URL
; Параметры ..: $sURL            - URL адрес
;               $iTimeout        - Время ожидания. По умолчанию 1000
;               $iNumberRequests - Количество проверок. Позволяет получить максимальное время отклика. По умолчанию 1
; Возвращает .: Время отклика
; ===============================================================================================================================
Func _Ping_HelperHTTP($sURL, $iTimeout = 1000, $iNumberRequests = 1)
	Local $iPing, $iPingMax = 0
	For $i = 1 To $iNumberRequests
		$iPing = Ping(StringRegExpReplace($sURL, '.+?//(.+?)(?:/.*)?', '$1'), $iTimeout)
		If $iPingMax < $iPing Then $iPingMax = $iPing
		Sleep(100)
	Next
	If $iPingMax Then Return $iPingMax
	Return SetError(1, __Debug_HelperHTTP('Ping error for URL: ' & $sURL, @ScriptLineNumber), 0) ; ошибка пинга для URL
EndFunc   ;==>_Ping_HelperHTTP

Func _ParsURL_HelperHTTP($sURL, $iUrlCode = 1)
	Local $aComplete[4] = ['http://', '', '', '']
	Local $aSplitURL = StringSplit($sURL, '?', 2)
	If UBound($aSplitURL) > 0 Then
		Local $aParsURL = StringRegExp($aSplitURL[0], '(^https?://|^)(.*?)(/.*|$)$', 1)
		If IsArray($aParsURL) Then
			If $aParsURL[0] Then $aComplete[0] = $aParsURL[0]
			If $aParsURL[1] Then $aComplete[1] = $aParsURL[1]
			If $aParsURL[2] Then $aComplete[2] = $aParsURL[2]
		EndIf
	EndIf
	If UBound($aSplitURL) = 2 Then
		$aComplete[3] = $aSplitURL[1]
		If $iUrlCode Then _EncodeURL_HelperHTTP($aComplete[3])
	EndIf
	Return $aComplete
EndFunc   ;==>_ParsURL_HelperHTTP


; #USER_FUNCTION# ===============================================================================================================
; Description ...: Returns a randomly generated character string
; Parameters ....: $iLength - The length of the generated string
; ===============================================================================================================================
; #ПОЛЬЗОВАТЕЛЬСКАЯ_ФУНКЦИЯ# ====================================================================================================
; Описание ...: Возвращает случайно сгенерированную строку символов
; Параметры ..: $iLength - Длина генерируемой строки. По умолчанию 64
; ===============================================================================================================================
Func _GenSessinoKey_HelperHTTP($iLength = 64)
	Local $sResult
	Local $sSequence = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789'
	Local $aSplit = StringSplit($sSequence, "", 2)
	For $i = 1 To $iLength
		$sResult &= $aSplit[Random(0, UBound($aSplit) - 1, 1)]
	Next
	Return $sResult
EndFunc   ;==>_GenSessinoKey_HelperHTTP

#EndRegion Пользовательские функции. User functions


#Region Внутренние системные функции. Internal functions

Func __GetResponse_HelperHTTP()
	Local $sRespData_HelperHTTP = ''
	If $ogObject_HelperHTTP.readyState = 4 Then
		Local $sRespData_HelperHTTP = $ogObject_HelperHTTP.ResponseText

		If Not @error Then
			AdlibUnRegister('__GetResponse_HelperHTTP')
			If $sgResponse_Function_HelperHTTP Then __Callback_HelperHTTP($sRespData_HelperHTTP)
			Return $sRespData_HelperHTTP
		EndIf
	ElseIf $hgTimer_HelperHTTP And TimerDiff($hgTimer_HelperHTTP) > $igTimeout_HelperHTTP Then
		AdlibUnRegister('__GetResponse_HelperHTTP')
		$hgTimer_HelperHTTP = 0
		If $sgResponse_Function_HelperHTTP Then __Callback_HelperHTTP('', 1)
	EndIf
	Return $sRespData_HelperHTTP
EndFunc   ;==>__GetResponse_HelperHTTP

Func __Callback_HelperHTTP($sData, $iTimeoutExceeded = 0)
	If Not $sgResponse_Function_HelperHTTP Then
		__Debug_HelperHTTP('Callback function not registered', @ScriptLineNumber) ; Функция обратного вызова не зарегистрирована
		Return SetError(1, 0, 0)
	EndIf
	If $iTimeoutExceeded Then __Debug_HelperHTTP('Timeout exceeded. Increase the waiting time for a response in $ igTimeout_HelperHTTP', @ScriptLineNumber) ; Превышен тайм-аут. Увеличьте время ожидания ответа в $ igTimeout_HelperHTTP
	Call($sgResponse_Function_HelperHTTP, $sData)
	If @error = 0xDEAD And @extended = 0xBEEF Then
		Call($sgResponse_Function_HelperHTTP, $sData, $iTimeoutExceeded)
		If @error = 0xDEAD And @extended = 0xBEEF Then __Debug_HelperHTTP('Failed to Call function ' & $sgResponse_Function_HelperHTTP & '. The wrong number of parameters may be specified', @ScriptLineNumber) ; Не удалось вызвать функцию. Может быть указано неверное количество параметров
	EndIf
EndFunc   ;==>__Callback_HelperHTTP

Func __WriteHeader_HelperHTTP($sHeader)
	If Not IsObj($ogObject_HelperHTTP) Then Return SetError(1, __Debug_HelperHTTP('Not object $ogObject_HelperHTTP', @ScriptLineNumber), 0) ; Нет объекта $ogObject_HelperHTTP
	Local $aHeader = StringSplit($sHeader, ':', 2)
	If UBound($aHeader) = 2 Then
		$ogObject_HelperHTTP.SetRequestHeader($aHeader[0], $aHeader[1])
		If @error Then Return SetError(2, __Debug_HelperHTTP('Not add header: ' & $aHeader[0] & '|' & $aHeader[1], @ScriptLineNumber), 0) ; Не удалось добавить заголовок
		Return 1
	Else
		Return SetError(2, __Debug_HelperHTTP('Incorrect header format: ' & $sHeader, @ScriptLineNumber), 0) ; Неправильный формат заголовка
	EndIf
	Return 0
EndFunc   ;==>__WriteHeader_HelperHTTP


Func __Debug_HelperHTTP($ogError_HelperHTTP, $iScriptLine = '')
	If IsString($ogError_HelperHTTP) Then
		Local $sSign = StringLeft($ogError_HelperHTTP, 1)
		Local $iSign = StringInStr('>!-+', $sSign) ? $sSign : ''
		$ogError_HelperHTTP = $iSign ? StringRight($ogError_HelperHTTP, StringLen($ogError_HelperHTTP) - 1) : $ogError_HelperHTTP
		Local $sMsg = $iSign & @ScriptName & ' MsgLine(' & $iScriptLine & ') : ==> ' & $ogError_HelperHTTP

		If $igDEBUG_HelperHTTP Then ConsoleWrite($sMsg & @CRLF)
		Return SetError(1, 0, 0)
	ElseIf Not IsObj($ogError_HelperHTTP) Then
		Return SetError(2, 0, 0)
	EndIf
	Local $iErrNumber = Hex($ogError_HelperHTTP.number, 8)
	If $igDEBUG_HelperHTTP Then
		Local $sWindescription = $ogError_HelperHTTP.windescription
		Local $sDescription = $ogError_HelperHTTP.description
		Local $sSource = $ogError_HelperHTTP.source
		Local $sHelpfile = $ogError_HelperHTTP.helpfile
		Local $sHelpcontext = $ogError_HelperHTTP.helpcontext
		Local $sLastdllerror = $ogError_HelperHTTP.lastdllerror
		Local $sRetcode = "0x" & Hex($ogError_HelperHTTP.retcode)
		ConsoleWrite(@ScriptName & " (" & $ogError_HelperHTTP.scriptline & ") : ==> COM Error!" & @CRLF) ; Получена ошибка COM
		ConsoleWrite("Number is: " & @TAB & @TAB & "0x" & $iErrNumber & @CRLF) ; Номер ошибки
		If $sWindescription Then ConsoleWrite("Windescription:" & @TAB & $sWindescription & @CRLF) ; Системное описание ошибки
		If $sDescription Then ConsoleWrite("Description is: " & @TAB & $sDescription & @CRLF) ; Описание
		If $sSource Then ConsoleWrite("Source is: " & @TAB & @TAB & $sSource & @CRLF) ; Источник
		If $sHelpfile Then ConsoleWrite("Helpfile is: " & @TAB & $sHelpfile & @CRLF) ; Файл справки
		If $sHelpcontext Then ConsoleWrite("Helpcontext is: " & @TAB & $sHelpcontext & @CRLF) ; Контекст помощи
		If $sLastdllerror Then ConsoleWrite("Lastdllerror is: " & @TAB & $sLastdllerror & @CRLF) ; Последняя ошибка dll
		If $sRetcode Then ConsoleWrite("Retcode is: " & @TAB & $sRetcode & @CRLF & @CRLF) ; Возвращённый код
	EndIf
	Return SetError(3, $iErrNumber, 0)
EndFunc   ;==>__Debug_HelperHTTP

#EndRegion Внутренние системные функции. Internal functions


