#include-once

;     _   _          _                         _   _   _____   _____   ____
;    | | | |   ___  | |  _ __     ___   _ __  | | | | |_   _| |_   _| |  _ \  ©Webarion
;    | |_| |  / _ \ | | | '_ \   / _ \ | '__| | |_| |   | |     | |   | |_) |
;    |  _  | |  __/ | | | |_) | |  __/ | |    |  _  |   | |     | |   |  __/
;    |_| |_|  \___| |_| | .__/   \___| |_|    |_| |_|   |_|     |_|   |_|
;                       |_|

; Note: English translation by Google Translate

; # ABOUT THE LIBRARY # =========================================================================================================
; Name .............: HelperHTTP
; Current version ..: 1.0.0
; AutoIt Version ...: 3.3.14.5
; Description ......: Synchronous and asynchronous HTTP request helper
; Author ...........: Webarion
; Links: ...........: http://webarion.ru, http://f91974ik.bget.ru
; Link library .....: https://github.com/webarion/HelperHTTP
; ===============================================================================================================================

#CS Version history:
	v1.0.0
	First published version
#CE History

; # О БИБЛИОТЕКЕ # ==============================================================================================================
; Название .........: HelperHTTP
; Текущая версия ...: 1.0.0
; AutoIt Версия ....: 3.3.14.5
; Описание .........: Помощник синхронных и асинхронных HTTP запросов
; Автор ............: Webarion
; Ссылки: ..........: http://webarion.ru, http://f91974ik.bget.ru
; Ссылка библиотеки : https://github.com/webarion/HelperHTTP
; ===============================================================================================================================

#CS История версий:
	v1.0.0
	Первая опубликованная версия
#CE History

#CS Brief description of user functions. Краткое описание пользовательских функций

; _Init_HelperHTTP              - Initializes the library | Инициализирует библиотеку
; _Response_Function_HelperHTTP - Registers the function that will receive the response | Регистрирует функцию, в которую будет приходить ответ
; _RequestHeader_HelperHTTP     - Adds a header to the request | Добавляет временный или постоянный заголовок в запрос
; _DelHeader_HelperHTTP         - Removes the header from the request | Удаляет заголовок из запроса
; _Request_HelperHTTP           - Sends a request and receives a response | Отправляет запрос и получает ответ
; _isTimeoutHTTP                - Lets you know if the asynchronous request timeout is exceeded | Позволяет узнать не превышен ли таймаут асинхронного запроса
; _EncodeURL_HelperHTTP         - Returns a string encoded according to the URL format | Возвращает строку, закодированную соответственно URL формату
; _DecodeURL_HelperHTTP         - Returns the decoded URL string | Возвращает декодированную URL-строку
; _Ping_HelperHTTP              - Determines if a URL is available | Определяет, доступен ли URL
; _ParsURL_HelperHTTP           - Parses the URL string | Парсит URL строку
; _GenSessinoKey_HelperHTTP     - Returns a randomly generated character string | Возвращает случайно сгенерированную строку символов

#CE

;~ #include <dev.au3>; Developer library. Библиотека разработчика

#Region User variables. Переменные пользователя
Global $igDEBUG_HelperHTTP = 0 ; Allows you to show additional information in the console. Позволяет показывать в консоли дополнительную информацию.
$igTimeout_HelperHTTP = 5000 ;  Request timeout. Таймаут запроса
#EndRegion User variables. Переменные пользователя

#Region Internal variables. Внутренние системные переменные
Global $igInit_HelperHTTP = 0, $hgTimer_HelperHTTP = 0, $sgResponse_Function_HelperHTTP = '', $ogObject_HelperHTTP
Global $ogRequestHeaders_HelperHTTP = ObjCreate('Scripting.Dictionary')
_RequestHeader_HelperHTTP('User-Agent', 'Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:82.0) Gecko/20100101 Firefox/82.0')
_RequestHeader_HelperHTTP('Content-Type', 'application/x-www-form-_EncodeURL_HelperHTTPd')
#EndRegion Internal variables. Внутренние системные переменные


#Region User functions. Пользовательские функции

; #USER FUNCTION# ===============================================================================================================
; Description .: Initializes the library
; Parameters ..: $sMethod - initialize the HTTP object. By default, it is detected automatically
; Returns .....: 1-success. 0 - in case of an error, if the object is not found in the system and @error is set.
;                               Shows an additional message if $igDEBUG_HelperHTTP is enabled
; ===============================================================================================================================
; #ПОЛЬЗОВАТЕЛЬСКАЯ ФУНКЦИЯ# ====================================================================================================
; Описание ....: Инициализирует библиотеку
; Параметры ...: $sMethod - инициализируемый HTTP объект. По умолчанию определяется автоматически
; Возвращает ..: 1 - успех. 0 - в случае ошибки, если объект не найден в системе и устанавливается @error.
;                               Показывает дополнительное сообщение если включён $igDEBUG_HelperHTTP
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
	If $i = UBound($aMethod_HelperHTTP) Then Return SetError(1, __Debug_HelperHTTP('!' & __TrHH('Not find HTTP object', 'Ошибка инициализации HTTP объекта'), @ScriptLineNumber), 0)
	__Debug_HelperHTTP('+' & __TrHH('Ok initialization object', 'Успешная инициализация объекта') & ' (' & $aMethod_HelperHTTP[$i] & ')', @ScriptLineNumber)
	Global $ogError_HelperHTTP = ObjEvent('AutoIt.Error', '__Debug_HelperHTTP')
	Return 1
EndFunc   ;==>_Init_HelperHTTP


; #USER FUNCTION# ===============================================================================================================
; Description ..: Registers the function that will receive the response
; Parameters ...: $sRespFunc-string with the function name
; ===============================================================================================================================
; #ПОЛЬЗОВАТЕЛЬСКАЯ ФУНКЦИЯ# ====================================================================================================
; Описание ....: Регистрирует функцию, в которую будет приходить ответ
; Параметры ...: $sRespFunc - строка с названием функции
; ===============================================================================================================================
Func _Response_Function_HelperHTTP($sRespFunc)
	If $sRespFunc Then $sgResponse_Function_HelperHTTP = $sRespFunc
EndFunc   ;==>_Response_Function_HelperHTTP


; #USER FUNCTION# ===============================================================================================================
; Description .: Adds a header to the request
; Parameters ..: $sHeaderName         - header name
;                $sHeaderValue        - header value
;                $iOnlyForNextRequest - 1 - the header will be set only for the next request
;                                       0 - the header will be created in all subsequent requests
; ===============================================================================================================================
; #ПОЛЬЗОВАТЕЛЬСКАЯ ФУНКЦИЯ# ====================================================================================================
; Описание ....: Добавляет временный или постоянный заголовок в запрос
; Параметры ...: $sHeaderName      - название заголовка
;                $sHeaderValue     - значение заголовка
;                $iOnlyForNextRequest - 0 - заголовок будет установлен во всех запросах [По умолчанию]
;                                       1 - заголовок будет установлен только в следующем запросе
;                                       n - число означающее, через какое количество запросов, заголовок будет стёрт
; ===============================================================================================================================
Func _RequestHeader_HelperHTTP($sHeaderName, $sHeaderValue, $iOnlyForNextRequest = 0)
	Local $aHeader = [$sHeaderValue, $iOnlyForNextRequest]
	$ogRequestHeaders_HelperHTTP.Item($sHeaderName) = $aHeader
EndFunc   ;==>_RequestHeader_HelperHTTP


; #USER FUNCTION# ===============================================================================================================
; Description .: Removes the header from the request
; Parameters ..: $sHeaderName - name of the header to delete
; ===============================================================================================================================
; #ПОЛЬЗОВАТЕЛЬСКАЯ ФУНКЦИЯ# ====================================================================================================
; Описание ....: Удаляет заголовок из запроса
; Параметры ...: $sHeaderName - название удаляемого заголовка
; ===============================================================================================================================
Func _DelHeader_HelperHTTP($sHeaderName)
	If $ogRequestHeaders_HelperHTTP.Exists($sHeaderName) Then
		$ogRequestHeaders_HelperHTTP.Remove($sHeaderName)
		Return 1
	EndIf
	Return 0
EndFunc   ;==>_DelHeader_HelperHTTP


; #USER FUNCTION# ===============================================================================================================
; Description .: Sends a request and receives a response
; Parameters ..: $sURL-request address
;                $sMethod           - method. By default, 'GET'
;                $sParams           - string of parameters in the request
;                $iAsync            - if 1, the request is asynchronous
;                $sCallbackFunction - Name of the function to which the response will be sent. By default, the response is via Return
;                $sHeader           - Header or headers. It can consist of several lines
; Returns .....: In the knock of a synchronous request and if no return function is specified, returns the response
; ===============================================================================================================================
; #ПОЛЬЗОВАТЕЛЬСКАЯ ФУНКЦИЯ# ====================================================================================================
; Описание ....: Отправляет запрос и получает ответ
; Параметры ...: $sURL              - адрес запроса
;                $sMethod           - метод. По умолчанию 'GET'
;                $sParams           - строка параметров в запросе: a=1&b=2&c=3
;                $iAsync            - Если 1, запрос асинхронный
;                $sCallbackFunction - Название функции, в которую будет отправлен ответ. По умолчанию ответ через Return
;                $sHeader           - Заголовок или заголовки. Может состоять из нескольких строк
; Возвращает ..: Возвращает ответ, в стучае синхронного запроса и если не указана функция возврата.
; ===============================================================================================================================
Func _Request_HelperHTTP($sUrl, $sMethod = 'GET', $sParams = '', $iAsync = 0, $sCallbackFunction = '', $sHeader = '')
	If Not IsObj($ogObject_HelperHTTP) Then _Init_HelperHTTP()
	If @error Then Return SetError(1, __Debug_HelperHTTP(__TrHH('No request object', 'Нет объекта запроса'), @ScriptLineNumber - 1), 0)
	If Not $sMethod Or $sMethod = Default Then $sMethod = 'GET'
	$ogObject_HelperHTTP.Open($sMethod, $sUrl, $iAsync)
	If @error Then Return SetError(2, __Debug_HelperHTTP(__TrHH('Failed to execute', 'Не удалось выполнить') & ' HTTP.Open', @ScriptLineNumber - 1), 0)
	; добавляем ранее указанные заголовки
	For $aKeyHeader In $ogRequestHeaders_HelperHTTP
		Local $aHeader = $ogRequestHeaders_HelperHTTP.Item($aKeyHeader)
		If UBound($aHeader) Then
			$ogObject_HelperHTTP.SetRequestHeader($aKeyHeader, $aHeader[0])
			If UBound($aHeader) = 2 Then
				If $aHeader[1] = 1 Then
					_DelHeader_HelperHTTP($aKeyHeader)
				ElseIf $aHeader[1] > 1 Then
					_RequestHeader_HelperHTTP($aKeyHeader, $aHeader[0], $aHeader[1] - 1)
				EndIf
			EndIf
		EndIf
	Next
	;
	If $sHeader Then ; если есть заголовок текущего запроса
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
	If @error Then Return SetError(3, __Debug_HelperHTTP(__TrHH('Failed to execute', 'Не удалось выполнить') & ' HTTP.Send', @ScriptLineNumber), 0)
	If $sCallbackFunction Then $sgResponse_Function_HelperHTTP = $sCallbackFunction
	If $iAsync Then
		AdlibRegister('__GetResponse_HelperHTTP', 300)
		Return 1
	Else
		Return __GetResponse_HelperHTTP()
	EndIf
EndFunc   ;==>_Request_HelperHTTP


; #USER FUNCTION# ===============================================================================================================
; Description .: Lets you know if the asynchronous request timeout is exceeded
; Returns .....: 1 - if the timeout is exceeded, 0 - if not
; ===============================================================================================================================
; #ПОЛЬЗОВАТЕЛЬСКАЯ ФУНКЦИЯ# ====================================================================================================
; Описание ....: Позволяет узнать не превышен ли таймаут асинхронного запроса
; Возвращает ..: 1 - если таймаут превышен, 0 - если нет
; ===============================================================================================================================
Func _isTimeoutHTTP()
	If $hgTimer_HelperHTTP And TimerDiff($hgTimer_HelperHTTP) > $igTimeout_HelperHTTP Then Return 1
	Return 0
EndFunc   ;==>_isTimeoutHTTP


; #USER FUNCTION# ===============================================================================================================
; Description .: Returns a string encoded according to the URL format
; Parameters ..: String to encode
; ===============================================================================================================================
; #ПОЛЬЗОВАТЕЛЬСКАЯ ФУНКЦИЯ# ====================================================================================================
; Описание ....: Возвращает строку, закодированную соответственно URL формату
; Параметры ...: $sRawStr - Строка для кодирования
; ===============================================================================================================================
Func _EncodeURL_HelperHTTP(ByRef $sRawStr)
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
	$sRawStr = $sUrl
	Return $sRawStr
EndFunc   ;==>_EncodeURL_HelperHTTP


; #USER FUNCTION# ===============================================================================================================
; Description .: Returns the decoded URL string
; Parameters ..: $urlText-String to decode
; ===============================================================================================================================
; #ПОЛЬЗОВАТЕЛЬСКАЯ ФУНКЦИЯ# ====================================================================================================
; Описание ....: Возвращает декодированную URL-строку
; Параметры ...: $urlText - Строка для декодирования
; ===============================================================================================================================
Func _DecodeURL_HelperHTTP(ByRef $urlText)
	$urlText = StringReplace($urlText, "+", " ")
	Local $matches = StringRegExp($urlText, "\%([abcdefABCDEF0-9]{2})", 3)
	If Not @error Then
		For $match In $matches
			$urlText = StringReplace($urlText, "%" & $match, BinaryToString('0x' & $match))
		Next
	EndIf
	Return $urlText
EndFunc   ;==>_DecodeURL_HelperHTTP


; #USER_FUNCTION# ===============================================================================================================
; Description ...: Determines if a URL is available
; Parameters ....: $sURL            - Url
;                  $iTimeout        - Timeout. Default 1000
;                  $iNumberRequests - Number of inspections. Allows you to get the maximum response time. The default is 1
; Return values .: Response time
; ===============================================================================================================================
; #ПОЛЬЗОВАТЕЛЬСКАЯ_ФУНКЦИЯ# ====================================================================================================
; Описание ...: Определяет, доступен ли URL
; Параметры ..: $sURL            - URL адрес
;               $iTimeout        - Время ожидания. По умолчанию 1000
;               $iNumberRequests - Количество проверок. Позволяет получить максимальное время отклика. По умолчанию 1
; Возвращает .: Время отклика
; ===============================================================================================================================
Func _Ping_HelperHTTP($sUrl, $iTimeout = 1000, $iNumberRequests = 1)
	Local $iPing, $iPingMax = 0
	For $i = 1 To $iNumberRequests
		$iPing = Ping(StringRegExpReplace($sUrl, '.+?//(.+?)(?:/.*)?', '$1'), $iTimeout)
		If $iPingMax < $iPing Then $iPingMax = $iPing
		Sleep(100)
	Next
	If $iPingMax Then Return $iPingMax
	Return SetError(1, __Debug_HelperHTTP(__TrHH('Ping error for URL', 'Ошибка пинга для URL') & ': ' & $sUrl, @ScriptLineNumber), 0)
EndFunc   ;==>_Ping_HelperHTTP


; #USER FUNCTION# ===============================================================================================================
; Description .: Parses the URL string
; Parameters ..: $sUrl - URL address
;                $iUrlCode - specifies how to convert the URL parameter string. By default, 2
;                            0 - leave as is
;                            1 - parameters will be encoded in URL format
;                            2 - parameters will be decoded in from URL format
; Returns .....: Array ['http://', 'domain', 'address', 'parameters string']
; ===============================================================================================================================
; #ПОЛЬЗОВАТЕЛЬСКАЯ ФУНКЦИЯ# ====================================================================================================
; Описание ....: Парсит URL строку
; Параметры ...: $sUrl     - URL адрес
;                $iUrlCode - указывает на способ преобразования строки параметров URL. По умолчанию 2
;                            0 - оставить как есть
;                            1 - параметры будут закодированы в URL формат
;                            2 - параметры будут декодированы в из URL формата
; Возвращает ..: Массив ['http://', 'домен', 'адрес', 'строка параметров']
; ===============================================================================================================================
Func _ParsURL_HelperHTTP($sUrl, $iUrlCode = 2)
	Local $aComplete[4] = ['http://', '', '', '']
	Local $aSplitURL = StringSplit($sUrl, '?', 2)
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
		Switch $iUrlCode
			Case 1
				_EncodeURL_HelperHTTP($aComplete[3])
			Case 2
				_DecodeURL_HelperHTTP($aComplete[3])
		EndSwitch
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

#EndRegion User functions. Пользовательские функции


#Region Internal functions. Внутренние системные функции

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
		__Debug_HelperHTTP(__TrHH('Callback function not registered', 'Функция обратного вызова не зарегистрирована'), @ScriptLineNumber)
		Return SetError(1, 0, 0)
	EndIf
	If $iTimeoutExceeded Then __Debug_HelperHTTP(__TrHH('Timeout exceeded. Increase the waiting time for a response in', 'Превышен тайм-аут. Увеличьте время ожидания ответа в') & ' $igTimeout_HelperHTTP', @ScriptLineNumber)
	Call($sgResponse_Function_HelperHTTP, $sData)
	If @error = 0xDEAD And @extended = 0xBEEF Then
		Call($sgResponse_Function_HelperHTTP, $sData, $iTimeoutExceeded)
		If @error = 0xDEAD And @extended = 0xBEEF Then __Debug_HelperHTTP(__TrHH('Failed to Call function', 'Не удалось вызвать функцию') & ' ' & $sgResponse_Function_HelperHTTP & '. ' & __TrHH('The wrong number of parameters may be specified', 'Может быть указано неверное количество параметров'), @ScriptLineNumber)
	EndIf
EndFunc   ;==>__Callback_HelperHTTP


Func __WriteHeader_HelperHTTP($sHeader)
	If Not IsObj($ogObject_HelperHTTP) Then Return SetError(1, __Debug_HelperHTTP(__TrHH('Not object', 'Нет объекта') & ' $ogObject_HelperHTTP', @ScriptLineNumber), 0)
	Local $aHeader = StringSplit($sHeader, ':', 2)
	If UBound($aHeader) = 2 Then
		$ogObject_HelperHTTP.SetRequestHeader($aHeader[0], $aHeader[1])
		If @error Then Return SetError(2, __Debug_HelperHTTP(__TrHH('Not add header', 'Не удалось добавить заголовок') & ': ' & $aHeader[0] & '|' & $aHeader[1], @ScriptLineNumber), 0)
		Return 1
	Else
		Return SetError(2, __Debug_HelperHTTP(__TrHH('Incorrect header format', 'Неправильный формат заголовка') & ': ' & $sHeader, @ScriptLineNumber), 0)
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
		ConsoleWrite(@ScriptName & " (" & $ogError_HelperHTTP.scriptline & ") : ==> COM " & __TrHH("Error", "Ошибка") & "!" & @CRLF) ; Получена ошибка COM
		ConsoleWrite(__TrHH("Number is", "Номер ошибки") & ": " & @TAB & @TAB & "0x" & $iErrNumber & @CRLF)
		If $sWindescription Then ConsoleWrite(__TrHH("Windescription", "Системное описание ошибки") & ":" & @TAB & $sWindescription & @CRLF)
		If $sDescription Then ConsoleWrite(__TrHH("Description is", "Описание") & ": " & @TAB & $sDescription & @CRLF)
		If $sSource Then ConsoleWrite(__TrHH("Source is", "Источник") & ": " & @TAB & @TAB & $sSource & @CRLF)
		If $sHelpfile Then ConsoleWrite(__TrHH("Helpfile is", "Файл справки") & ": " & @TAB & $sHelpfile & @CRLF)
		If $sHelpcontext Then ConsoleWrite(__TrHH("Helpcontext is", "Контекст помощи") & ": " & @TAB & $sHelpcontext & @CRLF)
		If $sLastdllerror Then ConsoleWrite(__TrHH("Lastdllerror is", "Последняя ошибка dll") & ": " & @TAB & $sLastdllerror & @CRLF)
		If $sRetcode Then ConsoleWrite(__TrHH("Retcode is", "Возвращённый код") & ": " & @TAB & $sRetcode & @CRLF & @CRLF)
	EndIf
	Return SetError(3, $iErrNumber, 0)
EndFunc   ;==>__Debug_HelperHTTP

; Translator. Переводчик
Func __TrHH($sEng_VP, $sRus_VP)
	Return @OSLang = 419 ? $sRus_VP : $sEng_VP
EndFunc   ;==>__TrHH

#EndRegion Internal functions. Внутренние системные функции



