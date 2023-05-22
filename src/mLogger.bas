Option Explicit
' mLogger.bas
'
' This module is for logging, and output log to sheet named "Log" in
' excel file. This module is compliant with RFC5424 and PSR-3 Logger
' Interface, has methods for each log level.
'
' @see RFC5424
' @see PSR-3 Logger Interface
'
Private Const LOG_SHEET_NAME As String = "Log"

' Output Log into "Log" sheet:
' @param level log level
' @param message
' @param context
' @return void
Public Sub log(level As LogLevel, message As String, Optional context As Dictionary = Nothing)
	Dim logSheet As Worksheet
	Set _
		logSheet = ThisWorkbook.Sheets(LOG_SHEET_NAME)
	Dim lastRow As Long
	lastRow = logSheet.Cells(Rows.Count, 1).End(xlUp).Row
	logSheet.Cells(lastRow + 1, 1).Value = level
	logSheet.Cells(lastRow + 1, 2).Value = message

	' output context if exists
	If Not context Is Nothing Then
    	Dim key As Variant
    	Dim i As Long
    	For Each key In context.Keys
			logSheet.Cells(lastRow + 1, i).Value = key
			logSheet.Cells(lastRow + 1, i + 1).Value = context(key)
		Next key
	End If
End Sub

' define log level
Public Enum LogLevel
	EMERGENCY = 0 
	ALERT = 1
	CRITICAL = 2
	ERROR = 3
	WARNING = 4
	NOTICE = 5
	INFO = 6
	DEBUGGING = 7 ' DEBUG is not available cause keyword is reserved.
	TRACE = 8
End Enum

' PSR-3 Compliant Interface
' @param message
' @param context

' Emergency
Public Sub emergency(message As String, Optional context As Dictionary = Nothing)
	log EMERGENCY, message, context
End Sub

' Alert
Public Sub alert(message As String, Optional context As Dictionary = Nothing)
	log ALERT, message, context
End Sub

' Critical
Public Sub critical(message As String, Optional context As Dictionary = Nothing)
	log CRITICAL, message, context
End Sub

' Error
Public Sub error(message As String, Optional context As Dictionary = Nothing)
	log ERROR, message, context
End Sub

' Warning
Public Sub warning(message As String, Optional context As Dictionary = Nothing)
	log WARNING, message, context
End Sub

' Notice
Public Sub notice(message As String, Optional context As Dictionary = Nothing)
	log NOTICE, message, context
End Sub

' Information
Public Sub info(message As String, Optional context As Dictionary = Nothing)
	log INFO, message, context
End Sub

' debugging
Public Sub debugging(message As String, Optional context As Dictionary = Nothing)
	log DEBUGGING, message, context
End Sub

' tracing
Public Sub trace(message As String, Optional context As Dictionary = Nothing)
	log TRACE, message, context
End Sub

' End of source
