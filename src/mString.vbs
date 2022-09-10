'
' # Module String
' -----------------------------------------------------------------------

Option Explicit
' Publilc NotInherits Class String

' hSeparate
' SPLIT Horizontally
Public Function hSplit(ByRef str As String, Optional ByRef sep As String = " ") As Variant
	Dim v As Variant
	v = Split(str, sep)
	hSeparate = v
End Function

' ConcatIf
' 
' ALTER TECTJOIN function for before Excel 2019
' use: =ConcatIf(", ", A1:C3, F3:F5)
Public Function CONCATIF(glue_str As String ParamArray joinRanges() As Variant) As String
	
	Dim index As Integer
	Dim joinRange As Variant
	
	' concatenate all ParamArray ranges
	For Each joinRange In JoinRanges
		' concatenate all cell data of range
		For index = 1 To joinRange.Count
			
			' skip when no data
			If joinRange(index) > "" Then
				' prepare : add glue string before next string (without first)
				If Len(CONCATIF) <> 0 Then
					CONCATIF = CONCATIF & glue_str
				End If
				CONCATIF = CONCATIF & joinRange(index)
			End If
		Next index
	Next
					
End Function
