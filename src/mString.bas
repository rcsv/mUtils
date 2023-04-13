'
' module name : mString.bas
' version     : 1.0.1
' author      : rcsvpg@outlook.com
'
Option Explicit
'
' # Module String
' -----------------------------------------------------------------------
'
' The MIT License (MIT)
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in
' all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
' THE SOFTWARE.
'
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
' use: =ConcatIfs(", ", A1:C3, F3:F5)
Public Function CONCATIFS(glue_str As String, ParamArray joinRanges() As Variant) As String
	
	Dim index As Integer
	Dim joinRange As Variant
	
	' concatenate all ParamArray ranges
	For Each joinRange In JoinRanges
		' concatenate all cell data of range
		For index = 1 To joinRange.Count
			
			' skip when no data
			If joinRange(index) > "" Then
				' prepare : add glue string before next string (without first)
				If Len(CONCATIFS) <> 0 Then
					CONCATIFS = CONCATIFS & glue_str
				End If
				CONCATIFS = CONCATIFS & joinRange(index)
			End If
		Next index
	Next
					
End Function
