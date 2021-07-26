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
