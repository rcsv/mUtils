'
' # Picker
' -----------------------------------------------------------------------

Option Explicit
' Public NotInherits Class Pickers

' フォルダー/ファイルピッカーに登場するボタンの文字列
Private Const strBUTTON_NAME As String = "選択"

' タイトル文字列
Private Const strTITLE_FOLDER As String = "フォルダを"
Private Const strTITLE_FILE As String = "ファイルを"

'
' getFolderPath
' -----------------------------------------------------------------------
' フォルダ選択ダイアログからフォルダ名を取得する
' キャンセルするとから文字列になるのが玉に瑕
'
Public Function getFolderPath() As String
  
  ' 初期値を置いとく
  getFolderPath = ""
  
  With Application.FileDialog(msoFileDialogFolderPicker)
    .AllowMultiSelect = False
    .ButtonName = strBUTTON_NAME
    .Title = strTITLE_FOLDER & strBUTTON_NAME
    
    ' フォルダ名を取得したら戻す
    If .Show = True Then
      getFolderPath = .SelectedItems(1)
    End If
  End With

End Function

'
' getFilePath
' -----------------------------------------------------------------------
' ファイル選択ダイアログからファイル名を取得する
' キャンセルするとから文字列が戻る
'
' initialDirectory String Optional 初期ディレクトリ
' flgWithPath Boolean フルパスで戻すかどうかTrueはフルパスFalse=ファイル名のみ
Public Function getFilePath(Optional initialDirectory As String = "", _
  Optional flgWithPath As Boolean = False) As String
  
  ' 初期値を作っておく
  getFilePath = ""
  
  With Application.FileDialog(msoFileDialogFilePicker)
    .InitialFileName = initialDirectory
    .AllowMultiSelect = False
    .ButtonName = strBUTTON_NAME
    .Title = strTITLE_FILE & strBUTTON_NAME
    
    ' ファイル名を取得したら戻す
    If .Show = True Then
      ' さらにファイル名をフルパスで返すか、ファイル名だけで返すか
      If flgWithPath Then
        getFilePath = Dir(.SelectedItems(1))
      Else
        getFilePath = .SelectedItems(1)
      End If
    End If
  End With
End Function

