'
' module name : mPickers.bas
' version     : 1.0.2
' author      : tomohiro awane <Tomohiro.AWANE@aisin.co.jp>
'
Option Explicit
'
' Windows Picker Utility for File and Folder
' ------------------------------------------------------------------------
'
' This script is used to pick a file or folder from Windows Explorer.
'
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
' Version 1.0.2 - 2023-04-13
' Version 1.0.1 - 2022-03-01
' Version 1.0.0 - 2021-02-01
'
' Description:
' ------------------------------------------------------------------------
' Windowsのファイル・フォルダ選択ダイアログを表示し選択したファイル、フォ
' ルダのパスを返す。ボタンから直接は呼ばれず、一旦Public Sub Main()を経由
' して、パラメータの初期化後に呼ばれる
'
' アイデア：
' Private Sub Worksheet_SelectionChange(ByVal Target As Range)で、
' 下記のメソッドを呼び出す
'
' C2, C4, C6 をターゲットにした際のPicker
Public Sub Main_C2()
    Dim targetCell As String
    ' パラメータの初期化
    targetCell = "C2" ' パスを入力するセル(targetDir)
    Call PickerFolder(targetCell)
End Sub

Public Sub Main_C4()
    Dim targetCell As String
    ' パラメータの初期化
    targetCell = "C4" ' パスを入力するセル(targetDir)
    Call PickerFolder(targetCell)
End Sub

Public Sub Main_C6()
    Dim targetCell As String
    ' パラメータの初期化
    targetCell = "C6" ' パスを入力するセル(targetDir)
    Call PickerFolder(targetCell)
End Sub


Private Sub PickerFile(cell As String)

    ' デフォルト値を取得
    Dim currentpath As String
    currentpath = ActiveSheet.Range(cell).value

    ' Pickerを表示
    Dim picker As FileDialog
    Set picker = application.FileDialog(msoFileDialogFilePicker)
    picker.InitialFileName = currentpath
    picker.AllowMultiSelect = False
    picker.Show

    ' 選択されたファイルのパスをセルに入力
    If picker.SelectedItems.Count > 0 Then
        ActiveSheet.Range(cell).value = picker.SelectedItems(1)
    End If

    Set picker = Nothing

End Sub

Private Sub PickerFolder(cell As String)
    ' デフォルト値を取得
    Dim currentpath As String
    currentpath = ActiveSheet.Range(cell).value

    ' Pickerを表示
    Dim picker As FileDialog
    Set picker = application.FileDialog(msoFileDialogFolderPicker)
    picker.InitialFileName = currentpath
    picker.Show

    ' 選択されたフォルダのパスをセルに入力
    If picker.SelectedItems.Count > 0 Then
        ActiveSheet.Range(cell).value = picker.SelectedItems(1)
    End If
    
End Sub
