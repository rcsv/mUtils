'
' module name : mWatchdogs.bas
' version     : 1.0.2
' author      : tomohiro awane <Tomohiro.Awane@aisin.co.jp>
'
Option Explicit
'
' Original -> boxDrive deploying script
' ------------------------------------------------------------------------
'
' This script is used to deploy the boxDrive client to a Windows machine.
' It is designed to be run from a network share, and will download the
' boxDrive client from the boxDrive website.
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

' Version 1.0.2 - 2023-04-13
' Version 1.0.1 - 2022-03-01
' Version 1.0.0 - 2021-02-01
'

' Description:
' HDD側にあるオリジナルフォルダの構成に応じて、boxDriveのフォルダ構成・ファイルを
' 配備します。
' ・pptx は、.pdfに変換してから配備
' ・ファイル名の先頭に「_」がある場合は、配備対象外
' ・配備したファイルは、オリジナルフォルダに「_」を付与して終了
' ・boxDriveのフォルダ構成は、オリジナルフォルダの構成に合わせて作成
' ・boxDrive側に独自ファイルがある場合は、退避フォルダに移動
' ・配備したファイルをシート「result」に日付、ファイル名、ファイルサイズ、配備先フォルダを記録

' Excel VBA で動作して、Public Sub Main()がエントリーポイントで、ボタンから実行される
' Main は、対象フォルダ、boxDriveのルートパス、退避フォルダの初期化。
' 上記に関連する変数がシート「config」に記載されていれば、それを読み込む
' なければ、初期値を設定
' その後、対象フォルダを再帰的に走査し、ファイルを配備する
' 同じ名前・同じ日付・時刻のファイルがboxDriveにあれば、配備しない


Private Const DS As String = "\"

' シート名の定義
Private Const MWD_SHEET_CONFIG As String = "Config"
Private Const MWD_SHEET_RESULT As String = "Result"

' デフォルト値の定義
Private Const MWD_DEFAULT_TARGET_FOLDER As String = "C:\Users\Public\Documents\boxDrive"
Private Const MWD_DEFAULT_BOXDRIVE_ROOT As String = "C:\BoxDrive\"
Private Const MWD_DEFAULT_BACKUP_FOLDER As String = "C:\Users\Public\Documents\boxDrive\_backup"

Public Sub Main()
    ' Entrypoint of this module
    Dim targetFolder As String
    Dim boxDriveRoot As String
    Dim backupFolder As String

    ' シート「config」から、対象フォルダ、boxDriveのルートパス、退避フォルダを取得
    ' なければ、デフォルト値を設定
    targetFolder = GetConfigValue(MWD_SHEET_CONFIG, "targetFolder", MWD_DEFAULT_TARGET_FOLDER)
    boxDriveRoot = GetConfigValue(MWD_SHEET_CONFIG, "boxDriveRoot", MWD_DEFAULT_BOXDRIVE_ROOT)
    backupFolder = GetConfigValue(MWD_SHEET_CONFIG, "backupFolder", MWD_DEFAULT_BACKUP_FOLDER)

    ' 対象フォルダを再帰的に走査し、ファイルを配備する
    ' 相対パスを設定して、ルートフォルダからスタート
    Call ScanDirs(targetFolder, boxDriveRoot, backupFolder)

    ' boxDriveのフォルダ構成を整理
    Call CleanUp(boxDriveRoot)

End Sub

Private Sub CleanUp(boxDriveRoot As String)
    ' boxDriveのフォルダ構成を整理

    ' boxDriveのルートフォルダを開く
    ' フォルダを取得
    ' フォルダがあれば、再帰的に走査
    ' サブフォルダもファイルもなければそのフォルダは削除
    Dim fs As New FileSystemObject
    Dim folder As folder
    Dim subFolder As folder
    Dim subFolders As Folders

    Set folder = fs.GetFolder(boxDriveRoot)
    Set subFolders = folder.subFolders

    For Each subFolder In subFolders
        Call CleanUp(subFolder.Path)
    Next

    If subFolders.Count = 0 And folder.Files.Count = 0 Then
        folder.Delete
    End If

End Sub

' ScanDirs
' ------------------------------------------------------------------------
' 対象フォルダを再帰的に走査し、ファイルを配備する
' 同じ名前・同じ日付・時刻のファイルがboxDriveにあれば、配備しない
' 配備したファイルをシート「result」に日付、ファイル名、ファイルサイズ、配備先フォルダを記録
' fsオブジェクトの作成→targetFolderを開く→サブフォルダを取得→サブフォルダがあれば、再帰的に走査
'
' Need: Tool->References->Microsoft Scripting Runtime
'
Private Sub ScanDirs(targetFolder As String, ByVal boxDriveRoot As String, backupFolder As String, Optional relativePath As String = "")

    ' fso でフォルダを探索し、フォルダがあれば相対パスを作って自己呼び出し
    ' ファイルがあれば、配備する
    Dim fso As FileSystemObject
    Dim folder As folder

    ' フォルダを開く
    Set fso = New FileSystemObject
    If Dir(targetFolder & DS & relativePath, vbDirectory) = "" Then
        MsgBox "folder not found"
        Exit Sub
    End If
    
    Set folder = fso.GetFolder(targetFolder & DS & relativePath)

    ' サブフォルダを取得
    Dim subFolder As folder
    For Each subFolder In folder.subFolders
        ' サブフォルダがあれば、再帰的に走査
        Call ScanDirs(targetFolder, boxDriveRoot, backupFolder, relativePath & DS & subFolder.Name)
    Next

    ' ファイルを取得
    Dim file As file
    For Each file In folder.Files
        ' ファイルがあれば、配備する
        Call DeployFile(file, boxDriveRoot, relativePath)
    Next


End Sub

' DeployFiles
Private Sub DeployFile(file As file, deployDest As String, relativePath As String)
    ' ファイルを配備する
    ' ファイル名とdeployDestでコピー先のフルパスを作成
    ' 同じ名前・同じ日付・時刻のファイルがboxDriveにあれば、配備しない
    ' 配備したファイルをシート「result」に日付、ファイル名、ファイルサイズ、配備先フォルダを記録
    Dim deployDestFullPath As String

    ' relativepath の有無で deployDestのアップデート
    ' relativePathが空文字の場合は、DSを付けない
    If relativePath <> "" Then
        deployDest = deployDest & relativePath
    End If

    ' deployDestがなければ作成
    If Dir(deployDest, vbDirectory) = "" Then
        MkDir deployDest
        ' フォルダを作成したことをシート「result」に記録
        Call WriteResultDir(deployDest, "")
        
    End If

    ' ファイル名とdeployDestでコピー先のフルパスを作成
    deployDestFullPath = deployDest & DS & file.Name

    ' 同じ名前・同じ日付・時刻のファイルがboxDriveにあれば、配備しない
    If Dir(deployDestFullPath) <> "" Then
        If file.DateLastModified = FileDateTime(deployDestFullPath) Then
            Exit Sub
        End If
    End If

    ' ファイルをコピー
    ' PPTX→PDFにしてからコピー
    If Right(file.Name, 5) = ".pptx" Then
        Call ConvertPPTXtoPDF(file, deployDestFullPath)
    Else
        file.Copy deployDestFullPath, True
    End If

    ' 配備したファイルをシート「result」に日付、ファイル名、ファイルサイズ、配備先フォルダを記録
    Call WriteResult(file, deployDestFullPath)

    
End Sub

' GetConfigValue
' ------------------------------------------------------------------------
' 指定したシートから、対象フォルダ、boxDriveのルートパス、退避フォルダを取得
' なければ、デフォルト値を設定
Private Function GetConfigValue(sheetName As String, key As String, defaultValue As String) As String
    ' シート「config」から、対象フォルダ、boxDriveのルートパス、退避フォルダを取得
    ' なければ、デフォルト値を設定
    Dim sheet As Worksheet
    Dim cell As Range
    Dim value As String

    Set sheet = ThisWorkbook.Sheets(sheetName)
    Set cell = sheet.Range(key)
    If cell.value = "" Then
        value = defaultValue
    Else
        value = cell.value
    End If

    GetConfigValue = value
End Function

' WriteResult
' ------------------------------------------------------------------------
Private Sub WriteResult(file As file, deployDestFullPath As String)
    ' 配備したファイルをシート「result」に日付、ファイル名、ファイルサイズ、配備先フォルダを記録
    Dim sheet As Worksheet
    Dim cell As Range
    Dim lastRow As Long

    Set sheet = ThisWorkbook.Sheets(MWD_SHEET_RESULT)
    lastRow = sheet.Range("B" & Rows.Count).End(xlUp).Row + 1
    Set cell = sheet.Range("B" & lastRow)
    cell.value = Now
    cell.Offset(0, 1).value = file.Name
    cell.Offset(0, 2).value = file.Size
    cell.Offset(0, 3).value = deployDestFullPath
End Sub

' WriteResultDir
' ------------------------------------------------------------------------
Private Sub WriteResultDir(deployDest As String, deployDestFullPath As String)
    ' ディレクトリ作成をシート「result」に記録
    Dim sheet As Worksheet
    Dim cell As Range
    Dim lastRow As Long

    Set sheet = ThisWorkbook.Sheets(MWD_SHEET_RESULT)
    lastRow = sheet.Range("B" & Rows.Count).End(xlUp).Row + 1
    Set cell = sheet.Range("B" & lastRow)
    cell.value = Now
    cell.Offset(0, 1).value = deployDest
    cell.Offset(0, 2).value = 0
    cell.Offset(0, 3).value = deployDestFullPath
    
End Sub
' ConvertPPTXtoPDF
' ------------------------------------------------------------------------
' Need:
' Microsoft PowerPoint 15.0 Object Library
' Microsoft Scripting Runtime
Private Sub ConvertPPTXtoPDF(file As file, deployDestFullPath As String)
    ' PPTXをPDFに変換してからコピー
    Dim ppt As PowerPoint.application
    Dim pptPresentation As PowerPoint.Presentation
    Dim pdfPath As String

    ' PPTXを開く
    Set ppt = New PowerPoint.application
    'ppt.Visible = False ' pptを非表示にする

    Set pptPresentation = ppt.Presentations.Open(file.Path)

    ' PDFに変換
    pdfPath = Left(deployDestFullPath, Len(deployDestFullPath) - 4) & "pdf"
    pptPresentation.SaveAs pdfPath, ppSaveAsPDF

    ' PPTXを閉じて解放する
    pptPresentation.Close
    ppt.Quit
    Set pptPresentation = Nothing
    Set ppt = Nothing


End Sub
