Option Explicit

' mReplica.bas
' ------------------------------------------------------------------------
'
' License : MIT
' https://opensource.org/licenses/MIT
' @version 1.0.0
'

' mReplica パッケージは、任意の場所にあるオリジナルファイルを基に、
' ダミーファイルを様々な名前で生成するためのスクリプトです。
' オリジナルファイルは、ボタンを設置したシートのどこかにあります
'
' 注意：このマクロで生成するダミーファイルは、エクセルと同じフォルダに固定
'

'
' 定数定義
' ------------------------------------------------------------------------
' 元ファイルのフルパスが記載されているセル
Private Const CELL_BASEFILENAME As String = "C3"
' コピーファイル名が存在する列
Private Const COL_RENAME_LIST As String = "B"
' コピーファイル名の一覧が開始する行
Private Const START_RENAME_LIST As Integer = 5


'
' 関数
' ------------------------------------------------------------------------
' ファイルをコピーするスクリプトのみ抽出
Private Sub makeReplica(original As String, destination As String)

    ' use FileSystemObject
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' get Extension
    Dim ext As String
    ext = fso.GetExtensionName(original)

    ' copy with original extension
    fso.CopyFile original, destination & "." & ext

    ' release
    Set fso = Nothing
    
End Sub

'
' 本モジュールの主たる部分
' ------------------------------------------------------------------------
' ボタンから呼ばれる部分
Public Sub Replication()

    ' Set Destination Base
    Dim dest_base As String
    dest_base = ThisWorkbook.Path

    ' 任意のセルから、複製元のファイル名を取得する
    Dim original_file_name As String
    original_file_name = Range(CELL_BASEFILENAME).Value

    ' 任意のテーブルの1行目から開始し、１行ずつ、複製ファイル名を取得
    Dim destination_file_name As String
    Dim i As Integer

    ' 繰り返し処理によるテーブル読み込み開始
    i = START_RENAME_LIST
    While Range(COL_RENAME_LIST & i).Value <> ""
      destination_file_name = Range(COL_RENAME_LIST & i).Value
      makeReplica original_file_name, dest_base & Application.PathSeparator & destination_file_name & ".mp4"
      i = i + 4
    Wend
    
End Sub
