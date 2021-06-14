'
' # Utility Class
' -----------------------------------------------------------------------

Option Explicit
' Public NotInherits Class Utils


    ' 設定項目が始まる行数
    Private Const rowSTART As Integer = 4 ' configure data starts 4

    ' ログファイルの名前 - 直下に作成します。
    Private Const strLOGFILE As String = "\_err.log"
    
    ' getConfigで使用する、設定項目をもとに取得済みにしておく hashmap
    Private hashMap As Object
    
    ' getConfigで使用する、データテーブルが始まる行・列を定義
    Private Const iDATA_START_ROW As Integer = 4            ' 4 行目 - テーブル開始行
    Private Const iDATA_KEY As Integer = 2                  ' 2 列目 - キー列
    Private Const iDATA_VALUE As Integer = 3   ' 3 列目 - 値列
    
    ' setResultでデバッグ情報やプログラムの進捗状況をログ形式ではなく
    ' サマリ形式で出力するための列情報
    Private Const iRESULT_KEY As Integer = 6
    Private Const iRESULT_VALUE As Integer = 7

    ' getParenthesisで使用する、ファイル名等、抽出する対象のカッコのタイプ
    ' 括弧のタイプを表現している列挙型
    Public Enum BracketType
        parenthesis = 1      ' () parentheses ASCII
        brace = 2            ' {} braces, Curly brackets
        bracket = 3          ' [] square brackets
        quote = 4            ' '' single quotes
        jp_brace = 5         ' 「」鍵括弧
        jp_bracket = 6       ' 【】
        wide_parenthesis = 7 ' （） full-width parenthesis
        wide_brace = 8       ' ｛｝ full-width braces, curly brackets
        wide_bracket = 9     ' ［］ full-width square brackets
        jp_double_bracket = 10 ' 『』 Japanese double hook brackets
    End Enum
    
    Private Const HANAKUSO As String = "●" ' 標準の見たよマーク
    
    ' Define a constant id number
    Private Const myContextID As Integer = 10


    ' DEBUG SWITCH
    ' configシートに記入しておくDEBUGスイッチのキー部分
    Private Const flgDEBUG As String = "DEBUG"
 
    '
    ' dPrint
    ' -------------------------------------------------------------------
    ' コンソールとログファイルの両方に、同じ文字列を出力する
    ' 出力する文字列は下記のフォーマットに準じる
    '
    ' YYYY-MM-DD HH:mm:ss メッセージ
    ' (時間とメッセージの間はタブ文字１つ)
    '
    Public Sub dPrint(ByVal msg As String)
        Dim buf, objFSO As Object, path

        buf = Now & vbTab & msg
        Debug.Print buf

        Set objFSO = CreateObject("Scripting.FileSystemObject")
        
        path = ThisWorkbook.path & strLOGFILE
        
        ' Dropbox / OneDrive / Sharepoint 等の仮想フォルダは FileSystemObject で操作できない
        path = cnvNetPath2Local(path)

        With objFSO
        If Not .fileExists(path) Then .createTextFile (path)
            With .OpenTextFile(path, 8)
            .writeline buf
            .Close
            End With
        End With

        Set objFSO = Nothing
    End Sub
    
    '
    ' getConfig
    ' -------------------------------------------------------------------
    ' シートにある設定項目を読み出す
    ' 入力は文字列、出力も文字列に限定。
    '
    Public Function getConfig(key As String) As String
        Dim hm As Object ' hashmap object 取得用の変数
        Set hm = getInstance()
        getConfig = hm.Item(key)
    End Function

    ' getInstance - singleton pattern
    Private Function getInstance() As Object
        ' getInstanceを呼び出して、何度も同じ変数の初期化をさせない
        If hashMap Is Nothing Then
            Call initHM(hashMap)
        End If
        Set getInstance = hashMap
    End Function
    
    ' getConfigで初期化がされていない場合getInstance経由で呼ばれる
    ' 実際の初期化メソッド #configシートから、設定情報をすべて取得
    ' HashMap形式で取り込むため、探索時間は短く済む (log(n)?)
    Private Sub initHM(ByRef hm As Object)
        ' getConfig初期設定、情報は wsCONF シートから取得
        ' hashMapの.NET名は Dictionary
        Set hm = CreateObject("Scripting.Dictionary")
        
        ' config table start with iDATA_START_ROW
        Dim i: i = iDATA_START_ROW
        
        ' 一時的に情報を保存する変数を用意し空文字で初期化する
        Dim key As String: key = ""
        Dim value As String: value = ""
        
        ' configシートの列"Key"に文字が入っている限り
        ' key / value のペアの取得を続ける
        Do While wsCONF.Cells(i, iDATA_KEY) <> ""
            key = wsCONF.Cells(i, iDATA_KEY)
            value = wsCONF.Cells(i, iDATA_VALUE)
            
            ' ペアを登録してカウンタをインクリメント
            hm.Add key, value
            i = i + 1
        Loop
    End Sub
    ' -------------------------------------------------------------------

    '
    ' checkDEBUG
    ' -------------------------------------------------------------------
    ' configシートに DEBUG という行がある場合はその値を持ってくる。
    ' DEBUG という行があり、なおかつ Trueが入っている場合のみ、Trueが返される
    Public Function checkDEBUG() As Boolean
        checkDEBUG = getConfig(flgDEBUG)
    End Function

    ' setResult
    ' --------------------------------------------------------------------
    ' config シートにメモを残す
    ' result テーブルは上から順に情報を確認していく
    '
    Public Sub setResult(ByVal key As String, ByVal value As String)
        Call writeConf(key, value, iRESULT_KEY, iRESULT_VALUE, wsCONF)
    End Sub

    '
    ' setConfig
    ' -------------------------------------------------------------------
    ' 設定項目の再設定を行う関数。割と使わない。
    '
    Public Sub setConfig(ByVal key As String, ByVal value As String)
        Call writeConf(key, value, iDATA_KEY, iDATA_VALUE)
    End Sub

    ' writeConf
    ' configファイルに対して情報を書き込むための関数
    ' 設定テーブルが何処でも大丈夫なよう、シート・設定テーブル開始行列は
    ' 同時に指定することができる。
    '
    Private Sub writeConf(ByVal key As String, ByVal value As String, colKey As Integer, colValue As Integer, _
        Optional ws As Object)
        Dim i As Integer: i = iDATA_START_ROW

        ' CONSTLAINT : wsCONF という名前のシートが必要
        If ws Is Nothing Then
            ws = wsCONF
        End If

        Do While ws.Cells(i, colKey) <> ""
            ' key に相当する文字列と同じ文字があった場合隣にvalue記入
            If ws.Cells(i, colKey) = key Then
                ws.Cells(i, colValue) = value
                ws.Cells(i, colValue + 1) = Now
                Exit Sub
            End If
            i = i + 1
        Loop

        ' unknown error テーブルのスキャンが行われなかった
        ws.Cells(i, colKey) = key
        ws.Cells(i, colValue) = value
        ws.Cells(i, colValue + 1) = Now

    End Sub

    '
    ' FolderPicker
    ' -------------------------------------------------------------------
    ' フォルダを選択するダイアログを表示して選んでもらう
    ' 正直数行しかないので、別途関数化する必要はない程度
    ' (だが頻繁に使うよね。)
    '
    Public Function FolderPicker() As String
        With Application.FileDialog(msoFileDialogFolderPicker)
            .AllowMultiSelect = False
            .Title = "対象となるフォルダの選択"
            If .Show = True Then
                FolderPicker = .SelectedItems(1)
            End If
        End With
    End Function


    '
    ' cnvNetPath2Local
    ' -------------------------------------------------------------------
    ' ファイルが開けないネットワークドライブパスを通常のフォルダパスに
    ' 変換。本メソッドでは通常 http: の削除のみ行う
    ' OneDrive形式：http://servername/c/sample/
    ' 通常の形式　：\\servername\c\sample\
    '
    Public Function cnvNetPath2Local(ByVal path_withScheme As String) As String
    
        If InStr(path_withScheme, "http://") > 0 Then
            ' http://で始まっていることを確認し、冒頭のスキーム定義部分のみを削除する
            cnvNetPath2Local = Replace(path_withScheme, "http:", "")
        Else
            ' 検出できなかった場合は、そのままパスを返す
            cnvNetPath2Local = path_withScheme
        End If
    End Function


    '
    ' getBraceInside
    ' --------------------------------------------------------------------
    ' 【】や（）、() 等、ブレースに限らずカッコで区切られた文字を抽出する
    ' 対応するカッコは10種類
    ' @param paren_type 列挙型 BracketType のどれか。
    ' デフォルト値は【】jp_bracket。
    '
    Public Function getBraceInside(ByVal str As String, _
        Optional paren_type As BracketType = jp_bracket) As String

        Dim pos_start As Integer, pos_end As Integer
        Dim strSTART As String, strEND As String
        getBraceInside = ""

        Select Case paren_type
            Case parenthesis
                strSTART = "(": strEND = ")"
            Case brace
                strSTART = "{": strEND = "}"
            Case bracket
                strSTART = "[": strEND = "]"
            Case quote
                strSTART = "'": strEND = "'"
            Case jp_brace
                strSTART = "「": strEND = "」"
            Case jp_bracket
                strSTART = "【": strEND = "】"
            Case wide_parenthesis
                strSTART = "（": strEND = "）"
            Case wide_brace
                strSTART = "｛": strEND = "｝"
            Case wide_bracket
                strSTART = "［": strEND = "］"
            Case jp_double_bracket
                strSTART = "『": strEND = "』"
            Case Else
                Err.Raise vbObjectError * myContextID + 515, Description:="unknwon parenthesis type"
        End Select

        pos_start = InStr(str, strSTART)
        pos_end = InStr(str, strEND)
        pos_end = pos_end - pos_start - 1

        getBraceInside = Mid(str, (pos_start + 1), pos_end)
        Call dPrint("String: " & getBraceInside & " distilled from: " & str)
    End Function

    '
    ' addMITAYO
    ' --------------------------------------------------------------------
    ' ファイル名文字列をパラメータで取得し、
    ' 末尾（拡張子の直前）に、●をつけた文字列を生成する
    ' 確認済みのマークがついている場合、●は増えない。同じ文字列が返る
    '
    ' @param str ファイル名
    ' @param flgCHAR 確認済みのマーク（任意、デフォルトは●HANAKUSO）
    ' @return 確認済みの文字がついた文字列
    '
    Public Function addMITAYO(ByVal str As String, _
            Optional flgCHAR As String = HANAKUSO) As String

        Dim posLastDot As Integer
        Dim strExt As String

        ' 文字列（この場合ファイル名）を調べて、拡張子以外の文字数を調べる
        posLastDot = InStrRev(str, ".")
        strExt = Right(str, Len(str) - posLastDot)
        If Mid(str, posLastDot - 1, 1) <> flgCHAR Then
            ' HANAKUSO を拡張子の前につける
            addMITAYO = Left(str, posLastDot - 1) & flgCHAR & "." & strExt
        Else
            ' すでに HANAKUSOがついているファイル名の場合は何もしない
            addMITAYO = str
        End If
    End Function

    '
    ' wrapXxxx
    ' - wrapOpen, wrapSave, wrapClose
    ' --------------------------------------------------------------------
    ' エクセルファイルから、別のエクセルファイルの操作を行うシーケンスを
    ' まとめたもの。
    ' 開く・保存・閉じる操作の前後で、警告のダイアログ表示を抑制
    '
    Public Function wrapOpen(target As String) As Workbook
        dPrint "try2open: " & target
        Application.DisplayAlerts = False
        Set wrapOpen = Workbooks.Open(target, False)
        Application.DisplayAlerts = True
    End Function

    Public Sub wrapSave(target As Workbook)
        Application.DisplayAlerts = False
        target.Save
        Application.DisplayAlerts = True
    End Sub

    Public Sub wrapClose(target As Workbook)
        Application.DisplayAlerts = False
        target.Close
        Application.DisplayAlerts = True
    End Sub

    '
    ' wrapMessageBox
    ' --------------------------------------------------------------------
    ' メッセージボックスに統一感を出すためのマクロ
    ' 要設定：タイトル
    '
    Public Function wrapMsgDlg(msg As String, Optional btns As VbMsgBoxStyle = vbOKOnly + vbExclamation) As Long
        wrapMsgDlg = MsgBox(msg, Buttons:=btns, Title:=getConfig("タイトル"))
    End Function

    '
    ' ResetTextFormatting
    ' --------------------------------------------------------------------
    ' wsを直接操作して、テキストフォーマットや条件付き書式をリセットする
    ' @param ws a worksheet object
    ' 要設定："行：開始"、"列：右端"
    '
    Public Sub ResetTextFormatting(ws As Worksheet)
        Dim tempx As Integer, tempy As Integer
        Dim r1 As Range, r2 As Range

        tempx = getConfig("行：開始")
        tempy = getConfig("列：右端") ' TODO: UNIFY orignal source
     
        Set r1 = ws.Cells(tempx, 1)
        Set r2 = ws.Cells(50, tempy)
        ws.Range(r1, r2).Select
        Cells.FormatConditions.Delete
    End Sub
    
    


