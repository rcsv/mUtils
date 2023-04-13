'
' module name : mConf.bas
' version     : 1.0.1
' author      : rcsvpg@outlook.com
'
Option Explicit
'
' KVS
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
    '
    ' getConfigを使用するために必要な情報
    ' テーブルがあるシート名
    Private Const gC_WS_NAME As String = "Conf"

    ' 1. テーブル「開始行」(初期値:4行目から開始) ヘッダは入れない
    Private Const gC_i_ROW_START As Integer = 4

    ' 2. 「キー」となる列
    Private Const gC_i_COL_KEY As Integer = 2

    ' 3. 「値」の列
    Private Const gC_i_COL_VALUE As Integer = 3

    ' getConfigで使用するHashmap Object
    private hashMap As Object

    '
    ' getConfig 本体
    ' ----------------------------------------------------
    ' シートにある設定項目を読み出す。入力は文字列
    ' 出力も文字列に限定
    Public Function getConfig(ByRef key As String) As String
        Dim hm As Object ' hash map object 取得用の変数
        Set hm = getInstance()
        getConfig = hm.Item(key)
    End Function

    ' getInstance - singleton pattern
    Private Function getInstance() As Object
        ' getInstanceを呼び出して、何度もHashMapの生成をさせない
        If hashMap Is Nothing Then
            Call initHM(hashMap)
        End If
        Set getInstance = hashMap
    End Function

    ' initHM - initialize HashMap Object
    ' getConfig を呼び出した際に、hashMap が初期化されていない場合に
    ' 限り呼び出される初期化サブルーチン
    Private Sub initHM(ByRef hm As Object)
        Set hm = CreateObject("Scripting.Dictionary")

        ' テーブル開始行からデータ取得開始
        Dim i As Integer: i = gC_i_ROW_START
        Dim key As String: key = ""
        Dim value As String: value = ""

        ' 該当するワークシートがないとエラー
        Dim ws As Worksheet: Set ws = Worksheets(gC_WS_NAME)

        ' key列に情報が入っている限り、key / value のペア取得を続ける
        Do While ws.Cells(i, gC_i_COL_KEY) <> ""
            key = ws.Cells(i, gC_i_COL_KEY)
            value = ws.Cells(i, gC_i_COL_VALUE)

            ' set into Key-Value-Store and increase counter number
            hm.add key, value
            i = i + 1
        Loop
    End Sub
