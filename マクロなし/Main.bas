Attribute VB_Name = "Main"
Option Explicit

Const SHEET_NAME_INTRO As String = "はじめに"
Const SHEET_NAME_CATAROG As String = "一覧"
Const SHEET_NAME_FIRST As String = "1人目"
Const SHEET_NAME_SECOND As String = "2人目"
Const SHEET_NAME_RESULT As String = "結果"

Const CLM_RECOMMENDED As Long = 1   'おすすめ欄
Const CLM_NG          As Long = 2   'NG欄

'デバッグ用
Public Sub forDebug()
    Worksheets(SHEET_NAME_FIRST).Visible = xlSheetVisible
    Worksheets(SHEET_NAME_SECOND).Visible = xlSheetVisible
    Worksheets(SHEET_NAME_RESULT).Visible = xlSheetVisible
End Sub

'1人目入力開始
Public Sub firstInputStart()
   MsgBox "1人目の入力を開始します。" & vbCrLf & _
            "2人目は画面を見ないようにしてください。" & vbCrLf & _
            "OKを押したら開始します。"
    Call Util.hiddenSheet(SHEET_NAME_SECOND)
    Call Util.dispSheet(SHEET_NAME_FIRST)
    Call initSheet(SHEET_NAME_FIRST)
    Worksheets(SHEET_NAME_FIRST).Select
End Sub

'1人目入力終了
Public Sub firstInputEnd()
    If canContinue(SHEET_NAME_FIRST) = vbOK Then
        Call Util.hiddenSheet(SHEET_NAME_FIRST)
        MsgBox "処理が完了しました。" & vbCrLf & "2人目に代わり操作を続けてください。"
    End If
End Sub

'2人目入力開始
Public Sub secondInputStart()
    MsgBox "2人目の入力を開始します。" & vbCrLf & _
            "1人目は画面を見ないようにしてください。" & vbCrLf & _
            "OKを押したら開始します。"
    Call Util.hiddenSheet(SHEET_NAME_FIRST)
    Call Util.dispSheet(SHEET_NAME_SECOND)
    Call initSheet(SHEET_NAME_SECOND)
    Worksheets(SHEET_NAME_SECOND).Select
End Sub

'2人目入力終了
Public Sub secondInputEnd()
    If canContinue(SHEET_NAME_SECOND) = vbOK Then
        Call Util.hiddenSheet(SHEET_NAME_SECOND)
        MsgBox "処理が完了しました。" & vbCrLf & "最後に2人で結果確認をしてください。"
    End If
End Sub

'結果表示
Public Sub showResult()
    MsgBox "結果を表示します。"
    Application.ScreenUpdating = False
    '結果シートの表示
    Call dispSheet(SHEET_NAME_RESULT)
    
    '結果の取得
    Dim resultList As New VBA.Collection
    Dim resFetchRow As Long
    For resFetchRow = 1 To Worksheets(SHEET_NAME_FIRST).Cells(Rows.Count, 1).End(xlUp).Row
        Dim res As Result
        Set res = fetchResult(resFetchRow)
        If Not (res Is Nothing) Then
            resultList.Add res
        End If
    Next resFetchRow
    
    '結果シートへの転記
    With Worksheets(SHEET_NAME_RESULT)
        .Cells(1, CLM_RECOMMENDED).Value = "おすすめ"
        .Cells(1, CLM_NG).Value = "ＮＧ"
    
        Dim okRow As Long: okRow = 2
        Dim ngRow As Long: ngRow = 2
        Dim iter As Result
        '2人とも歓迎の場合
        For Each iter In resultList
            If iter.getResult = RES_GREAT Then
                .Cells(okRow, CLM_RECOMMENDED).Value = iter.itemName
                okRow = okRow + 1
            End If
        Next iter
        '1人歓迎・1人NGでない場合
        For Each iter In resultList
            If iter.getResult = RES_GOOD Then
                .Cells(okRow, CLM_RECOMMENDED).Value = iter.itemName
                okRow = okRow + 1
            End If
        Next iter
        '2人ともNGでない場合
        For Each iter In resultList
            If iter.getResult = RES_NOT_BAD Then
                .Cells(okRow, CLM_RECOMMENDED).Value = iter.itemName
                okRow = okRow + 1
            End If
        Next iter
        '1人以上NGの場合
        For Each iter In resultList
            If iter.getResult = RES_BAD Then
                .Cells(ngRow, CLM_NG).Value = iter.itemName
                ngRow = ngRow + 1
            End If
        Next iter
        'セル幅調整
        .Columns("A:B").EntireColumn.AutoFit
    End With
    Application.ScreenUpdating = True
    MsgBox "集計が完了しました。" & vbCrLf & "結果を確認してください。"
End Sub

'指定したシートを未入力状態にする
Private Sub initSheet(ByVal sn As String)
    'アイテム列をコピペ
    Worksheets(SHEET_NAME_CATAROG).Select
    Worksheets(SHEET_NAME_CATAROG).Range(Cells(1, 1), Cells(Rows.Count, 1).End(xlUp).Rows).Copy
    Worksheets(sn).Cells(1, 1).PasteSpecial
    'プレイ列をコピペ
    Worksheets(SHEET_NAME_CATAROG).Range(Cells(1, 2), Cells(Rows.Count, 2).End(xlUp).Rows).Copy
    'プレイ行番号を保持
    Dim playRow As Long
    playRow = Worksheets(sn).Cells(Rows.Count, 1).End(xlUp).Row + 1
    Worksheets(sn).Cells(playRow, 1).PasteSpecial
    With Worksheets(sn)
        'セル幅調整
        .Columns("A:A").EntireColumn.AutoFit
        'コメント
        .Cells(1, 2).Value = "希望度"
        .Cells(playRow, 2).Value = "-"
        '背景色
        .Cells(1, 2).Interior.Color = .Cells(1, 1).Interior.Color
        .Cells(playRow, 2).Interior.Color = .Cells(1, 1).Interior.Color
        '罫線
        .Select
        .Cells(1, 1).CurrentRegion.Borders.LineStyle = xlContinuous
        '初期位置に戻す
        .Cells(1, 1).Select
    End With
End Sub

'結果取得
Private Function fetchResult(ByVal rw As Long) As Result
    '背景色がなければデータを取得する
    Set fetchResult = Nothing
    If Worksheets(SHEET_NAME_FIRST).Cells(rw, 1).Interior.ColorIndex = -4142 Then
        Dim res As New Result
        res.itemName = Worksheets(SHEET_NAME_FIRST).Cells(rw, 1).Value
        res.ansFirst = Worksheets(SHEET_NAME_FIRST).Cells(rw, 2).Value
        res.ansSecond = Worksheets(SHEET_NAME_SECOND).Cells(rw, 2).Value
        Set fetchResult = res
    End If
End Function
