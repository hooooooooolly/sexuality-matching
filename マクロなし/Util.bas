Attribute VB_Name = "Util"
Option Explicit

'シートを表示する
'存在しない場合は追加する
Public Sub dispSheet(ByVal sheetName As String)
    If existWorksheet(sheetName) Then
        Worksheets(sheetName).Visible = xlSheetVisible
    Else
        Dim oldSheet As Worksheet
        Set oldSheet = ActiveSheet
        Worksheets.Add After:=Worksheets(Worksheets.Count)
        ActiveSheet.Name = sheetName
        oldSheet.Activate
    End If
End Sub

'シートを非表示にする
Public Sub hiddenSheet(ByVal sheetName As String)
    If existWorksheet(sheetName) Then
        Worksheets(sheetName).Visible = xlSheetVeryHidden
    End If
End Sub

'入力完了時の操作確認
Public Function canContinue(ByVal sheetName As String) As Long
    If Util.existsBlank(sheetName) Then
        canContinue = MsgBox("未入力箇所があります。" & vbCrLf & "本当に入力終了してよろしいですか？", vbOKCancel)
    Else
        canContinue = MsgBox("本当に入力終了してよろしいですか？", vbOKCancel)
    End If
End Function

'未入力箇所がないことを確認する
Public Function existsBlank(ByVal sheetName As String) As Boolean
    Dim inputArea As Range
    Set inputArea = Worksheets(sheetName).Cells(1, 1).CurrentRegion
    existsBlank = (0 < WorksheetFunction.CountBlank(inputArea))
End Function

'シートの存在確認
Public Function existWorksheet(ByVal sn As String) As Boolean
    existWorksheet = False
    Dim ws As Worksheet
    For Each ws In Worksheets
        If ws.Name = sn Then
            existWorksheet = True
            Exit For
        End If
    Next ws
End Function

