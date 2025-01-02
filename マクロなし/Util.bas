Attribute VB_Name = "Util"
Option Explicit

'�V�[�g��\������
'���݂��Ȃ��ꍇ�͒ǉ�����
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

'�V�[�g���\���ɂ���
Public Sub hiddenSheet(ByVal sheetName As String)
    If existWorksheet(sheetName) Then
        Worksheets(sheetName).Visible = xlSheetVeryHidden
    End If
End Sub

'���͊������̑���m�F
Public Function canContinue(ByVal sheetName As String) As Long
    If Util.existsBlank(sheetName) Then
        canContinue = MsgBox("�����͉ӏ�������܂��B" & vbCrLf & "�{���ɓ��͏I�����Ă�낵���ł����H", vbOKCancel)
    Else
        canContinue = MsgBox("�{���ɓ��͏I�����Ă�낵���ł����H", vbOKCancel)
    End If
End Function

'�����͉ӏ����Ȃ����Ƃ��m�F����
Public Function existsBlank(ByVal sheetName As String) As Boolean
    Dim inputArea As Range
    Set inputArea = Worksheets(sheetName).Cells(1, 1).CurrentRegion
    existsBlank = (0 < WorksheetFunction.CountBlank(inputArea))
End Function

'�V�[�g�̑��݊m�F
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

