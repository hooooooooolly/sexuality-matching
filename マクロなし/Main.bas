Attribute VB_Name = "Main"
Option Explicit

Const SHEET_NAME_INTRO As String = "�͂��߂�"
Const SHEET_NAME_CATAROG As String = "�ꗗ"
Const SHEET_NAME_FIRST As String = "1�l��"
Const SHEET_NAME_SECOND As String = "2�l��"
Const SHEET_NAME_RESULT As String = "����"

Const CLM_RECOMMENDED As Long = 1   '�������ߗ�
Const CLM_NG          As Long = 2   'NG��

'�f�o�b�O�p
Public Sub forDebug()
    Worksheets(SHEET_NAME_FIRST).Visible = xlSheetVisible
    Worksheets(SHEET_NAME_SECOND).Visible = xlSheetVisible
    Worksheets(SHEET_NAME_RESULT).Visible = xlSheetVisible
End Sub

'1�l�ړ��͊J�n
Public Sub firstInputStart()
   MsgBox "1�l�ڂ̓��͂��J�n���܂��B" & vbCrLf & _
            "2�l�ڂ͉�ʂ����Ȃ��悤�ɂ��Ă��������B" & vbCrLf & _
            "OK����������J�n���܂��B"
    Call Util.hiddenSheet(SHEET_NAME_SECOND)
    Call Util.dispSheet(SHEET_NAME_FIRST)
    Call initSheet(SHEET_NAME_FIRST)
    Worksheets(SHEET_NAME_FIRST).Select
End Sub

'1�l�ړ��͏I��
Public Sub firstInputEnd()
    If canContinue(SHEET_NAME_FIRST) = vbOK Then
        Call Util.hiddenSheet(SHEET_NAME_FIRST)
        MsgBox "�������������܂����B" & vbCrLf & "2�l�ڂɑ��葀��𑱂��Ă��������B"
    End If
End Sub

'2�l�ړ��͊J�n
Public Sub secondInputStart()
    MsgBox "2�l�ڂ̓��͂��J�n���܂��B" & vbCrLf & _
            "1�l�ڂ͉�ʂ����Ȃ��悤�ɂ��Ă��������B" & vbCrLf & _
            "OK����������J�n���܂��B"
    Call Util.hiddenSheet(SHEET_NAME_FIRST)
    Call Util.dispSheet(SHEET_NAME_SECOND)
    Call initSheet(SHEET_NAME_SECOND)
    Worksheets(SHEET_NAME_SECOND).Select
End Sub

'2�l�ړ��͏I��
Public Sub secondInputEnd()
    If canContinue(SHEET_NAME_SECOND) = vbOK Then
        Call Util.hiddenSheet(SHEET_NAME_SECOND)
        MsgBox "�������������܂����B" & vbCrLf & "�Ō��2�l�Ō��ʊm�F�����Ă��������B"
    End If
End Sub

'���ʕ\��
Public Sub showResult()
    MsgBox "���ʂ�\�����܂��B"
    Application.ScreenUpdating = False
    '���ʃV�[�g�̕\��
    Call dispSheet(SHEET_NAME_RESULT)
    
    '���ʂ̎擾
    Dim resultList As New VBA.Collection
    Dim resFetchRow As Long
    For resFetchRow = 1 To Worksheets(SHEET_NAME_FIRST).Cells(Rows.Count, 1).End(xlUp).Row
        Dim res As Result
        Set res = fetchResult(resFetchRow)
        If Not (res Is Nothing) Then
            resultList.Add res
        End If
    Next resFetchRow
    
    '���ʃV�[�g�ւ̓]�L
    With Worksheets(SHEET_NAME_RESULT)
        .Cells(1, CLM_RECOMMENDED).Value = "��������"
        .Cells(1, CLM_NG).Value = "�m�f"
    
        Dim okRow As Long: okRow = 2
        Dim ngRow As Long: ngRow = 2
        Dim iter As Result
        '2�l�Ƃ����}�̏ꍇ
        For Each iter In resultList
            If iter.getResult = RES_GREAT Then
                .Cells(okRow, CLM_RECOMMENDED).Value = iter.itemName
                okRow = okRow + 1
            End If
        Next iter
        '1�l���}�E1�lNG�łȂ��ꍇ
        For Each iter In resultList
            If iter.getResult = RES_GOOD Then
                .Cells(okRow, CLM_RECOMMENDED).Value = iter.itemName
                okRow = okRow + 1
            End If
        Next iter
        '2�l�Ƃ�NG�łȂ��ꍇ
        For Each iter In resultList
            If iter.getResult = RES_NOT_BAD Then
                .Cells(okRow, CLM_RECOMMENDED).Value = iter.itemName
                okRow = okRow + 1
            End If
        Next iter
        '1�l�ȏ�NG�̏ꍇ
        For Each iter In resultList
            If iter.getResult = RES_BAD Then
                .Cells(ngRow, CLM_NG).Value = iter.itemName
                ngRow = ngRow + 1
            End If
        Next iter
        '�Z��������
        .Columns("A:B").EntireColumn.AutoFit
    End With
    Application.ScreenUpdating = True
    MsgBox "�W�v���������܂����B" & vbCrLf & "���ʂ��m�F���Ă��������B"
End Sub

'�w�肵���V�[�g�𖢓��͏�Ԃɂ���
Private Sub initSheet(ByVal sn As String)
    '�A�C�e������R�s�y
    Worksheets(SHEET_NAME_CATAROG).Select
    Worksheets(SHEET_NAME_CATAROG).Range(Cells(1, 1), Cells(Rows.Count, 1).End(xlUp).Rows).Copy
    Worksheets(sn).Cells(1, 1).PasteSpecial
    '�v���C����R�s�y
    Worksheets(SHEET_NAME_CATAROG).Range(Cells(1, 2), Cells(Rows.Count, 2).End(xlUp).Rows).Copy
    '�v���C�s�ԍ���ێ�
    Dim playRow As Long
    playRow = Worksheets(sn).Cells(Rows.Count, 1).End(xlUp).Row + 1
    Worksheets(sn).Cells(playRow, 1).PasteSpecial
    With Worksheets(sn)
        '�Z��������
        .Columns("A:A").EntireColumn.AutoFit
        '�R�����g
        .Cells(1, 2).Value = "��]�x"
        .Cells(playRow, 2).Value = "-"
        '�w�i�F
        .Cells(1, 2).Interior.Color = .Cells(1, 1).Interior.Color
        .Cells(playRow, 2).Interior.Color = .Cells(1, 1).Interior.Color
        '�r��
        .Select
        .Cells(1, 1).CurrentRegion.Borders.LineStyle = xlContinuous
        '�����ʒu�ɖ߂�
        .Cells(1, 1).Select
    End With
End Sub

'���ʎ擾
Private Function fetchResult(ByVal rw As Long) As Result
    '�w�i�F���Ȃ���΃f�[�^���擾����
    Set fetchResult = Nothing
    If Worksheets(SHEET_NAME_FIRST).Cells(rw, 1).Interior.ColorIndex = -4142 Then
        Dim res As New Result
        res.itemName = Worksheets(SHEET_NAME_FIRST).Cells(rw, 1).Value
        res.ansFirst = Worksheets(SHEET_NAME_FIRST).Cells(rw, 2).Value
        res.ansSecond = Worksheets(SHEET_NAME_SECOND).Cells(rw, 2).Value
        Set fetchResult = res
    End If
End Function
