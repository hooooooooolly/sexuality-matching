VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Result"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum ENUM_RESULT
    RES_GREAT = 4
    RES_GOOD = 2
    RES_NOT_BAD = 1
    RES_BAD = 0
End Enum

'çÄñ⁄ñº
Private itemName_  As String
'1êlñ⁄ÇÃâÒìö
Private ansFirst_  As Long
'2êlñ⁄ÇÃâÒìö
Private ansSecond_ As Long

'getter
Public Property Get itemName() As String
    itemName = itemName_
End Property

Public Property Get ansFirst() As Long
    ansFirst = ansFirst_
End Property

Public Property Get ansSecond() As Long
    ansSecond = ansSecond_
End Property

'setter
Public Property Let itemName(ByVal ans As String)
   itemName_ = ans
End Property

Public Property Let ansFirst(ByVal ans As Long)
    ansFirst_ = ans
End Property

Public Property Let ansSecond(ByVal ans As Long)
    ansSecond_ = ans
End Property

'åãâ éÊìæ
Public Function getResult() As ENUM_RESULT
    Dim ans As Long
    ans = ansFirst_ * ansSecond_
    Select Case ans
        Case 4
            getResult = RES_GREAT
        Case 2
            getResult = RES_GOOD
        Case 1
            getResult = RES_NOT_BAD
        Case 0
            getResult = RES_BAD
        Case Default
            getResult = RES_BAD
    End Select
End Function
