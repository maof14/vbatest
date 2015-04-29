VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CECB3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public ID
Public BU
Public segment
Public segmentDesc
Public productArea
Public productAreaName
Public pCode
Public productLine
Public pCodeOne
Public pCodeName
Public commodity
Public pcWarranty
Public regionOrNot
Public hmSalesOrNot
Public crossReference
Public paHead
Public paController
Public econApprover
Public comments

Public Function getOnePCode() As String
    getOnePCode = Me.pCodeOne
End Function

Public Function getPCode()
    getPCode = Me.pCode
End Function
