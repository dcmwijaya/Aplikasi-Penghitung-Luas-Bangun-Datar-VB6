VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassHitung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function LJG(ByVal a As Integer, ByVal t As Integer)
    LJG = a * t
End Function

Public Function LBK(ByVal d1 As Integer, ByVal d2 As Integer)
    LBK = 0.5 * d1 * d2
End Function

Public Function LSL(ByVal r As Integer)
    LSL = 22 / 7 * r
End Function
