VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Name        As String
Public Color       As Long
Private m_MaxSpeed As Long
'etc...

Public Property Get MaxSpeed() As Long
    MaxSpeed = m_MaxSpeed
End Property
Public Property Let MaxSpeed(lMax As Long)
    m_MaxSpeed = lMax
End Property

Public Function GetColorName() As String
Select Case Color
Case vbBlack: GetColorName = "Black"
Case vbRed: GetColorName = "Red"
Case vbBlue: GetColorName = "Blue"
Case Else: GetColorName = Color
End Select



End Function

