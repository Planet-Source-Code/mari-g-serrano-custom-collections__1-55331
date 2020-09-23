VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCol   As New Collection
Private mCar  As cCar '<UDC>

Friend Property Get Item(Key As String) As cCar '<UDC>
    If Not Exists(Key) Then
        Set mCar = New cCar '<UDC>
        mCol.Add mCar, Key
    End If
    Set Item = mCol.Item(Key)
End Property

Friend Property Let Item(Key As String, v As cCar) '<UDC>
    If Existe(Key) Then mCol.Remove Key
    mCol.Add v, Key
End Property
Friend Property Set Item(Key As String, v As cCar) '<UDC>
    'Add o borra y add
    If Existe(Key) Then mCol.Remove Key
    mCol.Add v, Key
End Property


Public Property Get Exists(mKey As String) As Boolean
    On Error GoTo NO
    Set mCar = mCol.Item(mKey) '<UDC>
    Exists = True
NO:
End Property
Friend Sub Remove(Key As String)
    On Error Resume Next
    mCol.Remove Key
    
End Sub

Friend Sub RemoveAll()
    Set mCol = Nothing
    Set mCol = New Collection
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "440"
    'Attribute NewEnum.VB_UserMemId = -4
    'Attribute NewEnum.VB_MemberFlags = "40"
    'hide member
    ' proc Id id = -4
    ' (for use with 'For each')
    Set NewEnum = mCol.[_NewEnum]
End Property


Public Property Get Count() As Long
    Count = mCol.Count
End Property

