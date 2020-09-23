VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim Cars As New cCars
'see intellisense...
    Cars.Item("OPEL").MaxSpeed = 100
    Cars.Item("OPEL").Name = "Opel Vectra"
    Cars.Item("OPEL").Color = vbBlue

    Cars.Item("FORD").Color = vbBlack
    Cars.Item("FORD").Name = "Ford Ka"

    With Cars.Item("FERRARI")
        .Name = "Ferrari Testarossa"
        .Color = vbRed
        .MaxSpeed = 350
    End With

    Dim aCar As cCar

    For Each aCar In Cars
        Debug.Print "Name:" & aCar.Name & _
                    " Speed:" & aCar.MaxSpeed & _
                    " Color:" & aCar.GetColorName
    Next


End Sub
