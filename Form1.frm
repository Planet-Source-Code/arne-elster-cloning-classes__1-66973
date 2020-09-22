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
   StartUpPosition =   3  'Windows-Standard
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim clsFirstInstace     As SimpleClass
    Dim clsSecondInstance   As SimpleClass

    Set clsFirstInstace = New SimpleClass

    With clsFirstInstace
        .MySingle = 1.5
        .MyString = "text"
        .MyVariable = 1000
        .MyUDTvalue1 = 2000
        .MyUDTvalue2 = "UDT text"
    End With

    ' clsSecondInstance is not a reference to clsFirstInstance,
    ' but a new instance, with all the values from clsFirstInstance!
    Set clsSecondInstance = clsFirstInstace.Clone()

    With clsSecondInstance
        Debug.Print .MySingle
        Debug.Print .MyString
        Debug.Print .MyVariable
        Debug.Print .MyUDTvalue1
        Debug.Print .MyUDTvalue2
    End With
End Sub
