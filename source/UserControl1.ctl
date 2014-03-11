VERSION 5.00
Begin VB.UserControl UserControl1 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Property Get p1()
    Dim x
    'hi
    MsgBox "1 +1"
    Dim z
    
    'jojo
101:
    Call f1: Call f2
End Property

Sub f1()

End Sub

Sub f2()

End Sub

