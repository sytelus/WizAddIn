Attribute VB_Name = "modAddToIni"
Option Explicit

Public Const gsADD_IN_PRJ_NAME As String = "WizAddIn"

Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)

Sub AddToINI()
    Dim rc As Long
    rc = WritePrivateProfileString("Add-Ins32", gsADD_IN_PRJ_NAME & ".AddIn", "0", "VBADDIN.INI")
End Sub


