Attribute VB_Name = "modErrorHandler"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''      Method: Report
'' Description: Report warnings and errors to KLog
''        Date: Jan 12, 2001
''       Input: None
''      Output: None
''Modification: Jan 12, 2001 - Aman Anwari
''              Original written in VB6
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub Report(ByVal lNumber As Long, ByVal sAppName As String, _
                    ByVal sObjectName As String, ByVal sMethodName As String, _
                    ByVal lLineNo As Long, Optional ByVal sLogMessage As String)

    ' Report error if subscribe
    If Not g_objKLog Is Nothing Then
        ' Report error
        Call g_objKLog.Report(lNumber, sAppName, sObjectName, sMethodName, lLineNo, sLogMessage)
    Else
        'Shital, 23Jan2k2, If KLog isn't there, show the message instead of remaining silent
        MsgBox "Following error occured in the application. This error is not logged because KLog is not available." & vbCrLf _
            & "Error " & lNumber & " : " & sLogMessage & vbCrLf & _
            "Method: " & sMethodName & ", Object: " & sObjectName & ", Application: " & sAppName & ", Line: " & lLineNo
    End If
    
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''      Method: IsKPMGCode
'' Description: Check if error generated is KPMG defined error code
''        Date: Mar 26, 2001
''       Input: Error code
''      Output: True if KPMG defined error code Else False
''Modification: Mar 26, 2001 - Aman Anwari
''              Original written in VB6
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function IsKPMGCode(ByVal lErrorNumber As Long) As Boolean
    
    ' Report error if subscribe
    If Not g_objKLog Is Nothing Then
        ' Check if KPMG defined error code
        IsKPMGCode = g_objKLog.IsKPMGCode(lErrorNumber)
    Else
        IsKPMGCode = False
    End If
    
End Function

'This function is to be used in future to replace repetitive KLog implementation
'code in all modules
Public Sub HandleError(ByVal MethodName As String, Optional ByVal ComponentTypeName As String = vbNullString, Optional ByVal CustomErrorNumber As Long = 0)
    
    '#NO_ERROR_HANDLER
    
    ' Report generated error
    Call Report(Err.Number, App.EXEName & " : " & Err.Source, ComponentTypeName, MethodName, Erl, Err.Description)

    'If there's custom error code
    If CustomErrorNumber <> 0 Then
        Call Report(CustomErrorNumber, App.EXEName, ComponentTypeName, MethodName, Erl)
        'Do not raise error if it is KPMG code (i.e. if error code has been entered in XML)
        If Not IsKPMGCode(CustomErrorNumber) Then ReRaiseError
    Else
        'For all other errors always reraise it
        Call ReRaiseError
    End If

End Sub



