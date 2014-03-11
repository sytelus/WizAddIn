Attribute VB_Name = "modLineNumbers"
Option Explicit

Public Sub AddLineNumbersToProjects(ByVal voVBInstance As VBIDE.VBE)
    Call AddOrRemoveLineNumbers(voVBInstance, True)
End Sub

Public Sub RemoveLineNumbersFromProjects(ByVal voVBInstance As VBIDE.VBE)
    Call AddOrRemoveLineNumbers(voVBInstance, False)
End Sub

Private Sub AddOrRemoveLineNumbers(ByVal voVBInstance As VBIDE.VBE, ByVal vblnAddOrRemove As Boolean)
    
    On Error GoTo ErrorTrap
    
    Dim oVBProject As VBIDE.VBProject
    Dim oVBComponent As VBIDE.VBComponent
    Dim oCodeModule As VBIDE.CodeModule
    Dim oVBMethod As VBIDE.Member
    Dim lMethodStartLine As Long
    Dim lMethodLineCount As Long
    Dim lMethodLineIndex As Long
    Dim sOriginalLineNumber As String
    Dim lFirstWordStart As Long
    Dim lFirstWordLen As Long
    Dim sMethodLine As String
    Dim bLineNumberExist As Boolean
    Dim evptVBMethodType As vbext_ProcKind
    Dim bValidMethod As Boolean
    Dim bManualLineNumber As Boolean
    Dim sReadOnlyModuleNames As String
    
    sReadOnlyModuleNames = vbNullString
    For Each oVBProject In voVBInstance.VBProjects
        For Each oVBComponent In oVBProject.VBComponents
            Set oCodeModule = oVBComponent.CodeModule
            If Not (oCodeModule Is Nothing) Then
                Dim lVBMethodIndex As Long
                lVBMethodIndex = 1
                Do While lVBMethodIndex <= oCodeModule.Members.Count
                    Set oVBMethod = oCodeModule.Members(lVBMethodIndex)
                    If (oVBMethod.Type = vbext_mt_Method) Or (oVBMethod.Type = vbext_mt_Property) Then
                        bValidMethod = True
                        Dim sMethodDefinationLine As String
                        sMethodDefinationLine = GetMethodStartLine(oCodeModule, oVBMethod)
                        Dim oWords As New Collection
                        Call MakeWordList(sMethodDefinationLine, oWords)
                        If oVBMethod.Type <> vbext_mt_Property Then
                            evptVBMethodType = vbext_pk_Proc
                            If oWords.Count >= 2 Then
                                If (LCase(oWords(1)) = "declare") Or (LCase(oWords(2)) = "declare") Then
                                    bValidMethod = False
                                Else
                                    'Method is not API declaration
                                End If
                            Else
                                'Method is not API declaration
                            End If
                        Else
                            If oWords.Count >= 2 Then   'Check the second word
                                Select Case LCase(oWords(2))
                                    Case "get"
                                        evptVBMethodType = vbext_pk_Get
                                    Case "let"
                                        evptVBMethodType = vbext_pk_Let
                                    Case "set"
                                        evptVBMethodType = vbext_pk_Set
                                    Case Else   'Check the 3rd word
                                        If oWords.Count >= 3 Then
                                            Select Case LCase(oWords(3))
                                                Case "get"
                                                    evptVBMethodType = vbext_pk_Get
                                                Case "let"
                                                    evptVBMethodType = vbext_pk_Let
                                                Case "set"
                                                    evptVBMethodType = vbext_pk_Set
                                                Case Else
                                                    bValidMethod = False
                                            End Select
                                        Else
                                            bValidMethod = False
                                        End If
                                End Select
                            Else
                                bValidMethod = False
                            End If
                            Set oWords = Nothing
                        End If
                        
                        If bValidMethod Then
                            
                            'Due to bug in VB, Members Collection does not includes Let if there also exist Get property. Here's the worka around
                            Dim lPropertyTypeIndex As Long
                            Dim lPropertyTypeCount As Long
                            
                            If ((evptVBMethodType = vbext_pk_Get) Or (evptVBMethodType = vbext_pk_Let) Or (evptVBMethodType = vbext_pk_Set)) And (oVBMethod.Type = vbext_mt_Property) Then
                                lPropertyTypeCount = 3
                            Else
                                lPropertyTypeCount = 1
                            End If
                            
                            For lPropertyTypeIndex = 1 To lPropertyTypeCount
                                
                                On Error Resume Next
                                lMethodStartLine = oCodeModule.ProcStartLine(oVBMethod.Name, evptVBMethodType)
                                If Err.Number = 35 Then 'Sub or Function does not exist
                                    'Let or Set does not exist
                                    Err.Clear
                                    GoTo NextPropertyType
                                ElseIf Err.Number <> 0 Then
                                    GoTo ErrorTrap
                                End If
                                On Error GoTo ErrorTrap
                                lMethodLineCount = oCodeModule.ProcCountLines(oVBMethod.Name, evptVBMethodType)
                                Dim bSplittedLineStarted As Boolean 'Lines ending with _ are splitted ones
                                Dim bSelectStatementStarted As Boolean
                                Dim bThisIsCaseStatement As Boolean
                                Dim sTrimmedLine As String
                                
                                bSplittedLineStarted = False
                                bSelectStatementStarted = True
                                bThisIsCaseStatement = False
                                
                                Dim lMethodLineIndexStart As Long
                                Dim lMethodLineIndexStop As Long
                                
                                If vblnAddOrRemove = True Then
                                    'Find the last line of method
                                    'Look for the non blank/non comment line from the end of the procedure
                                    Dim lMethodActualLastLineNumber As Long
                                    For lMethodActualLastLineNumber = (lMethodStartLine + lMethodLineCount - 1) To (lMethodStartLine + 1) Step -1
                                        sTrimmedLine = Trim(oCodeModule.Lines(lMethodActualLastLineNumber, 1))
                                        If (sTrimmedLine <> vbNullString) _
                                            And (Left(sTrimmedLine, 1) <> "'") _
                                            And (Right(sTrimmedLine, 1) <> "_") _
                                            And (Left(sTrimmedLine, 1) <> "#") _
                                            Then
                                            Exit For
                                        End If
                                    Next lMethodActualLastLineNumber
                                    
                                    'Find the method start line
                                    Dim lMethodActualStartLineNumber As Long
                                    For lMethodActualStartLineNumber = lMethodStartLine To lMethodActualLastLineNumber - 1
                                        sTrimmedLine = Trim(oCodeModule.Lines(lMethodActualStartLineNumber, 1))
                                        If (sTrimmedLine <> vbNullString) _
                                            And (Left(sTrimmedLine, 1) <> "'") _
                                            And (Right(sTrimmedLine, 1) <> "_") _
                                            And (Left(sTrimmedLine, 1) <> "#") _
                                            Then
                                            Exit For
                                        End If
                                    Next lMethodActualStartLineNumber
                                    
                                    lMethodLineIndexStart = lMethodActualStartLineNumber + 1
                                    lMethodLineIndexStop = lMethodActualLastLineNumber - 1
                                    
                                Else
                                    lMethodLineIndexStart = 1
                                    lMethodLineIndexStop = oCodeModule.CountOfLines
                                End If
                                                                
                                For lMethodLineIndex = lMethodLineIndexStart To lMethodLineIndexStop
                                    sMethodLine = oCodeModule.Lines(lMethodLineIndex, 1)
                                    sTrimmedLine = Trim(sMethodLine)
                                    If (sTrimmedLine <> vbNullString) And (Not (Left(sTrimmedLine, 1)) = "'") And (Not (Left(sTrimmedLine, 1) = "#")) And (Not bSplittedLineStarted) Then
                                        
                                        'Check if line number already exist
                                        sOriginalLineNumber = GetNextWord(sMethodLine, 1, , lFirstWordStart)
                                        lFirstWordLen = Len(sOriginalLineNumber)
                                        bLineNumberExist = False
                                        bManualLineNumber = False
                                        If lFirstWordLen <> 0 Then
                                            If (Left(sOriginalLineNumber, 1) <> "'") Then   'If not comment
                                                If (Right(sOriginalLineNumber, 1) = ":") Then 'Remove last colon
                                                    If lFirstWordLen <> 1 Then
                                                        sOriginalLineNumber = Left(sOriginalLineNumber, lFirstWordLen - 1)
                                                        bManualLineNumber = True
                                                    End If
                                                End If
                                                If IsNumeric(sOriginalLineNumber) Then
                                                    bLineNumberExist = True
                                                End If
                                            End If
                                        End If
                                        
                                        If InStr(1, sTrimmedLine, "Select Case", vbTextCompare) = 1 Then
                                            bSelectStatementStarted = True
                                        ElseIf bSelectStatementStarted = True Then
                                            If InStr(1, sTrimmedLine, "End Select", vbTextCompare) = 1 Then
                                                bSelectStatementStarted = False
                                            End If
                                        End If
                                        
                                        If bSelectStatementStarted = True Then
                                            If InStr(1, sTrimmedLine, "Case", vbTextCompare) = 1 Then
                                                bThisIsCaseStatement = True
                                            Else
                                                bThisIsCaseStatement = False
                                            End If
                                        Else
                                            bThisIsCaseStatement = False
                                        End If
                                        
                                        'VB bug: For API declarations in code VB returns lMethodStartLine=1 and lMethodLineCount=true line num!!!
                                        If Not ((lMethodStartLine = 1) And (GetMethodCodeLocation(oVBMethod) <> 1)) Then
                                            If vblnAddOrRemove = True Then 'Add
                                                If bLineNumberExist And (Not bManualLineNumber) And (Not bThisIsCaseStatement) Then    'Remove prev ones
                                                    sMethodLine = Mid(sMethodLine, lFirstWordStart + lFirstWordLen + 1)
                                                End If
                                                
                                                If ((Not bLineNumberExist) Or (bLineNumberExist And (Not bManualLineNumber))) And (Not bThisIsCaseStatement) Then
                                                    Call oCodeModule.ReplaceLine(lMethodLineIndex, lMethodLineIndex & " " & sMethodLine)
                                                End If
                                            Else                    'Remove
                                                If bLineNumberExist And (Not bManualLineNumber) And (Not bThisIsCaseStatement) Then
                                                    Call oCodeModule.ReplaceLine(lMethodLineIndex, Mid(sMethodLine, lFirstWordStart + lFirstWordLen + 1))
                                                End If
                                            End If
                                        End If
                                    End If
                                    If Right(sTrimmedLine, 1) = "_" Then
                                        bSplittedLineStarted = True
                                    Else
                                        bSplittedLineStarted = False
                                    End If
                                Next lMethodLineIndex
NextPropertyType:
                                Select Case evptVBMethodType
                                    Case vbext_pk_Get
                                        evptVBMethodType = vbext_pk_Let
                                    Case vbext_pk_Let
                                        evptVBMethodType = vbext_pk_Set
                                    Case vbext_pk_Set
                                        evptVBMethodType = vbext_pk_Get
                                End Select
                                
                            Next lPropertyTypeIndex
                        End If
                    End If
                    lVBMethodIndex = lVBMethodIndex + 1
                Loop
            End If
ResumeNextComponent:
        Next oVBComponent
    Next oVBProject
    
    If sReadOnlyModuleNames <> vbNullString Then
        Err.Raise 1000, , "Line numbers can not be add to/removed from following modules because they are Read Only: " & vbCrLf & sReadOnlyModuleNames
    End If
    
Exit Sub
ErrorTrap:
    If Err.Number = 40198 Then  'Can not edit module
        sReadOnlyModuleNames = IIf(sReadOnlyModuleNames <> vbNullString, ", ", vbNullString) & sReadOnlyModuleNames & oVBComponent.Name
        Resume ResumeNextComponent
    Else
        MsgBox "Error " & Err.Number & ": " & Err.Description
    End If
End Sub

Private Function GetMethodStartLine(ByVal voCodeModule As CodeModule, ByVal voCodeMember As Member) As String
    Dim sMethodDefinationLine As String
    Dim lSearchStartLineOffset As Long
    lSearchStartLineOffset = 0
    Do
        sMethodDefinationLine = Trim(voCodeModule.Lines(voCodeMember.CodeLocation + lSearchStartLineOffset, 1))
        lSearchStartLineOffset = lSearchStartLineOffset + 1
    Loop While ((sMethodDefinationLine = vbNullString) Or (Left(sMethodDefinationLine, 1) = "'"))
    
    GetMethodStartLine = sMethodDefinationLine
End Function

Public Sub AddErrorHandlerToProjects(ByVal voVBInstance As VBIDE.VBE)
    
    On Error GoTo ErrorTrap
    
    Dim oVBProject As VBIDE.VBProject
    Dim oVBComponent As VBIDE.VBComponent
    Dim oCodeModule As VBIDE.CodeModule
    Dim oVBMethod As VBIDE.Member
    Dim lMethodStartLine As Long
    Dim lMethodLineCount As Long
    Dim lMethodLineIndex As Long
    Dim sOriginalLineNumber As String
    Dim lFirstWordStart As Long
    Dim lFirstWordLen As Long
    Dim sMethodLine As String
    Dim bLineNumberExist As Boolean
    Dim evptVBMethodType As vbext_ProcKind
    Dim bValidMethod As Boolean
    Dim bManualLineNumber As Boolean
    Dim sReadOnlyModuleNames As String
    
    sReadOnlyModuleNames = vbNullString
    For Each oVBProject In voVBInstance.VBProjects
        For Each oVBComponent In oVBProject.VBComponents
            Set oCodeModule = oVBComponent.CodeModule
            If Not (oCodeModule Is Nothing) Then
                For Each oVBMethod In oCodeModule.Members
                    If (oVBMethod.Type = vbext_mt_Method) Or (oVBMethod.Type = vbext_mt_Property) Then
                        bValidMethod = True
                        If oVBMethod.Type <> vbext_mt_Property Then
                            evptVBMethodType = vbext_pk_Proc
                        Else
                            Dim sMethodDefinationLine As String
                            Dim oWords As New Collection
                            Dim lSearchStartLineOffset As Long
                            lSearchStartLineOffset = 0
                            Do
                                sMethodDefinationLine = Trim(oCodeModule.Lines(oVBMethod.CodeLocation + lSearchStartLineOffset, 1))
                                lSearchStartLineOffset = lSearchStartLineOffset + 1
                            Loop While ((sMethodDefinationLine = vbNullString) Or (Left(sMethodDefinationLine, 1) = "'"))  'Skipp comment and blank lines
                            
                            Call MakeWordList(sMethodDefinationLine, oWords)
                            If oWords.Count >= 2 Then   'Check the second word
                                Select Case LCase(oWords(2))
                                    Case "get"
                                        evptVBMethodType = vbext_pk_Get
                                    Case "let"
                                        evptVBMethodType = vbext_pk_Let
                                    Case "set"
                                        evptVBMethodType = vbext_pk_Set
                                    Case Else   'Check the 3rd word
                                        If oWords.Count >= 3 Then
                                            Select Case LCase(oWords(3))
                                                Case "get"
                                                    evptVBMethodType = vbext_pk_Get
                                                Case "let"
                                                    evptVBMethodType = vbext_pk_Let
                                                Case "set"
                                                    evptVBMethodType = vbext_pk_Set
                                                Case Else
                                                    bValidMethod = False
                                            End Select
                                        Else
                                            bValidMethod = False
                                        End If
                                End Select
                            Else
                                bValidMethod = False
                            End If
                            Set oWords = Nothing
                        End If
                        
                        If bValidMethod Then
                            
                            'Due to bug in VB, Members Collection does not includes Let if there also exist Get property. Here's the worka around
                            Dim lPropertyTypeIndex As Long
                            Dim lPropertyTypeCount As Long
                            
                            If ((evptVBMethodType = vbext_pk_Get) Or (evptVBMethodType = vbext_pk_Let) Or (evptVBMethodType = vbext_pk_Set)) And (oVBMethod.Type = vbext_mt_Property) Then
                                lPropertyTypeCount = 3
                            Else
                                lPropertyTypeCount = 1
                            End If
                            
                            For lPropertyTypeIndex = 1 To lPropertyTypeCount
                                
                                Dim sPropertyTypeString As String
                                If oVBMethod.Type = vbext_mt_Property Then
                                    Select Case evptVBMethodType
                                        Case vbext_pk_Get
                                            sPropertyTypeString = "Get_"
                                        Case vbext_pk_Let
                                            sPropertyTypeString = "Let_"
                                        Case vbext_pk_Set
                                            sPropertyTypeString = "Set_"
                                        Case Else
                                            sPropertyTypeString = vbNullString
                                    End Select
                                Else
                                    sPropertyTypeString = vbNullString
                                End If
                                
                                On Error Resume Next
                                lMethodStartLine = oCodeModule.ProcStartLine(oVBMethod.Name, evptVBMethodType)
                                If Err.Number = 35 Then 'Sub or Function does not exist
                                    'Let or Set does not exist
                                    Err.Clear
                                    GoTo NextPropertyType
                                ElseIf Err.Number <> 0 Then
                                    GoTo ErrorTrap
                                End If
                                On Error GoTo ErrorTrap
                                lMethodLineCount = oCodeModule.ProcCountLines(oVBMethod.Name, evptVBMethodType)
                                Dim bSplittedLineStarted As Boolean 'Lines ending with _ are splitted ones
                                Dim bSelectStatementStarted As Boolean
                                Dim bThisIsCaseStatement As Boolean
                                Dim sTrimmedLine As String
                                
                                bSplittedLineStarted = False
                                bSelectStatementStarted = True
                                bThisIsCaseStatement = False
                                
                                'Find the last line of method
                                'Look for the non blank/non comment line from the end of the procedure
                                Dim lMethodActualLastLineNumber As Long
                                For lMethodActualLastLineNumber = (lMethodStartLine + lMethodLineCount - 1) To (lMethodStartLine + 1) Step -1
                                    sTrimmedLine = Trim(oCodeModule.Lines(lMethodActualLastLineNumber, 1))
                                    If (sTrimmedLine <> vbNullString) _
                                        And (Left(sTrimmedLine, 1) <> "'") _
                                        And (Right(sTrimmedLine, 1) <> "_") _
                                        And (Left(sTrimmedLine, 1) <> "#") _
                                        Then
                                        Exit For
                                    End If
                                Next lMethodActualLastLineNumber
                                
                                'Find the method start line
                                Dim lMethodActualStartLineNumber As Long
                                For lMethodActualStartLineNumber = lMethodStartLine To lMethodActualLastLineNumber - 1
                                    sTrimmedLine = Trim(oCodeModule.Lines(lMethodActualStartLineNumber, 1))
                                    If (sTrimmedLine <> vbNullString) _
                                        And (Left(sTrimmedLine, 1) <> "'") _
                                        And (Right(sTrimmedLine, 1) <> "_") _
                                        And (Left(sTrimmedLine, 1) <> "#") _
                                        Then
                                        Exit For
                                    End If
                                Next lMethodActualStartLineNumber
                                
                                'Check if error handler already exists OR refrence to Err.Number or "'#NO_ERROR_HANDLER" is there
                                Dim bDontPutErrorHandler As Boolean
                                Dim bEmptyMethod As Boolean
                                Dim lMethodActualLineCount As Long
                                Dim bContinueChecks As Boolean
                                Dim bMethodMayContainCalls As Boolean
                                
                                bDontPutErrorHandler = False
                                bEmptyMethod = True
                                bContinueChecks = True
                                bMethodMayContainCalls = False
                                
                                lMethodActualLineCount = 0
                                For lMethodLineIndex = (lMethodActualStartLineNumber + 1) To (lMethodActualLastLineNumber - 1)
                                    sMethodLine = Trim(oCodeModule.Lines(lMethodLineIndex, 1))
                                    If UCase(sMethodLine) = "'#USE_ERROR_HANDLER" Then
                                        bDontPutErrorHandler = False
                                        bContinueChecks = False
                                    ElseIf UCase(sMethodLine) = "'#NO_ERROR_HANDLER" Then
                                        bDontPutErrorHandler = True
                                        bContinueChecks = False
                                    End If
                                    
                                    If (sMethodLine <> vbNullString) And (Left(sMethodLine, 1) <> "'") Then
                                        bEmptyMethod = False
                                        lMethodActualLineCount = lMethodActualLineCount + 1
                                        If Not bMethodMayContainCalls Then
                                            If ((InStr(1, sMethodLine, "(") <> 0) And (InStr(1, sMethodLine, ")") <> 0)) Or (InStr(1, sMethodLine, "call ") <> 0) Then
                                                bMethodMayContainCalls = True
                                            End If
                                        End If
                                        If bContinueChecks Then
                                            If InStr(1, sMethodLine, "err.", vbTextCompare) <> 0 Then
                                                bDontPutErrorHandler = True
                                                'Don't exit for loop. See next lines for commands.
                                            End If
                                            'Remove line number if any
                                            Dim oFirstWords As New Collection
                                            Call MakeWordList(sMethodLine, oFirstWords)
                                                'If first and second word is On Error
                                                If oFirstWords.Count >= 2 Then
                                                    If (oFirstWords(1) = "On") And (oFirstWords(2) = "Error") Then
                                                        bDontPutErrorHandler = True
                                                    Else
                                                        If oFirstWords.Count >= 3 Then
                                                            'First is line number and then On Error
                                                            If (oFirstWords(2) = "On") And (oFirstWords(3) = "Error") Then
                                                                bDontPutErrorHandler = True
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            Set oFirstWords = Nothing
                                        End If
                                    End If
                                Next lMethodLineIndex
                                
                                'VB bug: For API declarations in code VB returns lMethodStartLine=1 and lMethodLineCount=true line num!!!
                                If Not ((lMethodStartLine = 1) And (GetMethodCodeLocation(oVBMethod) <> 1)) Then
                                    If (Not bDontPutErrorHandler) And (Not bEmptyMethod) Then
                                        'No error handlers for simple property procedures to improve performance and avoide resetting of errors in error handler code when it tries to access properties
                                        If Not (((evptVBMethodType = vbext_pk_Get) Or (evptVBMethodType = vbext_pk_Let) Or (evptVBMethodType = vbext_pk_Set)) And (lMethodActualLineCount = 1) And (Not bMethodMayContainCalls)) Then
                                            'Insert start
                                            Call oCodeModule.InsertLines(lMethodActualStartLineNumber + 1, vbCrLf & "    Const sMETHOD_NAME as String = " & sQUOTE & oVBProject.Name & "." & oVBComponent.Name & "." & sPropertyTypeString & oVBMethod.Name & sQUOTE & vbCrLf & "    On Error Goto ErrorTrap" & vbCrLf)
                                            
                                            'update last line number
                                            lMethodActualLastLineNumber = lMethodActualLastLineNumber + oCodeModule.ProcCountLines(oVBMethod.Name, evptVBMethodType) - lMethodLineCount
                                            
                                            'Parse the last line to get what the type of method
                                            Dim sVBMethodType As String
                                            Call MakeWordList(oCodeModule.Lines(lMethodActualLastLineNumber, 1), oWords)
                                            If oWords.Count = 2 Then 'Proper end line
                                                sVBMethodType = oWords(2)
                                                Dim sErrorHandlerCode As String
                                                sErrorHandlerCode = vbCrLf & "Exit " & sVBMethodType & vbCrLf & "ErrorTrap:" & vbCrLf
                                                If InStr(1, oVBMethod.Name, "_") = 0 Then
                                                    sErrorHandlerCode = sErrorHandlerCode & "    Call HandleError(sMETHOD_NAME, " & sQUOTE & oVBProject.Name & "." & oVBComponent.Name & sQUOTE & ")"
                                                Else
                                                    sErrorHandlerCode = sErrorHandlerCode & "    Call HandleError(sMETHOD_NAME, " & oVBProject.Name & "." & oVBComponent.Name & ", *Put here custom error code for event*)"
                                                End If
                                                Call oCodeModule.InsertLines(lMethodActualLastLineNumber, sErrorHandlerCode)
                                            End If
                                        End If
                                    End If
                                End If
 
NextPropertyType:
                                Select Case evptVBMethodType
                                    Case vbext_pk_Get
                                        evptVBMethodType = vbext_pk_Let
                                    Case vbext_pk_Let
                                        evptVBMethodType = vbext_pk_Set
                                    Case vbext_pk_Set
                                        evptVBMethodType = vbext_pk_Get
                                End Select
                                
                            Next lPropertyTypeIndex
                        End If
                    End If
                Next oVBMethod
            End If
ResumeNextComponent:
        Next oVBComponent
    Next oVBProject
    
    If sReadOnlyModuleNames <> vbNullString Then
        Err.Raise 1000, , "Line numbers can not be add to/removed from following modules because they are Read Only: " & vbCrLf & sReadOnlyModuleNames
    End If
    
Exit Sub
ErrorTrap:
    If Err.Number = 40198 Then  'Can not edit module
        sReadOnlyModuleNames = IIf(sReadOnlyModuleNames <> vbNullString, ", ", vbNullString) & sReadOnlyModuleNames & oVBComponent.Name
        Resume ResumeNextComponent
    Else
        MsgBox "Error " & Err.Number & ": " & Err.Description
    End If
End Sub

Private Function GetMethodCodeLocation(ByVal voVBMethod As Object) As Long
    
    Dim lGetMethodCodeLocation As Long
        
    On Error Resume Next
    lGetMethodCodeLocation = voVBMethod.CodeLocation
    
    If Err.Number = 424 Then
        GetMethodCodeLocation = 0
    ElseIf Err.Number = 0 Then
        GetMethodCodeLocation = lGetMethodCodeLocation
    Else
        Dim lErrorNumber As Long
        Dim sErrorDescription As String
        
        lErrorNumber = Err.Number
        sErrorDescription = Err.Source
        
        On Error GoTo ErrorTrap
        Err.Raise Err.Number, , Err.Description
    End If
    
Exit Function
ErrorTrap:
    ReRaiseError
End Function
