VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'TODO: Change name of the project and description, change global constant in module, save prj as different file
'      Change the const in modAddToIni

Implements IDTExtensibility

Dim mcbAddLineNumbers         As Office.CommandBarControl
Dim mcbRemoveLineNumbers         As Office.CommandBarControl
Dim mcbAddErrorHandler         As Office.CommandBarControl
Private WithEvents mceAddLineNumbers As CommandBarEvents          'command bar event handler
Attribute mceAddLineNumbers.VB_VarHelpID = -1
Private WithEvents mceRemoveLineNumbers As CommandBarEvents          'command bar event handler
Attribute mceRemoveLineNumbers.VB_VarHelpID = -1
Private WithEvents mceAddErrorHandler As CommandBarEvents          'command bar event handler
Attribute mceAddErrorHandler.VB_VarHelpID = -1
Public VBInstance             As VBIDE.VBE
Attribute VBInstance.VB_VarHelpID = -1
Private WithEvents moVBProjectsEvents As VBProjectsEvents
Attribute moVBProjectsEvents.VB_VarHelpID = -1

Private Sub IDTExtensibility_OnConnection(ByVal VBInst As Object, ByVal ConnectMode As VBIDE.vbext_ConnectMode, ByVal AddInInst As VBIDE.AddIn, custom() As Variant)
    
    On Error GoTo error_handler
    
    'save the vb instance
    Set VBInstance = VBInst
    Set moVBProjectsEvents = VBInstance.Events.VBProjectsEvents
    
    If ConnectMode = vbext_cm_External Then
        'Used by the wizard toolbar to start this wizard
        'Me.Show
    Else
        'Set mcbAddLineNumbers = AddToAddInCommandBar("IntelliDataField")
        'sink the event
        'Set Me.MenuHandler = VBInst.Events.CommandBarEvents(mcbAddLineNumbers)
        Call AddToCommandBar
    End If
  
    If ConnectMode = vbext_cm_AfterStartup Then
'        If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
            'set this to display the form on connect
            'Me.Show
'        End If
    End If
  
    Exit Sub
    
error_handler:
    
    MsgBox Err.Description
    
End Sub

Private Sub IDTExtensibility_OnDisconnection(ByVal RemoveMode As VBIDE.vbext_DisconnectMode, custom() As Variant)
    
    On Error Resume Next
    
    'delete the command bar entry
    mcbAddLineNumbers.Delete
    mcbRemoveLineNumbers.Delete
    mcbAddErrorHandler.Delete
    
End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
    ' Comment to prevent procedure from being
    ' deleted on compilation.
End Sub

Private Sub IDTExtensibility_OnAddInsUpdate(custom() As Variant)
    ' Comment to prevent procedure from being
    ' deleted on compilation.
End Sub


'this event fires when the menu is clicked in the IDE
Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
    Dim cbMenuCommandBar As Office.CommandBarControl  'command bar object
    Dim cbMenu As Object
  
    On Error GoTo AddToAddInCommandBarErr
    
    'see if we can find the Add-Ins menu
    Set cbMenu = VBInstance.CommandBars("Add-Ins")
    If cbMenu Is Nothing Then
        'not available so we fail
        Exit Function
    End If
    
    'add it to the command bar
    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    'set the caption
    cbMenuCommandBar.Caption = sCaption
    
    Set AddToAddInCommandBar = cbMenuCommandBar
    
    Exit Function
    
AddToAddInCommandBarErr:

End Function


Sub AddToCommandBar()
  On Error GoTo AddToCommandBarErr
  
  'make sure the standard toolbar is visible
  VBInstance.CommandBars(2).Visible = True
  
  'add it to the command bar
  'the following line will add the TabOrder manager to the
  'Standard toolbar to the right of the ToolBox button
  Set mcbAddLineNumbers = VBInstance.CommandBars(2).Controls.Add(1, , , VBInstance.CommandBars(2).Controls.Count)
  'set the caption
  mcbAddLineNumbers.Caption = "Add Line Numbers"
  'copy the icon to the clipboard
  Clipboard.SetData LoadResPicture(101, 0)
  'set the icon for the button
  mcbAddLineNumbers.PasteFace
  
  Set mcbRemoveLineNumbers = VBInstance.CommandBars(2).Controls.Add(1, , , VBInstance.CommandBars(2).Controls.Count)
  'set the caption
  mcbRemoveLineNumbers.Caption = "Remove Line Numbers"
  'copy the icon to the clipboard
  Clipboard.SetData LoadResPicture(102, 0)
  'set the icon for the button
  mcbRemoveLineNumbers.PasteFace
  
  Set mcbAddErrorHandler = VBInstance.CommandBars(2).Controls.Add(1, , , VBInstance.CommandBars(2).Controls.Count)
  'set the caption
  mcbAddErrorHandler.Caption = "Add Error Handler"
  'copy the icon to the clipboard
  Clipboard.SetData LoadResPicture(103, 0)
  'set the icon for the button
  mcbAddErrorHandler.PasteFace
  
  
  'sink the event
  Set mceAddLineNumbers = VBInstance.Events.CommandBarEvents(mcbAddLineNumbers)
  Set mceRemoveLineNumbers = VBInstance.Events.CommandBarEvents(mcbRemoveLineNumbers)
  Set mceAddErrorHandler = VBInstance.Events.CommandBarEvents(mcbAddErrorHandler)
  
  Exit Sub
    
AddToCommandBarErr:
  MsgBox Err.Description
End Sub


Private Sub mceAddErrorHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Call AddErrorHandlerToProjects(VBInstance)
End Sub

Private Sub mceAddLineNumbers_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Call AddLineNumbersToProjects(VBInstance)
End Sub

Private Sub mceRemoveLineNumbers_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Call RemoveLineNumbersFromProjects(VBInstance)
End Sub

Private Sub moVBProjectsEvents_ItemAdded(ByVal VBProject As VBIDE.VBProject)
    If VBProject.Saved Then
        Dim sFileName As String
        
        If VBProject.Collection.FileName <> vbNullString Then
            sFileName = VBProject.Collection.FileName
        Else
            sFileName = VBProject.FileName
        End If
        
        If IsEmpty(gvMRUList) Then
            ReDim gvMRUList(1 To 4)
        End If
        Dim lMRULIndex As Long
        Dim bProjectExistsInMRU As Boolean
        bProjectExistsInMRU = False
        For lMRULIndex = LBound(gvMRUList) To UBound(gvMRUList)
            If gvMRUList(lMRULIndex) & vbNullString = sFileName Then
                bProjectExistsInMRU = True
                Exit For
            End If
        Next lMRULIndex
        
        If Not bProjectExistsInMRU Then
            Dim oRegistry As New clsRegistry
                Dim lMRUNum As Long
                lMRUNum = GetSetting("WizAddIn", "LastValues", "MRUNum", "1")
                Call oRegistry.UpdateKey(HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0\RecentFiles", CStr(lMRUNum), sFileName)
                lMRUNum = lMRUNum + 1
                If lMRUNum > 4 Then lMRUNum = 1
                Call SaveSetting("WizAddIn", "LastValues", "MRUNum", CStr(lMRUNum))
                gvMRUList(lMRUNum) = sFileName
            Set oRegistry = Nothing
        End If
    End If
End Sub
