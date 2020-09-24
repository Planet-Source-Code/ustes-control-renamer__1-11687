VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   7932
   ClientLeft      =   1740
   ClientTop       =   1548
   ClientWidth     =   8952
   _ExtentX        =   15790
   _ExtentY        =   13991
   _Version        =   393216
   Description     =   "When you change the name of control it finds that name in all modules and changes it."
   DisplayName     =   "Control Name Replacer"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Command Line / Startup"
   LoadBehavior    =   5
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public FormDisplayed          As Boolean
Public VBInstance             As VBIDE.VBE
Dim mcbMenuCommandBar         As Office.CommandBarControl
Public WithEvents MenuHandler As CommandBarEvents          'command bar event handler
Attribute MenuHandler.VB_VarHelpID = -1
Public WithEvents evtControlHook As VBIDE.VBControlsEvents
Attribute evtControlHook.VB_VarHelpID = -1

'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo error_handler
    
    
    'save the vb instance
    Set VBInstance = Application
    
    'this is a good place to set a breakpoint and
    'test various addin objects, properties and methods
    Debug.Print VBInstance.FullName

    If ConnectMode = ext_cm_External Then
        'Used by the wizard toolbar to start this wizard
    Else
        Set mcbMenuCommandBar = AddToAddInCommandBar("Control Renamer")
        'sink the event
        Set Me.MenuHandler = VBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
    End If
  
    If ConnectMode = ext_cm_AfterStartup Then
        If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
            'set this to display the form on connect
        End If
    End If
    
    Set Me.evtControlHook = VBInstance.Events.VBControlsEvents(Nothing, Nothing)

    Exit Sub
    
error_handler:
    
    MsgBox Err.Description
    
End Sub

'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    
    'delete the command bar entry
    mcbMenuCommandBar.Delete
    
    'shut down the Add-In
    If FormDisplayed Then
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "1"
        FormDisplayed = False
    Else
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "0"
    End If
    

End Sub

Private Sub evtControlHook_ItemRenamed(ByVal VBControl As VBIDE.VBControl, ByVal OldName As String, ByVal OldIndex As Long)
    
    Dim intChange As VbMsgBoxResult
    Dim vbC As VBComponent
    Dim i As Integer
    Dim lngStartLine As Long
    Dim lngEndLine As Long
    Dim lngStartCol As Long
    Dim lngEndCol As Long
    Dim bFound As Boolean
    Dim strTemp As String
    On Error Resume Next
    
    
    intChange = MsgBox("You are about to rename a control, would you like to change all references in code?", vbYesNo, "Control Renamer")
    
    If intChange = vbYes Then
    
        For Each vbC In VBInstance.ActiveVBProject.VBComponents
LookAgain:
                bFound = vbC.CodeModule.Find(OldName, lngStartLine, lngStartCol, lngEndLine, lngEndCol)
                If bFound Then
                    strTemp = vbC.CodeModule.Lines(lngStartLine, 1)
                    strTemp = Replace(strTemp, OldName, VBControl.Properties("Name").Value)
                    vbC.CodeModule.ReplaceLine lngStartLine, strTemp
                    lngStartLine = 0
                    lngEndLine = 0
                    lngStartCol = 0
                    lngEndCol = 0
                    GoTo LookAgain
                End If
        Next vbC
    Else
        Exit Sub
    End If
End Sub

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

