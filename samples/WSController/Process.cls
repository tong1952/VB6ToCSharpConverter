VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Process"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mobjController As Controller

Implements TalentProcessLib.TalentProcess

Public Property Get PID() As Long
Attribute PID.VB_Description = "Returns the PID of this process"

' Returns the PID of this process

' Returns
' --------------------------------------------------------
' The return value is a ...

    Const PROC_NAME = "PID"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String

    On Error GoTo Error

    bRaiseError = False

    PID = GetCurrentProcessId()

    GoTo Done

Error:
    sRaiseErrorText = Err.Description
    bRaiseError = True
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Call EH_RaiseError(OLEERR_PROPERTY_GET_FAILED, TypeName(Me), PROC_NAME, sRaiseErrorText)
    End If

    Exit Property

End Property

Public Property Get StatusAsText() As String
Attribute StatusAsText.VB_Description = "Returns a status information string"

' Returns a status information string

' Returns
' --------------------------------------------------------
' The return value is a ...

    Const PROC_NAME = "StatusAsText"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim fMean As Single, fMax As Single, fMin As Single, fStdDev As Single, fFreq As Single
    
    On Error GoTo Error

    bRaiseError = False

    If mobjController.Stats Is Nothing Then
        StatusAsText = LoadResString(ridRunning)
    Else
        With mobjController.Stats
            StatusAsText = LoadResString(ridRunningAt _
                                        , Format(1 / .Mean, "0.00") _
                                        , Format(.Minimum, "0.000") _
                                        , Format(.Maximum, "0.000") _
                                        , Format(.StdDev, "0.000000"))
        End With
    End If
                    

    GoTo Done

Error:
    StatusAsText = LoadResString(ridRunning)
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Call EH_RaiseError(OLEERR_PROPERTY_GET_FAILED, TypeName(Me), PROC_NAME, sRaiseErrorText)
    End If

    Exit Property

End Property

Public Sub Hide()
Attribute Hide.VB_Description = "Hides the program form(s)"

' Hides the program form(s)

    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String

    On Error GoTo Error

    bRaiseError = False


    GoTo Done

Error:
    sRaiseErrorText = Err.Description
    bRaiseError = True
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = "Process"
        EH.Procedure = "Hide"
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Call EH_RaiseError(OLEERR_METHOD_FAILED, "Process", "Hide", sRaiseErrorText)
    End If

    Exit Sub

End Sub

Public Sub Show()
Attribute Show.VB_Description = "Displays the program form"

' Displays the program form

    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String

    On Error GoTo Error

    bRaiseError = False


    GoTo Done

Error:
    sRaiseErrorText = Err.Description
    bRaiseError = True
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = "Process"
        EH.Procedure = "Show"
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Call EH_RaiseError(OLEERR_METHOD_FAILED, "Process", "Show", sRaiseErrorText)
    End If

    Exit Sub

End Sub

Public Property Get Visible() As Boolean
Attribute Visible.VB_Description = "Sets/returns the visibility"

' Sets/returns the visibility

' Returns
' --------------------------------------------------------
' True
' False

    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String

    On Error GoTo Error

    bRaiseError = False

        Visible = False

    GoTo Done

Error:
    sRaiseErrorText = Err.Description
    bRaiseError = True
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = "Process"
        EH.Procedure = "Visible"
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Call EH_RaiseError(OLEERR_PROPERTY_GET_FAILED, "Process", "Visible", sRaiseErrorText)
    End If

    Exit Property

End Property

Public Property Let Visible(vNewValue As Boolean)

' Sets/returns the visibility

' Part          Description
' ----------------------------------------------------------------------
'
' vNewValue     New value that property value is being set to

    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String

    On Error GoTo Error

    bRaiseError = False


    GoTo Done

Error:
    sRaiseErrorText = Err.Description
    bRaiseError = True
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = "Process"
        EH.Procedure = "Visible"
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Call EH_RaiseError(OLEERR_PROPERTY_SET_FAILED, "Process", "Visible", sRaiseErrorText)
    End If

    Exit Property

End Property

Public Property Get Events() As Events
Attribute Events.VB_Description = "Sets the events object"
Attribute Events.VB_MemberFlags = "40"

' Sets the events object

' Returns
' --------------------------------------------------------
' The return value is an object ...

    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String

    On Error GoTo Error

    bRaiseError = False

    Set Events = mobjController.Events

    GoTo Done

Error:
    sRaiseErrorText = Err.Description
    bRaiseError = True
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = "Process"
        EH.Procedure = "Events"
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Call EH_RaiseError(OLEERR_PROPERTY_GET_FAILED, "Process", "Events", sRaiseErrorText)
    End If

    Exit Property

End Property

Public Property Get Controller() As Controller
Attribute Controller.VB_Description = "Returns the controller object"

' Returns the controller object

' Returns
' --------------------------------------------------------
' The return value is an object ...

    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String

    On Error GoTo Error

    bRaiseError = False

    Set Controller = mobjController

    GoTo Done

Error:
    sRaiseErrorText = Err.Description
    bRaiseError = True
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = "Process"
        EH.Procedure = "Controller"
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Call EH_RaiseError(OLEERR_PROPERTY_GET_FAILED, "Process", "Controller", sRaiseErrorText)
    End If

    Exit Property

End Property

Public Property Get Name() As String

    Name = PROCESS_NAME
    
End Property

Public Property Get Description() As String

    Description = PROCESS_DESCRIPTION
    
End Property

Public Sub Quit()
Attribute Quit.VB_Description = "Stops this program"

' Stops this program

    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String

    On Error GoTo Error

    bRaiseError = False

    If mobjController Is Nothing Then
        End
    Else
        With mobjController.Commands
            .Insert ocShutdown, "Shutdown"
        End With
    End If
    
    GoTo Done

Error:
    sRaiseErrorText = Err.Description
    bRaiseError = True
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = "Process"
        EH.Procedure = "Quit"
        Call EH_MsgBox
    End If
    Resume ErrorExit
    
ErrorExit:
    On Error Resume Next
    mobjController.Shutdown
    GoTo Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Call EH_RaiseError(OLEERR_METHOD_FAILED, "Process", "Quit", sRaiseErrorText)
    End If

    Exit Sub

End Sub

Public Property Get CommandInterface() As Object
Attribute CommandInterface.VB_Description = "Returns a new Interface object to allow external interfaces with this application"


' Returns a new Interface object to allow external interfaces with this
' application

    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String

    On Error GoTo Error

    bRaiseError = False

    Select Case WindTunnel
    Case wtCWT
        Set CommandInterface = New WindspeedCommandsCWT
    Case Else
        Set CommandInterface = New WindspeedCommandsAAWT
    End Select
    
    GoTo Done

Error:
    sRaiseErrorText = Err.Description
    bRaiseError = True
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = "Process"
        EH.Procedure = "CommandInterface"
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Call EH_RaiseError(OLEERR_PROPERTY_GET_FAILED, "Application", "CommandInterface", sRaiseErrorText)
    End If

    Exit Property
        
End Property

Private Sub Class_Initialize()

'

    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim objSOM As SOM
    Dim objSPGM As SharedPropertyGroupManager
    Dim objSPG As SharedPropertyGroup
    Dim bExists As Boolean
    
    On Error GoTo Error

    bRaiseError = False
    Set objSOM = New SOM
    
    On Error Resume Next
    If IsWithinIDE() Then
        On Error Resume Next
        Set mobjController = objSOM.GetObject(SOM_CONTROLLER)
        If Err Then
            Set mobjController = New Controller
            objSOM.AddObject mobjController, SOM_CONTROLLER
        End If
    Else
        On Error Resume Next
        Set mobjController = objSOM.GetObject(SOM_CONTROLLER)
        If Err Then
            On Error Resume Next
            Set objSPGM = New SharedPropertyGroupManager
            Set objSPG = objSPGM.CreatePropertyGroup(SPG_CONSTANTS, 0, 1, bExists)
            If (bExists = False) And (Err = 0) Then
                On Error GoTo Error
                'LogMessage "PRocess Class Initialize - Creating Controller object"
                Set mobjController = New Controller
                'LogMessage "PRocess Class Initialize - Adding Controller object to SOM"
                objSOM.AddObject mobjController, SOM_CONTROLLER
            End If
        End If
    End If
    
    GoTo Done

Error:
'    sRaiseErrorText = Err.Description
'    bRaiseError = True
'    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = "Process"
        EH.Procedure = "Class_Initialize"
        Call EH_MsgBox
'    End If
    Resume Done

Done:
    On Error Resume Next
    Set objSOM = Nothing
    If bRaiseError Then
        On Error GoTo 0
        Call EH_RaiseError(OLEERR_METHOD_FAILED, "Process", "Class_Initialize", sRaiseErrorText)
    End If

    Exit Sub

End Sub

Private Sub Class_Terminate()

    Set mobjController = Nothing
    
End Sub

Private Property Get TalentProcess_Application() As Object

    Set TalentProcess_Application = Controller
    
End Property

Private Property Get TalentProcess_CommandInterface() As Object
    
    Set TalentProcess_CommandInterface = CommandInterface
    
End Property

Private Property Get TalentProcess_Computer() As String

    TalentProcess_Computer = GetComputerName()
    
End Property

Private Property Get TalentProcess_Description() As String
    
    TalentProcess_Description = Description

End Property

Private Property Get TalentProcess_Events() As TalentProcessLib.Events

    Set TalentProcess_Events = Events
    
End Property

Private Sub TalentProcess_Hide()

    Hide
    
End Sub

Private Property Get TalentProcess_Name() As String

    TalentProcess_Name = Name
    
End Property

Private Property Get TalentProcess_PID() As Long

    TalentProcess_PID = PID
    
End Property

Private Sub TalentProcess_Quit()

    Quit
    
End Sub

Private Property Get TalentProcess_Ready() As Boolean

    If mobjController Is Nothing Then
        TalentProcess_Ready = False
    Else
        TalentProcess_Ready = mobjController.Ready
    End If
    
End Property

Private Sub TalentProcess_Show()

    Show
    
End Sub

Private Property Get TalentProcess_StatusAsText() As String

    TalentProcess_StatusAsText = StatusAsText
    
End Property

Private Property Let TalentProcess_Visible(RHS As Boolean)

    Visible = RHS
    
End Property

Private Property Get TalentProcess_Visible() As Boolean

    TalentProcess_Visible = Visible
    
End Property
