VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AcousticCommands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text
Private mobjController As Controller
Private mobjConstants As SharedPropertyGroup
Private mobjStatuses As Statuses

Public Property Get AcousticInitialized() As Boolean
    AcousticInitialized = mobjStatuses.GetValue(STS_ACOUSTIC_INITIALIZED)
End Property

Public Property Get AcousticFaulted() As Boolean
    AcousticFaulted = mobjStatuses.GetValue(STS_ACOUSTIC_FAULTED)
End Property

Private Sub Class_Initialize()

    Const PROC_NAME = "Class_Initialize"
    Dim objSOM As SOM
    Dim objSPGM As SharedPropertyGroupManager
    Dim bExists As Boolean
    
    On Error GoTo Error

    Set objSOM = New SOM
    On Error Resume Next
    Set mobjController = objSOM.GetObject(SOM_CONTROLLER)
    If Err Then
        On Error GoTo Error
        Set mobjController = New Controller
        objSOM.AddObject mobjController, SOM_CONTROLLER
    End If
    On Error GoTo Error
    
    Set objSPGM = New SharedPropertyGroupManager
    Set mobjConstants = objSPGM.CreatePropertyGroup(SPG_CONSTANTS, 0, 0, bExists)
    ''Set mobjDataItems = objSPGM.CreatePropertyGroup(SPG_DATA, 0, 0, bExists)
    Set mobjStatuses = New Statuses
    mobjStatuses.UseSPG SPG_STATUSES
            
    GoTo Done

Error:
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next
    Set objSOM = Nothing
    Set objSPGM = Nothing
    Exit Sub

End Sub

Private Sub Class_Terminate()

    Set mobjController = Nothing
    Set mobjConstants = Nothing

End Sub

Private Function GetCC(CC As CCNames) As Single
Attribute GetCC.VB_Description = "Returns a specific control constant"

' Returns a specific control constant

' Returns
' --------------------------------------------------------
' The return value is a ...

    Const PROC_NAME = "GetCC"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String

    On Error GoTo Error

    bRaiseError = False

    GetCC = mobjConstants.PropertyByPosition(CC).Value

    GoTo Done

Error:
    sRaiseErrorText = "Can't get the value of control constant '" & CC & "': " & Err.Description & ". Make sure the controller is initialized"
    bRaiseError = True
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Err.Raise 5, TypeName(Me) & "." & PROC_NAME, sRaiseErrorText
    End If

    Exit Function

End Function

Public Sub InitializeAcoustic()
Attribute InitializeAcoustic.VB_Description = "Initialize the program"

' Initialize the program

    Const PROC_NAME = "InitializeAcoustic"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    Dim objCommand As Command
    
    On Error GoTo Error

    bRaiseError = False

    mobjController.Commands.Append Opcodes.ocInitialize, PROC_NAME
    
    GoTo Done

Error:
    bRaiseError = True
    lRaiseErrorNumber = CCommandError(Err.Number)
    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, PROC_NAME, Err.Description)
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
        Err.Raise lRaiseErrorNumber, PROCESS_NAME, sRaiseErrorText
    End If
    Exit Sub

End Sub

Public Sub ResetAcousticCommunications()
Attribute ResetAcousticCommunications.VB_Description = "Reset communications with BK Connect"

' Reset communications with BK Connect

    Const PROC_NAME = "ResetAcousticCommunications"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long

    On Error GoTo Error

    bRaiseError = False

    With mobjController
        .Commands.Clear PROC_NAME, .GetInvalidOpcodes(CStr(ocResetCommunications))
        .Commands.Append ocResetCommunications, PROC_NAME
    End With

    GoTo Done

Error:
    bRaiseError = True
    lRaiseErrorNumber = CCommandError(Err.Number)
    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, PROC_NAME, Err.Description)
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
        Err.Raise lRaiseErrorNumber, PROCESS_NAME, sRaiseErrorText
    End If

    Exit Sub

End Sub

Public Sub SendAcousticCommand(API As Variant)
Attribute SendAcousticCommand.VB_Description = "Send a BK Connect API (URL) representing a command the web interface understands"
Attribute SendAcousticCommand.VB_MemberFlags = "40"

' Send a BK Connect API (URL) representing a command the web interface understands

    Const PROC_NAME = "SendAcousticCommand"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    Dim oArg As ArgData

    On Error GoTo Error

    bRaiseError = False

    Set oArg = StringArgData(API)
    CheckMethodArgs Me, PROC_NAME, oArg
    
    With mobjController
        .CheckInterlock ILCK_GENERAL
        .Commands.Clear PROC_NAME, .GetInvalidOpcodes(CStr(ocSendCommand))
        .Commands.Append ocSendCommand, PROC_NAME, oArg
    End With

    GoTo Done

Error:
    bRaiseError = True
    lRaiseErrorNumber = CCommandError(Err.Number)
    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, PROC_NAME, Err.Description)
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
        Err.Raise lRaiseErrorNumber, PROCESS_NAME, sRaiseErrorText
    End If

    Exit Sub

End Sub

Public Sub SetAcousticIP(IPaddr As String)
Attribute SetAcousticIP.VB_Description = "Set the target IP address of computer running BK Connect"

' Set the target IP address of computer running BK Connect

    Const PROC_NAME = "SetAcousticIP"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    Dim oArg As ArgData

    On Error GoTo Error

    bRaiseError = False

    Set oArg = StringArgData(IPaddr)
    CheckMethodArgs Me, PROC_NAME, oArg
    
    With mobjController
        .CheckInterlock ILCK_GENERAL
        .Commands.Clear PROC_NAME, .GetInvalidOpcodes(CStr(ocSetIPAddr))
        .Commands.Append ocSetIPAddr, PROC_NAME, oArg
    End With

    GoTo Done

Error:
    bRaiseError = True
    lRaiseErrorNumber = CCommandError(Err.Number)
    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, PROC_NAME, Err.Description)
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
        Err.Raise lRaiseErrorNumber, PROCESS_NAME, sRaiseErrorText
    End If

    Exit Sub

End Sub

Public Sub StartAcousticLogging(Optional LogToMsgLog)
Attribute StartAcousticLogging.VB_Description = "Start logging communications to the diagnostic panel"

' Start logging communications to the diagnostic panel

    Const PROC_NAME = "StartAcousticLogging"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    Dim argMsgLog As ArgData

    On Error GoTo Error

    bRaiseError = False

    Set argMsgLog = BooleanArgData(LogToMsgLog, True)
    mobjController.Commands.Insert ocStartLogging, PROC_NAME, argMsgLog

    GoTo Done

Error:
    bRaiseError = True
    lRaiseErrorNumber = CCommandError(Err.Number)
    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, PROC_NAME, Err.Description)
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next
    Set argMsgLog = Nothing
    If bRaiseError Then
        On Error GoTo 0
        Err.Raise lRaiseErrorNumber, PROCESS_NAME, sRaiseErrorText
    End If

    Exit Sub

End Sub

Public Sub StartAcousticRecorder()
Attribute StartAcousticRecorder.VB_Description = "Start the BK Connect recorder. The Data Recorder Service must have been previously started."

' Start the BK Connect recorder. The Data Recorder Service must have been previously started.

    Const PROC_NAME = "StartAcousticRecorder"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long

    On Error GoTo Error

    bRaiseError = False

    With mobjController
        .Interlocks(ILCK_GENERAL).Check
        .Commands.Clear PROC_NAME, .GetInvalidOpcodes(CStr(ocStartRecord))
        .Commands.Append ocStartRecord, PROC_NAME
    End With
    
    GoTo Done

Error:
    bRaiseError = True
    lRaiseErrorNumber = CCommandError(Err.Number)
    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, PROC_NAME, Err.Description)
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
        Err.Raise lRaiseErrorNumber, PROCESS_NAME, sRaiseErrorText
    End If

    Exit Sub

End Sub

Public Sub StopAcousticLogging()
Attribute StopAcousticLogging.VB_Description = "Stop logging communications to the diagnostic panel"

' Stop logging communications to the diagnostic panel

    Const PROC_NAME = "StopAcousticLogging"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long

    On Error GoTo Error

    bRaiseError = False

    mobjController.Commands.Insert ocStopLogging, PROC_NAME

    GoTo Done

Error:
    bRaiseError = True
    lRaiseErrorNumber = CCommandError(Err.Number)
    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, PROC_NAME, Err.Description)
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
        Err.Raise lRaiseErrorNumber, PROCESS_NAME, sRaiseErrorText
    End If

    Exit Sub

End Sub

Public Sub StopAcousticRecorder()
Attribute StopAcousticRecorder.VB_Description = "Stop the BK Connect recorder. The Data Recorder Service must have been previously started."

' Stop the BK Connect recorder. The Data Recorder Service must have been previously started.

    Const PROC_NAME = "StopAcousticRecorder"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long

    On Error GoTo Error

    bRaiseError = False

    With mobjController
        .Interlocks(ILCK_GENERAL).Check
        .Commands.Clear PROC_NAME, .GetInvalidOpcodes(CStr(ocStopRecord))
        .Commands.Append ocStopRecord, PROC_NAME
    End With
    
    GoTo Done

Error:
    bRaiseError = True
    lRaiseErrorNumber = CCommandError(Err.Number)
    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, PROC_NAME, Err.Description)
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
        Err.Raise lRaiseErrorNumber, PROCESS_NAME, sRaiseErrorText
    End If

    Exit Sub

End Sub

Public Sub SuspendAcoustic()
Attribute SuspendAcoustic.VB_Description = "Suspend the control program"

' Suspend the control program

    Const PROC_NAME = "SuspendAcoustic"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    
    On Error GoTo Error

    bRaiseError = False
    
    With mobjController
        .Commands.Clear PROC_NAME
        .Commands.Append ocSuspend, PROC_NAME
    End With
    
    GoTo Done

Error:
    bRaiseError = True
    lRaiseErrorNumber = CCommandError(Err.Number)
    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, PROC_NAME, Err.Description)
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
        Err.Raise lRaiseErrorNumber, PROCESS_NAME, sRaiseErrorText
    End If

    Exit Sub

End Sub
