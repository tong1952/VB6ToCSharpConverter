head 1.4;
branch ;
access ;
symbols 4%2E00%2E0020:1.4 4%2E00%2E0019:1.3 4%2E00%2E0016:1.2 4%2E00%2E0005:1.1;
locks ; strict;
comment @@;


1.4
date 2025.02.19.21.25.39; author Ryc; state Develop;
branches ;
next 1.3;

1.3
date 2025.02.05.16.13.33; author Ryc; state Develop;
branches ;
next 1.2;

1.2
date 2024.08.20.17.27.17; author Ryc; state Develop;
branches ;
next 1.1;

1.1
date 2024.07.10.03.42.23; author Ryc; state Develop;
branches ;
next ;

ext
@project Z:/Client Facilities/GAC AAWT/Source Code/Applications/GACAAWTacController/GACAAWTacController.pj;
@

desc
@@


1.4
log
@mobjStatuses.Reset moved from Commands class to Controller class event mobjCommands_CommandAdded@
text
@VERSION 1.0 CLASS
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
@


1.3
log
@Removed trailing "/" from URLs. Added command SetAcousticIP.@
text
@a130 1
    mobjStatuses.Reset STS_ACOUSTIC_INITIALIZED, False
@


1.2
log
@Changed to BK Connect protocol via HTTP.@
text
@a154 43
'
'Private Sub ReleaseAcousticControl()
'
'' PRIVATE
'
'    Const PROC_NAME = "ReleaseAcousticControl"
'    Dim bRaiseError As Boolean
'    Dim sRaiseErrorText As String
'    Dim lRaiseErrorNumber As Long
'
'    On Error GoTo Error
'
'    bRaiseError = False
'
'    With mobjController
'        .CheckInterlock ILCK_GENERAL
'        .Commands.Clear PROC_NAME, .GetInvalidOpcodes(CStr(ocReleaseControl))
'        .Commands.Append ocReleaseControl, PROC_NAME
'    End With
'
'    GoTo Done
'
'Error:
'    bRaiseError = True
'    lRaiseErrorNumber = CCommandError(Err.Number)
'    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, PROC_NAME, Err.Description)
'    If (Err.Number And vbObjectError) <> vbObjectError Then
'        EH.Module = TypeName(Me)
'        EH.Procedure = PROC_NAME
'        Call EH_MsgBox
'    End If
'    Resume Done
'
'Done:
'    On Error Resume Next
'    If bRaiseError Then
'        On Error GoTo 0
'        Err.Raise lRaiseErrorNumber, PROCESS_NAME, sRaiseErrorText
'    End If
'
'    Exit Sub
'
'End Sub
d157 1
a157 1
Attribute ResetAcousticCommunications.VB_Description = "Resets communications with BK Connect"
d159 1
a159 1
' Resets communications with BK Connect
d248 2
d251 1
a251 85
'Private Sub SendAcousticParameters()
'
'' PRIVATE
'
'    Const PROC_NAME = "SendAcousticParameters"
'    Dim bRaiseError As Boolean
'    Dim sRaiseErrorText As String
'    Dim lRaiseErrorNumber As Long
'
'    On Error GoTo Error
'
'    bRaiseError = False
'
'    With mobjController
'        .CheckInterlock ILCK_GENERAL
'        .Commands.Clear PROC_NAME, .GetInvalidOpcodes(CStr(ocSendPars))
'        .Commands.Append ocSendPars, PROC_NAME
'    End With
'
'    GoTo Done
'
'Error:
'    bRaiseError = True
'    lRaiseErrorNumber = CCommandError(Err.Number)
'    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, PROC_NAME, Err.Description)
'    If (Err.Number And vbObjectError) <> vbObjectError Then
'        EH.Module = TypeName(Me)
'        EH.Procedure = PROC_NAME
'        Call EH_MsgBox
'    End If
'    Resume Done
'
'Done:
'    On Error Resume Next
'    If bRaiseError Then
'        On Error GoTo 0
'        Err.Raise lRaiseErrorNumber, PROCESS_NAME, sRaiseErrorText
'    End If
'
'    Exit Sub
'
'End Sub
'
'Private Sub SetAcousticTrigger()
'
'' Stops the main fan
'
'    Const PROC_NAME = "SetAcousticTrigger"
'    Dim bRaiseError As Boolean
'    Dim sRaiseErrorText As String
'    Dim lRaiseErrorNumber As Long
'
'    On Error GoTo Error
'
'    bRaiseError = False
'
'    With mobjController
'        .Interlocks(ILCK_GENERAL).Check
'        .Commands.Clear PROC_NAME, .GetInvalidOpcodes(CStr(ocSetTrigger))
'        .Commands.Append ocSetTrigger, PROC_NAME
'    End With
'
'    GoTo Done
'
'Error:
'    bRaiseError = True
'    lRaiseErrorNumber = CCommandError(Err.Number)
'    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, PROC_NAME, Err.Description)
'    If (Err.Number And vbObjectError) <> vbObjectError Then
'        EH.Module = TypeName(Me)
'        EH.Procedure = PROC_NAME
'        Call EH_MsgBox
'    End If
'    Resume Done
'
'Done:
'    On Error Resume Next
'    If bRaiseError Then
'        On Error GoTo 0
'        Err.Raise lRaiseErrorNumber, PROCESS_NAME, sRaiseErrorText
'    End If
'
'    Exit Sub
'
'End Sub
d253 43
a508 44

'
'Private Sub TakeAcousticControl()
'
''
'
'    Const PROC_NAME = "TakeAcousticControl"
'    Dim bRaiseError As Boolean
'    Dim sRaiseErrorText As String
'    Dim lRaiseErrorNumber As Long
'
'    On Error GoTo Error
'
'    bRaiseError = False
'
'    With mobjController
'        .CheckInterlock ILCK_GENERAL
'        .Commands.Clear PROC_NAME, .GetInvalidOpcodes(CStr(ocTakeControl))
'        .Commands.Append ocTakeControl, PROC_NAME
'    End With
'
'    GoTo Done
'
'Error:
'    bRaiseError = True
'    lRaiseErrorNumber = CCommandError(Err.Number)
'    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, PROC_NAME, Err.Description)
'    If (Err.Number And vbObjectError) <> vbObjectError Then
'        EH.Module = TypeName(Me)
'        EH.Procedure = PROC_NAME
'        Call EH_MsgBox
'    End If
'    Resume Done
'
'Done:
'    On Error Resume Next
'    If bRaiseError Then
'        On Error GoTo 0
'        Err.Raise lRaiseErrorNumber, PROCESS_NAME, sRaiseErrorText
'    End If
'
'    Exit Sub
'
'End Sub
@


1.1
log
@Initial revision@
text
@d21 1
a21 1
    AcousticInitialized = mobjStatuses.GetValue(STS_GFAI_INITIALIZED)
d24 4
d79 1
d116 1
d131 1
a131 1
    mobjStatuses.Reset STS_GFAI_INITIALIZED, False
d155 43
a198 43
Public Sub ReleaseAcousticControl()

' Stops the main fan

    Const PROC_NAME = "ReleaseAcousticControl"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long

    On Error GoTo Error

    bRaiseError = False

    With mobjController
        .CheckInterlock ILCK_GENERAL
        .Commands.Clear PROC_NAME, .GetInvalidOpcodes(CStr(ocReleaseControl))
        .Commands.Append ocReleaseControl, PROC_NAME
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

d200 1
d202 1
a202 1
' Resets the GFAI communications
d208 1
a208 1
    
d217 1
a217 1
    
d242 3
a244 1
Public Sub SendAcousticParameters()
d246 1
a246 1
' Stops the main fan
d248 1
a248 1
    Const PROC_NAME = "SendAcousticParameters"
d252 1
d258 3
d263 2
a264 2
        .Commands.Clear PROC_NAME, .GetInvalidOpcodes(CStr(ocSendPars))
        .Commands.Append ocSendPars, PROC_NAME
d266 1
a266 1
    
a290 1
Public Sub SetAcousticTrigger()
d292 85
a376 1
' Stops the main fan
d378 2
a379 4
    Const PROC_NAME = "SetAcousticTrigger"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
d381 1
a381 1
    On Error GoTo Error
a382 34
    bRaiseError = False

    With mobjController
        .Interlocks(ILCK_GENERAL).Check
        .Commands.Clear PROC_NAME, .GetInvalidOpcodes(CStr(ocSetTrigger))
        .Commands.Append ocSetTrigger, PROC_NAME
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

d388 1
a388 1
    
d392 1
a392 1
        
d395 1
a395 1
            
d421 2
a422 1
Public Sub StartAcousticRecord()
d424 1
a424 1
' Stops the main fan
d426 1
a426 1
    Const PROC_NAME = "StartAcousticRecord"
d466 1
d468 2
d474 1
a474 1
    
d478 1
a478 1
        
d480 1
a480 1
            
d505 2
a506 1
Public Sub StopAcousticRecord()
d508 1
a508 1
' Stops the main fan
d510 1
a510 1
    Const PROC_NAME = "StopAcousticRecord"
d521 2
a522 2
        .Commands.Clear PROC_NAME, .GetInvalidOpcodes(CStr(ocCancelRecord))
        .Commands.Append ocCancelRecord, PROC_NAME
d550 1
d592 43
a634 42
Public Sub TakeAcousticControl()

' Stops the main fan

    Const PROC_NAME = "TakeAcousticControl"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long

    On Error GoTo Error

    bRaiseError = False

    With mobjController
        .CheckInterlock ILCK_GENERAL
        .Commands.Clear PROC_NAME, .GetInvalidOpcodes(CStr(ocTakeControl))
        .Commands.Append ocTakeControl, PROC_NAME
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
@
