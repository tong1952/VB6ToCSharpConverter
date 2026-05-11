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
Public Sub StartAcousticRecord()

' Stops the main fan

    Const PROC_NAME = "StartAcousticRecord"
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
Public Sub SetAcousticTrigger()

' Stops the main fan

    Const PROC_NAME = "SetAcousticTrigger"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long

    On Error GoTo Error

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
Public Sub StopAcousticRecord()

' Stops the main fan

    Const PROC_NAME = "StopAcousticRecord"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long

    On Error GoTo Error

    bRaiseError = False

    With mobjController
        .Interlocks(ILCK_GENERAL).Check
        .Commands.Clear PROC_NAME, .GetInvalidOpcodes(CStr(ocCancelRecord))
        .Commands.Append ocCancelRecord, PROC_NAME
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
Public Sub SendAcousticParameters()

' Stops the main fan

    Const PROC_NAME = "SendAcousticParameters"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long

    On Error GoTo Error

    bRaiseError = False

    With mobjController
        .CheckInterlock ILCK_GENERAL
        .Commands.Clear PROC_NAME, .GetInvalidOpcodes(CStr(ocSendPars))
        .Commands.Append ocSendPars, PROC_NAME
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
Public Sub ResetAcousticCommunications()

' Resets the GFAI communications

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
Public Sub StartAcousticLogging(Optional LogToMsgLog)

    Const PROC_NAME = "StartLogging"
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
Public Sub StopAcousticLogging()

    Const PROC_NAME = "StopLogging"
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
Public Sub InitializeAcoustic()

' Initialize the program

    Const PROC_NAME = "InitializeAcoustic"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    Dim objCommand As Command
    
    On Error GoTo Error

    bRaiseError = False

    mobjController.Commands.Append Opcodes.ocInitialize, PROC_NAME
    mobjStatuses.Reset STS_GFAI_INITIALIZED, False
    
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
Public Property Get AcousticInitialized() As Boolean
    AcousticInitialized = mobjStatuses.GetValue(STS_GFAI_INITIALIZED)
End Property
Private Sub Class_Initialize()

'
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
        EH.Procedure = "Class_Initialize"
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next
    Set objSOM = Nothing
    Set objSPGM = Nothing
    Exit Sub

End Sub
Private Function GetCC(CC As CCNames) As Single

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

Private Sub Class_Terminate()

    Set mobjController = Nothing
    Set mobjConstants = Nothing

End Sub

