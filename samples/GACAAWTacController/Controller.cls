VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Controller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text

Private WithEvents mobjCommands As CommandListLib.Commands
Attribute mobjCommands.VB_VarHelpID = -1
Private WithEvents mobjStatuses As TalentProcessLib.Statuses
Attribute mobjStatuses.VB_VarHelpID = -1
Private mobjConstants As ControlLib.ControlConstants
Private mobjGlobalVars As ControlLib.GlobalVars
Private WithEvents mobjParManager As ParManager
Attribute mobjParManager.VB_VarHelpID = -1
Private WithEvents tmrMain As ccrpTimer
Attribute tmrMain.VB_VarHelpID = -1
Private WithEvents tmrRunOnce As ccrpTimer
Attribute tmrRunOnce.VB_VarHelpID = -1
Private tmrElapsed As ccrpStopWatch
Private WithEvents mobjInterlocks As Interlocks
Attribute mobjInterlocks.VB_VarHelpID = -1
Private mobjEvents As TalentProcessLib.Events
Private mobjForm As frmControl
Private mobjSPGConstants As SharedPropertyGroup
Private mobjSPGDataItems As SharedPropertyGroup
'Private mobjBKConnect As SocketPoll
Private mobjLogFile As TalentProcessLib.LogFile
Private mobjAcousticMessages As MailSlot
Private WithEvents mobjAcousticMessageLogFile As LogFile
Attribute mobjAcousticMessageLogFile.VB_VarHelpID = -1

Private WithEvents mobjWinHTTPRequest As WinHttp.WinHttpRequest
Attribute mobjWinHTTPRequest.VB_VarHelpID = -1

Implements ICcrpTimerNotify

Private lID As Long
Private bReady As Boolean
Private bInited As Boolean
Private Const bKeepStats As Boolean = False
Private bRealTime As Boolean
Private mgvdLastCommandTime As Date
Private mvInterlockCheckValue As Variant
Private mdStartTime As Date

Private msAPI() As String
Private msAPIValue() As String
Private msAPIValueDescription() As String
Private msAPIPrefix As String

' Invalid Opcode Collections
Dim mcolInvalidOpcodes As Collection

Private Enum InitTypes
    initAll = 0
    initConstants = 1
    initPars = 2
End Enum

' Module level control constants
Private mccsPriority As String
Private mccfUpdateTime As Single
Private mccfBKConnectTimeout As Single
Private mccsBKConnectAddress As String
Private mcclBKConnectPort As Long
Private Const mccbBKConnectAnsi As Boolean = False
Private mccfKeepAlivePacketInterval As Single
Private mccfMaxTimeBetweenSendPars As Single
Private mccbUseTargetValues As Boolean
Private mcclGetControlTimeout As Long
Private mccbSendParsOnlyIfUnderFCAS As Boolean

' Module level Inputs
'Private mipAirtemp As Single
'Private mipWindspeed As Single
'Private mipWsMode As Single
'Private mipWsAchieved As Single
'Private mipWsTarget As Single
'Private mipRelHumidity As Single
'Private mipYaw As Single
'Private mipYawAchieved As Single
'Private mipYawTarget As Single
'Private mipAbsPress As Single

' Module level Outputs
Private mopAcousticStatus As Single
'Private mopBKConnectState As Single
Private mopAcousticRecNo As Single
'Private mopFCASInControl As Single

' Module level Global variables
Private mgvbLogOutput As Boolean
Private mgvbLogInput As Boolean
Private mgvbLogToMsgLog As Boolean
Private mgvbLogKeepAlive As Boolean

Private mgvlNumReceiveErrors As Long
Private mlDefPriority As Long
Private mgvsLastSentCommand As String
Private mgvbSimulation As Boolean
'Private mgvbWsWasAchieved As Boolean
'Private mgvfWsNowAchievedTgt As Single
'Private mgvbYawWasAchieved As Boolean
'Private mgvfYawNowAchievedTgt As Boolean
Private mgvdLastBKConnectResponseTime As Date
Private mgvsLastBKConnectResponse As String
Private mgvdBKConnectDataLastRequestTime As Date
Private mgvdLastDataStreamLoggedTime As Date
Private mgvdLastSendParsTime As Date

Private mgvsLastBKConnectResponseErrMsg As String
Private mgvlLastBKConnectResponseErrNum As Long
Private mgvsLastBKConnectResponseErrSource As String
Private mgvsLastBKConnectResponseErrInput As String
Private mgvdLastBKConnectResponseErrTime As Date

Private mgvdLastKeepAlivePacketTime As Date

Private Const NUM_API = 25

Private Enum bkapi
    bkapiLoadProject = 1
    bkapiLoadTemplate = 2
    bkapiStartRecorder = 3
    bkapiStopRecorder = 4
    bkapiResetMonitor = 5
    bkapiManualStart = 6
    bkapiManualStop = 7
    bkapiDeleteLatest = 8
    bkapiSet_RecordingDirectory = 9
    bkapiGet_RecordingDirectory = 10
    bkapiSet_RecordingFileName = 11
    bkapiGet_RecordingFileName = 12
    bkapiGet_RecordingNumber = 13
    bkapiGet_FullRecordingFileName = 14
    bkapiSet_Recording = 15
    bkapiGet_Recording = 16
    bkapiSet_AutoPack = 17
    bkapiGet_AutoPack = 18
    bkapiSet_AutoStore = 19
    bkapiGet_AutoStore = 20
    bkapiSet_Marker = 21
    bkapiGet_Marker = 22
    bkapiAddMarker = 23
    bkapiDeleteMarker = 24
    bkapiSetTrigger = 25
End Enum

Private Sub AbortCurrentCommand(Optional CompletionMessage As String)
    If mobjCommands.Count > 0 Then
        mobjCommands(1).Abort CompletionMessage
    End If
End Sub

Private Sub AddOutput(ParameterName As String, ValueHolder As Single, ResetToDefaultUponComError As Boolean, IsCalculatedBySimulator As Boolean)
    If mgvbSimulation Then
        mobjParManager.AddWithoutFlag(ParameterName, ValueHolder, IIf(IsCalculatedBySimulator, ioInput, ioOutput)).Tag = ResetToDefaultUponComError
    Else
        mobjParManager.AddWithoutFlag(ParameterName, ValueHolder, ioOutput).Tag = ResetToDefaultUponComError
    End If
End Sub

Public Sub CheckInterlock(opcode, Optional CheckValue)

'Checks the specified interlock. If the interlock does not exist, then
' no error is raised.

    Const PROC_NAME = "CheckInterlock"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    Dim sIndex As String
    Dim objInterlock As Interlock

    On Error GoTo Error
    bRaiseError = False
    
    On Error Resume Next
    sIndex = opcode
    Set objInterlock = mobjInterlocks(sIndex)
    If Err Then
        GoTo Done
    End If
    
    mvInterlockCheckValue = CheckValue
    
    On Error GoTo Error
    objInterlock.Check

    GoTo Done
Error:
    sRaiseErrorText = Err.Description
    bRaiseError = True
    lRaiseErrorNumber = Err.Number
    Resume Done
Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Err.Raise lRaiseErrorNumber, TypeName(Me) & "." & PROC_NAME, sRaiseErrorText
    End If
    Exit Sub
End Sub

Private Sub Class_Initialize()
    
    Const PROC_NAME = "Class_Initialize"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim sClassName As String
    Dim i As Long
    
    On Error GoTo Error

    Call Main
    
    bRaiseError = False
    On Error GoTo Error_CreatingObjects
    
    sClassName = "ccrpTimers6.ccrpTimer"
    Set tmrMain = New ccrpTimers6.ccrpTimer
    If App.StartMode <> vbSModeStandalone Then
        Set tmrRunOnce = New ccrpTimers6.ccrpTimer
    End If
        
    On Error GoTo Error
        
    If App.StartMode <> vbSModeStandalone Then
        tmrRunOnce.Interval = 500
        tmrRunOnce.EventType = TimerOneShot
        tmrRunOnce.Enabled = True
    End If
    
    mgvbSimulation = Reg_GetValue(HKEY_LOCAL_MACHINE, REG_ACOUSTIC, , "Simulation", False)
    
'Command List

    ReDim msAPI(NUM_API) As String
    ReDim msAPIValue(NUM_API) As String
    ReDim msAPIValueDescription(NUM_API) As String

    i = 1
    msAPI(i) = "/ProgramService/LoadProject?name=": msAPIValueDescription(i) = "ProjectName": i = i + 1
    msAPI(i) = "/ProgramService/LoadTemplate?name=": msAPIValueDescription(i) = "TemplateName": i = i + 1
    msAPI(i) = "/DataRecorderService/StartRecorder": i = i + 1
    msAPI(i) = "/DataRecorderService/StopRecorder": i = i + 1
    msAPI(i) = "/DataRecorderService/ResetMonitor": i = i + 1
    msAPI(i) = "/DataRecorderService/ManualStart": i = i + 1
    msAPI(i) = "/DataRecorderService/ManualStop": i = i + 1
    msAPI(i) = "/DataRecorderService/DeleteLatest": i = i + 1
    msAPI(i) = "/DataRecorderService/set_RecordingDirectory?value=": msAPIValueDescription(i) = "newPath": i = i + 1
    msAPI(i) = "/DataRecorderService/get_RecordingDirectory": i = i + 1
    msAPI(i) = "/DataRecorderService/set_RecordingFileName?value=":  msAPIValueDescription(i) = "newPath": i = i + 1
    msAPI(i) = "/DataRecorderService/get_RecordingFileName": i = i + 1
    msAPI(i) = "/DataRecorderService/get_RecordingNumber": i = i + 1
    msAPI(i) = "/DataRecorderService/get_FullRecordingFileName": i = i + 1
    msAPI(i) = "/DataRecorderService/set_Recording?value=":  msAPIValueDescription(i) = "newPath": i = i + 1
    msAPI(i) = "/DataRecorderService/get_Recording": i = i + 1
    msAPI(i) = "/DataRecorderService/set_AutoPack?value=":  msAPIValueDescription(i) = "true/false": i = i + 1
    msAPI(i) = "/DataRecorderService/get_AutoPack": i = i + 1
    msAPI(i) = "/DataRecorderService/set_AutoStore?value=":  msAPIValueDescription(i) = "true/false": i = i + 1
    msAPI(i) = "/DataRecorderService/get_AutoStore": i = i + 1
    msAPI(i) = "/DataRecorderService/set_Marker?value=":  msAPIValueDescription(i) = "newName": i = i + 1
    msAPI(i) = "/DataRecorderService/get_Marker": i = i + 1
    msAPI(i) = "/DataRecorderService/AddMarker": i = i + 1
    msAPI(i) = "/DataRecorderService/DeleteMarker": i = i + 1
    msAPI(i) = "/TriggerSetupService/SetTrigger?triggerProperties=":  msAPIValueDescription(i) = "xml": i = i + 1

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

Error_CreatingObjects:
    sRaiseErrorText = "Can't create object '" & sClassName & "': " & Err.Description
    bRaiseError = True
    EH.Module = TypeName(Me)
    EH.Procedure = PROC_NAME
    EH.Item(1) = sClassName
    Call EH_MsgBox
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Call EH_RaiseError(OLEERR_METHOD_FAILED, TypeName(Me), PROC_NAME, sRaiseErrorText)
    End If

    Exit Sub

End Sub

Private Sub Class_Terminate()

    On Error Resume Next
    
'    If Not mobjBKConnect Is Nothing Then
'        mobjBKConnect.Disconnect
'        Set mobjBKConnect = Nothing
'    End If
'
    mobjAcousticMessages.CloseMailslot
    Set mobjAcousticMessages = Nothing
    Set mobjAcousticMessageLogFile.MailSlot = Nothing
    mobjAcousticMessageLogFile.CloseFile
    Set mobjAcousticMessageLogFile = Nothing
    
    Set mobjLogFile.MailSlot = Nothing
    mobjLogFile.CloseFile
    Set mobjLogFile = Nothing
    
    tmrMain.Enabled = False
    Set tmrMain.Notify = Nothing
    Set tmrMain = Nothing
    mobjInterlocks.Quit
    Set mobjInterlocks = Nothing
    Set tmrElapsed = Nothing
    Set mobjCommands = Nothing
    Set mobjConstants = Nothing
    Set mobjGlobalVars = Nothing
    Set mobjEvents = Nothing
    Set mobjStatuses = Nothing
    Set mcolInvalidOpcodes = Nothing
    Set mobjSPGConstants = Nothing
    Set mobjSPGDataItems = Nothing
    mobjParManager.Release
    Set mobjParManager = Nothing
    bInited = False
    Exit Sub

End Sub

Public Property Get Commands() As CommandListLib.Commands
Attribute Commands.VB_Description = "Returns the command list"

' Returns the command list

' Returns
' --------------------------------------------------------
' The return value is an object ...

    Const PROC_NAME = "Commands (Get)"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String

    On Error GoTo Error

    bRaiseError = False

    Set Commands = mobjCommands

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

Public Property Get ControlConstants() As ControlLib.ControlConstants
Attribute ControlConstants.VB_Description = "Returns the collection of control constants for use by an external debugger"

' Returns the collection of control constants for use by an
' external debugger (Look)

' Returns
' --------------------------------------------------------
' The return value is a ControlConstants object

    Const PROC_NAME = "ControlConstants (Get)"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String

    On Error GoTo Error

    bRaiseError = False
    Set ControlConstants = mobjConstants
    
    GoTo Done

Error:
    sRaiseErrorText = Err.Description
    bRaiseError = True
    EH.Module = TypeName(Me)
    EH.Procedure = PROC_NAME
    Call EH_MsgBox
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Call EH_RaiseError(OLEERR_PROPERTY_GET_FAILED, TypeName(Me), PROC_NAME, sRaiseErrorText)
    End If

    Exit Property

End Property

Friend Sub ControlLogic()

    Const PROC_NAME = "ControlLogic"
' This routine is executed at every control interval

    Dim lElapsed As Long    ' Milliseconds
    Dim sData As String
    Dim bExists As Boolean
    Dim fCommSendTime As Single
    Dim i As Long
    Dim sResponse As String
    Dim xByte As Variant
    Dim xPacket As Variant
    Dim lBytesAvailable As Long
    
    On Error GoTo Error

    If bInited = True Then
        mobjParManager.ReadInputs
    End If
    
' Process the responses:
    If mgvbSimulation = False Then
'        If Not mobjBKConnect Is Nothing Then
'            If (mobjBKConnect.State = socConnected) Then
'                For i = 1 To mobjBKConnect.PacketCount
'                    mgvdLastBKConnectResponseTime = Now
'                    sResponse = mobjBKConnect.Packets(i)
'                    ProcessBKConnectResponse sResponse
'                Next i
'                mobjBKConnect.Flush
'            End If
'        End If
    End If
        
    ExecuteCommands
    
    If bInited = True Then
    
' Perform control logic
        AcousticControlLogic
        
' Mode Dependent Monitoring
''        If (mgvbSpeedControl) Then      ' IF speed mode
''            mopVehSpeedSetpt = Ctl_Ramp(mgvfSpeedTarget, mopVehSpeedSetpt, mopVehSpeedRate, mccfUpdateTime, mgvfSpeedIncrement)
''            If mobjSpeedMonitor.MonitorParameter(mopVehSpeed) Then
''                If mobjCommands.Exists(Opcodes.ocSetSpeed) = False Then
''                    mobjStatuses(STS_SPEED_ACHIEVED).SetValue True
''                End If
''            Else
''                mobjStatuses(STS_SPEED_ACHIEVED).SetValue False
''            End If
''        Else
''            mgvfSpeedTarget = mopVehSpeed
''            mopVehSpeedSetpt = mopVehSpeed
''        End If
''
        
' Mode Independent Monitoring
        mobjStatuses(STS_ACOUSTIC_FAULTED).SetValue CBool(mopAcousticStatus)
        
' Special Output Conversion
' ------------------------------------

''        If mobjCommands.Count = 0 Then
''            mopVehForceTarget = mgvfForceTarget
''            mopVehSpeedTarget = mgvfSpeedTarget
''            mopPosTarget = mgvfPosTarget
''        End If
            
' Write control outputs
        WriteOutputs

        
    ElseIf mgvbSimulation = False Then
        On Error Resume Next
'        sData = mobjBKConnect.GetData
        
    End If
    
    If Not mobjStatuses Is Nothing Then
        mobjStatuses.FlushMessages
    End If
    
    GoTo Done

Error:
    If Err.Number And vbObjectError <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Suspend
    Resume Next

Done:
    On Error Resume Next

    Exit Sub

End Sub

Private Function CurrentCommandOpcode() As Long
    If mobjCommands.Count = 0 Then
        CurrentCommandOpcode = 0
    Else
        CurrentCommandOpcode = mobjCommands(1).opcode
    End If
End Function

Public Property Get Events() As Events
    Set Events = mobjEvents
End Property

Private Sub ExecuteCommands()
Attribute ExecuteCommands.VB_Description = "Executes the next command on the list"

    Const PROC_NAME = "ExecuteCommands"
' Executes the next command 1.0 the list

    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    Static ReEntered As Boolean
    Dim objCommand As CommandListLib.Command
    Dim i As Long
    Dim objEH As EH
    Dim bFirstTime As Boolean
    Dim iInt As Integer
    Dim fTargets(1 To 7) As Single
    Dim sAxis(1 To 3) As String
    
    On Error GoTo Error

    bRaiseError = False
    If ReEntered Then
        Exit Sub
    End If
    ReEntered = True

    If mobjCommands.Count = 0 Then
        GoTo Done
    End If

    Set objCommand = mobjCommands(1)
    With objCommand
        Select Case .opcode

' Initialize
        Case ocInitialize
            Me.Initialize , objCommand
            .Remove
            mobjStatuses(STS_ACOUSTIC_INITIALIZED).SetValue bInited
            mobjCommands.Clear "Initialize"
            If bInited Then
                mobjCommands.Insert ocResetCommunications, "ResetAcousticCommunications"
            End If
            
        Case ocSuspend
            .Messages.Flush
            'SendCommand objCommand
            bInited = False
            mobjParManager.ReleaseOutputs
            .Remove
            mobjStatuses(STS_ACOUSTIC_INITIALIZED).SetValue bInited
            SetProcessPriority priNormal
            If IsWithinIDE() Then
                Me.Shutdown
            End If
        
        Case ocShutdown
            .Messages.Flush
            .Remove
            Me.Shutdown
            
        Case ocResetCommunications
            If mgvbSimulation Then
                .Remove
            End If
    
            If Not mobjWinHTTPRequest Is Nothing Then
                mobjWinHTTPRequest.Abort
                Set mobjWinHTTPRequest = Nothing
            End If
                
            Set mobjWinHTTPRequest = New WinHttp.WinHttpRequest
            With mobjWinHTTPRequest
                .SetTimeouts 1000, 1000, 1000, 1000
            End With
            
'            If mobjBKConnect Is Nothing Then
'                Set mobjBKConnect = New TalentProcessLib.SocketPoll
'            End If
'            If (mobjBKConnect.State <> socConnected) Then
'                mobjBKConnect.Disconnect
'                mobjBKConnect.Caption = "BKConnect"
'                Set mobjBKConnect.LogFile = mobjAcousticMessageLogFile
'                mobjBKConnect.logToTalent = False
'                mobjBKConnect.ConnectTimeout = mccfBKConnectTimeout
'                mobjBKConnect.ConnectAsClient mccsBKConnectAddress, CInt(mcclBKConnectPort)
'                If mobjBKConnect.LogMode = logNone Then
'                    mobjAcousticMessageLogFile.LogMessage "Connected to " & mccsBKConnectAddress & " using Port " & mcclBKConnectPort & " as client"
'                End If
'                With mobjBKConnect.BinData
'                    .UseStartCharacter = True
'                    .StartCharacter = Chr(2)
'                    .UseEndCharacter = True
'                    .EndCharacter = Chr(3)
'                End With
'                Pause 1#
'                mobjCommands.Insert ocGetCurrentControl, "GetAcousticCurrentControl"
'                mobjCommands.Insert ocGetCurrentStatus, "GetAcousticCurrentStatus"
'            End If
            .Remove
                   
        Case ocStartLogging
            mobjLogFile.OpenFile "BKConnect", lfAppend, lfBinaryLog, True
            mobjAcousticMessages.OpenOutput
'            mobjBKConnect.LogMode = logNone ' logAll
            mgvbLogInput = True
            mgvbLogOutput = True
            If .Arguments.Count >= 1 Then
                mgvbLogToMsgLog = .Arguments(1).Value
            Else
                mgvbLogToMsgLog = False
            End If
            .Remove
        
        Case ocStopLogging
            On Error Resume Next
            mobjAcousticMessages.CloseMailslot
            mobjLogFile.CloseFile
'            mobjBKConnect.LogMode = logNone
            mgvbLogInput = False
            mgvbLogOutput = False
            .Remove
                                        
        Case Else       ' All other commands require initialized state
            
            CheckInterlock ocInitialize
            
            If mgvbSimulation = False Then
                If (mopAcousticStatus = 0#) Then
                    If .FirstTime Then
                        .ResetTime
                        .InsertPredecessor ocResetCommunications, "ResetAcousticCommunications"
                        GoTo Done
                    ElseIf SecondsSinceTime(.TimeSince) > 10 Then
                        .Abort "No connection to BK Connect"
                    End If
                End If
            End If
            
            Select Case .opcode
                            
'            Case ocSendPars
'                CheckInterlock ILCK_GENERAL
'                .Messages.Flush
'                SendParValues
'                .Remove
'
'            Case ocGetCurrentControl
'                CheckInterlock ILCK_GENERAL
'                SendCommand "GetCurrentControlOwnerRequest"
'                .Remove
'
'            Case ocGetCurrentStatus
'                CheckInterlock ILCK_GENERAL
'                SendCommand "GetAMSStatusRequest"
'                .Remove
'
'            Case ocTakeControl
'                CheckInterlock ILCK_GENERAL
'                If .FirstTime Then
'                    SendCommand "GetControlRequest", Array("timeout", mcclGetControlTimeout)
'                ElseIf mopFCASInControl = 1# Then
'                    .Remove
'                    SendParValues
'                ElseIf .TimeSince >= mcclGetControlTimeout Then
'                    .Abort , cmdTimedout
'                End If
'
'            Case ocReleaseControl
'                CheckInterlock ILCK_GENERAL
'                SendCommand "UnsolicitedControlReleasedEvent"
'                mopFCASInControl = 0#
'                .Remove
                
            Case ocStartRecord
                CheckInterlock ILCK_GENERAL
'                If mopFCASInControl = 0 Then
'                    .InsertPredecessor ocTakeControl, "TakeAcousticControl"
'                Else
'                    .Messages.Flush
'                    SendCommand "StartRecordCommand"
'                    .Remove
'                End If
                
                .Messages.Flush
                SendCommand bkapiStartRecorder
                .Remove
                
'            Case ocSetTrigger
'                CheckInterlock ILCK_GENERAL
'                If mopFCASInControl = 0 Then
'                    .InsertPredecessor ocTakeControl, "TakeAcousticControl"
'                Else
'                    .Messages.Flush
'                    SendCommand "SetTriggerCommand"
'                    .Remove
'                End If
                
            Case ocStopRecord
                CheckInterlock ILCK_GENERAL
'                If mopFCASInControl = 0 Then
'                    .InsertPredecessor ocTakeControl, "TakeAcousticControl"
'                Else
'                    .Messages.Flush
'                    SendCommand "CancelRecordCommand"
'                    .Remove
'                End If
                
                .Messages.Flush
                SendCommand bkapiStopRecorder
                .Remove
            
            Case ocSendCommand
                CheckInterlock ILCK_GENERAL
                .Messages.Flush
                SendCommand .Arguments(1).Value
                .Remove
            
            Case ocSetIPAddr
                CheckInterlock ILCK_GENERAL
                .Messages.Flush
                mccsBKConnectAddress = .Arguments(1).Value ' will be overwritten by an Initialize
                msAPIPrefix = "http://" & mccsBKConnectAddress & ":" & mcclBKConnectPort
                .Remove
            
            Case Else
                .Abort "Unrecognized opcode: " & .opcode
                
            End Select
        End Select
        
    End With
    GoTo Done

Error:
    sRaiseErrorText = Err.Description
    lRaiseErrorNumber = CCommandError(Err.Number)
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume ErrorExit
    
ErrorExit:
    On Error Resume Next
    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, objCommand.Text, sRaiseErrorText)
    objCommand.Abort sRaiseErrorText, cmdAborted
    Resume Done

Done:
    On Error Resume Next
    ReEntered = False
    Exit Sub

End Sub

Public Function GetInvalidOpcodes(opcode) As Variant
    
    Dim sIndex As String
    On Error Resume Next
    sIndex = opcode
    GetInvalidOpcodes = mcolInvalidOpcodes(sIndex)
    If Err Then
        GetInvalidOpcodes = 0
    End If
    Exit Function
    
End Function

Public Function GetShowStatusText() As String
    On Error Resume Next
    mobjInterlocks(CStr(ocShowStatus)).Check
    If Err Then
        GetShowStatusText = LoadResString(stsTheFollowingFaultsExist, Err.Description)
    Else
        GetShowStatusText = LoadResString(stsNoFaultsDetected)
    End If
    Exit Function
End Function

Private Sub AcousticControlLogic()
Attribute AcousticControlLogic.VB_Description = "Closed loop control logic"

' Closed loop control logic

    Const PROC_NAME = "AcousticControlLogic"
    Dim xData() As Byte
    Dim i As Long, sResponse As String
    Dim objPar As ManagedPar
    Dim bExists As Boolean
    Dim fValue As Single
    Dim bCmdReqd As Boolean
    
    On Error GoTo Error
    
    On Error Resume Next
    If mgvbSimulation Then
        mopAcousticStatus = NOFAULT
'    ElseIf mobjBKConnect.State = socConnected Then
'        mopAcousticStatus = NOFAULT
'    Else
'        mopAcousticStatus = NOFAULT
    End If
    If Err Then
        mopAcousticStatus = FAULT
    End If
            
'' Send out the parameter (windspeed, yaw, etc.) values whenever windspeed or yaw or airtemp is achieved, or a maximum time period has elapsed
'    If mgvbSimulation = False Then
'        If Not mobjBKConnect Is Nothing Then
'            If mobjBKConnect.State = soxConnected Then
'                If (mccbSendParsOnlyIfUnderFCAS = False) Or (mopFCASInControl = 1#) Then
'                    If WindspeedNowAchieved Or YawAngleNowAchieved Or ((mccfMaxTimeBetweenSendPars > 0) And (SecondsSinceTime(mgvdLastSendParsTime) >= mccfMaxTimeBetweenSendPars)) Then
'                        On Error Resume Next
'                        If mopBKConnectState <= gsPreview Then
'                            SendParValues
'                        End If
'                        On Error GoTo Error
'                    End If
'                End If
'            End If
'        End If
'    End If

' Send Keepalive packet if connected
    If mgvbSimulation = False Then
        If Not mobjWinHTTPRequest Is Nothing Then
'            If mopAcousticStatus = NOFAULT Then
                If mccfKeepAlivePacketInterval > 0# Then
                    If SecondsSinceTime(mgvdLastKeepAlivePacketTime) >= mccfKeepAlivePacketInterval Then
                        On Error Resume Next
                        SendCommand bkapiGet_RecordingNumber    ' safe, read-only, data item
                        mgvdLastKeepAlivePacketTime = Now
                        On Error GoTo Error
                    End If
                End If
'            End If
        End If
    End If
    
'' Safety monitoring
'    If mopAcousticStatus = 0 Then
'        mopBKConnectState = 0
''        mopFCASInControl = 0
'    End If
'
'' Simulate BKConnect responses
'    If mgvbSimulation Then
'    End If
    
    GoTo Done

Error:
    If bInited Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
        Suspend
    End If
    Resume Done

Done:
    On Error Resume Next

    Exit Sub

End Sub

Public Property Get GlobalVars() As ControlLib.GlobalVars

' Returns the collection of global variables for use by an
' external debugger (Look)

' Returns
' --------------------------------------------------------
' The return value is a GlobalVars object

    Const PROC_NAME = "GlobalVars (Get)"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String

    On Error GoTo Error

    bRaiseError = False
    Set GlobalVars = mobjGlobalVars
    
    GoTo Done

Error:
    sRaiseErrorText = Err.Description
    bRaiseError = True
    EH.Module = TypeName(Me)
    EH.Procedure = PROC_NAME
    Call EH_MsgBox
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Call EH_RaiseError(OLEERR_PROPERTY_GET_FAILED, TypeName(Me), PROC_NAME, sRaiseErrorText)
    End If

    Exit Property

End Property

Private Sub ICcrpTimerNotify_Timer(ByVal Milliseconds As Long)
    
    On Error Resume Next
    ControlLogic
    DoEvents
    Exit Sub
    
End Sub

Friend Sub Initialize(Optional InitType As InitTypes = initAll, _
            Optional Command As Command)

' Perform basic initialization/startup functions

    Const PROC_NAME = "Initialize"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    Dim bStatus As Boolean, sText As String
    Dim sFileName As String
    Dim objEH As EH
    Dim bExists As Boolean
    Dim sPath As String
    Dim sDefaultsFileName As String
    
    On Error GoTo Error

    bRaiseError = False

    Select Case InitType
    Case initConstants, initAll
        bStatus = True
        sPath = GetDefaultCCFolder(App.Path)
        sFileName = sPath & PROCESS_NAME & ".csv"
        sFileName = Reg_GetValue(HKEY_LOCAL_MACHINE, REG_ACOUSTIC, , "ControlConstantsFile", sFileName)
        sDefaultsFileName = App.Path & "\" & App.EXEName & "Defaults.csv"
        With mobjConstants
            .OpenFileWithDefaults sFileName, sDefaultsFileName
            
            .AddString "Priority", mccsPriority, bStatus, sText
            .AddSingle "UpdateInterval", mccfUpdateTime, bStatus, sText
            .AddSingle "BKConnectTimeout", mccfBKConnectTimeout, bStatus, sText
            .AddString "BKConnectAddress", mccsBKConnectAddress, bStatus, sText
            .AddLong "BKConnectPort", mcclBKConnectPort, bStatus, sText
            .AddSingle "KeepAliveInterval", mccfKeepAlivePacketInterval, bStatus, sText
'            .AddSingle "SendParsIntervalMax", mccfMaxTimeBetweenSendPars, bStatus, sText
'            .AddBoolean "SendParsUseTargets", mccbUseTargetValues, bStatus, sText
'            .AddLong "GetControlTimeout", mcclGetControlTimeout, bStatus, sText
'            .AddBoolean "SendParsFCASOnly", mccbSendParsOnlyIfUnderFCAS, bStatus, sText
            
            If bStatus = False Then
                bRaiseError = True
                sRaiseErrorText = "Can't find the following control constants in " & sFileName & ": " & sText
                lRaiseErrorNumber = tceCommandFailed
                If Command Is Nothing Then
                Else
                    Command.SuppressLogging = True
                End If
                Set objEH = New EH
                With objEH
                    .AppError = LoadResString(AE.CantLoadConstants)
                    .Item(1) = sText
                    .FileName = sFileName
                    .Module = TypeName(Me)
                    .Procedure = PROC_NAME
                    LogMessage2 objEH
                End With
                'GoTo Done
            End If
        End With
        
        With mobjSPGConstants
''            .CreatePropertyByPosition(CC_XMin, bExists).Value = mccfXMin
''            .CreatePropertyByPosition(CC_XMax, bExists).Value = mccfXMax
        End With
               
    End Select

    mlDefPriority = GetControllerPriority(mccsPriority)
    
    Select Case InitType
    Case initAll, initPars
        On Error Resume Next
        mobjParManager.RefreshObjects
        If Err Then
            lRaiseErrorNumber = CCommandError(Err.Number)
            sRaiseErrorText = Err.Description
            Set objEH = New EH
            If Command Is Nothing Then
            Else
                Command.SuppressLogging = True
            End If
            With objEH
                Set .VBErr = Err
                .AppError = LoadResString(AE.InitializeFailedParsMissing)
                .Item(1) = Command.Text
                .Module = TypeName(Me)
                .Procedure = PROC_NAME
                If Command Is Nothing Then
                Else
                    .Item(1) = Command.Text
                End If
                LogMessage2 objEH
            End With
            GoTo Done
        ElseIf mgvbSimulation Then
            ' Read initial values from simulation configuration/setup when running under simulation
            ''mopVehGTarget = mobjParManager("RLC_G").Value
        End If
        'mobjParManager.AttachOutputs
        bInited = mobjParManager.Ready
    End Select
    
    msAPIPrefix = "http://" & mccsBKConnectAddress & ":" & mcclBKConnectPort
    
    GoTo Done

Error:
    bInited = False
    lRaiseErrorNumber = CCommandError(Err.Number)
    sRaiseErrorText = Err.Description
    bRaiseError = True
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        If Command Is Nothing Then
        Else
            EH.Item(1) = Command.Text
        End If
        Call EH_MsgBox
        If Command Is Nothing Then
        Else
            Command.SuppressLogging = True
        End If
    End If
    Resume Done

Done:
    On Error Resume Next
    mobjStatuses(STS_ACOUSTIC_INITIALIZED).SetValue bInited
    mobjConstants.CloseFile
    Set objEH = Nothing
    If bInited Then
        SetProcessPriority mlDefPriority
    Else
        SetProcessPriority priNormal
    End If
    If bRaiseError Then
        On Error GoTo 0
        Call EH_RaiseError(lRaiseErrorNumber, TypeName(Me), PROC_NAME, sRaiseErrorText)
    End If

    Exit Sub

End Sub

Public Property Get Initialized() As Boolean
Attribute Initialized.VB_Description = "Returns whether or not the controller is initialized (controlling) or not"

' Returns whether or not the controller is initialized (controlling) or
' not

' Returns
' --------------------------------------------------------
' True
' False

    Const PROC_NAME = "Initialized (Get)"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String

    On Error GoTo Error

    bRaiseError = False

    Initialized = bInited

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
        EH_RaiseError OLEERR_PROPERTY_GET_FAILED, TypeName(Me), PROC_NAME, sRaiseErrorText
    End If

    Exit Property

End Property

Public Property Get Interlocks() As Interlocks
    Set Interlocks = mobjInterlocks
End Property

Private Sub LogMessageEH(MessageText As String, SendingRoutine As String, Optional Items, Optional MessageType As MessageTypes = msgInformation)

    Dim objEH As EH
    Dim sItems() As String
    Dim i As Long
    
    On Error Resume Next
    Set objEH = New EH
    With objEH
        .AppError = MessageText
        If Not IsMissing(Items) Then
            For i = 1 To 10
                Err.Clear
                .Item(i) = Items(i)
                If Err Then
                    Exit For
                End If
            Next i
        End If
        .Module = TypeName(Me)
        .Procedure = SendingRoutine
        .MessageType = MessageType
        LogMessage2 objEH
    End With
    Exit Sub
    
End Sub

Private Sub mobjCommands_CommandAdded(NewCommand As CommandListLib.Command)

    Dim objMessage As TalentProcessLib.Message
    Dim sStatus As String

    On Error Resume Next
    
    If NewCommand.SuccessorKey = "" Then
        If mgvbLogInput Then
            Set objMessage = New TalentProcessLib.Message
            With objMessage
                .Source = "Command"
                .Text = NewCommand.Text & GetCmdArgs(NewCommand.Arguments)
                .MessageType = mtCommand
            End With
            mobjAcousticMessageLogFile.LogMessage2 objMessage
        End If
    End If
        
    sStatus = ""
    
    Select Case NewCommand.opcode
        Case ocInitialize
            sStatus = STS_ACOUSTIC_INITIALIZED
    End Select
    
    If sStatus <> "" Then
        mobjStatuses.Reset sStatus, False
    End If
    
    Exit Sub
    
End Sub

Private Sub mobjCommands_CommandCompleted(Command As CommandListLib.Command, CompletionStatus As CommandListLib.CompletionStatuses, CompletionMessage As String)

'
    Const PROC_NAME = "mobjCommands_CommandCompleted"
    Dim objEH As EH
    Dim objMessage As TalentProcessLib.Message
    Dim i As Long
    
    On Error GoTo Error
    
    Select Case CompletionStatus
    Case cmdAborted, cmdTimedout
        If Command.SuppressLogging = False Then
            Set objEH = New EH
            Select Case Command.opcode
            Case ocResetCommunications
                objEH.AppError = LoadResString(AE.CommunicationsFaulted, Command.Text, CompletionMessage, mccsBKConnectAddress)
            Case Else
                objEH.AppError = AppErrorText(tceCommandFailed, Command.Text, CompletionMessage)
            End Select
            objEH.Module = TypeName(Me)
            objEH.Procedure = PROC_NAME
            Set objEH.Component = App
            LogPopupMessage objEH, Command.Text
            Command.SuppressLogging = True
        ElseIf Command.opcode = ocResetCommunications Then
            Set objEH = New EH
            objEH.AppError = LoadResString(AE.CommunicationsFaulted, Command.Text, CompletionMessage, mccsBKConnectAddress)
            objEH.Module = TypeName(Me)
            objEH.Procedure = PROC_NAME
            Set objEH.Component = App
            LogMessage2 objEH ' Command.Text & " Command failed -- " & CompletionMessage, mtFailure
        
        End If
                
    End Select
    
    On Error Resume Next
    If Command.SuccessorKey = "" Then
        If mgvbLogInput Then
            Set objMessage = New TalentProcessLib.Message
            With objMessage
                .Source = IIf(CompletionStatus = cmdCompleted, "Completed", "Failed")
                .Text = CompletionMessage
                .MessageType = IIf(CompletionStatus = cmdCompleted, mtInformation, mtFailure)
            End With
            mobjAcousticMessageLogFile.LogMessage2 objMessage
        End If
    End If
    
    GoTo Done

Error:
    Resume Done

Done:
    On Error Resume Next
    Set objEH = Nothing
    Exit Sub

End Sub

Private Sub mobjCommands_CommandInserted(NewCommand As CommandListLib.Command, ExistingCommand As CommandListLib.Command)

'

    Const PROC_NAME = "mobjCommands_CommandInserted"
    On Error GoTo Error

'    Select Case NewCommand.Opcode
'    Case ocStartFan
'        NewCommand.Messages.Add "Fan currently 0.0 -- Starting fan", mtInformation
'    Case ocStop
'        NewCommand.Messages.Add "Fan currently 1.0 -- Stopping fan", mtInformation
'    End Select
    
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

    Exit Sub

End Sub

Private Sub mobjCommands_MessageFlushed(MessageText As String, MessageType As Long)

'

    Const PROC_NAME = "mobjCommands_MessageFlushed"
    On Error GoTo Error

    LogMessage MessageText, MessageType

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

    Exit Sub

End Sub

Private Sub mobjAcousticMessageLogFile_MessageLogged(Message As Variant)

    Dim sText As String
    Dim lType As Long
    Dim objMessage1 As Object
    Dim objMessage2 As Object
    Dim sSource As String
    
    On Error Resume Next
    Set objMessage1 = Message
    CopyObject objMessage1, objMessage2
    If Err Then
        mobjLogFile.LogMessage2 Message
    Else
''        sSource = Trim(objMessage2.Source)
''        If Right(sSource, 1) = ":" Then
''            sSource = Left(sSource, Len(sSource) - 1)
''        End If
''        sSource = sSource & " (BKConnect):"
''        objMessage2.Source = sSource
        mobjLogFile.LogMessage2 objMessage2
    End If
    
    Exit Sub

End Sub

Private Sub mobjInterlocks_Evaluate(InterlockName As String, CurrentOKValue As Boolean, AdditionalMessages As Collection)

    
    Dim objCommand As CommandListLib.Command

    On Error Resume Next
    Select Case InterlockName
    Case ILCK_GENERAL
''        If mopBKConnectDataStatus = 1# Then
''            If mopBKConnectRemote <> 1# Then
''                AdditionalMessages.Add LoadResString(ridBKConnectNotInRemote)
''                CurrentOKValue = False
''            End If
''        End If
    End Select
    
    Exit Sub
    
End Sub

Private Sub mobjParManager_BeforeUnloaded(ConfigurationName As String, Cancel As Boolean, Cancellers As Collection)

'

    Const PROC_NAME = "mobjParManager_BeforeUnloaded"

    On Error GoTo Error

    If bInited = True Then
        Cancel = True
        Cancellers.Add PROCESS_NAME
    End If

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

    Exit Sub

End Sub

Private Sub mobjStatuses_StatusChanged(Message As String)
'

    Const PROC_NAME = "mobjStatuses_StatusChanged"

    On Error GoTo Error

''    If mipDCState = 2 Then      ' Don't log speed achieved/not achieved messages during a drive cycle.
''        Select Case Message
''        Case "Vehicle Speed Not Achieved", "Vehicle Speed Achieved"
''
''        Case Else
''            LogMessage Message, mtStatusChange
''        End Select
''    Else
        LogMessage Message, mtStatusChange
''    End If

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

    Exit Sub

End Sub

Public Property Get ParManager() As Object
    Set ParManager = mobjParManager
End Property

Private Sub ProcessBKConnectResponse(ByVal BKConnectResponse As String)
Attribute ProcessBKConnectResponse.VB_Description = "Process a received data packet"

' Process a received TCP/IP response message

' Part          Description
' ----------------------------------------------------------------------
' DataResponse

    Const PROC_NAME = "ProcessBKConnectResponse"
    Static vLastErrorMessageTime As Date
    Dim sItems() As String
    Dim lNumItems As Long
    Dim sSubItems() As String
    Dim lNumSubItems As Long
    Dim sCommand As String
    Dim sFaultText As String
    Dim lErrDescStart As Long
    Dim i As Long
    Dim lMessageType As MessageTypes
    Dim sLogText As String
    
    On Error GoTo Error
    
'    BKConnectResponse = Trim(BKConnectResponse)
'    If Left(BKConnectResponse, 1) = Chr(2) Then      ' leading terminator
'        If Len(BKConnectResponse) < 3 Then
'            GoTo Done
'        End If
'        BKConnectResponse = Mid(BKConnectResponse, 2)
'    End If
'    If Right(BKConnectResponse, 1) = Chr(3) Then     ' ETX
'        BKConnectResponse = Left(BKConnectResponse, Len(BKConnectResponse) - 1)
'    End If
'    mgvsLastBKConnectResponse = BKConnectResponse
'
'    If mgvbLogInput Then
'        sLogText = Replace(BKConnectResponse, vbLf, "")
'        sLogText = Replace(sLogText, vbCr, "")
'        mobjAcousticMessageLogFile.LogMessage sLogText, "Received:", mtInformation
'        If mgvbLogToMsgLog Then
'            LogMessage "Received: " & sLogText, mtInformation
'        End If
'    End If
'
'    If BKConnectResponse Like "*StillAliveReply*" Then
'        ' Do Nothing
'
'    ElseIf BKConnectResponse Like "*ErrorEvent*" Then
'        lErrDescStart = InStr(BKConnectResponse, """errorDescription"": ") + 21
'        For i = lErrDescStart To Len(BKConnectResponse)
'            If Mid(BKConnectResponse, i, 1) = """" Then
'                Exit For
'            End If
'            sFaultText = sFaultText & Mid(BKConnectResponse, i, 1)
'        Next i
'        If BKConnectResponse Like "*""Critical""*" Then
'            lMessageType = msgCritical
'        ElseIf BKConnectResponse Like "*""Error""*" Then
'            lMessageType = msgCritical
'        ElseIf BKConnectResponse Like "*""Warning""*" Then
'            lMessageType = msgExclamation
'        Else
'            lMessageType = msgInformation
'        End If
'        LogMessage sFaultText, lMessageType
'
'    ElseIf BKConnectResponse Like "*ShutdownEvent*" Then
'        ' BKConnect is shutting down
'        mobjBKConnect.Shutdown
'        mopAcousticStatus = 0#
'        mopFCASInControl = 0#
'    ElseIf (BKConnectResponse Like "*GetAMSStatusReply*") Or (BKConnectResponse Like "*AMSStatusChangedEvent*") Then
'        If BKConnectResponse Like "*Uninitialized*" Then
'            mopBKConnectState = gsUninitialized
'        ElseIf BKConnectResponse Like "*Preview*" Then
'            mopBKConnectState = gsPreview
'        ElseIf BKConnectResponse Like "*PreTrigger*" Then
'            mopBKConnectState = gsPreTrigger
'        ElseIf BKConnectResponse Like "*TriggerReady*" Then
'            mopBKConnectState = gsTriggerReady
'        ElseIf BKConnectResponse Like "*Recording*" Then
'            mopBKConnectState = gsRecording
'        ElseIf BKConnectResponse Like "*DataTransfer*" Then
'            mopBKConnectState = gsDataTransfer
'        ElseIf BKConnectResponse Like "*CancelRecord*" Then
'            mopBKConnectState = gsCancelRecord
'        ElseIf BKConnectResponse Like "*PostProcessing*" Then
'            mopBKConnectState = gsPostProcessing
'        ElseIf BKConnectResponse Like "*MeasurementFinished*" Then
'            mopBKConnectState = gsMeasurementFinished
'        Else
'            mopBKConnectState = -1
'        End If
'
'    ElseIf BKConnectResponse Like "*ReleaseControlRequest*" Then
'        SendCommand "ReleaseControlReply", Array("requestControlReply", "RequestGranted")
'        mopFCASInControl = 0#
'
'    ElseIf BKConnectResponse Like "*GetControlReply*" Then
'        If BKConnectResponse Like "*RequestGranted*" Then
'            mopFCASInControl = 1#
'        Else
'            mopFCASInControl = 0#
'            If CurrentCommandOpcode = ocTakeControl Then
'                AbortCurrentCommand "Denied by BKConnect system"
'            End If
'        End If
'
'    ElseIf BKConnectResponse Like "*GetCurrentControlOwnerReply*" Then
'        If BKConnectResponse Like "*You*" Then
'            mopFCASInControl = 1#
'        Else
'            mopFCASInControl = 0#
'        End If
'
'    ElseIf BKConnectResponse Like "*ForceReleaseControlEvent*" Then
'        mopFCASInControl = 0#
'
'
'    ElseIf BKConnectResponse Like "*UnsolicitedGetControlCommand*" Then
'        mopFCASInControl = 1#
'
'    End If
    
    GoTo Done

Error:
    mgvlNumReceiveErrors = mgvlNumReceiveErrors + 1
    mgvsLastBKConnectResponseErrMsg = Err.Description
    mgvlLastBKConnectResponseErrNum = Err.Number
    mgvsLastBKConnectResponseErrSource = Err.Description
    mgvsLastBKConnectResponseErrInput = BKConnectResponse
    
    If (Err.Number And vbObjectError) <> vbObjectError Then
        Set EH.VBErr = Err
        If (SecondsSinceTime(vLastErrorMessageTime) >= 300) Then
            EH.AppError = LoadResString(AE.BKConnectResponseError)
            EH.Item(1) = BKConnectResponse
            EH.Module = TypeName(Me)
            EH.Procedure = PROC_NAME
            LogMessage2 EH
            vLastErrorMessageTime = Now
        End If
        EH.AppError = Null
    End If
    Resume Done

Done:
    On Error Resume Next
    Exit Sub

End Sub

Public Property Get Ready() As Boolean

' Returns
' --------------------------------------------------------
' True
' False
 
    On Error Resume Next
    Ready = bReady
    Exit Property

End Property

'Public Sub SendCommand(CommandToSend As String, ParamArray ParPairs() As Variant)
Private Sub SendCommand(API As Variant, ParamArray ParValues() As Variant)

    Const PROC_NAME = "SendCommand"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    Dim sCmd As String
    Dim sParName As String
    Dim vParValue As Variant
    Dim sParValue As String
    Dim sArrayParName As String
    Dim vArrayParValue As Variant
    Dim i As Long
    Dim j As Long
    Dim xCommand() As Byte
    Dim sLogText As String
    Dim bLog As Boolean
    Dim APIindex As Long
    
    On Error GoTo Error

    bRaiseError = False
    
    bLog = True
    
    If IsNumeric(API) Then
        APIindex = API
        
        If APIindex = bkapiGet_RecordingNumber Then
            ' Normally do not log (repetitive) keep-alive communications
            bLog = mgvbLogKeepAlive
        End If
        
        If msAPIValueDescription(APIindex) <> "" Then
            If IsArray(ParValues) Then
                sParValue = ParValues(0)
                sCmd = msAPIPrefix & msAPI(APIindex) & URLEncode(sParValue)
            Else
                ' error - missing argument
                Err.Raise 5, , "API " & msAPI(APIindex) & " is missing required argument '" & msAPIValueDescription(APIindex) & "'"
            End If
        Else
            sCmd = msAPIPrefix & msAPI(APIindex)
        End If
    Else
        sCmd = API
    End If
    
    mgvsLastSentCommand = sCmd
    mgvdLastCommandTime = Now
    
    If mgvbLogOutput And bLog Then
        sLogText = sCmd
        mobjAcousticMessageLogFile.LogMessage sLogText, "Sending:", mtInformation
        If mgvbLogToMsgLog Then
            LogMessage "Sending: " & sLogText, mtInformation
        End If
    End If
                
    On Error GoTo Error
    If mgvbSimulation = False Then
        With mobjWinHTTPRequest
            .Open "GET", sCmd, False
            '.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
            
            On Error Resume Next
            .Send
            If Err Then
                mopAcousticStatus = FAULT
                .Abort
                GoTo Done
            End If
            
            .WaitForResponse
'            If Not mobjWinHTTPRequest Is Nothing Then
                If .Status = 200 Then
                    mopAcousticStatus = NOFAULT
                    If APIindex = bkapiGet_RecordingNumber Then
                        mopAcousticRecNo = .ResponseText
                    End If
                Else
                    mopAcousticStatus = FAULT
                End If
                If mgvbLogInput And bLog Then
                    mobjAcousticMessageLogFile.LogMessage .ResponseText, "Received:", mtInformation
                    If mgvbLogToMsgLog Then
                        LogMessage "Received: " & .ResponseText, mtInformation
                    End If
                End If
'                .Abort
'            End If
        End With
    End If
    
    GoTo Done

Error:
    sRaiseErrorText = LoadResString(ridBKConnectCmdError, sCmd, Err.Description)
    bRaiseError = True
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        EH_RaiseError 5, TypeName(Me), PROC_NAME, sRaiseErrorText
    End If

    Exit Sub

End Sub
'
'Private Sub SendParValues()
'
''
'
'    Const PROC_NAME = "SendParValues"
'    Dim bRaiseError As Boolean
'    Dim sRaiseErrorText As String
'    Dim lRaiseErrorNumber As Long
'    Dim vPars(0 To 4) As Variant
'    Dim fWindspeed As Single
'    Dim fYaw As Single
'    Dim fAirtemp As Single
'    Dim fPAbs As Single
'    Dim fRelHumidity As Single
'
'    On Error GoTo Error
'
'    bRaiseError = False
'    mgvdLastSendParsTime = Now
'
'    fWindspeed = mipWindspeed
'    fYaw = mipYaw
'    fAirtemp = mipAirtemp
'    fPAbs = mipAbsPress
'    fRelHumidity = mipRelHumidity
'    If mccbUseTargetValues Then
'        Select Case mipWsMode
'        Case 2
'            If mipWsAchieved = 1# Then
'                fWindspeed = mipWsTarget
'            End If
'        End Select
'        If mipYawAchieved = 1# Then
'            fYaw = mipYawTarget
'        End If
'    End If
'
'' Convert to BKConnect units as necessary
'    fWindspeed = fWindspeed / 3.6       ' Convert to m/s
'    fPAbs = fPAbs / 100#                ' Convert from kPa to bar
'
'    vPars(0) = Array("WindSpeed", fWindspeed)
'    vPars(1) = Array("YawAngle", fYaw)
'    vPars(2) = Array("AirTemperature", fAirtemp)
'    vPars(3) = Array("BarometricAirPressure", fPAbs)
'    vPars(4) = Array("AirHumidity", fRelHumidity)
'
'    SendCommand "SetParametersRequest", Array("parameters", vPars)
'
'    GoTo Done
'
'Error:
'    sRaiseErrorText = Err.Description
'    bRaiseError = True
'    lRaiseErrorNumber = 5
'    Resume Done
'
'Done:
'    On Error Resume Next
'    If bRaiseError Then
'        On Error GoTo 0
'        EH_RaiseError lRaiseErrorNumber, TypeName(Me), PROC_NAME, sRaiseErrorText
'    End If
'
'    Exit Sub
'
'End Sub

Public Sub Shutdown()

    Const PROC_NAME = "Shutdown"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim objSOM As SOM
    
    On Error GoTo Error

    bRaiseError = False
                
    SetProcessPriority priNormal
    
' Shut down actions
' ------------------
    
'    If Not mobjBKConnect Is Nothing Then
'        mobjBKConnect.Disconnect
'        Set mobjBKConnect = Nothing
'    End If
    
    If Not mobjWinHTTPRequest Is Nothing Then
        mobjWinHTTPRequest.Abort
        Set mobjWinHTTPRequest = Nothing
    End If
                
' Clear any pending commands
    bReady = False
    mobjCommands.Clear "Controller shutting down"
    DoEvents
        
    bInited = False
    mobjParManager.ReleaseOutputs
    
    mobjEvents.RaiseQuit
        
    DoEvents
    Unload mobjForm
    Set mobjForm = Nothing
    
    Set objSOM = New SOM
    objSOM.RemoveObject SOM_CONTROLLER
    Set objSOM = Nothing
    
    Class_Terminate
    
    Pause 2#
    
    If IsWithinIDE() Then
        End
    Else
        TerminateProcess GetCurrentProcess(), 0
    End If
    
    GoTo Done

Error:
    sRaiseErrorText = Err.Description
    bRaiseError = True
    EH.Module = TypeName(Me)
    EH.Procedure = PROC_NAME
    Call EH_MsgBox
    Resume Next

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Call EH_RaiseError(OLEERR_METHOD_FAILED, TypeName(Me), PROC_NAME, sRaiseErrorText)
    End If

    Exit Sub

End Sub

Private Sub Start()
Attribute Start.VB_Description = "Private start routine"

' Private start routine

    Const PROC_NAME = "Start"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim i As Long, objProcess As Object
    Dim sClassName As String
    Dim objProcessManagerProcess As Object
    Dim objSPGM As SharedPropertyGroupManager
    Dim bExists As Boolean
    
    On Error GoTo Error

    bRaiseError = False
        
    On Error GoTo Error_CreatingObjects
        
    sClassName = "CommandListLib.Commands"
    Set mobjCommands = New CommandListLib.Commands

    sClassName = "ControlLib.ControlConstants"
    Set mobjConstants = New ControlLib.ControlConstants
    
    sClassName = "ControlLib.Monitor"
    
    sClassName = "Events"
    Set mobjEvents = New Events
    
    sClassName = "Interlocks"
    Set mobjInterlocks = New Interlocks
        
    sClassName = "Invalid Opcodes Collection"
    Set mcolInvalidOpcodes = New Collection
            
    sClassName = "frmControl"
    Set mobjForm = New frmControl
    
    sClassName = "SharedPropertyGroupManager"
    Set objSPGM = New SharedPropertyGroupManager
    
    sClassName = "SharedPropertyGroup - Constants"
    Set mobjSPGConstants = objSPGM.CreatePropertyGroup(SPG_CONSTANTS, 0, 0, bExists)
    Set mobjSPGDataItems = objSPGM.CreatePropertyGroup(SPG_DATA, 0, 0, bExists)
    
    On Error GoTo Error
' Set up loop timing statistics
    Set tmrElapsed = New ccrpStopWatch
    
    DoEvents
    
' Set up interlock objects
    With mobjInterlocks.Add(CStr(ocInitialize))
        .Expressions.AddBoolean "Initialized", bInited, True, ilckVariableEqualValueMeansOK, LoadResString(stsControlSystemNotInitialized)
    End With
' General interlock
    With mobjInterlocks.Add(ILCK_GENERAL)
        .IncludedInterlocks.Add CStr(ocInitialize)
    End With
    With mobjInterlocks.Add(ILCK_COM_OK)
        .IncludedInterlocks.Add ILCK_GENERAL
        .Expressions.Add "TCPIP", mopAcousticStatus, 1#, ilckVariableEqualValueMeansOK, LoadResString(ridTCPCommsFaulted)
    End With
    With mobjInterlocks.Add(ILCK_OPERATION_OK)
        .IncludedInterlocks.Add ILCK_GENERAL
        .IncludedInterlocks.Add ILCK_COM_OK
        ' Add interlock related to safety, estops, etc.
    End With
    
'
'' Show Status interlock
'    With objInterlocks.Add(CStr(ocShowStatus))
'        .IncludedInterlocks.Add CStr(ocStartFan)
'        .IncludedInterlocks.Add CStr(ocStop)
'        .IncludedInterlocks.Add CStr(ocSetBKConnect)
'    End With
    
' Set up invalid opcode collections. Add arrays of invalid opcodes to the collection and use
' the opcode as a string as an index
    With mcolInvalidOpcodes
''        .Add Array(ocSetBKConnectMode), CStr(ocSetBKConnectMode)
''        .Add Array(ocSetBKConnectMode, ocSetDyThrMode), CStr(ocSetDyThrMode)
''        .Add Array(ocSetBKConnectMode, ocSetDyThrMode), CStr(ocStopDy)
    End With
                    
    On Error Resume Next
    Set mobjGlobalVars = New GlobalVars
    With mobjGlobalVars
        .AddBoolean "Running in Simulation Mode", mgvbSimulation
        .AddString "Last Sent Command", mgvsLastSentCommand
        .AddString "TCP Last Response", mgvsLastBKConnectResponse
        .AddDate "TCP Last Response Time", mgvdLastBKConnectResponseTime
                
        .AddLong "BKConnect Response Num Errors", mgvlNumReceiveErrors
        .AddDate "Last BKConnect Response Error Time", mgvdLastBKConnectResponseErrTime
        .AddString "Last BKConnect Response Error", mgvsLastBKConnectResponseErrMsg
        .AddLong "Last BKConnect Response Error Num", mgvlLastBKConnectResponseErrNum
        .AddString "Last BKConnect Response Error Source", mgvsLastBKConnectResponseErrSource
        .AddString "Last BKConnect Response Error Cause", mgvsLastBKConnectResponseErrInput
        .AddBoolean "Log (repetitive) Keep-Alive Communications", mgvbLogKeepAlive
    End With
    On Error GoTo Error
    
' Set up Statuses
    Set mobjStatuses = New Statuses
    With mobjStatuses
        .UseSPG SPG_STATUSES
        .Add STS_ACOUSTIC_INITIALIZED, False, "Acoustic Not Initialized", "Acoustic Initialized"
        .Add STS_ACOUSTIC_FAULTED, False, "Acoustic Faulted", "Acoustic Not Faulted"
    End With
    
' Create a process object, and then pass this Process
' object to the Process Manager

    Set objProcess = New Process
    bReady = True
    
    WaitForProcessesStartupComplete
    
    Set objProcessManagerProcess = AddTalentProcessWait(CStr(objProcess.Name), objProcess)
    lID = objProcessManagerProcess.ID
            
    Process.LogMode = logToTalentAndFile
    
' Load control constants
    On Error Resume Next
    Initialize initConstants
    On Error GoTo Error
    
' Set up parameter objects (after adding process and getting Process ID)

    On Error GoTo Error_CreatingObjects
    sClassName = "ParManager"
    Set mobjParManager = New ParManager
    mobjParManager.OwnerID = lID
    With mobjParManager
        .AutoRefreshEnabled = False
        .IgnoreLoadedEvents = True
        .IgnoreUnloadedEvents = True
    
    ' Input Parameters
'        .AddWithoutFlag "Airtemp", mipAirtemp, ioInput
'        .AddWithoutFlag "WindSpeed", mipWindspeed, ioInput
'        .AddWithoutFlag "WSMode", mipWsMode, ioInput
''        .AddWithoutFlag "WS_Achieved", mipWsAchieved, ioInput
'        .AddWithoutFlag "WindspeedTarget", mipWsTarget, ioInput
'        .AddWithoutFlag "RelHumidity", mipRelHumidity, ioInput
'        .AddWithoutFlag "TT_Yaw", mipYaw, ioInput
''        .AddWithoutFlag "TT_YawAchieved", mipYawAchieved, ioInput
'        .AddWithoutFlag "TT_PosTgt", mipYawTarget, ioInput
'        .AddWithoutFlag "PStatic", mipAbsPress, ioInput

    ' Output Parameters
        AddOutput "AcousticStatus", mopAcousticStatus, False, False
'        AddOutput "BKConnectState", mopBKConnectState, True, True
        AddOutput "AcousticRecNo", mopAcousticRecNo, True, True
'        AddOutput "BKConnectRemote", mopFCASInControl, False, True
             
    End With
    
' Now fully started up. Set process priority and enable timers
' If running within VB (the IDE) then use a form based timer, and don't
' use Real Time Priority
    If mccfUpdateTime < 0.05 Then
        mccfUpdateTime = 0.05
    End If

    If IsWithinIDE() Then
        mobjForm.tmrDesignTime.Interval = mccfUpdateTime * 1000
        mobjForm.tmrDesignTime.Enabled = True
    Else
        tmrMain.Stats.Frequency = Reg_GetValue(HKEY_LOCAL_MACHINE, REG_PROCESSES, , "TimerFrequency", 10, vbLong)
        tmrMain.Interval = mccfUpdateTime * 1000           ' 50 milliseconds
        tmrMain.EventType = TimerPeriodic
        Set tmrMain.Notify = Me
        tmrMain.Enabled = True
    End If

    Set mobjLogFile = New LogFile
    Set mobjAcousticMessages = New MailSlot
    With mobjAcousticMessages
        .MailSlotType = mstCustom
        .Name = "AcousticMessages"
        .OpenOutput
    End With
    Set mobjAcousticMessageLogFile = New LogFile
    Set mobjAcousticMessageLogFile.MailSlot = mobjAcousticMessages

' Insert reset command to set up communications
    If mobjParManager.Configurations.Count > 0 Then
        mopAcousticStatus = 1#
        mgvdLastBKConnectResponseTime = Now
        mdStartTime = Now
        mobjCommands.Insert ocInitialize, "Initialize"
    Else
        mobjCommands.Insert ocResetCommunications, "ResetAcousticCommunications"
    End If
    
    
' Clear Run Once timer
    Set tmrRunOnce = Nothing
    tmrElapsed.Reset

    GoTo Done
    
Error:
    sRaiseErrorText = Err.Description
'    bRaiseError = True
'    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
'    End If
    LogMessage "Error during controller startup: " & sRaiseErrorText & ". See log file for details", mtFailure
    Resume Next

Error_CreatingObjects:
    sRaiseErrorText = "Can't create object '" & sClassName & "': " & Err.Description
    bRaiseError = True
    EH.Module = TypeName(Me)
    EH.Procedure = PROC_NAME
    EH.Item(1) = sClassName
    Call EH_MsgBox
    Resume Done

Error_Configurations:
    sRaiseErrorText = "Unable to connect to Configurations collection: " & Err.Description
    EH.AppError = AE.CantCreateConfigurations
    EH.Module = TypeName(Me)
    EH.Procedure = PROC_NAME
    Call EH_MsgBox
    Resume Next
    
Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Call EH_RaiseError(OLEERR_METHOD_FAILED, TypeName(Me), PROC_NAME, sRaiseErrorText)
    End If

    Exit Sub

End Sub

Private Sub Suspend()
Attribute Suspend.VB_Description = "Suspend control"

' Suspend control

    On Error Resume Next

    bInited = False
    mobjParManager.ReleaseOutputs
    mobjStatuses(STS_ACOUSTIC_INITIALIZED).SetValue bInited
'    mobjBKConnect.Disconnect
    
    If Not mobjWinHTTPRequest Is Nothing Then
        mobjWinHTTPRequest.Abort
        Set mobjWinHTTPRequest = Nothing
    End If
                
    Exit Sub

End Sub

Private Sub tmrRunOnce_Timer(ByVal Milliseconds As Long)

    Const PROC_NAME = "tmrRunOnce_Timer"

    On Error GoTo Error
    
    tmrRunOnce.Enabled = False
    Set tmrRunOnce = Nothing

    Start

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

    Exit Sub


End Sub

Private Function URLEncode(EncodeStr As String) As String

'*** EXAMPLE *** from Curl to http
''    strRequest = "https://iam.cloud.ibm.com/identity/token"
''   - H 'Accept: application/json'
''   -H 'Content-Type: application/x-www-form-urlencoded'
''   --data-urlencode "apikey=D_wVsNM0o9ACq-Z3fc2VzXMAeJ3fZxz7SqVJ47gfKVrh"
''   --data-urlencode "response_type=cloud_iam"
''   --data-urlencode "grant_type=urn:ibm:params:oauth:grant-type:apikey"
'
'BECOMES
'    strRequest = "https://iam.cloud.ibm.com/identity/token?" & _
'    "apikey=" & URLEncode("D_wVsNM0o9ACq-Z3fc2VzXMAeJ3fZxz7SqVJ47gfKVrh") & _
'    "&response_type=" & URLEncode("cloud_iam") & _
'    "&grant_type=" & URLEncode("urn:ibm:params:oauth:grant-type:apikey")
'

    Dim i As Integer
    Dim erg As String
    
    erg = EncodeStr
 
    ' *** First replace '%' chr
    erg = Replace(erg, "%", Chr(1))
 
    ' *** then '+' chr
    erg = Replace(erg, "+", Chr(2))
    
    For i = 0 To 255
        Select Case i
            ' *** Allowed 'regular' characters
            Case 37, 43, 48 To 57, 65 To 90, 97 To 122
            
            Case 1  ' *** Replace original %
                erg = Replace(erg, Chr(i), "%25")
        
            Case 2  ' *** Replace original +
                erg = Replace(erg, Chr(i), "%2B")
                
            Case 32
                erg = Replace(erg, Chr(i), "+")
        
            Case 3 To 15
                erg = Replace(erg, Chr(i), "%0" & Hex(i))
        
            Case Else
                erg = Replace(erg, Chr(i), "%" & Hex(i))
                
        End Select
    Next
    
    URLEncode = erg
    
End Function
'
'Private Property Get WindspeedNowAchieved() As Boolean
'
'    Dim bAchieved As Boolean
'
'    On Error GoTo Error
'
'    bAchieved = False
'    If mipWsMode = 2 Then
'        If mipWsAchieved = 1# Then
'            bAchieved = True
'        End If
'    End If
'
'    If bAchieved Then
'        If mgvbWsWasAchieved = False Then
'            WindspeedNowAchieved = True
'            mgvfWsNowAchievedTgt = mipWsTarget
'        ElseIf mipWsTarget <> mgvfWsNowAchievedTgt Then
'            WindspeedNowAchieved = True
'            mgvfWsNowAchievedTgt = mipWsTarget
'        End If
'    End If
'    mgvbWsWasAchieved = bAchieved
'
'    GoTo Done
'
'Error:
'    WindspeedNowAchieved = False
'    Resume Done
'
'Done:
'    Exit Property
'
'End Property

Private Sub WriteOutputs()

    Const PROC_NAME = "WriteOutputs"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim objPar As ManagedPar
    
    On Error GoTo Error

    bRaiseError = False

' Calculate outputs that are dependent upon items received from BK Connect
    

' Write control outputs
    If mopAcousticStatus = 1# Then
        mobjParManager.WriteOutputs
    Else
        On Error Resume Next
        For Each objPar In mobjParManager
            If objPar.InputOrOutput = ioOutput Then
                If objPar.Tag = True Then
                Else
                    objPar.WriteValue
                End If
            End If
        Next objPar
    End If
        
    GoTo Done

Error:
    sRaiseErrorText = Err.Description
    bRaiseError = True
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        EH_RaiseError OLEERR_METHOD_FAILED, TypeName(Me), PROC_NAME, sRaiseErrorText
    End If

    Exit Sub

End Sub
'
'Private Property Get YawAngleNowAchieved() As Boolean
'
'    Dim bAchieved As Boolean
'
'    On Error GoTo Error
'
'    If mipYawAchieved = 1# Then
'        bAchieved = True
'    End If
'
'    If bAchieved Then
'        If mgvbYawWasAchieved = False Then
'            YawAngleNowAchieved = True
'            mgvfYawNowAchievedTgt = mipYawTarget
'        ElseIf mipYawTarget <> mgvfYawNowAchievedTgt Then
'            YawAngleNowAchieved = True
'            mgvfYawNowAchievedTgt = mipYawTarget
'        End If
'    End If
'    mgvbYawWasAchieved = bAchieved
'
'    GoTo Done
'
'Error:
'    YawAngleNowAchieved = False
'    Resume Done
'
'Done:
'    Exit Property
'
'End Property
