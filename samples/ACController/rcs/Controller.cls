head 1.5;
branch ;
access ;
symbols CD%204%2E00%2E0021:1.5 CD%204%2E00%2E0020:1.5 4%2E00%2E0015:1.5
        CD%204%2E00%2E0019:1.4 4%2E00%2E0014:1.4 4%2E00%2E0013:1.3 4%2E00%2E0002:1.2 4%2E00%2E0001:1.1;
locks ; strict;
comment @@;


1.5
date 2024.03.26.15.22.04; author Kavanagh; state Develop;
branches ;
next 1.4;

1.4
date 2020.01.21.03.05.37; author Kavanagh; state Develop;
branches ;
next 1.3;

1.3
date 2020.01.20.05.01.31; author Kavanagh; state Develop;
branches ;
next 1.2;

1.2
date 2020.01.20.02.12.48; author Kavanagh; state Develop;
branches ;
next 1.1;

1.1
date 2020.01.20.00.44.02; author Kavanagh; state Develop;
branches ;
next ;

ext
@project Z:/Client Facilities/CAERI/Source Code/Applications/ACController/CWTCSACController.pj;
@

desc
@@


1.5
log
@Loop Timer@
text
@VERSION 1.0 CLASS
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
Private WithEvents tmrMain As LoopTimer
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
Private mobjGFAI As SocketPoll
Private mobjLogFile As TalentProcessLib.LogFile
Private mobjGFAIMessages As MailSlot
Private WithEvents mobjGFAIMessageLogFile As LogFile
Attribute mobjGFAIMessageLogFile.VB_VarHelpID = -1

Implements ICcrpTimerNotify

Private lID As Long
Private bReady As Boolean
Private bInited As Boolean
Private Const bKeepStats As Boolean = False
Private bRealTime As Boolean
Private mgvdLastCommandTime As Date
Private mvInterlockCheckValue As Variant
Private mdStartTime As Date

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
Private mccfGFAITCPIPConnectTimeout As Single
Private mccsGFAITCPIPAddress As String
Private mcclGFAITCPIPPort As Long
Private Const mccbGFAITCPIPAnsi As Boolean = True
Private mccfKeepAlivePacketInterval As Single
Private mccfMaxTimeBetweenSendPars As Single
Private mccbUseTargetValues As Boolean
Private mcclGetControlTimeout As Long
Private mccbSendParsOnlyIfUnderWTCS As Boolean

' Module level Inputs
Private mipAirtemp As Single
Private mipWindspeed As Single
Private mipWsMode As Single
Private mipWsAchieved As Single
Private mipWsTarget As Single
Private mipRelHumidity As Single
Private mipYaw As Single
Private mipYawAchieved As Single
Private mipYawTarget As Single
Private mipAbsPress As Single

' Module level Outputs
Private mopGFAITCPIPStatus As Single
Private mopGFAIState As Single
Private mopWTCSInControl As Single

' Module level Global variables
Private mgvbLogOutput As Boolean
Private mgvbLogInput As Boolean
Private mgvbLogToMsgLog As Boolean

Private mgvlNumReceiveErrors As Long
Private mlDefPriority As Long
Private mgvsLastSentCommand As String
Private mgvbSimulation As Boolean
Private mgvbWsWasAchieved As Boolean
Private mgvfWsNowAchievedTgt As Single
Private mgvbYawWasAchieved As Boolean
Private mgvfYawNowAchievedTgt As Boolean
Private mgvdLastGFAIResponseTime As Date
Private mgvsLastGFAIResponse As String
Private mgvdGFAIDataLastRequestTime As Date
Private mgvdLastDataStreamLoggedTime As Date
Private mgvdLastSendParsTime As Date

Private mgvsLastGFAIResponseErrMsg As String
Private mgvlLastGFAIResponseErrNum As Long
Private mgvsLastGFAIResponseErrSource As String
Private mgvsLastGFAIResponseErrInput As String
Private mgvdLastGFAIResponseErrTime As Date

Private mgvdLastKeepAlivePacketTime As Date


Public Property Get ParManager() As Object
    Set ParManager = mobjParManager
End Property
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

Public Property Get Events() As Events
    Set Events = mobjEvents
End Property


Private Sub ProcessGFAIResponse(ByVal GFAIResponse As String)
Attribute ProcessGFAIResponse.VB_Description = "Process a received data packet"

' Process a received TCP/IP response message

' Part          Description
' ----------------------------------------------------------------------
' DataResponse

    Const PROC_NAME = "ProcessGFAIResponse"
    Static vLastErrorMessageTime As Date
    Dim sItems() As String
    Dim lNumItems As Long
    Dim sSubItems() As String
    Dim lNumSubItems As Long
    Dim sCommand As String
    Dim sFaultText As String
    Dim lErrDescStart As Long
    Dim I As Long
    Dim lMessageType As MessageTypes
    Dim sLogText As String
    
    On Error GoTo Error
    
    GFAIResponse = Trim(GFAIResponse)
    If Left(GFAIResponse, 1) = Chr(2) Then      ' leading terminator
        If Len(GFAIResponse) < 3 Then
            GoTo Done
        End If
        GFAIResponse = Mid(GFAIResponse, 2)
    End If
    If Right(GFAIResponse, 1) = Chr(3) Then     ' ETX
        GFAIResponse = Left(GFAIResponse, Len(GFAIResponse) - 1)
    End If
    mgvsLastGFAIResponse = GFAIResponse
                
    If mgvbLogInput Then
        sLogText = Replace(GFAIResponse, vbLf, "")
        sLogText = Replace(sLogText, vbCr, "")
        mobjGFAIMessageLogFile.LogMessage sLogText, "Received:", mtInformation
        If mgvbLogToMsgLog Then
            LogMessage "Received: " & sLogText, mtInformation
        End If
    End If
        
    If GFAIResponse Like "*StillAliveReply*" Then
        ' Do Nothing
            
    ElseIf GFAIResponse Like "*ErrorEvent*" Then
        lErrDescStart = InStr(GFAIResponse, """errorDescription"": ") + 21
        For I = lErrDescStart To Len(GFAIResponse)
            If Mid(GFAIResponse, I, 1) = """" Then
                Exit For
            End If
            sFaultText = sFaultText & Mid(GFAIResponse, I, 1)
        Next I
        If GFAIResponse Like "*""Critical""*" Then
            lMessageType = msgCritical
        ElseIf GFAIResponse Like "*""Error""*" Then
            lMessageType = msgCritical
        ElseIf GFAIResponse Like "*""Warning""*" Then
            lMessageType = msgExclamation
        Else
            lMessageType = msgInformation
        End If
        LogMessage sFaultText, lMessageType
        
    ElseIf GFAIResponse Like "*ShutdownEvent*" Then
        ' GFAI is shutting down
        mobjGFAI.Shutdown
        mopGFAITCPIPStatus = 0#
        mopWTCSInControl = 0#
    ElseIf (GFAIResponse Like "*GetAMSStatusReply*") Or (GFAIResponse Like "*AMSStatusChangedEvent*") Then
        If GFAIResponse Like "*Uninitialized*" Then
            mopGFAIState = gsUninitialized
        ElseIf GFAIResponse Like "*Preview*" Then
            mopGFAIState = gsPreview
        ElseIf GFAIResponse Like "*PreTrigger*" Then
            mopGFAIState = gsPreTrigger
        ElseIf GFAIResponse Like "*TriggerReady*" Then
            mopGFAIState = gsTriggerReady
        ElseIf GFAIResponse Like "*Recording*" Then
            mopGFAIState = gsRecording
        ElseIf GFAIResponse Like "*DataTransfer*" Then
            mopGFAIState = gsDataTransfer
        ElseIf GFAIResponse Like "*CancelRecord*" Then
            mopGFAIState = gsCancelRecord
        ElseIf GFAIResponse Like "*PostProcessing*" Then
            mopGFAIState = gsPostProcessing
        ElseIf GFAIResponse Like "*MeasurementFinished*" Then
            mopGFAIState = gsMeasurementFinished
        Else
            mopGFAIState = -1
        End If
        
    ElseIf GFAIResponse Like "*ReleaseControlRequest*" Then
        SendCommand "ReleaseControlReply", Array("requestControlReply", "RequestGranted")
        mopWTCSInControl = 0#
        
    ElseIf GFAIResponse Like "*GetControlReply*" Then
        If GFAIResponse Like "*RequestGranted*" Then
            mopWTCSInControl = 1#
        Else
            mopWTCSInControl = 0#
            If CurrentCommandOpcode = ocTakeControl Then
                AbortCurrentCommand "Denied by GFAI system"
            End If
        End If
    
    ElseIf GFAIResponse Like "*GetCurrentControlOwnerReply*" Then
        If GFAIResponse Like "*You*" Then
            mopWTCSInControl = 1#
        Else
            mopWTCSInControl = 0#
        End If
    
    ElseIf GFAIResponse Like "*ForceReleaseControlEvent*" Then
        mopWTCSInControl = 0#
   
    
    ElseIf GFAIResponse Like "*UnsolicitedGetControlCommand*" Then
        mopWTCSInControl = 1#
    
    End If
    
    GoTo Done

Error:
    mgvlNumReceiveErrors = mgvlNumReceiveErrors + 1
    mgvsLastGFAIResponseErrMsg = Err.Description
    mgvlLastGFAIResponseErrNum = Err.Number
    mgvsLastGFAIResponseErrSource = Err.Description
    mgvsLastGFAIResponseErrInput = GFAIResponse
    
    If (Err.Number And vbObjectError) <> vbObjectError Then
        Set EH.VBErr = Err
        If (SecondsSinceTime(vLastErrorMessageTime) >= 300) Then
            EH.AppError = LoadResString(AE.GFAIResponseError)
            EH.Item(1) = GFAIResponse
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
Private Function CurrentCommandOpcode() As Long
    If mobjCommands.Count = 0 Then
        CurrentCommandOpcode = 0
    Else
        CurrentCommandOpcode = mobjCommands(1).Opcode
    End If
End Function
Private Sub AbortCurrentCommand(Optional CompletionMessage As String)
    If mobjCommands.Count > 0 Then
        mobjCommands(1).Abort CompletionMessage
    End If
End Sub

Private Sub Suspend()
Attribute Suspend.VB_Description = "Suspend control"

' Suspend control

    On Error Resume Next

    bInited = False
    mobjParManager.ReleaseOutputs
    mobjStatuses(STS_GFAI_INITIALIZED).SetValue bInited
    mobjGFAI.Disconnect
    
    Exit Sub

End Sub
Private Sub GFAIControlLogic()
Attribute GFAIControlLogic.VB_Description = "Closed loop control logic"

' Closed loop control logic

    Const PROC_NAME = "GFAIControlLogic"
    Dim xData() As Byte
    Dim I As Long, sResponse As String
    Dim objPar As ManagedPar
    Dim bExists As Boolean
    Dim fValue As Single
    Dim bCmdReqd As Boolean
    
    On Error GoTo Error
    
    On Error Resume Next
    If mgvbSimulation Then
        mopGFAITCPIPStatus = NOFAULT
    ElseIf mobjGFAI.State = socConnected Then
        mopGFAITCPIPStatus = NOFAULT
    Else
        mopGFAITCPIPStatus = FAULT
    End If
    If Err Then
        mopGFAITCPIPStatus = FAULT
    End If
            
' Send out the parameter (windspeed, yaw, etc.) values whenever windspeed or yaw or airtemp is achieved, or a maximum time period has elapsed
    If mgvbSimulation = False Then
        If Not mobjGFAI Is Nothing Then
            If mobjGFAI.State = soxConnected Then
                If (mccbSendParsOnlyIfUnderWTCS = False) Or (mopWTCSInControl = 1#) Then
                    If WindspeedNowAchieved Or YawAngleNowAchieved Or ((mccfMaxTimeBetweenSendPars > 0) And (SecondsSinceTime(mgvdLastSendParsTime) >= mccfMaxTimeBetweenSendPars)) Then
                        On Error Resume Next
                        If mopGFAIState <= gsPreview Then
                            SendParValues
                        End If
                        On Error GoTo Error
                    End If
                End If
            End If
        End If
    End If
    
' Send Keepalive packet if connected
    If mgvbSimulation = False Then
        If Not mobjGFAI Is Nothing Then
            If mobjGFAI.State = soxConnected Then
                If mccfKeepAlivePacketInterval > 0# Then
                    If SecondsSinceTime(mgvdLastKeepAlivePacketTime) >= mccfKeepAlivePacketInterval Then
                        On Error Resume Next
                        SendCommand "StillAliveRequest"
                        mgvdLastKeepAlivePacketTime = Now
                        On Error GoTo Error
                    End If
                End If
            End If
        End If
    End If
    
' Safety monitoring
    If mopGFAITCPIPStatus = 0 Then
        mopGFAIState = 0
        mopWTCSInControl = 0
    End If
    
' Simulate GFAI responses
    If mgvbSimulation Then
    End If
    
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
Private Property Get WindspeedNowAchieved() As Boolean

    Dim bAchieved As Boolean
    
    On Error GoTo Error
    
    bAchieved = False
    If mipWsMode = 2 Then
        If mipWsAchieved = 1# Then
            bAchieved = True
        End If
    End If
    
    If bAchieved Then
        If mgvbWsWasAchieved = False Then
            WindspeedNowAchieved = True
            mgvfWsNowAchievedTgt = mipWsTarget
        ElseIf mipWsTarget <> mgvfWsNowAchievedTgt Then
            WindspeedNowAchieved = True
            mgvfWsNowAchievedTgt = mipWsTarget
        End If
    End If
    mgvbWsWasAchieved = bAchieved
            
    GoTo Done
    
Error:
    WindspeedNowAchieved = False
    Resume Done
    
Done:
    Exit Property
    
End Property
Private Property Get YawAngleNowAchieved() As Boolean

    Dim bAchieved As Boolean
    
    On Error GoTo Error
    
    If mipYawAchieved = 1# Then
        bAchieved = True
    End If
    
    If bAchieved Then
        If mgvbYawWasAchieved = False Then
            YawAngleNowAchieved = True
            mgvfYawNowAchievedTgt = mipYawTarget
        ElseIf mipYawTarget <> mgvfYawNowAchievedTgt Then
            YawAngleNowAchieved = True
            mgvfYawNowAchievedTgt = mipYawTarget
        End If
    End If
    mgvbYawWasAchieved = bAchieved
            
    GoTo Done
    
Error:
    YawAngleNowAchieved = False
    Resume Done
    
Done:
    Exit Property

End Property

Public Property Get Initialized() As Boolean
Attribute Initialized.VB_Description = "Returns whether or not the controller is initialized (controlling) or not"

' Returns whether or not the controller is initialized (controlling) or
' not

' Returns
' --------------------------------------------------------
' True
' False

    Const PROC_NAME = "Initialized"
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
Public Sub CheckInterlock(Opcode, Optional CheckValue)

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
    sIndex = Opcode
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

Public Property Get Interlocks() As Interlocks
    Set Interlocks = mobjInterlocks
End Property
Public Function GetInvalidOpcodes(Opcode) As Variant
    
    Dim sIndex As String
    On Error Resume Next
    sIndex = Opcode
    GetInvalidOpcodes = mcolInvalidOpcodes(sIndex)
    If Err Then
        GetInvalidOpcodes = 0
    End If
    Exit Function
    
End Function
Private Sub Start()
Attribute Start.VB_Description = "Private start routine"

' Private start routine

    Const PROC_NAME = "Start"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim I As Long, objProcess As Object
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
        .Expressions.Add "TCPIP", mopGFAITCPIPStatus, 1#, ilckVariableEqualValueMeansOK, LoadResString(ridTCPCommsFaulted)
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
'        .IncludedInterlocks.Add CStr(ocSetGFAI)
'    End With
    
' Set up invalid opcode collections. Add arrays of invalid opcodes to the collection and use
' the opcode as a string as an index
    With mcolInvalidOpcodes
''        .Add Array(ocSetGFAIMode), CStr(ocSetGFAIMode)
''        .Add Array(ocSetGFAIMode, ocSetDyThrMode), CStr(ocSetDyThrMode)
''        .Add Array(ocSetGFAIMode, ocSetDyThrMode), CStr(ocStopDy)
    End With
                    
    On Error Resume Next
    Set mobjGlobalVars = New GlobalVars
    With mobjGlobalVars
        .AddBoolean "Running in Simulation Mode", mgvbSimulation
        .AddString "Last Sent Command", mgvsLastSentCommand
        .AddString "TCP Last Response", mgvsLastGFAIResponse
        .AddDate "TCP Last Response Time", mgvdLastGFAIResponseTime
                
        .AddLong "GFAI Response Num Errors", mgvlNumReceiveErrors
        .AddDate "Last GFAI Response Error Time", mgvdLastGFAIResponseErrTime
        .AddString "Last GFAI Response Error", mgvsLastGFAIResponseErrMsg
        .AddLong "Last GFAI Response Error Num", mgvlLastGFAIResponseErrNum
        .AddString "Last GFAI Response Error Source", mgvsLastGFAIResponseErrSource
        .AddString "Last GFAI Response Error Cause", mgvsLastGFAIResponseErrInput
                        
    End With
    On Error GoTo Error
    
' Set up Statuses
    Set mobjStatuses = New Statuses
    With mobjStatuses
        .UseSPG SPG_STATUSES
        .Add STS_GFAI_INITIALIZED, False, "Acoustic Not Initialized", "Acoustic Initialized"
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
        .AddWithoutFlag "Airtemp", mipAirtemp, ioInput
        .AddWithoutFlag "WindSpeed", mipWindspeed, ioInput
        .AddWithoutFlag "WS_Mode", mipWsMode, ioInput
        .AddWithoutFlag "WS_Achieved", mipWsAchieved, ioInput
        .AddWithoutFlag "WS_Target", mipWsTarget, ioInput
        .AddWithoutFlag "RelHumidity", mipRelHumidity, ioInput
        .AddWithoutFlag "TT_Yaw", mipYaw, ioInput
        .AddWithoutFlag "TT_YawAchieved", mipYawAchieved, ioInput
        .AddWithoutFlag "TT_YawTarget", mipYawTarget, ioInput
        .AddWithoutFlag "PStatic", mipAbsPress, ioInput

    ' Output Parameters
        AddOutput "GFAIComStatus", mopGFAITCPIPStatus, False, False
        AddOutput "GFAIState", mopGFAIState, True, True
        AddOutput "GFAIRemote", mopWTCSInControl, False, True
             
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
    Set mobjGFAIMessages = New MailSlot
    With mobjGFAIMessages
        .MailSlotType = mstCustom
        .Name = "GFAIMessages"
        .OpenOutput
    End With
    Set mobjGFAIMessageLogFile = New LogFile
    Set mobjGFAIMessageLogFile.MailSlot = mobjGFAIMessages


' Insert reset command to set up communications
    If mobjParManager.Configurations.Count > 0 Then
        mopGFAITCPIPStatus = 1#
        mgvdLastGFAIResponseTime = Now
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
        EH.Module = "Controller"
        EH.Procedure = "Start"
        Call EH_MsgBox
'    End If
    LogMessage "Error during controller startup: " & sRaiseErrorText & ". See log file for details", mtFailure
    Resume Next

Error_CreatingObjects:
    sRaiseErrorText = "Can't create object '" & sClassName & "': " & Err.Description
    bRaiseError = True
    EH.Module = "Controller"
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
        Call EH_RaiseError(OLEERR_METHOD_FAILED, "Controller", "Start", sRaiseErrorText)
    End If

    Exit Sub

End Sub
Private Sub AddOutput(ParameterName As String, ValueHolder As Single, ResetToDefaultUponComError As Boolean, IsCalculatedBySimulator As Boolean)
    If mgvbSimulation Then
        mobjParManager.AddWithoutFlag(ParameterName, ValueHolder, IIf(IsCalculatedBySimulator, ioInput, ioOutput)).Tag = ResetToDefaultUponComError
    Else
        mobjParManager.AddWithoutFlag(ParameterName, ValueHolder, ioOutput).Tag = ResetToDefaultUponComError
    End If
End Sub
Public Property Get Commands() As CommandListLib.Commands
Attribute Commands.VB_Description = "Returns the command list"

' Returns the command list

' Returns
' --------------------------------------------------------
' The return value is an object ...

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
        EH.Module = "Controller"
        EH.Procedure = "Commands"
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Call EH_RaiseError(OLEERR_PROPERTY_GET_FAILED, "Controller", "Commands", sRaiseErrorText)
    End If

    Exit Property

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
    Dim I As Long
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
        Select Case .Opcode

' Initialize
        Case ocInitialize
            Me.Initialize , objCommand
            .Remove
            mobjStatuses(STS_GFAI_INITIALIZED).SetValue bInited
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
            mobjStatuses(STS_GFAI_INITIALIZED).SetValue bInited
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
            If mobjGFAI Is Nothing Then
                Set mobjGFAI = New TalentProcessLib.SocketPoll
            End If
            If (mobjGFAI.State <> socConnected) Then
                mobjGFAI.Disconnect
                mobjGFAI.Caption = "GFAI"
                Set mobjGFAI.LogFile = mobjGFAIMessageLogFile
                mobjGFAI.logToTalent = False
                mobjGFAI.ConnectTimeout = mccfGFAITCPIPConnectTimeout
                mobjGFAI.ConnectAsClient mccsGFAITCPIPAddress, CInt(mcclGFAITCPIPPort)
                If mobjGFAI.LogMode = logNone Then
                    mobjGFAIMessageLogFile.LogMessage "Connected to " & mccsGFAITCPIPAddress & " using Port " & mcclGFAITCPIPPort & " as client"
                End If
                With mobjGFAI.BinData
                    .UseStartCharacter = True
                    .StartCharacter = Chr(2)
                    .UseEndCharacter = True
                    .EndCharacter = Chr(3)
                End With
                Pause 1#
                mobjCommands.Insert ocGetCurrentControl, "GetAcousticCurrentControl"
                mobjCommands.Insert ocGetCurrentStatus, "GetAcousticCurrentStatus"
            End If
            .Remove
                   
        Case ocStartLogging
            mobjLogFile.OpenFile "GFAI", lfAppend, lfBinaryLog, True
            mobjGFAIMessages.OpenOutput
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
            mobjGFAIMessages.CloseMailslot
            mobjLogFile.CloseFile
            mobjGFAI.LogMode = logNone
            mgvbLogInput = False
            mgvbLogOutput = False
            .Remove
                                        
        Case Else       ' All other commands require initialized state
            
            CheckInterlock ocInitialize
            
            If mgvbSimulation = False Then
                If (mopGFAITCPIPStatus = 0#) Then
                    .InsertPredecessor ocResetCommunications, "ResetAcousticCommunications"
                    GoTo Done
                End If
            End If
            
            Select Case .Opcode
                            
            Case ocSendPars
                CheckInterlock ILCK_GENERAL
                .Messages.Flush
                SendParValues
                .Remove
                
            Case ocGetCurrentControl
                CheckInterlock ILCK_GENERAL
                SendCommand "GetCurrentControlOwnerRequest"
                .Remove
                
            Case ocGetCurrentStatus
                CheckInterlock ILCK_GENERAL
                SendCommand "GetAMSStatusRequest"
                .Remove
                
            Case ocTakeControl
                CheckInterlock ILCK_GENERAL
                If .FirstTime Then
                    SendCommand "GetControlRequest", Array("timeout", mcclGetControlTimeout)
                ElseIf mopWTCSInControl = 1# Then
                    .Remove
                    SendParValues
                ElseIf .TimeSince >= mcclGetControlTimeout Then
                    .Abort , cmdTimedout
                End If
                
            Case ocReleaseControl
                CheckInterlock ILCK_GENERAL
                SendCommand "UnsolicitedControlReleasedEvent"
                mopWTCSInControl = 0#
                .Remove
                
            Case ocStartRecord
                CheckInterlock ILCK_GENERAL
                If mopWTCSInControl = 0 Then
                    .InsertPredecessor ocTakeControl, "TakeAcousticControl"
                Else
                    .Messages.Flush
                    SendCommand "StartRecordCommand"
                    .Remove
                End If
                
            Case ocSetTrigger
                CheckInterlock ILCK_GENERAL
                If mopWTCSInControl = 0 Then
                    .InsertPredecessor ocTakeControl, "TakeAcousticControl"
                Else
                    .Messages.Flush
                    SendCommand "SetTriggerCommand"
                    .Remove
                End If
                
            Case ocCancelRecord
                CheckInterlock ILCK_GENERAL
                If mopWTCSInControl = 0 Then
                    .InsertPredecessor ocTakeControl, "TakeAcousticControl"
                Else
                    .Messages.Flush
                    SendCommand "CancelRecordCommand"
                    .Remove
                End If
                
            Case Else
                .Abort "Unrecognized opcode: " & .Opcode
                
            End Select
        End Select
        
    End With
    GoTo Done

Error:
    sRaiseErrorText = Err.Description
    lRaiseErrorNumber = CCommandError(Err.Number)
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = "Controller"
        EH.Procedure = "ExecuteCommands"
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
Friend Sub ControlLogic()

    Const PROC_NAME = "ControlLogic"
' This routine is executed at every control interval

    Dim lElapsed As Long    ' Milliseconds
    Dim sData As String
    Dim bExists As Boolean
    Dim fCommSendTime As Single
    Dim I As Long
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
        If Not mobjGFAI Is Nothing Then
            If (mobjGFAI.State = socConnected) Then
                For I = 1 To mobjGFAI.PacketCount
                    mgvdLastGFAIResponseTime = Now
                    sResponse = mobjGFAI.Packets(I)
                    ProcessGFAIResponse sResponse
                Next I
                mobjGFAI.Flush
            End If
        End If
    End If
        
    ExecuteCommands
    
    If bInited = True Then
    
' Perform control logic
        GFAIControlLogic
        
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
        sData = mobjGFAI.GetData
        
    End If
    
    If Not mobjStatuses Is Nothing Then
        mobjStatuses.FlushMessages
    End If
    
    GoTo Done

Error:
    If Err.Number And vbObjectError <> vbObjectError Then
        EH.Module = "MainLine"
        EH.Procedure = "ControlLogic"
        Call EH_MsgBox
    End If
    Suspend
    Resume Next

Done:
    On Error Resume Next

    Exit Sub

End Sub
Private Sub WriteOutputs()

'

    Const PROC_NAME = "WriteOutputs"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim objPar As ManagedPar
    
    On Error GoTo Error

    bRaiseError = False



' Calculate outputs that are dependent upon items received from the GFAI UDP Data
    

' Write control outputs
    If mopGFAITCPIPStatus = 1# Then
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

Friend Sub Initialize(Optional InitType As InitTypes = initAll, _
            Optional Command As Command)

' Perform basic initialization/startup functions

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
        sFileName = Reg_GetValue(HKEY_LOCAL_MACHINE, REG_GFAI, , "ControlConstantsFile", sFileName)
        sDefaultsFileName = App.Path & "\" & App.EXEName & "Defaults.csv"
        With mobjConstants
            .OpenFileWithDefaults sFileName, sDefaultsFileName
            
            .AddString "Priority", mccsPriority, bStatus, sText
            .AddSingle "UpdateInterval", mccfUpdateTime, bStatus, sText
            .AddSingle "GFAITCPIPConnectTimeout", mccfGFAITCPIPConnectTimeout, bStatus, sText
            .AddString "GFAITCPIPAddress", mccsGFAITCPIPAddress, bStatus, sText
            .AddLong "GFAITCPIPPort", mcclGFAITCPIPPort, bStatus, sText
            .AddSingle "KeepAliveInterval", mccfKeepAlivePacketInterval, bStatus, sText
            .AddSingle "SendParsIntervalMax", mccfMaxTimeBetweenSendPars, bStatus, sText
            .AddBoolean "SendParsUseTargets", mccbUseTargetValues, bStatus, sText
            .AddLong "GetControlTimeout", mcclGetControlTimeout, bStatus, sText
            .AddBoolean "SendParsWTCSOnly", mccbSendParsOnlyIfUnderWTCS, bStatus, sText
            
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
                    .Module = "Controller"
                    .Procedure = "Initialize"
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
                .Module = "Controller"
                .Procedure = "Initialize"
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
    
    GoTo Done

Error:
    bInited = False
    lRaiseErrorNumber = CCommandError(Err.Number)
    sRaiseErrorText = Err.Description
    bRaiseError = True
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = "Controller"
        EH.Procedure = "Initialize"
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
    mobjStatuses(STS_GFAI_INITIALIZED).SetValue bInited
    mobjConstants.CloseFile
    Set objEH = Nothing
    If bInited Then
        SetProcessPriority mlDefPriority
    Else
        SetProcessPriority priNormal
    End If
    If bRaiseError Then
        On Error GoTo 0
        Call EH_RaiseError(lRaiseErrorNumber, "Controller", "Initialize", sRaiseErrorText)
    End If

    Exit Sub

End Sub


Public Property Get ControlConstants() As ControlLib.ControlConstants
Attribute ControlConstants.VB_Description = "Returns the collection of control constants for use by an external debugger"

' Returns the collection of control constants for use by an
' external debugger (Look)

' Returns
' --------------------------------------------------------
' The return value is a ControlConstants object

    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String

    On Error GoTo Error

    bRaiseError = False
    Set ControlConstants = mobjConstants
    
    GoTo Done

Error:
    sRaiseErrorText = Err.Description
    bRaiseError = True
    EH.Module = "Controller"
    EH.Procedure = "ControlConstants"
    Call EH_MsgBox
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Call EH_RaiseError(OLEERR_PROPERTY_GET_FAILED, "Controller", "ControlConstants", sRaiseErrorText)
    End If

    Exit Property

End Property
Public Property Get GlobalVars() As ControlLib.GlobalVars

' Returns the collection of global variables for use by an
' external debugger (Look)

' Returns
' --------------------------------------------------------
' The return value is a GlobalVars object

    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String

    On Error GoTo Error

    bRaiseError = False
    Set GlobalVars = mobjGlobalVars
    
    GoTo Done

Error:
    sRaiseErrorText = Err.Description
    bRaiseError = True
    EH.Module = "Controller"
    EH.Procedure = "GlobalVars"
    Call EH_MsgBox
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Call EH_RaiseError(OLEERR_PROPERTY_GET_FAILED, "Controller", "GlobalVars", sRaiseErrorText)
    End If

    Exit Property

End Property
Public Property Get Ready() As Boolean

'
'
'


' Returns
' --------------------------------------------------------
' True
' False
 
    On Error Resume Next
    Ready = bReady
    Exit Property

End Property
Private Sub SendParValues()

'

    Const PROC_NAME = "SendParValues"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    Dim vPars(0 To 4) As Variant
    Dim fWindspeed As Single
    Dim fYaw As Single
    Dim fAirtemp As Single
    Dim fPAbs As Single
    Dim fRelHumidity As Single
    
    On Error GoTo Error

    bRaiseError = False
    mgvdLastSendParsTime = Now

    fWindspeed = mipWindspeed
    fYaw = mipYaw
    fAirtemp = mipAirtemp
    fPAbs = mipAbsPress
    fRelHumidity = mipRelHumidity
    If mccbUseTargetValues Then
        Select Case mipWsMode
        Case 2
            If mipWsAchieved = 1# Then
                fWindspeed = mipWsTarget
            End If
        End Select
        If mipYawAchieved = 1# Then
            fYaw = mipYawTarget
        End If
    End If
    
' Convert to GFAI units as necessary
    fWindspeed = fWindspeed / 3.6       ' Convert to m/s
    fPAbs = fPAbs / 100#                ' Convert from kPa to bar
    
    vPars(0) = Array("WindSpeed", fWindspeed)
    vPars(1) = Array("YawAngle", fYaw)
    vPars(2) = Array("AirTemperature", fAirtemp)
    vPars(3) = Array("BarometricAirPressure", fPAbs)
    vPars(4) = Array("AirHumidity", fRelHumidity)

    SendCommand "SetParametersRequest", Array("parameters", vPars)
    
    GoTo Done

Error:
    sRaiseErrorText = Err.Description
    bRaiseError = True
    lRaiseErrorNumber = 5
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        EH_RaiseError lRaiseErrorNumber, TypeName(Me), PROC_NAME, sRaiseErrorText
    End If

    Exit Sub

End Sub
Public Sub Shutdown()

'
'
'


    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim objSOM As SOM
    
    On Error GoTo Error

    bRaiseError = False
                
    SetProcessPriority priNormal
    
' Shut down actions
' ------------------
    
    If Not mobjGFAI Is Nothing Then
        mobjGFAI.Disconnect
        Set mobjGFAI = Nothing
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
    EH.Module = "Controller"
    EH.Procedure = "Shutdown"
    Call EH_MsgBox
    Resume Next

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Call EH_RaiseError(OLEERR_METHOD_FAILED, "Controller", "Shutdown", sRaiseErrorText)
    End If

    Exit Sub

End Sub




Private Sub Class_Initialize()

'

    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim sClassName As String
    
    On Error GoTo Error

    Call Main
    
    bRaiseError = False
    On Error GoTo Error_CreatingObjects
    
    sClassName = "ccrpTimers6.ccrpTimer"
    Set tmrMain = New LoopTimer
    If App.StartMode <> vbSModeStandalone Then
        Set tmrRunOnce = New ccrpTimers6.ccrpTimer
    End If
        
    On Error GoTo Error
        
    If App.StartMode <> vbSModeStandalone Then
        tmrRunOnce.Interval = 500
        tmrRunOnce.EventType = TimerOneShot
        tmrRunOnce.Enabled = True
    End If
    
    mgvbSimulation = Reg_GetValue(HKEY_LOCAL_MACHINE, REG_GFAI, , "Simulation", False)
    
    GoTo Done

Error:
    sRaiseErrorText = Err.Description
    bRaiseError = True
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = "Controller"
        EH.Procedure = "Class_Initialize"
        Call EH_MsgBox
    End If
    Resume Done

Error_CreatingObjects:
    sRaiseErrorText = "Can't create object '" & sClassName & "': " & Err.Description
    bRaiseError = True
    EH.Module = "Controller"
    EH.Procedure = "Class_Initialize"
    EH.Item(1) = sClassName
    Call EH_MsgBox
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Call EH_RaiseError(OLEERR_METHOD_FAILED, "Controller", "Class_Initialize", sRaiseErrorText)
    End If

    Exit Sub

End Sub

Private Sub Class_Terminate()

    On Error Resume Next
    
    If Not mobjGFAI Is Nothing Then
        mobjGFAI.Disconnect
        Set mobjGFAI = Nothing
    End If
    
    mobjGFAIMessages.CloseMailslot
    Set mobjGFAIMessages = Nothing
    Set mobjGFAIMessageLogFile.MailSlot = Nothing
    mobjGFAIMessageLogFile.CloseFile
    Set mobjGFAIMessageLogFile = Nothing
    
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



Private Sub ICcrpTimerNotify_Timer(ByVal Milliseconds As Long)

    
    On Error Resume Next
    ControlLogic
    DoEvents
    Exit Sub
    
End Sub

Private Sub mobjCommands_CommandAdded(NewCommand As CommandListLib.Command)

    Dim objMessage As TalentProcessLib.Message

    On Error Resume Next
    If NewCommand.SuccessorKey = "" Then
        If mgvbLogInput Then
            Set objMessage = New TalentProcessLib.Message
            With objMessage
                .Source = "Command"
                .Text = NewCommand.Text & GetCmdArgs(NewCommand.Arguments)
                .MessageType = mtCommand
            End With
            mobjGFAIMessageLogFile.LogMessage2 objMessage
        End If
    End If
        
    
    Exit Sub
    
End Sub

Private Sub mobjCommands_CommandCompleted(Command As CommandListLib.Command, CompletionStatus As CommandListLib.CompletionStatuses, CompletionMessage As String)

'
    Const PROC_NAME = "mobjCommands_CommandCompleted"
    Dim objEH As EH
    Dim objMessage As TalentProcessLib.Message
    Dim I As Long
    
    On Error GoTo Error
    
    Select Case CompletionStatus
    Case cmdAborted, cmdTimedout
        If Command.SuppressLogging = False Then
            Set objEH = New EH
            Select Case Command.Opcode
            Case ocResetCommunications
                objEH.AppError = LoadResString(AE.CommunicationsFaulted, Command.Text, CompletionMessage, mccsGFAITCPIPAddress)
            Case Else
                objEH.AppError = AppErrorText(tceCommandFailed, Command.Text, CompletionMessage)
            End Select
            objEH.Module = TypeName(Me)
            objEH.Procedure = PROC_NAME
            Set objEH.Component = App
            LogPopupMessage objEH, Command.Text
            Command.SuppressLogging = True
        ElseIf Command.Opcode = ocResetCommunications Then
            Set objEH = New EH
            objEH.AppError = LoadResString(AE.CommunicationsFaulted, Command.Text, CompletionMessage, mccsGFAITCPIPAddress)
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
            mobjGFAIMessageLogFile.LogMessage2 objMessage
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
        EH.Module = "Controller"
        EH.Procedure = "mobjCommands_CommandInserted"
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next

    Exit Sub

End Sub

Private Sub mobjCommands_MessageFlushed(MessageText As String, MessageType As Long)

'

    On Error GoTo Error

    LogMessage MessageText, MessageType

    GoTo Done

Error:
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = "Controller"
        EH.Procedure = "mobjCommands_MessageFlushed"
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next

    Exit Sub

End Sub
Private Sub mobjInterlocks_Evaluate(InterlockName As String, CurrentOKValue As Boolean, AdditionalMessages As Collection)

    
    Dim objCommand As CommandListLib.Command

    On Error Resume Next
    Select Case InterlockName
    Case ILCK_GENERAL
''        If mopGFAIDataStatus = 1# Then
''            If mopGFAIRemote <> 1# Then
''                AdditionalMessages.Add LoadResString(ridGFAINotInRemote)
''                CurrentOKValue = False
''            End If
''        End If
    End Select
    
    Exit Sub
    
End Sub

Private Sub mobjGFAIMessageLogFile_MessageLogged(Message As Variant)

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
''        sSource = sSource & " (GFAI):"
''        objMessage2.Source = sSource
        mobjLogFile.LogMessage2 objMessage2
    End If
    
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

Private Sub tmrRunOnce_Timer(ByVal Milliseconds As Long)

'

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

Public Sub SendCommand(CommandToSend As String, ParamArray ParPairs() As Variant)

'

' Part          Description
' ----------------------------------------------------------------------
' Command
' Param

    Const PROC_NAME = "SendCommand"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    Dim sCmd As String
    Dim sParName As String
    Dim vParValue As Variant
    Dim sArrayParName As String
    Dim vArrayParValue As Variant
    Dim I As Long
    Dim J As Long
    Dim xCommand() As Byte
    Dim sLogText As String
    
    On Error GoTo Error

    bRaiseError = False

    sCmd = vbCrLf & "{" & vbCrLf & "  " & """cmd"": """ & CommandToSend & """"
    For I = LBound(ParPairs, 1) To UBound(ParPairs, 1)
        sParName = ParPairs(I)(0)
        vParValue = ParPairs(I)(1)
        sCmd = sCmd & "," & vbCrLf & "  " & """" & sParName & """: "
        If IsArray(vParValue) Then
            sCmd = sCmd & vbCrLf & "  {"
            For J = LBound(vParValue, 1) To UBound(vParValue, 1)
                sArrayParName = vParValue(J)(0)
                vArrayParValue = vParValue(J)(1)
                Select Case J
                Case LBound(vParValue, 1) ', UBound(vParValue, 1)
                Case Else
                    sCmd = sCmd & ","
                End Select
                sCmd = sCmd & vbCrLf & "    " & """" & sArrayParName & """: "
                Select Case VarType(vArrayParValue)
                Case vbString
                    sCmd = sCmd & """" & vArrayParValue & """"
                Case Else
                    sCmd = sCmd & vArrayParValue
                End Select
            Next J
            sCmd = sCmd & vbCrLf & "  }"
        Else
            Select Case VarType(vParValue)
            Case vbString
                sCmd = sCmd & """" & vParValue & """"
            Case Else
                sCmd = sCmd & vParValue
            End Select
        End If
    Next I
    sCmd = sCmd & vbCrLf & "}" & vbCrLf
        
    mgvsLastSentCommand = sCmd
    mgvdLastCommandTime = Now
    
    On Error GoTo Error
    If mgvbSimulation = False Then
        If mccbGFAITCPIPAnsi Then
            xCommand = StrConv(Chr(2) & sCmd & Chr(3), vbFromUnicode)
            mobjGFAI.SendData xCommand                                   ' Add {STX} and {ETX} to message sent to GFAI
        Else
            mobjGFAI.SendData Chr(2) & sCmd & Chr(3)       ' Add {STX} and {ETX} to message sent to GFAI
        End If
    End If
    If mgvbLogOutput Then
        sLogText = Replace(sCmd, vbLf, "")
        sLogText = Replace(sLogText, vbCr, "")
        mobjGFAIMessageLogFile.LogMessage sLogText, "Sending:", mtInformation
        If mgvbLogToMsgLog Then
            LogMessage "Sending: " & sLogText, mtInformation
        End If
    End If
                
    GoTo Done

Error:
    sRaiseErrorText = LoadResString(ridGFAICmdError, CommandToSend, Err.Description)
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
Private Sub LogMessageEH(MessageText As String, SendingRoutine As String, Optional Items, Optional MessageType As MessageTypes = msgInformation)

    Dim objEH As EH
    Dim sItems() As String
    Dim I As Long
    
    On Error Resume Next
    Set objEH = New EH
    With objEH
        .AppError = MessageText
        If Not IsMissing(Items) Then
            For I = 1 To 10
                Err.Clear
                .Item(I) = Items(I)
                If Err Then
                    Exit For
                End If
            Next I
        End If
        .Module = "Controller"
        .Procedure = SendingRoutine
        .MessageType = MessageType
        LogMessage2 objEH
    End With
    Exit Sub
    
End Sub
@


1.4
log
@@
text
@d25 1
a25 1
Private WithEvents tmrMain As ccrpTimer
d1634 1
a1634 1
    Set tmrMain = New ccrpTimers6.ccrpTimer
@


1.3
log
@As tested at AAWT@
text
@d382 5
a386 1

@


1.2
log
@@
text
@d74 1
d96 1
d157 4
d177 6
a182 1
        mobjGFAIMessageLogFile.LogMessage GFAIResponse, "Received:", mtInformation
d187 20
d211 1
d237 2
a238 1
    
d244 3
d256 1
a256 1
     ElseIf GFAIResponse Like "*ForceReleaseControlEvent*" Then
d260 3
d293 13
d352 8
a359 4
                If WindspeedNowAchieved Or YawAngleNowAchieved Or ((mccfMaxTimeBetweenSendPars > 0) And (SecondsSinceTime(mgvdLastSendParsTime) >= mccfMaxTimeBetweenSendPars)) Then
                    On Error Resume Next
                    SendParValues
                    On Error GoTo Error
d958 5
d1009 1
d1017 1
d1023 1
a1023 1
                    .InsertPredecessor "TakeAcousticControl", ocTakeControl
d1033 1
a1033 1
                    .InsertPredecessor "TakeAcousticControl", ocTakeControl
d1043 1
a1043 1
                    .InsertPredecessor "TakeAcousticControl", ocTakeControl
d1271 1
d2051 1
a2051 1
                Case LBound(vParValue, 1), UBound(vParValue, 1)
d2088 6
a2093 1
        mobjGFAIMessageLogFile.LogMessage sCmd, "Sending:", mtInformation
@


1.1
log
@Initial revision@
text
@d313 7
a319 5
                If SecondsSinceTime(mgvdLastKeepAlivePacketTime) >= mccfKeepAlivePacketInterval Then
                    On Error Resume Next
                    SendCommand "StillAliveRequest"
                    mgvdLastKeepAlivePacketTime = Now
                    On Error GoTo Error
@
