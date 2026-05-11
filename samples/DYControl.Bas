Attribute VB_Name = "modDYControl"
Option Explicit
Option Compare Text

Private Const MOD_NAME = "modDyControl"

' Object, Shared property Group, and Shared Property names

Public Const SOM_CONTROLLER = "Controller"
Public Const SPG_CONSTANTS = "Constants"
Public Const SPG_DATA = "Data"

Public Enum ResIDs
    ridDyModeStopped = 107
    ridThrModeStopped = 113
    ridNoResponseFromDyno = 118
    ridDynoModeIsNow = 119
    ridThrModeIsNow = 120
    ridNoSpeedControl = 121
    ridNoForceControl = 122
    ridNoPositionControl = 123
    ridInvalidDyCmdArgument = 124
    ridErrorDyCmd = 125
    ridDynoRejectedCmd = 132
    ridBadDynoMode = 134
    ridBadThrMode = 135
    ridBadDynoThrMode = 136
    ridVehicleIDSet = 137
    ridBadDyConfig = 143
    ridConnectionRefused = 145
    ridSocketNotCreated = 146
    ridSocketNotConnected = 147
    ridSocketInUndefinedState = 148
    ridDynoNotStopped = 149
    ridThrottleNotStopped = 150
    ridDynoMustBeIn4WD = 151
    ridDynoMustBeIn4WDSpeedSyncOff = 152
    ridDynoMustBeIn4WDForceViaVehOff = 153
    ridDynoMustBeInSpeed = 154
    ridDynoMustBeInForce = 155
    ridDynoNotInRemote = 156
    ridDynoNotInVehicleTestMode = 157
    ridTCPCommsFaulted = 159
    ridUDPCommsFaulted = 160
    ridDynoError = 162
    ridUDPCommsStarted = 163
    ridSpeedSyncTimedOut = 164
    ridCantSendTCPCommand = 165
    ridWaitingForOfflineModes = 166
    ridCmdTimedOut = 167
    ridDyNoCarrVehs = 168
    ridBadActuator = 169
    ridRobotNotInUse = 170
    ridNoActuatorSelected = 171
    ridTCPIPSetupFail = 172
    ridCmdNotSupported = 173
End Enum

Public Enum CCNames
    ccSpeedMax = 0
    ccSpeedRampMax = 1
    ccSpeedRampDefault = 2
    ccK0Min = 3
    ccK0Max = 4
    ccK1Min = 5
    ccK1Max = 6
    ccK2Min = 7
    ccK2Max = 8
    ccK3Min = 9
    ccK3Max = 10
    ccGradientMin = 11
    ccGradientMax = 12
    ccGradientRateMax = 13
    ccGradientRateDefault = 14
    ccPosMax = 15
    ccPosMin = 16
    ccPosRampMax = 17
    ccPosRampMin = 18
    ccPosRampDefault = 19
    ccForceRampMax = 20
    ccForceRampMin = 21
    ccForceRampDefault = 22
    ccForceMax = 23
    ccForceMin = 24
    ccTRefMin = 25
    ccTRefMax = 26
    ccTRefDefault = 27
    ccMaxCarrPosnPosition = 28
    ccMinCarrPosnPosition = 29
    ccMinRobotDCReps = 30
    ccMaxRobotDCReps = 31
End Enum

Public Enum SharedDataItems
    sdDyMode = 0
    sdThrMode = 1
    sdDyConfig = 2
    sdTranType = 3
    sdRbtMode = 4
    sdRbtInUse = 5
End Enum

Public Enum DynoConfigs
    dycOff = 0
    dycFront = 1
    dycRear = 2
    dyc4WD = 3
    dycFrontTracking = 4
    dycRearTracking = 5
End Enum

Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
            (ByVal pDest As Any, ByVal pSource As Any, ByVal ByteLen As Long)

Public Enum AE
    InvalidActiveXLaunch = 101
    CantCreateConfigurations = 102
    InitializeFailedParsMissing = 103
    CantLoadConstants = 104
    DataPacketError = 105
    CommunicationsFaulted = 158
    DynoAsyncError = 161
End Enum

Public Enum Opcodes
    ocAddCarriageVeh = 1
    ocClearAxleSim = 2
    ocClearDynoLimits = 3
    ocClearRobotBuffer = 4          ' Internal
    ocClearTireSim = 5
    ocEStopDyno = 6
    ocForceSplitViaVehicle = 7
    ocGetRobotDriveCycles = 8
    ocGetRobotShiftTables = 9
    ocGetRobotVehicleTypes = 10
    ocInitialize = 11
    ocLimitForce = 12               ' Internal
    ocLimitSpeed = 13               ' Internal
    ocMoveCarriage = 14
    ocRefreshDynoLists = 15
    ocRefreshRobotLists = 16
    ocRefreshVehID = 17
    ocRemoveCarriageVeh = 18
    ocReqRbtData = 19               ' Internal
    ocReqRbtErrors = 20             ' Internal
    ocReqRbtFilenames = 21          ' Internal
    ocReqTranType = 22              ' Internal
    ocResetCarriage = 23
    ocResetCommunications = 24
    ocResetDy = 25
    ocResetRobot = 26               ' Internal
    ocSelectRobotDriveCycle = 27
    ocSelectRobotShiftTable = 28
    ocSelectRobotVehicleType = 29
    ocSendDynoCommand = 30
    ocSetAxleSim = 31
    ocSetCarriageVeh = 32
    ocSetDynoConfig = 33
    ocSetDynoLimits = 34
    ocSetDynoMode = 35
    ocSetDynoOption = 36
    ocSetDynoSync = 37
    ocSetDyThrMode = 38
    ocSetForce = 39
    ocSetForceZero = 40
    ocSetGrad = 41
    ocSetK0 = 42
    ocSetK1 = 43
    ocSetK2 = 44
    ocSetK3 = 45
    ocSetPosition = 46
    ocSetPositionZero = 47
    ocSetRbtOnline = 48             ' Internal
    ocSetRbtRemote = 49             ' Internal
    ocSetRbtStandby = 50            ' Internal
    ocSetRLS = 51
    ocSetRobotMode = 52
    ocSetSpeed = 53
    ocSetSpeedZero = 54
    ocSetThrMode = 55
    ocSetTireSim = 56
    ocSetVehicleGear = 57
    ocSetVehicleID = 58
    ocShowRobotVehicleFiles = 59
    ocShowStatus = 60
    ocShutdown = 61
    ocStartCalib = 62
    ocStartLogging = 63
    ocStartRobotDriveCycle = 64
    ocStartRobotLearnCycle = 65
    ocStartVehicleEngine = 66
    ocStartWarmup = 67
    ocStopAux = 68
    ocStopDy = 69
    ocStopLogging = 70
    ocStopRobot = 71
    ocStopRobotDriveCycle = 72
    ocStopVehicleEngine = 73
    ocSuspend = 74
    ocSyncSpeed = 75
    ocTurnKeyOn = 76
    ocZeroDist = 77
    ocResetRobotComm = 78
    ocSetActuator = 79
    ocStartRbtLogging = 80
    ocStopRbtLogging = 81
    ocSendRobotCommand = 82
    ocUseThrottlePot = 83
End Enum

Public Enum DynoModes
    dmStopped = 0
    dmSpeed = 1
    dmForce = 2
    dmRLS = 3
    dmWarmup = 4
    dmCalib = 5
End Enum

Public Enum ThrottleModes
    tmStopped = 0
    tmSpeed = 1
    tmForce = 2
    tmPosition = 3
End Enum

Public Enum RobotModes
    rmStopped = 0
    rmSpeed = 1
    rmForce = 2
    rmPosition = 3
    rmRPM = 4
    rmVacuum = 5
End Enum

Public Enum BrakeModes
    bmStopped = 0
    bmForce = 2
    bmPedalForce = 7
End Enum

Public Enum DriveCycleStatus
    DC_NOTLOADED = 0
    DC_LOADED = 1
    DC_RUNNING = 2
    DC_SUSPENDED = 3
End Enum

Public Enum TCPIPCompletionStatuses
    comWaiting = 0
    comACK = 1
    comNACK = 2
End Enum

Public Enum Actuators
    actNone = -1
    actDyno = 0
    actThrottle = 1
    actBrake = 3
    actRobot = 4
End Enum

Public Enum TransmissionTypes
    ttAutoTran = 0
    ttManualTran = 1
End Enum

Public Enum ControlActuators
    caThrottle = 0
'    caDyno = 0
'    caThrottle = 1
    caRobot = 1
End Enum

'Public Enum ThrottleActuators
'    taNone = 0
'    taCable = 1
'    taInvehicle = 2
'    taRobot = 3
'    taThrottleBrake = 4
'    taThrottleDynoBraking = 5
'End Enum

Public Const NUM_COMMANDS = 115
Public Const INACTIVE_TIMER = -1
Public Const NOFAULT = 1#
Public Const FAULT = 0#
Public Const UseDefaultValue = -10000000000#

Public Const ILCK_GENERAL As String = "General"
Public Const ILCK_GENERAL_ROBOT = "Robot"
Public Const ILCK_SPEED As String = "Speed"
Public Const ILCK_THROTTLE = "Throttle"
Public Const ILCK_FORCE As String = "Force"
Public Const ILCK_POS As String = "Pos"
Public Const ILCK_INIT As String = "1"
Public Const ILCK_SPEED_RECEIVE As String = "Speed2"
Public Const ILCK_SPEED_RECEIVE_DIFF As String = "Speed3"
Public Const ILCK_FORCE_RECEIVE As String = "Force2"
Public Const ILCK_FORCE_RECEIVE_DIFF As String = "Force3"
Public Const ILCK_POS_RECEIVE = "Pos2"
Public Const ILCK_COM_OK = "ComOK"

' Status Names
Public Const SPG_STATUSES = "Statuses"
Public Const STS_DYNO_INITIALIZED = "DYNO_INITIALIZED"
Public Const STS_DYNO_MODE_ACHIEVED = "DYNO_MODE_ACHIEVED"
Public Const STS_THROTTLE_MODE_ACHIEVED = "THROTTLE_MODE_ACHIEVED"
Public Const STS_FORCE_ACHIEVED = "FORCE_ACHIEVED"
Public Const STS_SPEED_ACHIEVED = "SPEED_ACHIEVED"
Public Const STS_THROTTLE_ACHIEVED = "THROTTLE_ACHIEVED"
Public Const STS_RBT_MODE_ACHIEVED = "ROBOT_MODE_ACHIEVED"
Public Const STS_CARR_POS_ACHIEVED = "CARR_POS_ACHIEVED"

Public Enum DynoCommandCodes
    dcResetDyno = 1
    dcSetDynoConfig = 2
    dcStartWarmnup = 3
    dcStartCalibration = 4
    dcSetDynoThrottleMode = 5
    dcZeroDistance = 7
    dcSetDynoLimit = 8
    dcSetThrottleLimit = 9
    dcSetSpeed = 10
    dcSetForce = 11
    dcSetPosition = 12
    dcSetVehicle = 13
    dcSetRLS = 16
    dcSetGradient = 18
    dcStopDynoSystem = 19
    dcStopDynoFast = 20
    dcStartDynoData = 21
    dcReportFault = 22
    dcSendWTData = 23
    dcGetVehicleIds = 24
    
    dcSetSpeedDiff = 50
    dcSetSpeedSync = 51
    dcSetForceRatio = 52
    dcSetForceSplitVeh = 53
    
    dcSetActiveMode = 54
    dcSetZeroSpeedMode = 55
    
    dcSetSpeedtoZero = 100
    dcSetForcetoZero = 101
    dcSetPositiontoZero = 102
    dcFreeText = 103
    dcStopDynoData = 104
    
    dcMoveCarriage = 120
    dcResetCarriage = 121
    dcSetActuator = 122
    dcSetThrPot = 123
    
    rcClearBuff = 201
    rcGetData = 202
    rcGetDriveCycleList = 203
    rcGetFaults = 204
    rcGetFileNames = 205
    rcGetShiftTables = 206
    rcGetVehicleTypes = 207
    rcGetTransType = 208
    rcReset = 209
    rcSelectDriveCycle = 210
    rcSetMode = 211
    rcSetOnlineDirect = 212
    rcSetRemote = 213
    rcSelectShiftTable = 214
    rcSetStandby = 215
    rcSelectVehicleType = 216
    rcSetGear = 217
    rcStartDriveCycle = 218
    rcStartLearnCycle = 219
    rcStartEngine = 220
    rcStop = 221
    rcStopDriveCycle = 222
    rcStopEngine = 223
    rcTurnKeyOn = 224
    rcPowerOn = 225
    rcPowerOff = 226
    rcFreeText = 227
End Enum

Public Const ACK_STRING = "ACK"
Public Const NACK_STRING = "NACK"

' Dyno Command Text
Public Const DYCMD_SetSpeed = "SETSPEED"
Public Const DYCMD_SetForce = "SETFORCE"
Public Const DYCMD_SetPosition = "SETPOSITION"

' Robot Command Text

' "AADF"    ' Get current vehicle file names
' "AFBT"       ' Gets list of vehicle types
' "AFGD"    ' Get Data on Selected Vehicle Type
' "AFPT"       ' Gets list of Drive Cycles
' "ASTZ"   ' Gets error status
' "AVST"       ' Gets list of shift tables
'@ "AWRX"     ' Get Robot Data
' "EFBW"       ' Select a vehicle type
' "EFPW"       ' Select a Drive Cycle
' "ESTW"       ' Select a shift table
' "SFOD"       ' Turns on Robot (online direct)
'@ "SFOD"    ' Sets robot to online direct mode.
'@ "SLDR"         ' Clears ramp time in RBT buffer
' "SLFB"     ' Start robot learn cycle
' "SPAU"      ' Turns off Robot's power
' "SREM"   ' Disables robot keyboard
' "SRES"         ' (SRES??) Aborts function, stops vehicle
'@ "SRES"        ' Sets Robot to offline
'@ "SSTF"     ' Start selected Drive Cycle
' "SSTM"    ' Start Vehicle Engine
'@ "STBY"      ' Stop Currernt Drive Cycle
' "STBY"     ' Turn on Power before starting vehicle
'@ "SVRD"     ' Set Robot mode
' "SZAU"     ' Stop Vehicle Engine
' "SZEI"      ' Start Vehicle Engine

Public Const RBTCMD_CLR_BUFF = "SLDR"         ' Clears ramp time in RBT buffer
Public Const RBTCMD_GET_RBT_Data = "AWRX"     ' Get Robot Data
Public Const RBTCMD_GET_RBT_DC = "AFPT"       ' Gets list of Drive Cycles
Public Const RBTCMD_GET_RBT_FAULTS = "ASTZ"   ' Gets error status
Public Const RBTCMD_GET_RBT_FILENAMES = "AADF"    ' Get current vehicle file names
Public Const RBTCMD_GET_RBT_ST = "AVST"       ' Gets list of shift tables
Public Const RBTCMD_GET_RBT_VT = "AFBT"       ' Gets list of vehicle types
Public Const RBTCMD_GET_TRAN_TYPE = "AFGD"    ' Get Data on Selected Vehicle Type
Public Const RBTCMD_RBT_RESET = "SRES"        ' Sets Robot to offline
Public Const RBTCMD_SET_RBT_DC = "EFPW"       ' Select a Drive Cycle
Public Const RBTCMD_SET_RBT_MODE = "SVRD"     ' Set Robot mode
Public Const RBTCMD_SET_RBT_OnDir = "SFOD"    ' Sets robot to online direct mode.
Public Const RBTCMD_SET_RBT_Remote = "SREM"   ' Disables robot keyboard
Public Const RBTCMD_SET_RBT_ST = "ESTW"       ' Select a shift table
Public Const RBTCMD_SET_RBT_STBY = "STBY"     ' Turn on Power before starting vehicle
Public Const RBTCMD_SET_RBT_VT = "EFBW"       ' Select a vehicle type
Public Const RBTCMD_SET_VEH_GEAR = "SetVehicleGear(internal only)"
Public Const RBTCMD_START_RBT_DC = "SSTF"     ' Start selected Drive Cycle
Public Const RBTCMD_START_RBT_LC = "SLFB"     ' Start robot learn cycle
Public Const RBTCMD_START_VEH_ENG = "SSTM"    ' Start Vehicle Engine
Public Const RBTCMD_STOP_RBT = "SRES"         ' (SRES??) Aborts function, stops vehicle
Public Const RBTCMD_STOP_RBT_DC = "STBY"      ' Stop Currernt Drive Cycle
Public Const RBTCMD_STOP_VEH_ENG = "SZAU"     ' Stop Vehicle Engine
Public Const RBTCMD_TURN_KEY_ON = "SZEI"      ' Start Vehicle Engine
Public Const RBTCMD_TURNOFF_RBT = "SPAU"      ' Turns off Robot's power
Public Const RBTCMD_TURNON_RBT = "SFOD"       ' Turns on Robot (online direct)

Public Const VEH_GEARSTRING_P = "P"
Public Const VEH_GEARSTRING_D = "D"
Public Const VEH_GEARSTRING_M = "M"
Public Const VEH_GEARSTRING_PLUS = "+"
Public Const VEH_GEARSTRING_MINUS = "-"
Public Const VEH_GEARSTRING_R = "R"
Public Const VEH_GEARSTRING_N = "N"
Public Const VEH_GEARSTRING_1 = "1"
Public Const VEH_GEARSTRING_2 = "2"
Public Const VEH_GEARSTRING_3 = "3"
Public Const VEH_GEARSTRING_4 = "4"
Public Const VEH_GEARSTRING_5 = "5"
Public Const VEH_GEARSTRING_6 = "6"
Public Const VEH_GEARSTRING_AUTO = "Auto"

Public Const AUTO_SELECT = 1    ' Robot determines appropriate gear
Public Const USER_SELECT = 0    ' User chooses gear
Public Const AUTO_GEAR_MINUS = 10
Public Const AUTO_GEAR_PLUS = 9
Public Const AUTO_GEAR_M = 8
Public Const AUTO_GEAR_P = 7
Public Const AUTO_GEAR_R = 6
Public Const AUTO_GEAR_D = 5
Public Const AUTO_GEAR_4 = 4
Public Const AUTO_GEAR_3 = 3
Public Const AUTO_GEAR_2 = 2
Public Const AUTO_GEAR_1 = 1
Public Const AUTO_GEAR_N = 0
Public Const MAN_GEAR_N = 0
Public Const MAN_GEAR_1 = 1
Public Const MAN_GEAR_2 = 2
Public Const MAN_GEAR_3 = 3
Public Const MAN_GEAR_4 = 4
Public Const MAN_GEAR_5 = 5
Public Const MAN_GEAR_6 = 6
Public Const MAN_GEAR_R = 7
Public Const MAN_GEAR_OFFSET = 10
Public Const MANUAL_TRAN = 1#
Public Const AUTO_TRAN = 0#
Public Const V3_AUTO_SELECT = 1    ' Robot determines appropriate gear

Public Const RBT_ONLINE_DIRECT = 1
Public Const RBT_OFFLINE = 0
Public Const RBT_OPEN_LOOP = 2
Public Const RBT_ONLINE = 3
Public Const RBT_STS_MSG_SEPARATOR = "$SI$"
'Public Const DC_RUNNING = 1#
Public Const DC_STOPPED = 0#
Public Const RBT_SLAVE_ID = "K0"

' Robot modes as seen by robot
Public Const RBT_FRC_MODE_CODE = 3
Public Const RBT_SPD_MODE_CODE = 0
Public Const RBT_POS_MODE_CODE = 1
Public Const RBT_POS_CLUTCH_MODE_CODE = 7

Public Const OFFLINE_MODE_TEXT = "Offline"
Public Const SPEED_MODE_TEXT = "Speed"
Public Const FORCE_MODE_TEXT = "Force"
Public Const RLS_MODE_TEXT = "RLS"
Public Const POSITION_MODE_TEXT = "Position"

Public Const PROCESS_NAME = "Dynamometer"
Public Const PROCESS_DESCRIPTION = "Dyno Controller"
Public Const REG_PROCESSES = "Software\ReACT Technologies\Talent\4.0\Processes"
Public Const REG_DYNO = "Software\ReACT Technologies\Talent\4.0\Dyno"
Public Const REG_SHARED_FOLDERS = "SOFTWARE\ReACT Technologies\Talent\4.0\Shared Folders"

Public Function ActuatorIDAsText(ActuatorID As ControlActuators) As String

    Select Case ActuatorID
'        Case caDyno
'            ActuatorIDAsText = "Dyno"
        Case caThrottle
            ActuatorIDAsText = "Throttle"
        Case caRobot
            ActuatorIDAsText = "Robot"
        Case Else
            ActuatorIDAsText = "Unknown (" & ActuatorID & ")"
    End Select
    
End Function

Public Function ActuatorIDFromText(ActuatorIDText As String) As ControlActuators

'

    Const PROC_NAME = "ActuatorIDFromText"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    Dim I As Long
    
    On Error GoTo Error

    bRaiseError = False

    For I = caThrottle To caRobot
        If ActuatorIDAsText(I) Like ActuatorIDText & "*" Then
            ActuatorIDFromText = I
            GoTo Done
        End If
    Next I
    
    sRaiseErrorText = LoadResString(ridBadActuator, ActuatorIDText)
    bRaiseError = True
    lRaiseErrorNumber = tceInvalidArgs

    GoTo Done

Error:
    sRaiseErrorText = LoadResString(ridBadActuator, Err.Description)
    bRaiseError = True
    lRaiseErrorNumber = tceInvalidArgs
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        EH_RaiseError lRaiseErrorNumber, "modDYControl", PROC_NAME, sRaiseErrorText
    End If

    Exit Function

End Function

Public Function ArgValue(Arg As Variant, Optional TargetVariableType As VbVarType = vbVariant) As Variant

    Dim vValue As Variant
    
    If TypeOf Arg Is CommandListLib.Argument Then
        vValue = Arg.Value
    Else
        vValue = Arg
    End If

    If TargetVariableType = vbVariant Then
        ArgValue = vValue
    Else
        ArgValue = Convert(vValue, TargetVariableType)
    End If
        
    Exit Function
        
End Function

Public Function ArgValueAsString(Expression As Variant) As String

    On Error Resume Next
    Select Case VarType(Expression)
    Case vbString
        ArgValueAsString = """" & CStr(Expression) & """"
    Case Else
        If Expression = UseDefaultValue Then
            ArgValueAsString = ""
        Else
            ArgValueAsString = CStrUS(Expression)
        End If
    End Select
    
    Exit Function
    
End Function

Public Function ASin(Expression As Double) As Double

    If Expression > 0.99999999 Then
        ASin = 3.14159 / 2#
    ElseIf Expression < -0.99999999 Then
        ASin = -3.14159 / 2#
    Else
        ASin = Atn(Expression / Sqr(-Expression * Expression + 1))
    End If
    
End Function

Public Function CalcRampTime(ByVal Target As Single, ByVal Actual As Single, ByVal RampRate As Single) As Single

    On Error Resume Next
    CalcRampTime = 1#
    CalcRampTime = Abs(Target - Actual) / RampRate
    Exit Function
    
End Function

Public Sub CheckModeCombination(DyMode As DynoModes, ThrMode As ThrottleModes)

'
    Const PROC_NAME = "CheckModeCombination"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    Dim I As Long
    
    On Error GoTo Error

    bRaiseError = False

    Select Case DyMode
    Case dmSpeed
        If ThrMode = tmSpeed Then
            Err.Raise 5
        End If
    Case dmRLS, dmForce
        If ThrMode = tmForce Then
            Err.Raise 5
        End If
    End Select
    
    Select Case ThrMode
    Case tmPosition
        Select Case DyMode
        Case dmSpeed, dmForce, dmRLS
        Case Else
            Err.Raise 5
        End Select
    End Select
        
    GoTo Done

Error:
    sRaiseErrorText = LoadResString(ridBadDynoThrMode, DynoModeAsText(DyMode), ThrModeAsText(ThrMode))
    bRaiseError = True
    lRaiseErrorNumber = tceInvalidArgs
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        EH_RaiseError lRaiseErrorNumber, MOD_NAME, PROC_NAME, sRaiseErrorText
    End If

    Exit Sub

End Sub

Public Function DyConfigAsText(DyConfig) As String

    Select Case DyConfig
    Case dycFront
        DyConfigAsText = "FWD"
    Case dycFrontTracking
        DyConfigAsText = "FWDT"
    Case dycRear
        DyConfigAsText = "RWD"
    Case dycRearTracking
        DyConfigAsText = "RWDT"
    Case dyc4WD
        DyConfigAsText = "4WD"
    Case dycOff
        DyConfigAsText = "Off"
    Case Else
        DyConfigAsText = "Unknown (" & DyConfig & ")"
    End Select
End Function

Public Function DyConfigFromText(DynoConfigText As String, Optional OffOK As Boolean) As DynoConfigs

'

    Const PROC_NAME = "DyConfigFromText"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    Dim I As Long
    
    On Error GoTo Error

    bRaiseError = False

    For I = IIf(OffOK, dycOff, dycFront) To dycRearTracking
        If DyConfigAsText(I) Like DynoConfigText & "*" Then
            DyConfigFromText = I
            GoTo Done
        End If
    Next I
    
    sRaiseErrorText = LoadResString(ridBadDyConfig, DynoConfigText)
    bRaiseError = True
    lRaiseErrorNumber = tceInvalidArgs

    GoTo Done

Error:
    sRaiseErrorText = LoadResString(ridBadDyConfig, Err.Description)
    bRaiseError = True
    lRaiseErrorNumber = tceInvalidArgs
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        EH_RaiseError lRaiseErrorNumber, MOD_NAME, PROC_NAME, sRaiseErrorText
    End If

    Exit Function

End Function


Public Function DynoModeAsText(DynoMode) As String

    Select Case DynoMode
    Case dmStopped
        DynoModeAsText = "Offline"
    Case dmSpeed
        DynoModeAsText = "Speed"
    Case dmForce
        DynoModeAsText = "Force"
    Case dmRLS
        DynoModeAsText = "RLS"
    Case dmWarmup
        DynoModeAsText = "Warmup"
    Case dmCalib
        DynoModeAsText = "Calib"
    Case Else
        DynoModeAsText = "Unknown (" & DynoMode & ")"
    End Select
    
End Function

Public Function DynoModeFromText(DynoModeText As String) As DynoModes

'

    Const PROC_NAME = "DynoModeFromText"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    Dim I As Long
    
    On Error GoTo Error

    bRaiseError = False

    For I = dmStopped To dmCalib
        If DynoModeAsText(I) Like DynoModeText & "*" Then
            DynoModeFromText = I
            GoTo Done
        End If
    Next I
    
    sRaiseErrorText = LoadResString(ridBadDynoMode, DynoModeText)
    bRaiseError = True
    lRaiseErrorNumber = tceInvalidArgs

    GoTo Done

Error:
    sRaiseErrorText = LoadResString(ridBadDynoMode, Err.Description)
    bRaiseError = True
    lRaiseErrorNumber = tceInvalidArgs
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        EH_RaiseError lRaiseErrorNumber, MOD_NAME, PROC_NAME, sRaiseErrorText
    End If

    Exit Function

End Function

Public Function GetCmdArgs(Arguments As CommandListLib.Arguments) As String

    Dim I As Integer
    
    On Error Resume Next
    For I = 1 To Arguments.Count
        If GetCmdArgs <> "" Then
            GetCmdArgs = GetCmdArgs & ", "
        End If
        GetCmdArgs = GetCmdArgs & ArgValueAsString(Arguments(I).Value)
    Next I
    If GetCmdArgs <> "" Then
        GetCmdArgs = " " & GetCmdArgs
    End If
    
Done:
    Exit Function
    
End Function

Public Function IsSameCommand(LastCommand As Command, ReceivedMemnomic As String)

    On Error GoTo Error
    If Trim(LastCommand.TagVariant.LastCommand.CommandText) = Trim(ReceivedMemnomic) Then
        IsSameCommand = True
    End If
    GoTo Done
Error:
    Resume Done
Done:
    Exit Function
    
End Function

Public Function LoadResString(ID As Long, ParamArray ItemTextArgs() As Variant) As String

' Replace standard VB LoadResString function with this enhanced version that
' provides string substitutions and better language support.

' Look at the ID value to see which resource file to use:
'
' If the &H2000 bit is set, then get the string from the standard Talent
' set of strings

' If you are using a separate, application specific resource information only
' type DLL to store localization strings, then make sure those ID's have the
' &H1000 bit set. If this bit is set, then call the application specific
' LoadxxResString method to load the string

' If neither of these bits are set, then load the string from this
' projects resource file, but use GetResString instead of the
' built in VB function. GetResString provides better langauge support
' than LoadResString. For example, if the resource file contains German
' (Standard) and English (US) strings, and the computer is running
' German (Austrian) language, GetResString will return the German string
' whereas LoadResString will return the English string.

    If (ID And STD_TALENT_STRING_BASE) = STD_TALENT_STRING_BASE Then
        LoadResString = LoadStdTalentString(ID, ItemTextArgs)
'    ElseIf (ID And SEPARATE_RES_FILE_BASE) = SEPARATE_RES_FILE_BASE Then
'        LoadResString = LoadxxResString(ID, ItemTextArgs)
    Else 'If IsWithinIDE() Then
        LoadResString = VB.LoadResString(ID)
        LoadResString = ReplaceItemText(LoadResString, ItemTextArgs)
'    Else
'        LoadResString = GetResString(App.hInstance, ID, ItemTextArgs)
    End If
    
End Function

Public Sub LogCommMessage(ByVal MessageText As String, DataPacket, Optional MessageType As TalentProcessLib.MessageTypes = mtInformation)

'

    Dim sText As String
    Dim I As Long
    
    On Error Resume Next
    
    sText = MessageText
        
    If IsMissing(DataPacket) Then
    Else
        If Right(sText, 1) <> " " Then
            sText = sText & " "
        End If
        If VarType(DataPacket) = vbString Then
            sText = sText & """"
            For I = 1 To Len(DataPacket)
                Select Case Asc(Mid(DataPacket, I, 1))
                Case 2
                    sText = sText & "{STX}"
                Case 3
                    sText = sText & "{ETX}"
                Case 10
                    sText = sText & "{LF}"
                Case 13
                    sText = sText & "{CR}"
                Case Else
                    sText = sText & Mid(DataPacket, I, 1)
                End Select
                If I >= 200 Then
                    sText = sText & ".... " & Len(DataPacket) - 200 & " more characters"
                    Exit For
                End If
            Next I
            sText = sText & """"
        Else
            sText = sText & """<" & LenB(DataPacket) & " bytes binary data>"""
        End If
    End If

    LogMessage sText, MessageType
    
    Exit Sub

End Sub

Sub Main()

' This is the main startup routine

    Const PROC_NAME = "Main"
    Dim bRunning As Boolean
    
    On Error Resume Next
    
' Stop this instance if already running
    If App.PrevInstance Then
        If App.StartMode <> vbSModeStandalone Then
            EH.AppError = AE.InvalidActiveXLaunch
            EH.Item(1) = PROCESS_NAME
            Call EH_MsgBox
        End If
        End
    End If
    
' Create the Main Controller object
    On Error GoTo Error
        
    App.StartLogging "", vbLogAuto
    App.OleServerBusyRaiseError = True
    App.OleServerBusyTimeout = 1000
    
    Process.Title = PROCESS_NAME
    Process.EHCaption = PROCESS_NAME
    'If Not IsWithinIDE() Then
        Process.EHMode = ehLogOnly
    'Else
    '    Process.EHMode = ehDisplayAndLog
    'End If
    Process.EHMsgsToLog = ehCriticalExclamation
    Process.LogMode = logToNT
                
    GoTo Done
    
Error:
    If Err <> 364 Then
        EH.Module = MOD_NAME
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume ErrorExit

ErrorExit:
    EH.AppError = ERR_CANT_START
    EH.Module = MOD_NAME
    EH.Procedure = PROC_NAME
    Call EH_MsgBox
    On Error Resume Next
    GoTo Done

Done:
    Exit Sub
    
End Sub

Public Function RbtModesCompatible(DynoMode As Long, ThrottleMode As Long, _
                                BrakeMode As Long, RobotMode As Long) As Boolean
                                
    Const PROC_NAME = "RbtModesCompatible"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    Dim bModesOK As Boolean

    On Error GoTo Error

    bRaiseError = False

' Acceptable combinations are:
' 1. All modes offline
' 2. Dyno mode specified, all others offline (speed, force or RLS)
' 3. Dyno mode specified, and one other actuator specified as follows:
'       DynoMode=Speed        ThrottleMode=Force
'       DynoMode=Speed        ThrottleMode=Position
'       DynoMode=Speed        BrakeMode=Force
'       DynoMode=Speed        BrakeMode=PedalForce
'       DynoMode=Speed        RobotMode=Force
'       DynoMode=Speed        RobotMode=Position
'       DynoMode=Force        ThrottleMode=Speed
'       DynoMode=Force        ThrottleMode=Position
'       DynoMode=Force        RobotMode=Speed
'       DynoMode=Force        RobotMode=Position
'       DynoMode=RLS          ThrottleMode=Speed
'       DynoMode=RLS          ThrottleMode=Position
'       DynoMode=RLS          RobotMode=Speed
'       DynoMode=RLS          RobotMode=Position

    bModesOK = True

    Select Case DynoMode
        Case dmStopped
            If ThrottleMode <> tmStopped Or _
                BrakeMode <> bmStopped Or _
                RobotMode <> rmStopped Then
                
                bModesOK = False
            End If
            
        Case dmSpeed
            Select Case ThrottleMode
                Case tmForce, tmPosition
                    If BrakeMode <> bmStopped Or RobotMode <> rmStopped Then
                        bModesOK = False
                    End If
                    
                Case tmStopped
                    Select Case BrakeMode
                        Case bmForce, bmPedalForce
                            If RobotMode <> rmStopped Then
                                bModesOK = False
                            End If
                        Case bmStopped
                            Select Case RobotMode
                                Case rmForce, rmPosition, rmStopped
                                    ' OK
                                Case Else
                                    bModesOK = False
                            End Select
                        Case Else
                            bModesOK = False
                    End Select
                Case Else
                    bModesOK = False
            End Select
            
        Case dmForce
            Select Case ThrottleMode
                Case tmSpeed, tmPosition
                    If BrakeMode <> bmStopped Or RobotMode <> rmStopped Then
                        bModesOK = False
                    End If
                    
                Case tmStopped
                    Select Case BrakeMode
                        Case bmStopped
                            Select Case RobotMode
                                Case rmSpeed, rmPosition, rmStopped
                                    ' OK
                                Case Else
                                    bModesOK = False
                            End Select
                        Case Else
                            bModesOK = False
                    End Select
                Case Else
                    bModesOK = False
            End Select
        
        Case dmRLS
            Select Case ThrottleMode
                Case tmSpeed, tmPosition
                    If BrakeMode <> bmStopped Or RobotMode <> rmStopped Then
                        bModesOK = False
                    End If
                    
                Case tmStopped
                    Select Case BrakeMode
                        Case bmStopped
                            Select Case RobotMode
                                Case rmSpeed, rmPosition, rmStopped
                                    ' OK
                                Case Else
                                    bModesOK = False
                            End Select
                        Case Else
                            bModesOK = False
                    End Select
                Case Else
                    bModesOK = False
            End Select
        
        Case Else
            bModesOK = False
            
    End Select

    RbtModesCompatible = bModesOK

    GoTo Done
    
Error:
    bModesOK = False
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Err.Raise lRaiseErrorNumber, MOD_NAME & "." & PROC_NAME, sRaiseErrorText
    End If

    Exit Function
End Function

Public Function ThrModeAsText(ThrMode) As String

    Select Case ThrMode
    Case tmStopped
        ThrModeAsText = "Offline"
    Case tmSpeed
        ThrModeAsText = "Speed"
    Case tmForce
        ThrModeAsText = "Force"
    Case tmPosition
        ThrModeAsText = "Position"
    Case Else
        ThrModeAsText = "Unknown (" & ThrMode & ")"
    End Select

End Function

Public Function ThrModeFromText(ThrottleModeText As String) As ThrottleModes

'

    Const PROC_NAME = "ThrModeFromText"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    Dim I As Long
    
    On Error GoTo Error

    bRaiseError = False

    For I = tmStopped To tmPosition
        If ThrModeAsText(I) Like ThrottleModeText & "*" Then
            ThrModeFromText = I
            GoTo Done
        End If
    Next I
    
    sRaiseErrorText = LoadResString(ridBadThrMode, ThrottleModeText)
    bRaiseError = True
    lRaiseErrorNumber = tceInvalidArgs

    GoTo Done

Error:
    sRaiseErrorText = LoadResString(ridBadThrMode, Err.Description)
    bRaiseError = True
    lRaiseErrorNumber = tceInvalidArgs
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        EH_RaiseError lRaiseErrorNumber, MOD_NAME, PROC_NAME, sRaiseErrorText
    End If

    Exit Function

End Function


