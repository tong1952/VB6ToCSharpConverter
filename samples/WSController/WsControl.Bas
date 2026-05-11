Attribute VB_Name = "modWsControl"
Option Explicit
Option Compare Text

' Note that this is a multithreaded application. Do not declare any variables in
' this module and expect them to be shared throughout the application.

Public Enum ResourceIDs
    ridAuxCurrentOff = 112
    ridMFCurrentOff = 113
    ridIdleCurrentOn = 114
    ridMFCurrentOn = 115
    ridModeManIlck = 116
    ridModeTrackIlck = 118
    ridIdleNozFlapNotOpen = 119
    ridIdleBypFlapNotClosed = 120
    ridFanStoppedIlck = 121
    ridFanDriveSafetyFault = 124
    ridFanDriveFault = 125
    ridFanDriveVibWarning = 126
    ridFanDriveVibFault = 127
    ridFanDriveTempFault = 128
    ridLowSpeedLmt1On = 129
    ridLowSpeedLmt2On = 130
    ridLowSpeedLmt3On = 131
    ridModeAuto = 132
    ridModeManual = 133
    ridModeTrack = 134
    ridMfTrippedOut = 135
    ridWsTooLowForPid = 139
    ridHxDpOverQFault = 140
    ridHxDpFault = 141
    ridCommonInterlockFaulted = 153
    ridFanspeedAchieved = 155
    ridIdleStarted = 156
    ridIdleStopped = 157
    ridMainFanAuxOff = 158
    ridMainFanAuxOn = 159
    ridMainfanOff = 160
    ridMainFanOn = 161
    ridWindspeedAchieved = 162
    ridWindspeedInitialized = 163
    ridFanspeedNotAchieved = 164
    ridWindspeedNotAchieved = 165
    ridWindspeedNotInitialized = 166
    ridCantCreateObject = 167
    ridRunning = 168
    ridRunningAt = 169
    ridCommandsClassCaption = 170
    ridErrDisplayingMessage = 174
    ridPIDIntegralAtMax = 177
    ridPIDIntegralAtMin = 178
    ridFanSpeedStillAbove = 179
    ridTrackParInvalidUnits = 180
    ridTrackFailed = 185
    ridIdlePermissiveNotOn = 201
    ridFanSystemFault = 203
    ridEstop = 204
    ridFanNotinComputerCtl = 205
    ridAuxNotReady = 206
    ridAuxPerm = 207
    ridAuxFault = 208
    ridLightCurtain = 209
    ridIdleDriveFault = 210
    ridFanLCFault = 211
    ridFanMCFault = 212
    ridCSDidNotStart = 213
    ridMCDidNotGetReady = 214
    ridLCDidNotGetReady = 215
    ridErrorLoadingWSCoeffs = 216
    ridErrorOpeningFile = 217
    ridErrorReadingData = 218
    ridMainFanRestartDelay = 219
    ridPurging = 220
    ridFanIsOnStopFanFirst = 221
    ridUnableToGetControlFromPLC = 222
    ridIdleMotorStopOn = 223
    ridWindspeedReduced = 224
    ridFanspeedReduced = 225
    ridCFLeftFault = 226
    ridCFLeftLockout = 227
    ridCFRightFault = 228
    ridCFRightLockout = 229
    ridCFTopFault = 230
    ridCFTopLockout = 231
End Enum


Public Const PROCESS_NAME = "Windspeed"
Public Const PROCESS_DESCRIPTION = "Windspeed Controller"
Public Const REG_PROCESSES = "SOFTWARE\React Technologies\Talent\4.0\Processes"
Public Const REG_WINDSPEED = "SOFTWARE\React Technologies\Talent\4.0\Windspeed"

' Object, Shared property Group, and Shared Property names
Public Const SOM_CONTROLLER = "Controller"
Public Const SPG_CONSTANTS = "Constants"
Public Const SPG_DATA = "Data"

' Define unique numeric indices for all control constants that are used
' outside of the Controller class. Use these numeric indices to save/recall
' the constants (by numberic index) in the Constants Shared Property Group
Public Enum SPGConstants
    CC_FSMin = 0
    CC_FSMax = 1
    CC_FSRampMax = 2
    CC_WSMin = 3
    CC_WSMax = 4
    CC_WSRampMax = 5
    CC_HeadwindMin = 6
    CC_HeadwindMax = 7
    CC_WSRampDef = 8
    CC_FSRampDef = 9
    CC_DefAutoIdle = 10
    CC_DefTrackParName = 11
    CC_CFLeft1Min = 12
    CC_CFLeft1Max = 13
    CC_CFLeft2Min = 14
    CC_CFLeft2Max = 15
    CC_CFRight1Min = 16
    CC_CFRight1Max = 17
    CC_CFRight2Min = 18
    CC_CFRight2Max = 19
    CC_CFTopMin = 20
    CC_CFTopMax = 21
End Enum

Public Enum Opcodes
    ocInitialize = 1
    ocSuspend = 2
    ocShutdown = 3
    ocShowStatus = 4
    ocStartMainFanAux = 5
    ocStopMainFanAux = 6
    ocStartMainFan = 7
    ocStopMainFan = 8
    ocSetFanSpeed = 9
    ocSetWindSpeed = 10
    ocTrackWindSpeed = 11
    
    ocStartIdle = 12
    ocStopIdle = 13
    ocShowIdleStatus = 14
    ocCancelIdle = 15
    
    ocSetModeManual = 20
    ocSetModeAuto = 21
    ocSetNozzleMethod = 22
    ocSetPlenumMethod = 23
    
    ocResetFaults = 30
    
    ocStartPurge = 40
    ocStopPurge = 41
    ocSetNozzleSize = 42
    ocStopNozzleMotion = 43
    
    ocSetCFlapLeft = 50
    ocSetCFlapLeft1 = 51
    ocSetCFlapLeft2 = 52
    ocSetCFlapRight = 53
    ocSetCFlapRight1 = 54
    ocSetCFlapRight2 = 55
    ocSetCFlapTop = 56
    ocStopCF = 57
End Enum

Public Enum Modes
    modeOff = 0
    modeManual = 1
    modeAuto = 2
    modeTrack = 3
End Enum

Public Const SPG_STATUSES = "Statuses"
Public Const STS_FANSPEED_ACHIEVED = "FanspeedAchieved"
Public Const STS_IDLE_STARTED = "IdleStarted"
Public Const STS_IDLE_STOPPED = "IdleStopped"
Public Const STS_MAINFAN_AUX_ON = "MainfanAuxOn"
Public Const STS_MAINFAN_AUX_OFF = "MainfanAuxOff"
Public Const STS_MAINFAN_ON = "MainfanOn"
Public Const STS_MAINFAN_OFF = "MainfanOff"
Public Const STS_WINDSPEED_ACHIEVED = "WindspeedAchieved"
Public Const STS_WINDSPEED_INITIALIZED = "WindspeedInitialized"

Public Enum IdleStatuses
    isNormal = 0
    isIdle = 1
    isMoving = 2
    isStopped = 3
    isFault = 4
End Enum

Public Enum NozzleSizes
    nsMoving = 0
    nsLarge = 1
    nsMedium = 2
    nsSmall = 3
End Enum


Public Const FAULT = 0#
Public Const NOFAULT = 1#

Public Const Rair As Double = 287.06
Public Const Rwater As Double = 461.04

Public Const DY_MODE_SPEED = 1
Public Const THR_MODE_SPEED = 1

Public Enum FilterSpeeds
    fsSLOW = 0
    fsFAST = 1
End Enum

Public Type FinalRampCC
    RampRate As Single
    TargetChangeBase As Single
    TargetChangeAdjust As Single
    Min As Single
    Max As Single
End Type

Public Declare Function GetCurrentProcessId Lib "KERNEL32" () As Long
Public Declare Function GetCurrentProcess Lib "KERNEL32" () As Long
Public Declare Function TerminateProcess Lib "KERNEL32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Function NozzleName(Index As Long) As String
    If Index = 2 Then
        NozzleName = "Medium"
    ElseIf Index = 3 Then
        NozzleName = "Small"
    Else
        NozzleName = "Large"
    End If
End Function
Public Function CalcFinalRampDeltakph(RampRate As Single, TargetChange As Single, NumRampRates As Long, RampRateCCs() As FinalRampCC) As Single

'

    Const PROC_NAME = "CalcFinalRampDeltakph"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    Dim fI As Single
    Dim DeltakphI As Single
    Dim DeltakphIm1 As Single
    Dim I As Long
    Dim colRRs As Collection
    Dim colIndex As Collection
    Dim J As Long
    Dim Im1 As Long
    
    On Error GoTo Error

' Sort the array
    Set colRRs = New Collection
    Set colIndex = New Collection
    colRRs.Add RampRateCCs(1).RampRate
    colIndex.Add 1
    For I = 2 To NumRampRates
        For J = 1 To colRRs.Count
            If RampRateCCs(I).RampRate <= colRRs(J) Then
                colRRs.Add RampRateCCs(I).RampRate, , J
                colIndex.Add I, , J
                Exit For
            ElseIf J = colRRs.Count Then
                colRRs.Add RampRateCCs(I).RampRate
                colIndex.Add I
                Exit For
            End If
        Next J
    Next I
        
    bRaiseError = False
    For J = 1 To NumRampRates
        If RampRate <= colRRs(J) Then
            If J = 1 Then
                fI = 1#
                I = colIndex(J)
                CalcFinalRampDeltakph = Ctl_Limit(RampRateCCs(I).TargetChangeBase + (RampRateCCs(I).TargetChangeAdjust * Abs(TargetChange)), RampRateCCs(I).Min, RampRateCCs(I).Max)
                GoTo Done
            Else
                I = colIndex(J)
                Im1 = colIndex(J - 1)
                fI = (RampRate - RampRateCCs(Im1).RampRate) / (RampRateCCs(I).RampRate - RampRateCCs(Im1).RampRate)
                DeltakphI = Ctl_Limit(RampRateCCs(I).TargetChangeBase + (RampRateCCs(I).TargetChangeAdjust * Abs(TargetChange)), RampRateCCs(I).Min, RampRateCCs(I).Max)
                DeltakphIm1 = Ctl_Limit(RampRateCCs(Im1).TargetChangeBase + (RampRateCCs(Im1).TargetChangeAdjust * Abs(TargetChange)), RampRateCCs(Im1).Min, RampRateCCs(Im1).Max)
                CalcFinalRampDeltakph = (fI * DeltakphI) + ((1# - fI) * DeltakphIm1)
                GoTo Done
            End If
        End If
    Next J

    I = colIndex(NumRampRates)
    CalcFinalRampDeltakph = Ctl_Limit(RampRateCCs(I).TargetChangeBase + (RampRateCCs(I).TargetChangeAdjust * Abs(TargetChange)), RampRateCCs(I).Min, RampRateCCs(I).Max)

    GoTo Done

Error:
    sRaiseErrorText = Err.Description
    bRaiseError = True
    lRaiseErrorNumber = Err.Number
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = "modWsControl"
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        EH_RaiseError lRaiseErrorNumber, "modWsControl", PROC_NAME, sRaiseErrorText
    End If

    Exit Function

    
    
End Function
Public Function TableLookup(ByVal InputValue As Single, InputValues() As Single, OutputValues() As Single, ArrayMaxSize As Long) As Single

    Dim I As Long
    Dim fFactor As Single
    
    On Error GoTo Error
    
    TableLookup = OutputValues(LBound(OutputValues))
    
    For I = LBound(InputValues) To ArrayMaxSize
        If InputValue <= InputValues(I) Then
            If I = LBound(InputValues) Then
                TableLookup = OutputValues(ArrayMaxSize)
            Else
                fFactor = (InputValue - InputValues(I - 1)) / (InputValues(I) - InputValues(I - 1))
                TableLookup = OutputValues(I - 1) + fFactor * (OutputValues(I) - OutputValues(I - 1))
            End If
            Exit For
        ElseIf I = ArrayMaxSize Then
            TableLookup = OutputValues(I)
        End If
    Next I
    
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
    ElseIf IsWithinIDE() Then
        LoadResString = VB.LoadResString(ID)
        LoadResString = ReplaceItemText(LoadResString, ItemTextArgs)
    Else
        LoadResString = GetResString(App.hInstance, ID, ItemTextArgs)
    End If
    
End Function

Public Function Log10(Number As Single) As Single

    Log10 = Log(Number) / Log(10)
    
End Function

Public Function LimitD(Expression As Double, Min As Single, Max As Single) As Double

    If Expression < Min Then
        LimitD = Min
    ElseIf Expression > Max Then
        LimitD = Max
    Else
        LimitD = Expression
    End If
    
End Function
