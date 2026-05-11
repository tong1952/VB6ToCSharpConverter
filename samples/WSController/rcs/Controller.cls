head 1.25;
branch ;
access ;
symbols CD%204%2E00%2E0021:1.25 CD%204%2E00%2E0020:1.25 4%2E00%2E0092:1.25
        4%2E00%2E0091:1.24 CD%204%2E00%2E0019:1.23 4%2E00%2E0090:1.23 CD%204%2E00%2E0018:1.22
        CD%204%2E00%2E0017:1.22 CD%204%2E00%2E0016:1.22 4%2E00%2E0088:1.22 4%2E00%2E0082:1.21
        4%2E00%2E0080:1.20 4%2E00%2E0078:1.19 4%2E00%2E0077:1.18 CD%204%2E00%2E0015:1.17
        4%2E00%2E0075:1.17 CD%204%2E00%2E0014:1.16 CD%204%2E00%2E0013:1.16 4%2E00%2E0074:1.16
        4%2E00%2E0066:1.15 4%2E00%2E0065:1.14 4%2E00%2E0062:1.13 4%2E00%2E0058:1.12
        4%2E00%2E0057:1.11 4%2E00%2E0053:1.10 4%2E00%2E0052:1.9 4%2E00%2E0049:1.8
        4%2E00%2E0044:1.7 4%2E00%2E0041:1.6 4%2E00%2E0040:1.5 4%2E00%2E0039:1.4 4%2E00%2E0037:1.3
        4%2E00%2E0034:1.2;
locks ; strict;
comment @@;


1.25
date 2024.03.26.18.40.21; author Kavanagh; state Develop;
branches ;
next 1.24;

1.24
date 2024.03.26.18.35.56; author Kavanagh; state Develop;
branches ;
next 1.23;

1.23
date 2020.01.21.03.22.07; author Kavanagh; state Develop;
branches ;
next 1.22;

1.22
date 2019.11.21.14.54.44; author Kavanagh; state Develop;
branches ;
next 1.21;

1.21
date 2019.11.20.23.27.13; author Kavanagh; state Develop;
branches ;
next 1.20;

1.20
date 2019.11.14.21.35.46; author Kavanagh; state Develop;
branches ;
next 1.19;

1.19
date 2019.10.15.22.23.49; author Kavanagh; state Develop;
branches ;
next 1.18;

1.18
date 2019.10.10.21.24.16; author Kavanagh; state Develop;
branches ;
next 1.17;

1.17
date 2019.09.29.06.24.24; author Kavanagh; state Develop;
branches ;
next 1.16;

1.16
date 2019.09.17.05.47.01; author Kavanagh; state Develop;
branches ;
next 1.15;

1.15
date 2019.09.17.01.59.31; author Kavanagh; state Develop;
branches ;
next 1.14;

1.14
date 2019.09.16.01.26.49; author Kavanagh; state Develop;
branches ;
next 1.13;

1.13
date 2019.09.11.21.31.24; author Kavanagh; state Develop;
branches ;
next 1.12;

1.12
date 2019.09.09.06.32.35; author Kavanagh; state Develop;
branches ;
next 1.11;

1.11
date 2019.09.09.05.06.45; author Kavanagh; state Develop;
branches ;
next 1.10;

1.10
date 2019.09.09.02.57.45; author Kavanagh; state Develop;
branches ;
next 1.9;

1.9
date 2019.09.09.02.05.10; author Kavanagh; state Develop;
branches ;
next 1.8;

1.8
date 2019.09.06.23.48.16; author Kavanagh; state Develop;
branches ;
next 1.7;

1.7
date 2019.07.04.05.31.25; author Kavanagh; state Develop;
branches ;
next 1.6;

1.6
date 2019.07.04.00.43.47; author Kavanagh; state Develop;
branches ;
next 1.5;

1.5
date 2019.07.02.21.21.10; author Kavanagh; state Develop;
branches ;
next 1.4;

1.4
date 2019.04.26.02.44.33; author Kavanagh; state Develop;
branches ;
next 1.3;

1.3
date 2019.04.22.04.35.43; author Kavanagh; state Develop;
branches ;
next 1.2;

1.2
date 2019.04.17.03.31.36; author Kavanagh; state Develop;
branches ;
next 1.1;

1.1
date 2019.04.16.23.31.55; author Kavanagh; state Develop;
branches ;
next ;

ext
@project Z:/Client Facilities/CAERI/Source Code/Applications/WSController/CWTCSWSController.pj;
@

desc
@@


1.25
log
@LoopTimer@
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
Private mobjStats As AdvUtilLib.Stats
Private WithEvents mobjParManager As ParManager
Attribute mobjParManager.VB_VarHelpID = -1
Private WithEvents tmrMain As LoopTimer
Attribute tmrMain.VB_VarHelpID = -1
Private WithEvents tmrRunOnce As ccrpTimer
Attribute tmrRunOnce.VB_VarHelpID = -1
Private tmrElapsed As ccrpStopWatch
Private WithEvents mobjInterlocks As Interlocks
Attribute mobjInterlocks.VB_VarHelpID = -1

Private mobjWSFilter As ControlLib.Filter
Private mobjWSFilter2 As ControlLib.Filter
Private mobjWSFilterBW As ConversionLib.Butterworth
Private mobjWSFilterBW2 As ConversionLib.Butterworth
Private mobjWSctlFilter As ControlLib.Filter
Private mobjTrackFilter As ControlLib.Filter
Private mobjFFLeadLag As ControlLib.Filter
Private mobjRampRateFilter As ControlLib.Filter
Private mobjOutputFilter As ControlLib.Filter
Private mobjWSMonitor As ControlLib.Monitor
Private mobjFSMonitor As ControlLib.Monitor
Private mobjPID As ControlLib.PIDControl
Private mobjTrackPar As ParametersLib.Parameter

Private mobjEvents As TalentProcessLib.Events
Private mobjForm As frmControl
Private mobjSPGConstants As SharedPropertyGroup
Private mcolInvalidOpcodes As Collection
Private mobjMFPermComponents As PermissiveComponents
Private mobjMFAuxPermComponents As PermissiveComponents
Private mobjIdleCmdPermComponents As PermissiveComponents

Implements ICcrpTimerNotify

Private mlID As Long
Private mbReady As Boolean
Private mbInited As Boolean
Private Const mbKeepStats As Boolean = True
Private mvCheckIlckValue As Variant

'Module level control constants
Private mccsPriority As String
Private mccfSlowFilterDelayTime As Single
Private mccfWsTgtTrackFilTc As Single
Private mccbFilterInput As Boolean
Private mccbFilterOutput As Boolean
Private mccfPidOutputFilTc As Single
Private mccfMinFilTc As Single
Private mccfMaxFilTc As Single
Private mccfMainFanTimeout As Single
Private mccfIdleTimeout As Single
Private mccfMainFanAuxTimeout As Single
Private mccfMainFanAuxPause As Single
Private mccfMainFanAuxAutoOffTime As Single
Private mccfResetFaultTime As Single
Private mccbAutoReset As Boolean        ' If True, then reset the fan when the main ResetFaults command is used
Private mccfMainFanRestartTime As Single

Private mcclNumCloseTimes(1 To 3) As Long
Private mccfCloseTimesLarge(1 To 10) As Single
Private mccfCloseTimeSpeedsLarge(1 To 10) As Single
Private mccfCloseTimesMedium(1 To 10) As Single
Private mccfCloseTimeSpeedsMedium(1 To 10) As Single
Private mccfCloseTimesSmall(1 To 10) As Single
Private mccfCloseTimeSpeedsSmall(1 To 10) As Single

Private mccfUpdateInt As Single
Private mccfCloseWsRamp(1 To 3) As Single
Private mccfCloseWsFilTc(1 To 3) As Single
Private mccfCtlMax As Single
Private mccfCtlMin As Single
Private mccfDisplayFilType As Single
Private mccfDefFsRamp As Single
Private mccfDefWsRamp As Single
Private mccfFfGain As Single
Private mccfFfLagCoef As Single
Private mccfFfLagCoefTrack As Single
Private mccfFfLeadCoef As Single
Private mccfFfLeadCoefTrack As Single
Private mccfWsFilterTc0(1 To 3) As Single
Private mccfWsFilterTc1(1 To 3) As Single
Private mccfTdFilterTC As Single
Private mccfWsFilterFast(1 To 3) As Single
Private mccfWsFilterFastTrack(1 To 3) As Single

Private mccfNomFs(1 To 3) As Single
Private mccfMaxFs(1 To 3) As Single
Private mccfMaxFsPrepMode As Single
Private mccfMaxFsStabMode As Single

Private mccfFsMonitorError As Single
Private mccfFsMonitorFilTc As Single
Private mccfFsMonitorPeriod As Single

Private mccfIntMax As Single
Private mccfIntMin As Single

Private mccfKp0(1 To 3) As Single
Private mccfKp1(1 To 3) As Single
Private mccfKp0Track(1 To 3) As Single
Private mccfKp1Track(1 To 3) As Single
Private mccfKpMin(1 To 3) As Single
Private mccfKpMax(1 To 3) As Single

Private mccfTi0(1 To 3) As Single
Private mccfTi1(1 To 3) As Single
Private mccfTi0Track(1 To 3) As Single
Private mccfTi1Track(1 To 3) As Single
Private mccfTiMin(1 To 3) As Single
Private mccfTiMax(1 To 3) As Single

Private mccfTd0(1 To 3) As Single
Private mccfTd1(1 To 3) As Single
Private mccfTd0Track(1 To 3) As Single
Private mccfTd1Track(1 To 3) As Single
Private mccfTdMin(1 To 3) As Single
Private mccfTdMax(1 To 3) As Single


Private mccfMinDeltaP As Single

Private mccfMaxFsRamp As Single
Private mccfMaxFsRampTrack As Single
Private mccfMaxSuspendTime As Single
Private mccfMinWsHeadwind As Single
Private mccfMaxWsHeadwind As Single
Private mccfMinWsRamp As Single
Private mccfMaxWsRamp As Single
Private mccfMaxWsError As Single
Private mccfMinFs As Single
Private mccfMinFsAutoMode As Single
Private mccfMinFsRamp As Single
Private mccfMinDeltaPFSMax As Single
Private mccfMinWs As Single
Private mccfMinWsForPid As Single
Private mccfMinWsForPidDelay As Single
Private mccfMinWsErrorSlow As Single
Private mccfMinWsErrorSlowSmall As Single
Private mccfMinDeltaPHxRatio As Single
Private mccfMonitorError As Single
Private mccfMonitorFilTc As Single
Private mccfMonitorPeriod As Single
Private mccfSlowRampFanspeed As Single
Private mccfMaxWsRampLowSpeeds As Single

Private mccfNomWs(1 To 3) As Single
Private mccfMaxWs(1 To 3) As Single
Private mccfMaxWsRateForI As Single
Private mccfMinWsRateForI As Single
Private mccfWsRateForI0 As Single
Private mccfWsRateForI1 As Single
Private mccfMaxWsPrepMode As Single
Private mccfMaxWsStabMode As Single

Private mccbDefAutoIdle As Boolean
Private mccfWSStartIdle As Single
Private mccfWSStopIdle As Single
Private mccfAutoIdleDeadband As Single
Private mccfIdleFanspeed As Single

Private mccfControlFilTc0(1 To 3) As Single
Private mccfControlFilTc1(1 To 3) As Single
Private mccfControlFilTcMin(1 To 3) As Single
Private mccfControlFilTcMax(1 To 3) As Single
Private mccfDrivePause As Single
Private mccsMFPermComponents As String
Private mccsMFAuxPermComponents As String
Private mccsIdleCmdPermComponents As String
Private mccfCancelIdleTime As Single

Private mccsDefTrackParName As String
Private mccfMaxFsLimit(1 To 3) As Single
Private mccfTurnOffFanTime As Single
Private mccfMaxRunTimeWithoutCooling As Single
Private mccbAutoFollowDyTarget As Boolean
Private mccbAutoFollowDyTargetUseTrackCoeffs As Boolean
Private mccbAutoIdleUsesLimitCommand As Boolean
Private mgvbWsTooLowForPID As Boolean
Private mgvdWsTooLowForPIDTime As Date
Private mccfNomWSReverse As Single
Private mccfNomFSReverse As Single
Private mccfTimeAtZeroToForceWSZero As Single

Private mccfUseTargetErrorForDerivLimit As Single
Private mccfFinalRampDeltakph As Single
Private mccfFinalRampTimeSec As Single
Private mccfFinalRampDeltakphSmall As Single
Private mccfFinalRampTimeSecSmall As Single
Private mccbFinalRampSnap As Boolean
Private mcclSetpointRampMethod As Long
Private mccbFinalRampUseSetpoint As Boolean
Private mccfWsRateForFinalRamp As Single
Private mccbFinalRampSnapBumpless As Boolean

Private mccsWCoeffsFile(1 To 3) As String
Private mccdKpCoeffs(1 To 3, 0 To 6) As Double
Private mccdKqnCoeffs(1 To 3, 0 To 6) As Double
Private mccdKqplCoeffs(1 To 3, 0 To 6) As Double
Private mccdWSvsFSCoeffs(1 To 3, 0 To 6) As Double
Private mccfKpCoeffMin(1 To 3) As Single
Private mccfKpCoeffMax(1 To 3) As Single
Private mccfKqnCoeffMin(1 To 3) As Single
Private mccfKqnCoeffMax(1 To 3) As Single
Private mccfKqplCoeffMin(1 To 3) As Single
Private mccfKqplCoeffMax(1 To 3) As Single

Private mccbOFDefault(1 To 18) As Boolean
Private mccbOFPurgeOpen As Boolean
Private mccfNozzleTimeout As Single

Private mccfCFLeft1Min As Single
Private mccfCFLeft1Max As Single
Private mccfCFLeft2Min As Single
Private mccfCFLeft2Max As Single
Private mccfCFRight1Min As Single
Private mccfCFRight1Max As Single
Private mccfCFRight2Min As Single
Private mccfCFRight2Max As Single
Private mccfCFTopMin As Single
Private mccfCFTopMax As Single
Private mccfCFTimeout As Single
Private mccfCFHoldTime As Single
Private mccfCFTolerance As Single

'Module level Inputs
Private mipAirtemp As Single
Private mipAirtempPlenum As Single
Private mipRelHumid As Single
Private mipDewpoint As Single
Private mipPabs As Single
Private mipDeltaP As Single
Private mipDeltaPPlenum As Single

Private mipEstopStatus As Single
Private mipToPLCResetCmd As Single
Private mipOpModeOff As Single
Private mipOpModeMaint As Single
Private mipOpModePrep As Single
Private mipOpModeStab As Single     ' AAWT Only
Private mipOpModeTest As Single
Private mipAuxDPLCCmdRequest As Single
Private mipCoolPLCCmdRequest As Single

Private mipMfPerm As Single
Private mipMfRemote As Single
Private mipMfLCFault As Single
Private mipMfMCFault As Single
Private mipCSStatus As Single
Private mipLCReady As Single
Private mipMCReady As Single
Private mipMfStatus As Single
Private mipFanspeed As Single
Private mipMFSpeedLimitCommand(1 To 3) As Single
Private mipMFSpeedLimitStatus(1 To 3) As Single
Private mipAuxFault As Single
Private mipAuxReady As Single
Private mipNozzleSmallLimSw As Single
Private mipNozzleMediumLimSw As Single
Private mipNozzleLargeLimSw As Single

Private mipIdleNozFlapMotorStop As Single               ' Idle condition is Nozzle flap closed (blocks the nozzle) and bypass flaps open (to allow bypass flow)
Private mipIdleNozFlapOpenStatus As Single              ' Normal condition is Nozzle Flap open and bypass flaps closed
Private mipIdleNozFlapClosedStatus As Single
Private mipIdleNozFlapMotorFault As Single
Private mipIdleBypFlapOpenStatus(1 To 3) As Single
Private mipIdleBypFlapClosedStatus(1 To 3) As Single

Private mipCFLeft1Pos As Single
Private mipCFLeft1Fault As Single
Private mipCFLeft1Lockout As Single
Private mipCFLeft2Pos As Single
Private mipCFLeft2Fault As Single
Private mipCFLeft2Lockout As Single
Private mipCFRight1Pos As Single
Private mipCFRight1Fault As Single
Private mipCFRight1Lockout As Single
Private mipCFRight2Pos As Single
Private mipCFRight2Fault As Single
Private mipCFRight2Lockout As Single
Private mipCFTopPos As Single
Private mipCFTopFault As Single
Private mipCFTopLockout As Single

Private mipBLFanStatus As Single
Private mipDySpeedTarget As Single
Private mipDySpeedRate As Single
Private mipDyMode As Single
Private mipThrMode As Single

'Module level Outputs
Private mopWsControl As Single
Private mopFilDeltaP As Single
Private mopFilDeltaPPlenum As Single

Private mopAutoIdleStatus As Single
Private mopIdleStatus As Single
Private mopIdleStopControl As Single
Private mopIdleStartControl As Single

Private mopFsSetpt As Single
Private mopAuxControl As Single
Private mopMfControl As Single
Private mopMFReset As Single
Private mopAirDensityKgm3 As Single
Private mopBlockageCorr As Single
Private mopFsRamp As Single
Private mopFsTargetDisp As Single
Private mopPstaticKpa As Single
Private mopQPa As Single
Private mopQnPa As Single
Private mopQplPa As Single
Private mopWSCalcMethod As Single
Private mopKp As Single
Private mopKqpl As Single
Private mopKqn As Single

Private mopWindspeed As Single
Private mopWsRamp As Single
Private mopWsSetpt As Single
Private mopWsSetptDisp As Single
Private mopWsTargetDisp As Single
Private mopHeadwind As Single
Private mopWsAchieved As Single

Private mopWindspeedmps As Single
Private mopWsSetptmps As Single
Private mopWsTargetmps As Single
Private mopWsRampmps As Single
Private mopHeadwindmps As Single

Private mopWsMode As Single
Private mopNozzle As Single
Private mopCFLeftStartCmd As Single
Private mopCFLeftStopCmd As Single
Private mopCFLeft1Target As Single
Private mopCFLeft2Target As Single
Private mopCFRightStartCmd As Single
Private mopCFRightStopCmd As Single
Private mopCFRight1Target As Single
Private mopCFRight2Target As Single
Private mopCFTopStartCmd As Single
Private mopCFTopStopCmd As Single
Private mopCFTopTarget As Single

Private mopOFOpenCmd(1 To 18) As Single
Private mopPurgeCmd As Single
Private mopNozzleCmd(1 To 3) As Single

'Module level Global variables
Private mgvlDefPriority As Priorities
Private mgvlWindTunnel As WindTunnels
Private mgvfCloseWsFilTc As Single
Private mgvfPurgeStatus As Single
Private mgvfHeadwind As Single
Private mgvfControlRamp As Single
Private mgvfControlFilTc As Single
Private mgvfCloseTime As Single
Private mgvfCloseTimeRamp As Single
Private mgvfFilterTmr As Single
Private mgvfFbNormFs As Single
Private mgvfFfNormFs As Single
Private mgvfFfRawNormFs As Single

Private mgvfAuxStatus As Single
Private mgvfFilDeltaPControl As Single
Private mgvfFilHxDeltaP As Single
Private mgvfFsRamp As Single
Private mgvfFsIncrement As Single
Private mgvfFsNormSetpt As Single
Private mgvfFsTarget As Single
Private mgvfFsLimitedTarget As Single
Private mgvfHxDpOverDp As Single
Private mgvfMinKp As Single
Private mgvfMaxKp As Single
Private mgvfMinTi As Single
Private mgvfMaxTi As Single
Private mgvfMinTd As Single
Private mgvfMaxTd As Single
Private mgvfMinBlockage As Single
Private mgvfMaxBlockage As Single
Private mgvfMaxFs As Single
Private mgvfMaxFsTarget As Single
Private mgvfMaxWs As Single
Private mgvfMfStartDelay As Single
Private mgvfNomFsMax As Single
Private mgvfNomWsMax As Single
Private mgvfNormError As Single
Private mgvfNormWsSetpt As Single
Private mgvfOldRemoteStatus As Single
Private mgvfOldCtlMode As Single
Private mgvfOldMfStatus As Single
Private mgvfOldAuxStatus As Single
Private mgvlOldNozzle As NozzleSizes
Private mgvlNozzle As NozzleSizes
Private mgvfOldPurgeStopFanRequest As Single
Private mgvfOldTsDoor4 As Single
Private mgvfOldWsIncrement As Single
Private mgvfPressureRatio As Single
Private mgvfPstaticControlKpa As Single
Private mgvfQControlPa As Single
Private mgvfRawHxDpOverDp As Single
Private mgvfKp As Single
Private mgvfKp0 As Single
Private mgvfKp1 As Single
Private mgvfKpEffect As Single
Private mgvfTd As Single
Private mgvfTd0 As Single
Private mgvfTd1 As Single
Private mgvfDeriviativeEffect As Single
Private mgvfTi As Single
Private mgvfTi0 As Single
Private mgvfTi1 As Single
Private mgvfIntegratorEffect As Single
Private mgvfOldIntegral As Single
Private mgvfWsError As Single
Private mgvfWsFilTc As Single
Private mgvfWsFilTc0 As Single
Private mgvfWsFilTc1 As Single
Private mgvfWsIncrement As Single
Private mgvfWsRamp As Single
Private mgvfWsRateForI As Single
Private mgvfWsTarget As Single
Private mgvfWSTargetLimited As Single

Private mgvlFilterSpeed As FilterSpeeds
Private mgvbSeparateControl As Boolean
Private mgvdAuxOnTime As Date
Private mgvbAuxAutoOff As Boolean
Private mgvdDriveOnTime As Date
Private mgvdDriveOffTime As Date
Private mgvbAutoIdle As Boolean
Private mgvsTrackParName As String
Private mgvdLastIdleTime As Date
Private mgvfTPlAbs As Single
Private mgvfOldWsTarget As Single
Private mgvfFfLeadCoeff As Single
Private mgvfFfLagCoeff As Single
Private mgvlOldIdleStatus As IdleStatuses
Private mgvfWsFilterFast As Single
Private mgvfControlFilTc0 As Single
Private mgvfControlFilTc1 As Single
Private mgvfControlFilTcMin As Single
Private mgvfControlFilTcMax As Single
Private mgvfTrackC0 As Single
Private mgvfTrackC1 As Single
Private mgvdWDTTime As Date
Private mgvdTurnOffFanTime As Date
Private mgvfGlycolPumpInRemoteOrOn As Single
Private mgvfGlycolSystemInRemoteOrOn As Single
Private mgvdLastAuxOnTime As Date
Private mgvbStopAutoIdle As Boolean
Private mgvbAutoIdleSpeedHasBeenHigh As Boolean
Private mgvbAutoIdleSpeedHasBeenLow As Boolean
Private mgvbAutoIdleWantsIdle As Boolean
Private mgvlOIdIdleFlapsControl As Long
Private mgvlWSControlMode As CWTCSWSController.Modes
Private mgvdLastTimeFSTargetNotZero As Date
Private mgvdLastTimeWSTargetNotZero As Date
Private mgvbUsingCloseTime As Boolean
Private mgvbWSComingBackToSetpt As Boolean
Private mgvbUsingFinalRamp As Boolean
Private mgvfWSTargetError As Single
Private mgvfTargetNormError As Single
Private mgvbFinalRampHasSnapped As Boolean
Private mgvbFinalRampOKToUse As Boolean
Private mgvlNozzleTarget As NozzleSizes
Private mgvdNozzleSizeCmdTime As Date
Private mgvdMFInterlockedTime As Date
Private mgvbHoldCRUntilFSOK As Boolean

Private mgvdCFLeftCmdTime As Date
Private mgvbCFLeftCtl As Boolean
Private mgvbCFLeft1Achieved As Boolean
Private mgvbCFLeft2Achieved As Boolean
Private mgvdCFLeft1AchievedTime As Date
Private mgvdCFLeft2AchievedTime As Date
Private mgvdCFRightCmdTime As Date
Private mgvbCFRightCtl As Boolean
Private mgvbCFRight1Achieved As Boolean
Private mgvbCFRight2Achieved As Boolean
Private mgvdCFRight1AchievedTime As Date
Private mgvdCFRight2AchievedTime As Date
Private mgvdCFTopCmdTime As Date
Private mgvbCFTopCtl As Boolean
Private mgvbCFTopAchieved As Boolean
Private mgvdCFTopAchievedTime As Date

Private mgvfKpCoeffMin As Single
Private mgvfKpCoeffMax As Single
Private mgvfKqnCoeffMin As Single
Private mgvfKqnCoeffMax As Single
Private mgvfKqplCoeffMin As Single
Private mgvfKqplCoeffMax As Single

Private mgvdVmps As Double
Private mgvdWSvsFSCoeffs(0 To 6) As Double
Private mgvdKpCoeffs(0 To 6) As Double
Private mgvdKqnCoeffs(0 To 6) As Double
Private mgvdKqplCoeffs(0 To 6) As Double
Private mgvdKp As Double
Private mgvdKqn As Double
Private mgvdKqpl As Double
Private mgvdPoSCPa As Double
Private mgvdPoTSPa As Double
Private mgvdTSCAbs As Double
Private mgvdTTSAbs As Double
Private mgvdXPPl As Double

'Wind tunnel Specific module level vars
Private Enum InitTypes
    initAll = 0
    initConstants = 1
    initPars = 2
End Enum

Public Property Get ParManager() As Object
    Set ParManager = mobjParManager
End Property

Public Function GetShowStatusText() As String
'
    On Error Resume Next
    mobjInterlocks(CStr(ocShowStatus)).Check
    If Err Then
        GetShowStatusText = LoadResString(stsTheFollowingFaultsExist, Err.Description)
    Else
        GetShowStatusText = LoadResString(stsNoFaultsDetected)
    End If
    Exit Function
End Function

Private Sub OutputConversions()

'Convert/calcs to set outputs
    
    Dim I As Long
    
    On Error Resume Next
        

    ' Update displayed target values to be actual target values if
    ' the corresponding command is not currently being executed
    If mobjCommands.Exists(ocSetFanSpeed) = False Then
        mopFsTargetDisp = mgvfFsTarget
        mopFsRamp = mgvfFsRamp
    End If
    If mobjCommands.Exists(ocSetWindSpeed) = False Then
        mopWsTargetDisp = mgvfWsTarget
        mopWsRamp = mgvfWsRamp
    End If
    If mobjCommands.Exists(ocTrackWindSpeed) = False Then
        mopHeadwind = mgvfHeadwind
    End If
            
    If mopWsMode >= modeTrack Then
        If mgvbAutoIdle Then
            mopAutoIdleStatus = 1#
        Else
            mopAutoIdleStatus = 0#
        End If
    Else
        If mccbDefAutoIdle Then
            mopAutoIdleStatus = 1#
        Else
            mopAutoIdleStatus = 0#
        End If
    End If
    
    
    If (mopWsMode = modeTrack) And (mgvbAutoIdle) Then
        If mgvbAutoIdleWantsIdle Then
            If Abs(mopFsSetpt) > mccfMaxFsLimit(1) Then
                ' Wait until setpoint < idle limit
            Else
                mopIdleStartControl = 1#
                mopIdleStopControl = 0#
            End If
        Else
            mopIdleStopControl = 1#
            mopIdleStartControl = 0#
        End If
    End If
    
    If mopWsMode = modeAuto Then
        If mobjStatuses.GetValue(STS_WINDSPEED_ACHIEVED) Then
            mopWsAchieved = 1#
        Else
            mopWsAchieved = 0#
        End If
    Else
        mopWsAchieved = 0#
    End If
    
    mopNozzle = mgvlNozzle
    
    If mgvlWindTunnel = wtCWT Then
        mopNozzleCmd(nsSmall) = 0#
        mopNozzleCmd(nsMedium) = 0#
        mopNozzleCmd(nsLarge) = 0#
        If mobjInterlocks(CStr(ocSetNozzleSize)).IsOK = False Then
            mgvdNozzleSizeCmdTime = 0#
        ElseIf (mipMfStatus = 0) And (SecondsSinceTime(mgvdNozzleSizeCmdTime) <= mccfNozzleTimeout) Then
            Select Case mgvlNozzleTarget
            Case nsSmall
                If mipNozzleSmallLimSw = 1 Then
                    mgvlNozzleTarget = 0
                Else
                    mopNozzleCmd(nsSmall) = 1#
                End If
            Case nsMedium
                If mipNozzleMediumLimSw = 1 Then
                    mgvlNozzleTarget = 0
                Else
                    mopNozzleCmd(nsMedium) = 1#
                End If
            Case nsLarge
                If mipNozzleLargeLimSw = 1 Then
                    mgvlNozzleTarget = 0
                Else
                    mopNozzleCmd(nsLarge) = 1#
                End If
            End Select
        End If
    End If
    
    If mgvlWindTunnel = wtCWT Then
        If mopIdleStartControl = 1 Then
            mopIdleStopControl = 0
            If mopIdleStatus = IdleStatuses.isIdle Then
                mopIdleStartControl = 0#
            End If
        ElseIf mopIdleStopControl = 1 Then
            mopIdleStartControl = 0
            If mopIdleStatus = IdleStatuses.isNormal Then
                mopIdleStopControl = 0#
            End If
        End If
    End If
    
    If mgvlWindTunnel = wtCWT Then
        mopCFLeftStartCmd = 0
        mopCFLeftStopCmd = 1
        If mipAuxDPLCCmdRequest = 0 Then
            mopCFLeft1Target = Ctl_Limit(mipCFLeft1Pos, mccfCFLeft1Min, mccfCFLeft1Max)
            mopCFLeft2Target = Ctl_Limit(mipCFLeft2Pos, mccfCFLeft2Min, mccfCFLeft2Max)
        ElseIf mgvbCFLeftCtl = False Then
            mgvdCFLeftCmdTime = 0
        ElseIf mobjInterlocks(CStr(ocSetCFlapLeft)).IsOK = False Then
            mgvdCFLeftCmdTime = 0#
        ElseIf (mipMfStatus = 0) And (SecondsSinceTime(mgvdCFLeftCmdTime) <= mccfCFTimeout) Then
            If (mgvbCFLeft1Achieved = False) Or (mgvbCFLeft2Achieved = False) Then
                mopCFLeftStartCmd = 1
                mopCFLeftStopCmd = 0
            ElseIf (SecondsSinceTime(mgvdCFLeft1AchievedTime) <= mccfCFHoldTime) Or (SecondsSinceTime(mgvdCFLeft2AchievedTime) <= mccfCFHoldTime) Then
                mopCFLeftStartCmd = 1
                mopCFLeftStopCmd = 0
            Else
                mgvbCFLeftCtl = False
            End If
        End If
        
        mopCFRightStartCmd = 0
        mopCFRightStopCmd = 1
        If mipAuxDPLCCmdRequest = 0 Then
            mopCFRight1Target = Ctl_Limit(mipCFRight1Pos, mccfCFRight1Min, mccfCFRight1Max)
            mopCFRight2Target = Ctl_Limit(mipCFRight2Pos, mccfCFRight2Min, mccfCFRight2Max)
        ElseIf mgvbCFRightCtl = False Then
            mgvdCFRightCmdTime = 0
        ElseIf mobjInterlocks(CStr(ocSetCFlapRight)).IsOK = False Then
            mgvdCFRightCmdTime = 0#
        ElseIf (mipMfStatus = 0) And (SecondsSinceTime(mgvdCFRightCmdTime) <= mccfCFTimeout) Then
            If (mgvbCFRight1Achieved = False) Or (mgvbCFRight2Achieved = False) Then
                mopCFRightStartCmd = 1
                mopCFRightStopCmd = 0
            ElseIf (SecondsSinceTime(mgvdCFRight1AchievedTime) <= mccfCFHoldTime) Or (SecondsSinceTime(mgvdCFRight2AchievedTime) <= mccfCFHoldTime) Then
                mopCFRightStartCmd = 1
                mopCFRightStopCmd = 0
            Else
                mgvbCFRightCtl = False
            End If
        End If
    
        mopCFTopStartCmd = 0
        mopCFTopStopCmd = 1
        If mipAuxDPLCCmdRequest = 0 Then
            mopCFTopTarget = Ctl_Limit(mipCFTopPos, mccfCFTopMin, mccfCFTopMax)
        ElseIf mgvbCFTopCtl = False Then
            mgvdCFTopCmdTime = 0
        ElseIf mobjInterlocks(CStr(ocSetCFlapTop)).IsOK = False Then
            mgvdCFTopCmdTime = 0#
        ElseIf (mipMfStatus = 0) And (SecondsSinceTime(mgvdCFTopCmdTime) <= mccfCFTimeout) Then
            If (mgvbCFTopAchieved = False) Then
                mopCFTopStartCmd = 1
                mopCFTopStopCmd = 0
            ElseIf (SecondsSinceTime(mgvdCFTopAchievedTime) <= mccfCFHoldTime) Then
                mopCFTopStartCmd = 1
                mopCFTopStopCmd = 0
            Else
                mgvbCFTopCtl = False
            End If
        End If
    
    End If
    
    If mopWsMode <> modeAuto Then
        mopWsSetptDisp = mopWsSetpt
    End If
    
    If mgvfFsTarget <> 0# Then
        mgvdLastTimeFSTargetNotZero = Now
    End If
    If mgvfWsTarget <> 0# Then
        mgvdLastTimeWSTargetNotZero = Now
    End If
    
    For I = 1 To 18
        If mccbOFDefault(I) Then
            mopOFOpenCmd(I) = 1#
        ElseIf mopPurgeCmd Then
            mopOFOpenCmd(I) = IIf(mccbOFPurgeOpen, 1#, 0#)
        Else
            mopOFOpenCmd(I) = 0#
        End If
    Next I
    
    If mccbAutoReset Then
        If mipToPLCResetCmd = 1# Then
            mopMFReset = 1#
        ElseIf mopMFReset = 1# Then
            If mobjCommands.Exists(ocResetFaults) = False Then
                mopMFReset = 0#
            End If
        End If
    End If
    
    mopPstaticKpa = mgvdPoTSPa / 1000#
    
    mopWindspeedmps = mopWindspeed / 3.6
    mopWsSetptmps = mopWsSetptDisp / 3.6
    mopWsTargetmps = mopWsTargetDisp / 3.6
    mopWsRampmps = mopWsRamp / 3.6
    mopHeadwindmps = mopHeadwind / 3.6
    
    
    Exit Sub
End Sub

Private Sub SaveOldStates()

'Save values for next time

    On Error Resume Next
    mgvfOldCtlMode = mopWsMode
    mgvfOldRemoteStatus = mipMfRemote
    mgvfOldMfStatus = mipMfStatus
    mgvfOldAuxStatus = mgvfAuxStatus
    mgvfOldWsIncrement = mgvfWsIncrement

    mgvfOldCtlMode = mopWsMode
    mgvlOldNozzle = mgvlNozzle
    mgvlOldIdleStatus = mopIdleStatus
    
    Exit Sub
End Sub
Private Sub MonitorStatuses()

'Monitor the status properties

    Const PROC_NAME = "MonitorStatuses"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long

    On Error GoTo Error
    bRaiseError = False
'Mode dependent monitoring
    Select Case mopWsMode
    Case modeAuto
        If mobjWSMonitor.MonitorParameter(mopWindspeed) Then
            If mobjCommands.Exists(Opcodes.ocSetWindSpeed) = False Then
                mobjStatuses.SetValue STS_WINDSPEED_ACHIEVED, True
            End If
        Else
            mobjStatuses.SetValue STS_WINDSPEED_ACHIEVED, False
        End If
    Case modeManual
        If mipMfStatus = 1# Then
            If mobjFSMonitor.MonitorParameter(mipFanspeed) Then
                If mobjCommands.Exists(Opcodes.ocSetFanSpeed) = False Then
                    mobjStatuses.SetValue STS_FANSPEED_ACHIEVED, True
                End If
            Else
                mobjStatuses.SetValue STS_FANSPEED_ACHIEVED, False
            End If
        ElseIf mobjStatuses.GetValue(STS_FANSPEED_ACHIEVED) = True Then
            mobjStatuses.SetValue STS_FANSPEED_ACHIEVED, False
        End If
    End Select

    If (mopWsMode <> mgvfOldCtlMode) Then   ' IF new mode
        Select Case mopWsMode
        Case modeManual
            LogMessage LoadResString(ridModeManual)
        Case modeAuto
            LogMessage LoadResString(ridModeAuto)
        Case modeTrack
            LogMessage LoadResString(ridModeTrack)
        End Select
    End If

'Mode Independent Monitoring
'---------------------------------------

    If (mgvfAuxStatus = 1#) Then
        If mgvdAuxOnTime = 0 Then
            mgvdAuxOnTime = Now
        End If
        mgvdLastAuxOnTime = Now
        mobjStatuses.SetValue STS_MAINFAN_AUX_OFF, False
        mobjStatuses.SetValue STS_MAINFAN_AUX_ON, True
    ElseIf (mgvfAuxStatus = 0#) Then
        mgvdAuxOnTime = 0
        mobjStatuses.SetValue STS_MAINFAN_AUX_OFF, True
        mobjStatuses.SetValue STS_MAINFAN_AUX_ON, False
    Else
        mobjStatuses.SetValue STS_MAINFAN_AUX_OFF, False
        mobjStatuses.SetValue STS_MAINFAN_AUX_ON, False
    End If

    If (mipMfStatus = 1#) Then
        If mgvdDriveOnTime = 0 Then
            mgvdDriveOnTime = Now
        End If
        mgvdDriveOffTime = 0
        mobjStatuses.SetValue STS_MAINFAN_OFF, False
        mobjStatuses.SetValue STS_MAINFAN_ON, True
    Else
        If mgvdDriveOnTime > 0 Then
            If mgvdDriveOffTime = 0 Then
                mgvdDriveOffTime = Now
            End If
        End If
        mgvdDriveOnTime = 0
        mobjStatuses.SetValue STS_MAINFAN_OFF, True
        mobjStatuses.SetValue STS_MAINFAN_ON, False
    End If


    If WindTunnel = wtCWT Then
        Select Case mopIdleStatus
        Case isIdle
            mobjStatuses.SetValue STS_IDLE_STARTED, True
            mobjStatuses.SetValue STS_IDLE_STOPPED, False
        Case isNormal
            mobjStatuses.SetValue STS_IDLE_STARTED, False
            mobjStatuses.SetValue STS_IDLE_STOPPED, True
        Case Else
            mobjStatuses.SetValue STS_IDLE_STARTED, False
            mobjStatuses.SetValue STS_IDLE_STOPPED, False
        End Select
        
        If mgvbCFLeftCtl Then
            If Abs(mopCFLeft1Target - mipCFLeft1Pos) <= mccfCFTolerance Then
                If mgvbCFLeft1Achieved = False Then
                    mgvdCFLeft1AchievedTime = Now
                End If
                mgvbCFLeft1Achieved = True
            Else
                mgvbCFLeft1Achieved = False
            End If
            If Abs(mopCFLeft2Target - mipCFLeft2Pos) <= mccfCFTolerance Then
                If mgvbCFLeft2Achieved = False Then
                    mgvdCFLeft2AchievedTime = Now
                End If
                mgvbCFLeft2Achieved = True
            Else
                mgvbCFLeft2Achieved = False
            End If
        Else
            mgvbCFLeft1Achieved = False
            mgvbCFLeft2Achieved = False
        End If
    
        If mgvbCFRightCtl Then
            If Abs(mopCFRight1Target - mipCFRight1Pos) <= mccfCFTolerance Then
                If mgvbCFRight1Achieved = False Then
                    mgvdCFRight1AchievedTime = Now
                End If
                mgvbCFRight1Achieved = True
            Else
                mgvbCFRight1Achieved = False
            End If
            If Abs(mopCFRight2Target - mipCFRight2Pos) <= mccfCFTolerance Then
                If mgvbCFRight2Achieved = False Then
                    mgvdCFRight2AchievedTime = Now
                End If
                mgvbCFRight2Achieved = True
            Else
                mgvbCFRight2Achieved = False
            End If
        Else
            mgvbCFRight1Achieved = False
            mgvbCFRight2Achieved = False
        End If
    
        If mgvbCFTopCtl Then
            If Abs(mopCFTopTarget - mipCFTopPos) <= mccfCFTolerance Then
                If mgvbCFTopAchieved = False Then
                    mgvdCFTopAchievedTime = Now
                End If
                mgvbCFTopAchieved = True
            Else
                mgvbCFTopAchieved = False
            End If
        Else
            mgvbCFTopAchieved = False
        End If
    
    End If
    
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
        EH_RaiseError lRaiseErrorNumber, TypeName(Me), PROC_NAME, sRaiseErrorText
    End If
    Exit Sub
End Sub

Private Sub CheckInterlocks()

    Const PROC_NAME = "CheckInterlocks"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim sExplanationText As String
    Dim objEH As EH
    Dim I As Long
    
    On Error GoTo Error
    bRaiseError = False
                        
'Auto Mode Interlock
    If mopWsMode > modeManual Then
        If Not mobjInterlocks(CStr(ocSetModeAuto)).IsOK Then    ' switch to manual mode
            On Error Resume Next
            mobjInterlocks(CStr(ocSetModeAuto)).Check
            sExplanationText = Err.Description
            On Error GoTo Error
            LogMessage LoadResString(ridModeManIlck, sExplanationText), mtWarning
            mopWsMode = modeManual
            mgvfFsTarget = mipFanspeed
            mgvfFsRamp = mccfMaxFsRamp
            If mipMfStatus = 1# Then
                mopMfControl = 1#
            Else
                mopMfControl = 0#
            End If
            If mgvfAuxStatus = 1# Then
                mopAuxControl = 1#
            Else
                mopAuxControl = 0#
            End If
        ElseIf mipMfStatus = 0# Then
            sExplanationText = LoadResString(ridMfTrippedOut)
            LogMessage LoadResString(ridModeManIlck, sExplanationText), mtWarning
            mopWsMode = modeManual
            mgvfFsTarget = 0#
            mopFsSetpt = 0#
            mopMfControl = 0#
            If mopAuxControl = 1# Then
                mopAuxControl = IIf(mgvfAuxStatus = 1#, 1#, 0#)
            Else
                mopAuxControl = 0#
            End If
        ElseIf (mopWsMode = modeTrack) And Not mobjInterlocks(CStr(ocTrackWindSpeed)).IsOK Then    ' switch to Auto mode
            On Error Resume Next
            mobjInterlocks(CStr(ocTrackWindSpeed)).Check
            sExplanationText = Err.Description
            On Error GoTo Error
            LogMessage LoadResString(ridModeTrackIlck, sExplanationText), mtWarning
            mopWsMode = modeAuto
        ElseIf (mopWsMode = modeAuto) And (mgvlWindTunnel = wtCWT) And (mopIdleStatus <> isNormal) Then         ' switch to Auto mode
            If mipIdleNozFlapOpenStatus <> 1# Then
                sExplanationText = IIf(sExplanationText <> "", ", ", "") & LoadResString(ridIdleNozFlapNotOpen)
            End If
            For I = 1 To 3
                If mipIdleBypFlapClosedStatus(I) <> 1# Then
                    sExplanationText = IIf(sExplanationText <> "", ", ", "") & LoadResString(ridIdleBypFlapNotClosed, I)
                End If
            Next I
            LogMessage LoadResString(ridModeManIlck, sExplanationText), mtWarning
            mopWsMode = modeManual
            mgvfFsTarget = 0
            mgvfFsRamp = mccfMaxFsRamp
        End If
    End If
            
    If (mopWsMode = modeManual) And (mgvlWindTunnel = wtCWT) And (mgvfFsTarget > mccfMaxFsLimit(1)) And (mopIdleStatus <> isNormal) Then
        If mipIdleNozFlapOpenStatus <> 1# Then
            sExplanationText = IIf(sExplanationText <> "", ", ", "") & LoadResString(ridIdleNozFlapNotOpen)
        End If
        For I = 1 To 3
            If mipIdleBypFlapClosedStatus(I) <> 1# Then
                sExplanationText = IIf(sExplanationText <> "", ", ", "") & LoadResString(ridIdleBypFlapNotClosed, I)
            End If
        Next I
        LogMessage LoadResString(ridFanStoppedIlck, sExplanationText), mtWarning
        mgvfFsTarget = 0
        mgvfFsRamp = mccfMaxFsRamp
    End If
    
    If (mopWsMode = modeAuto) And (mgvfWsTarget > mccfMaxWsPrepMode) And (mipOpModePrep = 1#) Then
        sExplanationText = GetCommonMessage2(cmInPrepModeMaxIs, mccfMaxWsPrepMode)
        LogMessage LoadResString(ridWindspeedReduced, mccfMaxWsPrepMode, sExplanationText), mtWarning
        mgvfWsTarget = mccfMaxWsPrepMode
        mgvfWsRamp = mccfMaxWsRamp
    End If
    If (mopWsMode = modeAuto) And (mgvfWsTarget > mccfMaxWsStabMode) And (mipOpModeStab = 1#) And (mgvlWindTunnel = wtAAWT) Then
        sExplanationText = GetCommonMessage2(cmInStabModeMaxIs, mccfMaxWsStabMode)
        LogMessage LoadResString(ridWindspeedReduced, mccfMaxWsStabMode, sExplanationText), mtWarning
        mgvfWsTarget = mccfMaxWsStabMode
        mgvfWsRamp = mccfMaxWsRamp
    End If
    If (mopWsMode = modeManual) And (mgvfFsTarget > mccfMaxFsPrepMode) And (mipOpModePrep = 1#) Then
        sExplanationText = GetCommonMessage2(cmInPrepModeMaxIs, mccfMaxFsPrepMode)
        LogMessage LoadResString(ridFanspeedReduced, mccfMaxFsPrepMode, sExplanationText), mtWarning
        mgvfFsTarget = mccfMaxFsPrepMode
        mgvfFsRamp = mccfMaxFsRamp
    End If
    If (mopWsMode = modeManual) And (mgvfFsTarget > mccfMaxFsStabMode) And (mipOpModeStab = 1#) And (mgvlWindTunnel = wtAAWT) Then
        sExplanationText = GetCommonMessage2(cmInStabModeMaxIs, mccfMaxFsStabMode)
        LogMessage LoadResString(ridFanspeedReduced, mccfMaxFsStabMode, sExplanationText), mtWarning
        mgvfFsTarget = mccfMaxFsStabMode
        mgvfFsRamp = mccfMaxFsRamp
    End If
    
    
    
    If mopMfControl = 1# Then
        If (mgvfAuxStatus <> 1#) And (SecondsSinceTime(mgvdLastAuxOnTime) >= mccfMaxRunTimeWithoutCooling) Then
''            Set objEH = New EH
''            objEH.AppError = LoadResString(AE.MainFanAuxTrippedOut)
''            objEH.Module = TypeName(Me)
''            objEH.Procedure = PROC_NAME
''            objEH.Item(1) = mccfMaxRunTimeWithoutCooling
''            LogMessage2 objEH
            mopWsMode = modeManual
            mgvfFsTarget = 0#
            mopFsSetpt = 0#
            mopMfControl = 0#
            mopAuxControl = 0#
        End If
    End If
    
    If Not mobjInterlocks(CStr(ocStartMainFan)).IsOK Then 'And (mopMFRunRequest = 1#) Then      ' Stop fan
        If mopWsMode >= modeManual Then
            mopWsMode = modeManual
            mgvfFsTarget = 0#
            If mipMfStatus = 1# Then
                If (mopMfControl = 1) Then
                    If SecondsSinceTime(mgvdMFInterlockedTime) >= 60# Then
                        mopMfControl = 0#
                    End If
                End If
            Else
                mopMfControl = 0#
            End If
            If mopAuxControl = 1# Then
                mopAuxControl = IIf(mgvfAuxStatus = 1#, 1#, 0#)
            Else
                mopAuxControl = 0#
            End If
        End If
    Else
        mgvdMFInterlockedTime = Now
    End If
    
    If Not mobjInterlocks(CStr(ocStartMainFanAux)).IsOK And (mopAuxControl = 1#) Then        ' Stop fan aux
        If mopWsMode > modeManual Then
            mopWsMode = modeManual
        End If
        mopFsSetpt = 0#
        mgvfFsTarget = 0#
        mopMfControl = 0#
        mopAuxControl = 0#
    End If
    
'Manual mode interlock
    If Not mobjInterlocks(CStr(ocSetModeManual)).IsOK And (mopWsMode > modeOff) Then ' switch to monitor mode
        mopWsMode = modeOff
        mopFsSetpt = 0#
        mgvfFsTarget = 0#
        mopMfControl = 0#
    End If

    If WindTunnel = wtCWT Then
        If Not mobjInterlocks(CStr(ocStartIdle)).IsOK Then
            If (mopWsMode = modeTrack) And mgvbAutoIdle Then
                mgvbAutoIdle = False
            End If
            mopIdleStartControl = 0#
            mopIdleStopControl = 0#
        End If
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

Private Sub CalcWindspeed()
'
    Const PROC_NAME = "CalcWindspeed"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim fDeltaP As Single
    Dim dPsatExp As Double
    Dim dPSat As Double
    Dim dPv As Double
    
    Static fLastBWFilterTC As Single
    
    On Error GoTo Error
    bRaiseError = False

    If ((mopWsMode = modeTrack) Or (mopWsMode = modeAuto)) Then
        If ((mgvfWsIncrement <> 0) And (mopWsMode = modeAuto)) Then
            mgvlFilterSpeed = fsFAST
            mgvfFilterTmr = mccfSlowFilterDelayTime
        ElseIf (Abs(mopWindspeed - mopWsSetpt) >= IIf(mgvlNozzle = nsSmall, mccfMinWsErrorSlowSmall, mccfMinWsErrorSlow)) Then
            mgvlFilterSpeed = fsFAST
            mgvfFilterTmr = mccfSlowFilterDelayTime * (1 - mgvfNormWsSetpt)
        Else
            If (mgvfFilterTmr <= 0#) Then
                mgvlFilterSpeed = fsSLOW
            Else
                mgvfFilterTmr = mgvfFilterTmr - mccfUpdateInt
            End If
        End If
    ElseIf (mopWsMode = modeManual) Then
        If (mgvfFsIncrement <> 0#) Then
            mgvlFilterSpeed = fsFAST
        ElseIf (Abs(mipFanspeed - mopFsSetpt) * mgvfNomWsMax / mgvfNomFsMax >= mccfMinWsErrorSlow) Then
            mgvlFilterSpeed = fsFAST
        Else
            mgvlFilterSpeed = fsSLOW
        End If
    Else
        mgvlFilterSpeed = fsFAST
    End If

    mgvfWsFilTc = mgvfWsFilTc0 + (mgvfWsFilTc1 * mipFanspeed / mgvfNomFsMax)
    mgvfWsFilTc = Ctl_Limit(mgvfWsFilTc, mccfMinFilTc, mccfMaxFilTc)

' Filter the Delta Pressure
    If (mccbFilterInput) Then
        If mccfDisplayFilType = 2 Then
            If mobjWSFilterBW Is Nothing Then
                Set mobjWSFilterBW = New Butterworth
                fLastBWFilterTC = mgvfWsFilTc0
                mobjWSFilterBW.InitializeLowPass 1 / mccfUpdateInt, (1# / (2# * 3.14159 * mgvfWsFilTc0)), 4
                Set mobjWSFilterBW2 = New Butterworth
                fLastBWFilterTC = mgvfWsFilTc0
                mobjWSFilterBW2.InitializeLowPass 1 / mccfUpdateInt, (1# / (2# * 3.14159 * mgvfWsFilTc0)), 4
            End If
            If (mgvlFilterSpeed = fsFAST) Then
                mobjWSFilterBW.Reset
                fLastBWFilterTC = -1
            End If
            If mgvfWsFilTc0 <> fLastBWFilterTC Then
                fLastBWFilterTC = mgvfWsFilTc0
                mobjWSFilterBW.InitializeLowPass 1 / mccfUpdateInt, (1# / (2# * 3.14159 * mgvfWsFilTc0)), 4
                mobjWSFilterBW2.InitializeLowPass 1 / mccfUpdateInt, (1# / (2# * 3.14159 * mgvfWsFilTc0)), 4
            End If
            mopFilDeltaP = mobjWSFilterBW.GetFilteredValue(CDbl(mipDeltaP))
            mopFilDeltaPPlenum = mobjWSFilterBW2.GetFilteredValue(CDbl(mipDeltaPPlenum))
        ElseIf (mobjWSFilter.FilterType = ctlRollingAverage) Then
            If (mgvlFilterSpeed = fsFAST) Then
                mobjWSFilter.ClearData
                mobjWSFilter2.ClearData
            End If
            mobjWSFilter.FilterConst(0) = mgvfWsFilTc / mccfUpdateInt
            mobjWSFilter2.FilterConst(0) = mgvfWsFilTc / mccfUpdateInt
            mopFilDeltaP = mobjWSFilter.FilteredValue(mipDeltaP)
            mopFilDeltaPPlenum = mobjWSFilter2.FilteredValue(mipDeltaPPlenum)
        Else
            If (mgvlFilterSpeed = fsFAST) Then
                mobjWSFilter.FilterConst(0) = mgvfWsFilterFast
                mobjWSFilter2.FilterConst(0) = mgvfWsFilterFast
            Else
                mobjWSFilter.FilterConst(0) = mgvfWsFilTc
                mobjWSFilter2.FilterConst(0) = mgvfWsFilTc
            End If
            mopFilDeltaP = mobjWSFilter.FilteredValue(mipDeltaP)
            mopFilDeltaPPlenum = mobjWSFilter2.FilteredValue(mipDeltaPPlenum)
        End If
    Else
      ' No filtering
        mopFilDeltaP = mipDeltaP
        mopFilDeltaPPlenum = mipDeltaPPlenum
    End If

    mgvfTPlAbs = mipAirtempPlenum + 273.15
    dPsatExp = -6094.4642 * 1# / mgvfTPlAbs + 21.1249952 - 0.027245552 * mgvfTPlAbs + 0.000016853396 * mgvfTPlAbs ^ 2 + 2.4575506 * Log(mgvfTPlAbs)
    dPSat = Exp(dPsatExp)
    dPv = dPSat * mipRelHumid / 100#
    
    ' Calculate WS dependent values of Kp, Kqn, and Kqpl
    If mopWsMode <= modeManual Then
        mgvdVmps = LimitD(PolyDouble(CDbl(mipFanspeed), mgvdWSvsFSCoeffs, 6), 0#, mgvfMaxWs) / 3.6
    Else
        mgvdVmps = mopWsSetpt / 3.6
    End If
    
    mgvdKp = LimitD(PolyDouble(mgvdVmps * 3.6, mgvdKpCoeffs, 6), mgvfKpCoeffMin, mgvfKpCoeffMax)
    mgvdKqn = LimitD(PolyDouble(mgvdVmps * 3.6, mgvdKqnCoeffs, 6), mgvfKqnCoeffMin, mgvfKqnCoeffMax)
    mgvdKqpl = LimitD(PolyDouble(mgvdVmps * 3.6, mgvdKqplCoeffs, 6), mgvfKqplCoeffMin, mgvfKqplCoeffMax)
    
    mgvdPoSCPa = mipPabs * 1000# + mopFilDeltaP
    mgvdPoTSPa = mipPabs * 1000# + mgvdKp * mopFilDeltaP
    mgvdTSCAbs = mipAirtemp + 273.15
    mgvdTTSAbs = (mgvdPoTSPa / mgvdPoSCPa) ^ ((1.4 - 1#) / 1.4) * mgvdTSCAbs
    
    mgvdXPPl = 0.6222 * (dPv / (mipPabs * 1000# - dPv))
        
    ' Calc thermo properties based upon computed static pressure, airtemp, and dewpoint
    mopAirDensityKgm3 = ((mgvdPoTSPa / (Rair * mgvdTSCAbs)) * (0.6222 / (0.6222 + mgvdXPPl))) + ((mgvdPoTSPa / (Rwater * mgvdTSCAbs)) * (mgvdXPPl / (0.6222 + mgvdXPPl)))
    
    mopQnPa = mgvdKqn * mopFilDeltaP
    If mopQnPa < 0# Then mopQnPa = 0#
    mopQplPa = mgvdKqpl * mopFilDeltaPPlenum
    If mopQplPa < 0# Then mopQplPa = 0#
    
    If mopWSCalcMethod = 1 Then      ' Plenum
        mopQPa = mopQplPa
    Else
        mopQPa = mopQnPa
    End If
    
    If (mopAirDensityKgm3 > 0#) Then      ' Compute mopWindspeed kph
        If (mopAirDensityKgm3 > 0#) And (mopQPa >= 0#) Then       ' Compute mopWindspeed kph
            mopWindspeed = 3.6 * Sqr((2# * mopQPa) / mopAirDensityKgm3)
        Else
            mopWindspeed = 0#
        End If
    Else
        mopWindspeed = 0#
    End If

    If (mgvbSeparateControl) Then
'Filter separate control mopWindspeed if required.
        mgvfControlFilTc = mgvfControlFilTc0 + mgvfControlFilTc1 * mgvfNormWsSetpt
        mgvfControlFilTc = Ctl_Limit(mgvfControlFilTc, mgvfControlFilTcMin, mgvfControlFilTcMax)
        If mgvfControlFilTc > mccfUpdateInt Then
            With mobjWSctlFilter
                .FilterConst(0) = mgvfControlFilTc
                mgvfFilDeltaPControl = .FilteredValue(IIf(mopWSCalcMethod = 1, mipDeltaPPlenum, mipDeltaP))
            End With
        Else
            mgvfFilDeltaPControl = IIf(mopWSCalcMethod = 1, mipDeltaPPlenum, mipDeltaP)
        End If
        If mopWSCalcMethod = 1 Then      ' Plenum
            mgvfQControlPa = mgvdKqpl * mgvfFilDeltaPControl
        Else
            mgvfQControlPa = mgvdKqn * mgvfFilDeltaPControl
        End If

        If (mopAirDensityKgm3 > 0#) And (mgvfQControlPa >= 0#) Then       ' Compute Windspeed kph
            mopWsControl = 3.6 * Sqr((2# * mgvfQControlPa) / mopAirDensityKgm3)
        Else
            mopWsControl = 0#
        End If
    Else
        mopWsControl = mopWindspeed
    End If
    
    If (mopIdleStatus = isIdle) Then
        mopWindspeed = 0#
    ElseIf (mipFanspeed <= mccfMinDeltaPFSMax) Then  ' FS can be negative
        If (mopFilDeltaP <= mccfMinDeltaP) Then         ' if DeltaP is < min
            mopWindspeed = 0#                           '   mopWindspeed is assumed zero
            mopWsControl = 0#
        ElseIf FanspeedIsSetToZero Then
            mopWindspeed = 0#                           '   mopWindspeed is assumed zero
            mopWsControl = 0#
        End If
    End If                  ' endif

' Handle reverse airflow cases
    If (((mopWsMode = modeManual) And (mgvfFsTarget < 0#)) Or ((mopWsMode > modeManual) And (mopWsSetpt < 0))) And (mipFanspeed < 0#) Then
        mopWindspeed = mccfNomWSReverse * mipFanspeed / mccfNomFSReverse
        mopWsControl = mopWindspeed
' Handle low measured Delta P cases
    ElseIf (mopWsMode >= modeAuto) And (mopWsSetpt <= mccfMinWsForPid) And (mopWsSetpt > 0) And (mopIdleStatus = isNormal) Then
        mopWindspeed = mgvfNomWsMax * mopFsSetpt / (mccfFfGain * mgvfNomFsMax)
        mopWsControl = mopWindspeed
    End If
    
    mopKp = mgvdKp
    mopKqpl = mgvdKqpl
    mopKqn = mgvdKqn
    
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
Private Function FanspeedIsSetToZero() As Boolean

' Returns True if the fan is at zero:
'       - fan off
'       - fan on, manual mode, fs target is zero
'       - fan on, auto mode, ws target is zero, and has been this way for over 60 seconds

    On Error GoTo Error
    
    If mopWsMode = modeOff Then
        FanspeedIsSetToZero = True
    ElseIf mopWsMode = modeManual Then
        If mccfTimeAtZeroToForceWSZero < 0 Then
        ElseIf SecondsSinceTime(mgvdLastTimeFSTargetNotZero) >= mccfTimeAtZeroToForceWSZero Then
            FanspeedIsSetToZero = True
        End If
    Else
        If mccfTimeAtZeroToForceWSZero < 0 Then
        ElseIf SecondsSinceTime(mgvdLastTimeWSTargetNotZero) >= mccfTimeAtZeroToForceWSZero Then
            FanspeedIsSetToZero = True
        End If
    End If
    
    GoTo Done
    
Error:
    Resume Done

Done:
    On Error Resume Next
    Exit Function
    
End Function
Private Sub InputConversions()
    
'Perform input conversions/calculations
    
    If mipPabs <= 10 Then
        mipPabs = 101#
    End If
    
    If (mipCSStatus = 1#) Then
        mgvfAuxStatus = 1#
    ElseIf (mipCSStatus = 0#) Then
        mgvfAuxStatus = 0#
    Else
        mgvfAuxStatus = 2#
    End If
    
    If WindTunnel = wtAAWT Then
        mgvlNozzle = nsLarge
        mopIdleStatus = isNormal
    Else
        If mipNozzleSmallLimSw = 1 Then
            mgvlNozzle = nsSmall
        ElseIf mipNozzleMediumLimSw = 1 Then
            mgvlNozzle = nsMedium
        ElseIf mipNozzleLargeLimSw = 1 Then
            mgvlNozzle = nsLarge
        Else
            mgvlNozzle = nsMoving
        End If
        
        If (mipIdleNozFlapOpenStatus = 1#) And (mipIdleBypFlapClosedStatus(1) = 1#) And (mipIdleBypFlapClosedStatus(2) = 1#) And (mipIdleBypFlapClosedStatus(3) = 1#) Then
            mopIdleStatus = isNormal
        ElseIf (mipIdleNozFlapClosedStatus = 1#) And (mipIdleBypFlapOpenStatus(1) = 1#) And (mipIdleBypFlapOpenStatus(2) = 1#) And (mipIdleBypFlapOpenStatus(3) = 1#) Then
            mopIdleStatus = isIdle
        ElseIf mipAuxDPLCCmdRequest = 0 Then
            mopIdleStatus = isMoving
        Else 'If (mopIdleStartControl = 1#) Or (mopIdleStopControl = 1#) Then
            mopIdleStatus = isMoving
        'Else
        '    mopIdleStatus = isFault
        End If
        
    End If
    
End Sub
Private Sub CalcNozzleModeDependentConstants()

    Dim bChanged As Boolean
    Dim bExists As Boolean
    Dim I As Long
    Dim J As Long
    
    If mgvlNozzle <> mgvlOldNozzle Then
        bChanged = True
    End If
    
    Select Case mgvlNozzle
    Case nsLarge, nsSmall, nsMedium
        I = mgvlNozzle
    Case Else
        I = nsLarge
    End Select
        
    'PID feedback Gains
    If (mgvlWindTunnel = wtCWT) And mccbAutoFollowDyTarget And (mccbAutoFollowDyTargetUseTrackCoeffs = False) And (mopWsMode = modeTrack) And (mgvsTrackParName = mccsDefTrackParName) And ((mipDyMode = DY_MODE_SPEED) Or (mipThrMode = THR_MODE_SPEED)) Then
        mgvfKp0 = mccfKp0(I)
        mgvfKp1 = mccfKp1(I)
        mgvfTi0 = mccfTi0(I)
        mgvfTi1 = mccfTi1(I)
        mgvfTd0 = mccfTd0(I)
        mgvfTd1 = mccfTd1(I)
        mgvfWsFilterFast = mccfWsFilterFast(I)
    ElseIf (mopWsMode = modeTrack) Then
        mgvfKp0 = mccfKp0Track(I)
        mgvfKp1 = mccfKp1Track(I)
        mgvfTi0 = mccfTi0Track(I)
        mgvfTi1 = mccfTi1Track(I)
        mgvfTd0 = mccfTd0Track(I)
        mgvfTd1 = mccfTd1Track(I)
        mgvfWsFilterFast = mccfWsFilterFastTrack(I)
    Else
        mgvfKp0 = mccfKp0(I)
        mgvfKp1 = mccfKp1(I)
        mgvfTi0 = mccfTi0(I)
        mgvfTi1 = mccfTi1(I)
        mgvfTd0 = mccfTd0(I)
        mgvfTd1 = mccfTd1(I)
        mgvfWsFilterFast = mccfWsFilterFast(I)
    End If
    mgvfMinKp = mccfKpMin(I)
    mgvfMaxKp = mccfKpMax(I)
    mgvfMinTi = mccfTiMin(I)
    mgvfMaxTi = mccfTiMax(I)
    mgvfMinTd = mccfTdMin(I)
    mgvfMaxTd = mccfTdMax(I)
    
    'Filter Coefs
    mgvfWsFilTc0 = mccfWsFilterTc0(I)
    mgvfWsFilTc1 = mccfWsFilterTc1(I)
    mgvfCloseWsFilTc = mccfCloseWsFilTc(I)
    
    mgvfControlFilTc0 = mccfControlFilTc0(I)
    mgvfControlFilTc1 = mccfControlFilTc1(I)
    mgvfControlFilTcMin = mccfControlFilTcMin(I)
    mgvfControlFilTcMax = mccfControlFilTcMax(I)
    
    'Ranges
    mgvfNomWsMax = mccfNomWs(I)
    mgvfMaxWs = mccfMaxWs(I)
    mgvfNomFsMax = mccfNomFs(I)
    mgvfMaxFs = mccfMaxFs(I)
    
    If bChanged Then
        With mobjSPGConstants
            .CreatePropertyByPosition(CC_FSMax, bExists).Value = mgvfMaxFs
            .CreatePropertyByPosition(CC_WSMax, bExists).Value = mgvfMaxWs
        End With
    End If
    
    ' Windspeed calculation coefficients
    
    For J = 0 To 6
        mgvdKpCoeffs(J) = mccdKpCoeffs(I, J)
        mgvdKqnCoeffs(J) = mccdKqnCoeffs(I, J)
        mgvdKqplCoeffs(J) = mccdKqplCoeffs(I, J)
        mgvdWSvsFSCoeffs(J) = mccdWSvsFSCoeffs(I, J)
    Next J
    mgvfKpCoeffMin = mccfKpCoeffMin(I)
    mgvfKpCoeffMax = mccfKpCoeffMax(I)
    mgvfKqnCoeffMin = mccfKqnCoeffMin(I)
    mgvfKqnCoeffMax = mccfKqnCoeffMax(I)
    mgvfKqplCoeffMin = mccfKqplCoeffMin(I)
    mgvfKqplCoeffMax = mccfKqplCoeffMax(I)

End Sub

Public Function GetParameter(ParName As String) As ParametersLib.Parameter

' Returns a specific parameter

' Part          Description
' ----------------------------------------------------------------------
' ParName       The parameter name

    Dim bRaiseError As Boolean
    Dim lRaiseErrorNumber As Long
    Dim sRaiseErrorText As String
    Dim sRaiseErrorSource As String
    
    On Error GoTo Error
    Set GetParameter = mobjParManager.Configurations.Parameters(ParName)
    
    GoTo Done
Error:
    bRaiseError = True
    sRaiseErrorText = Err.Description
    lRaiseErrorNumber = tceInvalidArgs
    sRaiseErrorSource = Err.Source
    Resume Done
Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Err.Raise lRaiseErrorNumber, sRaiseErrorSource, sRaiseErrorText
    End If
    Exit Function
End Function

Private Sub Start()

'Private start routine

    Const PROC_NAME = "Start"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim I As Long
    Dim objProcess As Object
    Dim sClassName As String
    Dim objProcessManagerProcess As Object
    Dim objSPGM As SharedPropertyGroupManager
    Dim bExists As Boolean
    Dim lTimerFreq As Long
    
    On Error GoTo Error
    bRaiseError = False
        
    On Error GoTo Error_CreatingObjects
    sClassName = "CommandListLib.Commands"
    Set mobjCommands = New CommandListLib.Commands

    sClassName = "ControlLib.ControlConstants"
    Set mobjConstants = New ControlLib.ControlConstants
    
    sClassName = "ControlLib.Monitor"
    Set mobjWSMonitor = New ControlLib.Monitor
    Set mobjFSMonitor = New ControlLib.Monitor
    
    sClassName = "ControlLib.Filter"
    Set mobjWSFilter = New ControlLib.Filter
    Set mobjWSFilter2 = New ControlLib.Filter
    Set mobjWSctlFilter = New ControlLib.Filter
    Set mobjTrackFilter = New ControlLib.Filter
    Set mobjFFLeadLag = New ControlLib.Filter
    Set mobjRampRateFilter = New ControlLib.Filter
    Set mobjOutputFilter = New ControlLib.Filter
    
    sClassName = "PIDControl"
    Set mobjPID = New ControlLib.PIDControl
    
    sClassName = "ControlLib.GlobalVars"
    Set mobjGlobalVars = New ControlLib.GlobalVars
    
    sClassName = "Events"
    Set mobjEvents = New Events
    
    sClassName = "Interlocks"
    Set mobjInterlocks = New Interlocks
        
    sClassName = "Invalid Opcodes Collection"
    Set mcolInvalidOpcodes = New Collection
            
    sClassName = "SharedPropertyGroupManager"
    Set objSPGM = New SharedPropertyGroupManager
    
    sClassName = "SharedPropertyGroup - Constants"
    Set mobjSPGConstants = objSPGM.CreatePropertyGroup(SPG_CONSTANTS, 0, 0, bExists)
        
    sClassName = "PermissiveComponents"
    Set mobjMFPermComponents = New PermissiveComponents
    Set mobjMFAuxPermComponents = New PermissiveComponents
    Set mobjIdleCmdPermComponents = New PermissiveComponents
        
    On Error GoTo Error
    
'Set up loop timing statistics
    Set mobjStats = New AdvUtilLib.Stats
    Set tmrElapsed = New ccrpStopWatch
    
    mgvlWindTunnel = WindTunnel
    
    DoEvents
    
'Set up interlock objects
' ToDo: add mode switch logic here once inputs for this switch available
    With mobjInterlocks
        With .Add(CStr(ocInitialize))
            .Expressions.AddBoolean "Initialized", mbInited, True, ilckVariableEqualValueMeansOK, LoadResString(stsControlSystemNotInitialized)
        End With
        
        With .Add(CStr(ocStopMainFanAux))
            .Expressions.Add "FanRemote", mipMfRemote, 1#, ilckVariableEqualValueMeansOK, LoadResString(ridFanNotinComputerCtl)  '"Main Fan Drive not in remote"
        End With
        With .Add(CStr(ocStartMainFanAux))
            .IncludedInterlocks.Add CStr(ocStopMainFanAux)
            .Expressions.Add "MfAuxFlt", mipAuxFault, 1#, ilckVariableEqualValueMeansOK, LoadResString(ridAuxFault)
            .Expressions.Add "MFAuxReady", mipAuxReady, 1#, ilckVariableEqualValueMeansOK, LoadResString(ridAuxNotReady)
        End With
        
        With .Add(CStr(ocSetModeManual))
            .Expressions.Add "FanRemote", mipMfRemote, 1#, ilckVariableEqualValueMeansOK, LoadResString(ridFanNotinComputerCtl)  '"Main Fan Drive not in remote"
        End With
        
        With .Add(CStr(ocStopMainFan))
            .IncludedInterlocks.Add CStr(ocSetModeManual)
        End With
        With .Add(CStr(ocStartMainFan))
            .IncludedInterlocks.Add CStr(ocSetModeManual)
            .IncludedInterlocks.Add CStr(ocStartMainFanAux)
            .Expressions.Add "MFRemote", mipMfRemote, 1#, ilckVariableEqualValueMeansOK, LoadResString(ridFanNotinComputerCtl)
            .Expressions.Add "MFFault", mipMfLCFault, 1#, ilckVariableEqualValueMeansOK, LoadResString(ridFanLCFault)
            .Expressions.Add "MFPerm", mipMfPerm, 1#, ilckVariableEqualValueMeansOK, LoadResString(ridFanDriveSafetyFault)
            ''.Expressions.Add "EStop", mipEstopStatus, 1#, ilckVariableEqualValueMeansOK, LoadResString(ridCommonInterlockFaulted) 'ToDo: Add back in
            .Expressions.Add "MFDrvFlt", mipMfMCFault, 1#, ilckVariableEqualValueMeansOK, LoadResString(ridFanMCFault)
            .Expressions.Add "Purge", mopPurgeCmd, 0, ilckVariableEqualValueMeansOK, LoadResString(ridPurging)
        End With
        With .Add(CStr(ocSetFanSpeed))
            .IncludedInterlocks.Add CStr(ocStartMainFan)
        End With
        With .Add(CStr(ocSetModeAuto))
            .IncludedInterlocks.Add CStr(ocStartMainFan)
        End With
        With .Add(CStr(ocSetWindSpeed))
            .IncludedInterlocks.Add CStr(ocSetModeAuto)
        End With
        With .Add(CStr(ocTrackWindSpeed))
            .IncludedInterlocks.Add CStr(ocSetModeAuto)
        End With
        With .Add(CStr(ocStartPurge))
            .Expressions.Add "MFOn", mipMfStatus, 0, ilckVariableEqualValueMeansOK, LoadResString(ridFanIsOnStopFanFirst)
        End With
        With .Add(CStr(ocShowStatus))
            .IncludedInterlocks.Add CStr(ocSetWindSpeed)
        End With
        
        If WindTunnel = wtCWT Then
            With .Add(CStr(ocStartIdle))
                .Expressions.Add "Fault", mipIdleNozFlapMotorFault, 1#, ilckVariableEqualValueMeansOK, LoadResString(ridIdleDriveFault)
                .Expressions.Add "IdleStop", mipIdleNozFlapMotorStop, 0#, ilckVariableEqualValueMeansOK, LoadResString(ridIdleMotorStopOn)
                .Expressions.Add "Estop", mipEstopStatus, 1#, ilckVariableEqualValueMeansOK, LoadResString(ridEstop)
            End With
            With .Add(CStr(ocStopIdle))
                .IncludedInterlocks.Add CStr(ocStartIdle)
            End With
            With .Add(CStr(ocShowIdleStatus))
                .IncludedInterlocks.Add CStr(ocStartIdle)
            End With
            With .Add(CStr(ocSetNozzleSize))
                .Expressions.Add "MFOn", mipMfStatus, 0, ilckVariableEqualValueMeansOK, LoadResString(ridFanIsOnStopFanFirst)
                .Expressions.Add "Estop", mipEstopStatus, 1#, ilckVariableEqualValueMeansOK, LoadResString(ridEstop)
            End With
            With .Add(CStr(ocSetCFlapLeft1))
                .Expressions.Add "MFOn", mipMfStatus, 0, ilckVariableEqualValueMeansOK, LoadResString(ridFanIsOnStopFanFirst)
                .Expressions.Add "Estop", mipEstopStatus, 1#, ilckVariableEqualValueMeansOK, LoadResString(ridEstop)
                .Expressions.Add "Fault1", mipCFLeft1Fault, 1#, ilckVariableEqualValueMeansOK, LoadResString(ridCFLeftFault, 1)
                .Expressions.Add "LockOut1", mipCFLeft1Lockout, 0#, ilckVariableEqualValueMeansOK, LoadResString(ridCFLeftLockout, 1)
            End With
            With .Add(CStr(ocSetCFlapLeft2))
                .Expressions.Add "MFOn", mipMfStatus, 0, ilckVariableEqualValueMeansOK, LoadResString(ridFanIsOnStopFanFirst)
                .Expressions.Add "Estop", mipEstopStatus, 1#, ilckVariableEqualValueMeansOK, LoadResString(ridEstop)
                .Expressions.Add "Fault2", mipCFLeft2Fault, 1#, ilckVariableEqualValueMeansOK, LoadResString(ridCFLeftFault, 2)
                .Expressions.Add "LockOut2", mipCFLeft2Lockout, 0#, ilckVariableEqualValueMeansOK, LoadResString(ridCFLeftLockout, 2)
            End With
            With .Add(CStr(ocSetCFlapLeft))
                .IncludedInterlocks.Add CStr(ocSetCFlapLeft1)
                .IncludedInterlocks.Add CStr(ocSetCFlapLeft2)
            End With
            With .Add(CStr(ocSetCFlapRight1))
                .Expressions.Add "MFOn", mipMfStatus, 0, ilckVariableEqualValueMeansOK, LoadResString(ridFanIsOnStopFanFirst)
                .Expressions.Add "Estop", mipEstopStatus, 1#, ilckVariableEqualValueMeansOK, LoadResString(ridEstop)
                .Expressions.Add "Fault1", mipCFRight1Fault, 1#, ilckVariableEqualValueMeansOK, LoadResString(ridCFRightFault, 1)
                .Expressions.Add "LockOut1", mipCFRight1Lockout, 0#, ilckVariableEqualValueMeansOK, LoadResString(ridCFRightLockout, 1)
            End With
            With .Add(CStr(ocSetCFlapRight2))
                .Expressions.Add "MFOn", mipMfStatus, 0, ilckVariableEqualValueMeansOK, LoadResString(ridFanIsOnStopFanFirst)
                .Expressions.Add "Estop", mipEstopStatus, 1#, ilckVariableEqualValueMeansOK, LoadResString(ridEstop)
                .Expressions.Add "Fault2", mipCFRight2Fault, 1#, ilckVariableEqualValueMeansOK, LoadResString(ridCFRightFault, 2)
                .Expressions.Add "LockOut2", mipCFRight2Lockout, 0#, ilckVariableEqualValueMeansOK, LoadResString(ridCFRightLockout, 2)
            End With
            With .Add(CStr(ocSetCFlapRight))
                .IncludedInterlocks.Add CStr(ocSetCFlapRight1)
                .IncludedInterlocks.Add CStr(ocSetCFlapRight2)
            End With
            With .Add(CStr(ocSetCFlapTop))
                .Expressions.Add "MFOn", mipMfStatus, 0, ilckVariableEqualValueMeansOK, LoadResString(ridFanIsOnStopFanFirst)
                .Expressions.Add "Estop", mipEstopStatus, 1#, ilckVariableEqualValueMeansOK, LoadResString(ridEstop)
                .Expressions.Add "Fault", mipCFTopFault, 1#, ilckVariableEqualValueMeansOK, LoadResString(ridCFTopFault)
                .Expressions.Add "LockOut", mipCFTopLockout, 0#, ilckVariableEqualValueMeansOK, LoadResString(ridCFTopLockout)
            End With
            With .Add(CStr(ocStopCF))
            End With
        End If
            
    End With
    
'Set up invalid opcode collections. Add arrays of invalid opcodes to the collection and use
' the opcode as a string as an index
    With mcolInvalidOpcodes
        .Add Array(ocSetModeAuto, ocSetFanSpeed, ocSetWindSpeed, ocTrackWindSpeed, ocStopMainFan, ocStopMainFanAux), CStr(ocSetFanSpeed)
        .Add Array(ocSetFanSpeed, ocSetWindSpeed, ocTrackWindSpeed, ocStopMainFan, ocStopMainFanAux, ocStartIdle), CStr(ocSetWindSpeed)
        .Add Array(ocSetFanSpeed, ocSetWindSpeed, ocTrackWindSpeed, ocStopMainFan, ocStopMainFanAux, ocStartIdle), CStr(ocTrackWindSpeed)
        .Add Array(ocShowStatus), CStr(ocShowStatus)
        .Add Array(ocStopMainFan, ocStopMainFanAux), CStr(ocStartMainFan)
        .Add Array(ocStopMainFanAux), CStr(ocStartMainFanAux)
        .Add Array(ocSetFanSpeed, ocSetWindSpeed, ocTrackWindSpeed, ocStartMainFan), CStr(ocStopMainFan)
        .Add Array(ocSetFanSpeed, ocSetWindSpeed, ocTrackWindSpeed, ocStartMainFan, ocStartMainFanAux), CStr(ocStopMainFanAux)
        If WindTunnel = wtCWT Then
            .Add Array(ocSetWindSpeed, ocTrackWindSpeed, ocStopIdle, ocStartIdle), CStr(ocStartIdle)
            .Add Array(ocStartIdle, ocStopIdle), CStr(ocStopIdle)
            .Add Array(ocStartIdle, ocStopIdle), CStr(ocCancelIdle)
            .Add Array(ocSetNozzleSize, ocStopNozzleMotion), CStr(ocSetNozzleSize)
            .Add Array(ocSetNozzleSize), CStr(ocStopNozzleMotion)
            .Add Array(ocSetCFlapLeft, ocSetCFlapLeft1, ocSetCFlapLeft2, ocSetCFlapRight, ocSetCFlapRight1, ocSetCFlapRight2, ocSetCFlapTop), CStr(ocStopCF)
        End If
    End With

'Load control constants
    On Error Resume Next
    Initialize initConstants
    On Error GoTo Error
    
    On Error Resume Next
    With mobjGlobalVars
        .AddLong "Nozzle Size (1=Normal/Large, 2=Medium, 3=Small) (mgvfNozzle)", mgvlNozzle
        .AddSingle "Aux Summary Status (0=Off, 1=On, 2=Partial)", mgvfAuxStatus
        .AddSingle "WS Mode (1=Man,2=Auto,3=Track)", mopWsMode
        .AddLong "WS Internal Control Mode (1=Man,2=Auto,3=Track)", mgvlWSControlMode
        .AddLong "Windspeed Filter Speed (0=Slow, 1=Fast)", mgvlFilterSpeed
        .AddSingle "Filter Timer (when <=0, OK to use slow filter)", mgvfFilterTmr
        .AddSingle "Filter Type (Display) (0=First Order, 1=Rolling Avg, 2=BW)", mccfDisplayFilType
        .AddSingle "Ws Fil Tc 0", mgvfWsFilTc0
        .AddSingle "Ws Fil Tc 1", mgvfWsFilTc1
        .AddSingle "Ws Fil Tc (sec)", mgvfWsFilTc
        .AddSingle "Ws Fast Filter TC (sec)", mgvfWsFilterFast
        .AddSingle "DeltaP ", mipDeltaP    'mipDeltaP
        .AddSingle "Filtered DeltaP (Pa)", mopFilDeltaP
        .AddSingle "DeltaP Plenum", mipDeltaPPlenum    'mipDeltaP
        .AddSingle "Filtered DeltaP Plenum (Pa)", mopFilDeltaPPlenum
        .AddBoolean "Filter Input for WS calc", mccbFilterInput
        .AddBoolean "Separate Filters for Display and Control", mgvbSeparateControl
        .AddSingle "Control Filter TC", mgvfControlFilTc
        .AddSingle "Filtered DeltaP (Control) (Pa)", mgvfFilDeltaPControl
        .AddSingle "Kp (Static Pressure Coefficient)", mopKp
        .AddSingle "Kq plenum (Dynamic Pressure Coefficient)", mopKqpl
        .AddSingle "Kq nozzle (Dynamic Pressure Coefficient)", mopKqn
        .AddSingle "Q Plenum", mopQplPa
        .AddSingle "Q Nozzle", mopQnPa
        .AddSingle "WS Calc Method (1=Plenum, 0=Nozzle)", mopWSCalcMethod
        .AddSingle "Q", mopQPa
        .AddSingle "Q Control", mgvfQControlPa
        .AddSingle "Windspeed", mopWindspeed
        .AddSingle "Windspeed (Control)", mopWsControl
        
        .AddSingle "Windspeed Target (kph) (Control)", mgvfWsTarget
        .AddSingle "Limited Windspeed Target (kph) (Control)", mgvfWSTargetLimited
        .AddSingle "Windspeed Target (kph) (Displayed)", mopWsTargetDisp
        .AddSingle "Windspeed Rate in use (kph/s)", mgvfControlRamp
        .AddSingle "Windspeed Rate (kph/s)", mgvfWsRamp
        .AddSingle "Windspeed Setpoint (kph)", mopWsSetpt
        .AddSingle "Windspeed Increment", mgvfWsIncrement
        .AddSingle "Max Windspeed Rate For I (when under PID)", mgvfWsRateForI
        .AddSingle "Close Time (sec)", mgvfCloseTime
        .AddSingle "Close Time Ramp (kph/s)", mgvfCloseTimeRamp
        .AddSingle "Close Time Filter TC (sec)", mgvfCloseWsFilTc
        .AddSingle "Nominal Max Windspeed (kph)", mgvfNomWsMax
        .AddSingle "Normalized Windspeed Setpoint (kph)", mgvfNormWsSetpt
        
        .AddLong "Final Setpoint Ramp Method (0=None,1=CloseTime,2=FinalSnap)", mcclSetpointRampMethod
        .AddBoolean "Final Ramp - Can be used", mgvbFinalRampOKToUse
        .AddBoolean "Final Snap - In effect", mgvbUsingFinalRamp
        .AddBoolean "Final Snap - Coming back to Setpt", mgvbWSComingBackToSetpt
        .AddSingle "FF Norm Fanspeed (raw)", mgvfFfRawNormFs
        .AddSingle "FF Lead Coeff (sec)", mgvfFfLeadCoeff
        .AddSingle "FF Lag Coeff (sec)", mgvfFfLagCoeff
        .AddSingle "FF Norm Fanspeed (after LL)", mgvfFfNormFs
        
        .AddSingle "Windspeed Target Error", mgvfWSTargetError
        .AddSingle "Normalized Windspeed TargetError", mgvfTargetNormError
        .AddSingle "Windspeed Error", mgvfWsError
        .AddSingle "Normalized Windspeed Error", mgvfNormError
        .AddSingle "PID Kp", mgvfKp
        .AddSingle "PID Ti", mgvfTi
        .AddSingle "PID Td", mgvfTd
        .AddSingle "Kp Effect", mgvfKpEffect
        .AddSingle "Integrator Effect", mgvfIntegratorEffect
        .AddSingle "Old Integral Effect", mgvfOldIntegral
        .AddSingle "Deriviative Effect", mgvfDeriviativeEffect
        .AddSingle "PID Output (Norm FS)", mgvfFbNormFs
        .AddSingle "Controller Output (FB+FF, norm)", mgvfFsNormSetpt
        .AddSingle "Nominal Max Fanspeed (rpm)", mgvfNomFsMax
        .AddSingle "Fanspeed Target (rpm)", mgvfFsTarget
        .AddSingle "Limited Target (rpm)", mgvfFsLimitedTarget
        .AddSingle "Fanspeed Setpoint (rpm)", mopFsSetpt
        .AddSingle "Kp 0", mgvfKp0
        .AddSingle "Kp 1", mgvfKp1
        .AddSingle "Kp Max", mgvfMaxKp
        .AddSingle "Kp Min", mgvfMinKp
        .AddSingle "Ti 0", mgvfTi0
        .AddSingle "Ti 1", mgvfTi1
        .AddSingle "Ti Max", mgvfMaxTi
        .AddSingle "Ti Min", mgvfMinTi
        .AddSingle "Td 0", mgvfTd0
        .AddSingle "Td 1", mgvfTd1
        .AddSingle "Td Max", mgvfMaxTd
        .AddSingle "Td Min", mgvfMinTd
        .AddBoolean "WS Too low for PID", mgvbWsTooLowForPID
        .AddDate "WS too low for PID Start Time", mgvdWsTooLowForPIDTime
        
        If WindTunnel = wtCWT Then
            .AddSingle "Idle Status (0=Stopped, 1=Started, 2=Moving)", mopIdleStatus
            .AddBoolean "Auto Idle Active", mgvbAutoIdle
            .AddBoolean "Auto Idle Speed has been high", mgvbAutoIdleSpeedHasBeenHigh
            .AddBoolean "Auto Idle Speed has been Low", mgvbAutoIdleSpeedHasBeenLow
            .AddBoolean "Auto Idle Stop Idle Overrride Active", mgvbStopAutoIdle
        Else
            .AddBoolean "Control Ramp Held because speed was low", mgvbHoldCRUntilFSOK
        End If
        
        .AddDate "Last time of non-zero FS target", mgvdLastTimeFSTargetNotZero
        .AddDate "Last time of non-zero WS target", mgvdLastTimeWSTargetNotZero
        
        .AddSingle "mgvfPurgeStatus", mgvfPurgeStatus
        
        .AddSingle "mgvfHeadwind", mgvfHeadwind
        .AddSingle "mgvfControlRamp", mgvfControlRamp
        .AddSingle "mgvfCloseTime", mgvfCloseTime
        .AddSingle "mgvfFilHxDeltaP", mgvfFilHxDeltaP
        .AddSingle "mgvfFsIncrement", mgvfFsIncrement
        .AddSingle "mgvfHxDpOverDp", mgvfHxDpOverDp
        .AddSingle "mgvfMinTi", mgvfMinTi
        .AddSingle "mgvfMaxTi", mgvfMaxTi
        .AddSingle "mgvfMinBlockage", mgvfMinBlockage
        .AddSingle "mgvfMaxBlockage", mgvfMaxBlockage
        .AddSingle "mgvfMaxFs", mgvfMaxFs
        .AddSingle "mgvfMaxFsTarget", mgvfMaxFsTarget
        .AddSingle "mgvfMaxWs", mgvfMaxWs
        .AddSingle "mgvfMfStartDelay", mgvfMfStartDelay
        .AddSingle "mgvfNormWsSetpt", mgvfNormWsSetpt
        .AddSingle "mgvfOldRemoteStatus", mgvfOldRemoteStatus
        .AddSingle "mgvfOldCtlMode", mgvfOldCtlMode
        .AddSingle "mgvfOldMfStatus", mgvfOldMfStatus
        .AddSingle "mgvfOldAuxStatus", mgvfOldAuxStatus
        .AddLong "mgvlOldNozzle", mgvlOldNozzle
        .AddSingle "mgvfOldPurgeStopFanRequest", mgvfOldPurgeStopFanRequest
        .AddSingle "mgvfOldTsDoor4", mgvfOldTsDoor4
        .AddSingle "mgvfOldWsIncrement", mgvfOldWsIncrement
        .AddSingle "mgvfPressureRatio", mgvfPressureRatio
        .AddSingle "mgvfPstaticControlKpa", mgvfPstaticControlKpa
        .AddSingle "mgvfQControlPa", mgvfQControlPa
        .AddSingle "mgvfRawHxDpOverDp", mgvfRawHxDpOverDp
        
        .AddSingle "mgvfWsError", mgvfWsError
                            
    End With
    On Error GoTo Error
    
'Set up Statuses
    Set mobjStatuses = New Statuses
    With mobjStatuses
        .UseSPG SPG_STATUSES
        .Add STS_WINDSPEED_INITIALIZED, False, LoadResString(ridWindspeedNotInitialized), LoadResString(ridWindspeedInitialized)
        .Add STS_FANSPEED_ACHIEVED, False, LoadResString(ridFanspeedNotAchieved), LoadResString(ridFanspeedAchieved)
        .Add STS_WINDSPEED_ACHIEVED, False, LoadResString(ridWindspeedNotAchieved), LoadResString(ridWindspeedAchieved)
        .Add STS_MAINFAN_ON, False, "", LoadResString(ridMainFanOn)
        .Add STS_MAINFAN_OFF, False, "", LoadResString(ridMainfanOff)
        .Add STS_MAINFAN_AUX_ON, False, "", LoadResString(ridMainFanAuxOn)
        .Add STS_MAINFAN_AUX_OFF, False, "", LoadResString(ridMainFanAuxOff)
        If WindTunnel = wtCWT Then
            .Add STS_IDLE_STARTED, False, "", LoadResString(ridIdleStarted)
            .Add STS_IDLE_STOPPED, True, "", LoadResString(ridIdleStopped)
        End If
    End With
        
'Create a process object, and then pass this Process
' object to the Process Manager

    Set objProcess = New Process
    mbReady = True
    
    WaitForProcessesStartupComplete
    
'    Set objProcessManagerProcess = AddTalentProcessWait(CStr(objProcess.Name), objProcess)
'    mlID = objProcessManagerProcess.ID
    mlID = AddTalentProcessGetID(CStr(objProcess.Name), objProcess)
            
'Set up parameter objects (after adding process and getting Process ID)

    On Error GoTo Error_Configurations
    Pause 2#
    sClassName = "ParManager"
    Set mobjParManager = New ParManager
    mobjParManager.OwnerID = mlID
    With mobjParManager
        .AutoRefreshEnabled = False
        .IgnoreLoadedEvents = True
        .IgnoreUnloadedEvents = True
    
'Input Parameters
        .AddWithoutFlag "Airtemp", mipAirtemp, ioInput
        .AddWithoutFlag "Dewpoint", mipDewpoint, ioInput
        .AddWithoutFlag "RelHumidity", mipRelHumid, ioInput
        .AddWithoutFlag "AbsPress", mipPabs, ioInput
        .AddWithoutFlag "DeltaPNozzle", mipDeltaP, ioInput
        .AddWithoutFlag "DeltaPPlenum", mipDeltaPPlenum, ioInput
        
        .AddWithoutFlag "OpModeOff", mipOpModeOff, ioInput
        .AddWithoutFlag "OpModeMaint", mipOpModeMaint, ioInput
        .AddWithoutFlag "OpModePrep", mipOpModePrep, ioInput
        If mgvlWindTunnel = wtAAWT Then
            .AddWithoutFlag "OpModeStab", mipOpModeStab, ioInput
        End If
        .AddWithoutFlag "OpModeTest", mipOpModeTest, ioInput
        .AddWithoutFlag "WTCS_AuxCmdRequest", mipAuxDPLCCmdRequest, ioInput
        .AddWithoutFlag "WTCS_CSCmdRequest", mipCoolPLCCmdRequest, ioInput
        
        .AddWithoutFlag "PLC_FaultReset", mipToPLCResetCmd, ioInput
        .AddWithoutFlag "MF_Remote", mipMfRemote, ioInput
        .AddWithoutFlag "MF_Permissive", mipMfPerm, ioInput
        .AddWithoutFlag "PLCEstopOut", mipEstopStatus, ioInput
        .AddWithoutFlag "MF_MCFault", mipMfMCFault, ioInput
        .AddWithoutFlag "MF_MCReady", mipMCReady, ioInput
        .AddWithoutFlag "MF_LCFault", mipMfLCFault, ioInput
        .AddWithoutFlag "MF_LCReady", mipLCReady, ioInput
        .AddWithoutFlag "MF_Status", mipMfStatus, ioInput
        .AddWithoutFlag "MF_Speed", mipFanspeed, ioInput
''        For I = 1 To 3
''            .AddWithoutFlag "MainFanSpeedLim" & I, mipMFSpeedLimitCommand(I), ioInput
''            .AddWithoutFlag "MainFanSpeedLim" & I & "Status", mipMFSpeedLimitStatus(I), ioInput
''        Next I
        .AddWithoutFlag "MF_CSReady", mipAuxReady, ioInput
        .AddWithoutFlag "MF_CSFault", mipAuxFault, ioInput
        .AddWithoutFlag "MF_CSStatus", mipCSStatus, ioInput
        
        If WindTunnel = wtCWT Then
            .AddWithoutFlag "NozTopLimSw", mipNozzleLargeLimSw, ioInput
            .AddWithoutFlag "NozMiddleLimSw", mipNozzleMediumLimSw, ioInput
            .AddWithoutFlag "NozBottomLimSw", mipNozzleSmallLimSw, ioInput
            
            .AddWithoutFlag "Idle_MotorStop", mipIdleNozFlapMotorStop, ioInput
            .AddWithoutFlag "Idle_FlapOpen", mipIdleNozFlapOpenStatus, ioInput
            .AddWithoutFlag "Idle_FlapClosed", mipIdleNozFlapClosedStatus, ioInput
            .AddWithoutFlag "Idle_MotorFault", mipIdleNozFlapMotorFault, ioInput
            For I = 1 To 3
                .AddWithoutFlag "BypFlap" & I & "Closed", mipIdleBypFlapClosedStatus(I), ioInput
                .AddWithoutFlag "BypFlap" & I & "Open", mipIdleBypFlapOpenStatus(I), ioInput
            Next I
            .AddWithoutFlag "DynoSpeedTgt", mipDySpeedTarget, ioInput
            .AddWithoutFlag "DynoSpeedRate", mipDySpeedRate, ioInput
            .AddWithoutFlag "DynoMode", mipDyMode, ioInput
            .AddWithoutFlag "ThrMode", mipThrMode, ioInput
''            .AddWithoutFlag "BLFanStatus", mipBLFanStatus, ioInput
        
            .AddWithoutFlag "CF_LeftPos1", mipCFLeft1Pos, ioInput
            .AddWithoutFlag "CF_Left1MotorFault", mipCFLeft1Fault, ioInput
            .AddWithoutFlag "CF_Left1MotorLockout", mipCFLeft1Lockout, ioInput
            .AddWithoutFlag "CF_LeftPos2", mipCFLeft2Pos, ioInput
            .AddWithoutFlag "CF_Left2MotorFault", mipCFLeft2Fault, ioInput
            .AddWithoutFlag "CF_Left2MotorLockout", mipCFLeft2Lockout, ioInput
            .AddWithoutFlag "CF_RightPos1", mipCFRight1Pos, ioInput
            .AddWithoutFlag "CF_Right1MotorFault", mipCFRight1Fault, ioInput
            .AddWithoutFlag "CF_Right1MotorLockout", mipCFRight1Lockout, ioInput
            .AddWithoutFlag "CF_RightPos2", mipCFRight2Pos, ioInput
            .AddWithoutFlag "CF_Right2MotorFault", mipCFRight2Fault, ioInput
            .AddWithoutFlag "CF_Right2MotorLockout", mipCFRight2Lockout, ioInput
            .AddWithoutFlag "CF_TopPos", mipCFTopPos, ioInput
            .AddWithoutFlag "CF_TopMotorFault", mipCFTopFault, ioInput
            .AddWithoutFlag "CF_TopMotorLockout", mipCFTopLockout, ioInput
                
        End If
        
'Output Parameters
''        .AddWithoutFlag "DeltaPFilt", mopFilDeltaP, ioOutput
''        .AddWithoutFlag "WindspeedRaw", mopWsControl, ioOutput
        .AddWithoutFlag "AirDensity", mopAirDensityKgm3, ioOutput
''        .AddWithoutFlag "Blockage", mopBlockageCorr, ioOutput
''        .AddWithoutFlag "BlockageType", mopBlockageType, ioOutput
        .AddWithoutFlag "PStatic", mopPstaticKpa, ioOutput
        .AddWithoutFlag "Q", mopQPa, ioOutput
        .AddWithoutFlag "WS_CalcMethod", mopWSCalcMethod, ioOutput
        .AddWithoutFlag "Kp", mopKp, ioOutput
        .AddWithoutFlag "Kqpl", mopKqpl, ioOutput
        .AddWithoutFlag "Kqn", mopKqn, ioOutput
        
        .AddWithoutFlag "MF_AuxStartCmd", mopAuxControl, ioOutput
        .AddWithoutFlag "MF_StartCmd", mopMfControl, ioOutput
        .AddWithoutFlag "MF_ResetCmd", mopMFReset, ioOutput
        .AddWithoutFlag "MF_SpeedSetpt", mopFsSetpt, ioOutput
        .AddWithoutFlag "MF_SpeedRate", mopFsRamp, ioOutput
        .AddWithoutFlag "MF_SpeedTgt", mopFsTargetDisp, ioOutput
        .AddWithoutFlag "Windspeed", mopWindspeed, ioOutput
        
        .AddWithoutFlag "WS_Setpt", mopWsSetptDisp, ioOutput
       '' .AddWithoutFlag "WindspeedSetptActual", mopWsSetpt, ioOutput
        .AddWithoutFlag "WS_Rate", mopWsRamp, ioOutput
        .AddWithoutFlag "WS_Target", mopWsTargetDisp, ioOutput
        .AddWithoutFlag "WS_Headwind", mopHeadwind, ioOutput
                
        .AddWithoutFlag "WS_Mode", mopWsMode, ioOutput
        
        If WindTunnel = wtCWT Then
            .AddWithoutFlag "NozzleSize", mopNozzle, ioOutput
            .AddWithoutFlag "AutoIdleStatus", mopAutoIdleStatus, ioOutput
            .AddWithoutFlag "Idle_Status", mopIdleStatus, ioOutput
            .AddWithoutFlag "Idle_StopCmd", mopIdleStopControl, ioOutput
            .AddWithoutFlag "Idle_StartCmd", mopIdleStartControl, ioOutput
            .AddWithoutFlag "Noz_Pos3Cmd", mopNozzleCmd(nsSmall), ioOutput
            .AddWithoutFlag "Noz_Pos2Cmd", mopNozzleCmd(nsMedium), ioOutput
            .AddWithoutFlag "Noz_Pos1Cmd", mopNozzleCmd(nsLarge), ioOutput
            .AddWithoutFlag "PurgeStartCmd", mopPurgeCmd, ioOutput
        
            .AddWithoutFlag "CF_LeftStartCmd", mopCFLeftStartCmd, ioOutput
            .AddWithoutFlag "CF_LeftStopCmd", mopCFLeftStopCmd, ioOutput
            .AddWithoutFlag "CF_LeftPos1SP", mopCFLeft1Target, ioOutput
            .AddWithoutFlag "CF_LeftPos2SP", mopCFLeft2Target, ioOutput
            .AddWithoutFlag "CF_RightStartCmd", mopCFRightStartCmd, ioOutput
            .AddWithoutFlag "CF_RightStopCmd", mopCFRightStopCmd, ioOutput
            .AddWithoutFlag "CF_RightPos1SP", mopCFRight1Target, ioOutput
            .AddWithoutFlag "CF_RightPos2SP", mopCFRight2Target, ioOutput
            .AddWithoutFlag "CF_TopStartCmd", mopCFTopStartCmd, ioOutput
            .AddWithoutFlag "CF_TopStopCmd", mopCFTopStopCmd, ioOutput
            .AddWithoutFlag "CF_TopPosSP", mopCFTopTarget, ioOutput
        End If
        
        If WindTunnel = wtAAWT Then
            .AddWithoutFlag "Windspeed_ms", mopWindspeedmps, ioOutput
            .AddWithoutFlag "WS_Setpt_ms", mopWsSetptmps, ioOutput
            .AddWithoutFlag "WS_Rate_ms", mopWsRampmps, ioOutput
            .AddWithoutFlag "WS_Target_ms", mopWsTargetmps, ioOutput
            .AddWithoutFlag "WS_Head_ms", mopHeadwindmps, ioOutput
            .AddWithoutFlag "PurgeCmd", mopPurgeCmd, ioOutput
            .AddWithoutFlag "WS_Achieved", mopWsAchieved, ioOutput
            For I = 1 To 18
                .AddWithoutFlag "OF_" & Format(I, "00") & "OpenCmd", mopOFOpenCmd(I), ioOutput
            Next I
        End If
    
    End With
        
    mobjMFPermComponents.Initialize mobjParManager.Configurations
    mobjMFAuxPermComponents.Initialize mobjParManager.Configurations
    mobjIdleCmdPermComponents.Initialize mobjParManager.Configurations
    
    On Error GoTo Error
' Now fully started up. Set process priority and enable timers
' If running within VB (the IDE) then use a form based timer, and don't
' use Real Time Priority
    If mccfUpdateInt < 0.01 Then
        mccfUpdateInt = 0.01
    End If

    If IsWithinIDE() Then
        Set mobjForm = New frmControl
        Load mobjForm
        mobjForm.tmrDesignTime.Interval = mccfUpdateInt * 1000
        mobjForm.tmrDesignTime.Enabled = True
    Else
' Calculate timer frequency. Start with update interval with 10% resolution
' Then overwrite this with the general setting in the Processes key if it exists.
' Then overwrite this with the process specific setting in the process specific key
' if it exists
        lTimerFreq = mccfUpdateInt * 1000 / 10
        lTimerFreq = Reg_GetValue(HKEY_LOCAL_MACHINE, REG_PROCESSES, , "TimerFrequency", lTimerFreq, vbLong)
        lTimerFreq = Reg_GetValue(HKEY_LOCAL_MACHINE, REG_PROCESSES, PROCESS_NAME, "TimerFrequency", lTimerFreq, vbLong)
        If lTimerFreq > mccfUpdateInt * 1000 Then
            lTimerFreq = mccfUpdateInt * 1000
        End If
        tmrMain.Stats.Frequency = lTimerFreq
        tmrMain.Interval = mccfUpdateInt * 1000
        tmrMain.EventType = TimerPeriodic
        Set tmrMain.Notify = Me
        SetProcessPriority mgvlDefPriority
        tmrMain.Enabled = True
    End If

    If mobjParManager.Configurations.Count > 0 Then
        mobjCommands.Append ocInitialize, "Initialize"
    End If

'Clear Run Once timer
    Set tmrRunOnce = Nothing
    mobjStats.Reset
    For I = 1 To 200
        mobjStats.AddValue CDbl(mccfUpdateInt)
    Next I
    tmrElapsed.Reset

    mgvlOldNozzle = -1
    mgvfOldCtlMode = -1
    
    GoTo Done
Error:
    sRaiseErrorText = Err.Description
    Set EH.VBErr = Err
    EH.AppError = StdTalentStrings.stsEHErrorDuringControllerStartup
    EH.Module = TypeName(Me)
    EH.Procedure = "Start"
    EH.Item(1) = GetComputerName
    Call EH_MsgBox
    Resume Next
Error_CreatingObjects:
    sRaiseErrorText = LoadResString(ridCantCreateObject, sClassName, Err.Description)
    bRaiseError = True
    Set EH.VBErr = Err
    EH.Module = TypeName(Me)
    EH.Procedure = "Start"
    EH.Item(1) = sClassName
    Call EH_MsgBox
    Resume Done
Error_Configurations:
    Set EH.VBErr = Err
    EH.AppError = LoadResString(StdTalentStrings.stsEHCantCreateConfigurations)
    EH.Module = TypeName(Me)
    EH.Procedure = PROC_NAME
    sRaiseErrorText = EH.Text
    Call EH_MsgBox
    Resume Next
Done:
    On Error Resume Next
    Set objSPGM = Nothing
    If bRaiseError Then
        On Error GoTo 0
        Call EH_RaiseError(OLEERR_METHOD_FAILED, TypeName(Me), "Start", sRaiseErrorText)
    End If
    Exit Sub
End Sub

Friend Sub ControlLogic()

'This routine is executed at every control interval

    Const PROC_NAME = "ControlLogic"
    Dim bInputsRead As Boolean
        
    On Error GoTo Error
    If mbInited = True Then
        mobjParManager.ReadInputs
        InputConversions
        CalcNozzleModeDependentConstants
        CalcWindspeed
        CheckInterlocks
        bInputsRead = True
    End If

    ExecuteCommands
    
    If mbInited Then
        If bInputsRead = False Then
            mobjParManager.ReadInputs
            InputConversions
            CalcNozzleModeDependentConstants
            CalcWindspeed
            CheckInterlocks
        End If
        WSControlLogic
        MonitorStatuses
        OutputConversions
        SaveOldStates
        
        mobjParManager.WriteOutputs
    End If
    
    If Not mobjStatuses Is Nothing Then
        mobjStatuses.FlushMessages
    End If
    
    GoTo Done
Error:
    If mbInited Then
        Set EH.VBErr = Err
        EH.AppError = stsEHErrorExecutingControlLogic
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume ErrorExit
ErrorExit:
    On Error Resume Next
    Suspend
    Resume Next
Done:
    On Error Resume Next
    Exit Sub
End Sub

Private Sub ExecuteCommands()

'Executes the next command on the list

    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    Static ReEntered As Boolean
    Dim objCommand As CommandListLib.Command
    Dim I As Long
    Dim fNozzleTopTarget As Single
    Dim fNozzleEastTarget As Single
    Dim fNozzleWestTarget As Single
    Dim sText As String
    
    On Error GoTo Error
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

'Initialize
        Case ocInitialize
            Me.Initialize , objCommand
            .Remove
            
        Case ocSuspend
'Shut down equipment if running, and then suspend
            If (.FirstTime) And (mbInited) And ((mopMfControl <> 0#) Or (mipMfStatus <> 0#)) Then
                mobjCommands.Insert ocStopMainFan, "StopMainFan"
            Else
                .Messages.Flush
                Suspend
                .Remove
                If IsWithinIDE() Then
                    Shutdown
                End If
            End If
        
        Case ocShutdown
'Shut down equipment if running, and then stop this application
            If .FirstTime And mbInited Then
                mobjCommands.Insert ocSuspend, "SuspendWindspeed"
            Else
                .Messages.Flush
                .Remove
                'Suspend
                Me.Shutdown
            End If
            
        Case Else       ' All other commands require initialized state
            
            CheckInterlock CStr(ocInitialize)
            
            Select Case .Opcode
                            
'Set windspeed controller operating mode to manual
            Case ocSetModeManual
                CheckInterlock CStr(ocSetModeManual)
                SetMode modeManual
                .Remove
                                                    
'Set windspeed controller operating mode to Auto (track done directly)
            Case ocSetModeAuto
                CheckInterlock CStr(ocSetModeAuto)
                If (mopIdleStatus <> isNormal) Or (mopIdleStartControl = 1) Then
                    .InsertPredecessor ocStopIdle, "StopIdle"
                Else
                    SetMode modeAuto
                    .Remove
                End If
                                                    
            Case ocResetFaults
                CheckInterlock CStr(ocResetFaults)
                If .FirstTime Then
                    .ResetTime
                    mopMFReset = 1#
                ElseIf .TimeSince >= mccfResetFaultTime Then
                    mopMFReset = 0#
                    .Remove
                End If
                
            Case ocStartMainFanAux
                CheckInterlock CStr(ocStartMainFanAux)
                mgvdTurnOffFanTime = Now()  'reset turnoff time
                .Messages.Flush
                mgvbAuxAutoOff = False
                mopAuxControl = 1       ' This will cause the Airtemp controller to start the HTa if it is not already running
                If .FirstTime Then
                    .ResetTime
                End If
                If mgvfAuxStatus = 1 Then
                    If mgvdAuxOnTime = 0 Then
                    ElseIf SecondsSinceTime(mgvdAuxOnTime) >= mccfMainFanAuxPause Then
                        .Remove
                    End If
                ElseIf .TimeSince >= mccfMainFanAuxTimeout Then
                    mopAuxControl = 0
                    sText = ""
                    If mipCSStatus <> 1# Then
                        sText = IIf(sText <> "", sText & vbCrLf & "- ", "- ") & LoadResString(ridCSDidNotStart)
                    End If
''                    If mipMCReady <> 1# Then
''                        sText = IIf(sText <> "", sText & vbCrLf & "- ", "- ") & LoadResString(ridMCDidNotGetReady)
''                    End If
''                    If mipLCReady <> 1# Then
''                        sText = IIf(sText <> "", sText & vbCrLf & "- ", "- ") & LoadResString(ridLCDidNotGetReady)
''                    End If
                    .Abort RaiseErrorText(tceDeviceTimedOut, .Text, sText), cmdTimedout
                End If

            Case ocStopMainFanAux
                CheckInterlock CStr(ocStopMainFanAux)
                If mopWsMode > modeManual Then
                    .InsertPredecessor ocSetModeManual, "SetModeManual"
                ElseIf (mopMfControl <> 0#) Or (mipMfStatus <> 0#) Then
                    .InsertPredecessor ocStopMainFan, "StopMainFan"
                Else
                    .Messages.Flush
                    mopAuxControl = 0
                    mgvfFsTarget = 0#
                    mopFsSetpt = 0#
                    mgvfWsTarget = 0#
                    mopWsSetpt = 0#
                    If (mgvfAuxStatus = 0#) Or (.TimeSince >= 30#) Then
                        .Remove
                    End If
                End If
                
            Case ocStartMainFan
                CheckInterlock CStr(ocStartMainFan)
                mgvdTurnOffFanTime = Now()  'reset turnoff time
                If mopWsMode < modeManual Then
                    .InsertPredecessor ocSetModeManual, "SetModeManual"
                ElseIf mgvfAuxStatus <> 1# Or mopAuxControl <> 1# Then
                    .InsertPredecessor ocStartMainFanAux, "StartMainFanAux"
                Else
                    .Messages.Flush
                    mopMfControl = 1
                    If mipMfStatus <> 1# Then
                        mgvfFsTarget = 0#
                    End If
                    If (mgvfFsTarget < mccfMinFs) And (mccfMinFs > 0) Then
                        mgvfFsTarget = mccfMinFs
                        mopFsSetpt = mccfMinFs
                    End If
                    If .FirstTime Then
                        .ResetTime
                    End If
                    If mipMfStatus = 1# Then
                        If mccfDrivePause <= 0 Then
                            .Remove
                        ElseIf mgvdDriveOnTime = 0 Then
                        
                        ElseIf SecondsSinceTime(mgvdDriveOnTime) >= mccfDrivePause Then
                            .Remove
                        End If
                    ElseIf .TimeSince >= mccfMainFanTimeout Then
                        mopMfControl = 0
                        .Abort , cmdTimedout
                    End If
                End If

            Case ocStopMainFan
                CheckInterlock CStr(ocStopMainFan)
                .Messages.Flush
                If mopWsMode <> modeManual Then
                    .InsertPredecessor ocSetModeManual, "SetModeManual"
                ElseIf (mgvfFsTarget <> 0#) Then
                    .InsertPredecessor ocSetFanSpeed, "SetFanSpeed", 0#
                    .ResetTime
                ElseIf (mopFsSetpt <> 0#) And (.TimeSince < 60#) Then
                    ' wait for fs setpoint to reach 0.0
                    .TagVariant = Now()
                ElseIf (Abs(mipFanspeed) > 5#) And (.TimeSince < 30#) Then
                    .TagVariant = Now()
                Else
                    mopMfControl = 0
                    If mccfMainFanAuxAutoOffTime > 0 Then
                        mgvbAuxAutoOff = True
                    End If
                    On Error Resume Next
                    If mipMfStatus <> 0# Then
                        If (.TimeSince >= 60#) Then
                            .Remove
                        Else
                            .TagVariant = Now()
                        End If
                    ElseIf SecondsSinceTime(.TagVariant) >= 5# Then
                        .Remove
                    End If
                End If
                
            Case ocSetWindSpeed
                CheckInterlock Array(CStr(ocSetWindSpeed), .Arguments(1).Value)
                mopWsTargetDisp = .Arguments(1).Value   ' Display updated target to user
                mopWsRamp = .Arguments(2).Value
                If mipMfStatus <> 1# Or mopMfControl <> 1# Then
                    .InsertPredecessor ocStartMainFan, "StartMainFan"
                ElseIf mopWsMode <> modeAuto Then
                    .InsertPredecessor ocSetModeAuto, "SetModeAuto"
                Else
                    .Messages.Flush
                    mgvfWsTarget = .Arguments(1).Value
                    mgvfWsRamp = .Arguments(2).Value
                    If mgvfWsRamp <= 0 Then
                        mgvfWsRamp = mccfDefWsRamp
                    End If
                    If mcclSetpointRampMethod = 1 Then
                        mobjRampRateFilter.ClearData
                        mgvfControlRamp = mobjRampRateFilter.FilteredValue(mgvfWsRamp)
                    End If
                    If Abs(mgvfWsTarget - mopWindspeed) <= (mccfFinalRampDeltakph + 1#) Then
                        mgvbFinalRampOKToUse = False
                    Else
                        mgvbFinalRampOKToUse = True
                    End If
                    mgvbFinalRampHasSnapped = False
                    mgvbUsingFinalRamp = False
                    mgvbHoldCRUntilFSOK = False
                    .Remove
                End If
                
            Case ocTrackWindSpeed
                CheckInterlock CStr(ocTrackWindSpeed)
                If mipMfStatus <> 1# Or mopMfControl <> 1# Then
                    .InsertPredecessor ocStartMainFan, "StartMainFan"
                ElseIf (.Arguments(3).Value = False) And (mopIdleStatus <> isNormal) Then
                    .InsertPredecessor ocStopIdle, "StopIdle"
                Else
                    .Messages.Flush
                    Set mobjTrackPar = mobjParManager.Configurations.Parameters(.Arguments(1))
                    GetConversionCoefficients mobjTrackPar.Units, "kph", mgvfTrackC1, mgvfTrackC0
                    SetMode modeTrack
                    mgvsTrackParName = .Arguments(1).Value
                    mgvfHeadwind = .Arguments(2).Value
                    If WindTunnel = wtCWT Then
                        mgvbAutoIdle = .Arguments(3).Value
                        mgvdLastIdleTime = Now
                        mgvbStopAutoIdle = False
                    End If
                    mgvfWsRamp = mccfMaxWsRamp
                    mgvfControlRamp = mgvfWsRamp
                    .Remove
                End If
                                             
            Case ocSetFanSpeed
                CheckInterlock CStr(ocSetFanSpeed)
                mopFsTargetDisp = .Arguments(1).Value  'Display updated target to user
                If .Arguments.Count >= 2 Then
                    mopFsRamp = .Arguments(2).Value
                End If
                If mopWsMode <> modeManual Then
                    .InsertPredecessor ocSetModeManual, "SetModeManual"
                ElseIf (mipMfStatus <> 1#) Or (mopMfControl <> 1#) Then
                    .InsertPredecessor ocStartMainFan, "StartMainFan"
                ElseIf (mopFsTargetDisp > mccfMaxFsLimit(1)) And LimitInEffect(1) Then
                    .InsertPredecessor ocStopIdle, "StopIdle"
                Else
                    .Messages.Flush
                    mgvfFsTarget = .Arguments(1).Value
                    If .Arguments.Count >= 2 Then
                        mgvfFsRamp = .Arguments(2).Value
                    Else
                        mgvfFsRamp = mccfDefFsRamp
                    End If
                    .Remove
                End If
            
            Case ocCancelIdle
                mopIdleStartControl = 0#
                mopIdleStopControl = 0#
                If .Arguments.Count = 0 Then
                    .Remove
                ElseIf .TimeSince >= .Arguments(1).Value Then
                    .Remove
                End If
                
            Case ocStartIdle
                CheckInterlock CStr(ocStartIdle)
                If mipAuxDPLCCmdRequest = 0 Then
                    If .FirstTime Then
                        RequestControlFromAuxDPLC
                    ElseIf .TimeSince >= 10# Then
                        .Abort LoadResString(ridUnableToGetControlFromPLC, "AuxD"), cmdTimedout
                    End If
                ElseIf mopWsMode > modeManual Then
                    .InsertPredecessor ocSetModeManual, "SetModeManual"
                ElseIf Abs(mgvfFsTarget) > mccfMaxFsLimit(1) Then
                    mobjCommands.Insert ocSetFanSpeed, "SetFanspeed", mccfMaxFsLimit(1), mccfMaxFsRamp
                ElseIf mgvfFsTarget < 0# Then
                    mobjCommands.Insert ocSetFanSpeed, "SetFanspeed", 0#, mccfMaxFsRamp
                ElseIf (Abs(mopFsSetpt) > mccfMaxFsLimit(1)) And (.TimeSince <= 10#) Then
                    ' wait
                ElseIf (mopIdleStopControl = 1#) Then
                    .InsertPredecessor ocCancelIdle, "CancelIdle", mccfCancelIdleTime
                Else
                    .Messages.Flush
                    mopIdleStartControl = 1#
                    mopIdleStopControl = 0#
                    
                    If mopIdleStatus = isIdle Then
                        If (mccfIdleFanspeed > 0) Then
                            If mobjInterlocks(CStr(ocSetFanSpeed)).IsOK Then
                                If (mipMfStatus = 1#) And (mgvfFsTarget < mccfIdleFanspeed) Then
                                    mobjCommands.Insert ocSetFanSpeed, "SetFanspeed", mccfIdleFanspeed
                                End If
                            End If
                        End If
                        .Remove
                    ElseIf .TimeSince > mccfIdleTimeout Then
                        .Abort , cmdTimedout
                    End If
                End If

            Case ocStopIdle
                CheckInterlock CStr(ocStopIdle)
                If mopIdleStatus = isNormal Then
                    mopIdleStartControl = 0#
                    mopIdleStopControl = 0#
                    If (mopWsMode = modeTrack) Then
                        mgvbAutoIdle = False
                    End If
                    .Remove
                ElseIf mipAuxDPLCCmdRequest = 0 Then
                    If .FirstTime Then
                        RequestControlFromAuxDPLC
                    ElseIf .TimeSince >= 10# Then
                        .Abort LoadResString(ridUnableToGetControlFromPLC, "AuxD"), cmdTimedout
                    End If
                ElseIf (mopWsMode = modeTrack) And (mgvbAutoIdle = True) Then
                    mgvbStopAutoIdle = True
                    .Remove
                ElseIf mopWsMode > modeManual Then
                    .InsertPredecessor ocSetModeManual, "SetModeManual"
                ElseIf mgvfFsTarget > mccfMaxFsLimit(1) Then
                    mobjCommands.Insert ocSetFanSpeed, "SetFanspeed", mccfMaxFsLimit(1), mccfMaxFsRamp
                ElseIf (mopFsSetpt > mccfMaxFsLimit(1)) And (.TimeSince <= 10#) Then
                    ' wait
                ElseIf (mopIdleStartControl = 1#) Then
                    .InsertPredecessor ocCancelIdle, "CancelIdle", mccfCancelIdleTime
                Else
                    .Messages.Flush
                    mopIdleStartControl = 0#
                    mopIdleStopControl = 1#
                    If .TimeSince > mccfIdleTimeout Then
                        .Abort , cmdTimedout
                    End If
                End If
                            
            Case ocStartPurge
                CheckInterlock CStr(ocStartPurge)
                If WindTunnel = wtAAWT Then
                    If mipAuxDPLCCmdRequest = 0 Then
                        If .FirstTime Then
                            RequestControlFromAuxDPLC
                        ElseIf .TimeSince >= 10# Then
                            .Abort LoadResString(ridUnableToGetControlFromPLC, "AuxD"), cmdTimedout
                        End If
                    Else
                        mopPurgeCmd = 1
                        .Remove
                    End If
                
                Else
                    If mipCoolPLCCmdRequest = 0 Then
                        If .FirstTime Then
                            RequestControlFromCoolPLC
                        ElseIf .TimeSince >= 10# Then
                            .Abort LoadResString(ridUnableToGetControlFromPLC, "Cooling"), cmdTimedout
                        End If
                    Else
                        mopPurgeCmd = 1
                        .Remove
                    End If
                End If
                
            Case ocStopPurge
                CheckInterlock CStr(ocStopPurge)
                If WindTunnel = wtAAWT Then
                    If mipAuxDPLCCmdRequest = 0 Then
                        If .FirstTime Then
                            RequestControlFromAuxDPLC
                        ElseIf .TimeSince >= 10# Then
                            .Abort LoadResString(ridUnableToGetControlFromPLC, "AuxD"), cmdTimedout
                        End If
                    Else
                        mopPurgeCmd = 0
                        .Remove
                    End If
                
                Else
                    If mipCoolPLCCmdRequest = 0 Then
                        If .FirstTime Then
                            RequestControlFromCoolPLC
                        ElseIf .TimeSince >= 10# Then
                            .Abort LoadResString(ridUnableToGetControlFromPLC, "Cooling"), cmdTimedout
                        End If
                    Else
                        mopPurgeCmd = 0
                        .Remove
                    End If
                End If
                
            Case ocSetNozzleMethod
                mopWSCalcMethod = 0
                .Remove
                
            Case ocSetPlenumMethod
                mopWSCalcMethod = 1
                .Remove
                
            Case ocSetNozzleSize
                CheckInterlock CStr(ocSetNozzleSize)
                If mipAuxDPLCCmdRequest = 0 Then
                    If .FirstTime Then
                        RequestControlFromAuxDPLC
                    ElseIf .TimeSince >= 10# Then
                        .Abort LoadResString(ridUnableToGetControlFromPLC, "AuxD"), cmdTimedout
                    End If
                Else
                    mgvlNozzleTarget = .Arguments(1)
                    mgvdNozzleSizeCmdTime = Now
                    .Remove
                End If
                
            Case ocStopNozzleMotion
                CheckInterlock CStr(ocStopNozzleMotion)
                If mipAuxDPLCCmdRequest = 0 Then
                    If .FirstTime Then
                        RequestControlFromAuxDPLC
                    ElseIf .TimeSince >= 10# Then
                        .Abort LoadResString(ridUnableToGetControlFromPLC, "AuxD"), cmdTimedout
                    End If
                Else
                    mgvlNozzleTarget = 0
                    .Remove
                End If
                
            Case ocSetCFlapLeft
                CheckInterlock CStr(ocSetCFlapLeft)
                If mipAuxDPLCCmdRequest = 0 Then
                    If .FirstTime Then
                        RequestControlFromAuxDPLC
                    ElseIf .TimeSince >= 10# Then
                        .Abort LoadResString(ridUnableToGetControlFromPLC, "AuxD"), cmdTimedout
                    End If
                Else
                    mopCFLeft1Target = .Arguments(1)
                    mopCFLeft2Target = .Arguments(2)
                    mgvdCFLeftCmdTime = Now
                    mgvbCFLeftCtl = True
                    .Remove
                End If
                
            Case ocSetCFlapLeft1
                CheckInterlock CStr(ocSetCFlapLeft1)
                If mipAuxDPLCCmdRequest = 0 Then
                    If .FirstTime Then
                        RequestControlFromAuxDPLC
                    ElseIf .TimeSince >= 10# Then
                        .Abort LoadResString(ridUnableToGetControlFromPLC, "AuxD"), cmdTimedout
                    End If
                Else
                    mopCFLeft1Target = .Arguments(1)
                    mgvdCFLeftCmdTime = Now
                    mgvbCFLeftCtl = True
                    .Remove
                End If
                
            Case ocSetCFlapLeft2
                CheckInterlock CStr(ocSetCFlapLeft2)
                If mipAuxDPLCCmdRequest = 0 Then
                    If .FirstTime Then
                        RequestControlFromAuxDPLC
                    ElseIf .TimeSince >= 10# Then
                        .Abort LoadResString(ridUnableToGetControlFromPLC, "AuxD"), cmdTimedout
                    End If
                Else
                    mopCFLeft2Target = .Arguments(1)
                    mgvdCFLeftCmdTime = Now
                    mgvbCFLeftCtl = True
                    .Remove
                End If
                
            Case ocSetCFlapRight
                CheckInterlock CStr(ocSetCFlapRight)
                If mipAuxDPLCCmdRequest = 0 Then
                    If .FirstTime Then
                        RequestControlFromAuxDPLC
                    ElseIf .TimeSince >= 10# Then
                        .Abort LoadResString(ridUnableToGetControlFromPLC, "AuxD"), cmdTimedout
                    End If
                Else
                    mopCFRight1Target = .Arguments(1)
                    mopCFRight2Target = .Arguments(2)
                    mgvdCFRightCmdTime = Now
                    mgvbCFRightCtl = True
                    .Remove
                End If
                
            Case ocSetCFlapRight1
                CheckInterlock CStr(ocSetCFlapRight1)
                If mipAuxDPLCCmdRequest = 0 Then
                    If .FirstTime Then
                        RequestControlFromAuxDPLC
                    ElseIf .TimeSince >= 10# Then
                        .Abort LoadResString(ridUnableToGetControlFromPLC, "AuxD"), cmdTimedout
                    End If
                Else
                    mopCFRight1Target = .Arguments(1)
                    mgvdCFRightCmdTime = Now
                    mgvbCFRightCtl = True
                    .Remove
                End If
                
            Case ocSetCFlapRight2
                CheckInterlock CStr(ocSetCFlapRight2)
                If mipAuxDPLCCmdRequest = 0 Then
                    If .FirstTime Then
                        RequestControlFromAuxDPLC
                    ElseIf .TimeSince >= 10# Then
                        .Abort LoadResString(ridUnableToGetControlFromPLC, "AuxD"), cmdTimedout
                    End If
                Else
                    mopCFRight2Target = .Arguments(1)
                    mgvdCFRightCmdTime = Now
                    mgvbCFRightCtl = True
                    .Remove
                End If
                
            Case ocSetCFlapTop
                CheckInterlock CStr(ocSetCFlapTop)
                If mipAuxDPLCCmdRequest = 0 Then
                    If .FirstTime Then
                        RequestControlFromAuxDPLC
                    ElseIf .TimeSince >= 10# Then
                        .Abort LoadResString(ridUnableToGetControlFromPLC, "AuxD"), cmdTimedout
                    End If
                Else
                    mopCFTopTarget = .Arguments(1)
                    mgvdCFTopCmdTime = Now
                    mgvbCFTopCtl = True
                    .Remove
                End If
                
            Case ocStopCF
                CheckInterlock CStr(ocStopCF)
                If mipAuxDPLCCmdRequest = 0 Then
                    If .FirstTime Then
                        RequestControlFromAuxDPLC
                    ElseIf .TimeSince >= 10# Then
                        .Abort LoadResString(ridUnableToGetControlFromPLC, "AuxD"), cmdTimedout
                    End If
                Else
                    mgvbCFLeftCtl = False
                    mgvbCFRightCtl = False
                    mgvbCFTopCtl = False
                    .Remove
                End If
                
            Case Else
                .Abort "Unrecognized opcode: " & .Opcode
            End Select
        End Select
    End With
    
    GoTo Done
Error:
    lRaiseErrorNumber = CCommandError(Err.Number)
    sRaiseErrorText = Err.Description
'    If (Err.Number And vbObjectError) <> vbObjectError Then
'        EH.Module = TypeName(Me)
'        EH.Procedure = "ExecuteCommands"
'        Call EH_MsgBox
'    End If
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
Private Sub mobjCommands_CommandInserted(NewCommand As CommandListLib.Command, ExistingCommand As CommandListLib.Command)
'
    On Error GoTo Error
    With NewCommand
        Select Case .Opcode
        
        Case ocStartMainFanAux
            .Messages.Add LoadResString(ridAuxCurrentOff), mtInformation
        
        Case ocStartMainFan
            .Messages.Add LoadResString(ridMFCurrentOff), mtInformation
        
        Case ocStopMainFan
            .Messages.Add LoadResString(ridMFCurrentOn), mtInformation
        
'        Case ocSetFanSpeed
'            .Messages.Add LoadResString(ridFanspeedSetTo, CStr(.Arguments(1))), mtInformation
        
        Case ocStopIdle
            .Messages.Add LoadResString(ridIdleCurrentOn), mtInformation
            
        End Select
    End With
    
    GoTo Done
Error:
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = "mobjCommands_CommandInserted"
        Call EH_MsgBox
    End If
    Resume Done
Done:
    On Error Resume Next
    Exit Sub
End Sub

Private Sub mobjInterlocks_Evaluate(InterlockName As String, CurrentOKValue As Boolean, AdditionalMessages As Collection)
'
    Dim fCheckValue As Single
    Dim I As Long
    
    On Error Resume Next
    Select Case InterlockName
    Case CStr(ocStartMainFan), CStr(ocStopMainFan), CStr(ocStartMainFanAux), CStr(ocStopMainFanAux), CStr(ocStartIdle), CStr(ocStopIdle), CStr(ocShowStatus), _
         CStr(ocSetNozzleSize), CStr(ocSetCFlapLeft1), CStr(ocSetCFlapLeft2), CStr(ocSetCFlapRight1), CStr(ocSetCFlapRight2), CStr(ocSetCFlapTop), CStr(ocStopCF)
        Select Case True
        Case mipOpModeTest = 1, mipOpModeStab = 1, mipOpModePrep = 1
        Case mipOpModeMaint = 1
            CurrentOKValue = False
            AdditionalMessages.Add GetCommonMessage(cmWTModeMaint)
        Case mipOpModeOff = 1
            CurrentOKValue = False
            AdditionalMessages.Add GetCommonMessage(cmWTModeOff)
        Case Else
            CurrentOKValue = False
            AdditionalMessages.Add GetCommonMessage(cmWTModeNone)
        End Select
    End Select
    
    Select Case InterlockName
    Case CStr(ocSetWindSpeed)
        On Error Resume Next
        fCheckValue = mvCheckIlckValue
        If Err = 0 Then
            If mipOpModePrep = 1 Then
                If (fCheckValue > mccfMaxWsPrepMode) Then
                    CurrentOKValue = False
                    AdditionalMessages.Add GetCommonMessage2(cmInPrepModeMaxIs, mccfMaxWsPrepMode)
                End If
            ElseIf (mipOpModeStab = 1) And (mgvlWindTunnel = wtAAWT) Then
                If (fCheckValue > mccfMaxWsStabMode) Then
                    CurrentOKValue = False
                    AdditionalMessages.Add GetCommonMessage2(cmInStabModeMaxIs, mccfMaxWsStabMode)
                End If
            End If
        End If
    Case CStr(ocSetFanSpeed)
        On Error Resume Next
        fCheckValue = mvCheckIlckValue
        If Err = 0 Then
            If mipOpModePrep = 1 Then
                If (fCheckValue > mccfMaxFsPrepMode) Then
                    CurrentOKValue = False
                    AdditionalMessages.Add GetCommonMessage2(cmInPrepModeMaxIs, mccfMaxFsPrepMode)
                End If
            ElseIf (mipOpModeStab = 1) And (mgvlWindTunnel = wtAAWT) Then
                If (fCheckValue > mccfMaxFsStabMode) Then
                    CurrentOKValue = False
                    AdditionalMessages.Add GetCommonMessage2(cmInStabModeMaxIs, mccfMaxFsStabMode)
                End If
            End If
        End If
    Case CStr(ocTrackWindSpeed)
        Select Case True
        Case mipOpModeTest = 1, mipOpModeStab = 1
        Case mipOpModePrep = 1
            CurrentOKValue = False
            AdditionalMessages.Add GetCommonMessage(cmWTModePrep)
        Case mipOpModeMaint = 1
            CurrentOKValue = False
            AdditionalMessages.Add GetCommonMessage(cmWTModeMaint)
        Case mipOpModeOff = 1
            CurrentOKValue = False
            AdditionalMessages.Add GetCommonMessage(cmWTModeOff)
        Case Else
            CurrentOKValue = False
            AdditionalMessages.Add GetCommonMessage(cmWTModeNone)
        End Select
    
    End Select
    
    Select Case InterlockName
    Case CStr(ocStartMainFan), CStr(ocShowStatus)
        If mipMfPerm = 0# Then
            ' Add known components of this interlock to the message
            If mccsMFPermComponents <> "" Then
                mobjMFPermComponents.Check CurrentOKValue, AdditionalMessages
            End If
        End If
        If (mipMfStatus = 0#) And (mopMfControl = 0#) Then
            If SecondsSinceTime(mgvdDriveOffTime) < mccfMainFanRestartTime Then
                CurrentOKValue = False
                AdditionalMessages.Add LoadResString(ridMainFanRestartDelay, CLng(mccfMainFanRestartTime - SecondsSinceTime(mgvdDriveOffTime)))
            End If
        End If
    End Select
    
    Select Case InterlockName
    Case CStr(ocStartIdle), CStr(ocStopIdle), CStr(ocShowStatus), CStr(ocShowIdleStatus)
        If mipIdleNozFlapMotorStop = 1# Then
            ' Add known components of this interlock to the message
            If mccsIdleCmdPermComponents <> "" Then
                mobjIdleCmdPermComponents.Check CurrentOKValue, AdditionalMessages
            End If
        End If
    End Select
    Select Case InterlockName
    Case CStr(ocStartMainFanAux), CStr(ocShowStatus)
        If mipAuxReady = 0# Then
            ' Add known components of this interlock to the message
            If mccsMFAuxPermComponents <> "" Then
                mobjMFAuxPermComponents.Check CurrentOKValue, AdditionalMessages
            End If
        End If
    End Select
        
    Select Case InterlockName
    Case CStr(ocShowStatus)
        For I = 1 To 3
            If LimitInEffect(I) Then
                If mopWsMode > modeManual Then
                    If mgvfWsTarget > (mccfMaxFsLimit(I) * mgvfNomWsMax / mgvfNomFsMax) Then
                        CurrentOKValue = False
                        AdditionalMessages.Add LoadResString(ridLowSpeedLmt1On + I - 1, mccfMaxFsLimit(I))
                    End If
                ElseIf mgvfFsTarget > mccfMaxFsLimit(I) Then
                    CurrentOKValue = False
                    AdditionalMessages.Add LoadResString(ridLowSpeedLmt1On + I - 1, mccfMaxFsLimit(I))
                End If
            End If
        Next I
        
        If mopWsMode >= modeAuto Then
            If mobjPID.States(ctlIntegral) >= mccfIntMax Then
                AdditionalMessages.Add LoadResString(ridPIDIntegralAtMax)
                CurrentOKValue = False
            End If
            If mobjPID.States(ctlIntegral) <= mccfIntMin Then
                AdditionalMessages.Add LoadResString(ridPIDIntegralAtMin)
                CurrentOKValue = False
            End If
        End If
        If mopWsMode = modeAuto Then
            If mgvfWsTarget <= mccfMinWsForPid Then
                AdditionalMessages.Add LoadResString(ridWsTooLowForPid, mccfMinWsForPid)
                CurrentOKValue = False
            End If
        End If
    Case CStr(ocSetWindSpeed)
        For I = 2 To 3      ' Dont include idle limit check here -- Set WS will stop the idle
            If LimitInEffect(I) Then
                On Error Resume Next
                fCheckValue = mvCheckIlckValue
                If Err = 0 Then
                    If (fCheckValue / mgvfNomWsMax) * mgvfNomFsMax >= mccfMaxFsLimit(I) Then
                        CurrentOKValue = False
                        AdditionalMessages.Add LoadResString(ridLowSpeedLmt1On + I - 1, mccfMaxFsLimit(I))
                    End If
                End If
            End If
        Next I
        
    Case CStr(ocSetFanSpeed)
        For I = 2 To 3
            If LimitInEffect(I) Then
                On Error Resume Next
                fCheckValue = mvCheckIlckValue
                If Err = 0 Then
                    If fCheckValue >= mccfMaxFsLimit(I) Then
                        CurrentOKValue = False
                        AdditionalMessages.Add LoadResString(ridLowSpeedLmt1On + I - 1, mccfMaxFsLimit(I))
                    End If
                End If
            End If
        Next I
        
    End Select
    
    Exit Sub
End Sub

Friend Sub Initialize(Optional InitType As InitTypes = initAll, _
            Optional Command As Command)

'Perform basic initialization/startup functions

    Const PROC_NAME = "Initialize"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    Dim bStatus As Boolean, sText As String
    Dim sFileName As String
    Dim objEH As EH
    Dim bExists As Boolean
    Dim sDefaultsFileName As String
    Dim sPath As String
    Dim lNumComponents As Long
    Dim sComponents() As String
    Dim lNumItems As Long
    Dim sItems() As String
    Dim I As Long
    Dim J As Long
    
    On Error GoTo Error
    bRaiseError = False

    Select Case InitType
    Case initConstants, initAll
        bStatus = True
        sPath = GetDefaultCCFolder(App.Path)
        sFileName = sPath & PROCESS_NAME & ".csv"
        sFileName = Reg_GetValue(HKEY_LOCAL_MACHINE, REG_PROCESSES, PROCESS_NAME, "ControlConstantsFile", sFileName)
        sDefaultsFileName = App.Path & "\" & App.EXEName & "Defaults.csv"
        With mobjConstants
            .OpenFileWithDefaults sFileName, sDefaultsFileName
            .AddString "Priority", mccsPriority, bStatus, sText
            .AddSingle "UpdateInterval", mccfUpdateInt, bStatus, sText
            .AddString "MFPermComponents", mccsMFPermComponents, bStatus, sText
            .AddString "MFAuxPermComponents", mccsMFAuxPermComponents, bStatus, sText
            .AddSingle "SlowFilterDelayTime", mccfSlowFilterDelayTime, bStatus, sText
            .AddSingle "TrackingFilterTC", mccfWsTgtTrackFilTc, bStatus, sText
            .AddBoolean "OutputFilterEnable", mccbFilterOutput, bStatus, sText
            .AddSingle "OutputFilterTC", mccfPidOutputFilTc, bStatus, sText
            .AddSingle "MainFanTimeout", mccfMainFanTimeout, bStatus, sText
            .AddSingle "MainFanReadyDelay", mccfDrivePause, bStatus, sText
            .AddSingle "MainFanAuxTimeout", mccfMainFanAuxTimeout, bStatus, sText
            .AddSingle "MainFanAuxPause", mccfMainFanAuxPause, bStatus, sText
            .AddSingle "MainFanAuxAutoOff", mccfMainFanAuxAutoOffTime, bStatus, sText
            .AddSingle "MainFanRestartTime", mccfMainFanRestartTime, bStatus, sText
            .AddSingle "IdleTimeout", mccfIdleTimeout, bStatus, sText
            .AddSingle "SuspendTimeout", mccfMaxSuspendTime, bStatus, sText
            .AddSingle "ResetTime", mccfResetFaultTime, bStatus, sText
            .AddBoolean "ResetAuto", mccbAutoReset, bStatus, sText
            
            
            ' CloseTime
            ' -----------------------
            For J = 1 To IIf(WindTunnel = wtCWT, 3, 1)
                mcclNumCloseTimes(J) = 1
                .AddLong "CloseTime" & NozzleName(J) & "Num", mcclNumCloseTimes(J), bStatus, sText
                If mcclNumCloseTimes(J) > 10 Then
                    mcclNumCloseTimes(J) = 10
                End If
                If J = nsLarge Then
                    For I = 1 To mcclNumCloseTimes(J)
                        .AddSingle "CloseTime" & NozzleName(J) & Format(I, "00") & "Time", mccfCloseTimesLarge(I), bStatus, sText
                        .AddSingle "CloseTime" & NozzleName(J) & Format(I, "00") & "Speed", mccfCloseTimeSpeedsLarge(I), bStatus, sText
                    Next I
                ElseIf J = nsMedium Then
                    For I = 1 To mcclNumCloseTimes(J)
                        .AddSingle "CloseTime" & NozzleName(J) & Format(I, "00") & "Time", mccfCloseTimesMedium(I), bStatus, sText
                        .AddSingle "CloseTime" & NozzleName(J) & Format(I, "00") & "Speed", mccfCloseTimeSpeedsMedium(I), bStatus, sText
                    Next I
                Else
                    For I = 1 To mcclNumCloseTimes(J)
                        .AddSingle "CloseTime" & NozzleName(J) & Format(I, "00") & "Time", mccfCloseTimesSmall(I), bStatus, sText
                        .AddSingle "CloseTime" & NozzleName(J) & Format(I, "00") & "Speed", mccfCloseTimeSpeedsSmall(I), bStatus, sText
                    Next I
                End If
                .AddSingle "CloseTime" & NozzleName(J) & "FilterTC", mccfCloseWsFilTc(J), bStatus, sText
                .AddSingle "CloseTime" & NozzleName(J) & "WSRate", mccfCloseWsRamp(J), bStatus, sText
            Next J
            
            ' Filter Coefs of DeltP, and FS
            ' -------------------------------------
            .AddBoolean "DeltaPFilterEnable", mccbFilterInput, bStatus, sText
            .AddSingle "DeltaPFilterType", mccfDisplayFilType, bStatus, sText
            .AddSingle "DeltaPFilterMin", mccfMinFilTc, bStatus, sText
            .AddSingle "DeltaPFilterMax", mccfMaxFilTc, bStatus, sText
            .AddSingle "DeltaPFilterSlowError", mccfMinWsErrorSlow, bStatus, sText
            .AddSingle "DeltaPFilterSlowErrorSmall", mccfMinWsErrorSlowSmall, bStatus, sText
            
            
            For J = 1 To IIf(WindTunnel = wtCWT, 3, 1)
                .AddSingle "DeltaPFilter" & NozzleName(J) & "Base", mccfWsFilterTc0(J), bStatus, sText
                .AddSingle "DeltaPFilter" & NozzleName(J) & "Adjust", mccfWsFilterTc1(J), bStatus, sText
                .AddSingle "DeltaPFilter" & NozzleName(J) & "Fast", mccfWsFilterFast(J), bStatus, sText
                .AddSingle "DeltaPFilter" & NozzleName(J) & "FastTrack", mccfWsFilterFastTrack(J), bStatus, sText
                .AddSingle "DeltaPControlFilter" & NozzleName(J) & "Base", mccfControlFilTc0(J), bStatus, sText
                .AddSingle "DeltaPControlFilter" & NozzleName(J) & "Adjust", mccfControlFilTc1(J), bStatus, sText
                .AddSingle "DeltaPControlFilter" & NozzleName(J) & "Min", mccfControlFilTcMin(J), bStatus, sText
                .AddSingle "DeltaPControlFilter" & NozzleName(J) & "Max", mccfControlFilTcMax(J), bStatus, sText
            Next J
                        
            ' Feedforward Constants
            .AddSingle "FFGain", mccfFfGain, bStatus, sText
            .AddSingle "FFLag", mccfFfLagCoef, bStatus, sText
            .AddSingle "FFLagTrack", mccfFfLagCoefTrack, bStatus, sText
            .AddSingle "FFLead", mccfFfLeadCoef, bStatus, sText
            .AddSingle "FFLeadTrack", mccfFfLeadCoefTrack, bStatus, sText
            For J = 1 To IIf(WindTunnel = wtCWT, 3, 1)
                .AddSingle "FF" & NozzleName(J) & "Fanspeed", mccfNomFs(J), bStatus, sText
                .AddSingle "FF" & NozzleName(J) & "Windspeed", mccfNomWs(J), bStatus, sText
            Next J
            
            ' Limits
            .AddSingle "FanspeedMin", mccfMinFs, bStatus, sText
            .AddSingle "FanspeedMinAutoMode", mccfMinFsAutoMode, bStatus, sText
            If WindTunnel = wtCWT Then
                .AddSingle "FanspeedMaxLarge", mccfMaxFs(nsLarge), bStatus, sText
                .AddSingle "FanspeedMaxMedium", mccfMaxFs(nsMedium), bStatus, sText
                .AddSingle "FanspeedMaxSmall", mccfMaxFs(nsSmall), bStatus, sText
            Else
                .AddSingle "FanspeedMaxLarge", mccfMaxFs(nsLarge), bStatus, sText
            End If
            .AddSingle "FanspeedMaxLimit1", mccfMaxFsLimit(1), bStatus, sText
            .AddSingle "FanspeedMaxLimit2", mccfMaxFsLimit(2), bStatus, sText
            .AddSingle "FanspeedMaxLimit3", mccfMaxFsLimit(3), bStatus, sText
            .AddSingle "FanspeedRampMin", mccfMinFsRamp, bStatus, sText
            .AddSingle "FanspeedRampMax", mccfMaxFsRamp, bStatus, sText
            .AddSingle "FanspeedRampMaxTrack", mccfMaxFsRampTrack, bStatus, sText
            .AddSingle "FanspeedRampDefault", mccfDefFsRamp, bStatus, sText
            .AddSingle "FanspeedMaxPrep", mccfMaxFsPrepMode, bStatus, sText
            If WindTunnel = wtAAWT Then
                .AddSingle "FanspeedMaxStab", mccfMaxFsStabMode, bStatus, sText
                .AddSingle "WindspeedRampMaxLowSpeed", mccfMaxWsRampLowSpeeds, bStatus, sText
                .AddSingle "FanspeedSlowRamp", mccfSlowRampFanspeed, bStatus, sText
            End If
                        
            .AddSingle "WindspeedMin", mccfMinWs, bStatus, sText
            If WindTunnel = wtCWT Then
                .AddSingle "WindspeedMaxLarge", mccfMaxWs(nsLarge), bStatus, sText
                .AddSingle "WindspeedMaxMedium", mccfMaxWs(nsMedium), bStatus, sText
                .AddSingle "WindspeedMaxSmall", mccfMaxWs(nsSmall), bStatus, sText
            Else
                .AddSingle "WindspeedMaxLarge", mccfMaxWs(nsLarge), bStatus, sText
            End If
            .AddSingle "WindspeedRampMin", mccfMinWsRamp, bStatus, sText
            .AddSingle "WindspeedRampMax", mccfMaxWsRamp, bStatus, sText
            .AddSingle "WindspeedRampDefault", mccfDefWsRamp, bStatus, sText
            .AddSingle "WindspeedMaxPrep", mccfMaxWsPrepMode, bStatus, sText
            If WindTunnel = wtAAWT Then
                .AddSingle "WindspeedMaxStab", mccfMaxWsStabMode, bStatus, sText
            End If
            
            .AddSingle "HeadwindMax", mccfMaxWsHeadwind, bStatus, sText
            .AddSingle "HeadwindMin", mccfMinWsHeadwind, bStatus, sText
            
            .AddSingle "FanspeedMonLimit", mccfFsMonitorError, bStatus, sText
            .AddSingle "FanspeedMonFiltTC", mccfFsMonitorFilTc, bStatus, sText
            .AddSingle "FanspeedMonPeriod", mccfFsMonitorPeriod, bStatus, sText
            .AddSingle "WindspeedMonLimit", mccfMonitorError, bStatus, sText
            .AddSingle "WindspeedMonFiltTC", mccfMonitorFilTc, bStatus, sText
            .AddSingle "WindspeedMonPeriod", mccfMonitorPeriod, bStatus, sText
           
            
            ' PID coefficients
            ' -------------------------------------
            .AddSingle "PIDIntegralMax", mccfIntMax, bStatus, sText
            .AddSingle "PIDIntegralMin", mccfIntMin, bStatus, sText
            .AddSingle "PIDOutputMax", mccfCtlMax, bStatus, sText
            .AddSingle "PIDOutputMin", mccfCtlMin, bStatus, sText
            .AddSingle "PIDDerivFilterTC", mccfTdFilterTC, bStatus, sText
            .AddSingle "PIDErrorMax", mccfMaxWsError, bStatus, sText
            .AddSingle "PIDWindspeedMin", mccfMinWsForPid, bStatus, sText
            .AddSingle "PIDWindspeedMinDelay", mccfMinWsForPidDelay, bStatus, sText
            
            For J = 1 To IIf(WindTunnel = wtCWT, 3, 1)
                 .AddSingle "PIDGain" & NozzleName(J) & "Base", mccfKp0(J), bStatus, sText
                 .AddSingle "PIDGain" & NozzleName(J) & "Adjust", mccfKp1(J), bStatus, sText
                 .AddSingle "PIDGain" & NozzleName(J) & "TrackBase", mccfKp0Track(J), bStatus, sText
                 .AddSingle "PIDGain" & NozzleName(J) & "TrackAdjust", mccfKp1Track(J), bStatus, sText
                 .AddSingle "PIDGain" & NozzleName(J) & "Min", mccfKpMin(J), bStatus, sText
                 .AddSingle "PIDGain" & NozzleName(J) & "Max", mccfKpMax(J), bStatus, sText
                 
                 .AddSingle "PIDDeriv" & NozzleName(J) & "Base", mccfTd0(J), bStatus, sText
                 .AddSingle "PIDDeriv" & NozzleName(J) & "Adjust", mccfTd1(J), bStatus, sText
                 .AddSingle "PIDDeriv" & NozzleName(J) & "TrackBase", mccfTd0Track(J), bStatus, sText
                 .AddSingle "PIDDeriv" & NozzleName(J) & "TrackAdjust", mccfTd1Track(J), bStatus, sText
                 .AddSingle "PIDDeriv" & NozzleName(J) & "Min", mccfTdMin(J), bStatus, sText
                 .AddSingle "PIDDeriv" & NozzleName(J) & "Max", mccfTdMax(J), bStatus, sText
                 
                 .AddSingle "PIDIntegral" & NozzleName(J) & "Base", mccfTi0(J), bStatus, sText
                 .AddSingle "PIDIntegral" & NozzleName(J) & "Adjust", mccfTi1(J), bStatus, sText
                 .AddSingle "PIDIntegral" & NozzleName(J) & "TrackBase", mccfTi0Track(J), bStatus, sText
                 .AddSingle "PIDIntegral" & NozzleName(J) & "TrackAdjust", mccfTi1Track(J), bStatus, sText
                 .AddSingle "PIDIntegral" & NozzleName(J) & "Min", mccfTiMin(J), bStatus, sText
                 .AddSingle "PIDIntegral" & NozzleName(J) & "Max", mccfTiMax(J), bStatus, sText
            Next J
            .AddSingle "PIDIntegralMaxWSRateMax", mccfMaxWsRateForI, bStatus, sText
            .AddSingle "PIDIntegralMaxWSRateMin", mccfMinWsRateForI, bStatus, sText
            .AddSingle "PIDIntegralMaxWSRateBase", mccfWsRateForI0, bStatus, sText
            .AddSingle "PIDIntegralMaxWSRateAdjust", mccfWsRateForI1, bStatus, sText
                                    
            .AddSingle "DeltaPMinForWS", mccfMinDeltaP, bStatus, sText
            .AddSingle "DeltaPMinForWSFSMax", mccfMinDeltaPFSMax, bStatus, sText
                                    
            If WindTunnel = wtCWT Then
                .AddSingle "IdleFanSpeed", mccfIdleFanspeed, bStatus, sText
                .AddBoolean "AutoIdleDefault", mccbDefAutoIdle, bStatus, sText
                .AddSingle "AutoIdleStartWS", mccfWSStartIdle, bStatus, sText
                .AddSingle "AutoIdleStopWS", mccfWSStopIdle, bStatus, sText
                .AddSingle "AutoIdleDeadband", mccfAutoIdleDeadband, bStatus, sText
                .AddString "IdlePermComponents", mccsIdleCmdPermComponents, bStatus, sText
                .AddSingle "IdleReversePauseTime", mccfCancelIdleTime, bStatus, sText
                .AddSingle "NozzleTimeout", mccfNozzleTimeout, bStatus, sText
                
                .AddSingle "CFLeft1Min", mccfCFLeft1Min, bStatus, sText
                .AddSingle "CFLeft1Max", mccfCFLeft1Max, bStatus, sText
                .AddSingle "CFLeft2Min", mccfCFLeft2Min, bStatus, sText
                .AddSingle "CFLeft2Max", mccfCFLeft2Max, bStatus, sText
                .AddSingle "CFRight1Min", mccfCFRight1Min, bStatus, sText
                .AddSingle "CFRight1Max", mccfCFRight1Max, bStatus, sText
                .AddSingle "CFRight2Min", mccfCFRight2Min, bStatus, sText
                .AddSingle "CFRight2Max", mccfCFRight2Max, bStatus, sText
                .AddSingle "CFTopMin", mccfCFTopMin, bStatus, sText
                .AddSingle "CFTopMax", mccfCFTopMax, bStatus, sText
                .AddSingle "CFHoldTime", mccfCFHoldTime, bStatus, sText
                .AddSingle "CFTimeout", mccfCFTimeout, bStatus, sText
                .AddSingle "CFTolerance", mccfCFTolerance, bStatus, sText
            End If
            
            .AddString "TrackParDefault", mccsDefTrackParName, bStatus, sText
            .AddSingle "TurnOffFanTime", mccfTurnOffFanTime, bStatus, sText
            
            
            ' Windspeed calc Coeffs
            ' ------------------------
            For I = 1 To IIf(WindTunnel = wtCWT, 3, 1)
                .AddString "WSCoeffsFile" & NozzleName(I), mccsWCoeffsFile(I), bStatus, sText
            Next I
            
            .AddSingle "FanRunTimeWithoutCooling", mccfMaxRunTimeWithoutCooling, bStatus, sText
            
'            .AddSingle "RevWsStdC0", mccfRevWsStdC(0), bStatus, sText
'            .AddSingle "RevWsStdC1", mccfRevWsStdC(1), bStatus, sText
'            .AddSingle "RevWsStdC2", mccfRevWsStdC(2), bStatus, sText
'            .AddSingle "RevWsSmallC0", mccfRevWsSmallC(0), bStatus, sText
'            .AddSingle "RevWsSmallC1", mccfRevWsSmallC(1), bStatus, sText
'            .AddSingle "RevWsSmallC2", mccfRevWsSmallC(2), bStatus, sText
                                                
            If mgvlWindTunnel = wtCWT Then
                .AddBoolean "TrackDyTarget", mccbAutoFollowDyTarget, bStatus, sText
                .AddBoolean "TrackDyTargetUseTrackCoeffs", mccbAutoFollowDyTargetUseTrackCoeffs, bStatus, sText
                .AddBoolean "AutoIdleUsesLimit1", mccbAutoIdleUsesLimitCommand, bStatus, sText
            End If
            
            .AddSingle "FFReverseWindspeed", mccfNomWSReverse, bStatus, sText
            .AddSingle "FFReverseFanspeed", mccfNomFSReverse, bStatus, sText
            
            .AddSingle "ForceWSZeroTime", mccfTimeAtZeroToForceWSZero, bStatus, sText
            
            .AddSingle "PIDDerivUseTgtErrorLimit", mccfUseTargetErrorForDerivLimit, bStatus, sText
            .AddLong "SetptRampMethod", mcclSetpointRampMethod, bStatus, sText
            .AddSingle "FinalRampDeltaTime", mccfFinalRampTimeSec, bStatus, sText
            .AddSingle "FinalRampDeltaSpeed", mccfFinalRampDeltakph, bStatus, sText
            .AddBoolean "FinalRampSnap", mccbFinalRampSnap, bStatus, sText
            .AddBoolean "FinalRampUseSetpt", mccbFinalRampUseSetpoint, bStatus, sText
            .AddSingle "FinalRampMinWSRate", mccfWsRateForFinalRamp, bStatus, sText
            .AddBoolean "FinalRampSnapBumpless", mccbFinalRampSnapBumpless, bStatus, sText
            .AddSingle "FinalRampDeltaTimeSmall", mccfFinalRampTimeSecSmall, bStatus, sText
            .AddSingle "FinalRampDeltaSpeedSmall", mccfFinalRampDeltakphSmall, bStatus, sText
            
            If WindTunnel = wtAAWT Then
                For I = 1 To 18
                    .AddBoolean "OverflowFlap" & Format(I, "00") & "Open", mccbOFDefault(I), bStatus, sText
                Next I
                .AddBoolean "OverflowFlapPurgeOpen", mccbOFPurgeOpen, bStatus, sText
            End If
            
            If bStatus = False Then
                bRaiseError = True
                lRaiseErrorNumber = tceCommandFailed
                If Command Is Nothing Then
                Else
                    Command.SuppressLogging = True
                End If
                Set objEH = New EH
                With objEH
                    .AppError = LoadResString(StdTalentStrings.stsEHCantLoadConstants)
                    .Item(1) = sText
                    .FileName = sFileName
                    .Module = TypeName(Me)
                    .Procedure = PROC_NAME
                    sRaiseErrorText = objEH.Text
                    LogMessage2 objEH
                End With
                GoTo Done
            End If
                        
' Add control constants that are used by other classes to the Constants shared
' property group.
            With mobjSPGConstants
                .CreatePropertyByPosition(CC_FSMin, bExists).Value = mccfMinFs
                .CreatePropertyByPosition(CC_FSMax, bExists).Value = mgvfMaxFs
                .CreatePropertyByPosition(CC_FSRampMax, bExists).Value = mccfMaxFsRamp
                .CreatePropertyByPosition(CC_WSMin, bExists).Value = mccfMinWs
                .CreatePropertyByPosition(CC_WSMax, bExists).Value = mgvfMaxWs
                .CreatePropertyByPosition(CC_WSRampMax, bExists).Value = mccfMaxWsRamp
                .CreatePropertyByPosition(CC_HeadwindMin, bExists).Value = mccfMinWsHeadwind
                .CreatePropertyByPosition(CC_HeadwindMax, bExists).Value = mccfMaxWsHeadwind
                .CreatePropertyByPosition(CC_WSRampDef, bExists).Value = mccfDefWsRamp
                .CreatePropertyByPosition(CC_FSRampDef, bExists).Value = mccfDefFsRamp
                .CreatePropertyByPosition(CC_DefAutoIdle, bExists).Value = mccbDefAutoIdle
                .CreatePropertyByPosition(CC_DefTrackParName, bExists).Value = mccsDefTrackParName
                .CreatePropertyByPosition(CC_CFLeft1Min, bExists).Value = mccfCFLeft1Min
                .CreatePropertyByPosition(CC_CFLeft1Max, bExists).Value = mccfCFLeft1Max
                .CreatePropertyByPosition(CC_CFLeft2Min, bExists).Value = mccfCFLeft2Min
                .CreatePropertyByPosition(CC_CFLeft2Max, bExists).Value = mccfCFLeft2Max
                .CreatePropertyByPosition(CC_CFRight1Min, bExists).Value = mccfCFRight1Min
                .CreatePropertyByPosition(CC_CFRight1Max, bExists).Value = mccfCFRight1Max
                .CreatePropertyByPosition(CC_CFRight2Min, bExists).Value = mccfCFRight2Min
                .CreatePropertyByPosition(CC_CFRight2Max, bExists).Value = mccfCFRight2Max
                .CreatePropertyByPosition(CC_CFTopMin, bExists).Value = mccfCFTopMin
                .CreatePropertyByPosition(CC_CFTopMax, bExists).Value = mccfCFTopMax
            End With
            mgvbSeparateControl = True
        End With
               
        With mobjWSFilter                        'mopWS_Nozzle
            If mccfDisplayFilType = 2# Then
                Set mobjWSFilterBW = New Butterworth
                mobjWSFilterBW.InitializeLowPass 1 / mccfUpdateInt, 1# / (2# * 3.14159 * mccfWsFilterTc0(nsLarge)), 4
            ElseIf mccfDisplayFilType = 1# Then
                .FilterType = ctlRollingAverage
                .FilterConst(0) = mccfWsFilterTc0(nsLarge) / mccfUpdateInt
            Else
                .FilterType = ctlFirstOrder
                .FilterConst(0) = mccfWsFilterTc0(nsLarge)
            End If
            .UpdateTime = mccfUpdateInt
        End With
        With mobjWSFilter2
            If mccfDisplayFilType = 2# Then
                Set mobjWSFilterBW2 = New Butterworth
                mobjWSFilterBW2.InitializeLowPass 1 / mccfUpdateInt, 1# / (2# * 3.14159 * mccfWsFilterTc0(nsLarge)), 4
            ElseIf mccfDisplayFilType = 1# Then
                .FilterType = ctlRollingAverage
                .FilterConst(0) = mccfWsFilterTc0(nsLarge) / mccfUpdateInt
            Else
                .FilterType = ctlFirstOrder
                .FilterConst(0) = mccfWsFilterTc0(nsLarge)
            End If
            .UpdateTime = mccfUpdateInt
        End With
        With mobjWSctlFilter
            .FilterType = ctlFirstOrder
            .FilterConst(0) = mccfControlFilTcMin(nsLarge)
            .UpdateTime = mccfUpdateInt
        End With
        With mobjTrackFilter
            If mccfWsTgtTrackFilTc < mccfUpdateInt Then
                mccfWsTgtTrackFilTc = mccfUpdateInt
            End If
            .FilterType = ctlFirstOrder
            .FilterConst(0) = mccfWsTgtTrackFilTc
            .UpdateTime = mccfUpdateInt
        End With
        With mobjRampRateFilter
            .FilterType = ctlFirstOrder
            .UpdateTime = mccfUpdateInt
        End With
        
        With mobjWSMonitor
            .ErrorBand = mccfMonitorError
            .FilterTimeConst = mccfMonitorFilTc
            .MonitoringPeriod = mccfMonitorPeriod
            .UpdateTime = mccfUpdateInt
        End With
        With mobjFSMonitor
            .ErrorBand = mccfFsMonitorError
            .FilterTimeConst = mccfFsMonitorFilTc
            .MonitoringPeriod = mccfFsMonitorPeriod
            .UpdateTime = mccfUpdateInt
        End With
        
        With mobjFFLeadLag
            .FilterType = ctlLeadLagCompensation
            .UpdateTime = mccfUpdateInt
        End With
        
        mgvlDefPriority = GetControllerPriority(mccsPriority)
        mgvlOldNozzle = -1
                        
        If mccfIdleFanspeed > mccfMaxFsLimit(1) Then
            mccfIdleFanspeed = mccfMaxFsLimit(1)
        End If
        
        ' Load WS coeffs from the Excel files -- one per nozzle
        For I = 1 To IIf(WindTunnel = wtCWT, 3, 1)
            LoadWSCoeffs I, sPath & mccsWCoeffsFile(I)
        Next I
        
        
    End Select

    Select Case InitType
    Case initAll, initPars
        On Error Resume Next
        mobjParManager.RefreshObjects
        If Err Then
            lRaiseErrorNumber = CCommandError(Err.Number)
            sRaiseErrorText = Err.Description
            Set objEH = New EH
            With objEH
                Set .VBErr = Err
                .AppError = LoadResString(StdTalentStrings.stsEHInitializeFailedParsMissing)
                .Item(1) = Command.Text
                .Module = TypeName(Me)
                .Procedure = PROC_NAME
                If Command Is Nothing Then
                Else
                    .Item(1) = Command.Text
                End If
                LogMessage2 objEH
                If Command Is Nothing Then
                Else
                    Command.SuppressLogging = True
                End If
            End With
            GoTo Done
        End If
        
        If mopBlockageCorr = 0# Then
            mopBlockageCorr = 1#
        End If
        
        'mobjParManager.AttachOutputs
        mbInited = mobjParManager.Ready
        If mbInited Then
            mobjParManager.ReadInputs
        End If
        
        ' Set up Perm interlocks to show the operator the reason for the permissive to be faulted
        mobjMFPermComponents.SetComponents mccsMFPermComponents
        mobjMFAuxPermComponents.SetComponents mccsMFAuxPermComponents
        mobjIdleCmdPermComponents.SetComponents mccsIdleCmdPermComponents
        
        mgvdTurnOffFanTime = Now()
    End Select
    
    GoTo Done
Error:
    mbInited = False
    lRaiseErrorNumber = CCommandError(Err.Number)
    sRaiseErrorText = Err.Description
    bRaiseError = True
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
    Resume Done
    Resume
Done:
    On Error Resume Next
    mobjStatuses.SetValue STS_WINDSPEED_INITIALIZED, mbInited
    mobjConstants.CloseFile
    If mobjForm Is Nothing Then
        tmrMain.Enabled = True
    End If
    If mbInited Then
        SetProcessPriority mgvlDefPriority
    Else
        SetProcessPriority priNormal
    End If
    If bRaiseError Then
        On Error GoTo 0
        Call EH_RaiseError(lRaiseErrorNumber, TypeName(Me), PROC_NAME, sRaiseErrorText)
    End If
    Exit Sub
End Sub
Private Sub LoadWSCoeffs(Nozzle As NozzleSizes, FileName As String)
    '
    Const PROC_NAME = "LoadWSCoeffs"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim objApp As Object
    Dim objWB As Object
    Dim vValue As Variant
    Dim vValue2 As Variant
    Dim I As Long
    Dim J As Long
    Dim lCol As Long
    Dim sNozzleName As String
    
    On Error GoTo Error

    bRaiseError = False
    sNozzleName = NozzleName(Nozzle)
    
    CheckFileOpenable FileName, , , True
    Set objApp = CreateObject("ReACT.XLSXExporterLib.ExcelReader")
        
    On Error GoTo Error_OpeningFile
    Set objWB = objApp.OpenWorkbook(FileName, True)
        
    On Error GoTo Error_ReadingData
    vValue2 = objWB.Worksheets(1).Range("B7:J10").Value
    ' Convert returned 0 index based array to a one based array to make consistent with the arrays returned by Excel (and so that
    ' vValue(2,1) is row 2, column 1 in the Excel sheet
    ReDim vValue(1 To UBound(vValue2, 1) - LBound(vValue2, 1) + 1, 1 To UBound(vValue2, 2) - LBound(vValue2, 2) + 1) As Variant
    For I = LBound(vValue2, 1) To UBound(vValue2, 1)
        For J = LBound(vValue2, 2) To UBound(vValue2, 2)
            vValue(I - LBound(vValue2, 1) + 1, J - LBound(vValue2, 2) + 1) = vValue2(I, J)
        Next J
    Next I
    
    On Error GoTo Error
    
    For I = 0 To 6
        lCol = 7 - I
        mccdKqnCoeffs(Nozzle, I) = vValue(1, lCol)
        mccdKqplCoeffs(Nozzle, I) = vValue(2, lCol)
        mccdKpCoeffs(Nozzle, I) = vValue(3, lCol)
        mccdWSvsFSCoeffs(Nozzle, I) = vValue(4, lCol)
    Next I
    mccfKqnCoeffMin(Nozzle) = vValue(1, 8)
    mccfKqnCoeffMax(Nozzle) = vValue(1, 9)
    mccfKqplCoeffMin(Nozzle) = vValue(2, 8)
    mccfKqplCoeffMax(Nozzle) = vValue(2, 9)
    mccfKpCoeffMin(Nozzle) = vValue(3, 8)
    mccfKpCoeffMax(Nozzle) = vValue(3, 9)
    
    GoTo Done

Error:
    sRaiseErrorText = LoadResString(ridErrorLoadingWSCoeffs, sNozzleName, FileName, Err.Description)
    bRaiseError = True
    Resume Done
        
Error_OpeningFile:
    sRaiseErrorText = LoadResString(ridErrorOpeningFile, FileName, Err.Description)
    bRaiseError = True
    Resume Done

Error_ReadingData:
    sRaiseErrorText = LoadResString(ridErrorReadingData, FileName, "B7:J10", Err.Description)
    bRaiseError = True
    Resume Done
        
    
Done:
    On Error Resume Next
    objWB.Close False
    Set objWB = Nothing
    If bRaiseError Then
        On Error GoTo 0
        EH_RaiseError OLEERR_METHOD_FAILED, TypeName(Me), PROC_NAME, sRaiseErrorText
    End If
    
    Exit Sub
End Sub

Private Sub Suspend()

'Suspend control

    On Error Resume Next
    If mbInited Then
        StopFan True
    End If
    mbInited = False
    mobjParManager.ReleaseOutputs
    mopFsSetpt = 0#
    mopFsTargetDisp = 0#
    mgvfFsTarget = 0#
    mopWsMode = 0#
    mopMfControl = 0#
    mopAuxControl = 0#
    mopIdleStartControl = 0#
    mopIdleStopControl = 0#
    
    mobjStatuses(STS_WINDSPEED_INITIALIZED).SetValue mbInited
    SetProcessPriority priNormal
    Exit Sub
End Sub

Private Sub StopFan(Optional WaitForComplete As Boolean = False)

    Dim vStart As Date
    
    On Error GoTo Error
    If mbInited = False Then
        GoTo Done
    End If
    
    vStart = Now
    mgvfFsTarget = 0#
    mopFsTargetDisp = mgvfFsTarget
    
    If mopWsMode > modeManual Then
        mopWsMode = modeManual
    ElseIf mopWsMode = modeOff Then
        GoTo Done
    End If
    
    Do Until mopFsSetpt = 0#
        mopFsSetpt = Ctl_Ramp(mgvfFsTarget, mopFsSetpt, mccfMaxFsRamp, mccfUpdateInt, mgvfFsIncrement)
        mobjParManager.WriteOutputs
        If SecondsSinceTime(vStart) >= mccfMaxSuspendTime Then
            Exit Do
        End If
        Pause mccfUpdateInt
    Loop
    
    mopMfControl = 0#
    mobjParManager.WriteOutputs
    
    Do Until mipMfStatus = 0#
        mobjParManager.ReadInputs
        InputConversions
        If SecondsSinceTime(vStart) >= mccfMaxSuspendTime Then
            Exit Do
        End If
        Pause mccfUpdateInt
    Loop
    
    mopAuxControl = 0#
    mobjParManager.WriteOutputs
    
    GoTo Done
Error:
    Resume Done
Done:
    On Error Resume Next
    Exit Sub
End Sub

Public Sub Shutdown()
'
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim objSOM As SOM
    
    On Error GoTo Error
    bRaiseError = False
                
    SetProcessPriority priNormal
    
' Shut down actions
' ------------------
    
' Clear any pending commands
    mbReady = False
    tmrMain.Enabled = False
    mobjCommands.Clear "Shutdown"
    DoEvents
        
    mbInited = False
    mobjParManager.ReleaseOutputs
    
    mobjEvents.RaiseQuit
        
    If Not mobjForm Is Nothing Then
        Unload mobjForm
        Set mobjForm = Nothing
    End If
    
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
    EH.Procedure = "Shutdown"
    Call EH_MsgBox
    Resume Next
Done:
    On Error Resume Next
    Set objSOM = Nothing
    If bRaiseError Then
        On Error GoTo 0
        Call EH_RaiseError(OLEERR_METHOD_FAILED, TypeName(Me), "Shutdown", sRaiseErrorText)
    End If
    Exit Sub
End Sub

Public Property Get ControlConstants() As ControlLib.ControlConstants

' Returns the collection of control constants for use by an
' external debugger (Look)

' Returns
' --------------------------------------------------------
' The return value is a ControlConstants object

    Const PROC_NAME = "ControlConstants"
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

Public Property Get GlobalVars() As ControlLib.GlobalVars

' Returns the collection of global variables for use by an
' external debugger (Look)

' Returns
' --------------------------------------------------------
' The return value is a GlobalVars object

    Const PROC_NAME = "GlobalVars"
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

Public Property Get Ready() As Boolean

' Returns
' --------------------------------------------------------
' True
' False
 
    On Error Resume Next
    Ready = mbReady
    Exit Property
End Property

Private Sub Class_Initialize()

    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim sClassName As String
    
    On Error GoTo Error
    bRaiseError = False
    
    On Error Resume Next
    'LogMessage "Windspeed Controller Class Initialize. Thread ID = " & App.ThreadID
    
'Stop this instance if already running
    If App.PrevInstance Then
        If App.StartMode <> vbSModeStandalone Then
            EH.AppError = StdTalentStrings.stsEHInvalidActiveXLaunch
            EH.Item(1) = PROCESS_NAME
            EH.Module = TypeName(Me)
            EH.Procedure = "Class_Initialize"
            Call EH_MsgBox
        End If
        End
    End If
    
    On Error GoTo Error
    App.StartLogging "", vbLogAuto
    App.OleServerBusyRaiseError = True
    App.OleServerBusyTimeout = 1000
    
    Process.Title = PROCESS_NAME
    Process.EHCaption = PROCESS_NAME
'    If Not IsWithinIDE() Then
        Process.EHMode = ehLogOnly
'    Else
'        Process.EHMode = ehDisplayAndLog
'    End If
    Process.EHMsgsToLog = ehCriticalExclamation
    Process.LogMode = logToTalentAndFile
    
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
        
    mgvlOldNozzle = -1
    mopBlockageCorr = 1#

    
    GoTo Done
Error:
    sRaiseErrorText = Err.Description
    bRaiseError = True
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = "Class_Initialize"
        Call EH_MsgBox
    End If
    Resume Done
Error_CreatingObjects:
    sRaiseErrorText = LoadResString(ridCantCreateObject, sClassName) & Err.Description
'    sRaiseErrorText = "Can't create object '" & sClassName & "': " & Err.Description
    bRaiseError = True
    EH.Module = TypeName(Me)
    EH.Procedure = "Class_Initialize"
    EH.Item(1) = sClassName
    Call EH_MsgBox
    Resume Done
Done:
    On Error Resume Next
'    If bRaiseError Then
'        On Error GoTo 0
'        Call EH_RaiseError(OLEERR_METHOD_FAILED, TypeName(Me), "Class_Initialize", sRaiseErrorText)
'    End If
    Exit Sub
End Sub

Private Sub Class_Terminate()
'
    On Error Resume Next
    tmrMain.Enabled = False
    Set tmrMain.Notify = Nothing
    Set tmrMain = Nothing
    mobjInterlocks.Quit
    Set mobjInterlocks = Nothing
    Set tmrElapsed = Nothing
    mobjStatuses.Quit
    Set mobjStatuses = Nothing
    mobjCommands.Clear "Terminate"
    Set mobjCommands = Nothing
    Set mobjConstants = Nothing
    Set mobjGlobalVars = Nothing
    Set mobjEvents = Nothing
    Set mobjStats = Nothing
    Set mcolInvalidOpcodes = Nothing
    mobjMFPermComponents.Release
    Set mobjMFPermComponents = Nothing
    mobjMFAuxPermComponents.Release
    Set mobjMFAuxPermComponents = Nothing
    mobjIdleCmdPermComponents.Release
    Set mobjIdleCmdPermComponents = Nothing
    mobjParManager.Release
    Set mobjParManager = Nothing
    Set mobjSPGConstants = Nothing
    
    Set mobjFFLeadLag = Nothing
    Set mobjForm = Nothing
    Set mobjFSMonitor = Nothing
    Set mobjPID = Nothing
    Set mobjWSctlFilter = Nothing
    Set mobjWSFilter = Nothing
    Set mobjWSFilter2 = Nothing
    Set mobjWSMonitor = Nothing
    Set mobjTrackFilter = Nothing
    Set mobjRampRateFilter = Nothing
    Set mobjOutputFilter = Nothing
    Set mobjTrackPar = Nothing
    mbInited = False
    Exit Sub
End Sub

Private Sub ICcrpTimerNotify_Timer(ByVal Milliseconds As Long)
'
    Dim lElapsed As Long
    
    On Error Resume Next
    ControlLogic
    If mbKeepStats And mbInited Then
        lElapsed = tmrElapsed.Elapsed
        tmrElapsed.Reset
        mobjStats.AddValue CDbl(lElapsed / 1000)
    End If
    DoEvents
    Exit Sub
End Sub

Private Sub mobjParManager_BeforeUnloaded(ConfigurationName As String, Cancel As Boolean, Cancellers As Collection)
'
    Const PROC_NAME = "mobjParManager_BeforeUnloaded"

    On Error GoTo Error
    If mbInited = True Then
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
    LogMessage Message, mtStatusChange

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
    On Error GoTo Error
    LogMessage MessageText, MessageType

    GoTo Done
Error:
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = "mobjCommands_MessageFlushed"
        Call EH_MsgBox
    End If
    Resume Done
Done:
    On Error Resume Next
    Exit Sub
End Sub

Private Sub mobjCommands_CommandCompleted(Command As CommandListLib.Command, CompletionStatus As CommandListLib.CompletionStatuses, CompletionMessage As String)
'
    Const PROC_NAME = "mobjCommands_CommandCompleted"
    Dim objEH As EH
            
    On Error GoTo Error
    Select Case CompletionStatus
    Case cmdAborted, cmdTimedout
        If Command.SuppressLogging = False Then
            Set objEH = New EH
            objEH.AppError = AppErrorText(tceCommandFailed, Command.Text, CompletionMessage)
            objEH.Module = TypeName(Me)
            objEH.Procedure = PROC_NAME
            Set objEH.Component = App
            Command.SuppressLogging = True
            Select Case Command.Opcode
            Case ocStartMainFanAux
                LogPopupMessage objEH, Command.Text
            Case Else
                LogMessage2 objEH
            End Select
        End If
    End Select

    Select Case Command.Opcode
    Case ocStartIdle, ocStopIdle
        If CompletionStatus <> cmdCompleted Then
            mopIdleStartControl = 0#
            mopIdleStopControl = 0#
        End If
    Case ocResetFaults
        mopMFReset = 0#
    End Select
    
    GoTo Done
Error:
    Resume Done
Done:
    On Error Resume Next
    Set objEH = Nothing
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

Public Property Get Events() As Events

'Returns the Events object

    Set Events = mobjEvents

End Property

'Public Property Get Interlocks() As Interlocks
'    Set Interlocks = mobjInterlocks
'End Property

Public Function InvalidOpcodes(Opcode) As Variant
    
    Dim sIndex As String
    
    On Error Resume Next
    sIndex = Opcode
    InvalidOpcodes = mcolInvalidOpcodes(sIndex)
    If Err Then
        InvalidOpcodes = 0
    End If
    Exit Function
End Function

Public Sub CheckInterlock(Opcode)

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
    If IsArray(Opcode) Then
        sIndex = Opcode(LBound(Opcode))
        mvCheckIlckValue = Opcode(LBound(Opcode) + 1)
    Else
        sIndex = Opcode
        mvCheckIlckValue = Null
    End If
    
    Set objInterlock = mobjInterlocks(sIndex)
    If Err Then
        GoTo Done
    End If
    
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
        EH_RaiseError lRaiseErrorNumber, TypeName(Me), PROC_NAME, sRaiseErrorText
    End If
    Exit Sub
End Sub

Public Property Get Stats() As Stats
    Set Stats = mobjStats
End Property

Public Property Get Commands() As CommandListLib.Commands

'Returns the command list

    Const PROC_NAME = "Commands"
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

Private Sub WSControlLogic()

'Closed loop control logic

    Const PROC_NAME = "WSControlLogic"
    Dim I As Long
    Dim fSpeedTgt As Single
    Dim fSpeedRate As Single
    Dim fIncrement As Single
    
    On Error GoTo Error
'Mode dependent Control algorithm


' Calculate the internal mode to use for control. If in track mode and controlling to the Dyno Speed, and the Dyno or Throttle
' is in speed mode, then convert to "auto" control and control to the dyno target using the dyno rate.

    mgvbUsingCloseTime = False
    mgvlWSControlMode = mopWsMode
    mgvbWSComingBackToSetpt = False
    
    If mgvlWindTunnel = wtCWT Then
        If (mopWsMode = modeTrack) And mccbAutoFollowDyTarget Then
            If mgvsTrackParName = mccsDefTrackParName Then
                If (mipDyMode = DY_MODE_SPEED) Or (mipThrMode = THR_MODE_SPEED) Then
                    mgvlWSControlMode = modeAuto
                    fSpeedTgt = Ctl_Limit(mipDySpeedTarget + mgvfHeadwind, mccfMinWs, mgvfMaxWs)
                    fSpeedRate = Ctl_Limit(mipDySpeedRate, mccfMinWsRamp, mccfMaxWsRamp)
                    If (mgvfWsTarget <> fSpeedTgt) Or (mgvfWsRamp <> fSpeedRate) Then
                        mgvfWsTarget = fSpeedTgt
                        mgvfWsTarget = Ctl_Limit(mgvfWsTarget, mccfMinWs, mgvfMaxWs)
                        mgvfWsRamp = fSpeedRate
                        If mgvfWsRamp <= 0 Then
                            mgvfWsRamp = mccfDefWsRamp
                        End If
                        If mcclSetpointRampMethod = 1 Then
                            mobjRampRateFilter.ClearData
                            mgvfControlRamp = mobjRampRateFilter.FilteredValue(mgvfWsRamp)
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    Select Case mgvlWSControlMode
    Case modeAuto, modeTrack
           
        ' If tracking, get target value
        If mgvlWSControlMode = modeTrack Then
            On Error Resume Next
            mgvfWsTarget = ((mobjTrackPar.Value * mgvfTrackC1) + mgvfTrackC0) + mgvfHeadwind
            mgvfWsTarget = Ctl_Limit(mgvfWsTarget, mccfMinWs, mgvfMaxWs)
            With mobjTrackFilter
                .FilterConst(0) = mccfWsTgtTrackFilTc
                mgvfWsTarget = .FilteredValue(mgvfWsTarget)
            End With
            If Err Then
                LogMessage LoadResString(ridTrackFailed, Err.Description), mtFailure
                SetMode modeAuto
            End If
            On Error GoTo Error
            mgvfControlRamp = mgvfWsRamp
        Else    '= modeAuto
            mobjWSMonitor.Target = mgvfWsTarget
            If (mopWsMode <> mgvfOldCtlMode) Then
                mobjWSMonitor.Reset
            End If
           
            Select Case mcclSetpointRampMethod
            Case 1                  ' Use Close Time
    'Use a slow ramp rate when near the target
                If mgvlNozzle = nsLarge Then
                    mgvfCloseTime = TableLookup(mopWsSetpt, mccfCloseTimeSpeedsLarge, mccfCloseTimesLarge, mcclNumCloseTimes(mgvlNozzle))
                    mgvfCloseTimeRamp = mccfCloseWsRamp(nsSmall)
                ElseIf mgvlNozzle = nsMedium Then
                    mgvfCloseTime = TableLookup(mopWsSetpt, mccfCloseTimeSpeedsMedium, mccfCloseTimesMedium, mcclNumCloseTimes(mgvlNozzle))
                    mgvfCloseTimeRamp = mccfCloseWsRamp(nsMedium)
                Else
                    mgvfCloseTime = TableLookup(mopWsSetpt, mccfCloseTimeSpeedsSmall, mccfCloseTimesSmall, mcclNumCloseTimes(mgvlNozzle))
                    mgvfCloseTimeRamp = mccfCloseWsRamp(nsSmall)
                End If
            
                If ((mgvfWsTarget > 0#) And (mgvfWsRamp > mgvfCloseTimeRamp) And ((Abs(mgvfWsTarget - mopWsSetpt) / mgvfWsRamp) <= mgvfCloseTime)) Then
                    With mobjRampRateFilter
                        .FilterConst(0) = mgvfCloseWsFilTc
                        mgvfControlRamp = .FilteredValue(mgvfCloseTimeRamp)
                    End With
                    mgvbUsingCloseTime = True
                Else
                    mobjRampRateFilter.ClearData
                    mgvfControlRamp = mobjRampRateFilter.FilteredValue(mgvfWsRamp)
                End If
            Case 2      ' Limit the ramp rate when the target error gets below some value, and then snap the setpoint to the final value once the WS
                        ' starts coming back to the setpoint
                If (mgvlWindTunnel = wtAAWT) And (mgvbHoldCRUntilFSOK = True) Then
                    If (mopWindspeed > mopWsSetpt) Then
                        mgvfControlRamp = mgvfWsRamp
                    End If
                Else
                    mgvfControlRamp = mgvfWsRamp
                End If
                
                If ((Abs(mgvfWsTarget - IIf(mccbFinalRampUseSetpoint, mopWsSetpt, mopWsControl)) <= mccfFinalRampDeltakph) And (mgvfWsRamp > mccfWsRateForFinalRamp) And mgvbFinalRampOKToUse) Then
                    mgvbUsingFinalRamp = True
                    If mgvlNozzle = nsSmall Then
                        mgvfControlRamp = mccfFinalRampDeltakphSmall / mccfFinalRampTimeSecSmall
                    Else
                        mgvfControlRamp = mccfFinalRampDeltakph / mccfFinalRampTimeSec
                    End If
                    ' When close, once the WS starts coming back to the setpoint, snap the setpoint to the target value
                    If (mopWsControl > mopWsSetpt) And (mobjPID.States(ctlDerivative) > 0#) And (mopWsSetpt < mgvfWsTarget) Then
                        mgvbWSComingBackToSetpt = True
                    ElseIf (mopWsControl < mopWsSetpt) And (mobjPID.States(ctlDerivative) < 0#) And (mopWsSetpt > mgvfWsTarget) Then
                        mgvbWSComingBackToSetpt = True
                    End If
                End If
            
            Case Else
                mgvfControlRamp = mgvfWsRamp
            End Select
        
        End If
        
        
        If (mopWsMode = modeTrack) And mgvbAutoIdle Then
            If mgvfWsTarget >= (mccfWSStartIdle + mccfAutoIdleDeadband) Then         ' If speed > Start + DB then Stop the Idle in all cases
                mgvbAutoIdleWantsIdle = False
                mgvbStopAutoIdle = False
                mgvbAutoIdleSpeedHasBeenHigh = True
                mgvbAutoIdleSpeedHasBeenLow = False
                
            ElseIf mgvfWsTarget >= mccfWSStartIdle Then                              ' If speed > Start then no change to idle state
                
            ElseIf mgvfWsTarget >= mccfWSStopIdle Then                              ' If speed < Start and > Stop then start idle if speed is going down and stop idle
                                                                                    ' if speed is going up
                If mgvbStopAutoIdle Then
                    mgvbAutoIdleWantsIdle = False
                ElseIf mgvbAutoIdleSpeedHasBeenHigh Then                            ' Speed has been high, so start the idle
                    mgvbAutoIdleWantsIdle = True
                    mgvbAutoIdleSpeedHasBeenLow = False
                    
                ElseIf mgvbAutoIdleSpeedHasBeenLow Then                             ' Speed has been low, so stop the idle
                    mgvbAutoIdleWantsIdle = False
                    mgvbStopAutoIdle = False
                    mgvbAutoIdleSpeedHasBeenHigh = False
                Else
                        ' no change
                End If
                
            ElseIf mgvfWsTarget >= IIf((mccfWSStopIdle - mccfAutoIdleDeadband) > 0.1, (mccfWSStopIdle - mccfAutoIdleDeadband), 0.1) Then ' Speed is < Stop but > Stop - DB. No change
            
            Else                                                                    ' Else speed < Stop - DB then Start the Idle in all cases
                If mgvbStopAutoIdle Then
                    mgvbAutoIdleWantsIdle = False
                Else
                    mgvbAutoIdleWantsIdle = True
                End If
                mgvbAutoIdleSpeedHasBeenLow = True
                mgvbAutoIdleSpeedHasBeenHigh = False
                
            End If
            
            mgvfOldWsTarget = mgvfWsTarget
            If mgvbAutoIdleWantsIdle Then
                If mopIdleStatus = isIdle Then
                    mgvfFsTarget = mccfIdleFanspeed
                    mgvfFsRamp = mccfMaxFsRamp
                Else
                    mgvfFsTarget = 0#
                    mgvfFsRamp = mccfMaxFsRamp
                End If
                GoTo ManualMode
            ElseIf (mopIdleStatus <> isNormal) Then
                mgvfWsTarget = Ctl_Limit(mgvfWsTarget, 0, mccfMaxFsLimit(1) * mgvfNomWsMax / mgvfNomFsMax)
            End If
            
        ElseIf mgvbAutoIdle = False Then
            mgvbAutoIdleSpeedHasBeenHigh = False
            mgvbAutoIdleSpeedHasBeenLow = False
            
        End If
        
        mgvfWSTargetLimited = mgvfWsTarget
        For I = 1 To 3
            If LimitInEffect(I) Then
                If mgvfWSTargetLimited > (mccfMaxFsLimit(I) * mgvfNomWsMax / mgvfNomFsMax) Then
                    mgvfWSTargetLimited = mccfMaxFsLimit(I) * mgvfNomWsMax / mgvfNomFsMax
                    mobjRampRateFilter.ClearData
                    mgvfControlRamp = mobjRampRateFilter.FilteredValue(mgvfWsRamp)
                End If
            End If
        Next I
        
        If mgvlWindTunnel = wtAAWT Then                         ' Stall protection in the AAWT
            If mgvfWSTargetLimited > mopWsSetpt Then
                If mipFanspeed <= mccfSlowRampFanspeed Then
                    If mgvfControlRamp > mccfMaxWsRampLowSpeeds Then
                        mgvfControlRamp = mccfMaxWsRampLowSpeeds
                        mgvbFinalRampOKToUse = False
                        mgvbHoldCRUntilFSOK = True
                    End If
                End If
            End If
        End If
        
        mopWsSetpt = Ctl_Ramp(mgvfWSTargetLimited, mopWsSetpt, mgvfControlRamp, mccfUpdateInt, mgvfWsIncrement)
        If mccbFinalRampSnap And mgvbUsingFinalRamp And mgvbWSComingBackToSetpt And IIf(Abs(mgvfWsTarget - mopWsControl) <= mccfUseTargetErrorForDerivLimit, True, False) And (mopWsMode = modeAuto) Then
            If mgvbFinalRampHasSnapped = False Then
                If mccbFinalRampSnapBumpless Then
                    ' Subtract from the current integral state the difference in KP effect that the snapping of the setpoint will cause
                    mobjPID.States(ctlIntegral) = mobjPID.States(ctlIntegral) - mgvfKp * ((mgvfWSTargetLimited - mopWsSetptDisp) / mgvfNomWsMax)
                End If
            End If
            mopWsSetpt = mgvfWSTargetLimited
            mgvbFinalRampHasSnapped = True
        End If
        
        If (mgvbUsingCloseTime Or mgvbUsingFinalRamp Or mgvbFinalRampHasSnapped) Then
            mopWsSetptDisp = Ctl_Ramp(mgvfWSTargetLimited, mopWsSetptDisp, mgvfWsRamp, mccfUpdateInt, fIncrement)
        Else
            mopWsSetptDisp = mopWsSetpt
        End If
        If mopWsSetpt < 0 Then
            mgvfFsTarget = mopWsSetpt / mccfNomWSReverse * mccfNomFSReverse
            mgvfFsTarget = Ctl_Limit(mgvfFsTarget, mccfMinFs, 0)
            mgvfFsLimitedTarget = mgvfFsTarget
            For I = 1 To 3
                If LimitInEffect(I) Then
                    mgvfFsLimitedTarget = Ctl_Limit(mgvfFsLimitedTarget, -mccfMaxFsLimit(I), mccfMaxFsLimit(I))
                End If
            Next I
        Else
            mgvfNormWsSetpt = mopWsSetptDisp / mgvfNomWsMax   ' Normalize setpoint (use smooth ramping display setpt, not the potentially snapped actual setpoint)
            
    'Calculate Feedforward component. Include lead/lag compensation
            If (mopBlockageCorr <> 0#) Then
                mgvfFfRawNormFs = mccfFfGain * mgvfNormWsSetpt / Sqr(mopBlockageCorr)
            Else
                mgvfFfRawNormFs = mccfFfGain * mgvfNormWsSetpt
            End If
            If (mgvlWSControlMode = modeTrack) Then
                mgvfFfLeadCoeff = mccfFfLeadCoefTrack
                mgvfFfLagCoeff = mccfFfLagCoefTrack
            Else
                mgvfFfLeadCoeff = mccfFfLeadCoef
                mgvfFfLagCoeff = mccfFfLagCoef
            End If
            With mobjFFLeadLag
                .FilterConst(0) = mgvfFfLeadCoeff
                .FilterConst(1) = mgvfFfLagCoeff
                .UpdateTime = mccfUpdateInt
                mgvfFfNormFs = .FilteredValue(mgvfFfRawNormFs)
            End With
            
    'reset PID states if entering PID control. Set the integral to give bumpless transfer of fanspeed
            If (mgvfOldCtlMode < modeAuto) Or (mgvlOldIdleStatus <> isNormal) Then
                mgvfFfNormFs = mgvfFfRawNormFs
                With mobjPID
                    'PID states
                    .States(ctlProportional) = 0#
                    .States(ctlIntegral) = (mopFsSetpt / mgvfNomFsMax) - mgvfFfNormFs
                    .States(ctlDerivative) = 0#
                End With
            End If
            
    'Calculate adaptive parameters, and then PID feedback component
    ' If not within FB range or low speed is on, then disable the integrator
            If IsWSTooLowForPID Or (mopIdleStatus <> isNormal) Then
                mobjPID.States(ctlIntegral) = 0
                If (mgvlWSControlMode = modeTrack) And mgvbAutoIdle And mgvbAutoIdleWantsIdle Then
                    '
                ElseIf (mgvlWSControlMode <> modeTrack) Then
                    mgvfFbNormFs = 0#
                ElseIf ((mgvfWsTarget > mopWsControl) And (mgvfFbNormFs < 0#)) Then
                    mgvfFbNormFs = 0#
                ElseIf (mgvfFbNormFs < -mgvfFfNormFs) Then
                    mgvfFbNormFs = -mgvfFfNormFs
                End If
            Else                      ' ELSE perform feedback control
                mgvfWsError = mopWsSetpt - mopWsControl
                mgvfWsError = Ctl_Limit(mgvfWsError, -mccfMaxWsError, mccfMaxWsError)
                mgvfNormError = mgvfWsError / mgvfNomWsMax
                
                mgvfWSTargetError = mgvfWSTargetLimited - mopWsControl
                mgvfTargetNormError = mgvfWSTargetError / mgvfNomWsMax
                
                mgvfKp = mgvfKp0 + mgvfKp1 * mgvfNormWsSetpt  'setup adaptive coeff
                mgvfKp = Ctl_Limit(mgvfKp, mgvfMinKp, mgvfMaxKp)
                
                mgvfTi = mgvfTi0 + mgvfTi1 * mgvfNormWsSetpt
                mgvfTi = Ctl_Limit(mgvfTi, mgvfMinTi, mgvfMaxTi)
                
                mgvfTd = mgvfTd0 + mgvfTd1 * mgvfNormWsSetpt
                mgvfTd = Ctl_Limit(mgvfTd, mgvfMinTd, mgvfMaxTd)
        
                mgvfOldIntegral = mobjPID.States(ctlIntegral)
                With mobjPID
                    .Alpha = mccfTdFilterTC
                    .ControlMax = mccfCtlMax
                    .ControlMin = mccfCtlMin
                    .IntegralMax = mccfIntMax
                    .IntegralMin = mccfIntMin
                    .Kp = mgvfKp
                    .Td = mgvfTd
                    .Ti = mgvfTi
                    .UpdateTime = mccfUpdateInt
                    If mopWsMode > modeAuto Then
                        mgvfFbNormFs = .Output(mgvfNormError)
                    Else
                        mgvfFbNormFs = .Output2(mgvfNormError, mgvfTargetNormError, IIf(Abs(mgvfWSTargetError) <= mccfUseTargetErrorForDerivLimit, True, False))
                    End If
                End With
            
    'Lock intergral during high ramping
                mgvfWsRateForI = mccfWsRateForI0 + mccfWsRateForI1 * mgvfNormWsSetpt
                mgvfWsRateForI = Ctl_Limit(mgvfWsRateForI, mccfMinWsRateForI, mccfMaxWsRateForI)
                If ((Abs(mgvfWsIncrement / mccfUpdateInt) > mgvfWsRateForI)) Then
                    mobjPID.States(ctlIntegral) = mgvfOldIntegral
    '            ElseIf ((mipMfCurrent >= mccfMaxCurrent) Or (mipMfCurrent <= mccfMinCurrent)) Then
    '                mobjPID.States(ctlIntegral) = mgvfOldIntegral
                'ElseIf ((mopWsControl <= mccfMinWsForITrack) And (mgvlWSControlMode = modeTrack)) Then
                '    mobjPID.States(ctlIntegral) = mgvfOldIntegral
                End If
            End If
           
    'Calculate FS setpoint
            mgvfFsNormSetpt = mgvfFfNormFs + mgvfFbNormFs
            mgvfFsTarget = mgvfNomFsMax * mgvfFsNormSetpt     ' Convert fan speed
            mgvfFsTarget = Ctl_Limit(mgvfFsTarget, mccfMinFsAutoMode, mgvfMaxFs)
            mgvfFsLimitedTarget = mgvfFsTarget
            For I = 1 To 3
                If LimitInEffect(I) Then
                    mgvfFsLimitedTarget = Ctl_Limit(mgvfFsLimitedTarget, mccfMinFsAutoMode, mccfMaxFsLimit(I))
                End If
            Next I
        End If
        
        If mopWsMode = modeTrack Then
            mopFsSetpt = Ctl_Ramp(mgvfFsLimitedTarget, mopFsSetpt, mccfMaxFsRampTrack, mccfUpdateInt, mgvfFsIncrement)
        Else
            mopFsSetpt = Ctl_Ramp(mgvfFsLimitedTarget, mopFsSetpt, mccfMaxFsRamp, mccfUpdateInt, mgvfFsIncrement)
        End If
        
        If mopWsSetpt >= 0 Then
            If mgvfFsLimitedTarget < mgvfFsTarget Then
                If mgvfNormError > 0 Then
                    mobjPID.States(ctlIntegral) = mgvfOldIntegral
                End If
            ElseIf mgvfFsLimitedTarget > mgvfFsTarget Then
                If mgvfNormError < 0 Then
                    mobjPID.States(ctlIntegral) = mgvfOldIntegral
                End If
            End If
        End If
        
        If (mccbFilterOutput) Then
            With mobjOutputFilter
                .FilterConst(0) = mccfPidOutputFilTc
                mopFsSetpt = .FilteredValue(mopFsSetpt)
            End With
        End If
        If mopWsSetpt >= 0 Then
            mopFsSetpt = Ctl_Limit(mopFsSetpt, mccfMinFsAutoMode, mgvfMaxFs)
        Else
            mopFsSetpt = Ctl_Limit(mopFsSetpt, mccfMinFs, mgvfMaxFs)
        End If

    Case modeManual
        mobjFSMonitor.Target = mgvfFsTarget
        If (mopWsMode <> mgvfOldCtlMode) Then
            mobjFSMonitor.Reset
        End If
ManualMode:
        mgvfNormWsSetpt = 0.5         '  (Norm WS setpt for DeltaP filtering use only)
        If (mipMfStatus = 0#) Or (mopMfControl = 0#) Then
            mgvfFsTarget = 0#
            mopFsSetpt = mgvfFsTarget
            mgvfWsTarget = 0#
            mopWsSetpt = mgvfWsTarget
        Else
            mgvfWsTarget = mopWindspeed
            mopWsSetpt = mgvfWsTarget
            mgvfFsTarget = Ctl_Limit(mgvfFsTarget, mccfMinFs, mgvfMaxFs)
            mgvfFsLimitedTarget = mgvfFsTarget
            For I = 1 To 3
                If LimitInEffect(I) Then
                    mgvfFsLimitedTarget = Ctl_Limit(mgvfFsLimitedTarget, mccfMinFs, mccfMaxFsLimit(I))
                End If
            Next I
            mopFsSetpt = Ctl_Ramp(mgvfFsTarget, mopFsSetpt, mgvfFsRamp, mccfUpdateInt, mgvfFsIncrement)
        End If
        
        ' Shutting down the Fan drive (but not Aux) if MainFanSpeed is 0 for long time.
        If mgvfFsTarget = 0# And mopFsSetpt = 0# Then
            If mccfTurnOffFanTime > 0 Then
                If SecondsSinceTime(mgvdTurnOffFanTime) > mccfTurnOffFanTime Then    '30mins
                    If mopMfControl = 1# Then
                        If mccfMainFanAuxAutoOffTime > 0 Then
                            mgvbAuxAutoOff = True
                        End If
                    End If
                    mopMfControl = 0#
                End If
            End If
        Else
            mgvdTurnOffFanTime = Now()
        End If
            
        ' Shutting down the main fan auxiliaries if the fan has been off for a long time
        If (mipMfStatus = 0#) And (mopMfControl = 0#) And (mopAuxControl = 1#) Then
            If (mccfMainFanAuxAutoOffTime > 0) And mgvbAuxAutoOff And (mgvdDriveOffTime <> 0) Then
                If SecondsSinceTime(mgvdDriveOffTime) >= mccfMainFanAuxAutoOffTime * 60# Then
                    mopAuxControl = 0#
                End If
            End If
        End If
            
    Case Else       ' Off, set outputs to zero
        mgvfNormWsSetpt = 0.5         '  (for DeltaP filtering only)
        mopMfControl = 0#
        mgvfWsTarget = 0#
        mopWsSetpt = 0#
        mgvfFsTarget = 0#
        mopFsSetpt = 0#
    End Select
    
    With mobjPID
        mgvfKpEffect = .States(ctlProportional)
        mgvfDeriviativeEffect = .States(ctlDerivative)
        mgvfIntegratorEffect = .States(ctlIntegral)
    End With
                
    GoTo Done
Error:
    EH.AppError = LoadResString(stsEHErrorExecutingControlLogic)
    EH.Module = TypeName(Me)
    EH.Procedure = PROC_NAME
    Call EH_MsgBox
    Resume ErrorExit
ErrorExit:
    On Error Resume Next
    Suspend
    GoTo Done
Done:
    On Error Resume Next
    Exit Sub
End Sub
Private Function LimitInEffect(LimitNumber As Long) As Boolean

    LimitInEffect = False
    
    If mgvlWindTunnel = wtCWT Then
        Select Case LimitNumber
        Case 1
            If mopIdleStatus <> IdleStatuses.isNormal Then
                LimitInEffect = True
            End If
        End Select
    End If
    
End Function
Private Function IsWSTooLowForPID() As Boolean

' Return True if the setpoint is less than min PID speed (and the min fanspeed durign Auto mode is >= 0)
' Return True if the setpoint is less than min PID speed and has been that way for a specific time (if the min fanspeed durign Auto mode is < 0)
    
    If (mopWsSetpt <= mccfMinWsForPid) Or _
            ((mopWsControl <= mccfMinWsForPid) And (mgvlWSControlMode = modeTrack)) Or _
            ((mopWsControl <= mccfMinWsForPid) And (mopWsSetpt < mopWsControl)) Then
        If mgvbWsTooLowForPID = False Then
            mgvdWsTooLowForPIDTime = Now
        End If
        mgvbWsTooLowForPID = True
    Else
        mgvbWsTooLowForPID = False
    End If
    
    If mgvbWsTooLowForPID Then
        If mccfMinFsAutoMode >= 0 Then
            IsWSTooLowForPID = True
        ElseIf SecondsSinceTime(mgvdWsTooLowForPIDTime) >= mccfMinWsForPidDelay Then
            IsWSTooLowForPID = True
        Else
            IsWSTooLowForPID = False
        End If
    Else
        IsWSTooLowForPID = False
    End If

End Function

Private Sub SetMode(vNewValue As Modes)

'Sets/returns the control mode

' Part          Description
' ----------------------------------------------------------------------
'
' vNewValue     New value that property value is being set to

    Const PROC_NAME = "SetMode"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String

    On Error GoTo Error
    bRaiseError = False

    If mopWsMode = vNewValue Then
        GoTo Done
    End If
    
    Select Case vNewValue
    Case modeTrack, modeAuto
        If mopWsMode < modeAuto Then
            mobjFFLeadLag.ClearData
            mobjPID.States(ctlIntegral) = 0#
        End If
    End Select
    
    Select Case vNewValue
    Case modeTrack
        mopWsMode = modeTrack
        mgvfWsRamp = mccfMaxWsRamp
        mgvfControlRamp = mgvfWsRamp
        
    Case modeAuto
        Set mobjTrackPar = Nothing
        mopWsMode = modeAuto
        If (Abs((mccfFfGain * mopWindspeed / mgvfNomWsMax) - (mipFanspeed / mgvfNomFsMax)) <= mccfIntMax) Then
            mgvfWsTarget = mopWindspeed             '   Bumpless transfer
        Else
            mgvfWsTarget = mgvfNomWsMax * (mopFsSetpt / mgvfNomFsMax) / mccfFfGain
        End If
        mgvfWsRamp = mccfDefWsRamp
        mopWsSetpt = mgvfWsTarget
        mgvfFsRamp = mccfMaxFsRamp
        mgvfControlRamp = mgvfWsRamp
        
    Case modeManual
        Set mobjTrackPar = Nothing
        mopWsMode = modeManual
        If mipMfStatus = 1# Then
            mgvfFsTarget = mipFanspeed
            mopFsSetpt = mipFanspeed
            mopMfControl = 1#
        Else
            mgvfFsTarget = 0#
            mopFsSetpt = 0#
            mopMfControl = 0#
        End If
        mgvfFsRamp = mccfDefFsRamp
        
    Case Else
        Set mobjTrackPar = Nothing
        mopWsMode = modeOff
        mgvfFsTarget = 0#
        mopFsSetpt = 0#
        mgvfWsTarget = 0#
        mopWsSetpt = 0#
        mopMfControl = 0
    End Select
    
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
    Exit Sub
End Sub

@


1.24
log
@AAWT Purge Update@
text
@d26 1
a26 1
Private WithEvents tmrMain As ccrpTimer
d3509 1
a3509 1
    If Not mobjForm Is Nothing Then
d3863 1
a3863 1
    Set tmrMain = New ccrpTimers6.ccrpTimer
@


1.23
log
@@
text
@d2589 10
a2598 5
                If mipCoolPLCCmdRequest = 0 Then
                    If .FirstTime Then
                        RequestControlFromCoolPLC
                    ElseIf .TimeSince >= 10# Then
                        .Abort LoadResString(ridUnableToGetControlFromPLC, "Cooling"), cmdTimedout
d2600 1
d2602 10
a2611 2
                    mopPurgeCmd = 1
                    .Remove
d2616 10
a2625 5
                If mipCoolPLCCmdRequest = 0 Then
                    If .FirstTime Then
                        RequestControlFromCoolPLC
                    ElseIf .TimeSince >= 10# Then
                        .Abort LoadResString(ridUnableToGetControlFromPLC, "Cooling"), cmdTimedout
d2627 1
d2629 10
a2638 2
                    mopPurgeCmd = 0
                    .Remove
@


1.22
log
@Idle improvements@
text
@d346 1
d492 1
d607 10
d767 1
d1860 2
d2076 1
a2076 1
            
d2456 1
d4268 8
a4275 1
                mgvfControlRamp = mgvfWsRamp
d4372 2
@


1.21
log
@Added basic collector flaps commands@
text
@d1410 1
a1410 1
        ElseIf (mopIdleStartControl = 1#) Or (mopIdleStopControl = 1#) Then
d1412 2
a1413 2
        Else
            mopIdleStatus = isFault
d2287 1
a2287 1
                If mopIdleStatus <> isNormal Then
d2477 2
d2974 1
a2974 1
        For I = 1 To 3
@


1.20
log
@Updated to keep the MF Start command on during an estop (while the fan fast ramps to 0 rpm).@
text
@d234 14
d291 16
d355 11
d492 17
d651 65
d871 58
d1684 40
d1745 1
d1982 17
d2040 12
d2630 122
d2822 2
a2823 1
    Case CStr(ocStartMainFan), CStr(ocStopMainFan), CStr(ocStartMainFanAux), CStr(ocStopMainFanAux), CStr(ocStartIdle), CStr(ocStopIdle), CStr(ocShowStatus), CStr(ocSetNozzleSize)
d3205 14
d3304 10
a3313 1
                
@


1.19
log
@@
text
@d449 1
d894 6
a899 4
            If mgvfFsTarget = 0# Then
                mopMfControl = 0#
            ElseIf mipMfStatus = 1# Then
                mopMfControl = 1#
d909 2
a1482 3
        With .Add(CStr(ocSetNozzleSize))
            .Expressions.Add "MFOn", mipMfStatus, 0, ilckVariableEqualValueMeansOK, LoadResString(ridFanIsOnStopFanFirst)
        End With
d1500 1
@


1.18
log
@Fixed Idle handling (included bypass flaps). Fixed speed limit handling with Prep and Stab operation modes. Fixed Purge commands.@
text
@d548 28
a575 24
    mopNozzleCmd(nsSmall) = 0#
    mopNozzleCmd(nsMedium) = 0#
    mopNozzleCmd(nsLarge) = 0#
    If (mipMfStatus = 0) And (SecondsSinceTime(mgvdNozzleSizeCmdTime) <= mccfNozzleTimeout) Then
        Select Case mgvlNozzleTarget
        Case nsSmall
            If mipNozzleSmallLimSw = 1 Then
                mgvlNozzleTarget = 0
            Else
                mopNozzleCmd(nsSmall) = 1#
            End If
        Case nsMedium
            If mipNozzleMediumLimSw = 1 Then
                mgvlNozzleTarget = 0
            Else
                mopNozzleCmd(nsMedium) = 1#
            End If
        Case nsLarge
            If mipNozzleLargeLimSw = 1 Then
                mgvlNozzleTarget = 0
            Else
                mopNozzleCmd(nsLarge) = 1#
            End If
        End Select
d1489 1
d1497 3
d2269 1
a2269 1
                                If (mipMfStatus = 0#) Or (mgvfFsTarget < mccfIdleFanspeed) Then
@


1.17
log
@@
text
@d270 6
a275 4
Private mipIdleMotorStop As Single
Private mipIdleStoppedStatus As Single
Private mipIdleStartedStatus As Single
Private mipIdleFault As Single
d767 1
d771 1
a771 1
            
d805 20
d827 14
d842 27
d1214 1
a1214 1
        If mipIdleStoppedStatus = 1# Then
d1216 1
a1216 1
        ElseIf mipIdleStartedStatus = 1# Then
d1218 4
d1223 1
a1223 1
            mopIdleStatus = isMoving
d1483 2
a1484 2
                .Expressions.Add "Fault", mipIdleFault, 1#, ilckVariableEqualValueMeansOK, LoadResString(ridIdleDriveFault)
                .Expressions.Add "IdleStop", mipIdleMotorStop, 0#, ilckVariableEqualValueMeansOK, LoadResString(ridIdleMotorStopOn)
d1736 8
a1743 5
            .AddWithoutFlag "Idle_MotorStop", mipIdleMotorStop, ioInput
            .AddWithoutFlag "Idle_FlapOpen", mipIdleStoppedStatus, ioInput
            .AddWithoutFlag "Idle_FlapClosed", mipIdleStartedStatus, ioInput
            .AddWithoutFlag "Idle_MotorFault", mipIdleFault, ioInput
            
d1789 1
d1798 1
d2254 1
d2299 1
d2486 17
d2523 1
a2523 1
        If mipIdleMotorStop = 1# Then
d4193 1
a4193 1
            If mipIdleStoppedStatus <> 1 Then
a4197 6
    ' ToDo: Update as necessary based upon key switch, etc. Maybe also idle system in CWT
''    If mipMFSpeedLimitCommand(LimitNumber) = 0# Then
''        LimitInEffect = True
''    ElseIf mipMFSpeedLimitStatus(LimitNumber) = 1# Then
''        LimitInEffect = True
''    End If
@


1.16
log
@Continued CWT work. Idle etc.@
text
@d1710 3
a1712 3
'            .AddWithoutFlag "NozzleSize", mopNozzle, ioOutput
'            .AddWithoutFlag "AutoIdleStatus", mopAutoIdleStatus, ioOutput
'            .AddWithoutFlag "Idle_Status", mopIdleStatus, ioOutput
@


1.15
log
@@
text
@d270 3
a272 4
Private mipIdleCmdPerm As Single
Private mipIdleOpenStatus As Single
Private mipIdleClosedStatus As Single
Private mipIdleMotionStatus As Single
a273 1
Private mipIdleLight As Single
d288 2
a289 2
Private mopIdleOpenControl As Single
Private mopIdleCloseControl As Single
d535 2
a536 2
                mopIdleCloseControl = 1#
                mopIdleOpenControl = 0#
d539 2
a540 2
            mopIdleOpenControl = 1#
            mopIdleCloseControl = 0#
d572 14
d863 2
a864 2
            mopIdleCloseControl = 0#
            mopIdleOpenControl = 0#
d981 1
a981 1
        mgvdVmps = LimitD(PolyDouble(CDbl(mipFanspeed), mgvdWSvsFSCoeffs, 6), 0#, mgvfMaxWs / 3.6)
d986 3
a988 3
    mgvdKp = LimitD(PolyDouble(mgvdVmps, mgvdKpCoeffs, 5), mgvfKpCoeffMin, mgvfKpCoeffMax)
    mgvdKqn = LimitD(PolyDouble(mgvdVmps, mgvdKqnCoeffs, 5), mgvfKqnCoeffMin, mgvfKqnCoeffMax)
    mgvdKqpl = LimitD(PolyDouble(mgvdVmps, mgvdKqplCoeffs, 5), mgvfKqplCoeffMin, mgvfKqplCoeffMax)
d1138 1
d1150 7
a1156 9
''        If mipIdleOpenStatus = 1# Then
''            mopIdleStatus = isNormal
''        ElseIf mipIdleClosedStatus = 1# Then
''            mopIdleStatus = isIdle
''        ElseIf mipIdleMotionStatus = 1# Then
''            mopIdleStatus = isMoving
''        Else
''            mopIdleStatus = isStopped
''        End If
a1157 2
        mopIdleStatus = isNormal
        
a1414 1
                ''.Expressions.Add "Light", mipIdleLight, 1#, ilckVariableEqualValueMeansOK, LoadResString(ridLightCurtain)
d1416 1
a1416 1
                .Expressions.Add "IdlePerm", mipIdleCmdPerm, 1#, ilckVariableEqualValueMeansOK, LoadResString(ridIdlePermissiveNotOn)
d1668 4
a1671 6
    ''        .AddWithoutFlag "IdlePermissive", mipIdleCmdPerm, ioInput
    ''        .AddWithoutFlag "IdleOpenStatus", mipIdleOpenStatus, ioInput
    ''        .AddWithoutFlag "IdleClosedStatus", mipIdleClosedStatus, ioInput
    ''        .AddWithoutFlag "IdleInMotion", mipIdleMotionStatus, ioInput
    ''        .AddWithoutFlag "IdleLightFault", mipIdleLight, ioInput
    ''        .AddWithoutFlag "IdleDriveFault", mipIdleFault, ioInput
d1710 5
a1714 6
    ''        .AddWithoutFlag "NozzleSize", mopNozzle, ioOutput
            
    ''        .AddWithoutFlag "AutoIdleStatus", mopAutoIdleStatus, ioOutput
    ''        .AddWithoutFlag "Idle_Status", mopIdleStatus, ioOutput
            .AddWithoutFlag "Idle_StartCmd", mopIdleOpenControl, ioOutput
            .AddWithoutFlag "Idle_StopCmd", mopIdleCloseControl, ioOutput
d2154 2
a2155 2
                mopIdleCloseControl = 0#
                mopIdleOpenControl = 0#
d2164 7
a2170 1
                If mopWsMode > modeManual Then
d2178 1
a2178 1
                ElseIf (mopIdleOpenControl = 1#) And (mipIdleMotionStatus = 1#) Then
d2181 2
a2182 2
                    mopIdleCloseControl = 1#
                    mopIdleOpenControl = 0#
d2201 2
a2202 2
                    mopIdleCloseControl = 0#
                    mopIdleOpenControl = 1#
d2207 6
d2222 1
a2222 1
                ElseIf (mopIdleCloseControl = 1#) And (mipIdleMotionStatus = 1#) Then
d2225 2
a2226 2
                    mopIdleCloseControl = 0#
                    mopIdleOpenControl = 1#
d2431 1
a2431 1
        If mipIdleCmdPerm = 0# Then
d3075 2
a3076 2
    mopIdleCloseControl = 0#
    mopIdleOpenControl = 0#
d3506 1
a3506 1
    Case ocStartIdle
d3508 2
a3509 2
            mopIdleCloseControl = 0#
            mopIdleOpenControl = 0#
d4098 8
@


1.14
log
@@
text
@d83 6
a88 4
Private mccfCloseTimesNormal() As Single
Private mccfCloseTimeSpeedsNormal() As Single
Private mccfCloseTimesSmall() As Single
Private mccfCloseTimeSpeedsSmall() As Single
d1447 1
a1447 1
        .AddLong "Nozzle Size (1=Normal, 2=Small) (mgvfNozzle)", mgvlNozzle
d2555 19
a2573 6
                ReDim mccfCloseTimesNormal(1 To mcclNumCloseTimes(J))
                ReDim mccfCloseTimeSpeedsNormal(1 To mcclNumCloseTimes(J))
                For I = 1 To mcclNumCloseTimes(J)
                    .AddSingle "CloseTime" & NozzleName(J) & Format(I, "00") & "Time", mccfCloseTimesNormal(I), bStatus, sText
                    .AddSingle "CloseTime" & NozzleName(J) & Format(I, "00") & "Speed", mccfCloseTimeSpeedsNormal(I), bStatus, sText
                Next I
d3710 2
a3711 2
                If mgvlNozzle = nsSmall Then
                    mgvfCloseTime = TableLookup(mopWsSetpt, mccfCloseTimeSpeedsNormal, mccfCloseTimesNormal)
d3713 3
d3717 2
a3718 2
                    mgvfCloseTime = TableLookup(mopWsSetpt, mccfCloseTimeSpeedsNormal, mccfCloseTimesNormal)
                    mgvfCloseTimeRamp = mccfCloseWsRamp(nsLarge)
@


1.13
log
@Added PUrge commands. Added control of overflow flaps (AAWT). Changed Relhumidity par name (AAWT) to make it consistent with CWT@
text
@d230 1
d248 2
d326 1
d445 2
d546 26
d1395 3
d1432 3
d1632 2
d1654 3
a1656 3
  ''          .AddWithoutFlag "NozzleSmall", mipNozzleTopLimSw, ioInput
   '         .AddWithoutFlag "NozzleSmall", mipNozzleMiddleLimSw, ioInput
   '         .AddWithoutFlag "NozzleSmall", mipNozzleBottomLimSw, ioInput
d1708 3
d2215 10
a2224 2
                mopPurgeCmd = 1
                .Remove
d2228 10
a2237 2
                mopPurgeCmd = 0
                .Remove
d2247 27
d2344 1
a2344 1
    Case CStr(ocStartMainFan), CStr(ocStopMainFan), CStr(ocStartMainFanAux), CStr(ocStopMainFanAux), CStr(ocStartIdle), CStr(ocStopIdle), CStr(ocShowStatus)
d2695 1
@


1.12
log
@@
text
@d228 3
d321 3
d490 2
d551 10
d1346 1
d1360 3
d1582 1
a1582 1
        .AddWithoutFlag "RelHumid", mipRelHumid, ioInput
d1676 4
d2170 10
d2651 7
@


1.11
log
@Fixed WS calculation plenum method. Added addition m/s output parameters. Added Kp and Kq coefficients output parameters@
text
@d164 2
d2493 2
d3676 10
@


1.10
log
@Fixed to actually use RH when calculating the windspeed@
text
@d296 3
d307 6
d553 6
d900 1
a900 1
            mopFilDeltaPPlenum = mobjWSFilter.FilteredValue(mipDeltaPPlenum)
a907 17
'Filter separate control mopWindspeed if required.

    If (mgvbSeparateControl) Then
        mgvfControlFilTc = mgvfControlFilTc0 + mgvfControlFilTc1 * mgvfNormWsSetpt
        mgvfControlFilTc = Ctl_Limit(mgvfControlFilTc, mgvfControlFilTcMin, mgvfControlFilTcMax)
        If mgvfControlFilTc > mccfUpdateInt Then
            With mobjWSctlFilter
                .FilterConst(0) = mgvfControlFilTc
                mgvfFilDeltaPControl = .FilteredValue(mipDeltaP)
            End With
        Else
            mgvfFilDeltaPControl = mipDeltaP
        End If
    Else
        mgvfFilDeltaPControl = mopFilDeltaP
    End If

d920 3
a922 3
    mgvdKp = LimitD(PolyDouble(mgvdVmps, mgvdKpCoeffs, 6), mgvfKpCoeffMin, mgvfKpCoeffMax)
    mgvdKqn = LimitD(PolyDouble(mgvdVmps, mgvdKqnCoeffs, 6), mgvfKqnCoeffMin, mgvfKqnCoeffMax)
    mgvdKqpl = LimitD(PolyDouble(mgvdVmps, mgvdKqplCoeffs, 6), mgvfKqplCoeffMin, mgvfKqplCoeffMax)
d956 24
a979 13
''        mgvfPDiffCorrControl = mgvfCDynamic_A0 + (mgvfCDynamic_A1 * mgvfFilDeltaPControl) + (mgvfCDynamic_A2 * mgvfFilDeltaPControl ^ 2)
''        If (mopAirDensityKgm3 > 0#) And (mgvfPDiffCorrControl >= 0#) Then       ' Compute mopWindspeed kph
''            mgvfMachControl = (5# * ((mgvfPDiffCorrControl / mgvfPstaticPa + 1#) ^ (2 / 7) - 1)) ^ (1 / 2)
''            mgvfQControlPa = (1.4 * mgvfPstaticPa * mgvfMachControl ^ 2 / 2#) * mopBlockageCorr
''            If mgvfQControlPa < 0# Then mgvfQControlPa = 0#
''            If (mopAirDensityKgm3 > 0#) And (mgvfQControlPa >= 0#) Then       ' Compute Windspeed kph
''                mopWsControl = 3.6 * Sqr((2# * mgvfQControlPa) / mopAirDensityKgm3)
''            Else
''                mopWsControl = 0#
''            End If
''        Else
            mopWsControl = mopWindspeed
''        End If
a993 1

d1004 4
d1396 2
d1402 10
d1617 3
d1646 8
@


1.9
log
@Used to tune windspeed control on the AAWT successfully@
text
@d917 1
a917 1
        mgvdVmps = LimitD(PolyDouble(CDbl(mipFanspeed), mgvdWSvsFSCoeffs, 6), 0#, mgvfMaxWs)
d1534 1
@


1.8
log
@@
text
@a302 1
Private mopWSToPLC As Single
a542 1
    mopWSToPLC = mopWindspeed / 3.6
d848 8
d862 2
a863 1
                mobjWSFilterBW.InitializeLowPass 1 / mccfUpdateInt, (1# / 2# * 3.14159 * mgvfWsFilTc0), 4
a1599 1
        .AddWithoutFlag "PLCWindspeed", mopWSToPLC, ioOutput
d1923 5
a1927 1
                        If SecondsSinceTime(mgvdDriveOnTime) >= mccfDrivePause Then
@


1.7
log
@Completed for initial testing at CWT. Added Nozzle/Plenum commands@
text
@d80 1
a81 1

d109 2
d171 2
d237 6
d255 1
a255 1
Private mipAuxPerm As Single
d310 1
d639 5
a644 3
        If mgvdDriveOffTime = 0 Then
            mgvdDriveOffTime = Now
        End If
d649 14
a662 11
    Select Case mopIdleStatus
    Case isIdle
        mobjStatuses.SetValue STS_IDLE_STARTED, True
        mobjStatuses.SetValue STS_IDLE_STOPPED, False
    Case isNormal
        mobjStatuses.SetValue STS_IDLE_STARTED, False
        mobjStatuses.SetValue STS_IDLE_STOPPED, True
    Case Else
        mobjStatuses.SetValue STS_IDLE_STARTED, False
        mobjStatuses.SetValue STS_IDLE_STOPPED, False
    End Select
a663 1
    
d779 7
a785 3
    If Not mobjInterlocks(CStr(ocStartIdle)).IsOK Then
        If (mopWsMode = modeTrack) And mgvbAutoIdle Then
            mgvbAutoIdle = False
a786 2
        mopIdleCloseControl = 0#
        mopIdleOpenControl = 0#
d788 1
a788 1

d1039 3
a1041 1
    mipPabs = mipPabs
d1098 1
a1098 1
    If mccbAutoFollowDyTarget And (mccbAutoFollowDyTargetUseTrackCoeffs = False) And (mopWsMode = modeTrack) And (mgvsTrackParName = mccsDefTrackParName) And ((mipDyMode = DY_MODE_SPEED) Or (mipThrMode = THR_MODE_SPEED)) Then
d1272 2
d1289 1
a1289 1
            .Expressions.Add "MFAuxPerm", mipAuxPerm, 1#, ilckVariableEqualValueMeansOK, LoadResString(ridAuxPermFault)
d1324 14
a1337 13
        With .Add(CStr(ocStartIdle))
            ''.Expressions.Add "Light", mipIdleLight, 1#, ilckVariableEqualValueMeansOK, LoadResString(ridLightCurtain)
            .Expressions.Add "Fault", mipIdleFault, 1#, ilckVariableEqualValueMeansOK, LoadResString(ridIdleDriveFault)
            .Expressions.Add "IdlePerm", mipIdleCmdPerm, 1#, ilckVariableEqualValueMeansOK, LoadResString(ridIdlePermissiveNotOn)
        End With
        With .Add(CStr(ocStopIdle))
            .IncludedInterlocks.Add CStr(ocStartIdle)
        End With
        With .Add(CStr(ocShowIdleStatus))
            .IncludedInterlocks.Add CStr(ocStartIdle)
        End With
                
    
d1351 4
a1354 2
        .Add Array(ocSetWindSpeed, ocTrackWindSpeed, ocStopIdle, ocStartIdle), CStr(ocStartIdle)
        .Add Array(ocStartIdle, ocStopIdle), CStr(ocStopIdle)
d1437 7
a1443 5
        .AddSingle "Idle Status (0=Stopped, 1=Started, 2=Moving)", mopIdleStatus
        .AddBoolean "Auto Idle Active", mgvbAutoIdle
        .AddBoolean "Auto Idle Speed has been high", mgvbAutoIdleSpeedHasBeenHigh
        .AddBoolean "Auto Idle Speed has been Low", mgvbAutoIdleSpeedHasBeenLow
        .AddBoolean "Auto Idle Stop Idle Overrride Active", mgvbStopAutoIdle
d1494 4
a1497 2
        .Add STS_IDLE_STARTED, False, "", LoadResString(ridIdleStarted)
        .Add STS_IDLE_STOPPED, True, "", LoadResString(ridIdleStopped)
d1531 8
d1553 1
a1553 1
        .AddWithoutFlag "MF_CSReady", mipAuxPerm, ioInput
d1998 5
a2002 1
                    mgvbAutoIdle = .Arguments(3).Value
a2004 2
                    mgvdLastIdleTime = Now
                    mgvbStopAutoIdle = False
d2174 51
d2232 6
d2239 1
d2251 1
a2251 1
        If mipAuxPerm = 0# Then
d2371 1
d2380 1
a2380 1
            For J = 1 To 1 ' IIf(WindTunnel = wtCWT, 3, 1)
d2442 5
a2446 1
            
d2458 4
d2485 22
a2506 65
            ' Large (Normal)
            .AddSingle "PIDGainLargeBase", mccfKp0(nsLarge), bStatus, sText
            .AddSingle "PIDGainLargeAdjust", mccfKp1(nsLarge), bStatus, sText
            .AddSingle "PIDGainLargeTrackBase", mccfKp0Track(nsLarge), bStatus, sText
            .AddSingle "PIDGainLargeTrackAdjust", mccfKp1Track(nsLarge), bStatus, sText
            .AddSingle "PIDGainLargeMin", mccfKpMin(nsLarge), bStatus, sText
            .AddSingle "PIDGainLargeMax", mccfKpMax(nsLarge), bStatus, sText
            
            .AddSingle "PIDDerivLargeBase", mccfTd0(nsLarge), bStatus, sText
            .AddSingle "PIDDerivLargeAdjust", mccfTd1(nsLarge), bStatus, sText
            .AddSingle "PIDDerivLargeTrackBase", mccfTd0Track(nsLarge), bStatus, sText
            .AddSingle "PIDDerivLargeTrackAdjust", mccfTd1Track(nsLarge), bStatus, sText
            .AddSingle "PIDDerivLargeMin", mccfTdMin(nsLarge), bStatus, sText
            .AddSingle "PIDDerivLargeMax", mccfTdMax(nsLarge), bStatus, sText
            
            .AddSingle "PIDIntegralLargeBase", mccfTi0(nsLarge), bStatus, sText
            .AddSingle "PIDIntegralLargeAdjust", mccfTi1(nsLarge), bStatus, sText
            .AddSingle "PIDIntegralLargeTrackBase", mccfTi0Track(nsLarge), bStatus, sText
            .AddSingle "PIDIntegralLargeTrackAdjust", mccfTi1Track(nsLarge), bStatus, sText
            .AddSingle "PIDIntegralLargeMin", mccfTiMin(nsLarge), bStatus, sText
            .AddSingle "PIDIntegralLargeMax", mccfTiMax(nsLarge), bStatus, sText
            
             ' Medium
            .AddSingle "PIDGainMediumBase", mccfKp0(nsMedium), bStatus, sText
            .AddSingle "PIDGainMediumAdjust", mccfKp1(nsMedium), bStatus, sText
            .AddSingle "PIDGainMediumTrackBase", mccfKp0Track(nsMedium), bStatus, sText
            .AddSingle "PIDGainMediumTrackAdjust", mccfKp1Track(nsMedium), bStatus, sText
            .AddSingle "PIDGainMediumMin", mccfKpMin(nsMedium), bStatus, sText
            .AddSingle "PIDGainMediumMax", mccfKpMax(nsMedium), bStatus, sText
            
            .AddSingle "PIDDerivMediumBase", mccfTd0(nsMedium), bStatus, sText
            .AddSingle "PIDDerivMediumAdjust", mccfTd1(nsMedium), bStatus, sText
            .AddSingle "PIDDerivMediumTrackBase", mccfTd0Track(nsMedium), bStatus, sText
            .AddSingle "PIDDerivMediumTrackAdjust", mccfTd1Track(nsMedium), bStatus, sText
            .AddSingle "PIDDerivMediumMin", mccfTdMin(nsMedium), bStatus, sText
            .AddSingle "PIDDerivMediumMax", mccfTdMax(nsMedium), bStatus, sText
            
            .AddSingle "PIDIntegralMediumBase", mccfTi0(nsMedium), bStatus, sText
            .AddSingle "PIDIntegralMediumAdjust", mccfTi1(nsMedium), bStatus, sText
            .AddSingle "PIDIntegralMediumTrackBase", mccfTi0Track(nsMedium), bStatus, sText
            .AddSingle "PIDIntegralMediumTrackAdjust", mccfTi1Track(nsMedium), bStatus, sText
            .AddSingle "PIDIntegralMediumMin", mccfTiMin(nsMedium), bStatus, sText
            .AddSingle "PIDIntegralMediumMax", mccfTiMax(nsMedium), bStatus, sText
           
            ' Small
            .AddSingle "PIDGainSmallBase", mccfKp0(nsSmall), bStatus, sText
            .AddSingle "PIDGainSmallAdjust", mccfKp1(nsSmall), bStatus, sText
            .AddSingle "PIDGainSmallTrackingBase", mccfKp0Track(nsSmall), bStatus, sText
            .AddSingle "PIDGainSmallTrackingAdjust", mccfKp1Track(nsSmall), bStatus, sText
            .AddSingle "PIDGainSmallMin", mccfKpMin(nsSmall), bStatus, sText
            .AddSingle "PIDGainSmallMax", mccfKpMax(nsSmall), bStatus, sText
            
            .AddSingle "PIDDerivSmallBase", mccfTd0(nsSmall), bStatus, sText
            .AddSingle "PIDDerivSmallAdjust", mccfTd1(nsSmall), bStatus, sText
            .AddSingle "PIDDerivSmallTrackBase", mccfTd0Track(nsSmall), bStatus, sText
            .AddSingle "PIDDerivSmallTrackAdjust", mccfTd1Track(nsSmall), bStatus, sText
            .AddSingle "PIDDerivSmallMin", mccfTdMin(nsSmall), bStatus, sText
            .AddSingle "PIDDerivSmallMax", mccfTdMax(nsSmall), bStatus, sText
            
            .AddSingle "PIDIntegralSmallBase", mccfTi0(nsSmall), bStatus, sText
            .AddSingle "PIDIntegralSmallAdjust", mccfTi1(nsSmall), bStatus, sText
            .AddSingle "PIDIntegralSmallTrackBase", mccfTi0Track(nsSmall), bStatus, sText
            .AddSingle "PIDIntegralSmallTrackAdjust", mccfTi1Track(nsSmall), bStatus, sText
            .AddSingle "PIDIntegralSmallMin", mccfTiMin(nsSmall), bStatus, sText
            .AddSingle "PIDIntegralSmallMax", mccfTiMax(nsSmall), bStatus, sText
d2515 9
a2523 7
            .AddSingle "IdleFanSpeed", mccfIdleFanspeed, bStatus, sText
            .AddBoolean "AutoIdleDefault", mccbDefAutoIdle, bStatus, sText
            .AddSingle "AutoIdleStartWS", mccfWSStartIdle, bStatus, sText
            .AddSingle "AutoIdleStopWS", mccfWSStopIdle, bStatus, sText
            .AddSingle "AutoIdleDeadband", mccfAutoIdleDeadband, bStatus, sText
            .AddString "IdlePermComponents", mccsIdleCmdPermComponents, bStatus, sText
            .AddSingle "IdleReversePauseTime", mccfCancelIdleTime, bStatus, sText
d2544 5
a2548 2
            .AddBoolean "TrackDyTarget", mccbAutoFollowDyTarget, bStatus, sText
            .AddBoolean "TrackDyTargetUseTrackCoeffs", mccbAutoFollowDyTargetUseTrackCoeffs, bStatus, sText
a2549 1
            .AddBoolean "AutoIdleUsesLimit1", mccbAutoIdleUsesLimitCommand, bStatus, sText
d3119 1
d3464 18
a3481 12
    If (mopWsMode = modeTrack) And mccbAutoFollowDyTarget Then
        If mgvsTrackParName = mccsDefTrackParName Then
            If (mipDyMode = DY_MODE_SPEED) Or (mipThrMode = THR_MODE_SPEED) Then
                mgvlWSControlMode = modeAuto
                fSpeedTgt = Ctl_Limit(mipDySpeedTarget + mgvfHeadwind, mccfMinWs, mgvfMaxWs)
                fSpeedRate = Ctl_Limit(mipDySpeedRate, mccfMinWsRamp, mccfMaxWsRamp)
                If (mgvfWsTarget <> fSpeedTgt) Or (mgvfWsRamp <> fSpeedRate) Then
                    mgvfWsTarget = fSpeedTgt
                    mgvfWsTarget = Ctl_Limit(mgvfWsTarget, mccfMinWs, mgvfMaxWs)
                    mgvfWsRamp = fSpeedRate
                    If mgvfWsRamp <= 0 Then
                        mgvfWsRamp = mccfDefWsRamp
a3482 4
                    If mcclSetpointRampMethod = 1 Then
                        mobjRampRateFilter.ClearData
                        mgvfControlRamp = mobjRampRateFilter.FilteredValue(mgvfWsRamp)
                    End If
d3487 1
a3487 1
            
d3830 1
a3830 1
        ' Shutting down the glycol cooling pump (GP8) if the fan has been off for a long time
@


1.6
log
@Initial version for WS tuning at CWT@
text
@d293 1
d533 2
d1045 12
a1056 9
        If mipIdleOpenStatus = 1# Then
            mopIdleStatus = isNormal
        ElseIf mipIdleClosedStatus = 1# Then
            mopIdleStatus = isIdle
        ElseIf mipIdleMotionStatus = 1# Then
            mopIdleStatus = isMoving
        Else
            mopIdleStatus = isStopped
        End If
d1304 1
a1304 1
            .Expressions.Add "Light", mipIdleLight, 1#, ilckVariableEqualValueMeansOK, LoadResString(ridLightCurtain)
d1548 1
d1557 1
d2057 9
a2065 1
                                  
d2739 3
a2741 3
        mccdKpCoeffs(Nozzle, I) = vValue(1, lCol)
        mccdKqnCoeffs(Nozzle, I) = vValue(2, lCol)
        mccdKqplCoeffs(Nozzle, I) = vValue(3, lCol)
d2744 6
a2749 6
    mccfKpCoeffMin(Nozzle) = vValue(1, 8)
    mccfKpCoeffMax(Nozzle) = vValue(1, 9)
    mccfKqnCoeffMin(Nozzle) = vValue(2, 8)
    mccfKqnCoeffMax(Nozzle) = vValue(2, 9)
    mccfKqplCoeffMin(Nozzle) = vValue(3, 8)
    mccfKqplCoeffMax(Nozzle) = vValue(3, 9)
@


1.5
log
@Temporary WS calc method based on rpm. Used for opening ceremony initial operation before delta P sensors available@
text
@d35 1
d37 1
a80 2
Private mccfMinBlockage(1 To 2) As Single
Private mccfMaxBlockage(1 To 2) As Single
d82 1
a82 14
Private mccfCStatic_A0(1 To 2) As Single
Private mccfCStatic_A1(1 To 2) As Single
Private mccfCStatic_A2(1 To 2) As Single
Private mccfCDynamic_A0(1 To 2) As Single
Private mccfCDynamic_A1(1 To 2) As Single
Private mccfCDynamic_A2(1 To 2) As Single
Private mccfCStatic_A0_BL(1 To 2) As Single
Private mccfCStatic_A1_BL(1 To 2) As Single
Private mccfCStatic_A2_BL(1 To 2) As Single
Private mccfCDynamic_A0_BL(1 To 2) As Single
Private mccfCDynamic_A1_BL(1 To 2) As Single
Private mccfCDynamic_A2_BL(1 To 2) As Single

Private mcclNumCloseTimes(1 To 2) As Long
d89 2
a90 2
Private mccfCloseWsRamp(1 To 2) As Single
Private mccfCloseWsFilTc(1 To 2) As Single
d101 2
a102 2
Private mccfWsFilterTc0(1 To 2) As Single
Private mccfWsFilterTc1(1 To 2) As Single
d104 2
a105 2
Private mccfWsFilterFast(1 To 2) As Single
Private mccfWsFilterFastTrack(1 To 2) As Single
d107 2
a108 2
Private mccfNomFs(1 To 2) As Single
Private mccfMaxFs(1 To 2) As Single
d117 6
a122 6
Private mccfKp0(1 To 2) As Single
Private mccfKp1(1 To 2) As Single
Private mccfKp0Track(1 To 2) As Single
Private mccfKp1Track(1 To 2) As Single
Private mccfKpMin(1 To 2) As Single
Private mccfKpMax(1 To 2) As Single
d124 6
a129 6
Private mccfTi0(1 To 2) As Single
Private mccfTi1(1 To 2) As Single
Private mccfTi0Track(1 To 2) As Single
Private mccfTi1Track(1 To 2) As Single
Private mccfTiMin(1 To 2) As Single
Private mccfTiMax(1 To 2) As Single
d131 6
a136 6
Private mccfTd0(1 To 2) As Single
Private mccfTd1(1 To 2) As Single
Private mccfTd0Track(1 To 2) As Single
Private mccfTd1Track(1 To 2) As Single
Private mccfTdMin(1 To 2) As Single
Private mccfTdMax(1 To 2) As Single
d163 2
a164 2
Private mccfNomWs(1 To 2) As Single
Private mccfMaxWs(1 To 2) As Single
d176 4
a179 4
Private mccfControlFilTc0(1 To 2) As Single
Private mccfControlFilTc1(1 To 2) As Single
Private mccfControlFilTcMin(1 To 2) As Single
Private mccfControlFilTcMax(1 To 2) As Single
d210 12
d224 2
d247 2
d266 1
d283 3
a310 12
Private mgvfCStatic_A0 As Single
Private mgvfCStatic_A1 As Single
Private mgvfCStatic_A2 As Single
Private mgvfCDynamic_A0 As Single
Private mgvfCDynamic_A1 As Single
Private mgvfCDynamic_A2 As Single
Private mgvfPstaticPa As Single
Private mgvfPDiffCorr As Single
Private mgvfMach As Single
Private mgvfPDiffCorrControl As Single
Private mgvfMachControl As Single

d381 1
d414 21
d531 1
d789 4
d839 1
d843 1
d846 1
d848 1
d852 1
d855 1
d858 1
d863 1
d883 4
a886 3
    mgvfPstaticPa = mipPabs * 1000# + mgvfCStatic_A0 + (mgvfCStatic_A1 * mopFilDeltaP) + (mgvfCStatic_A2 * mopFilDeltaP ^ 2)
    mopPstaticKpa = mgvfPstaticPa / 1000#
    mgvfPDiffCorr = mgvfCDynamic_A0 + (mgvfCDynamic_A1 * mopFilDeltaP) + (mgvfCDynamic_A2 * mopFilDeltaP ^ 2)
d888 18
d907 14
a920 7
    mopAirDensityKgm3 = Ctl_AirDensity(mipDewpoint, mipAirtemp, mopPstaticKpa)
'    mgvfPH2O = Ctl_VaporPressure(mipDewpoint)
'    mgvfR = mgvfPstaticPa / ((mgvfPH2O / 461.5) + ((mgvfPstaticPa - mgvfPH2O) / 287.058))     '
    If (mopAirDensityKgm3 > 0#) And (mgvfPDiffCorr >= 0#) Then       ' Compute mopWindspeed kph
        mgvfMach = (5# * ((mgvfPDiffCorr / mgvfPstaticPa + 1#) ^ (2 / 7) - 1)) ^ (1 / 2)
        mopQPa = (1.4 * mgvfPstaticPa * mgvfMach ^ 2 / 2#) * mopBlockageCorr
        If mopQPa < 0# Then mopQPa = 0#
a926 1
        mgvfMach = 0#
a929 10
'' Temporary Start
    If WindTunnel = wtCWT Then
        mopWindspeed = mipFanspeed * 0.333918
    Else
        mopWindspeed = mipFanspeed * 1.2943
    End If


'' Temporary End

d931 11
a941 11
        mgvfPDiffCorrControl = mgvfCDynamic_A0 + (mgvfCDynamic_A1 * mgvfFilDeltaPControl) + (mgvfCDynamic_A2 * mgvfFilDeltaPControl ^ 2)
        If (mopAirDensityKgm3 > 0#) And (mgvfPDiffCorrControl >= 0#) Then       ' Compute mopWindspeed kph
            mgvfMachControl = (5# * ((mgvfPDiffCorrControl / mgvfPstaticPa + 1#) ^ (2 / 7) - 1)) ^ (1 / 2)
            mgvfQControlPa = (1.4 * mgvfPstaticPa * mgvfMachControl ^ 2 / 2#) * mopBlockageCorr
            If mgvfQControlPa < 0# Then mgvfQControlPa = 0#
            If (mopAirDensityKgm3 > 0#) And (mgvfQControlPa >= 0#) Then       ' Compute Windspeed kph
                mopWsControl = 3.6 * Sqr((2# * mgvfQControlPa) / mopAirDensityKgm3)
            Else
                mopWsControl = 0#
            End If
        Else
d943 1
a943 1
        End If
d1019 1
a1019 1
    mipPabs = mipPabs / 10#             ' COnvert from hPa to kPa
d1029 2
a1030 2
    If mipNozzleSmallLimSw = 1 Then
        mgvlNozzle = nsSmall
d1032 19
a1050 1
        mgvlNozzle = nsNormal
d1052 1
a1052 11
   
    If mipIdleOpenStatus = 1# Then
        mopIdleStatus = isNormal
    ElseIf mipIdleClosedStatus = 1# Then
        mopIdleStatus = isIdle
    ElseIf mipIdleMotionStatus = 1# Then
        mopIdleStatus = isMoving
    Else
        mopIdleStatus = isStopped
    End If
        
d1059 1
d1066 1
a1066 1
    Case nsNormal, nsSmall
d1069 1
a1069 1
        I = nsNormal
a1119 31
        
    'Blockage Limitation
    mgvfMinBlockage = mccfMinBlockage(I)
    mgvfMaxBlockage = mccfMaxBlockage(I)
        


    Select Case mgvlNozzle
    Case nsNormal, nsSmall
        I = mgvlNozzle
    Case Else
        I = nsNormal
    End Select

    If mipBLFanStatus = 0 Then
        mgvfCStatic_A0 = mccfCStatic_A0(I)
        mgvfCStatic_A1 = mccfCStatic_A1(I)
        mgvfCStatic_A2 = mccfCStatic_A2(I)
        
        mgvfCDynamic_A0 = mccfCDynamic_A0(I)
        mgvfCDynamic_A1 = mccfCDynamic_A1(I)
        mgvfCDynamic_A2 = mccfCDynamic_A2(I)
    Else
        mgvfCStatic_A0 = mccfCStatic_A0_BL(I)
        mgvfCStatic_A1 = mccfCStatic_A1_BL(I)
        mgvfCStatic_A2 = mccfCStatic_A2_BL(I)
        
        mgvfCDynamic_A0 = mccfCDynamic_A0_BL(I)
        mgvfCDynamic_A1 = mccfCDynamic_A1_BL(I)
        mgvfCDynamic_A2 = mccfCDynamic_A2_BL(I)
    End If
a1124 2
            .CreatePropertyByPosition(CC_MinBlockage, bExists).Value = mgvfMinBlockage
            .CreatePropertyByPosition(CC_MaxBlockage, bExists).Value = mgvfMaxBlockage
d1128 15
d1208 1
a1324 1
        .Add Array(ocSetBlockage), CStr(ocSetBlockage)
a1350 6
        .AddSingle "CStatic_A0", mgvfCStatic_A0
        .AddSingle "CStatic_A1", mgvfCStatic_A1
        .AddSingle "CStatic_A2", mgvfCStatic_A2
        .AddSingle "CDynamic_A0", mgvfCDynamic_A0
        .AddSingle "CDynamic_A1", mgvfCDynamic_A1
        .AddSingle "CDynamic_A2", mgvfCDynamic_A2
a1351 7
        .AddSingle "PDiff (Corrected)", mgvfPDiffCorr
        .AddSingle "PDiff (Corrected) (Control)", mgvfPDiffCorrControl
'        .AddSingle "Rev WS c0", mgvfRevWsC(0)
'        .AddSingle "Rev WS c1", mgvfRevWsC(1)
'        .AddSingle "Rev WS c2", mgvfRevWsC(2)
        .AddSingle "Blockage correction factor", mopBlockageCorr
        
d2049 1
a2049 9
                
            Case ocSetBlockage
                CheckInterlock CStr(ocSetBlockage)
                mopBlockageCorr = .Arguments(1).Value
                If mopBlockageCorr <= 0# Then
                    mopBlockageCorr = 1#
                End If
                .Remove
                  
d2232 1
a2262 9

            ' Blockage
            ' ------------------------
            ' Std (Normal)
            .AddSingle "BlockageNormalMin", mccfMinBlockage(nsNormal), bStatus, sText
            .AddSingle "BlockageNormalMax", mccfMaxBlockage(nsNormal), bStatus, sText
            ' Small
            .AddSingle "BlockageSmallMin", mccfMinBlockage(nsSmall), bStatus, sText
            .AddSingle "BlockageSmallMax", mccfMaxBlockage(nsSmall), bStatus, sText
d2264 1
d2266 13
a2278 12
            ' ------------------------
            ' Std (Normal)
            mcclNumCloseTimes(nsNormal) = 1
            .AddLong "CloseTimeNormalNum", mcclNumCloseTimes(nsNormal), bStatus, sText
            ReDim mccfCloseTimesNormal(1 To mcclNumCloseTimes(nsNormal))
            ReDim mccfCloseTimeSpeedsNormal(1 To mcclNumCloseTimes(nsNormal))
            For I = 1 To mcclNumCloseTimes(nsNormal)
                .AddSingle "CloseTimeNormal" & Format(I, "00") & "Time", mccfCloseTimesNormal(I), bStatus, sText
                .AddSingle "CloseTimeNormal" & Format(I, "00") & "Speed", mccfCloseTimeSpeedsNormal(I), bStatus, sText
            Next I
            .AddSingle "CloseTimeNormalFilterTC", mccfCloseWsFilTc(nsNormal), bStatus, sText
            .AddSingle "CloseTimeNormalWSRate", mccfCloseWsRamp(nsNormal), bStatus, sText
a2279 12
            ' Small
            mcclNumCloseTimes(nsSmall) = 1
            .AddLong "CloseTimeSmallNum", mcclNumCloseTimes(nsSmall), bStatus, sText
            ReDim mccfCloseTimesSmall(1 To mcclNumCloseTimes(nsSmall))
            ReDim mccfCloseTimeSpeedsSmall(1 To mcclNumCloseTimes(nsSmall))
            For I = 1 To mcclNumCloseTimes(nsSmall)
                .AddSingle "CloseTimeSmall" & Format(I, "00") & "Time", mccfCloseTimesSmall(I), bStatus, sText
                .AddSingle "CloseTimeSmall" & Format(I, "00") & "Speed", mccfCloseTimeSpeedsSmall(I), bStatus, sText
            Next I
            .AddSingle "CloseTimeSmallFilterTC", mccfCloseWsFilTc(nsSmall), bStatus, sText
            .AddSingle "CloseTimeSmallWSRate", mccfCloseWsRamp(nsSmall), bStatus, sText
            
a2288 5
            ' Std (Normal)
            .AddSingle "DeltaPFilterNormalBase", mccfWsFilterTc0(nsNormal), bStatus, sText
            .AddSingle "DeltaPFilterNormalAdjust", mccfWsFilterTc1(nsNormal), bStatus, sText
            .AddSingle "DeltaPFilterNormalFast", mccfWsFilterFast(nsNormal), bStatus, sText
            .AddSingle "DeltaPFilterNormalFastTrack", mccfWsFilterFastTrack(nsNormal), bStatus, sText
d2290 11
a2300 19
            ' Small
            .AddSingle "DeltaPFilterSmallBase", mccfWsFilterTc0(nsSmall), bStatus, sText
            .AddSingle "DeltaPFilterSmallAdjust", mccfWsFilterTc1(nsSmall), bStatus, sText
            .AddSingle "DeltaPFilterSmallFast", mccfWsFilterFast(nsSmall), bStatus, sText
            .AddSingle "DeltaPFilterSmallFastTrack", mccfWsFilterFastTrack(nsSmall), bStatus, sText
            
            ' Filter Coefs of DeltP for Control use
            ' -------------------------------------
            ' Std (Normal)
            .AddSingle "DeltaPControlFilterNormalBase", mccfControlFilTc0(nsNormal), bStatus, sText
            .AddSingle "DeltaPControlFilterNormalAdjust", mccfControlFilTc1(nsNormal), bStatus, sText
            .AddSingle "DeltaPControlFilterNormalMin", mccfControlFilTcMin(nsNormal), bStatus, sText
            .AddSingle "DeltaPControlFilterNormalMax", mccfControlFilTcMax(nsNormal), bStatus, sText
            ' Small
            .AddSingle "DeltaPControlFilterSmallBase", mccfControlFilTc0(nsSmall), bStatus, sText
            .AddSingle "DeltaPControlFilterSmallAdjust", mccfControlFilTc1(nsSmall), bStatus, sText
            .AddSingle "DeltaPControlFilterSmallMin", mccfControlFilTcMin(nsSmall), bStatus, sText
            .AddSingle "DeltaPControlFilterSmallMax", mccfControlFilTcMax(nsSmall), bStatus, sText
            
d2307 4
a2310 4
            .AddSingle "FFNormalFanspeed", mccfNomFs(nsNormal), bStatus, sText
            .AddSingle "FFSmallFanspeed", mccfNomFs(nsSmall), bStatus, sText
            .AddSingle "FFNormalWindspeed", mccfNomWs(nsNormal), bStatus, sText
            .AddSingle "FFSmallWindspeed", mccfNomWs(nsSmall), bStatus, sText
d2315 7
a2321 2
            .AddSingle "FanspeedMaxNormal", mccfMaxFs(nsNormal), bStatus, sText
            .AddSingle "FanspeedMaxSmall", mccfMaxFs(nsSmall), bStatus, sText
d2331 7
a2337 2
            .AddSingle "WindspeedMaxNormal", mccfMaxWs(nsNormal), bStatus, sText
            .AddSingle "WindspeedMaxSmall", mccfMaxWs(nsSmall), bStatus, sText
d2364 7
a2370 7
            ' Std (Normal)
            .AddSingle "PIDGainNormalBase", mccfKp0(nsNormal), bStatus, sText
            .AddSingle "PIDGainNormalAdjust", mccfKp1(nsNormal), bStatus, sText
            .AddSingle "PIDGainNormalTrackBase", mccfKp0Track(nsNormal), bStatus, sText
            .AddSingle "PIDGainNormalTrackAdjust", mccfKp1Track(nsNormal), bStatus, sText
            .AddSingle "PIDGainNormalMin", mccfKpMin(nsNormal), bStatus, sText
            .AddSingle "PIDGainNormalMax", mccfKpMax(nsNormal), bStatus, sText
d2372 6
a2377 6
            .AddSingle "PIDDerivNormalBase", mccfTd0(nsNormal), bStatus, sText
            .AddSingle "PIDDerivNormalAdjust", mccfTd1(nsNormal), bStatus, sText
            .AddSingle "PIDDerivNormalTrackBase", mccfTd0Track(nsNormal), bStatus, sText
            .AddSingle "PIDDerivNormalTrackAdjust", mccfTd1Track(nsNormal), bStatus, sText
            .AddSingle "PIDDerivNormalMin", mccfTdMin(nsNormal), bStatus, sText
            .AddSingle "PIDDerivNormalMax", mccfTdMax(nsNormal), bStatus, sText
d2379 6
a2384 6
            .AddSingle "PIDIntegralNormalBase", mccfTi0(nsNormal), bStatus, sText
            .AddSingle "PIDIntegralNormalAdjust", mccfTi1(nsNormal), bStatus, sText
            .AddSingle "PIDIntegralNormalTrackBase", mccfTi0Track(nsNormal), bStatus, sText
            .AddSingle "PIDIntegralNormalTrackAdjust", mccfTi1Track(nsNormal), bStatus, sText
            .AddSingle "PIDIntegralNormalMin", mccfTiMin(nsNormal), bStatus, sText
            .AddSingle "PIDIntegralNormalMax", mccfTiMax(nsNormal), bStatus, sText
d2386 22
d2451 3
a2453 15
            ' Std (Normal) - Nozzle Method - No BL
            .AddSingle "CStaticNormC0", mccfCStatic_A0(nsNormal), bStatus, sText
            .AddSingle "CStaticNormC1", mccfCStatic_A1(nsNormal), bStatus, sText
            .AddSingle "CStaticNormC2", mccfCStatic_A2(nsNormal), bStatus, sText
            .AddSingle "CDynamicNormC0", mccfCDynamic_A0(nsNormal), bStatus, sText
            .AddSingle "CDynamicNormC1", mccfCDynamic_A1(nsNormal), bStatus, sText
            .AddSingle "CDynamicNormC2", mccfCDynamic_A2(nsNormal), bStatus, sText
                        
            ' Small - NO BL
            .AddSingle "CStaticSmallC0", mccfCStatic_A0(nsSmall), bStatus, sText
            .AddSingle "CStaticSmallC1", mccfCStatic_A1(nsSmall), bStatus, sText
            .AddSingle "CStaticSmallC2", mccfCStatic_A2(nsSmall), bStatus, sText
            .AddSingle "CDynamicSmallC0", mccfCDynamic_A0(nsSmall), bStatus, sText
            .AddSingle "CDynamicSmallC1", mccfCDynamic_A1(nsSmall), bStatus, sText
            .AddSingle "CDynamicSmallC2", mccfCDynamic_A2(nsSmall), bStatus, sText
a2454 16
            ' Std (Normal) - Nozzle Method - With BL
            .AddSingle "CStaticNorm_BL_C0", mccfCStatic_A0_BL(nsNormal), bStatus, sText
            .AddSingle "CStaticNorm_BL_C1", mccfCStatic_A1_BL(nsNormal), bStatus, sText
            .AddSingle "CStaticNorm_BL_C2", mccfCStatic_A2_BL(nsNormal), bStatus, sText
            .AddSingle "CDynamicNorm_BL_C0", mccfCDynamic_A0_BL(nsNormal), bStatus, sText
            .AddSingle "CDynamicNorm_BL_C1", mccfCDynamic_A1_BL(nsNormal), bStatus, sText
            .AddSingle "CDynamicNorm_BL_C2", mccfCDynamic_A2_BL(nsNormal), bStatus, sText
                        
            ' Small - With BL
            .AddSingle "CStaticSmall_BL_C0", mccfCStatic_A0_BL(nsSmall), bStatus, sText
            .AddSingle "CStaticSmall_BL_C1", mccfCStatic_A1_BL(nsSmall), bStatus, sText
            .AddSingle "CStaticSmall_BL_C2", mccfCStatic_A2_BL(nsSmall), bStatus, sText
            .AddSingle "CDynamicSmall_BL_C0", mccfCDynamic_A0_BL(nsSmall), bStatus, sText
            .AddSingle "CDynamicSmall_BL_C1", mccfCDynamic_A1_BL(nsSmall), bStatus, sText
            .AddSingle "CDynamicSmall_BL_C2", mccfCDynamic_A2_BL(nsSmall), bStatus, sText
                                                            
a2517 2
                .CreatePropertyByPosition(CC_MinBlockage, bExists).Value = mccfMinBlockage(nsNormal)
                .CreatePropertyByPosition(CC_MaxBlockage, bExists).Value = mccfMaxBlockage(nsNormal)
d2527 1
a2527 1
                mobjWSFilterBW.InitializeLowPass 1 / mccfUpdateInt, 1# / (2# * 3.14159 * mccfWsFilterTc0(nsNormal)), 4
d2530 1
a2530 1
                .FilterConst(0) = mccfWsFilterTc0(nsNormal) / mccfUpdateInt
d2533 1
a2533 1
                .FilterConst(0) = mccfWsFilterTc0(nsNormal)
d2537 13
d2552 1
a2552 1
            .FilterConst(0) = mccfControlFilTcMin(nsNormal)
d2593 6
d2664 1
d2683 15
d2699 66
d3100 1
d3435 1
a3435 1
                    mgvfCloseTimeRamp = mccfCloseWsRamp(nsNormal)
@


1.4
log
@Added ResetMainFanFaults command@
text
@d875 10
@


1.3
log
@Tested at CWT to start/stop fan and aux, and set fanspeed.@
text
@d76 2
d231 1
d272 1
d504 10
d1464 1
d1512 1
d1762 10
d2235 3
a2237 1
            
d3118 2
@


1.2
log
@Initial version for CWT@
text
@d952 1
a952 1
    If (mipCSStatus = 1#) And (mipLCReady = 1#) And (mipMCReady = 1#) Then
d954 1
a954 1
    ElseIf (mipCSStatus = 0#) And (mipLCReady = 0#) And (mipMCReady = 0#) Then
d1189 1
d1217 1
a1217 1
            .Expressions.Add "EStop", mipEstopStatus, 1#, ilckVariableEqualValueMeansOK, LoadResString(ridCommonInterlockFaulted)
d1766 6
a1771 6
                    If mipMCReady <> 1# Then
                        sText = IIf(sText <> "", sText & vbCrLf & "- ", "- ") & LoadResString(ridMCDidNotGetReady)
                    End If
                    If mipLCReady <> 1# Then
                        sText = IIf(sText <> "", sText & vbCrLf & "- ", "- ") & LoadResString(ridLCDidNotGetReady)
                    End If
d3658 8
a3665 5
    If mipMFSpeedLimitCommand(LimitNumber) = 0# Then
        LimitInEffect = True
    ElseIf mipMFSpeedLimitStatus(LimitNumber) = 1# Then
        LimitInEffect = True
    End If
@


1.1
log
@Initial revision@
text
@a228 1
Private mipReadyToRun As Single
d230 6
a235 3
Private mipMfReady As Single
Private mipMfFault As Single
Private mipMfDriveFault As Single
a239 1
Private mipAuxStatus As Single
a240 1
Private mipAuxReady As Single
a257 2
Private mopHxHighDpFault As Single
Private mopHxDpOverDpFault As Single
a264 1
Private mopIdleResetControl As Single
a268 1
Private mopMfReverseEnable As Single
d312 1
d457 1
a457 24
        
    Select Case mopWsMode
    Case modeAuto, modeTrack
        If mccfMinFsAutoMode < 0# Then
            mopMfReverseEnable = 1#
        ElseIf mopFsSetpt < 0 Then
            mopMfReverseEnable = 1#
        Else
            mopMfReverseEnable = 0#
        End If
    Case Else
        If mopMfControl = 1 Then
            If mopFsSetpt < 0 Then
                mopMfReverseEnable = 1#
            ElseIf mgvfFsTarget < 0# Then
                mopMfReverseEnable = 1#
            Else
                mopMfReverseEnable = 1#         ' Just leave it on, as fan faults happen if this is removed too soon
            End If
        ElseIf mopFsSetpt >= 0# Then
            mopMfReverseEnable = 0#
        End If
    End Select
    
d509 1
a509 1
    mgvfOldRemoteStatus = mipMfReady
d511 1
a511 1
    mgvfOldAuxStatus = mipAuxStatus
d569 1
a569 1
    If (mipAuxStatus = 1#) Then
d576 1
a576 1
    Else
d580 3
d656 5
a660 1
            mopAuxControl = mipAuxStatus
d669 1
a669 1
                mopAuxControl = mipAuxStatus
d677 1
a677 1
        If (mipAuxStatus = 0#) And (SecondsSinceTime(mgvdLastAuxOnTime) >= mccfMaxRunTimeWithoutCooling) Then
d704 1
a704 1
                mopAuxControl = mipAuxStatus
d718 1
a718 5
        If mopAuxControl = 1# Then
            mopAuxControl = mipAuxStatus
        Else
            mopAuxControl = 0#
        End If
d950 1
d952 8
d1195 1
a1195 1
            ''.Expressions.Add "AuxRemote", mipAuxReady, 1#, ilckVariableEqualValueMeansOK, LoadResString(ridAuxNotReady)
a1198 1
            .Expressions.Add "AuxRemote", mipAuxReady, 1#, ilckVariableEqualValueMeansOK, LoadResString(ridAuxNotReady)
d1200 1
a1200 1
            .Expressions.Add "MfAuxPerm", mipAuxPerm, 1#, ilckVariableEqualValueMeansOK, LoadResString(ridAuxPermFault)
d1204 1
a1204 1
            .Expressions.Add "FanRemote", mipMfReady, 1#, ilckVariableEqualValueMeansOK, LoadResString(ridFanNotinComputerCtl)  '"Main Fan Drive not in remote"
d1213 2
a1214 3
            .Expressions.Add "RTORun", mipReadyToRun, 1#, ilckVariableEqualValueMeansOK, LoadResString(ridReadyToRunFaulted)
            .Expressions.Add "MFRemote", mipMfReady, 1#, ilckVariableEqualValueMeansOK, LoadResString(ridFanNotinComputerCtl)
            .Expressions.Add "MFFault", mipMfFault, 1#, ilckVariableEqualValueMeansOK, LoadResString(ridFanSystemFault)
d1217 1
a1217 2
            .Expressions.Add "HighDP", mopHxHighDpFault, 1#, ilckVariableEqualValueMeansOK, LoadResString(ridHxDpFault)
            .Expressions.Add "MFDrvFlt", mipMfDriveFault, 1#, ilckVariableEqualValueMeansOK, LoadResString(ridFanDriveFault)
a1262 2
        .Add Array(ocStartIdle, ocStopIdle), CStr(ocCancelIdle)
        .Add Array(ocStartIdle, ocStopIdle), CStr(ocResetIdle)
d1274 1
d1445 1
a1445 1
        .AddWithoutFlag "Pabs", mipPabs, ioInput
d1449 8
a1456 8
        .AddWithoutFlag "PLCEstopOut", mipReadyToRun, ioInput
        .AddWithoutFlag "MF_KeySwitch", mipMfReady, ioInput
        .AddWithoutFlag "MainFanPermissive", mipMfPerm, ioInput
        .AddWithoutFlag "EStopStatus", mipEstopStatus, ioInput  'MfEstop is not available
        .AddWithoutFlag "MainFanFault", mipMfFault, ioInput
        .AddWithoutFlag "MainFanDriveFault", mipMfDriveFault, ioInput
        .AddWithoutFlag "MainFanStatus", mipMfStatus, ioInput
        
d1458 7
a1464 9
        For I = 1 To 3
            .AddWithoutFlag "MainFanSpeedLim" & I, mipMFSpeedLimitCommand(I), ioInput
            .AddWithoutFlag "MainFanSpeedLim" & I & "Status", mipMFSpeedLimitStatus(I), ioInput
        Next I
        .AddWithoutFlag "MainFanAuxStatus", mipAuxStatus, ioInput
        .AddWithoutFlag "MainFanAuxReady", mipAuxReady, ioInput
        .AddWithoutFlag "MainFanAuxPerm", mipAuxPerm, ioInput
        .AddWithoutFlag "MainFanAuxFault", mipAuxFault, ioInput
        .AddWithoutFlag "NozzleSmall", mipNozzleSmallLimSw, ioInput
d1466 18
a1483 6
''        .AddWithoutFlag "IdlePermissive", mipIdleCmdPerm, ioInput
''        .AddWithoutFlag "IdleOpenStatus", mipIdleOpenStatus, ioInput
''        .AddWithoutFlag "IdleClosedStatus", mipIdleClosedStatus, ioInput
''        .AddWithoutFlag "IdleInMotion", mipIdleMotionStatus, ioInput
''        .AddWithoutFlag "IdleLightFault", mipIdleLight, ioInput
''        .AddWithoutFlag "IdleDriveFault", mipIdleFault, ioInput
a1484 6
        .AddWithoutFlag "DynoSpeedTgt", mipDySpeedTarget, ioInput
        .AddWithoutFlag "DynoSpeedRate", mipDySpeedRate, ioInput
        .AddWithoutFlag "DynoMode", mipDyMode, ioInput
        .AddWithoutFlag "ThrMode", mipThrMode, ioInput
        .AddWithoutFlag "BLFanStatus", mipBLFanStatus, ioInput
        
a1485 1
        .AddWithoutFlag "NozzleSize", mopNozzle, ioOutput
d1487 1
a1487 1
        .AddWithoutFlag "WindspeedRaw", mopWsControl, ioOutput
d1489 1
a1489 1
        .AddWithoutFlag "Blockage", mopBlockageCorr, ioOutput
d1494 5
a1498 6
        .AddWithoutFlag "MainFanAuxControl", mopAuxControl, ioOutput
        .AddWithoutFlag "MainFanControl", mopMfControl, ioOutput
        .AddWithoutFlag "MainFanSpeedSetpt", mopFsSetpt, ioOutput
        .AddWithoutFlag "MainFanSpeedRamp", mopFsRamp, ioOutput
        .AddWithoutFlag "MainFanSpeedTarget", mopFsTargetDisp, ioOutput
        .AddWithoutFlag "MainFanRevEnable", mopMfReverseEnable, ioOutput
d1509 8
a1516 5
        .AddWithoutFlag "AutoIdleStatus", mopAutoIdleStatus, ioOutput
        .AddWithoutFlag "IdleStatus", mopIdleStatus, ioOutput
        .AddWithoutFlag "IdleOpenControl", mopIdleOpenControl, ioOutput
        .AddWithoutFlag "IdleCloseControl", mopIdleCloseControl, ioOutput
        .AddWithoutFlag "IdleResetControl", mopIdleResetControl, ioOutput
d1678 1
d1754 1
a1754 1
                If mipAuxStatus = 1 Then
d1761 11
a1771 1
                    .Abort , cmdTimedout
d1787 1
a1787 1
                    If (mipAuxStatus = 0#) Or (.TimeSince >= 30#) Then
d1797 1
a1797 1
                ElseIf mipAuxStatus <> 1# Or mopAuxControl <> 1# Then
a1989 11
            Case ocResetIdle
                mgvbAutoIdle = False
                mobjCommands.Clear "ResetIdle", ocStartIdle, ocStopIdle
                mopIdleCloseControl = 0#
                mopIdleOpenControl = 0#
                mopIdleResetControl = 1#
                If .TimeSince >= 2 Then
                    mopIdleResetControl = 0#
                    .Remove
                End If
                
a2656 1
    mopIdleResetControl = 0#
a3072 1
            LogMessage2 objEH
d3074 6
@
