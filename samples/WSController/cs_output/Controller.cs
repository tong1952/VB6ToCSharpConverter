// Converted from Controller.cls by vb6cs
// Date: 2026-05-11 03:43

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace Converted;
public class Controller : ICcrpTimerNotify
{
    private CommandListLib.Commands mobjCommands;
    private TalentProcessLib.Statuses mobjStatuses;
    private ControlLib.ControlConstants mobjConstants;
    private ControlLib.GlobalVars mobjGlobalVars;
    private AdvUtilLib.Stats mobjStats;
    private ParManager mobjParManager;
    private LoopTimer tmrMain;
    private ccrpTimer tmrRunOnce;
    private ccrpStopWatch tmrElapsed;
    private Interlocks mobjInterlocks;
    private ControlLib.Filter mobjWSFilter;
    private ControlLib.Filter mobjWSFilter2;
    private ConversionLib.Butterworth mobjWSFilterBW;
    private ConversionLib.Butterworth mobjWSFilterBW2;
    private ControlLib.Filter mobjWSctlFilter;
    private ControlLib.Filter mobjTrackFilter;
    private ControlLib.Filter mobjFFLeadLag;
    private ControlLib.Filter mobjRampRateFilter;
    private ControlLib.Filter mobjOutputFilter;
    private ControlLib.Monitor mobjWSMonitor;
    private ControlLib.Monitor mobjFSMonitor;
    private ControlLib.PIDControl mobjPID;
    private ParametersLib.Parameter mobjTrackPar;
    private TalentProcessLib.Events mobjEvents;
    private frmControl mobjForm;
    private SharedPropertyGroup mobjSPGConstants;
    private Collection mcolInvalidOpcodes;
    private PermissiveComponents mobjMFPermComponents;
    private PermissiveComponents mobjMFAuxPermComponents;
    private PermissiveComponents mobjIdleCmdPermComponents;
    private int mlID;
    private bool mbReady;
    private bool mbInited;
    private const bool mbKeepStats = true;
    private object mvCheckIlckValue;
    private string mccsPriority;
    private float mccfSlowFilterDelayTime;
    private float mccfWsTgtTrackFilTc;
    private bool mccbFilterInput;
    private bool mccbFilterOutput;
    private float mccfPidOutputFilTc;
    private float mccfMinFilTc;
    private float mccfMaxFilTc;
    private float mccfMainFanTimeout;
    private float mccfIdleTimeout;
    private float mccfMainFanAuxTimeout;
    private float mccfMainFanAuxPause;
    private float mccfMainFanAuxAutoOffTime;
    private float mccfResetFaultTime;
    private bool mccbAutoReset;
    private float mccfMainFanRestartTime;
    private int[] mcclNumCloseTimes = new int[3 + 1];
    private float[] mccfCloseTimesLarge = new float[10 + 1];
    private float[] mccfCloseTimeSpeedsLarge = new float[10 + 1];
    private float[] mccfCloseTimesMedium = new float[10 + 1];
    private float[] mccfCloseTimeSpeedsMedium = new float[10 + 1];
    private float[] mccfCloseTimesSmall = new float[10 + 1];
    private float[] mccfCloseTimeSpeedsSmall = new float[10 + 1];
    private float mccfUpdateInt;
    private float[] mccfCloseWsRamp = new float[3 + 1];
    private float[] mccfCloseWsFilTc = new float[3 + 1];
    private float mccfCtlMax;
    private float mccfCtlMin;
    private float mccfDisplayFilType;
    private float mccfDefFsRamp;
    private float mccfDefWsRamp;
    private float mccfFfGain;
    private float mccfFfLagCoef;
    private float mccfFfLagCoefTrack;
    private float mccfFfLeadCoef;
    private float mccfFfLeadCoefTrack;
    private float[] mccfWsFilterTc0 = new float[3 + 1];
    private float[] mccfWsFilterTc1 = new float[3 + 1];
    private float mccfTdFilterTC;
    private float[] mccfWsFilterFast = new float[3 + 1];
    private float[] mccfWsFilterFastTrack = new float[3 + 1];
    private float[] mccfNomFs = new float[3 + 1];
    private float[] mccfMaxFs = new float[3 + 1];
    private float mccfMaxFsPrepMode;
    private float mccfMaxFsStabMode;
    private float mccfFsMonitorError;
    private float mccfFsMonitorFilTc;
    private float mccfFsMonitorPeriod;
    private float mccfIntMax;
    private float mccfIntMin;
    private float[] mccfKp0 = new float[3 + 1];
    private float[] mccfKp1 = new float[3 + 1];
    private float[] mccfKp0Track = new float[3 + 1];
    private float[] mccfKp1Track = new float[3 + 1];
    private float[] mccfKpMin = new float[3 + 1];
    private float[] mccfKpMax = new float[3 + 1];
    private float[] mccfTi0 = new float[3 + 1];
    private float[] mccfTi1 = new float[3 + 1];
    private float[] mccfTi0Track = new float[3 + 1];
    private float[] mccfTi1Track = new float[3 + 1];
    private float[] mccfTiMin = new float[3 + 1];
    private float[] mccfTiMax = new float[3 + 1];
    private float[] mccfTd0 = new float[3 + 1];
    private float[] mccfTd1 = new float[3 + 1];
    private float[] mccfTd0Track = new float[3 + 1];
    private float[] mccfTd1Track = new float[3 + 1];
    private float[] mccfTdMin = new float[3 + 1];
    private float[] mccfTdMax = new float[3 + 1];
    private float mccfMinDeltaP;
    private float mccfMaxFsRamp;
    private float mccfMaxFsRampTrack;
    private float mccfMaxSuspendTime;
    private float mccfMinWsHeadwind;
    private float mccfMaxWsHeadwind;
    private float mccfMinWsRamp;
    private float mccfMaxWsRamp;
    private float mccfMaxWsError;
    private float mccfMinFs;
    private float mccfMinFsAutoMode;
    private float mccfMinFsRamp;
    private float mccfMinDeltaPFSMax;
    private float mccfMinWs;
    private float mccfMinWsForPid;
    private float mccfMinWsForPidDelay;
    private float mccfMinWsErrorSlow;
    private float mccfMinWsErrorSlowSmall;
    private float mccfMinDeltaPHxRatio;
    private float mccfMonitorError;
    private float mccfMonitorFilTc;
    private float mccfMonitorPeriod;
    private float mccfSlowRampFanspeed;
    private float mccfMaxWsRampLowSpeeds;
    private float[] mccfNomWs = new float[3 + 1];
    private float[] mccfMaxWs = new float[3 + 1];
    private float mccfMaxWsRateForI;
    private float mccfMinWsRateForI;
    private float mccfWsRateForI0;
    private float mccfWsRateForI1;
    private float mccfMaxWsPrepMode;
    private float mccfMaxWsStabMode;
    private bool mccbDefAutoIdle;
    private float mccfWSStartIdle;
    private float mccfWSStopIdle;
    private float mccfAutoIdleDeadband;
    private float mccfIdleFanspeed;
    private float[] mccfControlFilTc0 = new float[3 + 1];
    private float[] mccfControlFilTc1 = new float[3 + 1];
    private float[] mccfControlFilTcMin = new float[3 + 1];
    private float[] mccfControlFilTcMax = new float[3 + 1];
    private float mccfDrivePause;
    private string mccsMFPermComponents;
    private string mccsMFAuxPermComponents;
    private string mccsIdleCmdPermComponents;
    private float mccfCancelIdleTime;
    private string mccsDefTrackParName;
    private float[] mccfMaxFsLimit = new float[3 + 1];
    private float mccfTurnOffFanTime;
    private float mccfMaxRunTimeWithoutCooling;
    private bool mccbAutoFollowDyTarget;
    private bool mccbAutoFollowDyTargetUseTrackCoeffs;
    private bool mccbAutoIdleUsesLimitCommand;
    private bool mgvbWsTooLowForPID;
    private DateTime mgvdWsTooLowForPIDTime;
    private float mccfNomWSReverse;
    private float mccfNomFSReverse;
    private float mccfTimeAtZeroToForceWSZero;
    private float mccfUseTargetErrorForDerivLimit;
    private float mccfFinalRampDeltakph;
    private float mccfFinalRampTimeSec;
    private float mccfFinalRampDeltakphSmall;
    private float mccfFinalRampTimeSecSmall;
    private bool mccbFinalRampSnap;
    private int mcclSetpointRampMethod;
    private bool mccbFinalRampUseSetpoint;
    private float mccfWsRateForFinalRamp;
    private bool mccbFinalRampSnapBumpless;
    private string[] mccsWCoeffsFile = new string[3 + 1];
    private double[] mccdKpCoeffs = new double[3 + 1, 6 + 1];
    private double[] mccdKqnCoeffs = new double[3 + 1, 6 + 1];
    private double[] mccdKqplCoeffs = new double[3 + 1, 6 + 1];
    private double[] mccdWSvsFSCoeffs = new double[3 + 1, 6 + 1];
    private float[] mccfKpCoeffMin = new float[3 + 1];
    private float[] mccfKpCoeffMax = new float[3 + 1];
    private float[] mccfKqnCoeffMin = new float[3 + 1];
    private float[] mccfKqnCoeffMax = new float[3 + 1];
    private float[] mccfKqplCoeffMin = new float[3 + 1];
    private float[] mccfKqplCoeffMax = new float[3 + 1];
    private bool[] mccbOFDefault = new bool[18 + 1];
    private bool mccbOFPurgeOpen;
    private float mccfNozzleTimeout;
    private float mccfCFLeft1Min;
    private float mccfCFLeft1Max;
    private float mccfCFLeft2Min;
    private float mccfCFLeft2Max;
    private float mccfCFRight1Min;
    private float mccfCFRight1Max;
    private float mccfCFRight2Min;
    private float mccfCFRight2Max;
    private float mccfCFTopMin;
    private float mccfCFTopMax;
    private float mccfCFTimeout;
    private float mccfCFHoldTime;
    private float mccfCFTolerance;
    private float mipAirtemp;
    private float mipAirtempPlenum;
    private float mipRelHumid;
    private float mipDewpoint;
    private float mipPabs;
    private float mipDeltaP;
    private float mipDeltaPPlenum;
    private float mipEstopStatus;
    private float mipToPLCResetCmd;
    private float mipOpModeOff;
    private float mipOpModeMaint;
    private float mipOpModePrep;
    private float mipOpModeStab;
    private float mipOpModeTest;
    private float mipAuxDPLCCmdRequest;
    private float mipCoolPLCCmdRequest;
    private float mipMfPerm;
    private float mipMfRemote;
    private float mipMfLCFault;
    private float mipMfMCFault;
    private float mipCSStatus;
    private float mipLCReady;
    private float mipMCReady;
    private float mipMfStatus;
    private float mipFanspeed;
    private float[] mipMFSpeedLimitCommand = new float[3 + 1];
    private float[] mipMFSpeedLimitStatus = new float[3 + 1];
    private float mipAuxFault;
    private float mipAuxReady;
    private float mipNozzleSmallLimSw;
    private float mipNozzleMediumLimSw;
    private float mipNozzleLargeLimSw;
    private float mipIdleNozFlapMotorStop;
    private float mipIdleNozFlapOpenStatus;
    private float mipIdleNozFlapClosedStatus;
    private float mipIdleNozFlapMotorFault;
    private float[] mipIdleBypFlapOpenStatus = new float[3 + 1];
    private float[] mipIdleBypFlapClosedStatus = new float[3 + 1];
    private float mipCFLeft1Pos;
    private float mipCFLeft1Fault;
    private float mipCFLeft1Lockout;
    private float mipCFLeft2Pos;
    private float mipCFLeft2Fault;
    private float mipCFLeft2Lockout;
    private float mipCFRight1Pos;
    private float mipCFRight1Fault;
    private float mipCFRight1Lockout;
    private float mipCFRight2Pos;
    private float mipCFRight2Fault;
    private float mipCFRight2Lockout;
    private float mipCFTopPos;
    private float mipCFTopFault;
    private float mipCFTopLockout;
    private float mipBLFanStatus;
    private float mipDySpeedTarget;
    private float mipDySpeedRate;
    private float mipDyMode;
    private float mipThrMode;
    private float mopWsControl;
    private float mopFilDeltaP;
    private float mopFilDeltaPPlenum;
    private float mopAutoIdleStatus;
    private float mopIdleStatus;
    private float mopIdleStopControl;
    private float mopIdleStartControl;
    private float mopFsSetpt;
    private float mopAuxControl;
    private float mopMfControl;
    private float mopMFReset;
    private float mopAirDensityKgm3;
    private float mopBlockageCorr;
    private float mopFsRamp;
    private float mopFsTargetDisp;
    private float mopPstaticKpa;
    private float mopQPa;
    private float mopQnPa;
    private float mopQplPa;
    private float mopWSCalcMethod;
    private float mopKp;
    private float mopKqpl;
    private float mopKqn;
    private float mopWindspeed;
    private float mopWsRamp;
    private float mopWsSetpt;
    private float mopWsSetptDisp;
    private float mopWsTargetDisp;
    private float mopHeadwind;
    private float mopWsAchieved;
    private float mopWindspeedmps;
    private float mopWsSetptmps;
    private float mopWsTargetmps;
    private float mopWsRampmps;
    private float mopHeadwindmps;
    private float mopWsMode;
    private float mopNozzle;
    private float mopCFLeftStartCmd;
    private float mopCFLeftStopCmd;
    private float mopCFLeft1Target;
    private float mopCFLeft2Target;
    private float mopCFRightStartCmd;
    private float mopCFRightStopCmd;
    private float mopCFRight1Target;
    private float mopCFRight2Target;
    private float mopCFTopStartCmd;
    private float mopCFTopStopCmd;
    private float mopCFTopTarget;
    private float[] mopOFOpenCmd = new float[18 + 1];
    private float mopPurgeCmd;
    private float[] mopNozzleCmd = new float[3 + 1];
    private Priorities mgvlDefPriority;
    private WindTunnels mgvlWindTunnel;
    private float mgvfCloseWsFilTc;
    private float mgvfPurgeStatus;
    private float mgvfHeadwind;
    private float mgvfControlRamp;
    private float mgvfControlFilTc;
    private float mgvfCloseTime;
    private float mgvfCloseTimeRamp;
    private float mgvfFilterTmr;
    private float mgvfFbNormFs;
    private float mgvfFfNormFs;
    private float mgvfFfRawNormFs;
    private float mgvfAuxStatus;
    private float mgvfFilDeltaPControl;
    private float mgvfFilHxDeltaP;
    private float mgvfFsRamp;
    private float mgvfFsIncrement;
    private float mgvfFsNormSetpt;
    private float mgvfFsTarget;
    private float mgvfFsLimitedTarget;
    private float mgvfHxDpOverDp;
    private float mgvfMinKp;
    private float mgvfMaxKp;
    private float mgvfMinTi;
    private float mgvfMaxTi;
    private float mgvfMinTd;
    private float mgvfMaxTd;
    private float mgvfMinBlockage;
    private float mgvfMaxBlockage;
    private float mgvfMaxFs;
    private float mgvfMaxFsTarget;
    private float mgvfMaxWs;
    private float mgvfMfStartDelay;
    private float mgvfNomFsMax;
    private float mgvfNomWsMax;
    private float mgvfNormError;
    private float mgvfNormWsSetpt;
    private float mgvfOldRemoteStatus;
    private float mgvfOldCtlMode;
    private float mgvfOldMfStatus;
    private float mgvfOldAuxStatus;
    private NozzleSizes mgvlOldNozzle;
    private NozzleSizes mgvlNozzle;
    private float mgvfOldPurgeStopFanRequest;
    private float mgvfOldTsDoor4;
    private float mgvfOldWsIncrement;
    private float mgvfPressureRatio;
    private float mgvfPstaticControlKpa;
    private float mgvfQControlPa;
    private float mgvfRawHxDpOverDp;
    private float mgvfKp;
    private float mgvfKp0;
    private float mgvfKp1;
    private float mgvfKpEffect;
    private float mgvfTd;
    private float mgvfTd0;
    private float mgvfTd1;
    private float mgvfDeriviativeEffect;
    private float mgvfTi;
    private float mgvfTi0;
    private float mgvfTi1;
    private float mgvfIntegratorEffect;
    private float mgvfOldIntegral;
    private float mgvfWsError;
    private float mgvfWsFilTc;
    private float mgvfWsFilTc0;
    private float mgvfWsFilTc1;
    private float mgvfWsIncrement;
    private float mgvfWsRamp;
    private float mgvfWsRateForI;
    private float mgvfWsTarget;
    private float mgvfWSTargetLimited;
    private FilterSpeeds mgvlFilterSpeed;
    private bool mgvbSeparateControl;
    private DateTime mgvdAuxOnTime;
    private bool mgvbAuxAutoOff;
    private DateTime mgvdDriveOnTime;
    private DateTime mgvdDriveOffTime;
    private bool mgvbAutoIdle;
    private string mgvsTrackParName;
    private DateTime mgvdLastIdleTime;
    private float mgvfTPlAbs;
    private float mgvfOldWsTarget;
    private float mgvfFfLeadCoeff;
    private float mgvfFfLagCoeff;
    private IdleStatuses mgvlOldIdleStatus;
    private float mgvfWsFilterFast;
    private float mgvfControlFilTc0;
    private float mgvfControlFilTc1;
    private float mgvfControlFilTcMin;
    private float mgvfControlFilTcMax;
    private float mgvfTrackC0;
    private float mgvfTrackC1;
    private DateTime mgvdWDTTime;
    private DateTime mgvdTurnOffFanTime;
    private float mgvfGlycolPumpInRemoteOrOn;
    private float mgvfGlycolSystemInRemoteOrOn;
    private DateTime mgvdLastAuxOnTime;
    private bool mgvbStopAutoIdle;
    private bool mgvbAutoIdleSpeedHasBeenHigh;
    private bool mgvbAutoIdleSpeedHasBeenLow;
    private bool mgvbAutoIdleWantsIdle;
    private int mgvlOIdIdleFlapsControl;
    private CWTCSWSController.Modes mgvlWSControlMode;
    private DateTime mgvdLastTimeFSTargetNotZero;
    private DateTime mgvdLastTimeWSTargetNotZero;
    private bool mgvbUsingCloseTime;
    private bool mgvbWSComingBackToSetpt;
    private bool mgvbUsingFinalRamp;
    private float mgvfWSTargetError;
    private float mgvfTargetNormError;
    private bool mgvbFinalRampHasSnapped;
    private bool mgvbFinalRampOKToUse;
    private NozzleSizes mgvlNozzleTarget;
    private DateTime mgvdNozzleSizeCmdTime;
    private DateTime mgvdMFInterlockedTime;
    private bool mgvbHoldCRUntilFSOK;
    private DateTime mgvdCFLeftCmdTime;
    private bool mgvbCFLeftCtl;
    private bool mgvbCFLeft1Achieved;
    private bool mgvbCFLeft2Achieved;
    private DateTime mgvdCFLeft1AchievedTime;
    private DateTime mgvdCFLeft2AchievedTime;
    private DateTime mgvdCFRightCmdTime;
    private bool mgvbCFRightCtl;
    private bool mgvbCFRight1Achieved;
    private bool mgvbCFRight2Achieved;
    private DateTime mgvdCFRight1AchievedTime;
    private DateTime mgvdCFRight2AchievedTime;
    private DateTime mgvdCFTopCmdTime;
    private bool mgvbCFTopCtl;
    private bool mgvbCFTopAchieved;
    private DateTime mgvdCFTopAchievedTime;
    private float mgvfKpCoeffMin;
    private float mgvfKpCoeffMax;
    private float mgvfKqnCoeffMin;
    private float mgvfKqnCoeffMax;
    private float mgvfKqplCoeffMin;
    private float mgvfKqplCoeffMax;
    private double mgvdVmps;
    private double[] mgvdWSvsFSCoeffs = new double[6 + 1];
    private double[] mgvdKpCoeffs = new double[6 + 1];
    private double[] mgvdKqnCoeffs = new double[6 + 1];
    private double[] mgvdKqplCoeffs = new double[6 + 1];
    private double mgvdKp;
    private double mgvdKqn;
    private double mgvdKqpl;
    private double mgvdPoSCPa;
    private double mgvdPoTSPa;
    private double mgvdTSCAbs;
    private double mgvdTTSAbs;
    private double mgvdXPPl;
    private enum InitTypes
    {
        initAll = 0,
        initConstants = 1,
        initPars = 2
    }

    public object ParManager
    {
        get
        {
            object __result = null;
            __result = mobjParManager;
            return __result;
        }
    }

    public string GetShowStatusText()
    {
        string __result = "";
        try
        {
            mobjInterlocks(ocShowStatus.ToString()).Check();
            if (Err)
            {
                __result = LoadResString(stsTheFollowingFaultsExist, ex.Message);
            }
            else
            {
                __result = LoadResString(stsNoFaultsDetected);
            }

            return __result;
        }
        catch
        {
            // On Error Resume Next: exceptions suppressed
        }

        return __result;
    }

    private void OutputConversions()
    {
        int I = 0;
        try
        {
            if (mobjCommands.Exists(ocSetFanSpeed) == false)
            {
                mopFsTargetDisp = mgvfFsTarget;
                mopFsRamp = mgvfFsRamp;
            }

            if (mobjCommands.Exists(ocSetWindSpeed) == false)
            {
                mopWsTargetDisp = mgvfWsTarget;
                mopWsRamp = mgvfWsRamp;
            }

            if (mobjCommands.Exists(ocTrackWindSpeed) == false)
            {
                mopHeadwind = mgvfHeadwind;
            }

            if (mopWsMode >= modeTrack)
            {
                if (mgvbAutoIdle)
                {
                    mopAutoIdleStatus = 1;
                }
                else
                {
                    mopAutoIdleStatus = 0;
                }
            }
            else
            {
                if (mccbDefAutoIdle)
                {
                    mopAutoIdleStatus = 1;
                }
                else
                {
                    mopAutoIdleStatus = 0;
                }
            }

            if (mopWsMode == modeTrack & mgvbAutoIdle)
            {
                if (mgvbAutoIdleWantsIdle)
                {
                    if (Math.Abs(mopFsSetpt) > mccfMaxFsLimit(1))
                    {
                    }
                    else
                    {
                        mopIdleStartControl = 1;
                        mopIdleStopControl = 0;
                    }
                }
                else
                {
                    mopIdleStopControl = 1;
                    mopIdleStartControl = 0;
                }
            }

            if (mopWsMode == modeAuto)
            {
                if (mobjStatuses.GetValue(STS_WINDSPEED_ACHIEVED))
                {
                    mopWsAchieved = 1;
                }
                else
                {
                    mopWsAchieved = 0;
                }
            }
            else
            {
                mopWsAchieved = 0;
            }

            mopNozzle = mgvlNozzle;
            if (mgvlWindTunnel == wtCWT)
            {
                mopNozzleCmd(nsSmall) = 0;
                mopNozzleCmd(nsMedium) = 0;
                mopNozzleCmd(nsLarge) = 0;
                if (mobjInterlocks(ocSetNozzleSize.ToString()).IsOK == false)
                {
                    mgvdNozzleSizeCmdTime = 0;
                }
                else if (mipMfStatus == 0 & SecondsSinceTime(mgvdNozzleSizeCmdTime) <= mccfNozzleTimeout)
                {
                    switch (mgvlNozzleTarget)
                    {
                        case nsSmall:
                            if (mipNozzleSmallLimSw == 1)
                            {
                                mgvlNozzleTarget = 0;
                            }
                            else
                            {
                                mopNozzleCmd(nsSmall) = 1;
                            }

                            break;
                        case nsMedium:
                            if (mipNozzleMediumLimSw == 1)
                            {
                                mgvlNozzleTarget = 0;
                            }
                            else
                            {
                                mopNozzleCmd(nsMedium) = 1;
                            }

                            break;
                        case nsLarge:
                            if (mipNozzleLargeLimSw == 1)
                            {
                                mgvlNozzleTarget = 0;
                            }
                            else
                            {
                                mopNozzleCmd(nsLarge) = 1;
                            }

                            break;
                    }
                }
            }

            if (mgvlWindTunnel == wtCWT)
            {
                if (mopIdleStartControl == 1)
                {
                    mopIdleStopControl = 0;
                    if (mopIdleStatus == IdleStatuses.isIdle)
                    {
                        mopIdleStartControl = 0;
                    }
                }
                else if (mopIdleStopControl == 1)
                {
                    mopIdleStartControl = 0;
                    if (mopIdleStatus == IdleStatuses.isNormal)
                    {
                        mopIdleStopControl = 0;
                    }
                }
            }

            if (mgvlWindTunnel == wtCWT)
            {
                mopCFLeftStartCmd = 0;
                mopCFLeftStopCmd = 1;
                if (mipAuxDPLCCmdRequest == 0)
                {
                    mopCFLeft1Target = Ctl_Limit(mipCFLeft1Pos, mccfCFLeft1Min, mccfCFLeft1Max);
                    mopCFLeft2Target = Ctl_Limit(mipCFLeft2Pos, mccfCFLeft2Min, mccfCFLeft2Max);
                }
                else if (mgvbCFLeftCtl == false)
                {
                    mgvdCFLeftCmdTime = 0;
                }
                else if (mobjInterlocks(ocSetCFlapLeft.ToString()).IsOK == false)
                {
                    mgvdCFLeftCmdTime = 0;
                }
                else if (mipMfStatus == 0 & SecondsSinceTime(mgvdCFLeftCmdTime) <= mccfCFTimeout)
                {
                    if (mgvbCFLeft1Achieved == false | mgvbCFLeft2Achieved == false)
                    {
                        mopCFLeftStartCmd = 1;
                        mopCFLeftStopCmd = 0;
                    }
                    else if (SecondsSinceTime(mgvdCFLeft1AchievedTime) <= mccfCFHoldTime | SecondsSinceTime(mgvdCFLeft2AchievedTime) <= mccfCFHoldTime)
                    {
                        mopCFLeftStartCmd = 1;
                        mopCFLeftStopCmd = 0;
                    }
                    else
                    {
                        mgvbCFLeftCtl = false;
                    }
                }

                mopCFRightStartCmd = 0;
                mopCFRightStopCmd = 1;
                if (mipAuxDPLCCmdRequest == 0)
                {
                    mopCFRight1Target = Ctl_Limit(mipCFRight1Pos, mccfCFRight1Min, mccfCFRight1Max);
                    mopCFRight2Target = Ctl_Limit(mipCFRight2Pos, mccfCFRight2Min, mccfCFRight2Max);
                }
                else if (mgvbCFRightCtl == false)
                {
                    mgvdCFRightCmdTime = 0;
                }
                else if (mobjInterlocks(ocSetCFlapRight.ToString()).IsOK == false)
                {
                    mgvdCFRightCmdTime = 0;
                }
                else if (mipMfStatus == 0 & SecondsSinceTime(mgvdCFRightCmdTime) <= mccfCFTimeout)
                {
                    if (mgvbCFRight1Achieved == false | mgvbCFRight2Achieved == false)
                    {
                        mopCFRightStartCmd = 1;
                        mopCFRightStopCmd = 0;
                    }
                    else if (SecondsSinceTime(mgvdCFRight1AchievedTime) <= mccfCFHoldTime | SecondsSinceTime(mgvdCFRight2AchievedTime) <= mccfCFHoldTime)
                    {
                        mopCFRightStartCmd = 1;
                        mopCFRightStopCmd = 0;
                    }
                    else
                    {
                        mgvbCFRightCtl = false;
                    }
                }

                mopCFTopStartCmd = 0;
                mopCFTopStopCmd = 1;
                if (mipAuxDPLCCmdRequest == 0)
                {
                    mopCFTopTarget = Ctl_Limit(mipCFTopPos, mccfCFTopMin, mccfCFTopMax);
                }
                else if (mgvbCFTopCtl == false)
                {
                    mgvdCFTopCmdTime = 0;
                }
                else if (mobjInterlocks(ocSetCFlapTop.ToString()).IsOK == false)
                {
                    mgvdCFTopCmdTime = 0;
                }
                else if (mipMfStatus == 0 & SecondsSinceTime(mgvdCFTopCmdTime) <= mccfCFTimeout)
                {
                    if (mgvbCFTopAchieved == false)
                    {
                        mopCFTopStartCmd = 1;
                        mopCFTopStopCmd = 0;
                    }
                    else if (SecondsSinceTime(mgvdCFTopAchievedTime) <= mccfCFHoldTime)
                    {
                        mopCFTopStartCmd = 1;
                        mopCFTopStopCmd = 0;
                    }
                    else
                    {
                        mgvbCFTopCtl = false;
                    }
                }
            }

            if (mopWsMode != modeAuto)
            {
                mopWsSetptDisp = mopWsSetpt;
            }

            if (mgvfFsTarget != 0)
            {
                mgvdLastTimeFSTargetNotZero = Now;
            }

            if (mgvfWsTarget != 0)
            {
                mgvdLastTimeWSTargetNotZero = Now;
            }

            for (I = 1; I <= 18; I++)
            {
                if (mccbOFDefault(I))
                {
                    mopOFOpenCmd(I) = 1;
                }
                else if (mopPurgeCmd)
                {
                    mopOFOpenCmd(I) = IIf(mccbOFPurgeOpen, 1, 0);
                }
                else
                {
                    mopOFOpenCmd(I) = 0;
                }
            }

            if (mccbAutoReset)
            {
                if (mipToPLCResetCmd == 1)
                {
                    mopMFReset = 1;
                }
                else if (mopMFReset == 1)
                {
                    if (mobjCommands.Exists(ocResetFaults) == false)
                    {
                        mopMFReset = 0;
                    }
                }
            }

            mopPstaticKpa = mgvdPoTSPa / 1000;
            mopWindspeedmps = mopWindspeed / 3.6;
            mopWsSetptmps = mopWsSetptDisp / 3.6;
            mopWsTargetmps = mopWsTargetDisp / 3.6;
            mopWsRampmps = mopWsRamp / 3.6;
            mopHeadwindmps = mopHeadwind / 3.6;
            return;
        }
        catch
        {
            // On Error Resume Next: exceptions suppressed
        }
    }

    private void SaveOldStates()
    {
        try
        {
            mgvfOldCtlMode = mopWsMode;
            mgvfOldRemoteStatus = mipMfRemote;
            mgvfOldMfStatus = mipMfStatus;
            mgvfOldAuxStatus = mgvfAuxStatus;
            mgvfOldWsIncrement = mgvfWsIncrement;
            mgvfOldCtlMode = mopWsMode;
            mgvlOldNozzle = mgvlNozzle;
            mgvlOldIdleStatus = mopIdleStatus;
            return;
        }
        catch
        {
            // On Error Resume Next: exceptions suppressed
        }
    }

    private void MonitorStatuses()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        switch (mopWsMode)
        {
            case modeAuto:
                if (mobjWSMonitor.MonitorParameter(mopWindspeed))
                {
                    if (mobjCommands.Exists(Opcodes.ocSetWindSpeed) == false)
                    {
                        mobjStatuses.SetValue(STS_WINDSPEED_ACHIEVED, true);
                    }
                }
                else
                {
                    mobjStatuses.SetValue(STS_WINDSPEED_ACHIEVED, false);
                }

                break;
            case modeManual:
                if (mipMfStatus == 1)
                {
                    if (mobjFSMonitor.MonitorParameter(mipFanspeed))
                    {
                        if (mobjCommands.Exists(Opcodes.ocSetFanSpeed) == false)
                        {
                            mobjStatuses.SetValue(STS_FANSPEED_ACHIEVED, true);
                        }
                    }
                    else
                    {
                        mobjStatuses.SetValue(STS_FANSPEED_ACHIEVED, false);
                    }
                }
                else if (mobjStatuses.GetValue(STS_FANSPEED_ACHIEVED) == true)
                {
                    mobjStatuses.SetValue(STS_FANSPEED_ACHIEVED, false);
                }

                break;
        }

        if (mopWsMode != mgvfOldCtlMode)
        {
            switch (mopWsMode)
            {
                case modeManual:
                    LogMessage(LoadResString(ridModeManual));
                    break;
                case modeAuto:
                    LogMessage(LoadResString(ridModeAuto));
                    break;
                case modeTrack:
                    LogMessage(LoadResString(ridModeTrack));
                    break;
            }
        }

        if (mgvfAuxStatus == 1)
        {
            if (mgvdAuxOnTime == 0)
            {
                mgvdAuxOnTime = Now;
            }

            mgvdLastAuxOnTime = Now;
            mobjStatuses.SetValue(STS_MAINFAN_AUX_OFF, false);
            mobjStatuses.SetValue(STS_MAINFAN_AUX_ON, true);
        }
        else if (mgvfAuxStatus == 0)
        {
            mgvdAuxOnTime = 0;
            mobjStatuses.SetValue(STS_MAINFAN_AUX_OFF, true);
            mobjStatuses.SetValue(STS_MAINFAN_AUX_ON, false);
        }
        else
        {
            mobjStatuses.SetValue(STS_MAINFAN_AUX_OFF, false);
            mobjStatuses.SetValue(STS_MAINFAN_AUX_ON, false);
        }

        if (mipMfStatus == 1)
        {
            if (mgvdDriveOnTime == 0)
            {
                mgvdDriveOnTime = Now;
            }

            mgvdDriveOffTime = 0;
            mobjStatuses.SetValue(STS_MAINFAN_OFF, false);
            mobjStatuses.SetValue(STS_MAINFAN_ON, true);
        }
        else
        {
            if (mgvdDriveOnTime > 0)
            {
                if (mgvdDriveOffTime == 0)
                {
                    mgvdDriveOffTime = Now;
                }
            }

            mgvdDriveOnTime = 0;
            mobjStatuses.SetValue(STS_MAINFAN_OFF, true);
            mobjStatuses.SetValue(STS_MAINFAN_ON, false);
        }

        if (WindTunnel == wtCWT)
        {
            switch (mopIdleStatus)
            {
                case isIdle:
                    mobjStatuses.SetValue(STS_IDLE_STARTED, true);
                    mobjStatuses.SetValue(STS_IDLE_STOPPED, false);
                    break;
                case isNormal:
                    mobjStatuses.SetValue(STS_IDLE_STARTED, false);
                    mobjStatuses.SetValue(STS_IDLE_STOPPED, true);
                    break;
                default:
                    mobjStatuses.SetValue(STS_IDLE_STARTED, false);
                    mobjStatuses.SetValue(STS_IDLE_STOPPED, false);
                    break;
            }

            if (mgvbCFLeftCtl)
            {
                if (Math.Abs(mopCFLeft1Target - mipCFLeft1Pos) <= mccfCFTolerance)
                {
                    if (mgvbCFLeft1Achieved == false)
                    {
                        mgvdCFLeft1AchievedTime = Now;
                    }

                    mgvbCFLeft1Achieved = true;
                }
                else
                {
                    mgvbCFLeft1Achieved = false;
                }

                if (Math.Abs(mopCFLeft2Target - mipCFLeft2Pos) <= mccfCFTolerance)
                {
                    if (mgvbCFLeft2Achieved == false)
                    {
                        mgvdCFLeft2AchievedTime = Now;
                    }

                    mgvbCFLeft2Achieved = true;
                }
                else
                {
                    mgvbCFLeft2Achieved = false;
                }
            }
            else
            {
                mgvbCFLeft1Achieved = false;
                mgvbCFLeft2Achieved = false;
            }

            if (mgvbCFRightCtl)
            {
                if (Math.Abs(mopCFRight1Target - mipCFRight1Pos) <= mccfCFTolerance)
                {
                    if (mgvbCFRight1Achieved == false)
                    {
                        mgvdCFRight1AchievedTime = Now;
                    }

                    mgvbCFRight1Achieved = true;
                }
                else
                {
                    mgvbCFRight1Achieved = false;
                }

                if (Math.Abs(mopCFRight2Target - mipCFRight2Pos) <= mccfCFTolerance)
                {
                    if (mgvbCFRight2Achieved == false)
                    {
                        mgvdCFRight2AchievedTime = Now;
                    }

                    mgvbCFRight2Achieved = true;
                }
                else
                {
                    mgvbCFRight2Achieved = false;
                }
            }
            else
            {
                mgvbCFRight1Achieved = false;
                mgvbCFRight2Achieved = false;
            }

            if (mgvbCFTopCtl)
            {
                if (Math.Abs(mopCFTopTarget - mipCFTopPos) <= mccfCFTolerance)
                {
                    if (mgvbCFTopAchieved == false)
                    {
                        mgvdCFTopAchievedTime = Now;
                    }

                    mgvbCFTopAchieved = true;
                }
                else
                {
                    mgvbCFTopAchieved = false;
                }
            }
            else
            {
                mgvbCFTopAchieved = false;
            }
        }

        goto Done;
        Error();
        sRaiseErrorText = ex.Message;
        bRaiseError = true;
        lRaiseErrorNumber = ex.HResult;
        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        if (bRaiseError)
        {
            EH_RaiseError(lRaiseErrorNumber, TypeName(this), PROC_NAME, sRaiseErrorText);
        }

        return;
    }

    private void CheckInterlocks()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        if (mopWsMode > modeManual)
        {
            if (!mobjInterlocks(ocSetModeAuto.ToString()).IsOK)
            {
                // On Error Resume Next — nested handler; see enclosing try/catch
                mobjInterlocks(ocSetModeAuto.ToString()).Check();
                sExplanationText = ex.Message;
                // On Error GoTo Error — nested handler not restructured
                LogMessage(LoadResString(ridModeManIlck, sExplanationText), mtWarning);
                mopWsMode = modeManual;
                mgvfFsTarget = mipFanspeed;
                mgvfFsRamp = mccfMaxFsRamp;
                if (mipMfStatus == 1)
                {
                    mopMfControl = 1;
                }
                else
                {
                    mopMfControl = 0;
                }

                if (mgvfAuxStatus == 1)
                {
                    mopAuxControl = 1;
                }
                else
                {
                    mopAuxControl = 0;
                }
            }
            else if (mipMfStatus == 0)
            {
                sExplanationText = LoadResString(ridMfTrippedOut);
                LogMessage(LoadResString(ridModeManIlck, sExplanationText), mtWarning);
                mopWsMode = modeManual;
                mgvfFsTarget = 0;
                mopFsSetpt = 0;
                mopMfControl = 0;
                if (mopAuxControl == 1)
                {
                    mopAuxControl = IIf(mgvfAuxStatus == 1, 1, 0);
                }
                else
                {
                    mopAuxControl = 0;
                }
            }
            else if (mopWsMode == modeTrack & !mobjInterlocks(ocTrackWindSpeed.ToString()).IsOK)
            {
                // On Error Resume Next — nested handler; see enclosing try/catch
                mobjInterlocks(ocTrackWindSpeed.ToString()).Check();
                sExplanationText = ex.Message;
                // On Error GoTo Error — nested handler not restructured
                LogMessage(LoadResString(ridModeTrackIlck, sExplanationText), mtWarning);
                mopWsMode = modeAuto;
            }
            else if (mopWsMode == modeAuto & mgvlWindTunnel == wtCWT & mopIdleStatus != isNormal)
            {
                if (mipIdleNozFlapOpenStatus != 1)
                {
                    sExplanationText = string.Concat(IIf(sExplanationText != "", ", ", ""), LoadResString(ridIdleNozFlapNotOpen));
                }

                for (I = 1; I <= 3; I++)
                {
                    if (mipIdleBypFlapClosedStatus(I) != 1)
                    {
                        sExplanationText = string.Concat(IIf(sExplanationText != "", ", ", ""), LoadResString(ridIdleBypFlapNotClosed, I));
                    }
                }

                LogMessage(LoadResString(ridModeManIlck, sExplanationText), mtWarning);
                mopWsMode = modeManual;
                mgvfFsTarget = 0;
                mgvfFsRamp = mccfMaxFsRamp;
            }
        }

        if (mopWsMode == modeManual & mgvlWindTunnel == wtCWT & mgvfFsTarget > mccfMaxFsLimit(1) & mopIdleStatus != isNormal)
        {
            if (mipIdleNozFlapOpenStatus != 1)
            {
                sExplanationText = string.Concat(IIf(sExplanationText != "", ", ", ""), LoadResString(ridIdleNozFlapNotOpen));
            }

            for (I = 1; I <= 3; I++)
            {
                if (mipIdleBypFlapClosedStatus(I) != 1)
                {
                    sExplanationText = string.Concat(IIf(sExplanationText != "", ", ", ""), LoadResString(ridIdleBypFlapNotClosed, I));
                }
            }

            LogMessage(LoadResString(ridFanStoppedIlck, sExplanationText), mtWarning);
            mgvfFsTarget = 0;
            mgvfFsRamp = mccfMaxFsRamp;
        }

        if (mopWsMode == modeAuto & mgvfWsTarget > mccfMaxWsPrepMode & mipOpModePrep == 1)
        {
            sExplanationText = GetCommonMessage2(cmInPrepModeMaxIs, mccfMaxWsPrepMode);
            LogMessage(LoadResString(ridWindspeedReduced, mccfMaxWsPrepMode, sExplanationText), mtWarning);
            mgvfWsTarget = mccfMaxWsPrepMode;
            mgvfWsRamp = mccfMaxWsRamp;
        }

        if (mopWsMode == modeAuto & mgvfWsTarget > mccfMaxWsStabMode & mipOpModeStab == 1 & mgvlWindTunnel == wtAAWT)
        {
            sExplanationText = GetCommonMessage2(cmInStabModeMaxIs, mccfMaxWsStabMode);
            LogMessage(LoadResString(ridWindspeedReduced, mccfMaxWsStabMode, sExplanationText), mtWarning);
            mgvfWsTarget = mccfMaxWsStabMode;
            mgvfWsRamp = mccfMaxWsRamp;
        }

        if (mopWsMode == modeManual & mgvfFsTarget > mccfMaxFsPrepMode & mipOpModePrep == 1)
        {
            sExplanationText = GetCommonMessage2(cmInPrepModeMaxIs, mccfMaxFsPrepMode);
            LogMessage(LoadResString(ridFanspeedReduced, mccfMaxFsPrepMode, sExplanationText), mtWarning);
            mgvfFsTarget = mccfMaxFsPrepMode;
            mgvfFsRamp = mccfMaxFsRamp;
        }

        if (mopWsMode == modeManual & mgvfFsTarget > mccfMaxFsStabMode & mipOpModeStab == 1 & mgvlWindTunnel == wtAAWT)
        {
            sExplanationText = GetCommonMessage2(cmInStabModeMaxIs, mccfMaxFsStabMode);
            LogMessage(LoadResString(ridFanspeedReduced, mccfMaxFsStabMode, sExplanationText), mtWarning);
            mgvfFsTarget = mccfMaxFsStabMode;
            mgvfFsRamp = mccfMaxFsRamp;
        }

        if (mopMfControl == 1)
        {
            if (mgvfAuxStatus != 1 & SecondsSinceTime(mgvdLastAuxOnTime) >= mccfMaxRunTimeWithoutCooling)
            {
                mopWsMode = modeManual;
                mgvfFsTarget = 0;
                mopFsSetpt = 0;
                mopMfControl = 0;
                mopAuxControl = 0;
            }
        }

        if (!mobjInterlocks(ocStartMainFan.ToString()).IsOK)
        {
            if (mopWsMode >= modeManual)
            {
                mopWsMode = modeManual;
                mgvfFsTarget = 0;
                if (mipMfStatus == 1)
                {
                    if (mopMfControl == 1)
                    {
                        if (SecondsSinceTime(mgvdMFInterlockedTime) >= 60)
                        {
                            mopMfControl = 0;
                        }
                    }
                }
                else
                {
                    mopMfControl = 0;
                }

                if (mopAuxControl == 1)
                {
                    mopAuxControl = IIf(mgvfAuxStatus == 1, 1, 0);
                }
                else
                {
                    mopAuxControl = 0;
                }
            }
        }
        else
        {
            mgvdMFInterlockedTime = Now;
        }

        if (!mobjInterlocks(ocStartMainFanAux.ToString()).IsOK & mopAuxControl == 1)
        {
            if (mopWsMode > modeManual)
            {
                mopWsMode = modeManual;
            }

            mopFsSetpt = 0;
            mgvfFsTarget = 0;
            mopMfControl = 0;
            mopAuxControl = 0;
        }

        if (!mobjInterlocks(ocSetModeManual.ToString()).IsOK & mopWsMode > modeOff)
        {
            mopWsMode = modeOff;
            mopFsSetpt = 0;
            mgvfFsTarget = 0;
            mopMfControl = 0;
        }

        if (WindTunnel == wtCWT)
        {
            if (!mobjInterlocks(ocStartIdle.ToString()).IsOK)
            {
                if (mopWsMode == modeTrack & mgvbAutoIdle)
                {
                    mgvbAutoIdle = false;
                }

                mopIdleStartControl = 0;
                mopIdleStopControl = 0;
            }
        }

        goto Done;
        Error();
        sRaiseErrorText = ex.Message;
        bRaiseError = true;
        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        if (bRaiseError)
        {
            EH_RaiseError(OLEERR_METHOD_FAILED, TypeName(this), PROC_NAME, sRaiseErrorText);
        }

        return;
    }

    private void CalcWindspeed()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        if (mopWsMode == modeTrack | mopWsMode == modeAuto)
        {
            if (mgvfWsIncrement != 0 & mopWsMode == modeAuto)
            {
                mgvlFilterSpeed = fsFAST;
                mgvfFilterTmr = mccfSlowFilterDelayTime;
            }
            else if (Math.Abs(mopWindspeed - mopWsSetpt) >= IIf(mgvlNozzle == nsSmall, mccfMinWsErrorSlowSmall, mccfMinWsErrorSlow))
            {
                mgvlFilterSpeed = fsFAST;
                mgvfFilterTmr = mccfSlowFilterDelayTime * 1 - mgvfNormWsSetpt;
            }
            else
            {
                if (mgvfFilterTmr <= 0)
                {
                    mgvlFilterSpeed = fsSLOW;
                }
                else
                {
                    mgvfFilterTmr = mgvfFilterTmr - mccfUpdateInt;
                }
            }
        }
        else if (mopWsMode == modeManual)
        {
            if (mgvfFsIncrement != 0)
            {
                mgvlFilterSpeed = fsFAST;
            }
            else if (Math.Abs(mipFanspeed - mopFsSetpt) * mgvfNomWsMax / mgvfNomFsMax >= mccfMinWsErrorSlow)
            {
                mgvlFilterSpeed = fsFAST;
            }
            else
            {
                mgvlFilterSpeed = fsSLOW;
            }
        }
        else
        {
            mgvlFilterSpeed = fsFAST;
        }

        mgvfWsFilTc = mgvfWsFilTc0 + mgvfWsFilTc1 * mipFanspeed / mgvfNomFsMax;
        mgvfWsFilTc = Ctl_Limit(mgvfWsFilTc, mccfMinFilTc, mccfMaxFilTc);
        if (mccbFilterInput)
        {
            if (mccfDisplayFilType == 2)
            {
                if (mobjWSFilterBW == null)
                {
                    mobjWSFilterBW = new Butterworth();
                    fLastBWFilterTC = mgvfWsFilTc0;
                    mobjWSFilterBW.InitializeLowPass(1 / mccfUpdateInt, 1 / 2 * 3.14159 * mgvfWsFilTc0, 4);
                    mobjWSFilterBW2 = new Butterworth();
                    fLastBWFilterTC = mgvfWsFilTc0;
                    mobjWSFilterBW2.InitializeLowPass(1 / mccfUpdateInt, 1 / 2 * 3.14159 * mgvfWsFilTc0, 4);
                }

                if (mgvlFilterSpeed == fsFAST)
                {
                    mobjWSFilterBW.Reset();
                    fLastBWFilterTC = -1;
                }

                if (mgvfWsFilTc0 != fLastBWFilterTC)
                {
                    fLastBWFilterTC = mgvfWsFilTc0;
                    mobjWSFilterBW.InitializeLowPass(1 / mccfUpdateInt, 1 / 2 * 3.14159 * mgvfWsFilTc0, 4);
                    mobjWSFilterBW2.InitializeLowPass(1 / mccfUpdateInt, 1 / 2 * 3.14159 * mgvfWsFilTc0, 4);
                }

                mopFilDeltaP = mobjWSFilterBW.GetFilteredValue((double)mipDeltaP);
                mopFilDeltaPPlenum = mobjWSFilterBW2.GetFilteredValue((double)mipDeltaPPlenum);
            }
            else if (mobjWSFilter.FilterType == ctlRollingAverage)
            {
                if (mgvlFilterSpeed == fsFAST)
                {
                    mobjWSFilter.ClearData();
                    mobjWSFilter2.ClearData();
                }

                mobjWSFilter.FilterConst(0) = mgvfWsFilTc / mccfUpdateInt;
                mobjWSFilter2.FilterConst(0) = mgvfWsFilTc / mccfUpdateInt;
                mopFilDeltaP = mobjWSFilter.FilteredValue(mipDeltaP);
                mopFilDeltaPPlenum = mobjWSFilter2.FilteredValue(mipDeltaPPlenum);
            }
            else
            {
                if (mgvlFilterSpeed == fsFAST)
                {
                    mobjWSFilter.FilterConst(0) = mgvfWsFilterFast;
                    mobjWSFilter2.FilterConst(0) = mgvfWsFilterFast;
                }
                else
                {
                    mobjWSFilter.FilterConst(0) = mgvfWsFilTc;
                    mobjWSFilter2.FilterConst(0) = mgvfWsFilTc;
                }

                mopFilDeltaP = mobjWSFilter.FilteredValue(mipDeltaP);
                mopFilDeltaPPlenum = mobjWSFilter2.FilteredValue(mipDeltaPPlenum);
            }
        }
        else
        {
            mopFilDeltaP = mipDeltaP;
            mopFilDeltaPPlenum = mipDeltaPPlenum;
        }

        mgvfTPlAbs = mipAirtempPlenum + 273.15;
        dPsatExp = -6094.4642 * 1 / mgvfTPlAbs + 21.1249952 - 0.027245552 * mgvfTPlAbs + 1.6853396E-05 * Math.Pow(mgvfTPlAbs, 2) + 2.4575506 * Math.Log(mgvfTPlAbs);
        dPSat = Math.Exp(dPsatExp);
        dPv = dPSat * mipRelHumid / 100;
        if (mopWsMode <= modeManual)
        {
            mgvdVmps = LimitD(PolyDouble((double)mipFanspeed, mgvdWSvsFSCoeffs, 6), 0, mgvfMaxWs) / 3.6;
        }
        else
        {
            mgvdVmps = mopWsSetpt / 3.6;
        }

        mgvdKp = LimitD(PolyDouble(mgvdVmps * 3.6, mgvdKpCoeffs, 6), mgvfKpCoeffMin, mgvfKpCoeffMax);
        mgvdKqn = LimitD(PolyDouble(mgvdVmps * 3.6, mgvdKqnCoeffs, 6), mgvfKqnCoeffMin, mgvfKqnCoeffMax);
        mgvdKqpl = LimitD(PolyDouble(mgvdVmps * 3.6, mgvdKqplCoeffs, 6), mgvfKqplCoeffMin, mgvfKqplCoeffMax);
        mgvdPoSCPa = mipPabs * 1000 + mopFilDeltaP;
        mgvdPoTSPa = mipPabs * 1000 + mgvdKp * mopFilDeltaP;
        mgvdTSCAbs = mipAirtemp + 273.15;
        mgvdTTSAbs = Math.Pow(mgvdPoTSPa / mgvdPoSCPa, 1.4 - 1 / 1.4) * mgvdTSCAbs;
        mgvdXPPl = 0.6222 * dPv / mipPabs * 1000 - dPv;
        mopAirDensityKgm3 = mgvdPoTSPa / Rair * mgvdTSCAbs * 0.6222 / 0.6222 + mgvdXPPl + mgvdPoTSPa / Rwater * mgvdTSCAbs * mgvdXPPl / 0.6222 + mgvdXPPl;
        mopQnPa = mgvdKqn * mopFilDeltaP;
        if (mopQnPa < 0)
        {
            mopQnPa = 0;
        }

        mopQplPa = mgvdKqpl * mopFilDeltaPPlenum;
        if (mopQplPa < 0)
        {
            mopQplPa = 0;
        }

        if (mopWSCalcMethod == 1)
        {
            mopQPa = mopQplPa;
        }
        else
        {
            mopQPa = mopQnPa;
        }

        if (mopAirDensityKgm3 > 0)
        {
            if (mopAirDensityKgm3 > 0 & mopQPa >= 0)
            {
                mopWindspeed = 3.6 * Math.Sqrt(2 * mopQPa / mopAirDensityKgm3);
            }
            else
            {
                mopWindspeed = 0;
            }
        }
        else
        {
            mopWindspeed = 0;
        }

        if (mgvbSeparateControl)
        {
            mgvfControlFilTc = mgvfControlFilTc0 + mgvfControlFilTc1 * mgvfNormWsSetpt;
            mgvfControlFilTc = Ctl_Limit(mgvfControlFilTc, mgvfControlFilTcMin, mgvfControlFilTcMax);
            if (mgvfControlFilTc > mccfUpdateInt)
            {
                mobjWSctlFilter.FilterConst(0) = mgvfControlFilTc;
                mgvfFilDeltaPControl = mobjWSctlFilter.FilteredValue(IIf(mopWSCalcMethod == 1, mipDeltaPPlenum, mipDeltaP));
            }
            else
            {
                mgvfFilDeltaPControl = IIf(mopWSCalcMethod == 1, mipDeltaPPlenum, mipDeltaP);
            }

            if (mopWSCalcMethod == 1)
            {
                mgvfQControlPa = mgvdKqpl * mgvfFilDeltaPControl;
            }
            else
            {
                mgvfQControlPa = mgvdKqn * mgvfFilDeltaPControl;
            }

            if (mopAirDensityKgm3 > 0 & mgvfQControlPa >= 0)
            {
                mopWsControl = 3.6 * Math.Sqrt(2 * mgvfQControlPa / mopAirDensityKgm3);
            }
            else
            {
                mopWsControl = 0;
            }
        }
        else
        {
            mopWsControl = mopWindspeed;
        }

        if (mopIdleStatus == isIdle)
        {
            mopWindspeed = 0;
        }
        else if (mipFanspeed <= mccfMinDeltaPFSMax)
        {
            if (mopFilDeltaP <= mccfMinDeltaP)
            {
                mopWindspeed = 0;
                mopWsControl = 0;
            }
            else if (FanspeedIsSetToZero)
            {
                mopWindspeed = 0;
                mopWsControl = 0;
            }
        }

        if (mopWsMode == modeManual & mgvfFsTarget < 0 | mopWsMode > modeManual & mopWsSetpt < 0 & mipFanspeed < 0)
        {
            mopWindspeed = mccfNomWSReverse * mipFanspeed / mccfNomFSReverse;
            mopWsControl = mopWindspeed;
        }
        else if (mopWsMode >= modeAuto & mopWsSetpt <= mccfMinWsForPid & mopWsSetpt > 0 & mopIdleStatus == isNormal)
        {
            mopWindspeed = mgvfNomWsMax * mopFsSetpt / mccfFfGain * mgvfNomFsMax;
            mopWsControl = mopWindspeed;
        }

        mopKp = mgvdKp;
        mopKqpl = mgvdKqpl;
        mopKqn = mgvdKqn;
        goto Done;
        Error();
        sRaiseErrorText = ex.Message;
        bRaiseError = true;
        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        if (bRaiseError)
        {
            EH_RaiseError(OLEERR_METHOD_FAILED, TypeName(this), PROC_NAME, sRaiseErrorText);
        }

        return;
    }

    private bool FanspeedIsSetToZero()
    {
        bool __result = false;
        // TODO: On Error GoTo Error — handler label not found in this scope
        if (mopWsMode == modeOff)
        {
            __result = true;
        }
        else if (mopWsMode == modeManual)
        {
            if (mccfTimeAtZeroToForceWSZero < 0)
            {
            }
            else if (SecondsSinceTime(mgvdLastTimeFSTargetNotZero) >= mccfTimeAtZeroToForceWSZero)
            {
                __result = true;
            }
        }
        else
        {
            if (mccfTimeAtZeroToForceWSZero < 0)
            {
            }
            else if (SecondsSinceTime(mgvdLastTimeWSTargetNotZero) >= mccfTimeAtZeroToForceWSZero)
            {
                __result = true;
            }
        }

        goto Done;
        Error();
        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        return __result;
    }

    private void InputConversions()
    {
        if (mipPabs <= 10)
        {
            mipPabs = 101;
        }

        if (mipCSStatus == 1)
        {
            mgvfAuxStatus = 1;
        }
        else if (mipCSStatus == 0)
        {
            mgvfAuxStatus = 0;
        }
        else
        {
            mgvfAuxStatus = 2;
        }

        if (WindTunnel == wtAAWT)
        {
            mgvlNozzle = nsLarge;
            mopIdleStatus = isNormal;
        }
        else
        {
            if (mipNozzleSmallLimSw == 1)
            {
                mgvlNozzle = nsSmall;
            }
            else if (mipNozzleMediumLimSw == 1)
            {
                mgvlNozzle = nsMedium;
            }
            else if (mipNozzleLargeLimSw == 1)
            {
                mgvlNozzle = nsLarge;
            }
            else
            {
                mgvlNozzle = nsMoving;
            }

            if (mipIdleNozFlapOpenStatus == 1 & mipIdleBypFlapClosedStatus(1) == 1 & mipIdleBypFlapClosedStatus(2) == 1 & mipIdleBypFlapClosedStatus(3) == 1)
            {
                mopIdleStatus = isNormal;
            }
            else if (mipIdleNozFlapClosedStatus == 1 & mipIdleBypFlapOpenStatus(1) == 1 & mipIdleBypFlapOpenStatus(2) == 1 & mipIdleBypFlapOpenStatus(3) == 1)
            {
                mopIdleStatus = isIdle;
            }
            else if (mipAuxDPLCCmdRequest == 0)
            {
                mopIdleStatus = isMoving;
            }
            else
            {
                mopIdleStatus = isMoving;
            }
        }
    }

    private void CalcNozzleModeDependentConstants()
    {
        bool bChanged = false;
        bool bExists = false;
        int I = 0;
        int J = 0;
        if (mgvlNozzle != mgvlOldNozzle)
        {
            bChanged = true;
        }

        switch (mgvlNozzle)
        {
            case nsLarge:
            case nsSmall:
            case nsMedium:
                I = mgvlNozzle;
                break;
            default:
                I = nsLarge;
                break;
        }

        if (mgvlWindTunnel == wtCWT & mccbAutoFollowDyTarget & mccbAutoFollowDyTargetUseTrackCoeffs == false & mopWsMode == modeTrack & mgvsTrackParName == mccsDefTrackParName & mipDyMode == DY_MODE_SPEED | mipThrMode == THR_MODE_SPEED)
        {
            mgvfKp0 = mccfKp0(I);
            mgvfKp1 = mccfKp1(I);
            mgvfTi0 = mccfTi0(I);
            mgvfTi1 = mccfTi1(I);
            mgvfTd0 = mccfTd0(I);
            mgvfTd1 = mccfTd1(I);
            mgvfWsFilterFast = mccfWsFilterFast(I);
        }
        else if (mopWsMode == modeTrack)
        {
            mgvfKp0 = mccfKp0Track(I);
            mgvfKp1 = mccfKp1Track(I);
            mgvfTi0 = mccfTi0Track(I);
            mgvfTi1 = mccfTi1Track(I);
            mgvfTd0 = mccfTd0Track(I);
            mgvfTd1 = mccfTd1Track(I);
            mgvfWsFilterFast = mccfWsFilterFastTrack(I);
        }
        else
        {
            mgvfKp0 = mccfKp0(I);
            mgvfKp1 = mccfKp1(I);
            mgvfTi0 = mccfTi0(I);
            mgvfTi1 = mccfTi1(I);
            mgvfTd0 = mccfTd0(I);
            mgvfTd1 = mccfTd1(I);
            mgvfWsFilterFast = mccfWsFilterFast(I);
        }

        mgvfMinKp = mccfKpMin(I);
        mgvfMaxKp = mccfKpMax(I);
        mgvfMinTi = mccfTiMin(I);
        mgvfMaxTi = mccfTiMax(I);
        mgvfMinTd = mccfTdMin(I);
        mgvfMaxTd = mccfTdMax(I);
        mgvfWsFilTc0 = mccfWsFilterTc0(I);
        mgvfWsFilTc1 = mccfWsFilterTc1(I);
        mgvfCloseWsFilTc = mccfCloseWsFilTc(I);
        mgvfControlFilTc0 = mccfControlFilTc0(I);
        mgvfControlFilTc1 = mccfControlFilTc1(I);
        mgvfControlFilTcMin = mccfControlFilTcMin(I);
        mgvfControlFilTcMax = mccfControlFilTcMax(I);
        mgvfNomWsMax = mccfNomWs(I);
        mgvfMaxWs = mccfMaxWs(I);
        mgvfNomFsMax = mccfNomFs(I);
        mgvfMaxFs = mccfMaxFs(I);
        if (bChanged)
        {
            mobjSPGConstants.CreatePropertyByPosition(CC_FSMax, bExists).Value = mgvfMaxFs;
            mobjSPGConstants.CreatePropertyByPosition(CC_WSMax, bExists).Value = mgvfMaxWs;
        }

        for (J = 0; J <= 6; J++)
        {
            mgvdKpCoeffs(J) = mccdKpCoeffs(I, J);
            mgvdKqnCoeffs(J) = mccdKqnCoeffs(I, J);
            mgvdKqplCoeffs(J) = mccdKqplCoeffs(I, J);
            mgvdWSvsFSCoeffs(J) = mccdWSvsFSCoeffs(I, J);
        }

        mgvfKpCoeffMin = mccfKpCoeffMin(I);
        mgvfKpCoeffMax = mccfKpCoeffMax(I);
        mgvfKqnCoeffMin = mccfKqnCoeffMin(I);
        mgvfKqnCoeffMax = mccfKqnCoeffMax(I);
        mgvfKqplCoeffMin = mccfKqplCoeffMin(I);
        mgvfKqplCoeffMax = mccfKqplCoeffMax(I);
    }

    public ParametersLib.Parameter GetParameter(ref string ParName)
    {
        ParametersLib.Parameter __result = 0;
        // TODO: On Error GoTo Error — handler label not found in this scope
        __result = mobjParManager.Configurations.Parameters(ParName);
        goto Done;
        Error();
        bRaiseError = true;
        sRaiseErrorText = ex.Message;
        lRaiseErrorNumber = tceInvalidArgs;
        sRaiseErrorSource = ex.Source;
        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        if (bRaiseError)
        {
            throw new Exception(sRaiseErrorText);
        }

        return __result;
    }

    private void Start()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        // On Error GoTo Error_CreatingObjects — nested handler not restructured
        sClassName = "CommandListLib.Commands";
        mobjCommands = new CommandListLib.Commands();
        sClassName = "ControlLib.ControlConstants";
        mobjConstants = new ControlLib.ControlConstants();
        sClassName = "ControlLib.Monitor";
        mobjWSMonitor = new ControlLib.Monitor();
        mobjFSMonitor = new ControlLib.Monitor();
        sClassName = "ControlLib.Filter";
        mobjWSFilter = new ControlLib.Filter();
        mobjWSFilter2 = new ControlLib.Filter();
        mobjWSctlFilter = new ControlLib.Filter();
        mobjTrackFilter = new ControlLib.Filter();
        mobjFFLeadLag = new ControlLib.Filter();
        mobjRampRateFilter = new ControlLib.Filter();
        mobjOutputFilter = new ControlLib.Filter();
        sClassName = "PIDControl";
        mobjPID = new ControlLib.PIDControl();
        sClassName = "ControlLib.GlobalVars";
        mobjGlobalVars = new ControlLib.GlobalVars();
        sClassName = "Events";
        mobjEvents = new Events();
        sClassName = "Interlocks";
        mobjInterlocks = new Interlocks();
        sClassName = "Invalid Opcodes Collection";
        mcolInvalidOpcodes = new Collection();
        sClassName = "SharedPropertyGroupManager";
        objSPGM = new SharedPropertyGroupManager();
        sClassName = "SharedPropertyGroup - Constants";
        mobjSPGConstants = objSPGM.CreatePropertyGroup(SPG_CONSTANTS, 0, 0, bExists);
        sClassName = "PermissiveComponents";
        mobjMFPermComponents = new PermissiveComponents();
        mobjMFAuxPermComponents = new PermissiveComponents();
        mobjIdleCmdPermComponents = new PermissiveComponents();
        // On Error GoTo Error — nested handler not restructured
        mobjStats = new AdvUtilLib.Stats();
        tmrElapsed = new ccrpStopWatch();
        mgvlWindTunnel = WindTunnel;
        DoEvents();
        mobjInterlocks.Add(ocInitialize.ToString()).Expressions.AddBoolean("Initialized", mbInited, true, ilckVariableEqualValueMeansOK, LoadResString(stsControlSystemNotInitialized));
        mobjInterlocks.Add(ocStopMainFanAux.ToString()).Expressions.Add("FanRemote", mipMfRemote, 1, ilckVariableEqualValueMeansOK, LoadResString(ridFanNotinComputerCtl));
        mobjInterlocks.Add(ocStartMainFanAux.ToString()).IncludedInterlocks.Add(ocStopMainFanAux.ToString());
        mobjInterlocks.Add(ocStartMainFanAux.ToString()).Expressions.Add("MfAuxFlt", mipAuxFault, 1, ilckVariableEqualValueMeansOK, LoadResString(ridAuxFault));
        mobjInterlocks.Add(ocStartMainFanAux.ToString()).Expressions.Add("MFAuxReady", mipAuxReady, 1, ilckVariableEqualValueMeansOK, LoadResString(ridAuxNotReady));
        mobjInterlocks.Add(ocSetModeManual.ToString()).Expressions.Add("FanRemote", mipMfRemote, 1, ilckVariableEqualValueMeansOK, LoadResString(ridFanNotinComputerCtl));
        mobjInterlocks.Add(ocStopMainFan.ToString()).IncludedInterlocks.Add(ocSetModeManual.ToString());
        mobjInterlocks.Add(ocStartMainFan.ToString()).IncludedInterlocks.Add(ocSetModeManual.ToString());
        mobjInterlocks.Add(ocStartMainFan.ToString()).IncludedInterlocks.Add(ocStartMainFanAux.ToString());
        mobjInterlocks.Add(ocStartMainFan.ToString()).Expressions.Add("MFRemote", mipMfRemote, 1, ilckVariableEqualValueMeansOK, LoadResString(ridFanNotinComputerCtl));
        mobjInterlocks.Add(ocStartMainFan.ToString()).Expressions.Add("MFFault", mipMfLCFault, 1, ilckVariableEqualValueMeansOK, LoadResString(ridFanLCFault));
        mobjInterlocks.Add(ocStartMainFan.ToString()).Expressions.Add("MFPerm", mipMfPerm, 1, ilckVariableEqualValueMeansOK, LoadResString(ridFanDriveSafetyFault));
        mobjInterlocks.Add(ocStartMainFan.ToString()).Expressions.Add("MFDrvFlt", mipMfMCFault, 1, ilckVariableEqualValueMeansOK, LoadResString(ridFanMCFault));
        mobjInterlocks.Add(ocStartMainFan.ToString()).Expressions.Add("Purge", mopPurgeCmd, 0, ilckVariableEqualValueMeansOK, LoadResString(ridPurging));
        mobjInterlocks.Add(ocSetFanSpeed.ToString()).IncludedInterlocks.Add(ocStartMainFan.ToString());
        mobjInterlocks.Add(ocSetModeAuto.ToString()).IncludedInterlocks.Add(ocStartMainFan.ToString());
        mobjInterlocks.Add(ocSetWindSpeed.ToString()).IncludedInterlocks.Add(ocSetModeAuto.ToString());
        mobjInterlocks.Add(ocTrackWindSpeed.ToString()).IncludedInterlocks.Add(ocSetModeAuto.ToString());
        mobjInterlocks.Add(ocStartPurge.ToString()).Expressions.Add("MFOn", mipMfStatus, 0, ilckVariableEqualValueMeansOK, LoadResString(ridFanIsOnStopFanFirst));
        mobjInterlocks.Add(ocShowStatus.ToString()).IncludedInterlocks.Add(ocSetWindSpeed.ToString());
        if (WindTunnel == wtCWT)
        {
            mobjInterlocks.Add(ocStartIdle.ToString()).Expressions.Add("Fault", mipIdleNozFlapMotorFault, 1, ilckVariableEqualValueMeansOK, LoadResString(ridIdleDriveFault));
            mobjInterlocks.Add(ocStartIdle.ToString()).Expressions.Add("IdleStop", mipIdleNozFlapMotorStop, 0, ilckVariableEqualValueMeansOK, LoadResString(ridIdleMotorStopOn));
            mobjInterlocks.Add(ocStartIdle.ToString()).Expressions.Add("Estop", mipEstopStatus, 1, ilckVariableEqualValueMeansOK, LoadResString(ridEstop));
            mobjInterlocks.Add(ocStopIdle.ToString()).IncludedInterlocks.Add(ocStartIdle.ToString());
            mobjInterlocks.Add(ocShowIdleStatus.ToString()).IncludedInterlocks.Add(ocStartIdle.ToString());
            mobjInterlocks.Add(ocSetNozzleSize.ToString()).Expressions.Add("MFOn", mipMfStatus, 0, ilckVariableEqualValueMeansOK, LoadResString(ridFanIsOnStopFanFirst));
            mobjInterlocks.Add(ocSetNozzleSize.ToString()).Expressions.Add("Estop", mipEstopStatus, 1, ilckVariableEqualValueMeansOK, LoadResString(ridEstop));
            mobjInterlocks.Add(ocSetCFlapLeft1.ToString()).Expressions.Add("MFOn", mipMfStatus, 0, ilckVariableEqualValueMeansOK, LoadResString(ridFanIsOnStopFanFirst));
            mobjInterlocks.Add(ocSetCFlapLeft1.ToString()).Expressions.Add("Estop", mipEstopStatus, 1, ilckVariableEqualValueMeansOK, LoadResString(ridEstop));
            mobjInterlocks.Add(ocSetCFlapLeft1.ToString()).Expressions.Add("Fault1", mipCFLeft1Fault, 1, ilckVariableEqualValueMeansOK, LoadResString(ridCFLeftFault, 1));
            mobjInterlocks.Add(ocSetCFlapLeft1.ToString()).Expressions.Add("LockOut1", mipCFLeft1Lockout, 0, ilckVariableEqualValueMeansOK, LoadResString(ridCFLeftLockout, 1));
            mobjInterlocks.Add(ocSetCFlapLeft2.ToString()).Expressions.Add("MFOn", mipMfStatus, 0, ilckVariableEqualValueMeansOK, LoadResString(ridFanIsOnStopFanFirst));
            mobjInterlocks.Add(ocSetCFlapLeft2.ToString()).Expressions.Add("Estop", mipEstopStatus, 1, ilckVariableEqualValueMeansOK, LoadResString(ridEstop));
            mobjInterlocks.Add(ocSetCFlapLeft2.ToString()).Expressions.Add("Fault2", mipCFLeft2Fault, 1, ilckVariableEqualValueMeansOK, LoadResString(ridCFLeftFault, 2));
            mobjInterlocks.Add(ocSetCFlapLeft2.ToString()).Expressions.Add("LockOut2", mipCFLeft2Lockout, 0, ilckVariableEqualValueMeansOK, LoadResString(ridCFLeftLockout, 2));
            mobjInterlocks.Add(ocSetCFlapLeft.ToString()).IncludedInterlocks.Add(ocSetCFlapLeft1.ToString());
            mobjInterlocks.Add(ocSetCFlapLeft.ToString()).IncludedInterlocks.Add(ocSetCFlapLeft2.ToString());
            mobjInterlocks.Add(ocSetCFlapRight1.ToString()).Expressions.Add("MFOn", mipMfStatus, 0, ilckVariableEqualValueMeansOK, LoadResString(ridFanIsOnStopFanFirst));
            mobjInterlocks.Add(ocSetCFlapRight1.ToString()).Expressions.Add("Estop", mipEstopStatus, 1, ilckVariableEqualValueMeansOK, LoadResString(ridEstop));
            mobjInterlocks.Add(ocSetCFlapRight1.ToString()).Expressions.Add("Fault1", mipCFRight1Fault, 1, ilckVariableEqualValueMeansOK, LoadResString(ridCFRightFault, 1));
            mobjInterlocks.Add(ocSetCFlapRight1.ToString()).Expressions.Add("LockOut1", mipCFRight1Lockout, 0, ilckVariableEqualValueMeansOK, LoadResString(ridCFRightLockout, 1));
            mobjInterlocks.Add(ocSetCFlapRight2.ToString()).Expressions.Add("MFOn", mipMfStatus, 0, ilckVariableEqualValueMeansOK, LoadResString(ridFanIsOnStopFanFirst));
            mobjInterlocks.Add(ocSetCFlapRight2.ToString()).Expressions.Add("Estop", mipEstopStatus, 1, ilckVariableEqualValueMeansOK, LoadResString(ridEstop));
            mobjInterlocks.Add(ocSetCFlapRight2.ToString()).Expressions.Add("Fault2", mipCFRight2Fault, 1, ilckVariableEqualValueMeansOK, LoadResString(ridCFRightFault, 2));
            mobjInterlocks.Add(ocSetCFlapRight2.ToString()).Expressions.Add("LockOut2", mipCFRight2Lockout, 0, ilckVariableEqualValueMeansOK, LoadResString(ridCFRightLockout, 2));
            mobjInterlocks.Add(ocSetCFlapRight.ToString()).IncludedInterlocks.Add(ocSetCFlapRight1.ToString());
            mobjInterlocks.Add(ocSetCFlapRight.ToString()).IncludedInterlocks.Add(ocSetCFlapRight2.ToString());
            mobjInterlocks.Add(ocSetCFlapTop.ToString()).Expressions.Add("MFOn", mipMfStatus, 0, ilckVariableEqualValueMeansOK, LoadResString(ridFanIsOnStopFanFirst));
            mobjInterlocks.Add(ocSetCFlapTop.ToString()).Expressions.Add("Estop", mipEstopStatus, 1, ilckVariableEqualValueMeansOK, LoadResString(ridEstop));
            mobjInterlocks.Add(ocSetCFlapTop.ToString()).Expressions.Add("Fault", mipCFTopFault, 1, ilckVariableEqualValueMeansOK, LoadResString(ridCFTopFault));
            mobjInterlocks.Add(ocSetCFlapTop.ToString()).Expressions.Add("LockOut", mipCFTopLockout, 0, ilckVariableEqualValueMeansOK, LoadResString(ridCFTopLockout));
        }

        mcolInvalidOpcodes.Add(Array(ocSetModeAuto, ocSetFanSpeed, ocSetWindSpeed, ocTrackWindSpeed, ocStopMainFan, ocStopMainFanAux), ocSetFanSpeed.ToString());
        mcolInvalidOpcodes.Add(Array(ocSetFanSpeed, ocSetWindSpeed, ocTrackWindSpeed, ocStopMainFan, ocStopMainFanAux, ocStartIdle), ocSetWindSpeed.ToString());
        mcolInvalidOpcodes.Add(Array(ocSetFanSpeed, ocSetWindSpeed, ocTrackWindSpeed, ocStopMainFan, ocStopMainFanAux, ocStartIdle), ocTrackWindSpeed.ToString());
        mcolInvalidOpcodes.Add(Array(ocShowStatus), ocShowStatus.ToString());
        mcolInvalidOpcodes.Add(Array(ocStopMainFan, ocStopMainFanAux), ocStartMainFan.ToString());
        mcolInvalidOpcodes.Add(Array(ocStopMainFanAux), ocStartMainFanAux.ToString());
        mcolInvalidOpcodes.Add(Array(ocSetFanSpeed, ocSetWindSpeed, ocTrackWindSpeed, ocStartMainFan), ocStopMainFan.ToString());
        mcolInvalidOpcodes.Add(Array(ocSetFanSpeed, ocSetWindSpeed, ocTrackWindSpeed, ocStartMainFan, ocStartMainFanAux), ocStopMainFanAux.ToString());
        if (WindTunnel == wtCWT)
        {
            mcolInvalidOpcodes.Add(Array(ocSetWindSpeed, ocTrackWindSpeed, ocStopIdle, ocStartIdle), ocStartIdle.ToString());
            mcolInvalidOpcodes.Add(Array(ocStartIdle, ocStopIdle), ocStopIdle.ToString());
            mcolInvalidOpcodes.Add(Array(ocStartIdle, ocStopIdle), ocCancelIdle.ToString());
            mcolInvalidOpcodes.Add(Array(ocSetNozzleSize, ocStopNozzleMotion), ocSetNozzleSize.ToString());
            mcolInvalidOpcodes.Add(Array(ocSetNozzleSize), ocStopNozzleMotion.ToString());
            mcolInvalidOpcodes.Add(Array(ocSetCFlapLeft, ocSetCFlapLeft1, ocSetCFlapLeft2, ocSetCFlapRight, ocSetCFlapRight1, ocSetCFlapRight2, ocSetCFlapTop), ocStopCF.ToString());
        }// On Error Resume Next — nested handler; see enclosing try/catch
        Initialize(initConstants);
        // On Error GoTo Error — nested handler not restructured
        // On Error Resume Next — nested handler; see enclosing try/catch
        mobjGlobalVars.AddLong("Nozzle Size (1=Normal/Large, 2=Medium, 3=Small) (mgvfNozzle)", mgvlNozzle);
        mobjGlobalVars.AddSingle("Aux Summary Status (0=Off, 1=On, 2=Partial)", mgvfAuxStatus);
        mobjGlobalVars.AddSingle("WS Mode (1=Man,2=Auto,3=Track)", mopWsMode);
        mobjGlobalVars.AddLong("WS Internal Control Mode (1=Man,2=Auto,3=Track)", mgvlWSControlMode);
        mobjGlobalVars.AddLong("Windspeed Filter Speed (0=Slow, 1=Fast)", mgvlFilterSpeed);
        mobjGlobalVars.AddSingle("Filter Timer (when <=0, OK to use slow filter)", mgvfFilterTmr);
        mobjGlobalVars.AddSingle("Filter Type (Display) (0=First Order, 1=Rolling Avg, 2=BW)", mccfDisplayFilType);
        mobjGlobalVars.AddSingle("Ws Fil Tc 0", mgvfWsFilTc0);
        mobjGlobalVars.AddSingle("Ws Fil Tc 1", mgvfWsFilTc1);
        mobjGlobalVars.AddSingle("Ws Fil Tc (sec)", mgvfWsFilTc);
        mobjGlobalVars.AddSingle("Ws Fast Filter TC (sec)", mgvfWsFilterFast);
        mobjGlobalVars.AddSingle("DeltaP ", mipDeltaP);
        mobjGlobalVars.AddSingle("Filtered DeltaP (Pa)", mopFilDeltaP);
        mobjGlobalVars.AddSingle("DeltaP Plenum", mipDeltaPPlenum);
        mobjGlobalVars.AddSingle("Filtered DeltaP Plenum (Pa)", mopFilDeltaPPlenum);
        mobjGlobalVars.AddBoolean("Filter Input for WS calc", mccbFilterInput);
        mobjGlobalVars.AddBoolean("Separate Filters for Display and Control", mgvbSeparateControl);
        mobjGlobalVars.AddSingle("Control Filter TC", mgvfControlFilTc);
        mobjGlobalVars.AddSingle("Filtered DeltaP (Control) (Pa)", mgvfFilDeltaPControl);
        mobjGlobalVars.AddSingle("Kp (Static Pressure Coefficient)", mopKp);
        mobjGlobalVars.AddSingle("Kq plenum (Dynamic Pressure Coefficient)", mopKqpl);
        mobjGlobalVars.AddSingle("Kq nozzle (Dynamic Pressure Coefficient)", mopKqn);
        mobjGlobalVars.AddSingle("Q Plenum", mopQplPa);
        mobjGlobalVars.AddSingle("Q Nozzle", mopQnPa);
        mobjGlobalVars.AddSingle("WS Calc Method (1=Plenum, 0=Nozzle)", mopWSCalcMethod);
        mobjGlobalVars.AddSingle("Q", mopQPa);
        mobjGlobalVars.AddSingle("Q Control", mgvfQControlPa);
        mobjGlobalVars.AddSingle("Windspeed", mopWindspeed);
        mobjGlobalVars.AddSingle("Windspeed (Control)", mopWsControl);
        mobjGlobalVars.AddSingle("Windspeed Target (kph) (Control)", mgvfWsTarget);
        mobjGlobalVars.AddSingle("Limited Windspeed Target (kph) (Control)", mgvfWSTargetLimited);
        mobjGlobalVars.AddSingle("Windspeed Target (kph) (Displayed)", mopWsTargetDisp);
        mobjGlobalVars.AddSingle("Windspeed Rate in use (kph/s)", mgvfControlRamp);
        mobjGlobalVars.AddSingle("Windspeed Rate (kph/s)", mgvfWsRamp);
        mobjGlobalVars.AddSingle("Windspeed Setpoint (kph)", mopWsSetpt);
        mobjGlobalVars.AddSingle("Windspeed Increment", mgvfWsIncrement);
        mobjGlobalVars.AddSingle("Max Windspeed Rate For I (when under PID)", mgvfWsRateForI);
        mobjGlobalVars.AddSingle("Close Time (sec)", mgvfCloseTime);
        mobjGlobalVars.AddSingle("Close Time Ramp (kph/s)", mgvfCloseTimeRamp);
        mobjGlobalVars.AddSingle("Close Time Filter TC (sec)", mgvfCloseWsFilTc);
        mobjGlobalVars.AddSingle("Nominal Max Windspeed (kph)", mgvfNomWsMax);
        mobjGlobalVars.AddSingle("Normalized Windspeed Setpoint (kph)", mgvfNormWsSetpt);
        mobjGlobalVars.AddLong("Final Setpoint Ramp Method (0=None,1=CloseTime,2=FinalSnap)", mcclSetpointRampMethod);
        mobjGlobalVars.AddBoolean("Final Ramp - Can be used", mgvbFinalRampOKToUse);
        mobjGlobalVars.AddBoolean("Final Snap - In effect", mgvbUsingFinalRamp);
        mobjGlobalVars.AddBoolean("Final Snap - Coming back to Setpt", mgvbWSComingBackToSetpt);
        mobjGlobalVars.AddSingle("FF Norm Fanspeed (raw)", mgvfFfRawNormFs);
        mobjGlobalVars.AddSingle("FF Lead Coeff (sec)", mgvfFfLeadCoeff);
        mobjGlobalVars.AddSingle("FF Lag Coeff (sec)", mgvfFfLagCoeff);
        mobjGlobalVars.AddSingle("FF Norm Fanspeed (after LL)", mgvfFfNormFs);
        mobjGlobalVars.AddSingle("Windspeed Target Error", mgvfWSTargetError);
        mobjGlobalVars.AddSingle("Normalized Windspeed TargetError", mgvfTargetNormError);
        mobjGlobalVars.AddSingle("Windspeed Error", mgvfWsError);
        mobjGlobalVars.AddSingle("Normalized Windspeed Error", mgvfNormError);
        mobjGlobalVars.AddSingle("PID Kp", mgvfKp);
        mobjGlobalVars.AddSingle("PID Ti", mgvfTi);
        mobjGlobalVars.AddSingle("PID Td", mgvfTd);
        mobjGlobalVars.AddSingle("Kp Effect", mgvfKpEffect);
        mobjGlobalVars.AddSingle("Integrator Effect", mgvfIntegratorEffect);
        mobjGlobalVars.AddSingle("Old Integral Effect", mgvfOldIntegral);
        mobjGlobalVars.AddSingle("Deriviative Effect", mgvfDeriviativeEffect);
        mobjGlobalVars.AddSingle("PID Output (Norm FS)", mgvfFbNormFs);
        mobjGlobalVars.AddSingle("Controller Output (FB+FF, norm)", mgvfFsNormSetpt);
        mobjGlobalVars.AddSingle("Nominal Max Fanspeed (rpm)", mgvfNomFsMax);
        mobjGlobalVars.AddSingle("Fanspeed Target (rpm)", mgvfFsTarget);
        mobjGlobalVars.AddSingle("Limited Target (rpm)", mgvfFsLimitedTarget);
        mobjGlobalVars.AddSingle("Fanspeed Setpoint (rpm)", mopFsSetpt);
        mobjGlobalVars.AddSingle("Kp 0", mgvfKp0);
        mobjGlobalVars.AddSingle("Kp 1", mgvfKp1);
        mobjGlobalVars.AddSingle("Kp Max", mgvfMaxKp);
        mobjGlobalVars.AddSingle("Kp Min", mgvfMinKp);
        mobjGlobalVars.AddSingle("Ti 0", mgvfTi0);
        mobjGlobalVars.AddSingle("Ti 1", mgvfTi1);
        mobjGlobalVars.AddSingle("Ti Max", mgvfMaxTi);
        mobjGlobalVars.AddSingle("Ti Min", mgvfMinTi);
        mobjGlobalVars.AddSingle("Td 0", mgvfTd0);
        mobjGlobalVars.AddSingle("Td 1", mgvfTd1);
        mobjGlobalVars.AddSingle("Td Max", mgvfMaxTd);
        mobjGlobalVars.AddSingle("Td Min", mgvfMinTd);
        mobjGlobalVars.AddBoolean("WS Too low for PID", mgvbWsTooLowForPID);
        mobjGlobalVars.AddDate("WS too low for PID Start Time", mgvdWsTooLowForPIDTime);
        if (WindTunnel == wtCWT)
        {
            mobjGlobalVars.AddSingle("Idle Status (0=Stopped, 1=Started, 2=Moving)", mopIdleStatus);
            mobjGlobalVars.AddBoolean("Auto Idle Active", mgvbAutoIdle);
            mobjGlobalVars.AddBoolean("Auto Idle Speed has been high", mgvbAutoIdleSpeedHasBeenHigh);
            mobjGlobalVars.AddBoolean("Auto Idle Speed has been Low", mgvbAutoIdleSpeedHasBeenLow);
            mobjGlobalVars.AddBoolean("Auto Idle Stop Idle Overrride Active", mgvbStopAutoIdle);
        }
        else
        {
            mobjGlobalVars.AddBoolean("Control Ramp Held because speed was low", mgvbHoldCRUntilFSOK);
        }

        mobjGlobalVars.AddDate("Last time of non-zero FS target", mgvdLastTimeFSTargetNotZero);
        mobjGlobalVars.AddDate("Last time of non-zero WS target", mgvdLastTimeWSTargetNotZero);
        mobjGlobalVars.AddSingle("mgvfPurgeStatus", mgvfPurgeStatus);
        mobjGlobalVars.AddSingle("mgvfHeadwind", mgvfHeadwind);
        mobjGlobalVars.AddSingle("mgvfControlRamp", mgvfControlRamp);
        mobjGlobalVars.AddSingle("mgvfCloseTime", mgvfCloseTime);
        mobjGlobalVars.AddSingle("mgvfFilHxDeltaP", mgvfFilHxDeltaP);
        mobjGlobalVars.AddSingle("mgvfFsIncrement", mgvfFsIncrement);
        mobjGlobalVars.AddSingle("mgvfHxDpOverDp", mgvfHxDpOverDp);
        mobjGlobalVars.AddSingle("mgvfMinTi", mgvfMinTi);
        mobjGlobalVars.AddSingle("mgvfMaxTi", mgvfMaxTi);
        mobjGlobalVars.AddSingle("mgvfMinBlockage", mgvfMinBlockage);
        mobjGlobalVars.AddSingle("mgvfMaxBlockage", mgvfMaxBlockage);
        mobjGlobalVars.AddSingle("mgvfMaxFs", mgvfMaxFs);
        mobjGlobalVars.AddSingle("mgvfMaxFsTarget", mgvfMaxFsTarget);
        mobjGlobalVars.AddSingle("mgvfMaxWs", mgvfMaxWs);
        mobjGlobalVars.AddSingle("mgvfMfStartDelay", mgvfMfStartDelay);
        mobjGlobalVars.AddSingle("mgvfNormWsSetpt", mgvfNormWsSetpt);
        mobjGlobalVars.AddSingle("mgvfOldRemoteStatus", mgvfOldRemoteStatus);
        mobjGlobalVars.AddSingle("mgvfOldCtlMode", mgvfOldCtlMode);
        mobjGlobalVars.AddSingle("mgvfOldMfStatus", mgvfOldMfStatus);
        mobjGlobalVars.AddSingle("mgvfOldAuxStatus", mgvfOldAuxStatus);
        mobjGlobalVars.AddLong("mgvlOldNozzle", mgvlOldNozzle);
        mobjGlobalVars.AddSingle("mgvfOldPurgeStopFanRequest", mgvfOldPurgeStopFanRequest);
        mobjGlobalVars.AddSingle("mgvfOldTsDoor4", mgvfOldTsDoor4);
        mobjGlobalVars.AddSingle("mgvfOldWsIncrement", mgvfOldWsIncrement);
        mobjGlobalVars.AddSingle("mgvfPressureRatio", mgvfPressureRatio);
        mobjGlobalVars.AddSingle("mgvfPstaticControlKpa", mgvfPstaticControlKpa);
        mobjGlobalVars.AddSingle("mgvfQControlPa", mgvfQControlPa);
        mobjGlobalVars.AddSingle("mgvfRawHxDpOverDp", mgvfRawHxDpOverDp);
        mobjGlobalVars.AddSingle("mgvfWsError", mgvfWsError);
        // On Error GoTo Error — nested handler not restructured
        mobjStatuses = new Statuses();
        mobjStatuses.UseSPG(SPG_STATUSES);
        mobjStatuses.Add(STS_WINDSPEED_INITIALIZED, false, LoadResString(ridWindspeedNotInitialized), LoadResString(ridWindspeedInitialized));
        mobjStatuses.Add(STS_FANSPEED_ACHIEVED, false, LoadResString(ridFanspeedNotAchieved), LoadResString(ridFanspeedAchieved));
        mobjStatuses.Add(STS_WINDSPEED_ACHIEVED, false, LoadResString(ridWindspeedNotAchieved), LoadResString(ridWindspeedAchieved));
        mobjStatuses.Add(STS_MAINFAN_ON, false, "", LoadResString(ridMainFanOn));
        mobjStatuses.Add(STS_MAINFAN_OFF, false, "", LoadResString(ridMainfanOff));
        mobjStatuses.Add(STS_MAINFAN_AUX_ON, false, "", LoadResString(ridMainFanAuxOn));
        mobjStatuses.Add(STS_MAINFAN_AUX_OFF, false, "", LoadResString(ridMainFanAuxOff));
        if (WindTunnel == wtCWT)
        {
            mobjStatuses.Add(STS_IDLE_STARTED, false, "", LoadResString(ridIdleStarted));
            mobjStatuses.Add(STS_IDLE_STOPPED, true, "", LoadResString(ridIdleStopped));
        }

        objProcess = new Process();
        mbReady = true;
        WaitForProcessesStartupComplete();
        mlID = AddTalentProcessGetID(objProcess.Name.ToString(), objProcess);
        // On Error GoTo Error_Configurations — nested handler not restructured
        Pause(2);
        sClassName = "ParManager";
        mobjParManager = new ParManager();
        mobjParManager.OwnerID = mlID;
        mobjParManager.AutoRefreshEnabled = false;
        mobjParManager.IgnoreLoadedEvents = true;
        mobjParManager.IgnoreUnloadedEvents = true;
        mobjParManager.AddWithoutFlag("Airtemp", mipAirtemp, ioInput);
        mobjParManager.AddWithoutFlag("Dewpoint", mipDewpoint, ioInput);
        mobjParManager.AddWithoutFlag("RelHumidity", mipRelHumid, ioInput);
        mobjParManager.AddWithoutFlag("AbsPress", mipPabs, ioInput);
        mobjParManager.AddWithoutFlag("DeltaPNozzle", mipDeltaP, ioInput);
        mobjParManager.AddWithoutFlag("DeltaPPlenum", mipDeltaPPlenum, ioInput);
        mobjParManager.AddWithoutFlag("OpModeOff", mipOpModeOff, ioInput);
        mobjParManager.AddWithoutFlag("OpModeMaint", mipOpModeMaint, ioInput);
        mobjParManager.AddWithoutFlag("OpModePrep", mipOpModePrep, ioInput);
        if (mgvlWindTunnel == wtAAWT)
        {
            mobjParManager.AddWithoutFlag("OpModeStab", mipOpModeStab, ioInput);
        }

        mobjParManager.AddWithoutFlag("OpModeTest", mipOpModeTest, ioInput);
        mobjParManager.AddWithoutFlag("WTCS_AuxCmdRequest", mipAuxDPLCCmdRequest, ioInput);
        mobjParManager.AddWithoutFlag("WTCS_CSCmdRequest", mipCoolPLCCmdRequest, ioInput);
        mobjParManager.AddWithoutFlag("PLC_FaultReset", mipToPLCResetCmd, ioInput);
        mobjParManager.AddWithoutFlag("MF_Remote", mipMfRemote, ioInput);
        mobjParManager.AddWithoutFlag("MF_Permissive", mipMfPerm, ioInput);
        mobjParManager.AddWithoutFlag("PLCEstopOut", mipEstopStatus, ioInput);
        mobjParManager.AddWithoutFlag("MF_MCFault", mipMfMCFault, ioInput);
        mobjParManager.AddWithoutFlag("MF_MCReady", mipMCReady, ioInput);
        mobjParManager.AddWithoutFlag("MF_LCFault", mipMfLCFault, ioInput);
        mobjParManager.AddWithoutFlag("MF_LCReady", mipLCReady, ioInput);
        mobjParManager.AddWithoutFlag("MF_Status", mipMfStatus, ioInput);
        mobjParManager.AddWithoutFlag("MF_Speed", mipFanspeed, ioInput);
        mobjParManager.AddWithoutFlag("MF_CSReady", mipAuxReady, ioInput);
        mobjParManager.AddWithoutFlag("MF_CSFault", mipAuxFault, ioInput);
        mobjParManager.AddWithoutFlag("MF_CSStatus", mipCSStatus, ioInput);
        if (WindTunnel == wtCWT)
        {
            mobjParManager.AddWithoutFlag("NozTopLimSw", mipNozzleLargeLimSw, ioInput);
            mobjParManager.AddWithoutFlag("NozMiddleLimSw", mipNozzleMediumLimSw, ioInput);
            mobjParManager.AddWithoutFlag("NozBottomLimSw", mipNozzleSmallLimSw, ioInput);
            mobjParManager.AddWithoutFlag("Idle_MotorStop", mipIdleNozFlapMotorStop, ioInput);
            mobjParManager.AddWithoutFlag("Idle_FlapOpen", mipIdleNozFlapOpenStatus, ioInput);
            mobjParManager.AddWithoutFlag("Idle_FlapClosed", mipIdleNozFlapClosedStatus, ioInput);
            mobjParManager.AddWithoutFlag("Idle_MotorFault", mipIdleNozFlapMotorFault, ioInput);
            for (I = 1; I <= 3; I++)
            {
                mobjParManager.AddWithoutFlag(string.Concat(string.Concat("BypFlap", I), "Closed"), mipIdleBypFlapClosedStatus(I), ioInput);
                mobjParManager.AddWithoutFlag(string.Concat(string.Concat("BypFlap", I), "Open"), mipIdleBypFlapOpenStatus(I), ioInput);
            }

            mobjParManager.AddWithoutFlag("DynoSpeedTgt", mipDySpeedTarget, ioInput);
            mobjParManager.AddWithoutFlag("DynoSpeedRate", mipDySpeedRate, ioInput);
            mobjParManager.AddWithoutFlag("DynoMode", mipDyMode, ioInput);
            mobjParManager.AddWithoutFlag("ThrMode", mipThrMode, ioInput);
            mobjParManager.AddWithoutFlag("CF_LeftPos1", mipCFLeft1Pos, ioInput);
            mobjParManager.AddWithoutFlag("CF_Left1MotorFault", mipCFLeft1Fault, ioInput);
            mobjParManager.AddWithoutFlag("CF_Left1MotorLockout", mipCFLeft1Lockout, ioInput);
            mobjParManager.AddWithoutFlag("CF_LeftPos2", mipCFLeft2Pos, ioInput);
            mobjParManager.AddWithoutFlag("CF_Left2MotorFault", mipCFLeft2Fault, ioInput);
            mobjParManager.AddWithoutFlag("CF_Left2MotorLockout", mipCFLeft2Lockout, ioInput);
            mobjParManager.AddWithoutFlag("CF_RightPos1", mipCFRight1Pos, ioInput);
            mobjParManager.AddWithoutFlag("CF_Right1MotorFault", mipCFRight1Fault, ioInput);
            mobjParManager.AddWithoutFlag("CF_Right1MotorLockout", mipCFRight1Lockout, ioInput);
            mobjParManager.AddWithoutFlag("CF_RightPos2", mipCFRight2Pos, ioInput);
            mobjParManager.AddWithoutFlag("CF_Right2MotorFault", mipCFRight2Fault, ioInput);
            mobjParManager.AddWithoutFlag("CF_Right2MotorLockout", mipCFRight2Lockout, ioInput);
            mobjParManager.AddWithoutFlag("CF_TopPos", mipCFTopPos, ioInput);
            mobjParManager.AddWithoutFlag("CF_TopMotorFault", mipCFTopFault, ioInput);
            mobjParManager.AddWithoutFlag("CF_TopMotorLockout", mipCFTopLockout, ioInput);
        }

        mobjParManager.AddWithoutFlag("AirDensity", mopAirDensityKgm3, ioOutput);
        mobjParManager.AddWithoutFlag("PStatic", mopPstaticKpa, ioOutput);
        mobjParManager.AddWithoutFlag("Q", mopQPa, ioOutput);
        mobjParManager.AddWithoutFlag("WS_CalcMethod", mopWSCalcMethod, ioOutput);
        mobjParManager.AddWithoutFlag("Kp", mopKp, ioOutput);
        mobjParManager.AddWithoutFlag("Kqpl", mopKqpl, ioOutput);
        mobjParManager.AddWithoutFlag("Kqn", mopKqn, ioOutput);
        mobjParManager.AddWithoutFlag("MF_AuxStartCmd", mopAuxControl, ioOutput);
        mobjParManager.AddWithoutFlag("MF_StartCmd", mopMfControl, ioOutput);
        mobjParManager.AddWithoutFlag("MF_ResetCmd", mopMFReset, ioOutput);
        mobjParManager.AddWithoutFlag("MF_SpeedSetpt", mopFsSetpt, ioOutput);
        mobjParManager.AddWithoutFlag("MF_SpeedRate", mopFsRamp, ioOutput);
        mobjParManager.AddWithoutFlag("MF_SpeedTgt", mopFsTargetDisp, ioOutput);
        mobjParManager.AddWithoutFlag("Windspeed", mopWindspeed, ioOutput);
        mobjParManager.AddWithoutFlag("WS_Setpt", mopWsSetptDisp, ioOutput);
        mobjParManager.AddWithoutFlag("WS_Rate", mopWsRamp, ioOutput);
        mobjParManager.AddWithoutFlag("WS_Target", mopWsTargetDisp, ioOutput);
        mobjParManager.AddWithoutFlag("WS_Headwind", mopHeadwind, ioOutput);
        mobjParManager.AddWithoutFlag("WS_Mode", mopWsMode, ioOutput);
        if (WindTunnel == wtCWT)
        {
            mobjParManager.AddWithoutFlag("NozzleSize", mopNozzle, ioOutput);
            mobjParManager.AddWithoutFlag("AutoIdleStatus", mopAutoIdleStatus, ioOutput);
            mobjParManager.AddWithoutFlag("Idle_Status", mopIdleStatus, ioOutput);
            mobjParManager.AddWithoutFlag("Idle_StopCmd", mopIdleStopControl, ioOutput);
            mobjParManager.AddWithoutFlag("Idle_StartCmd", mopIdleStartControl, ioOutput);
            mobjParManager.AddWithoutFlag("Noz_Pos3Cmd", mopNozzleCmd(nsSmall), ioOutput);
            mobjParManager.AddWithoutFlag("Noz_Pos2Cmd", mopNozzleCmd(nsMedium), ioOutput);
            mobjParManager.AddWithoutFlag("Noz_Pos1Cmd", mopNozzleCmd(nsLarge), ioOutput);
            mobjParManager.AddWithoutFlag("PurgeStartCmd", mopPurgeCmd, ioOutput);
            mobjParManager.AddWithoutFlag("CF_LeftStartCmd", mopCFLeftStartCmd, ioOutput);
            mobjParManager.AddWithoutFlag("CF_LeftStopCmd", mopCFLeftStopCmd, ioOutput);
            mobjParManager.AddWithoutFlag("CF_LeftPos1SP", mopCFLeft1Target, ioOutput);
            mobjParManager.AddWithoutFlag("CF_LeftPos2SP", mopCFLeft2Target, ioOutput);
            mobjParManager.AddWithoutFlag("CF_RightStartCmd", mopCFRightStartCmd, ioOutput);
            mobjParManager.AddWithoutFlag("CF_RightStopCmd", mopCFRightStopCmd, ioOutput);
            mobjParManager.AddWithoutFlag("CF_RightPos1SP", mopCFRight1Target, ioOutput);
            mobjParManager.AddWithoutFlag("CF_RightPos2SP", mopCFRight2Target, ioOutput);
            mobjParManager.AddWithoutFlag("CF_TopStartCmd", mopCFTopStartCmd, ioOutput);
            mobjParManager.AddWithoutFlag("CF_TopStopCmd", mopCFTopStopCmd, ioOutput);
            mobjParManager.AddWithoutFlag("CF_TopPosSP", mopCFTopTarget, ioOutput);
        }

        if (WindTunnel == wtAAWT)
        {
            mobjParManager.AddWithoutFlag("Windspeed_ms", mopWindspeedmps, ioOutput);
            mobjParManager.AddWithoutFlag("WS_Setpt_ms", mopWsSetptmps, ioOutput);
            mobjParManager.AddWithoutFlag("WS_Rate_ms", mopWsRampmps, ioOutput);
            mobjParManager.AddWithoutFlag("WS_Target_ms", mopWsTargetmps, ioOutput);
            mobjParManager.AddWithoutFlag("WS_Head_ms", mopHeadwindmps, ioOutput);
            mobjParManager.AddWithoutFlag("PurgeCmd", mopPurgeCmd, ioOutput);
            mobjParManager.AddWithoutFlag("WS_Achieved", mopWsAchieved, ioOutput);
            for (I = 1; I <= 18; I++)
            {
                mobjParManager.AddWithoutFlag(string.Concat(string.Concat("OF_", Format(I, "00")), "OpenCmd"), mopOFOpenCmd(I), ioOutput);
            }
        }

        mobjMFPermComponents.Initialize(mobjParManager.Configurations);
        mobjMFAuxPermComponents.Initialize(mobjParManager.Configurations);
        mobjIdleCmdPermComponents.Initialize(mobjParManager.Configurations);
        // On Error GoTo Error — nested handler not restructured
        if (mccfUpdateInt < 0.01)
        {
            mccfUpdateInt = 0.01;
        }

        if (IsWithinIDE())
        {
            mobjForm = new frmControl();
            Load(mobjForm);
            mobjForm.tmrDesignTime.Interval = mccfUpdateInt * 1000;
            mobjForm.tmrDesignTime.Enabled = true;
        }
        else
        {
            lTimerFreq = mccfUpdateInt * 1000 / 10;
            lTimerFreq = Reg_GetValue(HKEY_LOCAL_MACHINE, REG_PROCESSES, null, "TimerFrequency", lTimerFreq, vbLong);
            lTimerFreq = Reg_GetValue(HKEY_LOCAL_MACHINE, REG_PROCESSES, PROCESS_NAME, "TimerFrequency", lTimerFreq, vbLong);
            if (lTimerFreq > mccfUpdateInt * 1000)
            {
                lTimerFreq = mccfUpdateInt * 1000;
            }

            tmrMain.Stats.Frequency = lTimerFreq;
            tmrMain.Interval = mccfUpdateInt * 1000;
            tmrMain.EventType = TimerPeriodic;
            tmrMain.Notify = this;
            SetProcessPriority(mgvlDefPriority);
            tmrMain.Enabled = true;
        }

        if (mobjParManager.Configurations.Count > 0)
        {
            mobjCommands.Append(ocInitialize, "Initialize");
        }

        tmrRunOnce = null;
        mobjStats.Reset();
        for (I = 1; I <= 200; I++)
        {
            mobjStats.AddValue((double)mccfUpdateInt);
        }

        tmrElapsed.Reset();
        mgvlOldNozzle = -1;
        mgvfOldCtlMode = -1;
        goto Done;
        Error();
        sRaiseErrorText = ex.Message;
        EH.VBErr = Err;
        EH.AppError = StdTalentStrings.stsEHErrorDuringControllerStartup;
        EH.Module = TypeName(this);
        EH.Procedure = "Start";
        EH.Item(1) = GetComputerName;
        EH_MsgBox();
    Error_CreatingObjects:
        ;
        sRaiseErrorText = LoadResString(ridCantCreateObject, sClassName, ex.Message);
        bRaiseError = true;
        EH.VBErr = Err;
        EH.Module = TypeName(this);
        EH.Procedure = "Start";
        EH.Item(1) = sClassName;
        EH_MsgBox();
        goto Done;
    Error_Configurations:
        ;
        EH.VBErr = Err;
        EH.AppError = LoadResString(StdTalentStrings.stsEHCantCreateConfigurations);
        EH.Module = TypeName(this);
        EH.Procedure = PROC_NAME;
        sRaiseErrorText = EH.Text;
        EH_MsgBox();
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        objSPGM = null;
        if (bRaiseError)
        {
            EH_RaiseError(OLEERR_METHOD_FAILED, TypeName(this), "Start", sRaiseErrorText)();
        }

        return;
    }

    internal void ControlLogic()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        if (mbInited == true)
        {
            mobjParManager.ReadInputs();
            InputConversions();
            CalcNozzleModeDependentConstants();
            CalcWindspeed();
            CheckInterlocks();
            bInputsRead = true;
        }

        ExecuteCommands();
        if (mbInited)
        {
            if (bInputsRead == false)
            {
                mobjParManager.ReadInputs();
                InputConversions();
                CalcNozzleModeDependentConstants();
                CalcWindspeed();
                CheckInterlocks();
            }

            WSControlLogic();
            MonitorStatuses();
            OutputConversions();
            SaveOldStates();
            mobjParManager.WriteOutputs();
        }

        if (!mobjStatuses == null)
        {
            mobjStatuses.FlushMessages();
        }

        goto Done;
        Error();
        if (mbInited)
        {
            EH.VBErr = Err;
            EH.AppError = stsEHErrorExecutingControlLogic;
            EH.Module = TypeName(this);
            EH.Procedure = PROC_NAME;
            EH_MsgBox();
        }

        goto ErrorExit;
    ErrorExit:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        Suspend();
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        return;
    }

    private void ExecuteCommands()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        if (ReEntered)
        {
            return;
        }

        ReEntered = true;
        if (mobjCommands.Count == 0)
        {
            goto Done;
        }

        objCommand = mobjCommands(1);
        switch (objCommand.Opcode)
        {
            case ocInitialize:
                this.Initialize(,);
                objCommand();
                objCommand.Remove();
                break;
            case ocSuspend:
                if (objCommand.FirstTime & mbInited & mopMfControl != 0 | mipMfStatus != 0)
                {
                    mobjCommands.Insert(ocStopMainFan, "StopMainFan");
                }
                else
                {
                    objCommand.Messages.Flush();
                    Suspend();
                    objCommand.Remove();
                    if (IsWithinIDE())
                    {
                        Shutdown();
                    }
                }

                break;
            case ocShutdown:
                if (objCommand.FirstTime & mbInited)
                {
                    mobjCommands.Insert(ocSuspend, "SuspendWindspeed");
                }
                else
                {
                    objCommand.Messages.Flush();
                    objCommand.Remove();
                    this.Shutdown();
                }

                break;
            default:
                CheckInterlock(ocInitialize.ToString());
                switch (objCommand.Opcode)
                {
                    case ocSetModeManual:
                        CheckInterlock(ocSetModeManual.ToString());
                        SetMode(modeManual);
                        objCommand.Remove();
                        break;
                    case ocSetModeAuto:
                        CheckInterlock(ocSetModeAuto.ToString());
                        if (mopIdleStatus != isNormal | mopIdleStartControl == 1)
                        {
                            objCommand.InsertPredecessor(ocStopIdle, "StopIdle");
                        }
                        else
                        {
                            SetMode(modeAuto);
                            objCommand.Remove();
                        }

                        break;
                    case ocResetFaults:
                        CheckInterlock(ocResetFaults.ToString());
                        if (objCommand.FirstTime)
                        {
                            objCommand.ResetTime();
                            mopMFReset = 1;
                        }
                        else if (objCommand.TimeSince >= mccfResetFaultTime)
                        {
                            mopMFReset = 0;
                            objCommand.Remove();
                        }

                        break;
                    case ocStartMainFanAux:
                        CheckInterlock(ocStartMainFanAux.ToString());
                        mgvdTurnOffFanTime = DateTime.Now;
                        objCommand.Messages.Flush();
                        mgvbAuxAutoOff = false;
                        mopAuxControl = 1;
                        if (objCommand.FirstTime)
                        {
                            objCommand.ResetTime();
                        }

                        if (mgvfAuxStatus == 1)
                        {
                            if (mgvdAuxOnTime == 0)
                            {
                            }
                            else if (SecondsSinceTime(mgvdAuxOnTime) >= mccfMainFanAuxPause)
                            {
                                objCommand.Remove();
                            }
                        }
                        else if (objCommand.TimeSince >= mccfMainFanAuxTimeout)
                        {
                            mopAuxControl = 0;
                            sText = "";
                            if (mipCSStatus != 1)
                            {
                                sText = string.Concat(IIf(sText != "", string.Concat(string.Concat(sText, vbCrLf), "- "), "- "), LoadResString(ridCSDidNotStart));
                            }

                            objCommand.Abort(RaiseErrorText(tceDeviceTimedOut, objCommand.Text, sText), cmdTimedout);
                        }

                        break;
                    case ocStopMainFanAux:
                        CheckInterlock(ocStopMainFanAux.ToString());
                        if (mopWsMode > modeManual)
                        {
                            objCommand.InsertPredecessor(ocSetModeManual, "SetModeManual");
                        }
                        else if (mopMfControl != 0 | mipMfStatus != 0)
                        {
                            objCommand.InsertPredecessor(ocStopMainFan, "StopMainFan");
                        }
                        else
                        {
                            objCommand.Messages.Flush();
                            mopAuxControl = 0;
                            mgvfFsTarget = 0;
                            mopFsSetpt = 0;
                            mgvfWsTarget = 0;
                            mopWsSetpt = 0;
                            if (mgvfAuxStatus == 0 | objCommand.TimeSince >= 30)
                            {
                                objCommand.Remove();
                            }
                        }

                        break;
                    case ocStartMainFan:
                        CheckInterlock(ocStartMainFan.ToString());
                        mgvdTurnOffFanTime = DateTime.Now;
                        if (mopWsMode < modeManual)
                        {
                            objCommand.InsertPredecessor(ocSetModeManual, "SetModeManual");
                        }
                        else if (mgvfAuxStatus != 1 | mopAuxControl != 1)
                        {
                            objCommand.InsertPredecessor(ocStartMainFanAux, "StartMainFanAux");
                        }
                        else
                        {
                            objCommand.Messages.Flush();
                            mopMfControl = 1;
                            if (mipMfStatus != 1)
                            {
                                mgvfFsTarget = 0;
                            }

                            if (mgvfFsTarget < mccfMinFs & mccfMinFs > 0)
                            {
                                mgvfFsTarget = mccfMinFs;
                                mopFsSetpt = mccfMinFs;
                            }

                            if (objCommand.FirstTime)
                            {
                                objCommand.ResetTime();
                            }

                            if (mipMfStatus == 1)
                            {
                                if (mccfDrivePause <= 0)
                                {
                                    objCommand.Remove();
                                }
                                else if (mgvdDriveOnTime == 0)
                                {
                                }
                                else if (SecondsSinceTime(mgvdDriveOnTime) >= mccfDrivePause)
                                {
                                    objCommand.Remove();
                                }
                            }
                            else if (objCommand.TimeSince >= mccfMainFanTimeout)
                            {
                                mopMfControl = 0;
                                objCommand.Abort(,);
                                cmdTimedout();
                            }
                        }

                        break;
                    case ocStopMainFan:
                        CheckInterlock(ocStopMainFan.ToString());
                        objCommand.Messages.Flush();
                        if (mopWsMode != modeManual)
                        {
                            objCommand.InsertPredecessor(ocSetModeManual, "SetModeManual");
                        }
                        else if (mgvfFsTarget != 0)
                        {
                            objCommand.InsertPredecessor(ocSetFanSpeed, "SetFanSpeed", 0);
                            objCommand.ResetTime();
                        }
                        else if (mopFsSetpt != 0 & objCommand.TimeSince < 60)
                        {
                            objCommand.TagVariant = DateTime.Now;
                        }
                        else if (Math.Abs(mipFanspeed) > 5 & objCommand.TimeSince < 30)
                        {
                            objCommand.TagVariant = DateTime.Now;
                        }
                        else
                        {
                            mopMfControl = 0;
                            if (mccfMainFanAuxAutoOffTime > 0)
                            {
                                mgvbAuxAutoOff = true;
                            }// On Error Resume Next — nested handler; see enclosing try/catch
                            if (mipMfStatus != 0)
                            {
                                if (objCommand.TimeSince >= 60)
                                {
                                    objCommand.Remove();
                                }
                                else
                                {
                                    objCommand.TagVariant = DateTime.Now;
                                }
                            }
                            else if (SecondsSinceTime(objCommand.TagVariant) >= 5)
                            {
                                objCommand.Remove();
                            }
                        }

                        break;
                    case ocSetWindSpeed:
                        CheckInterlock(Array(ocSetWindSpeed.ToString(), objCommand.Arguments(1).Value));
                        mopWsTargetDisp = objCommand.Arguments(1).Value;
                        mopWsRamp = objCommand.Arguments(2).Value;
                        if (mipMfStatus != 1 | mopMfControl != 1)
                        {
                            objCommand.InsertPredecessor(ocStartMainFan, "StartMainFan");
                        }
                        else if (mopWsMode != modeAuto)
                        {
                            objCommand.InsertPredecessor(ocSetModeAuto, "SetModeAuto");
                        }
                        else
                        {
                            objCommand.Messages.Flush();
                            mgvfWsTarget = objCommand.Arguments(1).Value;
                            mgvfWsRamp = objCommand.Arguments(2).Value;
                            if (mgvfWsRamp <= 0)
                            {
                                mgvfWsRamp = mccfDefWsRamp;
                            }

                            if (mcclSetpointRampMethod == 1)
                            {
                                mobjRampRateFilter.ClearData();
                                mgvfControlRamp = mobjRampRateFilter.FilteredValue(mgvfWsRamp);
                            }

                            if (Math.Abs(mgvfWsTarget - mopWindspeed) <= mccfFinalRampDeltakph + 1)
                            {
                                mgvbFinalRampOKToUse = false;
                            }
                            else
                            {
                                mgvbFinalRampOKToUse = true;
                            }

                            mgvbFinalRampHasSnapped = false;
                            mgvbUsingFinalRamp = false;
                            mgvbHoldCRUntilFSOK = false;
                            objCommand.Remove();
                        }

                        break;
                    case ocTrackWindSpeed:
                        CheckInterlock(ocTrackWindSpeed.ToString());
                        if (mipMfStatus != 1 | mopMfControl != 1)
                        {
                            objCommand.InsertPredecessor(ocStartMainFan, "StartMainFan");
                        }
                        else if (objCommand.Arguments(3).Value == false & mopIdleStatus != isNormal)
                        {
                            objCommand.InsertPredecessor(ocStopIdle, "StopIdle");
                        }
                        else
                        {
                            objCommand.Messages.Flush();
                            mobjTrackPar = mobjParManager.Configurations.Parameters(objCommand.Arguments(1));
                            GetConversionCoefficients(mobjTrackPar.Units, "kph", mgvfTrackC1, mgvfTrackC0);
                            SetMode(modeTrack);
                            mgvsTrackParName = objCommand.Arguments(1).Value;
                            mgvfHeadwind = objCommand.Arguments(2).Value;
                            if (WindTunnel == wtCWT)
                            {
                                mgvbAutoIdle = objCommand.Arguments(3).Value;
                                mgvdLastIdleTime = Now;
                                mgvbStopAutoIdle = false;
                            }

                            mgvfWsRamp = mccfMaxWsRamp;
                            mgvfControlRamp = mgvfWsRamp;
                            objCommand.Remove();
                        }

                        break;
                    case ocSetFanSpeed:
                        CheckInterlock(ocSetFanSpeed.ToString());
                        mopFsTargetDisp = objCommand.Arguments(1).Value;
                        if (objCommand.Arguments.Count >= 2)
                        {
                            mopFsRamp = objCommand.Arguments(2).Value;
                        }

                        if (mopWsMode != modeManual)
                        {
                            objCommand.InsertPredecessor(ocSetModeManual, "SetModeManual");
                        }
                        else if (mipMfStatus != 1 | mopMfControl != 1)
                        {
                            objCommand.InsertPredecessor(ocStartMainFan, "StartMainFan");
                        }
                        else if (mopFsTargetDisp > mccfMaxFsLimit(1) & LimitInEffect(1))
                        {
                            objCommand.InsertPredecessor(ocStopIdle, "StopIdle");
                        }
                        else
                        {
                            objCommand.Messages.Flush();
                            mgvfFsTarget = objCommand.Arguments(1).Value;
                            if (objCommand.Arguments.Count >= 2)
                            {
                                mgvfFsRamp = objCommand.Arguments(2).Value;
                            }
                            else
                            {
                                mgvfFsRamp = mccfDefFsRamp;
                            }

                            objCommand.Remove();
                        }

                        break;
                    case ocCancelIdle:
                        mopIdleStartControl = 0;
                        mopIdleStopControl = 0;
                        if (objCommand.Arguments.Count == 0)
                        {
                            objCommand.Remove();
                        }
                        else if (objCommand.TimeSince >= objCommand.Arguments(1).Value)
                        {
                            objCommand.Remove();
                        }

                        break;
                    case ocStartIdle:
                        CheckInterlock(ocStartIdle.ToString());
                        if (mipAuxDPLCCmdRequest == 0)
                        {
                            if (objCommand.FirstTime)
                            {
                                RequestControlFromAuxDPLC();
                            }
                            else if (objCommand.TimeSince >= 10)
                            {
                                objCommand.Abort(LoadResString(ridUnableToGetControlFromPLC, "AuxD"), cmdTimedout);
                            }
                        }
                        else if (mopWsMode > modeManual)
                        {
                            objCommand.InsertPredecessor(ocSetModeManual, "SetModeManual");
                        }
                        else if (Math.Abs(mgvfFsTarget) > mccfMaxFsLimit(1))
                        {
                            mobjCommands.Insert(ocSetFanSpeed, "SetFanspeed", mccfMaxFsLimit(1), mccfMaxFsRamp);
                        }
                        else if (mgvfFsTarget < 0)
                        {
                            mobjCommands.Insert(ocSetFanSpeed, "SetFanspeed", 0, mccfMaxFsRamp);
                        }
                        else if (Math.Abs(mopFsSetpt) > mccfMaxFsLimit(1) & objCommand.TimeSince <= 10)
                        {
                        }
                        else if (mopIdleStopControl == 1)
                        {
                            objCommand.InsertPredecessor(ocCancelIdle, "CancelIdle", mccfCancelIdleTime);
                        }
                        else
                        {
                            objCommand.Messages.Flush();
                            mopIdleStartControl = 1;
                            mopIdleStopControl = 0;
                            if (mopIdleStatus == isIdle)
                            {
                                if (mccfIdleFanspeed > 0)
                                {
                                    if (mobjInterlocks(ocSetFanSpeed.ToString()).IsOK)
                                    {
                                        if (mipMfStatus == 1 & mgvfFsTarget < mccfIdleFanspeed)
                                        {
                                            mobjCommands.Insert(ocSetFanSpeed, "SetFanspeed", mccfIdleFanspeed);
                                        }
                                    }
                                }

                                objCommand.Remove();
                            }
                            else if (objCommand.TimeSince > mccfIdleTimeout)
                            {
                                objCommand.Abort(,);
                                cmdTimedout();
                            }
                        }

                        break;
                    case ocStopIdle:
                        CheckInterlock(ocStopIdle.ToString());
                        if (mopIdleStatus == isNormal)
                        {
                            mopIdleStartControl = 0;
                            mopIdleStopControl = 0;
                            if (mopWsMode == modeTrack)
                            {
                                mgvbAutoIdle = false;
                            }

                            objCommand.Remove();
                        }
                        else if (mipAuxDPLCCmdRequest == 0)
                        {
                            if (objCommand.FirstTime)
                            {
                                RequestControlFromAuxDPLC();
                            }
                            else if (objCommand.TimeSince >= 10)
                            {
                                objCommand.Abort(LoadResString(ridUnableToGetControlFromPLC, "AuxD"), cmdTimedout);
                            }
                        }
                        else if (mopWsMode == modeTrack & mgvbAutoIdle == true)
                        {
                            mgvbStopAutoIdle = true;
                            objCommand.Remove();
                        }
                        else if (mopWsMode > modeManual)
                        {
                            objCommand.InsertPredecessor(ocSetModeManual, "SetModeManual");
                        }
                        else if (mgvfFsTarget > mccfMaxFsLimit(1))
                        {
                            mobjCommands.Insert(ocSetFanSpeed, "SetFanspeed", mccfMaxFsLimit(1), mccfMaxFsRamp);
                        }
                        else if (mopFsSetpt > mccfMaxFsLimit(1) & objCommand.TimeSince <= 10)
                        {
                        }
                        else if (mopIdleStartControl == 1)
                        {
                            objCommand.InsertPredecessor(ocCancelIdle, "CancelIdle", mccfCancelIdleTime);
                        }
                        else
                        {
                            objCommand.Messages.Flush();
                            mopIdleStartControl = 0;
                            mopIdleStopControl = 1;
                            if (objCommand.TimeSince > mccfIdleTimeout)
                            {
                                objCommand.Abort(,);
                                cmdTimedout();
                            }
                        }

                        break;
                    case ocStartPurge:
                        CheckInterlock(ocStartPurge.ToString());
                        if (WindTunnel == wtAAWT)
                        {
                            if (mipAuxDPLCCmdRequest == 0)
                            {
                                if (objCommand.FirstTime)
                                {
                                    RequestControlFromAuxDPLC();
                                }
                                else if (objCommand.TimeSince >= 10)
                                {
                                    objCommand.Abort(LoadResString(ridUnableToGetControlFromPLC, "AuxD"), cmdTimedout);
                                }
                            }
                            else
                            {
                                mopPurgeCmd = 1;
                                objCommand.Remove();
                            }
                        }
                        else
                        {
                            if (mipCoolPLCCmdRequest == 0)
                            {
                                if (objCommand.FirstTime)
                                {
                                    RequestControlFromCoolPLC();
                                }
                                else if (objCommand.TimeSince >= 10)
                                {
                                    objCommand.Abort(LoadResString(ridUnableToGetControlFromPLC, "Cooling"), cmdTimedout);
                                }
                            }
                            else
                            {
                                mopPurgeCmd = 1;
                                objCommand.Remove();
                            }
                        }

                        break;
                    case ocStopPurge:
                        CheckInterlock(ocStopPurge.ToString());
                        if (WindTunnel == wtAAWT)
                        {
                            if (mipAuxDPLCCmdRequest == 0)
                            {
                                if (objCommand.FirstTime)
                                {
                                    RequestControlFromAuxDPLC();
                                }
                                else if (objCommand.TimeSince >= 10)
                                {
                                    objCommand.Abort(LoadResString(ridUnableToGetControlFromPLC, "AuxD"), cmdTimedout);
                                }
                            }
                            else
                            {
                                mopPurgeCmd = 0;
                                objCommand.Remove();
                            }
                        }
                        else
                        {
                            if (mipCoolPLCCmdRequest == 0)
                            {
                                if (objCommand.FirstTime)
                                {
                                    RequestControlFromCoolPLC();
                                }
                                else if (objCommand.TimeSince >= 10)
                                {
                                    objCommand.Abort(LoadResString(ridUnableToGetControlFromPLC, "Cooling"), cmdTimedout);
                                }
                            }
                            else
                            {
                                mopPurgeCmd = 0;
                                objCommand.Remove();
                            }
                        }

                        break;
                    case ocSetNozzleMethod:
                        mopWSCalcMethod = 0;
                        objCommand.Remove();
                        break;
                    case ocSetPlenumMethod:
                        mopWSCalcMethod = 1;
                        objCommand.Remove();
                        break;
                    case ocSetNozzleSize:
                        CheckInterlock(ocSetNozzleSize.ToString());
                        if (mipAuxDPLCCmdRequest == 0)
                        {
                            if (objCommand.FirstTime)
                            {
                                RequestControlFromAuxDPLC();
                            }
                            else if (objCommand.TimeSince >= 10)
                            {
                                objCommand.Abort(LoadResString(ridUnableToGetControlFromPLC, "AuxD"), cmdTimedout);
                            }
                        }
                        else
                        {
                            mgvlNozzleTarget = objCommand.Arguments(1);
                            mgvdNozzleSizeCmdTime = Now;
                            objCommand.Remove();
                        }

                        break;
                    case ocStopNozzleMotion:
                        CheckInterlock(ocStopNozzleMotion.ToString());
                        if (mipAuxDPLCCmdRequest == 0)
                        {
                            if (objCommand.FirstTime)
                            {
                                RequestControlFromAuxDPLC();
                            }
                            else if (objCommand.TimeSince >= 10)
                            {
                                objCommand.Abort(LoadResString(ridUnableToGetControlFromPLC, "AuxD"), cmdTimedout);
                            }
                        }
                        else
                        {
                            mgvlNozzleTarget = 0;
                            objCommand.Remove();
                        }

                        break;
                    case ocSetCFlapLeft:
                        CheckInterlock(ocSetCFlapLeft.ToString());
                        if (mipAuxDPLCCmdRequest == 0)
                        {
                            if (objCommand.FirstTime)
                            {
                                RequestControlFromAuxDPLC();
                            }
                            else if (objCommand.TimeSince >= 10)
                            {
                                objCommand.Abort(LoadResString(ridUnableToGetControlFromPLC, "AuxD"), cmdTimedout);
                            }
                        }
                        else
                        {
                            mopCFLeft1Target = objCommand.Arguments(1);
                            mopCFLeft2Target = objCommand.Arguments(2);
                            mgvdCFLeftCmdTime = Now;
                            mgvbCFLeftCtl = true;
                            objCommand.Remove();
                        }

                        break;
                    case ocSetCFlapLeft1:
                        CheckInterlock(ocSetCFlapLeft1.ToString());
                        if (mipAuxDPLCCmdRequest == 0)
                        {
                            if (objCommand.FirstTime)
                            {
                                RequestControlFromAuxDPLC();
                            }
                            else if (objCommand.TimeSince >= 10)
                            {
                                objCommand.Abort(LoadResString(ridUnableToGetControlFromPLC, "AuxD"), cmdTimedout);
                            }
                        }
                        else
                        {
                            mopCFLeft1Target = objCommand.Arguments(1);
                            mgvdCFLeftCmdTime = Now;
                            mgvbCFLeftCtl = true;
                            objCommand.Remove();
                        }

                        break;
                    case ocSetCFlapLeft2:
                        CheckInterlock(ocSetCFlapLeft2.ToString());
                        if (mipAuxDPLCCmdRequest == 0)
                        {
                            if (objCommand.FirstTime)
                            {
                                RequestControlFromAuxDPLC();
                            }
                            else if (objCommand.TimeSince >= 10)
                            {
                                objCommand.Abort(LoadResString(ridUnableToGetControlFromPLC, "AuxD"), cmdTimedout);
                            }
                        }
                        else
                        {
                            mopCFLeft2Target = objCommand.Arguments(1);
                            mgvdCFLeftCmdTime = Now;
                            mgvbCFLeftCtl = true;
                            objCommand.Remove();
                        }

                        break;
                    case ocSetCFlapRight:
                        CheckInterlock(ocSetCFlapRight.ToString());
                        if (mipAuxDPLCCmdRequest == 0)
                        {
                            if (objCommand.FirstTime)
                            {
                                RequestControlFromAuxDPLC();
                            }
                            else if (objCommand.TimeSince >= 10)
                            {
                                objCommand.Abort(LoadResString(ridUnableToGetControlFromPLC, "AuxD"), cmdTimedout);
                            }
                        }
                        else
                        {
                            mopCFRight1Target = objCommand.Arguments(1);
                            mopCFRight2Target = objCommand.Arguments(2);
                            mgvdCFRightCmdTime = Now;
                            mgvbCFRightCtl = true;
                            objCommand.Remove();
                        }

                        break;
                    case ocSetCFlapRight1:
                        CheckInterlock(ocSetCFlapRight1.ToString());
                        if (mipAuxDPLCCmdRequest == 0)
                        {
                            if (objCommand.FirstTime)
                            {
                                RequestControlFromAuxDPLC();
                            }
                            else if (objCommand.TimeSince >= 10)
                            {
                                objCommand.Abort(LoadResString(ridUnableToGetControlFromPLC, "AuxD"), cmdTimedout);
                            }
                        }
                        else
                        {
                            mopCFRight1Target = objCommand.Arguments(1);
                            mgvdCFRightCmdTime = Now;
                            mgvbCFRightCtl = true;
                            objCommand.Remove();
                        }

                        break;
                    case ocSetCFlapRight2:
                        CheckInterlock(ocSetCFlapRight2.ToString());
                        if (mipAuxDPLCCmdRequest == 0)
                        {
                            if (objCommand.FirstTime)
                            {
                                RequestControlFromAuxDPLC();
                            }
                            else if (objCommand.TimeSince >= 10)
                            {
                                objCommand.Abort(LoadResString(ridUnableToGetControlFromPLC, "AuxD"), cmdTimedout);
                            }
                        }
                        else
                        {
                            mopCFRight2Target = objCommand.Arguments(1);
                            mgvdCFRightCmdTime = Now;
                            mgvbCFRightCtl = true;
                            objCommand.Remove();
                        }

                        break;
                    case ocSetCFlapTop:
                        CheckInterlock(ocSetCFlapTop.ToString());
                        if (mipAuxDPLCCmdRequest == 0)
                        {
                            if (objCommand.FirstTime)
                            {
                                RequestControlFromAuxDPLC();
                            }
                            else if (objCommand.TimeSince >= 10)
                            {
                                objCommand.Abort(LoadResString(ridUnableToGetControlFromPLC, "AuxD"), cmdTimedout);
                            }
                        }
                        else
                        {
                            mopCFTopTarget = objCommand.Arguments(1);
                            mgvdCFTopCmdTime = Now;
                            mgvbCFTopCtl = true;
                            objCommand.Remove();
                        }

                        break;
                    case ocStopCF:
                        CheckInterlock(ocStopCF.ToString());
                        if (mipAuxDPLCCmdRequest == 0)
                        {
                            if (objCommand.FirstTime)
                            {
                                RequestControlFromAuxDPLC();
                            }
                            else if (objCommand.TimeSince >= 10)
                            {
                                objCommand.Abort(LoadResString(ridUnableToGetControlFromPLC, "AuxD"), cmdTimedout);
                            }
                        }
                        else
                        {
                            mgvbCFLeftCtl = false;
                            mgvbCFRightCtl = false;
                            mgvbCFTopCtl = false;
                            objCommand.Remove();
                        }

                        break;
                    default:
                        objCommand.Abort(string.Concat("Unrecognized opcode: ", objCommand.Opcode));
                        break;
                }

                break;
        }

        goto Done;
        Error();
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = ex.Message;
        goto ErrorExit;
    ErrorExit:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, objCommand.Text, sRaiseErrorText);
        objCommand.Abort(sRaiseErrorText, cmdAborted);
        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        ReEntered = false;
        return;
    }

    private void mobjCommands_CommandInserted(ref CommandListLib.Command NewCommand, ref CommandListLib.Command ExistingCommand)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        switch (NewCommand.Opcode)
        {
            case ocStartMainFanAux:
                NewCommand.Messages.Add(LoadResString(ridAuxCurrentOff), mtInformation);
                break;
            case ocStartMainFan:
                NewCommand.Messages.Add(LoadResString(ridMFCurrentOff), mtInformation);
                break;
            case ocStopMainFan:
                NewCommand.Messages.Add(LoadResString(ridMFCurrentOn), mtInformation);
                break;
            case ocStopIdle:
                NewCommand.Messages.Add(LoadResString(ridIdleCurrentOn), mtInformation);
                break;
        }

        goto Done;
        Error();
        if (ex.HResult & vbObjectError != vbObjectError)
        {
            EH.Module = TypeName(this);
            EH.Procedure = "mobjCommands_CommandInserted";
            EH_MsgBox();
        }

        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        return;
    }

    private void mobjInterlocks_Evaluate(ref string InterlockName, ref bool CurrentOKValue, ref Collection AdditionalMessages)
    {
        float fCheckValue = 0;
        int I = 0;
        try
        {
            switch (InterlockName)
            {
                case ocStartMainFan.ToString():
                case ocStopMainFan.ToString():
                case ocStartMainFanAux.ToString():
                case ocStopMainFanAux.ToString():
                case ocStartIdle.ToString():
                case ocStopIdle.ToString():
                case ocShowStatus.ToString():
                case ocSetNozzleSize.ToString():
                case ocSetCFlapLeft1.ToString():
                case ocSetCFlapLeft2.ToString():
                case ocSetCFlapRight1.ToString():
                case ocSetCFlapRight2.ToString():
                case ocSetCFlapTop.ToString():
                case ocStopCF.ToString():
                    switch (true)
                    {
                        case mipOpModeTest == 1:
                        case mipOpModeStab == 1:
                        case mipOpModePrep == 1:
                            break;
                        case mipOpModeMaint == 1:
                            CurrentOKValue = false;
                            AdditionalMessages.Add(GetCommonMessage(cmWTModeMaint));
                            break;
                        case mipOpModeOff == 1:
                            CurrentOKValue = false;
                            AdditionalMessages.Add(GetCommonMessage(cmWTModeOff));
                            break;
                        default:
                            CurrentOKValue = false;
                            AdditionalMessages.Add(GetCommonMessage(cmWTModeNone));
                            break;
                    }

                    break;
            }

            switch (InterlockName)
            {
                case ocSetWindSpeed.ToString():
                    // On Error Resume Next — nested handler; see enclosing try/catch
                    ;
                    fCheckValue = mvCheckIlckValue;
                    if (Err == 0)
                    {
                        if (mipOpModePrep == 1)
                        {
                            if (fCheckValue > mccfMaxWsPrepMode)
                            {
                                CurrentOKValue = false;
                                AdditionalMessages.Add(GetCommonMessage2(cmInPrepModeMaxIs, mccfMaxWsPrepMode));
                            }
                        }
                        else if (mipOpModeStab == 1 & mgvlWindTunnel == wtAAWT)
                        {
                            if (fCheckValue > mccfMaxWsStabMode)
                            {
                                CurrentOKValue = false;
                                AdditionalMessages.Add(GetCommonMessage2(cmInStabModeMaxIs, mccfMaxWsStabMode));
                            }
                        }
                    }

                    break;
                case ocSetFanSpeed.ToString():
                    // On Error Resume Next — nested handler; see enclosing try/catch
                    ;
                    fCheckValue = mvCheckIlckValue;
                    if (Err == 0)
                    {
                        if (mipOpModePrep == 1)
                        {
                            if (fCheckValue > mccfMaxFsPrepMode)
                            {
                                CurrentOKValue = false;
                                AdditionalMessages.Add(GetCommonMessage2(cmInPrepModeMaxIs, mccfMaxFsPrepMode));
                            }
                        }
                        else if (mipOpModeStab == 1 & mgvlWindTunnel == wtAAWT)
                        {
                            if (fCheckValue > mccfMaxFsStabMode)
                            {
                                CurrentOKValue = false;
                                AdditionalMessages.Add(GetCommonMessage2(cmInStabModeMaxIs, mccfMaxFsStabMode));
                            }
                        }
                    }

                    break;
                case ocTrackWindSpeed.ToString():
                    switch (true)
                    {
                        case mipOpModeTest == 1:
                        case mipOpModeStab == 1:
                            break;
                        case mipOpModePrep == 1:
                            CurrentOKValue = false;
                            AdditionalMessages.Add(GetCommonMessage(cmWTModePrep));
                            break;
                        case mipOpModeMaint == 1:
                            CurrentOKValue = false;
                            AdditionalMessages.Add(GetCommonMessage(cmWTModeMaint));
                            break;
                        case mipOpModeOff == 1:
                            CurrentOKValue = false;
                            AdditionalMessages.Add(GetCommonMessage(cmWTModeOff));
                            break;
                        default:
                            CurrentOKValue = false;
                            AdditionalMessages.Add(GetCommonMessage(cmWTModeNone));
                            break;
                    }

                    break;
            }

            switch (InterlockName)
            {
                case ocStartMainFan.ToString():
                case ocShowStatus.ToString():
                    if (mipMfPerm == 0)
                    {
                        if (mccsMFPermComponents != "")
                        {
                            mobjMFPermComponents.Check(CurrentOKValue, AdditionalMessages);
                        }
                    }

                    if (mipMfStatus == 0 & mopMfControl == 0)
                    {
                        if (SecondsSinceTime(mgvdDriveOffTime) < mccfMainFanRestartTime)
                        {
                            CurrentOKValue = false;
                            AdditionalMessages.Add(LoadResString(ridMainFanRestartDelay, (int)mccfMainFanRestartTime - SecondsSinceTime(mgvdDriveOffTime)));
                        }
                    }

                    break;
            }

            switch (InterlockName)
            {
                case ocStartIdle.ToString():
                case ocStopIdle.ToString():
                case ocShowStatus.ToString():
                case ocShowIdleStatus.ToString():
                    if (mipIdleNozFlapMotorStop == 1)
                    {
                        if (mccsIdleCmdPermComponents != "")
                        {
                            mobjIdleCmdPermComponents.Check(CurrentOKValue, AdditionalMessages);
                        }
                    }

                    break;
            }

            switch (InterlockName)
            {
                case ocStartMainFanAux.ToString():
                case ocShowStatus.ToString():
                    if (mipAuxReady == 0)
                    {
                        if (mccsMFAuxPermComponents != "")
                        {
                            mobjMFAuxPermComponents.Check(CurrentOKValue, AdditionalMessages);
                        }
                    }

                    break;
            }

            switch (InterlockName)
            {
                case ocShowStatus.ToString():
                    for (I = 1; I <= 3; I++)
                    {
                        if (LimitInEffect(I))
                        {
                            if (mopWsMode > modeManual)
                            {
                                if (mgvfWsTarget > mccfMaxFsLimit(I) * mgvfNomWsMax / mgvfNomFsMax)
                                {
                                    CurrentOKValue = false;
                                    AdditionalMessages.Add(LoadResString(ridLowSpeedLmt1On + I - 1, mccfMaxFsLimit(I)));
                                }
                            }
                            else if (mgvfFsTarget > mccfMaxFsLimit(I))
                            {
                                CurrentOKValue = false;
                                AdditionalMessages.Add(LoadResString(ridLowSpeedLmt1On + I - 1, mccfMaxFsLimit(I)));
                            }
                        }
                    }

                    if (mopWsMode >= modeAuto)
                    {
                        if (mobjPID.States(ctlIntegral) >= mccfIntMax)
                        {
                            AdditionalMessages.Add(LoadResString(ridPIDIntegralAtMax));
                            CurrentOKValue = false;
                        }

                        if (mobjPID.States(ctlIntegral) <= mccfIntMin)
                        {
                            AdditionalMessages.Add(LoadResString(ridPIDIntegralAtMin));
                            CurrentOKValue = false;
                        }
                    }

                    if (mopWsMode == modeAuto)
                    {
                        if (mgvfWsTarget <= mccfMinWsForPid)
                        {
                            AdditionalMessages.Add(LoadResString(ridWsTooLowForPid, mccfMinWsForPid));
                            CurrentOKValue = false;
                        }
                    }

                    break;
                case ocSetWindSpeed.ToString():
                    for (I = 2; I <= 3; I++)
                    {
                        if (LimitInEffect(I))
                        {
                            // On Error Resume Next — nested handler; see enclosing try/catch
                            fCheckValue = mvCheckIlckValue;
                            if (Err == 0)
                            {
                                if (fCheckValue / mgvfNomWsMax * mgvfNomFsMax >= mccfMaxFsLimit(I))
                                {
                                    CurrentOKValue = false;
                                    AdditionalMessages.Add(LoadResString(ridLowSpeedLmt1On + I - 1, mccfMaxFsLimit(I)));
                                }
                            }
                        }
                    }

                    break;
                case ocSetFanSpeed.ToString():
                    for (I = 2; I <= 3; I++)
                    {
                        if (LimitInEffect(I))
                        {
                            // On Error Resume Next — nested handler; see enclosing try/catch
                            fCheckValue = mvCheckIlckValue;
                            if (Err == 0)
                            {
                                if (fCheckValue >= mccfMaxFsLimit(I))
                                {
                                    CurrentOKValue = false;
                                    AdditionalMessages.Add(LoadResString(ridLowSpeedLmt1On + I - 1, mccfMaxFsLimit(I)));
                                }
                            }
                        }
                    }

                    break;
            }

            return;
        }
        catch
        {
            // On Error Resume Next: exceptions suppressed
        }
    }

    internal void Initialize(InitTypes InitType = initAll, Command Command = 0)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        switch (InitType)
        {
            case initConstants:
            case initAll:
                bStatus = true;
                sPath = GetDefaultCCFolder(App.Path);
                sFileName = string.Concat(string.Concat(sPath, PROCESS_NAME), ".csv");
                sFileName = Reg_GetValue(HKEY_LOCAL_MACHINE, REG_PROCESSES, PROCESS_NAME, "ControlConstantsFile", sFileName);
                sDefaultsFileName = string.Concat(string.Concat(string.Concat(App.Path, "\\"), App.EXEName), "Defaults.csv");
                mobjConstants.OpenFileWithDefaults(sFileName, sDefaultsFileName);
                mobjConstants.AddString("Priority", mccsPriority, bStatus, sText);
                mobjConstants.AddSingle("UpdateInterval", mccfUpdateInt, bStatus, sText);
                mobjConstants.AddString("MFPermComponents", mccsMFPermComponents, bStatus, sText);
                mobjConstants.AddString("MFAuxPermComponents", mccsMFAuxPermComponents, bStatus, sText);
                mobjConstants.AddSingle("SlowFilterDelayTime", mccfSlowFilterDelayTime, bStatus, sText);
                mobjConstants.AddSingle("TrackingFilterTC", mccfWsTgtTrackFilTc, bStatus, sText);
                mobjConstants.AddBoolean("OutputFilterEnable", mccbFilterOutput, bStatus, sText);
                mobjConstants.AddSingle("OutputFilterTC", mccfPidOutputFilTc, bStatus, sText);
                mobjConstants.AddSingle("MainFanTimeout", mccfMainFanTimeout, bStatus, sText);
                mobjConstants.AddSingle("MainFanReadyDelay", mccfDrivePause, bStatus, sText);
                mobjConstants.AddSingle("MainFanAuxTimeout", mccfMainFanAuxTimeout, bStatus, sText);
                mobjConstants.AddSingle("MainFanAuxPause", mccfMainFanAuxPause, bStatus, sText);
                mobjConstants.AddSingle("MainFanAuxAutoOff", mccfMainFanAuxAutoOffTime, bStatus, sText);
                mobjConstants.AddSingle("MainFanRestartTime", mccfMainFanRestartTime, bStatus, sText);
                mobjConstants.AddSingle("IdleTimeout", mccfIdleTimeout, bStatus, sText);
                mobjConstants.AddSingle("SuspendTimeout", mccfMaxSuspendTime, bStatus, sText);
                mobjConstants.AddSingle("ResetTime", mccfResetFaultTime, bStatus, sText);
                mobjConstants.AddBoolean("ResetAuto", mccbAutoReset, bStatus, sText);
                for (J = 1; J <= IIf(WindTunnel == wtCWT, 3, 1); J++)
                {
                    mcclNumCloseTimes(J) = 1;
                    mobjConstants.AddLong(string.Concat(string.Concat("CloseTime", NozzleName(J)), "Num"), mcclNumCloseTimes(J), bStatus, sText);
                    if (mcclNumCloseTimes(J) > 10)
                    {
                        mcclNumCloseTimes(J) = 10;
                    }

                    if (J == nsLarge)
                    {
                        for (I = 1; I <= mcclNumCloseTimes(J); I++)
                        {
                            mobjConstants.AddSingle(string.Concat(string.Concat(string.Concat("CloseTime", NozzleName(J)), Format(I, "00")), "Time"), mccfCloseTimesLarge(I), bStatus, sText);
                            mobjConstants.AddSingle(string.Concat(string.Concat(string.Concat("CloseTime", NozzleName(J)), Format(I, "00")), "Speed"), mccfCloseTimeSpeedsLarge(I), bStatus, sText);
                        }
                    }
                    else if (J == nsMedium)
                    {
                        for (I = 1; I <= mcclNumCloseTimes(J); I++)
                        {
                            mobjConstants.AddSingle(string.Concat(string.Concat(string.Concat("CloseTime", NozzleName(J)), Format(I, "00")), "Time"), mccfCloseTimesMedium(I), bStatus, sText);
                            mobjConstants.AddSingle(string.Concat(string.Concat(string.Concat("CloseTime", NozzleName(J)), Format(I, "00")), "Speed"), mccfCloseTimeSpeedsMedium(I), bStatus, sText);
                        }
                    }
                    else
                    {
                        for (I = 1; I <= mcclNumCloseTimes(J); I++)
                        {
                            mobjConstants.AddSingle(string.Concat(string.Concat(string.Concat("CloseTime", NozzleName(J)), Format(I, "00")), "Time"), mccfCloseTimesSmall(I), bStatus, sText);
                            mobjConstants.AddSingle(string.Concat(string.Concat(string.Concat("CloseTime", NozzleName(J)), Format(I, "00")), "Speed"), mccfCloseTimeSpeedsSmall(I), bStatus, sText);
                        }
                    }

                    mobjConstants.AddSingle(string.Concat(string.Concat("CloseTime", NozzleName(J)), "FilterTC"), mccfCloseWsFilTc(J), bStatus, sText);
                    mobjConstants.AddSingle(string.Concat(string.Concat("CloseTime", NozzleName(J)), "WSRate"), mccfCloseWsRamp(J), bStatus, sText);
                }

                mobjConstants.AddBoolean("DeltaPFilterEnable", mccbFilterInput, bStatus, sText);
                mobjConstants.AddSingle("DeltaPFilterType", mccfDisplayFilType, bStatus, sText);
                mobjConstants.AddSingle("DeltaPFilterMin", mccfMinFilTc, bStatus, sText);
                mobjConstants.AddSingle("DeltaPFilterMax", mccfMaxFilTc, bStatus, sText);
                mobjConstants.AddSingle("DeltaPFilterSlowError", mccfMinWsErrorSlow, bStatus, sText);
                mobjConstants.AddSingle("DeltaPFilterSlowErrorSmall", mccfMinWsErrorSlowSmall, bStatus, sText);
                for (J = 1; J <= IIf(WindTunnel == wtCWT, 3, 1); J++)
                {
                    mobjConstants.AddSingle(string.Concat(string.Concat("DeltaPFilter", NozzleName(J)), "Base"), mccfWsFilterTc0(J), bStatus, sText);
                    mobjConstants.AddSingle(string.Concat(string.Concat("DeltaPFilter", NozzleName(J)), "Adjust"), mccfWsFilterTc1(J), bStatus, sText);
                    mobjConstants.AddSingle(string.Concat(string.Concat("DeltaPFilter", NozzleName(J)), "Fast"), mccfWsFilterFast(J), bStatus, sText);
                    mobjConstants.AddSingle(string.Concat(string.Concat("DeltaPFilter", NozzleName(J)), "FastTrack"), mccfWsFilterFastTrack(J), bStatus, sText);
                    mobjConstants.AddSingle(string.Concat(string.Concat("DeltaPControlFilter", NozzleName(J)), "Base"), mccfControlFilTc0(J), bStatus, sText);
                    mobjConstants.AddSingle(string.Concat(string.Concat("DeltaPControlFilter", NozzleName(J)), "Adjust"), mccfControlFilTc1(J), bStatus, sText);
                    mobjConstants.AddSingle(string.Concat(string.Concat("DeltaPControlFilter", NozzleName(J)), "Min"), mccfControlFilTcMin(J), bStatus, sText);
                    mobjConstants.AddSingle(string.Concat(string.Concat("DeltaPControlFilter", NozzleName(J)), "Max"), mccfControlFilTcMax(J), bStatus, sText);
                }

                mobjConstants.AddSingle("FFGain", mccfFfGain, bStatus, sText);
                mobjConstants.AddSingle("FFLag", mccfFfLagCoef, bStatus, sText);
                mobjConstants.AddSingle("FFLagTrack", mccfFfLagCoefTrack, bStatus, sText);
                mobjConstants.AddSingle("FFLead", mccfFfLeadCoef, bStatus, sText);
                mobjConstants.AddSingle("FFLeadTrack", mccfFfLeadCoefTrack, bStatus, sText);
                for (J = 1; J <= IIf(WindTunnel == wtCWT, 3, 1); J++)
                {
                    mobjConstants.AddSingle(string.Concat(string.Concat("FF", NozzleName(J)), "Fanspeed"), mccfNomFs(J), bStatus, sText);
                    mobjConstants.AddSingle(string.Concat(string.Concat("FF", NozzleName(J)), "Windspeed"), mccfNomWs(J), bStatus, sText);
                }

                mobjConstants.AddSingle("FanspeedMin", mccfMinFs, bStatus, sText);
                mobjConstants.AddSingle("FanspeedMinAutoMode", mccfMinFsAutoMode, bStatus, sText);
                if (WindTunnel == wtCWT)
                {
                    mobjConstants.AddSingle("FanspeedMaxLarge", mccfMaxFs(nsLarge), bStatus, sText);
                    mobjConstants.AddSingle("FanspeedMaxMedium", mccfMaxFs(nsMedium), bStatus, sText);
                    mobjConstants.AddSingle("FanspeedMaxSmall", mccfMaxFs(nsSmall), bStatus, sText);
                }
                else
                {
                    mobjConstants.AddSingle("FanspeedMaxLarge", mccfMaxFs(nsLarge), bStatus, sText);
                }

                mobjConstants.AddSingle("FanspeedMaxLimit1", mccfMaxFsLimit(1), bStatus, sText);
                mobjConstants.AddSingle("FanspeedMaxLimit2", mccfMaxFsLimit(2), bStatus, sText);
                mobjConstants.AddSingle("FanspeedMaxLimit3", mccfMaxFsLimit(3), bStatus, sText);
                mobjConstants.AddSingle("FanspeedRampMin", mccfMinFsRamp, bStatus, sText);
                mobjConstants.AddSingle("FanspeedRampMax", mccfMaxFsRamp, bStatus, sText);
                mobjConstants.AddSingle("FanspeedRampMaxTrack", mccfMaxFsRampTrack, bStatus, sText);
                mobjConstants.AddSingle("FanspeedRampDefault", mccfDefFsRamp, bStatus, sText);
                mobjConstants.AddSingle("FanspeedMaxPrep", mccfMaxFsPrepMode, bStatus, sText);
                if (WindTunnel == wtAAWT)
                {
                    mobjConstants.AddSingle("FanspeedMaxStab", mccfMaxFsStabMode, bStatus, sText);
                    mobjConstants.AddSingle("WindspeedRampMaxLowSpeed", mccfMaxWsRampLowSpeeds, bStatus, sText);
                    mobjConstants.AddSingle("FanspeedSlowRamp", mccfSlowRampFanspeed, bStatus, sText);
                }

                mobjConstants.AddSingle("WindspeedMin", mccfMinWs, bStatus, sText);
                if (WindTunnel == wtCWT)
                {
                    mobjConstants.AddSingle("WindspeedMaxLarge", mccfMaxWs(nsLarge), bStatus, sText);
                    mobjConstants.AddSingle("WindspeedMaxMedium", mccfMaxWs(nsMedium), bStatus, sText);
                    mobjConstants.AddSingle("WindspeedMaxSmall", mccfMaxWs(nsSmall), bStatus, sText);
                }
                else
                {
                    mobjConstants.AddSingle("WindspeedMaxLarge", mccfMaxWs(nsLarge), bStatus, sText);
                }

                mobjConstants.AddSingle("WindspeedRampMin", mccfMinWsRamp, bStatus, sText);
                mobjConstants.AddSingle("WindspeedRampMax", mccfMaxWsRamp, bStatus, sText);
                mobjConstants.AddSingle("WindspeedRampDefault", mccfDefWsRamp, bStatus, sText);
                mobjConstants.AddSingle("WindspeedMaxPrep", mccfMaxWsPrepMode, bStatus, sText);
                if (WindTunnel == wtAAWT)
                {
                    mobjConstants.AddSingle("WindspeedMaxStab", mccfMaxWsStabMode, bStatus, sText);
                }

                mobjConstants.AddSingle("HeadwindMax", mccfMaxWsHeadwind, bStatus, sText);
                mobjConstants.AddSingle("HeadwindMin", mccfMinWsHeadwind, bStatus, sText);
                mobjConstants.AddSingle("FanspeedMonLimit", mccfFsMonitorError, bStatus, sText);
                mobjConstants.AddSingle("FanspeedMonFiltTC", mccfFsMonitorFilTc, bStatus, sText);
                mobjConstants.AddSingle("FanspeedMonPeriod", mccfFsMonitorPeriod, bStatus, sText);
                mobjConstants.AddSingle("WindspeedMonLimit", mccfMonitorError, bStatus, sText);
                mobjConstants.AddSingle("WindspeedMonFiltTC", mccfMonitorFilTc, bStatus, sText);
                mobjConstants.AddSingle("WindspeedMonPeriod", mccfMonitorPeriod, bStatus, sText);
                mobjConstants.AddSingle("PIDIntegralMax", mccfIntMax, bStatus, sText);
                mobjConstants.AddSingle("PIDIntegralMin", mccfIntMin, bStatus, sText);
                mobjConstants.AddSingle("PIDOutputMax", mccfCtlMax, bStatus, sText);
                mobjConstants.AddSingle("PIDOutputMin", mccfCtlMin, bStatus, sText);
                mobjConstants.AddSingle("PIDDerivFilterTC", mccfTdFilterTC, bStatus, sText);
                mobjConstants.AddSingle("PIDErrorMax", mccfMaxWsError, bStatus, sText);
                mobjConstants.AddSingle("PIDWindspeedMin", mccfMinWsForPid, bStatus, sText);
                mobjConstants.AddSingle("PIDWindspeedMinDelay", mccfMinWsForPidDelay, bStatus, sText);
                for (J = 1; J <= IIf(WindTunnel == wtCWT, 3, 1); J++)
                {
                    mobjConstants.AddSingle(string.Concat(string.Concat("PIDGain", NozzleName(J)), "Base"), mccfKp0(J), bStatus, sText);
                    mobjConstants.AddSingle(string.Concat(string.Concat("PIDGain", NozzleName(J)), "Adjust"), mccfKp1(J), bStatus, sText);
                    mobjConstants.AddSingle(string.Concat(string.Concat("PIDGain", NozzleName(J)), "TrackBase"), mccfKp0Track(J), bStatus, sText);
                    mobjConstants.AddSingle(string.Concat(string.Concat("PIDGain", NozzleName(J)), "TrackAdjust"), mccfKp1Track(J), bStatus, sText);
                    mobjConstants.AddSingle(string.Concat(string.Concat("PIDGain", NozzleName(J)), "Min"), mccfKpMin(J), bStatus, sText);
                    mobjConstants.AddSingle(string.Concat(string.Concat("PIDGain", NozzleName(J)), "Max"), mccfKpMax(J), bStatus, sText);
                    mobjConstants.AddSingle(string.Concat(string.Concat("PIDDeriv", NozzleName(J)), "Base"), mccfTd0(J), bStatus, sText);
                    mobjConstants.AddSingle(string.Concat(string.Concat("PIDDeriv", NozzleName(J)), "Adjust"), mccfTd1(J), bStatus, sText);
                    mobjConstants.AddSingle(string.Concat(string.Concat("PIDDeriv", NozzleName(J)), "TrackBase"), mccfTd0Track(J), bStatus, sText);
                    mobjConstants.AddSingle(string.Concat(string.Concat("PIDDeriv", NozzleName(J)), "TrackAdjust"), mccfTd1Track(J), bStatus, sText);
                    mobjConstants.AddSingle(string.Concat(string.Concat("PIDDeriv", NozzleName(J)), "Min"), mccfTdMin(J), bStatus, sText);
                    mobjConstants.AddSingle(string.Concat(string.Concat("PIDDeriv", NozzleName(J)), "Max"), mccfTdMax(J), bStatus, sText);
                    mobjConstants.AddSingle(string.Concat(string.Concat("PIDIntegral", NozzleName(J)), "Base"), mccfTi0(J), bStatus, sText);
                    mobjConstants.AddSingle(string.Concat(string.Concat("PIDIntegral", NozzleName(J)), "Adjust"), mccfTi1(J), bStatus, sText);
                    mobjConstants.AddSingle(string.Concat(string.Concat("PIDIntegral", NozzleName(J)), "TrackBase"), mccfTi0Track(J), bStatus, sText);
                    mobjConstants.AddSingle(string.Concat(string.Concat("PIDIntegral", NozzleName(J)), "TrackAdjust"), mccfTi1Track(J), bStatus, sText);
                    mobjConstants.AddSingle(string.Concat(string.Concat("PIDIntegral", NozzleName(J)), "Min"), mccfTiMin(J), bStatus, sText);
                    mobjConstants.AddSingle(string.Concat(string.Concat("PIDIntegral", NozzleName(J)), "Max"), mccfTiMax(J), bStatus, sText);
                }

                mobjConstants.AddSingle("PIDIntegralMaxWSRateMax", mccfMaxWsRateForI, bStatus, sText);
                mobjConstants.AddSingle("PIDIntegralMaxWSRateMin", mccfMinWsRateForI, bStatus, sText);
                mobjConstants.AddSingle("PIDIntegralMaxWSRateBase", mccfWsRateForI0, bStatus, sText);
                mobjConstants.AddSingle("PIDIntegralMaxWSRateAdjust", mccfWsRateForI1, bStatus, sText);
                mobjConstants.AddSingle("DeltaPMinForWS", mccfMinDeltaP, bStatus, sText);
                mobjConstants.AddSingle("DeltaPMinForWSFSMax", mccfMinDeltaPFSMax, bStatus, sText);
                if (WindTunnel == wtCWT)
                {
                    mobjConstants.AddSingle("IdleFanSpeed", mccfIdleFanspeed, bStatus, sText);
                    mobjConstants.AddBoolean("AutoIdleDefault", mccbDefAutoIdle, bStatus, sText);
                    mobjConstants.AddSingle("AutoIdleStartWS", mccfWSStartIdle, bStatus, sText);
                    mobjConstants.AddSingle("AutoIdleStopWS", mccfWSStopIdle, bStatus, sText);
                    mobjConstants.AddSingle("AutoIdleDeadband", mccfAutoIdleDeadband, bStatus, sText);
                    mobjConstants.AddString("IdlePermComponents", mccsIdleCmdPermComponents, bStatus, sText);
                    mobjConstants.AddSingle("IdleReversePauseTime", mccfCancelIdleTime, bStatus, sText);
                    mobjConstants.AddSingle("NozzleTimeout", mccfNozzleTimeout, bStatus, sText);
                    mobjConstants.AddSingle("CFLeft1Min", mccfCFLeft1Min, bStatus, sText);
                    mobjConstants.AddSingle("CFLeft1Max", mccfCFLeft1Max, bStatus, sText);
                    mobjConstants.AddSingle("CFLeft2Min", mccfCFLeft2Min, bStatus, sText);
                    mobjConstants.AddSingle("CFLeft2Max", mccfCFLeft2Max, bStatus, sText);
                    mobjConstants.AddSingle("CFRight1Min", mccfCFRight1Min, bStatus, sText);
                    mobjConstants.AddSingle("CFRight1Max", mccfCFRight1Max, bStatus, sText);
                    mobjConstants.AddSingle("CFRight2Min", mccfCFRight2Min, bStatus, sText);
                    mobjConstants.AddSingle("CFRight2Max", mccfCFRight2Max, bStatus, sText);
                    mobjConstants.AddSingle("CFTopMin", mccfCFTopMin, bStatus, sText);
                    mobjConstants.AddSingle("CFTopMax", mccfCFTopMax, bStatus, sText);
                    mobjConstants.AddSingle("CFHoldTime", mccfCFHoldTime, bStatus, sText);
                    mobjConstants.AddSingle("CFTimeout", mccfCFTimeout, bStatus, sText);
                    mobjConstants.AddSingle("CFTolerance", mccfCFTolerance, bStatus, sText);
                }

                mobjConstants.AddString("TrackParDefault", mccsDefTrackParName, bStatus, sText);
                mobjConstants.AddSingle("TurnOffFanTime", mccfTurnOffFanTime, bStatus, sText);
                for (I = 1; I <= IIf(WindTunnel == wtCWT, 3, 1); I++)
                {
                    mobjConstants.AddString(string.Concat("WSCoeffsFile", NozzleName(I)), mccsWCoeffsFile(I), bStatus, sText);
                }

                mobjConstants.AddSingle("FanRunTimeWithoutCooling", mccfMaxRunTimeWithoutCooling, bStatus, sText);
                if (mgvlWindTunnel == wtCWT)
                {
                    mobjConstants.AddBoolean("TrackDyTarget", mccbAutoFollowDyTarget, bStatus, sText);
                    mobjConstants.AddBoolean("TrackDyTargetUseTrackCoeffs", mccbAutoFollowDyTargetUseTrackCoeffs, bStatus, sText);
                    mobjConstants.AddBoolean("AutoIdleUsesLimit1", mccbAutoIdleUsesLimitCommand, bStatus, sText);
                }

                mobjConstants.AddSingle("FFReverseWindspeed", mccfNomWSReverse, bStatus, sText);
                mobjConstants.AddSingle("FFReverseFanspeed", mccfNomFSReverse, bStatus, sText);
                mobjConstants.AddSingle("ForceWSZeroTime", mccfTimeAtZeroToForceWSZero, bStatus, sText);
                mobjConstants.AddSingle("PIDDerivUseTgtErrorLimit", mccfUseTargetErrorForDerivLimit, bStatus, sText);
                mobjConstants.AddLong("SetptRampMethod", mcclSetpointRampMethod, bStatus, sText);
                mobjConstants.AddSingle("FinalRampDeltaTime", mccfFinalRampTimeSec, bStatus, sText);
                mobjConstants.AddSingle("FinalRampDeltaSpeed", mccfFinalRampDeltakph, bStatus, sText);
                mobjConstants.AddBoolean("FinalRampSnap", mccbFinalRampSnap, bStatus, sText);
                mobjConstants.AddBoolean("FinalRampUseSetpt", mccbFinalRampUseSetpoint, bStatus, sText);
                mobjConstants.AddSingle("FinalRampMinWSRate", mccfWsRateForFinalRamp, bStatus, sText);
                mobjConstants.AddBoolean("FinalRampSnapBumpless", mccbFinalRampSnapBumpless, bStatus, sText);
                mobjConstants.AddSingle("FinalRampDeltaTimeSmall", mccfFinalRampTimeSecSmall, bStatus, sText);
                mobjConstants.AddSingle("FinalRampDeltaSpeedSmall", mccfFinalRampDeltakphSmall, bStatus, sText);
                if (WindTunnel == wtAAWT)
                {
                    for (I = 1; I <= 18; I++)
                    {
                        mobjConstants.AddBoolean(string.Concat(string.Concat("OverflowFlap", Format(I, "00")), "Open"), mccbOFDefault(I), bStatus, sText);
                    }

                    mobjConstants.AddBoolean("OverflowFlapPurgeOpen", mccbOFPurgeOpen, bStatus, sText);
                }

                if (bStatus == false)
                {
                    bRaiseError = true;
                    lRaiseErrorNumber = tceCommandFailed;
                    if (Command == null)
                    {
                    }
                    else
                    {
                        Command.SuppressLogging = true;
                    }

                    objEH = new EH();
                    objEH.AppError = LoadResString(StdTalentStrings.stsEHCantLoadConstants);
                    objEH.Item(1) = sText;
                    objEH.FileName = sFileName;
                    objEH.Module = TypeName(this);
                    objEH.Procedure = PROC_NAME;
                    sRaiseErrorText = objEH.Text;
                    LogMessage2(objEH);
                    goto Done;
                }

                mobjSPGConstants.CreatePropertyByPosition(CC_FSMin, bExists).Value = mccfMinFs;
                mobjSPGConstants.CreatePropertyByPosition(CC_FSMax, bExists).Value = mgvfMaxFs;
                mobjSPGConstants.CreatePropertyByPosition(CC_FSRampMax, bExists).Value = mccfMaxFsRamp;
                mobjSPGConstants.CreatePropertyByPosition(CC_WSMin, bExists).Value = mccfMinWs;
                mobjSPGConstants.CreatePropertyByPosition(CC_WSMax, bExists).Value = mgvfMaxWs;
                mobjSPGConstants.CreatePropertyByPosition(CC_WSRampMax, bExists).Value = mccfMaxWsRamp;
                mobjSPGConstants.CreatePropertyByPosition(CC_HeadwindMin, bExists).Value = mccfMinWsHeadwind;
                mobjSPGConstants.CreatePropertyByPosition(CC_HeadwindMax, bExists).Value = mccfMaxWsHeadwind;
                mobjSPGConstants.CreatePropertyByPosition(CC_WSRampDef, bExists).Value = mccfDefWsRamp;
                mobjSPGConstants.CreatePropertyByPosition(CC_FSRampDef, bExists).Value = mccfDefFsRamp;
                mobjSPGConstants.CreatePropertyByPosition(CC_DefAutoIdle, bExists).Value = mccbDefAutoIdle;
                mobjSPGConstants.CreatePropertyByPosition(CC_DefTrackParName, bExists).Value = mccsDefTrackParName;
                mobjSPGConstants.CreatePropertyByPosition(CC_CFLeft1Min, bExists).Value = mccfCFLeft1Min;
                mobjSPGConstants.CreatePropertyByPosition(CC_CFLeft1Max, bExists).Value = mccfCFLeft1Max;
                mobjSPGConstants.CreatePropertyByPosition(CC_CFLeft2Min, bExists).Value = mccfCFLeft2Min;
                mobjSPGConstants.CreatePropertyByPosition(CC_CFLeft2Max, bExists).Value = mccfCFLeft2Max;
                mobjSPGConstants.CreatePropertyByPosition(CC_CFRight1Min, bExists).Value = mccfCFRight1Min;
                mobjSPGConstants.CreatePropertyByPosition(CC_CFRight1Max, bExists).Value = mccfCFRight1Max;
                mobjSPGConstants.CreatePropertyByPosition(CC_CFRight2Min, bExists).Value = mccfCFRight2Min;
                mobjSPGConstants.CreatePropertyByPosition(CC_CFRight2Max, bExists).Value = mccfCFRight2Max;
                mobjSPGConstants.CreatePropertyByPosition(CC_CFTopMin, bExists).Value = mccfCFTopMin;
                mobjSPGConstants.CreatePropertyByPosition(CC_CFTopMax, bExists).Value = mccfCFTopMax;
                mgvbSeparateControl = true;
                if (mccfDisplayFilType == 2)
                {
                    mobjWSFilterBW = new Butterworth();
                    mobjWSFilterBW.InitializeLowPass(1 / mccfUpdateInt, 1 / 2 * 3.14159 * mccfWsFilterTc0(nsLarge), 4);
                }
                else if (mccfDisplayFilType == 1)
                {
                    mobjWSFilter.FilterType = ctlRollingAverage;
                    mobjWSFilter.FilterConst(0) = mccfWsFilterTc0(nsLarge) / mccfUpdateInt;
                }
                else
                {
                    mobjWSFilter.FilterType = ctlFirstOrder;
                    mobjWSFilter.FilterConst(0) = mccfWsFilterTc0(nsLarge);
                }

                mobjWSFilter.UpdateTime = mccfUpdateInt;
                if (mccfDisplayFilType == 2)
                {
                    mobjWSFilterBW2 = new Butterworth();
                    mobjWSFilterBW2.InitializeLowPass(1 / mccfUpdateInt, 1 / 2 * 3.14159 * mccfWsFilterTc0(nsLarge), 4);
                }
                else if (mccfDisplayFilType == 1)
                {
                    mobjWSFilter2.FilterType = ctlRollingAverage;
                    mobjWSFilter2.FilterConst(0) = mccfWsFilterTc0(nsLarge) / mccfUpdateInt;
                }
                else
                {
                    mobjWSFilter2.FilterType = ctlFirstOrder;
                    mobjWSFilter2.FilterConst(0) = mccfWsFilterTc0(nsLarge);
                }

                mobjWSFilter2.UpdateTime = mccfUpdateInt;
                mobjWSctlFilter.FilterType = ctlFirstOrder;
                mobjWSctlFilter.FilterConst(0) = mccfControlFilTcMin(nsLarge);
                mobjWSctlFilter.UpdateTime = mccfUpdateInt;
                if (mccfWsTgtTrackFilTc < mccfUpdateInt)
                {
                    mccfWsTgtTrackFilTc = mccfUpdateInt;
                }

                mobjTrackFilter.FilterType = ctlFirstOrder;
                mobjTrackFilter.FilterConst(0) = mccfWsTgtTrackFilTc;
                mobjTrackFilter.UpdateTime = mccfUpdateInt;
                mobjRampRateFilter.FilterType = ctlFirstOrder;
                mobjRampRateFilter.UpdateTime = mccfUpdateInt;
                mobjWSMonitor.ErrorBand = mccfMonitorError;
                mobjWSMonitor.FilterTimeConst = mccfMonitorFilTc;
                mobjWSMonitor.MonitoringPeriod = mccfMonitorPeriod;
                mobjWSMonitor.UpdateTime = mccfUpdateInt;
                mobjFSMonitor.ErrorBand = mccfFsMonitorError;
                mobjFSMonitor.FilterTimeConst = mccfFsMonitorFilTc;
                mobjFSMonitor.MonitoringPeriod = mccfFsMonitorPeriod;
                mobjFSMonitor.UpdateTime = mccfUpdateInt;
                mobjFFLeadLag.FilterType = ctlLeadLagCompensation;
                mobjFFLeadLag.UpdateTime = mccfUpdateInt;
                mgvlDefPriority = GetControllerPriority(mccsPriority);
                mgvlOldNozzle = -1;
                if (mccfIdleFanspeed > mccfMaxFsLimit(1))
                {
                    mccfIdleFanspeed = mccfMaxFsLimit(1);
                }

                for (I = 1; I <= IIf(WindTunnel == wtCWT, 3, 1); I++)
                {
                    LoadWSCoeffs(I, string.Concat(sPath, mccsWCoeffsFile(I)));
                }

                break;
        }

        switch (InitType)
        {
            case initAll:
            case initPars:
                // On Error Resume Next — nested handler; see enclosing try/catch
                ;
                mobjParManager.RefreshObjects();
                if (Err)
                {
                    lRaiseErrorNumber = CCommandError(ex.HResult);
                    sRaiseErrorText = ex.Message;
                    objEH = new EH();
                    objEH.VBErr = Err;
                    objEH.AppError = LoadResString(StdTalentStrings.stsEHInitializeFailedParsMissing);
                    objEH.Item(1) = Command.Text;
                    objEH.Module = TypeName(this);
                    objEH.Procedure = PROC_NAME;
                    if (Command == null)
                    {
                    }
                    else
                    {
                        objEH.Item(1) = Command.Text;
                    }

                    LogMessage2(objEH);
                    if (Command == null)
                    {
                    }
                    else
                    {
                        Command.SuppressLogging = true;
                    }

                    goto Done;
                }

                if (mopBlockageCorr == 0)
                {
                    mopBlockageCorr = 1;
                }

                mbInited = mobjParManager.Ready;
                if (mbInited)
                {
                    mobjParManager.ReadInputs();
                }

                mobjMFPermComponents.SetComponents(mccsMFPermComponents);
                mobjMFAuxPermComponents.SetComponents(mccsMFAuxPermComponents);
                mobjIdleCmdPermComponents.SetComponents(mccsIdleCmdPermComponents);
                mgvdTurnOffFanTime = DateTime.Now;
                break;
        }

        goto Done;
        Error();
        mbInited = false;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = ex.Message;
        bRaiseError = true;
        EH.Module = TypeName(this);
        EH.Procedure = PROC_NAME;
        if (Command == null)
        {
        }
        else
        {
            EH.Item(1) = Command.Text;
        }

        EH_MsgBox();
        if (Command == null)
        {
        }
        else
        {
            Command.SuppressLogging = true;
        }

        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        mobjStatuses.SetValue(STS_WINDSPEED_INITIALIZED, mbInited);
        mobjConstants.CloseFile();
        if (mobjForm == null)
        {
            tmrMain.Enabled = true;
        }

        if (mbInited)
        {
            SetProcessPriority(mgvlDefPriority);
        }
        else
        {
            SetProcessPriority(priNormal);
        }

        if (bRaiseError)
        {
            EH_RaiseError(lRaiseErrorNumber, TypeName(this), PROC_NAME, sRaiseErrorText)();
        }

        return;
    }

    private void LoadWSCoeffs(ref NozzleSizes Nozzle, ref string FileName)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        sNozzleName = NozzleName(Nozzle);
        CheckFileOpenable(FileName, ,, true);
        objApp = CreateObject("ReACT.XLSXExporterLib.ExcelReader");
        // On Error GoTo Error_OpeningFile — nested handler not restructured
        objWB = objApp.OpenWorkbook(FileName, true);
        // On Error GoTo Error_ReadingData — nested handler not restructured
        vValue2 = objWB.Worksheets(1).Range("B7:J10").Value;
        vValue = new object[UBound(vValue2, 1) - 0 + 1 + 1, UBound(vValue2, 2) - 0 + 1 + 1];
        for (I = 0; I <= UBound(vValue2, 1); I++)
        {
            for (J = 0; J <= UBound(vValue2, 2); J++)
            {
                vValue(I - 0 + 1, J - 0 + 1) = vValue2(I, J);
            }
        }// On Error GoTo Error — nested handler not restructured
        for (I = 0; I <= 6; I++)
        {
            lCol = 7 - I;
            mccdKqnCoeffs(Nozzle, I) = vValue(1, lCol);
            mccdKqplCoeffs(Nozzle, I) = vValue(2, lCol);
            mccdKpCoeffs(Nozzle, I) = vValue(3, lCol);
            mccdWSvsFSCoeffs(Nozzle, I) = vValue(4, lCol);
        }

        mccfKqnCoeffMin(Nozzle) = vValue(1, 8);
        mccfKqnCoeffMax(Nozzle) = vValue(1, 9);
        mccfKqplCoeffMin(Nozzle) = vValue(2, 8);
        mccfKqplCoeffMax(Nozzle) = vValue(2, 9);
        mccfKpCoeffMin(Nozzle) = vValue(3, 8);
        mccfKpCoeffMax(Nozzle) = vValue(3, 9);
        goto Done;
        Error();
        sRaiseErrorText = LoadResString(ridErrorLoadingWSCoeffs, sNozzleName, FileName, ex.Message);
        bRaiseError = true;
        goto Done;
    Error_OpeningFile:
        ;
        sRaiseErrorText = LoadResString(ridErrorOpeningFile, FileName, ex.Message);
        bRaiseError = true;
        goto Done;
    Error_ReadingData:
        ;
        sRaiseErrorText = LoadResString(ridErrorReadingData, FileName, "B7:J10", ex.Message);
        bRaiseError = true;
        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        objWB.Close(false);
        objWB = null;
        if (bRaiseError)
        {
            EH_RaiseError(OLEERR_METHOD_FAILED, TypeName(this), PROC_NAME, sRaiseErrorText);
        }

        return;
    }

    private void Suspend()
    {
        try
        {
            if (mbInited)
            {
                StopFan(true);
            }

            mbInited = false;
            mobjParManager.ReleaseOutputs();
            mopFsSetpt = 0;
            mopFsTargetDisp = 0;
            mgvfFsTarget = 0;
            mopWsMode = 0;
            mopMfControl = 0;
            mopAuxControl = 0;
            mopIdleStartControl = 0;
            mopIdleStopControl = 0;
            mobjStatuses(STS_WINDSPEED_INITIALIZED).SetValue(mbInited);
            SetProcessPriority(priNormal);
            return;
        }
        catch
        {
            // On Error Resume Next: exceptions suppressed
        }
    }

    private void StopFan(bool WaitForComplete = false)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        if (mbInited == false)
        {
            goto Done;
        }

        vStart = Now;
        mgvfFsTarget = 0;
        mopFsTargetDisp = mgvfFsTarget;
        if (mopWsMode > modeManual)
        {
            mopWsMode = modeManual;
        }
        else if (mopWsMode == modeOff)
        {
            goto Done;
        }

        while (!(mopFsSetpt == 0))
        {
            mopFsSetpt = Ctl_Ramp(mgvfFsTarget, mopFsSetpt, mccfMaxFsRamp, mccfUpdateInt, mgvfFsIncrement);
            mobjParManager.WriteOutputs();
            if (SecondsSinceTime(vStart) >= mccfMaxSuspendTime)
            {
                break;
            }

            Pause(mccfUpdateInt);
        }

        mopMfControl = 0;
        mobjParManager.WriteOutputs();
        while (!(mipMfStatus == 0))
        {
            mobjParManager.ReadInputs();
            InputConversions();
            if (SecondsSinceTime(vStart) >= mccfMaxSuspendTime)
            {
                break;
            }

            Pause(mccfUpdateInt);
        }

        mopAuxControl = 0;
        mobjParManager.WriteOutputs();
        goto Done;
        Error();
        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        return;
    }

    public void Shutdown()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        SetProcessPriority(priNormal);
        mbReady = false;
        tmrMain.Enabled = false;
        mobjCommands.Clear("Shutdown");
        DoEvents();
        mbInited = false;
        mobjParManager.ReleaseOutputs();
        mobjEvents.RaiseQuit();
        if (!mobjForm == null)
        {
            Unload(mobjForm);
            mobjForm = null;
        }

        objSOM = new SOM();
        objSOM.RemoveObject(SOM_CONTROLLER);
        objSOM = null;
        Class_Terminate();
        Pause(2);
        if (IsWithinIDE())
        {
            End();
        }
        else
        {
            TerminateProcess(GetCurrentProcess(), 0);
        }

        goto Done;
        Error();
        sRaiseErrorText = ex.Message;
        bRaiseError = true;
        EH.Module = TypeName(this);
        EH.Procedure = "Shutdown";
        EH_MsgBox();
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        objSOM = null;
        if (bRaiseError)
        {
            EH_RaiseError(OLEERR_METHOD_FAILED, TypeName(this), "Shutdown", sRaiseErrorText)();
        }

        return;
    }

    public ControlLib.ControlConstants ControlConstants
    {
        get
        {
            ControlLib.ControlConstants __result = 0;
            // TODO: On Error GoTo Error — handler label not found in this scope
            bRaiseError = false;
            __result = mobjConstants;
            goto Done;
            Error();
            sRaiseErrorText = ex.Message;
            bRaiseError = true;
            EH.Module = TypeName(this);
            EH.Procedure = PROC_NAME;
            EH_MsgBox();
            goto Done;
        Done:
            ;
            // On Error Resume Next — nested handler; see enclosing try/catch
            if (bRaiseError)
            {
                EH_RaiseError(OLEERR_PROPERTY_GET_FAILED, TypeName(this), PROC_NAME, sRaiseErrorText)();
            }

            return __result;
        }
    }

    public ControlLib.GlobalVars GlobalVars
    {
        get
        {
            ControlLib.GlobalVars __result = 0;
            // TODO: On Error GoTo Error — handler label not found in this scope
            bRaiseError = false;
            __result = mobjGlobalVars;
            goto Done;
            Error();
            sRaiseErrorText = ex.Message;
            bRaiseError = true;
            EH.Module = TypeName(this);
            EH.Procedure = PROC_NAME;
            EH_MsgBox();
            goto Done;
        Done:
            ;
            // On Error Resume Next — nested handler; see enclosing try/catch
            if (bRaiseError)
            {
                EH_RaiseError(OLEERR_PROPERTY_GET_FAILED, TypeName(this), PROC_NAME, sRaiseErrorText)();
            }

            return __result;
        }
    }

    public bool Ready
    {
        get
        {
            bool __result = false;
            try
            {
                __result = mbReady;
                return __result;
            }
            catch
            {
                // On Error Resume Next: exceptions suppressed
            }

            return __result;
        }
    }

    private void Class_Initialize()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        // On Error Resume Next — nested handler; see enclosing try/catch
        if (App.PrevInstance)
        {
            if (App.StartMode != vbSModeStandalone)
            {
                EH.AppError = StdTalentStrings.stsEHInvalidActiveXLaunch;
                EH.Item(1) = PROCESS_NAME;
                EH.Module = TypeName(this);
                EH.Procedure = "Class_Initialize";
                EH_MsgBox();
            }

            End();
        }// On Error GoTo Error — nested handler not restructured
        App.StartLogging("", vbLogAuto);
        App.OleServerBusyRaiseError = true;
        App.OleServerBusyTimeout = 1000;
        Process.Title = PROCESS_NAME;
        Process.EHCaption = PROCESS_NAME;
        Process.EHMode = ehLogOnly;
        Process.EHMsgsToLog = ehCriticalExclamation;
        Process.LogMode = logToTalentAndFile;
        // On Error GoTo Error_CreatingObjects — nested handler not restructured
        sClassName = "ccrpTimers6.ccrpTimer";
        tmrMain = new LoopTimer();
        if (App.StartMode != vbSModeStandalone)
        {
            tmrRunOnce = new ccrpTimers6.ccrpTimer();
        }// On Error GoTo Error — nested handler not restructured
        if (App.StartMode != vbSModeStandalone)
        {
            tmrRunOnce.Interval = 500;
            tmrRunOnce.EventType = TimerOneShot;
            tmrRunOnce.Enabled = true;
        }

        mgvlOldNozzle = -1;
        mopBlockageCorr = 1;
        goto Done;
        Error();
        sRaiseErrorText = ex.Message;
        bRaiseError = true;
        if (ex.HResult & vbObjectError != vbObjectError)
        {
            EH.Module = TypeName(this);
            EH.Procedure = "Class_Initialize";
            EH_MsgBox();
        }

        goto Done;
    Error_CreatingObjects:
        ;
        sRaiseErrorText = string.Concat(LoadResString(ridCantCreateObject, sClassName), ex.Message);
        bRaiseError = true;
        EH.Module = TypeName(this);
        EH.Procedure = "Class_Initialize";
        EH.Item(1) = sClassName;
        EH_MsgBox();
        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        return;
    }

    private void Class_Terminate()
    {
        try
        {
            tmrMain.Enabled = false;
            tmrMain.Notify = null;
            tmrMain = null;
            mobjInterlocks.Quit();
            mobjInterlocks = null;
            tmrElapsed = null;
            mobjStatuses.Quit();
            mobjStatuses = null;
            mobjCommands.Clear("Terminate");
            mobjCommands = null;
            mobjConstants = null;
            mobjGlobalVars = null;
            mobjEvents = null;
            mobjStats = null;
            mcolInvalidOpcodes = null;
            mobjMFPermComponents.Release();
            mobjMFPermComponents = null;
            mobjMFAuxPermComponents.Release();
            mobjMFAuxPermComponents = null;
            mobjIdleCmdPermComponents.Release();
            mobjIdleCmdPermComponents = null;
            mobjParManager.Release();
            mobjParManager = null;
            mobjSPGConstants = null;
            mobjFFLeadLag = null;
            mobjForm = null;
            mobjFSMonitor = null;
            mobjPID = null;
            mobjWSctlFilter = null;
            mobjWSFilter = null;
            mobjWSFilter2 = null;
            mobjWSMonitor = null;
            mobjTrackFilter = null;
            mobjRampRateFilter = null;
            mobjOutputFilter = null;
            mobjTrackPar = null;
            mbInited = false;
            return;
        }
        catch
        {
            // On Error Resume Next: exceptions suppressed
        }
    }

    private void ICcrpTimerNotify_Timer(int Milliseconds)
    {
        int lElapsed = 0;
        try
        {
            ControlLogic();
            if (mbKeepStats & mbInited)
            {
                lElapsed = tmrElapsed.Elapsed;
                tmrElapsed.Reset();
                mobjStats.AddValue((double)lElapsed / 1000);
            }

            DoEvents();
            return;
        }
        catch
        {
            // On Error Resume Next: exceptions suppressed
        }
    }

    private void mobjParManager_BeforeUnloaded(ref string ConfigurationName, ref bool Cancel, ref Collection Cancellers)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        if (mbInited == true)
        {
            Cancel = true;
            Cancellers.Add(PROCESS_NAME);
        }

        goto Done;
        Error();
        if (ex.HResult & vbObjectError != vbObjectError)
        {
            EH.Module = TypeName(this);
            EH.Procedure = PROC_NAME;
            EH_MsgBox();
        }

        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        return;
    }

    private void mobjStatuses_StatusChanged(ref string Message)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        LogMessage(Message, mtStatusChange);
        goto Done;
        Error();
        if (ex.HResult & vbObjectError != vbObjectError)
        {
            EH.Module = TypeName(this);
            EH.Procedure = PROC_NAME;
            EH_MsgBox();
        }

        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        return;
    }

    private void mobjCommands_MessageFlushed(ref string MessageText, ref int MessageType)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        LogMessage(MessageText, MessageType);
        goto Done;
        Error();
        if (ex.HResult & vbObjectError != vbObjectError)
        {
            EH.Module = TypeName(this);
            EH.Procedure = "mobjCommands_MessageFlushed";
            EH_MsgBox();
        }

        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        return;
    }

    private void mobjCommands_CommandCompleted(ref CommandListLib.Command Command, ref CommandListLib.CompletionStatuses CompletionStatus, ref string CompletionMessage)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        switch (CompletionStatus)
        {
            case cmdAborted:
            case cmdTimedout:
                if (Command.SuppressLogging == false)
                {
                    objEH = new EH();
                    objEH.AppError = AppErrorText(tceCommandFailed, Command.Text, CompletionMessage);
                    objEH.Module = TypeName(this);
                    objEH.Procedure = PROC_NAME;
                    objEH.Component = App;
                    Command.SuppressLogging = true;
                    switch (Command.Opcode)
                    {
                        case ocStartMainFanAux:
                            LogPopupMessage(objEH, Command.Text);
                            break;
                        default:
                            LogMessage2(objEH);
                            break;
                    }
                }

                break;
        }

        switch (Command.Opcode)
        {
            case ocStartIdle:
            case ocStopIdle:
                if (CompletionStatus != cmdCompleted)
                {
                    mopIdleStartControl = 0;
                    mopIdleStopControl = 0;
                }

                break;
            case ocResetFaults:
                mopMFReset = 0;
                break;
        }

        goto Done;
        Error();
        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        objEH = null;
        return;
    }

    private void tmrRunOnce_Timer(int Milliseconds)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        tmrRunOnce.Enabled = false;
        tmrRunOnce = null;
        Start();
        goto Done;
        Error();
        if (ex.HResult & vbObjectError != vbObjectError)
        {
            EH.Module = TypeName(this);
            EH.Procedure = PROC_NAME;
            EH_MsgBox();
        }

        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        return;
    }

    public Events Events
    {
        get
        {
            Events __result = 0;
            __result = mobjEvents;
            return __result;
        }
    }

    public object InvalidOpcodes(ref object Opcode)
    {
        object __result = null;
        string sIndex = "";
        try
        {
            sIndex = Opcode;
            __result = mcolInvalidOpcodes(sIndex);
            if (Err)
            {
                __result = 0;
            }

            return __result;
        }
        catch
        {
            // On Error Resume Next: exceptions suppressed
        }

        return __result;
    }

    public void CheckInterlock(ref object Opcode)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        // On Error Resume Next — nested handler; see enclosing try/catch
        if (Opcode is object[])
        {
            sIndex = Opcode(0);
            mvCheckIlckValue = Opcode(0 + 1);
        }
        else
        {
            sIndex = Opcode;
            mvCheckIlckValue = null;
        }

        objInterlock = mobjInterlocks(sIndex);
        if (Err)
        {
            goto Done;
        }// On Error GoTo Error — nested handler not restructured
        objInterlock.Check();
        goto Done;
        Error();
        sRaiseErrorText = ex.Message;
        bRaiseError = true;
        lRaiseErrorNumber = ex.HResult;
        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        if (bRaiseError)
        {
            EH_RaiseError(lRaiseErrorNumber, TypeName(this), PROC_NAME, sRaiseErrorText);
        }

        return;
    }

    public Stats Stats
    {
        get
        {
            Stats __result = 0;
            __result = mobjStats;
            return __result;
        }
    }

    public CommandListLib.Commands Commands
    {
        get
        {
            CommandListLib.Commands __result = 0;
            // TODO: On Error GoTo Error — handler label not found in this scope
            bRaiseError = false;
            __result = mobjCommands;
            goto Done;
            Error();
            sRaiseErrorText = ex.Message;
            bRaiseError = true;
            if (ex.HResult & vbObjectError != vbObjectError)
            {
                EH.Module = TypeName(this);
                EH.Procedure = PROC_NAME;
                EH_MsgBox();
            }

            goto Done;
        Done:
            ;
            // On Error Resume Next — nested handler; see enclosing try/catch
            if (bRaiseError)
            {
                EH_RaiseError(OLEERR_PROPERTY_GET_FAILED, TypeName(this), PROC_NAME, sRaiseErrorText)();
            }

            return __result;
        }
    }

    private void WSControlLogic()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        mgvbUsingCloseTime = false;
        mgvlWSControlMode = mopWsMode;
        mgvbWSComingBackToSetpt = false;
        if (mgvlWindTunnel == wtCWT)
        {
            if (mopWsMode == modeTrack & mccbAutoFollowDyTarget)
            {
                if (mgvsTrackParName == mccsDefTrackParName)
                {
                    if (mipDyMode == DY_MODE_SPEED | mipThrMode == THR_MODE_SPEED)
                    {
                        mgvlWSControlMode = modeAuto;
                        fSpeedTgt = Ctl_Limit(mipDySpeedTarget + mgvfHeadwind, mccfMinWs, mgvfMaxWs);
                        fSpeedRate = Ctl_Limit(mipDySpeedRate, mccfMinWsRamp, mccfMaxWsRamp);
                        if (mgvfWsTarget != fSpeedTgt | mgvfWsRamp != fSpeedRate)
                        {
                            mgvfWsTarget = fSpeedTgt;
                            mgvfWsTarget = Ctl_Limit(mgvfWsTarget, mccfMinWs, mgvfMaxWs);
                            mgvfWsRamp = fSpeedRate;
                            if (mgvfWsRamp <= 0)
                            {
                                mgvfWsRamp = mccfDefWsRamp;
                            }

                            if (mcclSetpointRampMethod == 1)
                            {
                                mobjRampRateFilter.ClearData();
                                mgvfControlRamp = mobjRampRateFilter.FilteredValue(mgvfWsRamp);
                            }
                        }
                    }
                }
            }
        }

        switch (mgvlWSControlMode)
        {
            case modeAuto:
            case modeTrack:
                if (mgvlWSControlMode == modeTrack)
                {
                    // On Error Resume Next — nested handler; see enclosing try/catch
                    mgvfWsTarget = mobjTrackPar.Value * mgvfTrackC1 + mgvfTrackC0 + mgvfHeadwind;
                    mgvfWsTarget = Ctl_Limit(mgvfWsTarget, mccfMinWs, mgvfMaxWs);
                    mobjTrackFilter.FilterConst(0) = mccfWsTgtTrackFilTc;
                    mgvfWsTarget = mobjTrackFilter.FilteredValue(mgvfWsTarget);
                    if (Err)
                    {
                        LogMessage(LoadResString(ridTrackFailed, ex.Message), mtFailure);
                        SetMode(modeAuto);
                    }// On Error GoTo Error — nested handler not restructured
                    mgvfControlRamp = mgvfWsRamp;
                }
                else
                {
                    mobjWSMonitor.Target = mgvfWsTarget;
                    if (mopWsMode != mgvfOldCtlMode)
                    {
                        mobjWSMonitor.Reset();
                    }

                    switch (mcclSetpointRampMethod)
                    {
                        case 1:
                            if (mgvlNozzle == nsLarge)
                            {
                                mgvfCloseTime = TableLookup(mopWsSetpt, mccfCloseTimeSpeedsLarge, mccfCloseTimesLarge, mcclNumCloseTimes(mgvlNozzle));
                                mgvfCloseTimeRamp = mccfCloseWsRamp(nsSmall);
                            }
                            else if (mgvlNozzle == nsMedium)
                            {
                                mgvfCloseTime = TableLookup(mopWsSetpt, mccfCloseTimeSpeedsMedium, mccfCloseTimesMedium, mcclNumCloseTimes(mgvlNozzle));
                                mgvfCloseTimeRamp = mccfCloseWsRamp(nsMedium);
                            }
                            else
                            {
                                mgvfCloseTime = TableLookup(mopWsSetpt, mccfCloseTimeSpeedsSmall, mccfCloseTimesSmall, mcclNumCloseTimes(mgvlNozzle));
                                mgvfCloseTimeRamp = mccfCloseWsRamp(nsSmall);
                            }

                            if (mgvfWsTarget > 0 & mgvfWsRamp > mgvfCloseTimeRamp & Math.Abs(mgvfWsTarget - mopWsSetpt) / mgvfWsRamp <= mgvfCloseTime)
                            {
                                mobjRampRateFilter.FilterConst(0) = mgvfCloseWsFilTc;
                                mgvfControlRamp = mobjRampRateFilter.FilteredValue(mgvfCloseTimeRamp);
                                mgvbUsingCloseTime = true;
                            }
                            else
                            {
                                mobjRampRateFilter.ClearData();
                                mgvfControlRamp = mobjRampRateFilter.FilteredValue(mgvfWsRamp);
                            }

                            break;
                        case 2:
                            if (mgvlWindTunnel == wtAAWT & mgvbHoldCRUntilFSOK == true)
                            {
                                if (mopWindspeed > mopWsSetpt)
                                {
                                    mgvfControlRamp = mgvfWsRamp;
                                }
                            }
                            else
                            {
                                mgvfControlRamp = mgvfWsRamp;
                            }

                            if (Math.Abs(mgvfWsTarget - IIf(mccbFinalRampUseSetpoint, mopWsSetpt, mopWsControl)) <= mccfFinalRampDeltakph & mgvfWsRamp > mccfWsRateForFinalRamp & mgvbFinalRampOKToUse)
                            {
                                mgvbUsingFinalRamp = true;
                                if (mgvlNozzle == nsSmall)
                                {
                                    mgvfControlRamp = mccfFinalRampDeltakphSmall / mccfFinalRampTimeSecSmall;
                                }
                                else
                                {
                                    mgvfControlRamp = mccfFinalRampDeltakph / mccfFinalRampTimeSec;
                                }

                                if (mopWsControl > mopWsSetpt & mobjPID.States(ctlDerivative) > 0 & mopWsSetpt < mgvfWsTarget)
                                {
                                    mgvbWSComingBackToSetpt = true;
                                }
                                else if (mopWsControl < mopWsSetpt & mobjPID.States(ctlDerivative) < 0 & mopWsSetpt > mgvfWsTarget)
                                {
                                    mgvbWSComingBackToSetpt = true;
                                }
                            }

                            break;
                        default:
                            mgvfControlRamp = mgvfWsRamp;
                            break;
                    }
                }

                if (mopWsMode == modeTrack & mgvbAutoIdle)
                {
                    if (mgvfWsTarget >= mccfWSStartIdle + mccfAutoIdleDeadband)
                    {
                        mgvbAutoIdleWantsIdle = false;
                        mgvbStopAutoIdle = false;
                        mgvbAutoIdleSpeedHasBeenHigh = true;
                        mgvbAutoIdleSpeedHasBeenLow = false;
                    }
                    else if (mgvfWsTarget >= mccfWSStartIdle)
                    {
                    }
                    else if (mgvfWsTarget >= mccfWSStopIdle)
                    {
                        if (mgvbStopAutoIdle)
                        {
                            mgvbAutoIdleWantsIdle = false;
                        }
                        else if (mgvbAutoIdleSpeedHasBeenHigh)
                        {
                            mgvbAutoIdleWantsIdle = true;
                            mgvbAutoIdleSpeedHasBeenLow = false;
                        }
                        else if (mgvbAutoIdleSpeedHasBeenLow)
                        {
                            mgvbAutoIdleWantsIdle = false;
                            mgvbStopAutoIdle = false;
                            mgvbAutoIdleSpeedHasBeenHigh = false;
                        }
                        else
                        {
                        }
                    }
                    else if (mgvfWsTarget >= IIf(mccfWSStopIdle - mccfAutoIdleDeadband > 0.1, mccfWSStopIdle - mccfAutoIdleDeadband, 0.1))
                    {
                    }
                    else
                    {
                        if (mgvbStopAutoIdle)
                        {
                            mgvbAutoIdleWantsIdle = false;
                        }
                        else
                        {
                            mgvbAutoIdleWantsIdle = true;
                        }

                        mgvbAutoIdleSpeedHasBeenLow = true;
                        mgvbAutoIdleSpeedHasBeenHigh = false;
                    }

                    mgvfOldWsTarget = mgvfWsTarget;
                    if (mgvbAutoIdleWantsIdle)
                    {
                        if (mopIdleStatus == isIdle)
                        {
                            mgvfFsTarget = mccfIdleFanspeed;
                            mgvfFsRamp = mccfMaxFsRamp;
                        }
                        else
                        {
                            mgvfFsTarget = 0;
                            mgvfFsRamp = mccfMaxFsRamp;
                        }

                        goto ManualMode;
                    }
                    else if (mopIdleStatus != isNormal)
                    {
                        mgvfWsTarget = Ctl_Limit(mgvfWsTarget, 0, mccfMaxFsLimit(1) * mgvfNomWsMax / mgvfNomFsMax);
                    }
                }
                else if (mgvbAutoIdle == false)
                {
                    mgvbAutoIdleSpeedHasBeenHigh = false;
                    mgvbAutoIdleSpeedHasBeenLow = false;
                }

                mgvfWSTargetLimited = mgvfWsTarget;
                for (I = 1; I <= 3; I++)
                {
                    if (LimitInEffect(I))
                    {
                        if (mgvfWSTargetLimited > mccfMaxFsLimit(I) * mgvfNomWsMax / mgvfNomFsMax)
                        {
                            mgvfWSTargetLimited = mccfMaxFsLimit(I) * mgvfNomWsMax / mgvfNomFsMax;
                            mobjRampRateFilter.ClearData();
                            mgvfControlRamp = mobjRampRateFilter.FilteredValue(mgvfWsRamp);
                        }
                    }
                }

                if (mgvlWindTunnel == wtAAWT)
                {
                    if (mgvfWSTargetLimited > mopWsSetpt)
                    {
                        if (mipFanspeed <= mccfSlowRampFanspeed)
                        {
                            if (mgvfControlRamp > mccfMaxWsRampLowSpeeds)
                            {
                                mgvfControlRamp = mccfMaxWsRampLowSpeeds;
                                mgvbFinalRampOKToUse = false;
                                mgvbHoldCRUntilFSOK = true;
                            }
                        }
                    }
                }

                mopWsSetpt = Ctl_Ramp(mgvfWSTargetLimited, mopWsSetpt, mgvfControlRamp, mccfUpdateInt, mgvfWsIncrement);
                if (mccbFinalRampSnap & mgvbUsingFinalRamp & mgvbWSComingBackToSetpt & IIf(Math.Abs(mgvfWsTarget - mopWsControl) <= mccfUseTargetErrorForDerivLimit, true, false) & mopWsMode == modeAuto)
                {
                    if (mgvbFinalRampHasSnapped == false)
                    {
                        if (mccbFinalRampSnapBumpless)
                        {
                            mobjPID.States(ctlIntegral) = mobjPID.States(ctlIntegral) - mgvfKp * mgvfWSTargetLimited - mopWsSetptDisp / mgvfNomWsMax;
                        }
                    }

                    mopWsSetpt = mgvfWSTargetLimited;
                    mgvbFinalRampHasSnapped = true;
                }

                if (mgvbUsingCloseTime | mgvbUsingFinalRamp | mgvbFinalRampHasSnapped)
                {
                    mopWsSetptDisp = Ctl_Ramp(mgvfWSTargetLimited, mopWsSetptDisp, mgvfWsRamp, mccfUpdateInt, fIncrement);
                }
                else
                {
                    mopWsSetptDisp = mopWsSetpt;
                }

                if (mopWsSetpt < 0)
                {
                    mgvfFsTarget = mopWsSetpt / mccfNomWSReverse * mccfNomFSReverse;
                    mgvfFsTarget = Ctl_Limit(mgvfFsTarget, mccfMinFs, 0);
                    mgvfFsLimitedTarget = mgvfFsTarget;
                    for (I = 1; I <= 3; I++)
                    {
                        if (LimitInEffect(I))
                        {
                            mgvfFsLimitedTarget = Ctl_Limit(mgvfFsLimitedTarget, -mccfMaxFsLimit(I), mccfMaxFsLimit(I));
                        }
                    }
                }
                else
                {
                    mgvfNormWsSetpt = mopWsSetptDisp / mgvfNomWsMax;
                    if (mopBlockageCorr != 0)
                    {
                        mgvfFfRawNormFs = mccfFfGain * mgvfNormWsSetpt / Math.Sqrt(mopBlockageCorr);
                    }
                    else
                    {
                        mgvfFfRawNormFs = mccfFfGain * mgvfNormWsSetpt;
                    }

                    if (mgvlWSControlMode == modeTrack)
                    {
                        mgvfFfLeadCoeff = mccfFfLeadCoefTrack;
                        mgvfFfLagCoeff = mccfFfLagCoefTrack;
                    }
                    else
                    {
                        mgvfFfLeadCoeff = mccfFfLeadCoef;
                        mgvfFfLagCoeff = mccfFfLagCoef;
                    }

                    mobjFFLeadLag.FilterConst(0) = mgvfFfLeadCoeff;
                    mobjFFLeadLag.FilterConst(1) = mgvfFfLagCoeff;
                    mobjFFLeadLag.UpdateTime = mccfUpdateInt;
                    mgvfFfNormFs = mobjFFLeadLag.FilteredValue(mgvfFfRawNormFs);
                    if (mgvfOldCtlMode < modeAuto | mgvlOldIdleStatus != isNormal)
                    {
                        mgvfFfNormFs = mgvfFfRawNormFs;
                        mobjPID.States(ctlProportional) = 0;
                        mobjPID.States(ctlIntegral) = mopFsSetpt / mgvfNomFsMax - mgvfFfNormFs;
                        mobjPID.States(ctlDerivative) = 0;
                    }

                    if (IsWSTooLowForPID | mopIdleStatus != isNormal)
                    {
                        mobjPID.States(ctlIntegral) = 0;
                        if (mgvlWSControlMode == modeTrack & mgvbAutoIdle & mgvbAutoIdleWantsIdle)
                        {
                        }
                        else if (mgvlWSControlMode != modeTrack)
                        {
                            mgvfFbNormFs = 0;
                        }
                        else if (mgvfWsTarget > mopWsControl & mgvfFbNormFs < 0)
                        {
                            mgvfFbNormFs = 0;
                        }
                        else if (mgvfFbNormFs < -mgvfFfNormFs)
                        {
                            mgvfFbNormFs = -mgvfFfNormFs;
                        }
                    }
                    else
                    {
                        mgvfWsError = mopWsSetpt - mopWsControl;
                        mgvfWsError = Ctl_Limit(mgvfWsError, -mccfMaxWsError, mccfMaxWsError);
                        mgvfNormError = mgvfWsError / mgvfNomWsMax;
                        mgvfWSTargetError = mgvfWSTargetLimited - mopWsControl;
                        mgvfTargetNormError = mgvfWSTargetError / mgvfNomWsMax;
                        mgvfKp = mgvfKp0 + mgvfKp1 * mgvfNormWsSetpt;
                        mgvfKp = Ctl_Limit(mgvfKp, mgvfMinKp, mgvfMaxKp);
                        mgvfTi = mgvfTi0 + mgvfTi1 * mgvfNormWsSetpt;
                        mgvfTi = Ctl_Limit(mgvfTi, mgvfMinTi, mgvfMaxTi);
                        mgvfTd = mgvfTd0 + mgvfTd1 * mgvfNormWsSetpt;
                        mgvfTd = Ctl_Limit(mgvfTd, mgvfMinTd, mgvfMaxTd);
                        mgvfOldIntegral = mobjPID.States(ctlIntegral);
                        mobjPID.Alpha = mccfTdFilterTC;
                        mobjPID.ControlMax = mccfCtlMax;
                        mobjPID.ControlMin = mccfCtlMin;
                        mobjPID.IntegralMax = mccfIntMax;
                        mobjPID.IntegralMin = mccfIntMin;
                        mobjPID.Kp = mgvfKp;
                        mobjPID.Td = mgvfTd;
                        mobjPID.Ti = mgvfTi;
                        mobjPID.UpdateTime = mccfUpdateInt;
                        if (mopWsMode > modeAuto)
                        {
                            mgvfFbNormFs = mobjPID.Output(mgvfNormError);
                        }
                        else
                        {
                            mgvfFbNormFs = mobjPID.Output2(mgvfNormError, mgvfTargetNormError, IIf(Math.Abs(mgvfWSTargetError) <= mccfUseTargetErrorForDerivLimit, true, false));
                        }

                        mgvfWsRateForI = mccfWsRateForI0 + mccfWsRateForI1 * mgvfNormWsSetpt;
                        mgvfWsRateForI = Ctl_Limit(mgvfWsRateForI, mccfMinWsRateForI, mccfMaxWsRateForI);
                        if (Math.Abs(mgvfWsIncrement / mccfUpdateInt) > mgvfWsRateForI)
                        {
                            mobjPID.States(ctlIntegral) = mgvfOldIntegral;
                        }
                    }

                    mgvfFsNormSetpt = mgvfFfNormFs + mgvfFbNormFs;
                    mgvfFsTarget = mgvfNomFsMax * mgvfFsNormSetpt;
                    mgvfFsTarget = Ctl_Limit(mgvfFsTarget, mccfMinFsAutoMode, mgvfMaxFs);
                    mgvfFsLimitedTarget = mgvfFsTarget;
                    for (I = 1; I <= 3; I++)
                    {
                        if (LimitInEffect(I))
                        {
                            mgvfFsLimitedTarget = Ctl_Limit(mgvfFsLimitedTarget, mccfMinFsAutoMode, mccfMaxFsLimit(I));
                        }
                    }
                }

                if (mopWsMode == modeTrack)
                {
                    mopFsSetpt = Ctl_Ramp(mgvfFsLimitedTarget, mopFsSetpt, mccfMaxFsRampTrack, mccfUpdateInt, mgvfFsIncrement);
                }
                else
                {
                    mopFsSetpt = Ctl_Ramp(mgvfFsLimitedTarget, mopFsSetpt, mccfMaxFsRamp, mccfUpdateInt, mgvfFsIncrement);
                }

                if (mopWsSetpt >= 0)
                {
                    if (mgvfFsLimitedTarget < mgvfFsTarget)
                    {
                        if (mgvfNormError > 0)
                        {
                            mobjPID.States(ctlIntegral) = mgvfOldIntegral;
                        }
                    }
                    else if (mgvfFsLimitedTarget > mgvfFsTarget)
                    {
                        if (mgvfNormError < 0)
                        {
                            mobjPID.States(ctlIntegral) = mgvfOldIntegral;
                        }
                    }
                }

                if (mccbFilterOutput)
                {
                    mobjOutputFilter.FilterConst(0) = mccfPidOutputFilTc;
                    mopFsSetpt = mobjOutputFilter.FilteredValue(mopFsSetpt);
                }

                if (mopWsSetpt >= 0)
                {
                    mopFsSetpt = Ctl_Limit(mopFsSetpt, mccfMinFsAutoMode, mgvfMaxFs);
                }
                else
                {
                    mopFsSetpt = Ctl_Limit(mopFsSetpt, mccfMinFs, mgvfMaxFs);
                }

                break;
            case modeManual:
                mobjFSMonitor.Target = mgvfFsTarget;
                if (mopWsMode != mgvfOldCtlMode)
                {
                    mobjFSMonitor.Reset();
                }

            ManualMode:
                ;
                mgvfNormWsSetpt = 0.5;
                if (mipMfStatus == 0 | mopMfControl == 0)
                {
                    mgvfFsTarget = 0;
                    mopFsSetpt = mgvfFsTarget;
                    mgvfWsTarget = 0;
                    mopWsSetpt = mgvfWsTarget;
                }
                else
                {
                    mgvfWsTarget = mopWindspeed;
                    mopWsSetpt = mgvfWsTarget;
                    mgvfFsTarget = Ctl_Limit(mgvfFsTarget, mccfMinFs, mgvfMaxFs);
                    mgvfFsLimitedTarget = mgvfFsTarget;
                    for (I = 1; I <= 3; I++)
                    {
                        if (LimitInEffect(I))
                        {
                            mgvfFsLimitedTarget = Ctl_Limit(mgvfFsLimitedTarget, mccfMinFs, mccfMaxFsLimit(I));
                        }
                    }

                    mopFsSetpt = Ctl_Ramp(mgvfFsTarget, mopFsSetpt, mgvfFsRamp, mccfUpdateInt, mgvfFsIncrement);
                }

                if (mgvfFsTarget == 0 & mopFsSetpt == 0)
                {
                    if (mccfTurnOffFanTime > 0)
                    {
                        if (SecondsSinceTime(mgvdTurnOffFanTime) > mccfTurnOffFanTime)
                        {
                            if (mopMfControl == 1)
                            {
                                if (mccfMainFanAuxAutoOffTime > 0)
                                {
                                    mgvbAuxAutoOff = true;
                                }
                            }

                            mopMfControl = 0;
                        }
                    }
                }
                else
                {
                    mgvdTurnOffFanTime = DateTime.Now;
                }

                if (mipMfStatus == 0 & mopMfControl == 0 & mopAuxControl == 1)
                {
                    if (mccfMainFanAuxAutoOffTime > 0 & mgvbAuxAutoOff & mgvdDriveOffTime != 0)
                    {
                        if (SecondsSinceTime(mgvdDriveOffTime) >= mccfMainFanAuxAutoOffTime * 60)
                        {
                            mopAuxControl = 0;
                        }
                    }
                }

                break;
            default:
                mgvfNormWsSetpt = 0.5;
                mopMfControl = 0;
                mgvfWsTarget = 0;
                mopWsSetpt = 0;
                mgvfFsTarget = 0;
                mopFsSetpt = 0;
                break;
        }

        mgvfKpEffect = mobjPID.States(ctlProportional);
        mgvfDeriviativeEffect = mobjPID.States(ctlDerivative);
        mgvfIntegratorEffect = mobjPID.States(ctlIntegral);
        goto Done;
        Error();
        EH.AppError = LoadResString(stsEHErrorExecutingControlLogic);
        EH.Module = TypeName(this);
        EH.Procedure = PROC_NAME;
        EH_MsgBox();
        goto ErrorExit;
    ErrorExit:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        Suspend();
        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        return;
    }

    private bool LimitInEffect(ref int LimitNumber)
    {
        bool __result = false;
        __result = false;
        if (mgvlWindTunnel == wtCWT)
        {
            switch (LimitNumber)
            {
                case 1:
                    if (mopIdleStatus != IdleStatuses.isNormal)
                    {
                        __result = true;
                    }

                    break;
            }
        }

        return __result;
    }

    private bool IsWSTooLowForPID()
    {
        bool __result = false;
        if (mopWsSetpt <= mccfMinWsForPid | mopWsControl <= mccfMinWsForPid & mgvlWSControlMode == modeTrack | mopWsControl <= mccfMinWsForPid & mopWsSetpt < mopWsControl)
        {
            if (mgvbWsTooLowForPID == false)
            {
                mgvdWsTooLowForPIDTime = Now;
            }

            mgvbWsTooLowForPID = true;
        }
        else
        {
            mgvbWsTooLowForPID = false;
        }

        if (mgvbWsTooLowForPID)
        {
            if (mccfMinFsAutoMode >= 0)
            {
                __result = true;
            }
            else if (SecondsSinceTime(mgvdWsTooLowForPIDTime) >= mccfMinWsForPidDelay)
            {
                __result = true;
            }
            else
            {
                __result = false;
            }
        }
        else
        {
            __result = false;
        }

        return __result;
    }

    private void SetMode(ref Modes vNewValue)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        if (mopWsMode == vNewValue)
        {
            goto Done;
        }

        switch (vNewValue)
        {
            case modeTrack:
            case modeAuto:
                if (mopWsMode < modeAuto)
                {
                    mobjFFLeadLag.ClearData();
                    mobjPID.States(ctlIntegral) = 0;
                }

                break;
        }

        switch (vNewValue)
        {
            case modeTrack:
                mopWsMode = modeTrack;
                mgvfWsRamp = mccfMaxWsRamp;
                mgvfControlRamp = mgvfWsRamp;
                break;
            case modeAuto:
                mobjTrackPar = null;
                mopWsMode = modeAuto;
                if (Math.Abs(mccfFfGain * mopWindspeed / mgvfNomWsMax - mipFanspeed / mgvfNomFsMax) <= mccfIntMax)
                {
                    mgvfWsTarget = mopWindspeed;
                }
                else
                {
                    mgvfWsTarget = mgvfNomWsMax * mopFsSetpt / mgvfNomFsMax / mccfFfGain;
                }

                mgvfWsRamp = mccfDefWsRamp;
                mopWsSetpt = mgvfWsTarget;
                mgvfFsRamp = mccfMaxFsRamp;
                mgvfControlRamp = mgvfWsRamp;
                break;
            case modeManual:
                mobjTrackPar = null;
                mopWsMode = modeManual;
                if (mipMfStatus == 1)
                {
                    mgvfFsTarget = mipFanspeed;
                    mopFsSetpt = mipFanspeed;
                    mopMfControl = 1;
                }
                else
                {
                    mgvfFsTarget = 0;
                    mopFsSetpt = 0;
                    mopMfControl = 0;
                }

                mgvfFsRamp = mccfDefFsRamp;
                break;
            default:
                mobjTrackPar = null;
                mopWsMode = modeOff;
                mgvfFsTarget = 0;
                mopFsSetpt = 0;
                mgvfWsTarget = 0;
                mopWsSetpt = 0;
                mopMfControl = 0;
                break;
        }

        goto Done;
        Error();
        sRaiseErrorText = ex.Message;
        bRaiseError = true;
        if (ex.HResult & vbObjectError != vbObjectError)
        {
            EH.Module = TypeName(this);
            EH.Procedure = PROC_NAME;
            EH_MsgBox();
        }

        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        if (bRaiseError)
        {
            EH_RaiseError(OLEERR_PROPERTY_GET_FAILED, TypeName(this), PROC_NAME, sRaiseErrorText)();
        }

        return;
    }
}