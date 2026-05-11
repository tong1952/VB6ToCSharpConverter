// Converted from WsControl.Bas by vb6cs
// Date: 2026-05-11 03:43

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace Converted;
public static class WsControl
{
    public enum ResourceIDs
    {
        ridAuxCurrentOff = 112,
        ridMFCurrentOff = 113,
        ridIdleCurrentOn = 114,
        ridMFCurrentOn = 115,
        ridModeManIlck = 116,
        ridModeTrackIlck = 118,
        ridIdleNozFlapNotOpen = 119,
        ridIdleBypFlapNotClosed = 120,
        ridFanStoppedIlck = 121,
        ridFanDriveSafetyFault = 124,
        ridFanDriveFault = 125,
        ridFanDriveVibWarning = 126,
        ridFanDriveVibFault = 127,
        ridFanDriveTempFault = 128,
        ridLowSpeedLmt1On = 129,
        ridLowSpeedLmt2On = 130,
        ridLowSpeedLmt3On = 131,
        ridModeAuto = 132,
        ridModeManual = 133,
        ridModeTrack = 134,
        ridMfTrippedOut = 135,
        ridWsTooLowForPid = 139,
        ridHxDpOverQFault = 140,
        ridHxDpFault = 141,
        ridCommonInterlockFaulted = 153,
        ridFanspeedAchieved = 155,
        ridIdleStarted = 156,
        ridIdleStopped = 157,
        ridMainFanAuxOff = 158,
        ridMainFanAuxOn = 159,
        ridMainfanOff = 160,
        ridMainFanOn = 161,
        ridWindspeedAchieved = 162,
        ridWindspeedInitialized = 163,
        ridFanspeedNotAchieved = 164,
        ridWindspeedNotAchieved = 165,
        ridWindspeedNotInitialized = 166,
        ridCantCreateObject = 167,
        ridRunning = 168,
        ridRunningAt = 169,
        ridCommandsClassCaption = 170,
        ridErrDisplayingMessage = 174,
        ridPIDIntegralAtMax = 177,
        ridPIDIntegralAtMin = 178,
        ridFanSpeedStillAbove = 179,
        ridTrackParInvalidUnits = 180,
        ridTrackFailed = 185,
        ridIdlePermissiveNotOn = 201,
        ridFanSystemFault = 203,
        ridEstop = 204,
        ridFanNotinComputerCtl = 205,
        ridAuxNotReady = 206,
        ridAuxPerm = 207,
        ridAuxFault = 208,
        ridLightCurtain = 209,
        ridIdleDriveFault = 210,
        ridFanLCFault = 211,
        ridFanMCFault = 212,
        ridCSDidNotStart = 213,
        ridMCDidNotGetReady = 214,
        ridLCDidNotGetReady = 215,
        ridErrorLoadingWSCoeffs = 216,
        ridErrorOpeningFile = 217,
        ridErrorReadingData = 218,
        ridMainFanRestartDelay = 219,
        ridPurging = 220,
        ridFanIsOnStopFanFirst = 221,
        ridUnableToGetControlFromPLC = 222,
        ridIdleMotorStopOn = 223,
        ridWindspeedReduced = 224,
        ridFanspeedReduced = 225,
        ridCFLeftFault = 226,
        ridCFLeftLockout = 227,
        ridCFRightFault = 228,
        ridCFRightLockout = 229,
        ridCFTopFault = 230,
        ridCFTopLockout = 231
    }

    public const object PROCESS_NAME = "Windspeed";
    public const object PROCESS_DESCRIPTION = "Windspeed Controller";
    public const object REG_PROCESSES = "SOFTWARE\\React Technologies\\Talent\\4.0\\Processes";
    public const object REG_WINDSPEED = "SOFTWARE\\React Technologies\\Talent\\4.0\\Windspeed";
    public const object SOM_CONTROLLER = "Controller";
    public const object SPG_CONSTANTS = "Constants";
    public const object SPG_DATA = "Data";
    public enum SPGConstants
    {
        CC_FSMin = 0,
        CC_FSMax = 1,
        CC_FSRampMax = 2,
        CC_WSMin = 3,
        CC_WSMax = 4,
        CC_WSRampMax = 5,
        CC_HeadwindMin = 6,
        CC_HeadwindMax = 7,
        CC_WSRampDef = 8,
        CC_FSRampDef = 9,
        CC_DefAutoIdle = 10,
        CC_DefTrackParName = 11,
        CC_CFLeft1Min = 12,
        CC_CFLeft1Max = 13,
        CC_CFLeft2Min = 14,
        CC_CFLeft2Max = 15,
        CC_CFRight1Min = 16,
        CC_CFRight1Max = 17,
        CC_CFRight2Min = 18,
        CC_CFRight2Max = 19,
        CC_CFTopMin = 20,
        CC_CFTopMax = 21
    }

    public enum Opcodes
    {
        ocInitialize = 1,
        ocSuspend = 2,
        ocShutdown = 3,
        ocShowStatus = 4,
        ocStartMainFanAux = 5,
        ocStopMainFanAux = 6,
        ocStartMainFan = 7,
        ocStopMainFan = 8,
        ocSetFanSpeed = 9,
        ocSetWindSpeed = 10,
        ocTrackWindSpeed = 11,
        ocStartIdle = 12,
        ocStopIdle = 13,
        ocShowIdleStatus = 14,
        ocCancelIdle = 15,
        ocSetModeManual = 20,
        ocSetModeAuto = 21,
        ocSetNozzleMethod = 22,
        ocSetPlenumMethod = 23,
        ocResetFaults = 30,
        ocStartPurge = 40,
        ocStopPurge = 41,
        ocSetNozzleSize = 42,
        ocStopNozzleMotion = 43,
        ocSetCFlapLeft = 50,
        ocSetCFlapLeft1 = 51,
        ocSetCFlapLeft2 = 52,
        ocSetCFlapRight = 53,
        ocSetCFlapRight1 = 54,
        ocSetCFlapRight2 = 55,
        ocSetCFlapTop = 56,
        ocStopCF = 57
    }

    public enum Modes
    {
        modeOff = 0,
        modeManual = 1,
        modeAuto = 2,
        modeTrack = 3
    }

    public const object SPG_STATUSES = "Statuses";
    public const object STS_FANSPEED_ACHIEVED = "FanspeedAchieved";
    public const object STS_IDLE_STARTED = "IdleStarted";
    public const object STS_IDLE_STOPPED = "IdleStopped";
    public const object STS_MAINFAN_AUX_ON = "MainfanAuxOn";
    public const object STS_MAINFAN_AUX_OFF = "MainfanAuxOff";
    public const object STS_MAINFAN_ON = "MainfanOn";
    public const object STS_MAINFAN_OFF = "MainfanOff";
    public const object STS_WINDSPEED_ACHIEVED = "WindspeedAchieved";
    public const object STS_WINDSPEED_INITIALIZED = "WindspeedInitialized";
    public enum IdleStatuses
    {
        isNormal = 0,
        isIdle = 1,
        isMoving = 2,
        isStopped = 3,
        isFault = 4
    }

    public enum NozzleSizes
    {
        nsMoving = 0,
        nsLarge = 1,
        nsMedium = 2,
        nsSmall = 3
    }

    public const object FAULT = 0;
    public const object NOFAULT = 1;
    public const double Rair = 287.06;
    public const double Rwater = 461.04;
    public const object DY_MODE_SPEED = 1;
    public const object THR_MODE_SPEED = 1;
    public enum FilterSpeeds
    {
        fsSLOW = 0,
        fsFAST = 1
    }

    public struct FinalRampCC
    {
        public float RampRate;
        public float TargetChangeBase;
        public float TargetChangeAdjust;
        public float Min;
        public float Max;
    }

    [DllImport("KERNEL32", EntryPoint = "GetCurrentProcessId")]
    public static extern int GetCurrentProcessId();
    [DllImport("KERNEL32", EntryPoint = "GetCurrentProcess")]
    public static extern int GetCurrentProcess();
    [DllImport("KERNEL32", EntryPoint = "TerminateProcess")]
    public static extern int TerminateProcess(int hProcess, int uExitCode);
    public static string NozzleName(ref int Index)
    {
        string __result = "";
        if (Index == 2)
        {
            __result = "Medium";
        }
        else if (Index == 3)
        {
            __result = "Small";
        }
        else
        {
            __result = "Large";
        }

        return __result;
    }

    public static float CalcFinalRampDeltakph(ref float RampRate, ref float TargetChange, ref int NumRampRates, ref FinalRampCC[] RampRateCCs)
    {
        float __result = 0;
        // TODO: On Error GoTo Error — handler label not found in this scope
        colRRs = new Collection();
        colIndex = new Collection();
        colRRs.Add(RampRateCCs(1).RampRate);
        colIndex.Add(1);
        for (I = 2; I <= NumRampRates; I++)
        {
            for (J = 1; J <= colRRs.Count; J++)
            {
                if (RampRateCCs(I).RampRate <= colRRs(J))
                {
                    colRRs.Add(RampRateCCs(I).RampRate, ,);
                    J();
                    colIndex.Add(I, ,);
                    J();
                    break;
                }
                else if (J == colRRs.Count)
                {
                    colRRs.Add(RampRateCCs(I).RampRate);
                    colIndex.Add(I);
                    break;
                }
            }
        }

        bRaiseError = false;
        for (J = 1; J <= NumRampRates; J++)
        {
            if (RampRate <= colRRs(J))
            {
                if (J == 1)
                {
                    fI = 1;
                    I = colIndex(J);
                    __result = Ctl_Limit(RampRateCCs(I).TargetChangeBase + RampRateCCs(I).TargetChangeAdjust * Math.Abs(TargetChange), RampRateCCs(I).Min, RampRateCCs(I).Max);
                    goto Done;
                }
                else
                {
                    I = colIndex(J);
                    Im1 = colIndex(J - 1);
                    fI = RampRate - RampRateCCs(Im1).RampRate / RampRateCCs(I).RampRate - RampRateCCs(Im1).RampRate;
                    DeltakphI = Ctl_Limit(RampRateCCs(I).TargetChangeBase + RampRateCCs(I).TargetChangeAdjust * Math.Abs(TargetChange), RampRateCCs(I).Min, RampRateCCs(I).Max);
                    DeltakphIm1 = Ctl_Limit(RampRateCCs(Im1).TargetChangeBase + RampRateCCs(Im1).TargetChangeAdjust * Math.Abs(TargetChange), RampRateCCs(Im1).Min, RampRateCCs(Im1).Max);
                    __result = fI * DeltakphI + 1 - fI * DeltakphIm1;
                    goto Done;
                }
            }
        }

        I = colIndex(NumRampRates);
        __result = Ctl_Limit(RampRateCCs(I).TargetChangeBase + RampRateCCs(I).TargetChangeAdjust * Math.Abs(TargetChange), RampRateCCs(I).Min, RampRateCCs(I).Max);
        goto Done;
        Error();
        sRaiseErrorText = ex.Message;
        bRaiseError = true;
        lRaiseErrorNumber = ex.HResult;
        if (ex.HResult & vbObjectError != vbObjectError)
        {
            EH.Module = "modWsControl";
            EH.Procedure = PROC_NAME;
            EH_MsgBox();
        }

        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        if (bRaiseError)
        {
            EH_RaiseError(lRaiseErrorNumber, "modWsControl", PROC_NAME, sRaiseErrorText);
        }

        return __result;
    }

    public static float TableLookup(float InputValue, ref float[] InputValues, ref float[] OutputValues, ref int ArrayMaxSize)
    {
        float __result = 0;
        // TODO: On Error GoTo Error — handler label not found in this scope
        __result = OutputValues(0);
        for (I = 0; I <= ArrayMaxSize; I++)
        {
            if (InputValue <= InputValues(I))
            {
                if (I == 0)
                {
                    __result = OutputValues(ArrayMaxSize);
                }
                else
                {
                    fFactor = InputValue - InputValues(I - 1) / InputValues(I) - InputValues(I - 1);
                    __result = OutputValues(I - 1) + fFactor * OutputValues(I) - OutputValues(I - 1);
                }

                break;
            }
            else if (I == ArrayMaxSize)
            {
                __result = OutputValues(I);
            }
        }

        goto Done;
        Error();
        goto Done;
    Done:
        ;
        return __result;
    }

    public static string LoadResString(ref int ID, params object[] ItemTextArgs)
    {
        string __result = "";
        if (ID & STD_TALENT_STRING_BASE == STD_TALENT_STRING_BASE)
        {
            __result = LoadStdTalentString(ID, ItemTextArgs);
        }
        else if (IsWithinIDE())
        {
            __result = VB.LoadResString(ID);
            __result = ReplaceItemText(LoadResString, ItemTextArgs);
        }
        else
        {
            __result = GetResString(App.hInstance, ID, ItemTextArgs);
        }

        return __result;
    }

    public static float Log10(ref float Number)
    {
        float __result = 0;
        __result = Math.Log(Number) / Math.Log(10);
        return __result;
    }

    public static double LimitD(ref double Expression, ref float Min, ref float Max)
    {
        double __result = 0;
        if (Expression < Min)
        {
            __result = Min;
        }
        else if (Expression > Max)
        {
            __result = Max;
        }
        else
        {
            __result = Expression;
        }

        return __result;
    }
}