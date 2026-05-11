// Converted from DYControl.Bas by vb6cs
// Date: 2026-05-11 03:23

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace Converted;
public static class DYControl
{
    private const object MOD_NAME = "modDyControl";
    public const object SOM_CONTROLLER = "Controller";
    public const object SPG_CONSTANTS = "Constants";
    public const object SPG_DATA = "Data";
    public enum ResIDs
    {
        ridDyModeStopped = 107,
        ridThrModeStopped = 113,
        ridNoResponseFromDyno = 118,
        ridDynoModeIsNow = 119,
        ridThrModeIsNow = 120,
        ridNoSpeedControl = 121,
        ridNoForceControl = 122,
        ridNoPositionControl = 123,
        ridInvalidDyCmdArgument = 124,
        ridErrorDyCmd = 125,
        ridDynoRejectedCmd = 132,
        ridBadDynoMode = 134,
        ridBadThrMode = 135,
        ridBadDynoThrMode = 136,
        ridVehicleIDSet = 137,
        ridBadDyConfig = 143,
        ridConnectionRefused = 145,
        ridSocketNotCreated = 146,
        ridSocketNotConnected = 147,
        ridSocketInUndefinedState = 148,
        ridDynoNotStopped = 149,
        ridThrottleNotStopped = 150,
        ridDynoMustBeIn4WD = 151,
        ridDynoMustBeIn4WDSpeedSyncOff = 152,
        ridDynoMustBeIn4WDForceViaVehOff = 153,
        ridDynoMustBeInSpeed = 154,
        ridDynoMustBeInForce = 155,
        ridDynoNotInRemote = 156,
        ridDynoNotInVehicleTestMode = 157,
        ridTCPCommsFaulted = 159,
        ridUDPCommsFaulted = 160,
        ridDynoError = 162,
        ridUDPCommsStarted = 163,
        ridSpeedSyncTimedOut = 164,
        ridCantSendTCPCommand = 165,
        ridWaitingForOfflineModes = 166,
        ridCmdTimedOut = 167,
        ridDyNoCarrVehs = 168,
        ridBadActuator = 169,
        ridRobotNotInUse = 170,
        ridNoActuatorSelected = 171,
        ridTCPIPSetupFail = 172,
        ridCmdNotSupported = 173
    }

    public enum CCNames
    {
        ccSpeedMax = 0,
        ccSpeedRampMax = 1,
        ccSpeedRampDefault = 2,
        ccK0Min = 3,
        ccK0Max = 4,
        ccK1Min = 5,
        ccK1Max = 6,
        ccK2Min = 7,
        ccK2Max = 8,
        ccK3Min = 9,
        ccK3Max = 10,
        ccGradientMin = 11,
        ccGradientMax = 12,
        ccGradientRateMax = 13,
        ccGradientRateDefault = 14,
        ccPosMax = 15,
        ccPosMin = 16,
        ccPosRampMax = 17,
        ccPosRampMin = 18,
        ccPosRampDefault = 19,
        ccForceRampMax = 20,
        ccForceRampMin = 21,
        ccForceRampDefault = 22,
        ccForceMax = 23,
        ccForceMin = 24,
        ccTRefMin = 25,
        ccTRefMax = 26,
        ccTRefDefault = 27,
        ccMaxCarrPosnPosition = 28,
        ccMinCarrPosnPosition = 29,
        ccMinRobotDCReps = 30,
        ccMaxRobotDCReps = 31
    }

    public enum SharedDataItems
    {
        sdDyMode = 0,
        sdThrMode = 1,
        sdDyConfig = 2,
        sdTranType = 3,
        sdRbtMode = 4,
        sdRbtInUse = 5
    }

    public enum DynoConfigs
    {
        dycOff = 0,
        dycFront = 1,
        dycRear = 2,
        dyc4WD = 3,
        dycFrontTracking = 4,
        dycRearTracking = 5
    }

    [DllImport("kernel32", EntryPoint = "GetCurrentProcessId")]
    public static extern int GetCurrentProcessId();
    [DllImport("kernel32", EntryPoint = "GetCurrentProcess")]
    public static extern int GetCurrentProcess();
    [DllImport("kernel32", EntryPoint = "TerminateProcess")]
    public static extern int TerminateProcess(int hProcess, int uExitCode);
    [DllImport("kernel32", EntryPoint = "RtlMoveMemory")]
    public static extern void CopyMemory(IntPtr pDest, IntPtr pSource, int ByteLen);
    public enum AE
    {
        InvalidActiveXLaunch = 101,
        CantCreateConfigurations = 102,
        InitializeFailedParsMissing = 103,
        CantLoadConstants = 104,
        DataPacketError = 105,
        CommunicationsFaulted = 158,
        DynoAsyncError = 161
    }

    public enum Opcodes
    {
        ocAddCarriageVeh = 1,
        ocClearAxleSim = 2,
        ocClearDynoLimits = 3,
        ocClearRobotBuffer = 4,
        ocClearTireSim = 5,
        ocEStopDyno = 6,
        ocForceSplitViaVehicle = 7,
        ocGetRobotDriveCycles = 8,
        ocGetRobotShiftTables = 9,
        ocGetRobotVehicleTypes = 10,
        ocInitialize = 11,
        ocLimitForce = 12,
        ocLimitSpeed = 13,
        ocMoveCarriage = 14,
        ocRefreshDynoLists = 15,
        ocRefreshRobotLists = 16,
        ocRefreshVehID = 17,
        ocRemoveCarriageVeh = 18,
        ocReqRbtData = 19,
        ocReqRbtErrors = 20,
        ocReqRbtFilenames = 21,
        ocReqTranType = 22,
        ocResetCarriage = 23,
        ocResetCommunications = 24,
        ocResetDy = 25,
        ocResetRobot = 26,
        ocSelectRobotDriveCycle = 27,
        ocSelectRobotShiftTable = 28,
        ocSelectRobotVehicleType = 29,
        ocSendDynoCommand = 30,
        ocSetAxleSim = 31,
        ocSetCarriageVeh = 32,
        ocSetDynoConfig = 33,
        ocSetDynoLimits = 34,
        ocSetDynoMode = 35,
        ocSetDynoOption = 36,
        ocSetDynoSync = 37,
        ocSetDyThrMode = 38,
        ocSetForce = 39,
        ocSetForceZero = 40,
        ocSetGrad = 41,
        ocSetK0 = 42,
        ocSetK1 = 43,
        ocSetK2 = 44,
        ocSetK3 = 45,
        ocSetPosition = 46,
        ocSetPositionZero = 47,
        ocSetRbtOnline = 48,
        ocSetRbtRemote = 49,
        ocSetRbtStandby = 50,
        ocSetRLS = 51,
        ocSetRobotMode = 52,
        ocSetSpeed = 53,
        ocSetSpeedZero = 54,
        ocSetThrMode = 55,
        ocSetTireSim = 56,
        ocSetVehicleGear = 57,
        ocSetVehicleID = 58,
        ocShowRobotVehicleFiles = 59,
        ocShowStatus = 60,
        ocShutdown = 61,
        ocStartCalib = 62,
        ocStartLogging = 63,
        ocStartRobotDriveCycle = 64,
        ocStartRobotLearnCycle = 65,
        ocStartVehicleEngine = 66,
        ocStartWarmup = 67,
        ocStopAux = 68,
        ocStopDy = 69,
        ocStopLogging = 70,
        ocStopRobot = 71,
        ocStopRobotDriveCycle = 72,
        ocStopVehicleEngine = 73,
        ocSuspend = 74,
        ocSyncSpeed = 75,
        ocTurnKeyOn = 76,
        ocZeroDist = 77,
        ocResetRobotComm = 78,
        ocSetActuator = 79,
        ocStartRbtLogging = 80,
        ocStopRbtLogging = 81,
        ocSendRobotCommand = 82,
        ocUseThrottlePot = 83
    }

    public enum DynoModes
    {
        dmStopped = 0,
        dmSpeed = 1,
        dmForce = 2,
        dmRLS = 3,
        dmWarmup = 4,
        dmCalib = 5
    }

    public enum ThrottleModes
    {
        tmStopped = 0,
        tmSpeed = 1,
        tmForce = 2,
        tmPosition = 3
    }

    public enum RobotModes
    {
        rmStopped = 0,
        rmSpeed = 1,
        rmForce = 2,
        rmPosition = 3,
        rmRPM = 4,
        rmVacuum = 5
    }

    public enum BrakeModes
    {
        bmStopped = 0,
        bmForce = 2,
        bmPedalForce = 7
    }

    public enum DriveCycleStatus
    {
        DC_NOTLOADED = 0,
        DC_LOADED = 1,
        DC_RUNNING = 2,
        DC_SUSPENDED = 3
    }

    public enum TCPIPCompletionStatuses
    {
        comWaiting = 0,
        comACK = 1,
        comNACK = 2
    }

    public enum Actuators
    {
        actNone = -1,
        actDyno = 0,
        actThrottle = 1,
        actBrake = 3,
        actRobot = 4
    }

    public enum TransmissionTypes
    {
        ttAutoTran = 0,
        ttManualTran = 1
    }

    public enum ControlActuators
    {
        caThrottle = 0,
        caRobot = 1
    }

    public const object NUM_COMMANDS = 115;
    public const object INACTIVE_TIMER = -1;
    public const object NOFAULT = 1;
    public const object FAULT = 0;
    public const object UseDefaultValue = -10000000000;
    public const string ILCK_GENERAL = "General";
    public const object ILCK_GENERAL_ROBOT = "Robot";
    public const string ILCK_SPEED = "Speed";
    public const object ILCK_THROTTLE = "Throttle";
    public const string ILCK_FORCE = "Force";
    public const string ILCK_POS = "Pos";
    public const string ILCK_INIT = "1";
    public const string ILCK_SPEED_RECEIVE = "Speed2";
    public const string ILCK_SPEED_RECEIVE_DIFF = "Speed3";
    public const string ILCK_FORCE_RECEIVE = "Force2";
    public const string ILCK_FORCE_RECEIVE_DIFF = "Force3";
    public const object ILCK_POS_RECEIVE = "Pos2";
    public const object ILCK_COM_OK = "ComOK";
    public const object SPG_STATUSES = "Statuses";
    public const object STS_DYNO_INITIALIZED = "DYNO_INITIALIZED";
    public const object STS_DYNO_MODE_ACHIEVED = "DYNO_MODE_ACHIEVED";
    public const object STS_THROTTLE_MODE_ACHIEVED = "THROTTLE_MODE_ACHIEVED";
    public const object STS_FORCE_ACHIEVED = "FORCE_ACHIEVED";
    public const object STS_SPEED_ACHIEVED = "SPEED_ACHIEVED";
    public const object STS_THROTTLE_ACHIEVED = "THROTTLE_ACHIEVED";
    public const object STS_RBT_MODE_ACHIEVED = "ROBOT_MODE_ACHIEVED";
    public const object STS_CARR_POS_ACHIEVED = "CARR_POS_ACHIEVED";
    public enum DynoCommandCodes
    {
        dcResetDyno = 1,
        dcSetDynoConfig = 2,
        dcStartWarmnup = 3,
        dcStartCalibration = 4,
        dcSetDynoThrottleMode = 5,
        dcZeroDistance = 7,
        dcSetDynoLimit = 8,
        dcSetThrottleLimit = 9,
        dcSetSpeed = 10,
        dcSetForce = 11,
        dcSetPosition = 12,
        dcSetVehicle = 13,
        dcSetRLS = 16,
        dcSetGradient = 18,
        dcStopDynoSystem = 19,
        dcStopDynoFast = 20,
        dcStartDynoData = 21,
        dcReportFault = 22,
        dcSendWTData = 23,
        dcGetVehicleIds = 24,
        dcSetSpeedDiff = 50,
        dcSetSpeedSync = 51,
        dcSetForceRatio = 52,
        dcSetForceSplitVeh = 53,
        dcSetActiveMode = 54,
        dcSetZeroSpeedMode = 55,
        dcSetSpeedtoZero = 100,
        dcSetForcetoZero = 101,
        dcSetPositiontoZero = 102,
        dcFreeText = 103,
        dcStopDynoData = 104,
        dcMoveCarriage = 120,
        dcResetCarriage = 121,
        dcSetActuator = 122,
        dcSetThrPot = 123,
        rcClearBuff = 201,
        rcGetData = 202,
        rcGetDriveCycleList = 203,
        rcGetFaults = 204,
        rcGetFileNames = 205,
        rcGetShiftTables = 206,
        rcGetVehicleTypes = 207,
        rcGetTransType = 208,
        rcReset = 209,
        rcSelectDriveCycle = 210,
        rcSetMode = 211,
        rcSetOnlineDirect = 212,
        rcSetRemote = 213,
        rcSelectShiftTable = 214,
        rcSetStandby = 215,
        rcSelectVehicleType = 216,
        rcSetGear = 217,
        rcStartDriveCycle = 218,
        rcStartLearnCycle = 219,
        rcStartEngine = 220,
        rcStop = 221,
        rcStopDriveCycle = 222,
        rcStopEngine = 223,
        rcTurnKeyOn = 224,
        rcPowerOn = 225,
        rcPowerOff = 226,
        rcFreeText = 227
    }

    public const object ACK_STRING = "ACK";
    public const object NACK_STRING = "NACK";
    public const object DYCMD_SetSpeed = "SETSPEED";
    public const object DYCMD_SetForce = "SETFORCE";
    public const object DYCMD_SetPosition = "SETPOSITION";
    public const object RBTCMD_CLR_BUFF = "SLDR";
    public const object RBTCMD_GET_RBT_Data = "AWRX";
    public const object RBTCMD_GET_RBT_DC = "AFPT";
    public const object RBTCMD_GET_RBT_FAULTS = "ASTZ";
    public const object RBTCMD_GET_RBT_FILENAMES = "AADF";
    public const object RBTCMD_GET_RBT_ST = "AVST";
    public const object RBTCMD_GET_RBT_VT = "AFBT";
    public const object RBTCMD_GET_TRAN_TYPE = "AFGD";
    public const object RBTCMD_RBT_RESET = "SRES";
    public const object RBTCMD_SET_RBT_DC = "EFPW";
    public const object RBTCMD_SET_RBT_MODE = "SVRD";
    public const object RBTCMD_SET_RBT_OnDir = "SFOD";
    public const object RBTCMD_SET_RBT_Remote = "SREM";
    public const object RBTCMD_SET_RBT_ST = "ESTW";
    public const object RBTCMD_SET_RBT_STBY = "STBY";
    public const object RBTCMD_SET_RBT_VT = "EFBW";
    public const object RBTCMD_SET_VEH_GEAR = "SetVehicleGear(internal only)";
    public const object RBTCMD_START_RBT_DC = "SSTF";
    public const object RBTCMD_START_RBT_LC = "SLFB";
    public const object RBTCMD_START_VEH_ENG = "SSTM";
    public const object RBTCMD_STOP_RBT = "SRES";
    public const object RBTCMD_STOP_RBT_DC = "STBY";
    public const object RBTCMD_STOP_VEH_ENG = "SZAU";
    public const object RBTCMD_TURN_KEY_ON = "SZEI";
    public const object RBTCMD_TURNOFF_RBT = "SPAU";
    public const object RBTCMD_TURNON_RBT = "SFOD";
    public const object VEH_GEARSTRING_P = "P";
    public const object VEH_GEARSTRING_D = "D";
    public const object VEH_GEARSTRING_M = "M";
    public const object VEH_GEARSTRING_PLUS = "+";
    public const object VEH_GEARSTRING_MINUS = "-";
    public const object VEH_GEARSTRING_R = "R";
    public const object VEH_GEARSTRING_N = "N";
    public const object VEH_GEARSTRING_1 = "1";
    public const object VEH_GEARSTRING_2 = "2";
    public const object VEH_GEARSTRING_3 = "3";
    public const object VEH_GEARSTRING_4 = "4";
    public const object VEH_GEARSTRING_5 = "5";
    public const object VEH_GEARSTRING_6 = "6";
    public const object VEH_GEARSTRING_AUTO = "Auto";
    public const object AUTO_SELECT = 1;
    public const object USER_SELECT = 0;
    public const object AUTO_GEAR_MINUS = 10;
    public const object AUTO_GEAR_PLUS = 9;
    public const object AUTO_GEAR_M = 8;
    public const object AUTO_GEAR_P = 7;
    public const object AUTO_GEAR_R = 6;
    public const object AUTO_GEAR_D = 5;
    public const object AUTO_GEAR_4 = 4;
    public const object AUTO_GEAR_3 = 3;
    public const object AUTO_GEAR_2 = 2;
    public const object AUTO_GEAR_1 = 1;
    public const object AUTO_GEAR_N = 0;
    public const object MAN_GEAR_N = 0;
    public const object MAN_GEAR_1 = 1;
    public const object MAN_GEAR_2 = 2;
    public const object MAN_GEAR_3 = 3;
    public const object MAN_GEAR_4 = 4;
    public const object MAN_GEAR_5 = 5;
    public const object MAN_GEAR_6 = 6;
    public const object MAN_GEAR_R = 7;
    public const object MAN_GEAR_OFFSET = 10;
    public const object MANUAL_TRAN = 1;
    public const object AUTO_TRAN = 0;
    public const object V3_AUTO_SELECT = 1;
    public const object RBT_ONLINE_DIRECT = 1;
    public const object RBT_OFFLINE = 0;
    public const object RBT_OPEN_LOOP = 2;
    public const object RBT_ONLINE = 3;
    public const object RBT_STS_MSG_SEPARATOR = "$SI$";
    public const object DC_STOPPED = 0;
    public const object RBT_SLAVE_ID = "K0";
    public const object RBT_FRC_MODE_CODE = 3;
    public const object RBT_SPD_MODE_CODE = 0;
    public const object RBT_POS_MODE_CODE = 1;
    public const object RBT_POS_CLUTCH_MODE_CODE = 7;
    public const object OFFLINE_MODE_TEXT = "Offline";
    public const object SPEED_MODE_TEXT = "Speed";
    public const object FORCE_MODE_TEXT = "Force";
    public const object RLS_MODE_TEXT = "RLS";
    public const object POSITION_MODE_TEXT = "Position";
    public const object PROCESS_NAME = "Dynamometer";
    public const object PROCESS_DESCRIPTION = "Dyno Controller";
    public const object REG_PROCESSES = "Software\\ReACT Technologies\\Talent\\4.0\\Processes";
    public const object REG_DYNO = "Software\\ReACT Technologies\\Talent\\4.0\\Dyno";
    public const object REG_SHARED_FOLDERS = "SOFTWARE\\ReACT Technologies\\Talent\\4.0\\Shared Folders";
    public static string ActuatorIDAsText(ref ControlActuators ActuatorID)
    {
        string __result = "";
        switch (ActuatorID)
        {
            case caThrottle:
                __result = "Throttle";
                break;
            case caRobot:
                __result = "Robot";
                break;
            default:
                __result = string.Concat(string.Concat("Unknown (", ActuatorID), ")");
                break;
        }

        return __result;
    }

    public static ControlActuators ActuatorIDFromText(ref string ActuatorIDText)
    {
        ControlActuators __result = 0;
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        for (I = caThrottle; I <= caRobot; I++)
        {
            if (Regex.IsMatch(ActuatorIDAsText(I), string.Concat(ActuatorIDText, "*")))
            {
                __result = I;
                goto Done;
            }
        }

        sRaiseErrorText = LoadResString(ridBadActuator, ActuatorIDText);
        bRaiseError = true;
        lRaiseErrorNumber = tceInvalidArgs;
        goto Done;
        Error();
        sRaiseErrorText = LoadResString(ridBadActuator, ex.Message);
        bRaiseError = true;
        lRaiseErrorNumber = tceInvalidArgs;
        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        if (bRaiseError)
        {
            EH_RaiseError(lRaiseErrorNumber, "modDYControl", PROC_NAME, sRaiseErrorText);
        }

        return __result;
    }

    public static object ArgValue(ref object Arg, VbVarType TargetVariableType = vbVariant)
    {
        object __result = null;
        object vValue = null;
        if (Arg is CommandListLib.Argument)
        {
            vValue = Arg.Value;
        }
        else
        {
            vValue = Arg;
        }

        if (TargetVariableType == vbVariant)
        {
            __result = vValue;
        }
        else
        {
            __result = Convert(vValue, TargetVariableType);
        }

        return __result;
    }

    public static string ArgValueAsString(ref object Expression)
    {
        string __result = "";
        try
        {
            switch (VarType(Expression))
            {
                case vbString:
                    __result = string.Concat(string.Concat("\"", Expression.ToString()), "\"");
                    break;
                default:
                    if (Expression == UseDefaultValue)
                    {
                        __result = "";
                    }
                    else
                    {
                        __result = CStrUS(Expression);
                    }

                    break;
            }

            return __result;
        }
        catch
        {
            // On Error Resume Next: exceptions suppressed
        }

        return __result;
    }

    public static double ASin(ref double Expression)
    {
        double __result = 0;
        if (Expression > 0.99999999)
        {
            __result = 3.14159 / 2;
        }
        else if (Expression < -0.99999999)
        {
            __result = -3.14159 / 2;
        }
        else
        {
            __result = Math.Atan(Expression / Math.Sqrt(-Expression * Expression + 1));
        }

        return __result;
    }

    public static float CalcRampTime(float Target, float Actual, float RampRate)
    {
        float __result = 0;
        try
        {
            __result = 1;
            __result = Math.Abs(Target - Actual) / RampRate;
            return __result;
        }
        catch
        {
            // On Error Resume Next: exceptions suppressed
        }

        return __result;
    }

    public static void CheckModeCombination(ref DynoModes DyMode, ref ThrottleModes ThrMode)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        switch (DyMode)
        {
            case dmSpeed:
                if (ThrMode == tmSpeed)
                {
                    throw new Exception(5);
                }

                break;
            case dmRLS:
            case dmForce:
                if (ThrMode == tmForce)
                {
                    throw new Exception(5);
                }

                break;
        }

        switch (ThrMode)
        {
            case tmPosition:
                switch (DyMode)
                {
                    case dmSpeed:
                    case dmForce:
                    case dmRLS:
                        break;
                    default:
                        throw new Exception(5);
                        break;
                }

                break;
        }

        goto Done;
        Error();
        sRaiseErrorText = LoadResString(ridBadDynoThrMode, DynoModeAsText(DyMode), ThrModeAsText(ThrMode));
        bRaiseError = true;
        lRaiseErrorNumber = tceInvalidArgs;
        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        if (bRaiseError)
        {
            EH_RaiseError(lRaiseErrorNumber, MOD_NAME, PROC_NAME, sRaiseErrorText);
        }

        return;
    }

    public static string DyConfigAsText(ref object DyConfig)
    {
        string __result = "";
        switch (DyConfig)
        {
            case dycFront:
                __result = "FWD";
                break;
            case dycFrontTracking:
                __result = "FWDT";
                break;
            case dycRear:
                __result = "RWD";
                break;
            case dycRearTracking:
                __result = "RWDT";
                break;
            case dyc4WD:
                __result = "4WD";
                break;
            case dycOff:
                __result = "Off";
                break;
            default:
                __result = string.Concat(string.Concat("Unknown (", DyConfig), ")");
                break;
        }

        return __result;
    }

    public static DynoConfigs DyConfigFromText(ref string DynoConfigText, bool OffOK = false)
    {
        DynoConfigs __result = 0;
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        for (I = IIf(OffOK, dycOff, dycFront); I <= dycRearTracking; I++)
        {
            if (Regex.IsMatch(DyConfigAsText(I), string.Concat(DynoConfigText, "*")))
            {
                __result = I;
                goto Done;
            }
        }

        sRaiseErrorText = LoadResString(ridBadDyConfig, DynoConfigText);
        bRaiseError = true;
        lRaiseErrorNumber = tceInvalidArgs;
        goto Done;
        Error();
        sRaiseErrorText = LoadResString(ridBadDyConfig, ex.Message);
        bRaiseError = true;
        lRaiseErrorNumber = tceInvalidArgs;
        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        if (bRaiseError)
        {
            EH_RaiseError(lRaiseErrorNumber, MOD_NAME, PROC_NAME, sRaiseErrorText);
        }

        return __result;
    }

    public static string DynoModeAsText(ref object DynoMode)
    {
        string __result = "";
        switch (DynoMode)
        {
            case dmStopped:
                __result = "Offline";
                break;
            case dmSpeed:
                __result = "Speed";
                break;
            case dmForce:
                __result = "Force";
                break;
            case dmRLS:
                __result = "RLS";
                break;
            case dmWarmup:
                __result = "Warmup";
                break;
            case dmCalib:
                __result = "Calib";
                break;
            default:
                __result = string.Concat(string.Concat("Unknown (", DynoMode), ")");
                break;
        }

        return __result;
    }

    public static DynoModes DynoModeFromText(ref string DynoModeText)
    {
        DynoModes __result = 0;
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        for (I = dmStopped; I <= dmCalib; I++)
        {
            if (Regex.IsMatch(DynoModeAsText(I), string.Concat(DynoModeText, "*")))
            {
                __result = I;
                goto Done;
            }
        }

        sRaiseErrorText = LoadResString(ridBadDynoMode, DynoModeText);
        bRaiseError = true;
        lRaiseErrorNumber = tceInvalidArgs;
        goto Done;
        Error();
        sRaiseErrorText = LoadResString(ridBadDynoMode, ex.Message);
        bRaiseError = true;
        lRaiseErrorNumber = tceInvalidArgs;
        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        if (bRaiseError)
        {
            EH_RaiseError(lRaiseErrorNumber, MOD_NAME, PROC_NAME, sRaiseErrorText);
        }

        return __result;
    }

    public static string GetCmdArgs(ref CommandListLib.Arguments Arguments)
    {
        string __result = "";
        short I = 0;
        try
        {
            for (I = 1; I <= Arguments.Count; I++)
            {
                if (GetCmdArgs != "")
                {
                    __result = string.Concat(GetCmdArgs, ", ");
                }

                __result = string.Concat(GetCmdArgs, ArgValueAsString(Arguments(I).Value));
            }

            if (GetCmdArgs != "")
            {
                __result = string.Concat(" ", GetCmdArgs);
            }

        Done:
            ;
            return __result;
        }
        catch
        {
            // On Error Resume Next: exceptions suppressed
        }

        return __result;
    }

    public static object IsSameCommand(ref Command LastCommand, ref string ReceivedMemnomic)
    {
        object __result = null;
        // TODO: On Error GoTo Error — handler label not found in this scope
        if (LastCommand.TagVariant.LastCommand.CommandText.Trim() == ReceivedMemnomic.Trim())
        {
            __result = true;
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
        else
        {
            __result = VB.LoadResString(ID);
            __result = ReplaceItemText(LoadResString, ItemTextArgs);
        }

        return __result;
    }

    public static void LogCommMessage(string MessageText, ref object DataPacket, TalentProcessLib.MessageTypes MessageType = mtInformation)
    {
        string sText = "";
        int I = 0;
        try
        {
            sText = MessageText;
            if (IsMissing(DataPacket))
            {
            }
            else
            {
                if (sText.Substring(sText.Length - 1) != " ")
                {
                    sText = string.Concat(sText, " ");
                }

                if (VarType(DataPacket) == vbString)
                {
                    sText = string.Concat(sText, "\"");
                    for (I = 1; I <= DataPacket.Length; I++)
                    {
                        switch ((short)Microsoft.VisualBasic.Strings.Mid(DataPacket, I, 1).FirstOrDefault())
                        {
                            case 2:
                                sText = string.Concat(sText, "{STX}");
                                break;
                            case 3:
                                sText = string.Concat(sText, "{ETX}");
                                break;
                            case 10:
                                sText = string.Concat(sText, "{LF}");
                                break;
                            case 13:
                                sText = string.Concat(sText, "{CR}");
                                break;
                            default:
                                sText = string.Concat(sText, Microsoft.VisualBasic.Strings.Mid(DataPacket, I, 1));
                                break;
                        }

                        if (I >= 200)
                        {
                            sText = string.Concat(string.Concat(string.Concat(sText, ".... "), DataPacket.Length - 200), " more characters");
                            break;
                        }
                    }

                    sText = string.Concat(sText, "\"");
                }
                else
                {
                    sText = string.Concat(string.Concat(string.Concat(sText, "\"<"), LenB(DataPacket)), " bytes binary data>\"");
                }
            }

            LogMessage(sText, MessageType);
            return;
        }
        catch
        {
            // On Error Resume Next: exceptions suppressed
        }
    }

    public static void Main()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        App.StartLogging("", vbLogAuto);
        App.OleServerBusyRaiseError = true;
        App.OleServerBusyTimeout = 1000;
        Process.Title = PROCESS_NAME;
        Process.EHCaption = PROCESS_NAME;
        Process.EHMode = ehLogOnly;
        Process.EHMsgsToLog = ehCriticalExclamation;
        Process.LogMode = logToNT;
        goto Done;
        Error();
        if (Err != 364)
        {
            EH.Module = MOD_NAME;
            EH.Procedure = PROC_NAME;
            EH_MsgBox();
        }

        goto ErrorExit;
    ErrorExit:
        ;
        EH.AppError = ERR_CANT_START;
        EH.Module = MOD_NAME;
        EH.Procedure = PROC_NAME;
        EH_MsgBox();
        // On Error Resume Next — nested handler; see enclosing try/catch
        goto Done;
    Done:
        ;
        return;
    }

    public static bool RbtModesCompatible(ref int DynoMode, ref int ThrottleMode, ref int BrakeMode, ref int RobotMode)
    {
        bool __result = false;
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        bModesOK = true;
        switch (DynoMode)
        {
            case dmStopped:
                if (ThrottleMode != tmStopped | BrakeMode != bmStopped | RobotMode != rmStopped)
                {
                    bModesOK = false;
                }

                break;
            case dmSpeed:
                switch (ThrottleMode)
                {
                    case tmForce:
                    case tmPosition:
                        if (BrakeMode != bmStopped | RobotMode != rmStopped)
                        {
                            bModesOK = false;
                        }

                        break;
                    case tmStopped:
                        switch (BrakeMode)
                        {
                            case bmForce:
                            case bmPedalForce:
                                if (RobotMode != rmStopped)
                                {
                                    bModesOK = false;
                                }

                                break;
                            case bmStopped:
                                switch (RobotMode)
                                {
                                    case rmForce:
                                    case rmPosition:
                                    case rmStopped:
                                        break;
                                    default:
                                        bModesOK = false;
                                        break;
                                }

                                break;
                            default:
                                bModesOK = false;
                                break;
                        }

                        break;
                    default:
                        bModesOK = false;
                        break;
                }

                break;
            case dmForce:
                switch (ThrottleMode)
                {
                    case tmSpeed:
                    case tmPosition:
                        if (BrakeMode != bmStopped | RobotMode != rmStopped)
                        {
                            bModesOK = false;
                        }

                        break;
                    case tmStopped:
                        switch (BrakeMode)
                        {
                            case bmStopped:
                                switch (RobotMode)
                                {
                                    case rmSpeed:
                                    case rmPosition:
                                    case rmStopped:
                                        break;
                                    default:
                                        bModesOK = false;
                                        break;
                                }

                                break;
                            default:
                                bModesOK = false;
                                break;
                        }

                        break;
                    default:
                        bModesOK = false;
                        break;
                }

                break;
            case dmRLS:
                switch (ThrottleMode)
                {
                    case tmSpeed:
                    case tmPosition:
                        if (BrakeMode != bmStopped | RobotMode != rmStopped)
                        {
                            bModesOK = false;
                        }

                        break;
                    case tmStopped:
                        switch (BrakeMode)
                        {
                            case bmStopped:
                                switch (RobotMode)
                                {
                                    case rmSpeed:
                                    case rmPosition:
                                    case rmStopped:
                                        break;
                                    default:
                                        bModesOK = false;
                                        break;
                                }

                                break;
                            default:
                                bModesOK = false;
                                break;
                        }

                        break;
                    default:
                        bModesOK = false;
                        break;
                }

                break;
            default:
                bModesOK = false;
                break;
        }

        __result = bModesOK;
        goto Done;
        Error();
        bModesOK = false;
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

    public static string ThrModeAsText(ref object ThrMode)
    {
        string __result = "";
        switch (ThrMode)
        {
            case tmStopped:
                __result = "Offline";
                break;
            case tmSpeed:
                __result = "Speed";
                break;
            case tmForce:
                __result = "Force";
                break;
            case tmPosition:
                __result = "Position";
                break;
            default:
                __result = string.Concat(string.Concat("Unknown (", ThrMode), ")");
                break;
        }

        return __result;
    }

    public static ThrottleModes ThrModeFromText(ref string ThrottleModeText)
    {
        ThrottleModes __result = 0;
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        for (I = tmStopped; I <= tmPosition; I++)
        {
            if (Regex.IsMatch(ThrModeAsText(I), string.Concat(ThrottleModeText, "*")))
            {
                __result = I;
                goto Done;
            }
        }

        sRaiseErrorText = LoadResString(ridBadThrMode, ThrottleModeText);
        bRaiseError = true;
        lRaiseErrorNumber = tceInvalidArgs;
        goto Done;
        Error();
        sRaiseErrorText = LoadResString(ridBadThrMode, ex.Message);
        bRaiseError = true;
        lRaiseErrorNumber = tceInvalidArgs;
        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        if (bRaiseError)
        {
            EH_RaiseError(lRaiseErrorNumber, MOD_NAME, PROC_NAME, sRaiseErrorText);
        }

        return __result;
    }
}