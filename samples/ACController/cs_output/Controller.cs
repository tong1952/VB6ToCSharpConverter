// Converted from Controller.cls by vb6cs
// Date: 2026-05-11 03:41

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
    private ParManager mobjParManager;
    private LoopTimer tmrMain;
    private ccrpTimer tmrRunOnce;
    private ccrpStopWatch tmrElapsed;
    private Interlocks mobjInterlocks;
    private TalentProcessLib.Events mobjEvents;
    private frmControl mobjForm;
    private SharedPropertyGroup mobjSPGConstants;
    private SharedPropertyGroup mobjSPGDataItems;
    private SocketPoll mobjGFAI;
    private TalentProcessLib.LogFile mobjLogFile;
    private MailSlot mobjGFAIMessages;
    private LogFile mobjGFAIMessageLogFile;
    private int lID;
    private bool bReady;
    private bool bInited;
    private const bool bKeepStats = false;
    private bool bRealTime;
    private DateTime mgvdLastCommandTime;
    private object mvInterlockCheckValue;
    private DateTime mdStartTime;
    public Collection mcolInvalidOpcodes;
    private enum InitTypes
    {
        initAll = 0,
        initConstants = 1,
        initPars = 2
    }

    private string mccsPriority;
    private float mccfUpdateTime;
    private float mccfGFAITCPIPConnectTimeout;
    private string mccsGFAITCPIPAddress;
    private int mcclGFAITCPIPPort;
    private const bool mccbGFAITCPIPAnsi = true;
    private float mccfKeepAlivePacketInterval;
    private float mccfMaxTimeBetweenSendPars;
    private bool mccbUseTargetValues;
    private int mcclGetControlTimeout;
    private bool mccbSendParsOnlyIfUnderWTCS;
    private float mipAirtemp;
    private float mipWindspeed;
    private float mipWsMode;
    private float mipWsAchieved;
    private float mipWsTarget;
    private float mipRelHumidity;
    private float mipYaw;
    private float mipYawAchieved;
    private float mipYawTarget;
    private float mipAbsPress;
    private float mopGFAITCPIPStatus;
    private float mopGFAIState;
    private float mopWTCSInControl;
    private bool mgvbLogOutput;
    private bool mgvbLogInput;
    private bool mgvbLogToMsgLog;
    private int mgvlNumReceiveErrors;
    private int mlDefPriority;
    private string mgvsLastSentCommand;
    private bool mgvbSimulation;
    private bool mgvbWsWasAchieved;
    private float mgvfWsNowAchievedTgt;
    private bool mgvbYawWasAchieved;
    private bool mgvfYawNowAchievedTgt;
    private DateTime mgvdLastGFAIResponseTime;
    private string mgvsLastGFAIResponse;
    private DateTime mgvdGFAIDataLastRequestTime;
    private DateTime mgvdLastDataStreamLoggedTime;
    private DateTime mgvdLastSendParsTime;
    private string mgvsLastGFAIResponseErrMsg;
    private int mgvlLastGFAIResponseErrNum;
    private string mgvsLastGFAIResponseErrSource;
    private string mgvsLastGFAIResponseErrInput;
    private DateTime mgvdLastGFAIResponseErrTime;
    private DateTime mgvdLastKeepAlivePacketTime;
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

    public Events Events
    {
        get
        {
            Events __result = 0;
            __result = mobjEvents;
            return __result;
        }
    }

    private void ProcessGFAIResponse(string GFAIResponse)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        GFAIResponse = GFAIResponse.Trim();
        if (GFAIResponse.Substring(0, 1) == (char)2)
        {
            if (GFAIResponse.Length < 3)
            {
                goto Done;
            }

            GFAIResponse = GFAIResponse.Substring(2 - 1);
        }

        if (GFAIResponse.Substring(GFAIResponse.Length - 1) == (char)3)
        {
            GFAIResponse = GFAIResponse.Substring(0, GFAIResponse.Length - 1);
        }

        mgvsLastGFAIResponse = GFAIResponse;
        if (mgvbLogInput)
        {
            sLogText = GFAIResponse.Replace(vbLf, "");
            sLogText = sLogText.Replace(vbCr, "");
            mobjGFAIMessageLogFile.LogMessage(sLogText, "Received:", mtInformation);
            if (mgvbLogToMsgLog)
            {
                LogMessage(string.Concat("Received: ", sLogText), mtInformation);
            }
        }

        if (Regex.IsMatch(GFAIResponse, "*StillAliveReply*"))
        {
        }
        else if (Regex.IsMatch(GFAIResponse, "*ErrorEvent*"))
        {
            lErrDescStart = GFAIResponse.IndexOf("\"errorDescription\": ", StringComparison.Ordinal) + 1 + 21;
            for (I = lErrDescStart; I <= GFAIResponse.Length; I++)
            {
                if (Microsoft.VisualBasic.Strings.Mid(GFAIResponse, I, 1) == "\"")
                {
                    break;
                }

                sFaultText = string.Concat(sFaultText, Microsoft.VisualBasic.Strings.Mid(GFAIResponse, I, 1));
            }

            if (Regex.IsMatch(GFAIResponse, "*\"Critical\"*"))
            {
                lMessageType = msgCritical;
            }
            else if (Regex.IsMatch(GFAIResponse, "*\"Error\"*"))
            {
                lMessageType = msgCritical;
            }
            else if (Regex.IsMatch(GFAIResponse, "*\"Warning\"*"))
            {
                lMessageType = msgExclamation;
            }
            else
            {
                lMessageType = msgInformation;
            }

            LogMessage(sFaultText, lMessageType);
        }
        else if (Regex.IsMatch(GFAIResponse, "*ShutdownEvent*"))
        {
            mobjGFAI.Shutdown();
            mopGFAITCPIPStatus = 0;
            mopWTCSInControl = 0;
        }
        else if (Regex.IsMatch(GFAIResponse, "*GetAMSStatusReply*") | Regex.IsMatch(GFAIResponse, "*AMSStatusChangedEvent*"))
        {
            if (Regex.IsMatch(GFAIResponse, "*Uninitialized*"))
            {
                mopGFAIState = gsUninitialized;
            }
            else if (Regex.IsMatch(GFAIResponse, "*Preview*"))
            {
                mopGFAIState = gsPreview;
            }
            else if (Regex.IsMatch(GFAIResponse, "*PreTrigger*"))
            {
                mopGFAIState = gsPreTrigger;
            }
            else if (Regex.IsMatch(GFAIResponse, "*TriggerReady*"))
            {
                mopGFAIState = gsTriggerReady;
            }
            else if (Regex.IsMatch(GFAIResponse, "*Recording*"))
            {
                mopGFAIState = gsRecording;
            }
            else if (Regex.IsMatch(GFAIResponse, "*DataTransfer*"))
            {
                mopGFAIState = gsDataTransfer;
            }
            else if (Regex.IsMatch(GFAIResponse, "*CancelRecord*"))
            {
                mopGFAIState = gsCancelRecord;
            }
            else if (Regex.IsMatch(GFAIResponse, "*PostProcessing*"))
            {
                mopGFAIState = gsPostProcessing;
            }
            else if (Regex.IsMatch(GFAIResponse, "*MeasurementFinished*"))
            {
                mopGFAIState = gsMeasurementFinished;
            }
            else
            {
                mopGFAIState = -1;
            }
        }
        else if (Regex.IsMatch(GFAIResponse, "*ReleaseControlRequest*"))
        {
            SendCommand("ReleaseControlReply", Array("requestControlReply", "RequestGranted"));
            mopWTCSInControl = 0;
        }
        else if (Regex.IsMatch(GFAIResponse, "*GetControlReply*"))
        {
            if (Regex.IsMatch(GFAIResponse, "*RequestGranted*"))
            {
                mopWTCSInControl = 1;
            }
            else
            {
                mopWTCSInControl = 0;
                if (CurrentCommandOpcode == ocTakeControl)
                {
                    AbortCurrentCommand("Denied by GFAI system");
                }
            }
        }
        else if (Regex.IsMatch(GFAIResponse, "*GetCurrentControlOwnerReply*"))
        {
            if (Regex.IsMatch(GFAIResponse, "*You*"))
            {
                mopWTCSInControl = 1;
            }
            else
            {
                mopWTCSInControl = 0;
            }
        }
        else if (Regex.IsMatch(GFAIResponse, "*ForceReleaseControlEvent*"))
        {
            mopWTCSInControl = 0;
        }
        else if (Regex.IsMatch(GFAIResponse, "*UnsolicitedGetControlCommand*"))
        {
            mopWTCSInControl = 1;
        }

        goto Done;
        Error();
        mgvlNumReceiveErrors = mgvlNumReceiveErrors + 1;
        mgvsLastGFAIResponseErrMsg = ex.Message;
        mgvlLastGFAIResponseErrNum = ex.HResult;
        mgvsLastGFAIResponseErrSource = ex.Message;
        mgvsLastGFAIResponseErrInput = GFAIResponse;
        if (ex.HResult & vbObjectError != vbObjectError)
        {
            EH.VBErr = Err;
            if (SecondsSinceTime(vLastErrorMessageTime) >= 300)
            {
                EH.AppError = LoadResString(AE.GFAIResponseError);
                EH.Item(1) = GFAIResponse;
                EH.Module = TypeName(this);
                EH.Procedure = PROC_NAME;
                LogMessage2(EH);
                vLastErrorMessageTime = Now;
            }

            EH.AppError = null;
        }

        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        return;
    }

    private int CurrentCommandOpcode()
    {
        int __result = 0;
        if (mobjCommands.Count == 0)
        {
            __result = 0;
        }
        else
        {
            __result = mobjCommands(1).Opcode;
        }

        return __result;
    }

    private void AbortCurrentCommand(string CompletionMessage = "")
    {
        if (mobjCommands.Count > 0)
        {
            mobjCommands(1).Abort(CompletionMessage);
        }
    }

    private void Suspend()
    {
        Attribute(Suspend.VB_Description == "Suspend control");
        try
        {
            bInited = false;
            mobjParManager.ReleaseOutputs();
            mobjStatuses(STS_GFAI_INITIALIZED).SetValue(bInited);
            mobjGFAI.Disconnect();
            return;
        }
        catch
        {
            // On Error Resume Next: exceptions suppressed
        }
    }

    private void GFAIControlLogic()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        // On Error Resume Next — nested handler; see enclosing try/catch
        if (mgvbSimulation)
        {
            mopGFAITCPIPStatus = NOFAULT;
        }
        else if (mobjGFAI.State == socConnected)
        {
            mopGFAITCPIPStatus = NOFAULT;
        }
        else
        {
            mopGFAITCPIPStatus = FAULT;
        }

        if (Err)
        {
            mopGFAITCPIPStatus = FAULT;
        }

        if (mgvbSimulation == false)
        {
            if (!mobjGFAI == null)
            {
                if (mobjGFAI.State == soxConnected)
                {
                    if (mccbSendParsOnlyIfUnderWTCS == false | mopWTCSInControl == 1)
                    {
                        if (WindspeedNowAchieved | YawAngleNowAchieved | mccfMaxTimeBetweenSendPars > 0 & SecondsSinceTime(mgvdLastSendParsTime) >= mccfMaxTimeBetweenSendPars)
                        {
                            // On Error Resume Next — nested handler; see enclosing try/catch
                            if (mopGFAIState <= gsPreview)
                            {
                                SendParValues();
                            }// On Error GoTo Error — nested handler not restructured
                        }
                    }
                }
            }
        }

        if (mgvbSimulation == false)
        {
            if (!mobjGFAI == null)
            {
                if (mobjGFAI.State == soxConnected)
                {
                    if (mccfKeepAlivePacketInterval > 0)
                    {
                        if (SecondsSinceTime(mgvdLastKeepAlivePacketTime) >= mccfKeepAlivePacketInterval)
                        {
                            // On Error Resume Next — nested handler; see enclosing try/catch
                            SendCommand("StillAliveRequest");
                            mgvdLastKeepAlivePacketTime = Now;
                            // On Error GoTo Error — nested handler not restructured
                        }
                    }
                }
            }
        }

        if (mopGFAITCPIPStatus == 0)
        {
            mopGFAIState = 0;
            mopWTCSInControl = 0;
        }

        if (mgvbSimulation)
        {
        }

        goto Done;
        Error();
        if (bInited)
        {
            EH.Module = TypeName(this);
            EH.Procedure = PROC_NAME;
            EH_MsgBox();
            Suspend();
        }

        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        return;
    }

    private bool WindspeedNowAchieved
    {
        get
        {
            bool __result = false;
            // TODO: On Error GoTo Error — handler label not found in this scope
            bAchieved = false;
            if (mipWsMode == 2)
            {
                if (mipWsAchieved == 1)
                {
                    bAchieved = true;
                }
            }

            if (bAchieved)
            {
                if (mgvbWsWasAchieved == false)
                {
                    __result = true;
                    mgvfWsNowAchievedTgt = mipWsTarget;
                }
                else if (mipWsTarget != mgvfWsNowAchievedTgt)
                {
                    __result = true;
                    mgvfWsNowAchievedTgt = mipWsTarget;
                }
            }

            mgvbWsWasAchieved = bAchieved;
            goto Done;
            Error();
            __result = false;
            goto Done;
        Done:
            ;
            return __result;
        }
    }

    private bool YawAngleNowAchieved
    {
        get
        {
            bool __result = false;
            // TODO: On Error GoTo Error — handler label not found in this scope
            if (mipYawAchieved == 1)
            {
                bAchieved = true;
            }

            if (bAchieved)
            {
                if (mgvbYawWasAchieved == false)
                {
                    __result = true;
                    mgvfYawNowAchievedTgt = mipYawTarget;
                }
                else if (mipYawTarget != mgvfYawNowAchievedTgt)
                {
                    __result = true;
                    mgvfYawNowAchievedTgt = mipYawTarget;
                }
            }

            mgvbYawWasAchieved = bAchieved;
            goto Done;
            Error();
            __result = false;
            goto Done;
        Done:
            ;
            return __result;
        }
    }

    public bool Initialized
    {
        get
        {
            bool __result = false;
            // TODO: On Error GoTo Error — handler label not found in this scope
            bRaiseError = false;
            __result = bInited;
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
                EH_RaiseError(OLEERR_PROPERTY_GET_FAILED, TypeName(this), PROC_NAME, sRaiseErrorText);
            }

            return __result;
        }
    }

    public void CheckInterlock(ref object Opcode, object CheckValue = null)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        // On Error Resume Next — nested handler; see enclosing try/catch
        sIndex = Opcode;
        objInterlock = mobjInterlocks(sIndex);
        if (Err)
        {
            goto Done;
        }

        mvInterlockCheckValue = CheckValue;
        // On Error GoTo Error — nested handler not restructured
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
            throw new Exception(sRaiseErrorText);
        }

        return;
    }

    public Interlocks Interlocks
    {
        get
        {
            Interlocks __result = 0;
            __result = mobjInterlocks;
            return __result;
        }
    }

    public object GetInvalidOpcodes(ref object Opcode)
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
        sClassName = "Events";
        mobjEvents = new Events();
        sClassName = "Interlocks";
        mobjInterlocks = new Interlocks();
        sClassName = "Invalid Opcodes Collection";
        mcolInvalidOpcodes = new Collection();
        sClassName = "frmControl";
        mobjForm = new frmControl();
        sClassName = "SharedPropertyGroupManager";
        objSPGM = new SharedPropertyGroupManager();
        sClassName = "SharedPropertyGroup - Constants";
        mobjSPGConstants = objSPGM.CreatePropertyGroup(SPG_CONSTANTS, 0, 0, bExists);
        mobjSPGDataItems = objSPGM.CreatePropertyGroup(SPG_DATA, 0, 0, bExists);
        // On Error GoTo Error — nested handler not restructured
        tmrElapsed = new ccrpStopWatch();
        DoEvents();
        mobjInterlocks.Add(ocInitialize.ToString()).Expressions.AddBoolean("Initialized", bInited, true, ilckVariableEqualValueMeansOK, LoadResString(stsControlSystemNotInitialized));
        mobjInterlocks.Add(ILCK_GENERAL).IncludedInterlocks.Add(ocInitialize.ToString());
        mobjInterlocks.Add(ILCK_COM_OK).IncludedInterlocks.Add(ILCK_GENERAL);
        mobjInterlocks.Add(ILCK_COM_OK).Expressions.Add("TCPIP", mopGFAITCPIPStatus, 1, ilckVariableEqualValueMeansOK, LoadResString(ridTCPCommsFaulted));
        mobjInterlocks.Add(ILCK_OPERATION_OK).IncludedInterlocks.Add(ILCK_GENERAL);
        mobjInterlocks.Add(ILCK_OPERATION_OK).IncludedInterlocks.Add(ILCK_COM_OK);
        // On Error Resume Next — nested handler; see enclosing try/catch
        mobjGlobalVars = new GlobalVars();
        mobjGlobalVars.AddBoolean("Running in Simulation Mode", mgvbSimulation);
        mobjGlobalVars.AddString("Last Sent Command", mgvsLastSentCommand);
        mobjGlobalVars.AddString("TCP Last Response", mgvsLastGFAIResponse);
        mobjGlobalVars.AddDate("TCP Last Response Time", mgvdLastGFAIResponseTime);
        mobjGlobalVars.AddLong("GFAI Response Num Errors", mgvlNumReceiveErrors);
        mobjGlobalVars.AddDate("Last GFAI Response Error Time", mgvdLastGFAIResponseErrTime);
        mobjGlobalVars.AddString("Last GFAI Response Error", mgvsLastGFAIResponseErrMsg);
        mobjGlobalVars.AddLong("Last GFAI Response Error Num", mgvlLastGFAIResponseErrNum);
        mobjGlobalVars.AddString("Last GFAI Response Error Source", mgvsLastGFAIResponseErrSource);
        mobjGlobalVars.AddString("Last GFAI Response Error Cause", mgvsLastGFAIResponseErrInput);
        // On Error GoTo Error — nested handler not restructured
        mobjStatuses = new Statuses();
        mobjStatuses.UseSPG(SPG_STATUSES);
        mobjStatuses.Add(STS_GFAI_INITIALIZED, false, "Acoustic Not Initialized", "Acoustic Initialized");
        objProcess = new Process();
        bReady = true;
        WaitForProcessesStartupComplete();
        objProcessManagerProcess = AddTalentProcessWait(objProcess.Name.ToString(), objProcess);
        lID = objProcessManagerProcess.ID;
        Process.LogMode = logToTalentAndFile;
        // On Error Resume Next — nested handler; see enclosing try/catch
        Initialize(initConstants);
        // On Error GoTo Error — nested handler not restructured
        // On Error GoTo Error_CreatingObjects — nested handler not restructured
        sClassName = "ParManager";
        mobjParManager = new ParManager();
        mobjParManager.OwnerID = lID;
        mobjParManager.AutoRefreshEnabled = false;
        mobjParManager.IgnoreLoadedEvents = true;
        mobjParManager.IgnoreUnloadedEvents = true;
        mobjParManager.AddWithoutFlag("Airtemp", mipAirtemp, ioInput);
        mobjParManager.AddWithoutFlag("WindSpeed", mipWindspeed, ioInput);
        mobjParManager.AddWithoutFlag("WS_Mode", mipWsMode, ioInput);
        mobjParManager.AddWithoutFlag("WS_Achieved", mipWsAchieved, ioInput);
        mobjParManager.AddWithoutFlag("WS_Target", mipWsTarget, ioInput);
        mobjParManager.AddWithoutFlag("RelHumidity", mipRelHumidity, ioInput);
        mobjParManager.AddWithoutFlag("TT_Yaw", mipYaw, ioInput);
        mobjParManager.AddWithoutFlag("TT_YawAchieved", mipYawAchieved, ioInput);
        mobjParManager.AddWithoutFlag("TT_YawTarget", mipYawTarget, ioInput);
        mobjParManager.AddWithoutFlag("PStatic", mipAbsPress, ioInput);
        AddOutput("GFAIComStatus", mopGFAITCPIPStatus, false, false);
        AddOutput("GFAIState", mopGFAIState, true, true);
        AddOutput("GFAIRemote", mopWTCSInControl, false, true);
        if (mccfUpdateTime < 0.05)
        {
            mccfUpdateTime = 0.05;
        }

        if (IsWithinIDE())
        {
            mobjForm.tmrDesignTime.Interval = mccfUpdateTime * 1000;
            mobjForm.tmrDesignTime.Enabled = true;
        }
        else
        {
            tmrMain.Stats.Frequency = Reg_GetValue(HKEY_LOCAL_MACHINE, REG_PROCESSES, null, "TimerFrequency", 10, vbLong);
            tmrMain.Interval = mccfUpdateTime * 1000;
            tmrMain.EventType = TimerPeriodic;
            tmrMain.Notify = this;
            tmrMain.Enabled = true;
        }

        mobjLogFile = new LogFile();
        mobjGFAIMessages = new MailSlot();
        mobjGFAIMessages.MailSlotType = mstCustom;
        mobjGFAIMessages.Name = "GFAIMessages";
        mobjGFAIMessages.OpenOutput();
        mobjGFAIMessageLogFile = new LogFile();
        mobjGFAIMessageLogFile.MailSlot = mobjGFAIMessages;
        if (mobjParManager.Configurations.Count > 0)
        {
            mopGFAITCPIPStatus = 1;
            mgvdLastGFAIResponseTime = Now;
            mdStartTime = Now;
            mobjCommands.Insert(ocInitialize, "Initialize");
        }
        else
        {
            mobjCommands.Insert(ocResetCommunications, "ResetAcousticCommunications");
        }

        tmrRunOnce = null;
        tmrElapsed.Reset();
        goto Done;
        Error();
        sRaiseErrorText = ex.Message;
        EH.Module = "Controller";
        EH.Procedure = "Start";
        EH_MsgBox();
        LogMessage(string.Concat(string.Concat("Error during controller startup: ", sRaiseErrorText), ". See log file for details"), mtFailure);
    Error_CreatingObjects:
        ;
        sRaiseErrorText = string.Concat(string.Concat(string.Concat("Can't create object '", sClassName), "': "), ex.Message);
        bRaiseError = true;
        EH.Module = "Controller";
        EH.Procedure = PROC_NAME;
        EH.Item(1) = sClassName;
        EH_MsgBox();
        goto Done;
    Error_Configurations:
        ;
        sRaiseErrorText = string.Concat("Unable to connect to Configurations collection: ", ex.Message);
        EH.AppError = AE.CantCreateConfigurations;
        EH.Module = TypeName(this);
        EH.Procedure = PROC_NAME;
        EH_MsgBox();
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        if (bRaiseError)
        {
            EH_RaiseError(OLEERR_METHOD_FAILED, "Controller", "Start", sRaiseErrorText)();
        }

        return;
    }

    private void AddOutput(ref string ParameterName, ref float ValueHolder, ref bool ResetToDefaultUponComError, ref bool IsCalculatedBySimulator)
    {
        if (mgvbSimulation)
        {
            mobjParManager.AddWithoutFlag(ParameterName, ValueHolder, IIf(IsCalculatedBySimulator, ioInput, ioOutput)).Tag = ResetToDefaultUponComError;
        }
        else
        {
            mobjParManager.AddWithoutFlag(ParameterName, ValueHolder, ioOutput).Tag = ResetToDefaultUponComError;
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
                EH.Module = "Controller";
                EH.Procedure = "Commands";
                EH_MsgBox();
            }

            goto Done;
        Done:
            ;
            // On Error Resume Next — nested handler; see enclosing try/catch
            if (bRaiseError)
            {
                EH_RaiseError(OLEERR_PROPERTY_GET_FAILED, "Controller", "Commands", sRaiseErrorText)();
            }

            return __result;
        }
    }

    private void ExecuteCommands()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
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
                mobjStatuses(STS_GFAI_INITIALIZED).SetValue(bInited);
                mobjCommands.Clear("Initialize");
                if (bInited)
                {
                    mobjCommands.Insert(ocResetCommunications, "ResetAcousticCommunications");
                }

                break;
            case ocSuspend:
                objCommand.Messages.Flush();
                bInited = false;
                mobjParManager.ReleaseOutputs();
                objCommand.Remove();
                mobjStatuses(STS_GFAI_INITIALIZED).SetValue(bInited);
                SetProcessPriority(priNormal);
                if (IsWithinIDE())
                {
                    this.Shutdown();
                }

                break;
            case ocShutdown:
                objCommand.Messages.Flush();
                objCommand.Remove();
                this.Shutdown();
                break;
            case ocResetCommunications:
                if (mgvbSimulation)
                {
                    objCommand.Remove();
                }

                if (mobjGFAI == null)
                {
                    mobjGFAI = new TalentProcessLib.SocketPoll();
                }

                if (mobjGFAI.State != socConnected)
                {
                    mobjGFAI.Disconnect();
                    mobjGFAI.Caption = "GFAI";
                    mobjGFAI.LogFile = mobjGFAIMessageLogFile;
                    mobjGFAI.logToTalent = false;
                    mobjGFAI.ConnectTimeout = mccfGFAITCPIPConnectTimeout;
                    mobjGFAI.ConnectAsClient(mccsGFAITCPIPAddress, (short)mcclGFAITCPIPPort);
                    if (mobjGFAI.LogMode == logNone)
                    {
                        mobjGFAIMessageLogFile.LogMessage(string.Concat(string.Concat(string.Concat(string.Concat("Connected to ", mccsGFAITCPIPAddress), " using Port "), mcclGFAITCPIPPort), " as client"));
                    }

                    mobjGFAI.BinData.UseStartCharacter = true;
                    mobjGFAI.BinData.StartCharacter = (char)2;
                    mobjGFAI.BinData.UseEndCharacter = true;
                    mobjGFAI.BinData.EndCharacter = (char)3;
                    Pause(1);
                    mobjCommands.Insert(ocGetCurrentControl, "GetAcousticCurrentControl");
                    mobjCommands.Insert(ocGetCurrentStatus, "GetAcousticCurrentStatus");
                }

                objCommand.Remove();
                break;
            case ocStartLogging:
                mobjLogFile.OpenFile("GFAI", lfAppend, lfBinaryLog, true);
                mobjGFAIMessages.OpenOutput();
                mgvbLogInput = true;
                mgvbLogOutput = true;
                if (objCommand.Arguments.Count >= 1)
                {
                    mgvbLogToMsgLog = objCommand.Arguments(1).Value;
                }
                else
                {
                    mgvbLogToMsgLog = false;
                }

                objCommand.Remove();
                break;
            case ocStopLogging:
                // On Error Resume Next — nested handler; see enclosing try/catch
                ;
                mobjGFAIMessages.CloseMailslot();
                mobjLogFile.CloseFile();
                mobjGFAI.LogMode = logNone;
                mgvbLogInput = false;
                mgvbLogOutput = false;
                objCommand.Remove();
                break;
            default:
                CheckInterlock(ocInitialize);
                if (mgvbSimulation == false)
                {
                    if (mopGFAITCPIPStatus == 0)
                    {
                        objCommand.InsertPredecessor(ocResetCommunications, "ResetAcousticCommunications");
                        goto Done;
                    }
                }

                switch (objCommand.Opcode)
                {
                    case ocSendPars:
                        CheckInterlock(ILCK_GENERAL);
                        objCommand.Messages.Flush();
                        SendParValues();
                        objCommand.Remove();
                        break;
                    case ocGetCurrentControl:
                        CheckInterlock(ILCK_GENERAL);
                        SendCommand("GetCurrentControlOwnerRequest");
                        objCommand.Remove();
                        break;
                    case ocGetCurrentStatus:
                        CheckInterlock(ILCK_GENERAL);
                        SendCommand("GetAMSStatusRequest");
                        objCommand.Remove();
                        break;
                    case ocTakeControl:
                        CheckInterlock(ILCK_GENERAL);
                        if (objCommand.FirstTime)
                        {
                            SendCommand("GetControlRequest", Array("timeout", mcclGetControlTimeout));
                        }
                        else if (mopWTCSInControl == 1)
                        {
                            objCommand.Remove();
                            SendParValues();
                        }
                        else if (objCommand.TimeSince >= mcclGetControlTimeout)
                        {
                            objCommand.Abort(,);
                            cmdTimedout();
                        }

                        break;
                    case ocReleaseControl:
                        CheckInterlock(ILCK_GENERAL);
                        SendCommand("UnsolicitedControlReleasedEvent");
                        mopWTCSInControl = 0;
                        objCommand.Remove();
                        break;
                    case ocStartRecord:
                        CheckInterlock(ILCK_GENERAL);
                        if (mopWTCSInControl == 0)
                        {
                            objCommand.InsertPredecessor(ocTakeControl, "TakeAcousticControl");
                        }
                        else
                        {
                            objCommand.Messages.Flush();
                            SendCommand("StartRecordCommand");
                            objCommand.Remove();
                        }

                        break;
                    case ocSetTrigger:
                        CheckInterlock(ILCK_GENERAL);
                        if (mopWTCSInControl == 0)
                        {
                            objCommand.InsertPredecessor(ocTakeControl, "TakeAcousticControl");
                        }
                        else
                        {
                            objCommand.Messages.Flush();
                            SendCommand("SetTriggerCommand");
                            objCommand.Remove();
                        }

                        break;
                    case ocCancelRecord:
                        CheckInterlock(ILCK_GENERAL);
                        if (mopWTCSInControl == 0)
                        {
                            objCommand.InsertPredecessor(ocTakeControl, "TakeAcousticControl");
                        }
                        else
                        {
                            objCommand.Messages.Flush();
                            SendCommand("CancelRecordCommand");
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
        sRaiseErrorText = ex.Message;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        if (ex.HResult & vbObjectError != vbObjectError)
        {
            EH.Module = "Controller";
            EH.Procedure = "ExecuteCommands";
            EH_MsgBox();
        }

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

    internal void ControlLogic()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        if (bInited == true)
        {
            mobjParManager.ReadInputs();
        }

        if (mgvbSimulation == false)
        {
            if (!mobjGFAI == null)
            {
                if (mobjGFAI.State == socConnected)
                {
                    for (I = 1; I <= mobjGFAI.PacketCount; I++)
                    {
                        mgvdLastGFAIResponseTime = Now;
                        sResponse = mobjGFAI.Packets(I);
                        ProcessGFAIResponse(sResponse);
                    }

                    mobjGFAI.Flush();
                }
            }
        }

        ExecuteCommands();
        if (bInited == true)
        {
            GFAIControlLogic();
            WriteOutputs();
        }
        else if (mgvbSimulation == false)
        {
            // On Error Resume Next — nested handler; see enclosing try/catch
            sData = mobjGFAI.GetData;
        }

        if (!mobjStatuses == null)
        {
            mobjStatuses.FlushMessages();
        }

        goto Done;
        Error();
        if (ex.HResult & vbObjectError != vbObjectError)
        {
            EH.Module = "MainLine";
            EH.Procedure = "ControlLogic";
            EH_MsgBox();
        }

        Suspend();
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        return;
    }

    private void WriteOutputs()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        if (mopGFAITCPIPStatus == 1)
        {
            mobjParManager.WriteOutputs();
        }
        else
        {
            // On Error Resume Next — nested handler; see enclosing try/catch
            foreach (var objPar in mobjParManager)
            {
                if (objPar.InputOrOutput == ioOutput)
                {
                    if (objPar.Tag == true)
                    {
                    }
                    else
                    {
                        objPar.WriteValue();
                    }
                }
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
                sFileName = Reg_GetValue(HKEY_LOCAL_MACHINE, REG_GFAI, null, "ControlConstantsFile", sFileName);
                sDefaultsFileName = string.Concat(string.Concat(string.Concat(App.Path, "\\"), App.EXEName), "Defaults.csv");
                mobjConstants.OpenFileWithDefaults(sFileName, sDefaultsFileName);
                mobjConstants.AddString("Priority", mccsPriority, bStatus, sText);
                mobjConstants.AddSingle("UpdateInterval", mccfUpdateTime, bStatus, sText);
                mobjConstants.AddSingle("GFAITCPIPConnectTimeout", mccfGFAITCPIPConnectTimeout, bStatus, sText);
                mobjConstants.AddString("GFAITCPIPAddress", mccsGFAITCPIPAddress, bStatus, sText);
                mobjConstants.AddLong("GFAITCPIPPort", mcclGFAITCPIPPort, bStatus, sText);
                mobjConstants.AddSingle("KeepAliveInterval", mccfKeepAlivePacketInterval, bStatus, sText);
                mobjConstants.AddSingle("SendParsIntervalMax", mccfMaxTimeBetweenSendPars, bStatus, sText);
                mobjConstants.AddBoolean("SendParsUseTargets", mccbUseTargetValues, bStatus, sText);
                mobjConstants.AddLong("GetControlTimeout", mcclGetControlTimeout, bStatus, sText);
                mobjConstants.AddBoolean("SendParsWTCSOnly", mccbSendParsOnlyIfUnderWTCS, bStatus, sText);
                if (bStatus == false)
                {
                    bRaiseError = true;
                    sRaiseErrorText = string.Concat(string.Concat(string.Concat("Can't find the following control constants in ", sFileName), ": "), sText);
                    lRaiseErrorNumber = tceCommandFailed;
                    if (Command == null)
                    {
                    }
                    else
                    {
                        Command.SuppressLogging = true;
                    }

                    objEH = new EH();
                    objEH.AppError = LoadResString(AE.CantLoadConstants);
                    objEH.Item(1) = sText;
                    objEH.FileName = sFileName;
                    objEH.Module = "Controller";
                    objEH.Procedure = "Initialize";
                    LogMessage2(objEH);
                }

                break;
        }

        mlDefPriority = GetControllerPriority(mccsPriority);
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
                    if (Command == null)
                    {
                    }
                    else
                    {
                        Command.SuppressLogging = true;
                    }

                    objEH.VBErr = Err;
                    objEH.AppError = LoadResString(AE.InitializeFailedParsMissing);
                    objEH.Item(1) = Command.Text;
                    objEH.Module = "Controller";
                    objEH.Procedure = "Initialize";
                    if (Command == null)
                    {
                    }
                    else
                    {
                        objEH.Item(1) = Command.Text;
                    }

                    LogMessage2(objEH);
                    goto Done;
                }
                else if (mgvbSimulation)
                {
                }

                bInited = mobjParManager.Ready;
                break;
        }

        goto Done;
        Error();
        bInited = false;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = ex.Message;
        bRaiseError = true;
        if (ex.HResult & vbObjectError != vbObjectError)
        {
            EH.Module = "Controller";
            EH.Procedure = "Initialize";
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
        }

        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        mobjStatuses(STS_GFAI_INITIALIZED).SetValue(bInited);
        mobjConstants.CloseFile();
        objEH = null;
        if (bInited)
        {
            SetProcessPriority(mlDefPriority);
        }
        else
        {
            SetProcessPriority(priNormal);
        }

        if (bRaiseError)
        {
            EH_RaiseError(lRaiseErrorNumber, "Controller", "Initialize", sRaiseErrorText)();
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
            EH.Module = "Controller";
            EH.Procedure = "ControlConstants";
            EH_MsgBox();
            goto Done;
        Done:
            ;
            // On Error Resume Next — nested handler; see enclosing try/catch
            if (bRaiseError)
            {
                EH_RaiseError(OLEERR_PROPERTY_GET_FAILED, "Controller", "ControlConstants", sRaiseErrorText)();
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
            EH.Module = "Controller";
            EH.Procedure = "GlobalVars";
            EH_MsgBox();
            goto Done;
        Done:
            ;
            // On Error Resume Next — nested handler; see enclosing try/catch
            if (bRaiseError)
            {
                EH_RaiseError(OLEERR_PROPERTY_GET_FAILED, "Controller", "GlobalVars", sRaiseErrorText)();
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
                __result = bReady;
                return __result;
            }
            catch
            {
                // On Error Resume Next: exceptions suppressed
            }

            return __result;
        }
    }

    private void SendParValues()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        mgvdLastSendParsTime = Now;
        fWindspeed = mipWindspeed;
        fYaw = mipYaw;
        fAirtemp = mipAirtemp;
        fPAbs = mipAbsPress;
        fRelHumidity = mipRelHumidity;
        if (mccbUseTargetValues)
        {
            switch (mipWsMode)
            {
                case 2:
                    if (mipWsAchieved == 1)
                    {
                        fWindspeed = mipWsTarget;
                    }

                    break;
            }

            if (mipYawAchieved == 1)
            {
                fYaw = mipYawTarget;
            }
        }

        fWindspeed = fWindspeed / 3.6;
        fPAbs = fPAbs / 100;
        vPars(0) = Array("WindSpeed", fWindspeed);
        vPars(1) = Array("YawAngle", fYaw);
        vPars(2) = Array("AirTemperature", fAirtemp);
        vPars(3) = Array("BarometricAirPressure", fPAbs);
        vPars(4) = Array("AirHumidity", fRelHumidity);
        SendCommand("SetParametersRequest", Array("parameters", vPars));
        goto Done;
        Error();
        sRaiseErrorText = ex.Message;
        bRaiseError = true;
        lRaiseErrorNumber = 5;
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

    public void Shutdown()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        SetProcessPriority(priNormal);
        if (!mobjGFAI == null)
        {
            mobjGFAI.Disconnect();
            mobjGFAI = null;
        }

        bReady = false;
        mobjCommands.Clear("Controller shutting down");
        DoEvents();
        bInited = false;
        mobjParManager.ReleaseOutputs();
        mobjEvents.RaiseQuit();
        DoEvents();
        Unload(mobjForm);
        mobjForm = null;
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
        EH.Module = "Controller";
        EH.Procedure = "Shutdown";
        EH_MsgBox();
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        if (bRaiseError)
        {
            EH_RaiseError(OLEERR_METHOD_FAILED, "Controller", "Shutdown", sRaiseErrorText)();
        }

        return;
    }

    private void Class_Initialize()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        Main();
        bRaiseError = false;
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

        mgvbSimulation = Reg_GetValue(HKEY_LOCAL_MACHINE, REG_GFAI, null, "Simulation", false);
        goto Done;
        Error();
        sRaiseErrorText = ex.Message;
        bRaiseError = true;
        if (ex.HResult & vbObjectError != vbObjectError)
        {
            EH.Module = "Controller";
            EH.Procedure = "Class_Initialize";
            EH_MsgBox();
        }

        goto Done;
    Error_CreatingObjects:
        ;
        sRaiseErrorText = string.Concat(string.Concat(string.Concat("Can't create object '", sClassName), "': "), ex.Message);
        bRaiseError = true;
        EH.Module = "Controller";
        EH.Procedure = "Class_Initialize";
        EH.Item(1) = sClassName;
        EH_MsgBox();
        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        if (bRaiseError)
        {
            EH_RaiseError(OLEERR_METHOD_FAILED, "Controller", "Class_Initialize", sRaiseErrorText)();
        }

        return;
    }

    private void Class_Terminate()
    {
        try
        {
            if (!mobjGFAI == null)
            {
                mobjGFAI.Disconnect();
                mobjGFAI = null;
            }

            mobjGFAIMessages.CloseMailslot();
            mobjGFAIMessages = null;
            mobjGFAIMessageLogFile.MailSlot = null;
            mobjGFAIMessageLogFile.CloseFile();
            mobjGFAIMessageLogFile = null;
            mobjLogFile.MailSlot = null;
            mobjLogFile.CloseFile();
            mobjLogFile = null;
            tmrMain.Enabled = false;
            tmrMain.Notify = null;
            tmrMain = null;
            mobjInterlocks.Quit();
            mobjInterlocks = null;
            tmrElapsed = null;
            mobjCommands = null;
            mobjConstants = null;
            mobjGlobalVars = null;
            mobjEvents = null;
            mobjStatuses = null;
            mcolInvalidOpcodes = null;
            mobjSPGConstants = null;
            mobjSPGDataItems = null;
            mobjParManager.Release();
            mobjParManager = null;
            bInited = false;
            return;
        }
        catch
        {
            // On Error Resume Next: exceptions suppressed
        }
    }

    private void ICcrpTimerNotify_Timer(int Milliseconds)
    {
        try
        {
            ControlLogic();
            DoEvents();
            return;
        }
        catch
        {
            // On Error Resume Next: exceptions suppressed
        }
    }

    private void mobjCommands_CommandAdded(ref CommandListLib.Command NewCommand)
    {
        TalentProcessLib.Message objMessage = 0;
        try
        {
            if (NewCommand.SuccessorKey == "")
            {
                if (mgvbLogInput)
                {
                    objMessage = new TalentProcessLib.Message();
                    objMessage.Source = "Command";
                    objMessage.Text = string.Concat(NewCommand.Text, GetCmdArgs(NewCommand.Arguments));
                    objMessage.MessageType = mtCommand;
                    mobjGFAIMessageLogFile.LogMessage2(objMessage);
                }
            }

            return;
        }
        catch
        {
            // On Error Resume Next: exceptions suppressed
        }
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
                    switch (Command.Opcode)
                    {
                        case ocResetCommunications:
                            objEH.AppError = LoadResString(AE.CommunicationsFaulted, Command.Text, CompletionMessage, mccsGFAITCPIPAddress);
                            break;
                        default:
                            objEH.AppError = AppErrorText(tceCommandFailed, Command.Text, CompletionMessage);
                            break;
                    }

                    objEH.Module = TypeName(this);
                    objEH.Procedure = PROC_NAME;
                    objEH.Component = App;
                    LogPopupMessage(objEH, Command.Text);
                    Command.SuppressLogging = true;
                }
                else if (Command.Opcode == ocResetCommunications)
                {
                    objEH = new EH();
                    objEH.AppError = LoadResString(AE.CommunicationsFaulted, Command.Text, CompletionMessage, mccsGFAITCPIPAddress);
                    objEH.Module = TypeName(this);
                    objEH.Procedure = PROC_NAME;
                    objEH.Component = App;
                    LogMessage2(objEH);
                }

                break;
        }// On Error Resume Next — nested handler; see enclosing try/catch
        if (Command.SuccessorKey == "")
        {
            if (mgvbLogInput)
            {
                objMessage = new TalentProcessLib.Message();
                objMessage.Source = IIf(CompletionStatus == cmdCompleted, "Completed", "Failed");
                objMessage.Text = CompletionMessage;
                objMessage.MessageType = IIf(CompletionStatus == cmdCompleted, mtInformation, mtFailure);
                mobjGFAIMessageLogFile.LogMessage2(objMessage);
            }
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

    private void mobjCommands_CommandInserted(ref CommandListLib.Command NewCommand, ref CommandListLib.Command ExistingCommand)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        goto Done;
        Error();
        if (ex.HResult & vbObjectError != vbObjectError)
        {
            EH.Module = "Controller";
            EH.Procedure = "mobjCommands_CommandInserted";
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
            EH.Module = "Controller";
            EH.Procedure = "mobjCommands_MessageFlushed";
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
        CommandListLib.Command objCommand = 0;
        try
        {
            switch (InterlockName)
            {
                case ILCK_GENERAL:
                    break;
            }

            return;
        }
        catch
        {
            // On Error Resume Next: exceptions suppressed
        }
    }

    private void mobjGFAIMessageLogFile_MessageLogged(ref object Message)
    {
        string sText = "";
        int lType = 0;
        object objMessage1 = null;
        object objMessage2 = null;
        string sSource = "";
        try
        {
            objMessage1 = Message;
            CopyObject(objMessage1, objMessage2);
            if (Err)
            {
                mobjLogFile.LogMessage2(Message);
            }
            else
            {
                mobjLogFile.LogMessage2(objMessage2);
            }

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
        if (bInited == true)
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

    public void SendCommand(ref string CommandToSend, params object[] ParPairs)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        sCmd = string.Concat(string.Concat(string.Concat(string.Concat(string.Concat(string.Concat(vbCrLf, "{"), vbCrLf), "  "), "\"cmd\": \""), CommandToSend), "\"");
        for (I = 0; I <= UBound(ParPairs, 1); I++)
        {
            sParName = ParPairs(I)(0);
            vParValue = ParPairs(I)(1);
            sCmd = string.Concat(string.Concat(string.Concat(string.Concat(string.Concat(string.Concat(sCmd, ","), vbCrLf), "  "), "\""), sParName), "\": ");
            if (vParValue is object[])
            {
                sCmd = string.Concat(string.Concat(sCmd, vbCrLf), "  {");
                for (J = 0; J <= UBound(vParValue, 1); J++)
                {
                    sArrayParName = vParValue(J)(0);
                    vArrayParValue = vParValue(J)(1);
                    switch (J)
                    {
                        case 0:
                            break;
                        default:
                            sCmd = string.Concat(sCmd, ",");
                            break;
                    }

                    sCmd = string.Concat(string.Concat(string.Concat(string.Concat(string.Concat(sCmd, vbCrLf), "    "), "\""), sArrayParName), "\": ");
                    switch (VarType(vArrayParValue))
                    {
                        case vbString:
                            sCmd = string.Concat(string.Concat(string.Concat(sCmd, "\""), vArrayParValue), "\"");
                            break;
                        default:
                            sCmd = string.Concat(sCmd, vArrayParValue);
                            break;
                    }
                }

                sCmd = string.Concat(string.Concat(sCmd, vbCrLf), "  }");
            }
            else
            {
                switch (VarType(vParValue))
                {
                    case vbString:
                        sCmd = string.Concat(string.Concat(string.Concat(sCmd, "\""), vParValue), "\"");
                        break;
                    default:
                        sCmd = string.Concat(sCmd, vParValue);
                        break;
                }
            }
        }

        sCmd = string.Concat(string.Concat(string.Concat(sCmd, vbCrLf), "}"), vbCrLf);
        mgvsLastSentCommand = sCmd;
        mgvdLastCommandTime = Now;
        // On Error GoTo Error — nested handler not restructured
        if (mgvbSimulation == false)
        {
            if (mccbGFAITCPIPAnsi)
            {
                xCommand = StrConv(string.Concat(string.Concat((char)2, sCmd), (char)3), vbFromUnicode);
                mobjGFAI.SendData(xCommand);
            }
            else
            {
                mobjGFAI.SendData(string.Concat(string.Concat((char)2, sCmd), (char)3));
            }
        }

        if (mgvbLogOutput)
        {
            sLogText = sCmd.Replace(vbLf, "");
            sLogText = sLogText.Replace(vbCr, "");
            mobjGFAIMessageLogFile.LogMessage(sLogText, "Sending:", mtInformation);
            if (mgvbLogToMsgLog)
            {
                LogMessage(string.Concat("Sending: ", sLogText), mtInformation);
            }
        }

        goto Done;
        Error();
        sRaiseErrorText = LoadResString(ridGFAICmdError, CommandToSend, ex.Message);
        bRaiseError = true;
        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        if (bRaiseError)
        {
            EH_RaiseError(5, TypeName(this), PROC_NAME, sRaiseErrorText);
        }

        return;
    }

    private void LogMessageEH(ref string MessageText, ref string SendingRoutine, object Items = null, MessageTypes MessageType = msgInformation)
    {
        EH objEH = 0;
        string[] sItems;
        int I = 0;
        try
        {
            objEH = new EH();
            objEH.AppError = MessageText;
            if (!IsMissing(Items))
            {
                for (I = 1; I <= 10; I++)
                {
                    // Err.Clear()
                    objEH.Item(I) = Items(I);
                    if (Err)
                    {
                        break;
                    }
                }
            }

            objEH.Module = "Controller";
            objEH.Procedure = SendingRoutine;
            objEH.MessageType = MessageType;
            LogMessage2(objEH);
            return;
        }
        catch
        {
            // On Error Resume Next: exceptions suppressed
        }
    }
}