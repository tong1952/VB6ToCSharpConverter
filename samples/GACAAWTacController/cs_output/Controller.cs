// Converted from Controller.cls by vb6cs
// Date: 2026-05-11 05:50

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
    private ccrpTimer tmrMain;
    private ccrpTimer tmrRunOnce;
    private ccrpStopWatch tmrElapsed;
    private Interlocks mobjInterlocks;
    private TalentProcessLib.Events mobjEvents;
    private frmControl mobjForm;
    private SharedPropertyGroup mobjSPGConstants;
    private SharedPropertyGroup mobjSPGDataItems;
    private TalentProcessLib.LogFile mobjLogFile;
    private MailSlot mobjAcousticMessages;
    private LogFile mobjAcousticMessageLogFile;
    private WinHttp.WinHttpRequest mobjWinHTTPRequest;
    private int lID;
    private bool bReady;
    private bool bInited;
    private const bool bKeepStats = false;
    private bool bRealTime;
    private DateTime mgvdLastCommandTime;
    private object mvInterlockCheckValue;
    private DateTime mdStartTime;
    private string[] msAPI;
    private string[] msAPIValue;
    private string[] msAPIValueDescription;
    private string msAPIPrefix;
    public Collection mcolInvalidOpcodes;
    private enum InitTypes
    {
        initAll = 0,
        initConstants = 1,
        initPars = 2
    }

    private string mccsPriority;
    private float mccfUpdateTime;
    private float mccfBKConnectTimeout;
    private string mccsBKConnectAddress;
    private int mcclBKConnectPort;
    private const bool mccbBKConnectAnsi = false;
    private float mccfKeepAlivePacketInterval;
    private float mccfMaxTimeBetweenSendPars;
    private bool mccbUseTargetValues;
    private int mcclGetControlTimeout;
    private bool mccbSendParsOnlyIfUnderFCAS;
    private float mopAcousticStatus;
    private float mopAcousticRecNo;
    private bool mgvbLogOutput;
    private bool mgvbLogInput;
    private bool mgvbLogToMsgLog;
    private bool mgvbLogKeepAlive;
    private int mgvlNumReceiveErrors;
    private int mlDefPriority;
    private string mgvsLastSentCommand;
    private bool mgvbSimulation;
    private DateTime mgvdLastBKConnectResponseTime;
    private string mgvsLastBKConnectResponse;
    private DateTime mgvdBKConnectDataLastRequestTime;
    private DateTime mgvdLastDataStreamLoggedTime;
    private DateTime mgvdLastSendParsTime;
    private string mgvsLastBKConnectResponseErrMsg;
    private int mgvlLastBKConnectResponseErrNum;
    private string mgvsLastBKConnectResponseErrSource;
    private string mgvsLastBKConnectResponseErrInput;
    private DateTime mgvdLastBKConnectResponseErrTime;
    private DateTime mgvdLastKeepAlivePacketTime;
    private const object NUM_API = 25;
    private enum bkapi
    {
        bkapiLoadProject = 1,
        bkapiLoadTemplate = 2,
        bkapiStartRecorder = 3,
        bkapiStopRecorder = 4,
        bkapiResetMonitor = 5,
        bkapiManualStart = 6,
        bkapiManualStop = 7,
        bkapiDeleteLatest = 8,
        bkapiSet_RecordingDirectory = 9,
        bkapiGet_RecordingDirectory = 10,
        bkapiSet_RecordingFileName = 11,
        bkapiGet_RecordingFileName = 12,
        bkapiGet_RecordingNumber = 13,
        bkapiGet_FullRecordingFileName = 14,
        bkapiSet_Recording = 15,
        bkapiGet_Recording = 16,
        bkapiSet_AutoPack = 17,
        bkapiGet_AutoPack = 18,
        bkapiSet_AutoStore = 19,
        bkapiGet_AutoStore = 20,
        bkapiSet_Marker = 21,
        bkapiGet_Marker = 22,
        bkapiAddMarker = 23,
        bkapiDeleteMarker = 24,
        bkapiSetTrigger = 25
    }

    private void AbortCurrentCommand(string CompletionMessage = "")
    {
        if (mobjCommands.Count > 0)
        {
            mobjCommands(1).Abort(CompletionMessage);
        }
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

    public void CheckInterlock(ref object opcode, object CheckValue = null)
    {
        const object PROC_NAME = "CheckInterlock";
        bool bRaiseError = false;
        string sRaiseErrorText = "";
        int lRaiseErrorNumber = 0;
        string sIndex = "";
        Interlock objInterlock = 0;
        try
        {
            bRaiseError = false;
            // On Error Resume Next — nested handler; see enclosing try/catch
            sIndex = opcode;
            objInterlock = mobjInterlocks(sIndex);
            if (Err)
            {
                goto Done;
            }

            mvInterlockCheckValue = CheckValue;
            // On Error GoTo Error — nested handler not restructured
            objInterlock.Check();
        }
        catch (Exception ex)
        {
            sRaiseErrorText = ex.Message;
            bRaiseError = true;
            lRaiseErrorNumber = ex.HResult;
        }

    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        if (bRaiseError)
        {
            throw new Exception(sRaiseErrorText);
        }

        return;
    }

    private void Class_Initialize()
    {
        const object PROC_NAME = "Class_Initialize";
        bool bRaiseError = false;
        string sRaiseErrorText = "";
        string sClassName = "";
        int i = 0;
        try
        {
            Main();
            bRaiseError = false;
            // On Error GoTo Error_CreatingObjects — nested handler not restructured
            sClassName = "ccrpTimers6.ccrpTimer";
            tmrMain = new ccrpTimers6.ccrpTimer();
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

            mgvbSimulation = Reg_GetValue(HKEY_LOCAL_MACHINE, REG_ACOUSTIC, null, "Simulation", false);
            msAPI = new string[NUM_API + 1];
            msAPIValue = new string[NUM_API + 1];
            msAPIValueDescription = new string[NUM_API + 1];
            i = 1;
            msAPI(i) = "/ProgramService/LoadProject?name=";
            msAPIValueDescription(i) = "ProjectName";
            i = i + 1;
            msAPI(i) = "/ProgramService/LoadTemplate?name=";
            msAPIValueDescription(i) = "TemplateName";
            i = i + 1;
            msAPI(i) = "/DataRecorderService/StartRecorder";
            i = i + 1;
            msAPI(i) = "/DataRecorderService/StopRecorder";
            i = i + 1;
            msAPI(i) = "/DataRecorderService/ResetMonitor";
            i = i + 1;
            msAPI(i) = "/DataRecorderService/ManualStart";
            i = i + 1;
            msAPI(i) = "/DataRecorderService/ManualStop";
            i = i + 1;
            msAPI(i) = "/DataRecorderService/DeleteLatest";
            i = i + 1;
            msAPI(i) = "/DataRecorderService/set_RecordingDirectory?value=";
            msAPIValueDescription(i) = "newPath";
            i = i + 1;
            msAPI(i) = "/DataRecorderService/get_RecordingDirectory";
            i = i + 1;
            msAPI(i) = "/DataRecorderService/set_RecordingFileName?value=";
            msAPIValueDescription(i) = "newPath";
            i = i + 1;
            msAPI(i) = "/DataRecorderService/get_RecordingFileName";
            i = i + 1;
            msAPI(i) = "/DataRecorderService/get_RecordingNumber";
            i = i + 1;
            msAPI(i) = "/DataRecorderService/get_FullRecordingFileName";
            i = i + 1;
            msAPI(i) = "/DataRecorderService/set_Recording?value=";
            msAPIValueDescription(i) = "newPath";
            i = i + 1;
            msAPI(i) = "/DataRecorderService/get_Recording";
            i = i + 1;
            msAPI(i) = "/DataRecorderService/set_AutoPack?value=";
            msAPIValueDescription(i) = "true/false";
            i = i + 1;
            msAPI(i) = "/DataRecorderService/get_AutoPack";
            i = i + 1;
            msAPI(i) = "/DataRecorderService/set_AutoStore?value=";
            msAPIValueDescription(i) = "true/false";
            i = i + 1;
            msAPI(i) = "/DataRecorderService/get_AutoStore";
            i = i + 1;
            msAPI(i) = "/DataRecorderService/set_Marker?value=";
            msAPIValueDescription(i) = "newName";
            i = i + 1;
            msAPI(i) = "/DataRecorderService/get_Marker";
            i = i + 1;
            msAPI(i) = "/DataRecorderService/AddMarker";
            i = i + 1;
            msAPI(i) = "/DataRecorderService/DeleteMarker";
            i = i + 1;
            msAPI(i) = "/TriggerSetupService/SetTrigger?triggerProperties=";
            msAPIValueDescription(i) = "xml";
            i = i + 1;
        }
        catch (Exception ex)
        {
            sRaiseErrorText = ex.Message;
            bRaiseError = true;
            if (ex.HResult & vbObjectError != vbObjectError)
            {
                EH.Module = TypeName(this);
                EH.Procedure = PROC_NAME;
                EH_MsgBox();
            }

        Error_CreatingObjects:
            ;
            sRaiseErrorText = string.Concat(string.Concat(string.Concat("Can't create object '", sClassName), "': "), ex.Message);
            bRaiseError = true;
            EH.Module = TypeName(this);
            EH.Procedure = PROC_NAME;
            EH.Item(1) = sClassName;
            EH_MsgBox();
        }

    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        if (bRaiseError)
        {
            EH_RaiseError(OLEERR_METHOD_FAILED, TypeName(this), PROC_NAME, sRaiseErrorText)();
        }

        return;
    }

    private void Class_Terminate()
    {
        try
        {
            mobjAcousticMessages.CloseMailslot();
            mobjAcousticMessages = null;
            mobjAcousticMessageLogFile.MailSlot = null;
            mobjAcousticMessageLogFile.CloseFile();
            mobjAcousticMessageLogFile = null;
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

    public CommandListLib.Commands Commands
    {
        get
        {
            CommandListLib.Commands __result = 0;
            Attribute(Commands.VB_Description == "Returns the command list");
            const object PROC_NAME = "Commands (Get)";
            bool bRaiseError = false;
            string sRaiseErrorText = "";
            try
            {
                bRaiseError = false;
                __result = mobjCommands;
            }
            catch (Exception ex)
            {
                sRaiseErrorText = ex.Message;
                bRaiseError = true;
                if (ex.HResult & vbObjectError != vbObjectError)
                {
                    EH.Module = TypeName(this);
                    EH.Procedure = PROC_NAME;
                    EH_MsgBox();
                }
            }

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

    public ControlLib.ControlConstants ControlConstants
    {
        get
        {
            ControlLib.ControlConstants __result = 0;
            Attribute(ControlConstants.VB_Description == "Returns the collection of control constants for use by an external debugger");
            const object PROC_NAME = "ControlConstants (Get)";
            bool bRaiseError = false;
            string sRaiseErrorText = "";
            try
            {
                bRaiseError = false;
                __result = mobjConstants;
            }
            catch (Exception ex)
            {
                sRaiseErrorText = ex.Message;
                bRaiseError = true;
                EH.Module = TypeName(this);
                EH.Procedure = PROC_NAME;
                EH_MsgBox();
            }

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

    internal void ControlLogic()
    {
        const object PROC_NAME = "ControlLogic";
        int lElapsed = 0;
        string sData = "";
        bool bExists = false;
        float fCommSendTime = 0;
        int i = 0;
        string sResponse = "";
        object xByte = null;
        object xPacket = null;
        int lBytesAvailable = 0;
        try
        {
            if (bInited == true)
            {
                mobjParManager.ReadInputs();
            }

            if (mgvbSimulation == false)
            {
            }

            ExecuteCommands();
            if (bInited == true)
            {
                AcousticControlLogic();
                mobjStatuses(STS_ACOUSTIC_FAULTED).SetValue((bool)mopAcousticStatus);
                WriteOutputs();
            }
            else if (mgvbSimulation == false)
            {
                // On Error Resume Next — nested handler; see enclosing try/catch
            }

            if (!mobjStatuses == null)
            {
                mobjStatuses.FlushMessages();
            }

            goto Done;
        }
        catch (Exception ex)
        {
            if (ex.HResult & vbObjectError != vbObjectError)
            {
                EH.Module = TypeName(this);
                EH.Procedure = PROC_NAME;
                EH_MsgBox();
            }

            Suspend();
        // Resume Next
        Done:
            ;
            // On Error Resume Next — nested handler; see enclosing try/catch
            return;
        }
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
            __result = mobjCommands(1).opcode;
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

    private void ExecuteCommands()
    {
        Attribute(ExecuteCommands.VB_Description == "Executes the next command on the list");
        const object PROC_NAME = "ExecuteCommands";
        bool bRaiseError = false;
        string sRaiseErrorText = "";
        int lRaiseErrorNumber = 0;
        // TODO: Static local 'ReEntered' — move to a class-level static field
        bool ReEntered = false;
        CommandListLib.Command objCommand = 0;
        int i = 0;
        EH objEH = 0;
        bool bFirstTime = false;
        short iInt = 0;
        float[] fTargets = new float[7 + 1];
        string[] sAxis = new string[3 + 1];
        try
        {
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
            switch (objCommand.opcode)
            {
                case ocInitialize:
                    this.Initialize(,);
                    objCommand();
                    objCommand.Remove();
                    mobjStatuses(STS_ACOUSTIC_INITIALIZED).SetValue(bInited);
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
                    mobjStatuses(STS_ACOUSTIC_INITIALIZED).SetValue(bInited);
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

                    if (!mobjWinHTTPRequest == null)
                    {
                        mobjWinHTTPRequest.Abort();
                        mobjWinHTTPRequest = null;
                    }

                    mobjWinHTTPRequest = new WinHttp.WinHttpRequest();
                    mobjWinHTTPRequest.SetTimeouts(1000, 1000, 1000, 1000);
                    objCommand.Remove();
                    break;
                case ocStartLogging:
                    mobjLogFile.OpenFile("BKConnect", lfAppend, lfBinaryLog, true);
                    mobjAcousticMessages.OpenOutput();
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
                    mobjAcousticMessages.CloseMailslot();
                    mobjLogFile.CloseFile();
                    mgvbLogInput = false;
                    mgvbLogOutput = false;
                    objCommand.Remove();
                    break;
                default:
                    CheckInterlock(ocInitialize);
                    if (mgvbSimulation == false)
                    {
                        if (mopAcousticStatus == 0)
                        {
                            if (objCommand.FirstTime)
                            {
                                objCommand.ResetTime();
                                objCommand.InsertPredecessor(ocResetCommunications, "ResetAcousticCommunications");
                                goto Done;
                            }
                            else if (SecondsSinceTime(objCommand.TimeSince) > 10)
                            {
                                objCommand.Abort("No connection to BK Connect");
                            }
                        }
                    }

                    switch (objCommand.opcode)
                    {
                        case ocStartRecord:
                            CheckInterlock(ILCK_GENERAL);
                            objCommand.Messages.Flush();
                            SendCommand(bkapiStartRecorder);
                            objCommand.Remove();
                            break;
                        case ocStopRecord:
                            CheckInterlock(ILCK_GENERAL);
                            objCommand.Messages.Flush();
                            SendCommand(bkapiStopRecorder);
                            objCommand.Remove();
                            break;
                        case ocSendCommand:
                            CheckInterlock(ILCK_GENERAL);
                            objCommand.Messages.Flush();
                            SendCommand.Arguments(1).Value();
                            objCommand.Remove();
                            break;
                        case ocSetIPAddr:
                            CheckInterlock(ILCK_GENERAL);
                            objCommand.Messages.Flush();
                            mccsBKConnectAddress = objCommand.Arguments(1).Value;
                            msAPIPrefix = string.Concat(string.Concat(string.Concat("http://", mccsBKConnectAddress), ":"), mcclBKConnectPort);
                            objCommand.Remove();
                            break;
                        default:
                            objCommand.Abort(string.Concat("Unrecognized opcode: ", objCommand.opcode));
                            break;
                    }

                    break;
            }

            goto Done;
        }
        catch (Exception ex)
        {
            sRaiseErrorText = ex.Message;
            lRaiseErrorNumber = CCommandError(ex.HResult);
            if (ex.HResult & vbObjectError != vbObjectError)
            {
                EH.Module = TypeName(this);
                EH.Procedure = PROC_NAME;
                EH_MsgBox();
            }
        }

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

    public object GetInvalidOpcodes(ref object opcode)
    {
        object __result = null;
        string sIndex = "";
        try
        {
            sIndex = opcode;
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

    private void AcousticControlLogic()
    {
        Attribute(AcousticControlLogic.VB_Description == "Closed loop control logic");
        const object PROC_NAME = "AcousticControlLogic";
        byte[] xData;
        int i = 0;
        string sResponse = "";
        ManagedPar objPar = 0;
        bool bExists = false;
        float fValue = 0;
        bool bCmdReqd = false;
        try
        {
            // On Error Resume Next — nested handler; see enclosing try/catch
            if (mgvbSimulation)
            {
                mopAcousticStatus = NOFAULT;
            }

            if (Err)
            {
                mopAcousticStatus = FAULT;
            }

            if (mgvbSimulation == false)
            {
                if (!mobjWinHTTPRequest == null)
                {
                    if (mccfKeepAlivePacketInterval > 0)
                    {
                        if (SecondsSinceTime(mgvdLastKeepAlivePacketTime) >= mccfKeepAlivePacketInterval)
                        {
                            // On Error Resume Next — nested handler; see enclosing try/catch
                            SendCommand(bkapiGet_RecordingNumber);
                            mgvdLastKeepAlivePacketTime = Now;
                            // On Error GoTo Error — nested handler not restructured
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            if (bInited)
            {
                EH.Module = TypeName(this);
                EH.Procedure = PROC_NAME;
                EH_MsgBox();
                Suspend();
            }
        }

    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        return;
    }

    public ControlLib.GlobalVars GlobalVars
    {
        get
        {
            ControlLib.GlobalVars __result = 0;
            const object PROC_NAME = "GlobalVars (Get)";
            bool bRaiseError = false;
            string sRaiseErrorText = "";
            try
            {
                bRaiseError = false;
                __result = mobjGlobalVars;
            }
            catch (Exception ex)
            {
                sRaiseErrorText = ex.Message;
                bRaiseError = true;
                EH.Module = TypeName(this);
                EH.Procedure = PROC_NAME;
                EH_MsgBox();
            }

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

    internal void Initialize(InitTypes InitType = initAll, Command Command = 0)
    {
        const object PROC_NAME = "Initialize";
        bool bRaiseError = false;
        string sRaiseErrorText = "";
        int lRaiseErrorNumber = 0;
        bool bStatus = false;
        string sText = "";
        string sFileName = "";
        EH objEH = 0;
        bool bExists = false;
        string sPath = "";
        string sDefaultsFileName = "";
        try
        {
            bRaiseError = false;
            switch (InitType)
            {
                case initConstants:
                case initAll:
                    bStatus = true;
                    sPath = GetDefaultCCFolder(App.Path);
                    sFileName = string.Concat(string.Concat(sPath, PROCESS_NAME), ".csv");
                    sFileName = Reg_GetValue(HKEY_LOCAL_MACHINE, REG_ACOUSTIC, null, "ControlConstantsFile", sFileName);
                    sDefaultsFileName = string.Concat(string.Concat(string.Concat(App.Path, "\\"), App.EXEName), "Defaults.csv");
                    mobjConstants.OpenFileWithDefaults(sFileName, sDefaultsFileName);
                    mobjConstants.AddString("Priority", mccsPriority, bStatus, sText);
                    mobjConstants.AddSingle("UpdateInterval", mccfUpdateTime, bStatus, sText);
                    mobjConstants.AddSingle("BKConnectTimeout", mccfBKConnectTimeout, bStatus, sText);
                    mobjConstants.AddString("BKConnectAddress", mccsBKConnectAddress, bStatus, sText);
                    mobjConstants.AddLong("BKConnectPort", mcclBKConnectPort, bStatus, sText);
                    mobjConstants.AddSingle("KeepAliveInterval", mccfKeepAlivePacketInterval, bStatus, sText);
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
                        objEH.Module = TypeName(this);
                        objEH.Procedure = PROC_NAME;
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
                        goto Done;
                    }
                    else if (mgvbSimulation)
                    {
                    }

                    bInited = mobjParManager.Ready;
                    break;
            }

            msAPIPrefix = string.Concat(string.Concat(string.Concat("http://", mccsBKConnectAddress), ":"), mcclBKConnectPort);
        }
        catch (Exception ex)
        {
            bInited = false;
            lRaiseErrorNumber = CCommandError(ex.HResult);
            sRaiseErrorText = ex.Message;
            bRaiseError = true;
            if (ex.HResult & vbObjectError != vbObjectError)
            {
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
            }
        }

    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        mobjStatuses(STS_ACOUSTIC_INITIALIZED).SetValue(bInited);
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
            EH_RaiseError(lRaiseErrorNumber, TypeName(this), PROC_NAME, sRaiseErrorText)();
        }

        return;
    }

    public bool Initialized
    {
        get
        {
            bool __result = false;
            Attribute(Initialized.VB_Description == "Returns whether or not the controller is initialized (controlling) or not");
            const object PROC_NAME = "Initialized (Get)";
            bool bRaiseError = false;
            string sRaiseErrorText = "";
            try
            {
                bRaiseError = false;
                __result = bInited;
            }
            catch (Exception ex)
            {
                sRaiseErrorText = ex.Message;
                bRaiseError = true;
                if (ex.HResult & vbObjectError != vbObjectError)
                {
                    EH.Module = TypeName(this);
                    EH.Procedure = PROC_NAME;
                    EH_MsgBox();
                }
            }

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

    public Interlocks Interlocks
    {
        get
        {
            Interlocks __result = 0;
            __result = mobjInterlocks;
            return __result;
        }
    }

    private void LogMessageEH(ref string MessageText, ref string SendingRoutine, object Items = null, MessageTypes MessageType = msgInformation)
    {
        EH objEH = 0;
        string[] sItems;
        int i = 0;
        try
        {
            objEH = new EH();
            objEH.AppError = MessageText;
            if (!IsMissing(Items))
            {
                for (i = 1; i <= 10; i++)
                {
                    // Err.Clear()
                    objEH.Item(i) = Items(i);
                    if (Err)
                    {
                        break;
                    }
                }
            }

            objEH.Module = TypeName(this);
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

    private void mobjCommands_CommandAdded(ref CommandListLib.Command NewCommand)
    {
        TalentProcessLib.Message objMessage = 0;
        string sStatus = "";
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
                    mobjAcousticMessageLogFile.LogMessage2(objMessage);
                }
            }

            sStatus = "";
            switch (NewCommand.opcode)
            {
                case ocInitialize:
                    sStatus = STS_ACOUSTIC_INITIALIZED;
                    break;
            }

            if (sStatus != "")
            {
                mobjStatuses.Reset(sStatus, false);
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
        const object PROC_NAME = "mobjCommands_CommandCompleted";
        EH objEH = 0;
        TalentProcessLib.Message objMessage = 0;
        int i = 0;
        try
        {
            switch (CompletionStatus)
            {
                case cmdAborted:
                case cmdTimedout:
                    if (Command.SuppressLogging == false)
                    {
                        objEH = new EH();
                        switch (Command.opcode)
                        {
                            case ocResetCommunications:
                                objEH.AppError = LoadResString(AE.CommunicationsFaulted, Command.Text, CompletionMessage, mccsBKConnectAddress);
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
                    else if (Command.opcode == ocResetCommunications)
                    {
                        objEH = new EH();
                        objEH.AppError = LoadResString(AE.CommunicationsFaulted, Command.Text, CompletionMessage, mccsBKConnectAddress);
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
                    mobjAcousticMessageLogFile.LogMessage2(objMessage);
                }
            }
        }
        catch (Exception ex)
        {
        }

    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        objEH = null;
        return;
    }

    private void mobjCommands_CommandInserted(ref CommandListLib.Command NewCommand, ref CommandListLib.Command ExistingCommand)
    {
        const object PROC_NAME = "mobjCommands_CommandInserted";
        try
        {
        }
        catch (Exception ex)
        {
            if (ex.HResult & vbObjectError != vbObjectError)
            {
                EH.Module = TypeName(this);
                EH.Procedure = PROC_NAME;
                EH_MsgBox();
            }
        }

    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        return;
    }

    private void mobjCommands_MessageFlushed(ref string MessageText, ref int MessageType)
    {
        const object PROC_NAME = "mobjCommands_MessageFlushed";
        try
        {
            LogMessage(MessageText, MessageType);
        }
        catch (Exception ex)
        {
            if (ex.HResult & vbObjectError != vbObjectError)
            {
                EH.Module = TypeName(this);
                EH.Procedure = PROC_NAME;
                EH_MsgBox();
            }
        }

    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        return;
    }

    private void mobjAcousticMessageLogFile_MessageLogged(ref object Message)
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

    private void mobjParManager_BeforeUnloaded(ref string ConfigurationName, ref bool Cancel, ref Collection Cancellers)
    {
        const object PROC_NAME = "mobjParManager_BeforeUnloaded";
        try
        {
            if (bInited == true)
            {
                Cancel = true;
                Cancellers.Add(PROCESS_NAME);
            }
        }
        catch (Exception ex)
        {
            if (ex.HResult & vbObjectError != vbObjectError)
            {
                EH.Module = TypeName(this);
                EH.Procedure = PROC_NAME;
                EH_MsgBox();
            }
        }

    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        return;
    }

    private void mobjStatuses_StatusChanged(ref string Message)
    {
        const object PROC_NAME = "mobjStatuses_StatusChanged";
        try
        {
            LogMessage(Message, mtStatusChange);
        }
        catch (Exception ex)
        {
            if (ex.HResult & vbObjectError != vbObjectError)
            {
                EH.Module = TypeName(this);
                EH.Procedure = PROC_NAME;
                EH_MsgBox();
            }
        }

    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        return;
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

    private void ProcessBKConnectResponse(string BKConnectResponse)
    {
        Attribute(ProcessBKConnectResponse.VB_Description == "Process a received data packet");
        const object PROC_NAME = "ProcessBKConnectResponse";
        // TODO: Static local 'vLastErrorMessageTime' — move to a class-level static field
        DateTime vLastErrorMessageTime = DateTime.MinValue;
        string[] sItems;
        int lNumItems = 0;
        string[] sSubItems;
        int lNumSubItems = 0;
        string sCommand = "";
        string sFaultText = "";
        int lErrDescStart = 0;
        int i = 0;
        MessageTypes lMessageType = 0;
        string sLogText = "";
        try
        {
        }
        catch (Exception ex)
        {
            mgvlNumReceiveErrors = mgvlNumReceiveErrors + 1;
            mgvsLastBKConnectResponseErrMsg = ex.Message;
            mgvlLastBKConnectResponseErrNum = ex.HResult;
            mgvsLastBKConnectResponseErrSource = ex.Message;
            mgvsLastBKConnectResponseErrInput = BKConnectResponse;
            if (ex.HResult & vbObjectError != vbObjectError)
            {
                EH.VBErr = Err;
                if (SecondsSinceTime(vLastErrorMessageTime) >= 300)
                {
                    EH.AppError = LoadResString(AE.BKConnectResponseError);
                    EH.Item(1) = BKConnectResponse;
                    EH.Module = TypeName(this);
                    EH.Procedure = PROC_NAME;
                    LogMessage2(EH);
                    vLastErrorMessageTime = Now;
                }

                EH.AppError = null;
            }
        }

    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        return;
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

    private void SendCommand(ref object API, params object[] ParValues)
    {
        const object PROC_NAME = "SendCommand";
        bool bRaiseError = false;
        string sRaiseErrorText = "";
        int lRaiseErrorNumber = 0;
        string sCmd = "";
        string sParName = "";
        object vParValue = null;
        string sParValue = "";
        string sArrayParName = "";
        object vArrayParValue = null;
        int i = 0;
        int j = 0;
        byte[] xCommand;
        string sLogText = "";
        bool bLog = false;
        int APIindex = 0;
        try
        {
            bRaiseError = false;
            bLog = true;
            if (double.TryParse(API.ToString(), _ _))
            {
                APIindex = API;
                if (APIindex == bkapiGet_RecordingNumber)
                {
                    bLog = mgvbLogKeepAlive;
                }

                if (msAPIValueDescription(APIindex) != "")
                {
                    if (ParValues is object[])
                    {
                        sParValue = ParValues(0);
                        sCmd = string.Concat(string.Concat(msAPIPrefix, msAPI(APIindex)), URLEncode(sParValue));
                    }
                    else
                    {
                        throw new Exception(5);
                        "API "(&);
                        msAPI(APIindex)(&);
                        " is missing required argument '"(&);
                        msAPIValueDescription(APIindex)(&);
                        "'"();
                    }
                }
                else
                {
                    sCmd = string.Concat(msAPIPrefix, msAPI(APIindex));
                }
            }
            else
            {
                sCmd = API;
            }

            mgvsLastSentCommand = sCmd;
            mgvdLastCommandTime = Now;
            if (mgvbLogOutput & bLog)
            {
                sLogText = sCmd;
                mobjAcousticMessageLogFile.LogMessage(sLogText, "Sending:", mtInformation);
                if (mgvbLogToMsgLog)
                {
                    LogMessage(string.Concat("Sending: ", sLogText), mtInformation);
                }
            }// On Error GoTo Error — nested handler not restructured
            if (mgvbSimulation == false)
            {
                mobjWinHTTPRequest.Open("GET", sCmd, false);
                // On Error Resume Next — nested handler; see enclosing try/catch
                mobjWinHTTPRequest.Send();
                if (Err)
                {
                    mopAcousticStatus = FAULT;
                    mobjWinHTTPRequest.Abort();
                    goto Done;
                }

                mobjWinHTTPRequest.WaitForResponse();
                if (mobjWinHTTPRequest.Status == 200)
                {
                    mopAcousticStatus = NOFAULT;
                    if (APIindex == bkapiGet_RecordingNumber)
                    {
                        mopAcousticRecNo = mobjWinHTTPRequest.ResponseText;
                    }
                }
                else
                {
                    mopAcousticStatus = FAULT;
                }

                if (mgvbLogInput & bLog)
                {
                    mobjAcousticMessageLogFile.LogMessage.ResponseText(,);
                    "Received:"(,);
                    mtInformation();
                    if (mgvbLogToMsgLog)
                    {
                        LogMessage(string.Concat("Received: ", mobjWinHTTPRequest.ResponseText), mtInformation);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            sRaiseErrorText = LoadResString(ridBKConnectCmdError, sCmd, ex.Message);
            bRaiseError = true;
        }

    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        if (bRaiseError)
        {
            EH_RaiseError(5, TypeName(this), PROC_NAME, sRaiseErrorText);
        }

        return;
    }

    public void Shutdown()
    {
        const object PROC_NAME = "Shutdown";
        bool bRaiseError = false;
        string sRaiseErrorText = "";
        SOM objSOM = 0;
        try
        {
            bRaiseError = false;
            SetProcessPriority(priNormal);
            if (!mobjWinHTTPRequest == null)
            {
                mobjWinHTTPRequest.Abort();
                mobjWinHTTPRequest = null;
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
        }
        catch (Exception ex)
        {
            sRaiseErrorText = ex.Message;
            bRaiseError = true;
            EH.Module = TypeName(this);
            EH.Procedure = PROC_NAME;
            EH_MsgBox();
        // Resume Next
        Done:
            ;
            // On Error Resume Next — nested handler; see enclosing try/catch
            if (bRaiseError)
            {
                EH_RaiseError(OLEERR_METHOD_FAILED, TypeName(this), PROC_NAME, sRaiseErrorText)();
            }

            return;
        }
    }

    private void Start()
    {
        Attribute(Start.VB_Description == "Private start routine");
        const object PROC_NAME = "Start";
        bool bRaiseError = false;
        string sRaiseErrorText = "";
        int i = 0;
        object objProcess = null;
        string sClassName = "";
        object objProcessManagerProcess = null;
        SharedPropertyGroupManager objSPGM = 0;
        bool bExists = false;
        try
        {
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
            mobjInterlocks.Add(ILCK_COM_OK).Expressions.Add("TCPIP", mopAcousticStatus, 1, ilckVariableEqualValueMeansOK, LoadResString(ridTCPCommsFaulted));
            mobjInterlocks.Add(ILCK_OPERATION_OK).IncludedInterlocks.Add(ILCK_GENERAL);
            mobjInterlocks.Add(ILCK_OPERATION_OK).IncludedInterlocks.Add(ILCK_COM_OK);
            // On Error Resume Next — nested handler; see enclosing try/catch
            mobjGlobalVars = new GlobalVars();
            mobjGlobalVars.AddBoolean("Running in Simulation Mode", mgvbSimulation);
            mobjGlobalVars.AddString("Last Sent Command", mgvsLastSentCommand);
            mobjGlobalVars.AddString("TCP Last Response", mgvsLastBKConnectResponse);
            mobjGlobalVars.AddDate("TCP Last Response Time", mgvdLastBKConnectResponseTime);
            mobjGlobalVars.AddLong("BKConnect Response Num Errors", mgvlNumReceiveErrors);
            mobjGlobalVars.AddDate("Last BKConnect Response Error Time", mgvdLastBKConnectResponseErrTime);
            mobjGlobalVars.AddString("Last BKConnect Response Error", mgvsLastBKConnectResponseErrMsg);
            mobjGlobalVars.AddLong("Last BKConnect Response Error Num", mgvlLastBKConnectResponseErrNum);
            mobjGlobalVars.AddString("Last BKConnect Response Error Source", mgvsLastBKConnectResponseErrSource);
            mobjGlobalVars.AddString("Last BKConnect Response Error Cause", mgvsLastBKConnectResponseErrInput);
            mobjGlobalVars.AddBoolean("Log (repetitive) Keep-Alive Communications", mgvbLogKeepAlive);
            // On Error GoTo Error — nested handler not restructured
            mobjStatuses = new Statuses();
            mobjStatuses.UseSPG(SPG_STATUSES);
            mobjStatuses.Add(STS_ACOUSTIC_INITIALIZED, false, "Acoustic Not Initialized", "Acoustic Initialized");
            mobjStatuses.Add(STS_ACOUSTIC_FAULTED, false, "Acoustic Faulted", "Acoustic Not Faulted");
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
            AddOutput("AcousticStatus", mopAcousticStatus, false, false);
            AddOutput("AcousticRecNo", mopAcousticRecNo, true, true);
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
            mobjAcousticMessages = new MailSlot();
            mobjAcousticMessages.MailSlotType = mstCustom;
            mobjAcousticMessages.Name = "AcousticMessages";
            mobjAcousticMessages.OpenOutput();
            mobjAcousticMessageLogFile = new LogFile();
            mobjAcousticMessageLogFile.MailSlot = mobjAcousticMessages;
            if (mobjParManager.Configurations.Count > 0)
            {
                mopAcousticStatus = 1;
                mgvdLastBKConnectResponseTime = Now;
                mdStartTime = Now;
                mobjCommands.Insert(ocInitialize, "Initialize");
            }
            else
            {
                mobjCommands.Insert(ocResetCommunications, "ResetAcousticCommunications");
            }

            tmrRunOnce = null;
            tmrElapsed.Reset();
        }
        catch (Exception ex)
        {
            sRaiseErrorText = ex.Message;
            EH.Module = TypeName(this);
            EH.Procedure = PROC_NAME;
            EH_MsgBox();
            LogMessage(string.Concat(string.Concat("Error during controller startup: ", sRaiseErrorText), ". See log file for details"), mtFailure);
        // Resume Next
        Error_CreatingObjects:
            ;
            sRaiseErrorText = string.Concat(string.Concat(string.Concat("Can't create object '", sClassName), "': "), ex.Message);
            bRaiseError = true;
            EH.Module = TypeName(this);
            EH.Procedure = PROC_NAME;
            EH.Item(1) = sClassName;
            EH_MsgBox();
        Error_Configurations:
            ;
            sRaiseErrorText = string.Concat("Unable to connect to Configurations collection: ", ex.Message);
            EH.AppError = AE.CantCreateConfigurations;
            EH.Module = TypeName(this);
            EH.Procedure = PROC_NAME;
            EH_MsgBox();
        }

    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        if (bRaiseError)
        {
            EH_RaiseError(OLEERR_METHOD_FAILED, TypeName(this), PROC_NAME, sRaiseErrorText)();
        }

        return;
    }

    private void Suspend()
    {
        Attribute(Suspend.VB_Description == "Suspend control");
        try
        {
            bInited = false;
            mobjParManager.ReleaseOutputs();
            mobjStatuses(STS_ACOUSTIC_INITIALIZED).SetValue(bInited);
            if (!mobjWinHTTPRequest == null)
            {
                mobjWinHTTPRequest.Abort();
                mobjWinHTTPRequest = null;
            }

            return;
        }
        catch
        {
            // On Error Resume Next: exceptions suppressed
        }
    }

    private void tmrRunOnce_Timer(int Milliseconds)
    {
        const object PROC_NAME = "tmrRunOnce_Timer";
        try
        {
            tmrRunOnce.Enabled = false;
            tmrRunOnce = null;
            Start();
        }
        catch (Exception ex)
        {
            if (ex.HResult & vbObjectError != vbObjectError)
            {
                EH.Module = TypeName(this);
                EH.Procedure = PROC_NAME;
                EH_MsgBox();
            }
        }

    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        return;
    }

    private string URLEncode(ref string EncodeStr)
    {
        string __result = "";
        short i = 0;
        string erg = "";
        erg = EncodeStr;
        erg = erg.Replace("%", (char)1);
        erg = erg.Replace("+", (char)2);
        for (i = 0; i <= 255; i++)
        {
            switch (i)
            {
                case 37:
                case 43:
                default:
                default:
                default:
                    break;
                case 1:
                    erg = erg.Replace((char)i, "%25");
                    break;
                case 2:
                    erg = erg.Replace((char)i, "%2B");
                    break;
                case 32:
                    erg = erg.Replace((char)i, "+");
                    break;
                default:
                    erg = erg.Replace((char)i, string.Concat("%0", Convert.ToString(i, 16)));
                    break;
                default:
                    erg = erg.Replace((char)i, string.Concat("%", Convert.ToString(i, 16)));
                    break;
            }
        }

        __result = erg;
        return __result;
    }

    private void WriteOutputs()
    {
        const object PROC_NAME = "WriteOutputs";
        bool bRaiseError = false;
        string sRaiseErrorText = "";
        ManagedPar objPar = 0;
        try
        {
            bRaiseError = false;
            if (mopAcousticStatus == 1)
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
        }
        catch (Exception ex)
        {
            sRaiseErrorText = ex.Message;
            bRaiseError = true;
        }

    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        if (bRaiseError)
        {
            EH_RaiseError(OLEERR_METHOD_FAILED, TypeName(this), PROC_NAME, sRaiseErrorText);
        }

        return;
    }
}