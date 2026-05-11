// Converted from AcousticCommands.cls by vb6cs
// Date: 2026-05-11 05:50

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace Converted;
public class AcousticCommands
{
    private Controller mobjController;
    private SharedPropertyGroup mobjConstants;
    private Statuses mobjStatuses;
    public bool AcousticInitialized
    {
        get
        {
            bool __result = false;
            __result = mobjStatuses.GetValue(STS_ACOUSTIC_INITIALIZED);
            return __result;
        }
    }

    public bool AcousticFaulted
    {
        get
        {
            bool __result = false;
            __result = mobjStatuses.GetValue(STS_ACOUSTIC_FAULTED);
            return __result;
        }
    }

    private void Class_Initialize()
    {
        const object PROC_NAME = "Class_Initialize";
        SOM objSOM = 0;
        SharedPropertyGroupManager objSPGM = 0;
        bool bExists = false;
        try
        {
            objSOM = new SOM();
            // On Error Resume Next — nested handler; see enclosing try/catch
            mobjController = objSOM.GetObject(SOM_CONTROLLER);
            if (Err)
            {
                // On Error GoTo Error — nested handler not restructured
                mobjController = new Controller();
                objSOM.AddObject(mobjController, SOM_CONTROLLER);
            }// On Error GoTo Error — nested handler not restructured
            objSPGM = new SharedPropertyGroupManager();
            mobjConstants = objSPGM.CreatePropertyGroup(SPG_CONSTANTS, 0, 0, bExists);
            mobjStatuses = new Statuses();
            mobjStatuses.UseSPG(SPG_STATUSES);
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
        objSOM = null;
        objSPGM = null;
        return;
    }

    private void Class_Terminate()
    {
        mobjController = null;
        mobjConstants = null;
    }

    private float GetCC(ref CCNames CC)
    {
        float __result = 0;
        Attribute(GetCC.VB_Description == "Returns a specific control constant");
        const object PROC_NAME = "GetCC";
        bool bRaiseError = false;
        string sRaiseErrorText = "";
        try
        {
            bRaiseError = false;
            __result = mobjConstants.PropertyByPosition(CC).Value;
        }
        catch (Exception ex)
        {
            sRaiseErrorText = string.Concat(string.Concat(string.Concat(string.Concat("Can't get the value of control constant '", CC), "': "), ex.Message), ". Make sure the controller is initialized");
            bRaiseError = true;
        }

    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        if (bRaiseError)
        {
            throw new Exception(sRaiseErrorText);
        }

        return __result;
    }

    public void InitializeAcoustic()
    {
        Attribute(InitializeAcoustic.VB_Description == "Initialize the program");
        const object PROC_NAME = "InitializeAcoustic";
        bool bRaiseError = false;
        string sRaiseErrorText = "";
        int lRaiseErrorNumber = 0;
        Command objCommand = 0;
        try
        {
            bRaiseError = false;
            mobjController.Commands.Append(Opcodes.ocInitialize, PROC_NAME);
        }
        catch (Exception ex)
        {
            bRaiseError = true;
            lRaiseErrorNumber = CCommandError(ex.HResult);
            sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, PROC_NAME, ex.Message);
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
            throw new Exception(sRaiseErrorText);
        }

        return;
    }

    public void ResetAcousticCommunications()
    {
        Attribute(ResetAcousticCommunications.VB_Description == "Reset communications with BK Connect");
        const object PROC_NAME = "ResetAcousticCommunications";
        bool bRaiseError = false;
        string sRaiseErrorText = "";
        int lRaiseErrorNumber = 0;
        try
        {
            bRaiseError = false;
            mobjController.Commands.Clear(PROC_NAME, mobjController.GetInvalidOpcodes(ocResetCommunications.ToString()));
            mobjController.Commands.Append(ocResetCommunications, PROC_NAME);
        }
        catch (Exception ex)
        {
            bRaiseError = true;
            lRaiseErrorNumber = CCommandError(ex.HResult);
            sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, PROC_NAME, ex.Message);
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
            throw new Exception(sRaiseErrorText);
        }

        return;
    }

    public void SendAcousticCommand(ref object API)
    {
        Attribute(SendAcousticCommand.VB_Description == "Send a BK Connect API (URL) representing a command the web interface understands");
        Attribute(SendAcousticCommand.VB_MemberFlags == "40");
        const object PROC_NAME = "SendAcousticCommand";
        bool bRaiseError = false;
        string sRaiseErrorText = "";
        int lRaiseErrorNumber = 0;
        ArgData oArg = 0;
        try
        {
            bRaiseError = false;
            oArg = StringArgData(API);
            CheckMethodArgs(this, PROC_NAME, oArg);
            mobjController.CheckInterlock(ILCK_GENERAL);
            mobjController.Commands.Clear(PROC_NAME, mobjController.GetInvalidOpcodes(ocSendCommand.ToString()));
            mobjController.Commands.Append(ocSendCommand, PROC_NAME, oArg);
        }
        catch (Exception ex)
        {
            bRaiseError = true;
            lRaiseErrorNumber = CCommandError(ex.HResult);
            sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, PROC_NAME, ex.Message);
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
            throw new Exception(sRaiseErrorText);
        }

        return;
    }

    public void SetAcousticIP(ref string IPaddr)
    {
        Attribute(SetAcousticIP.VB_Description == "Set the target IP address of computer running BK Connect");
        const object PROC_NAME = "SetAcousticIP";
        bool bRaiseError = false;
        string sRaiseErrorText = "";
        int lRaiseErrorNumber = 0;
        ArgData oArg = 0;
        try
        {
            bRaiseError = false;
            oArg = StringArgData(IPaddr);
            CheckMethodArgs(this, PROC_NAME, oArg);
            mobjController.CheckInterlock(ILCK_GENERAL);
            mobjController.Commands.Clear(PROC_NAME, mobjController.GetInvalidOpcodes(ocSetIPAddr.ToString()));
            mobjController.Commands.Append(ocSetIPAddr, PROC_NAME, oArg);
        }
        catch (Exception ex)
        {
            bRaiseError = true;
            lRaiseErrorNumber = CCommandError(ex.HResult);
            sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, PROC_NAME, ex.Message);
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
            throw new Exception(sRaiseErrorText);
        }

        return;
    }

    public void StartAcousticLogging(object LogToMsgLog = null)
    {
        Attribute(StartAcousticLogging.VB_Description == "Start logging communications to the diagnostic panel");
        const object PROC_NAME = "StartAcousticLogging";
        bool bRaiseError = false;
        string sRaiseErrorText = "";
        int lRaiseErrorNumber = 0;
        ArgData argMsgLog = 0;
        try
        {
            bRaiseError = false;
            argMsgLog = BooleanArgData(LogToMsgLog, true);
            mobjController.Commands.Insert(ocStartLogging, PROC_NAME, argMsgLog);
        }
        catch (Exception ex)
        {
            bRaiseError = true;
            lRaiseErrorNumber = CCommandError(ex.HResult);
            sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, PROC_NAME, ex.Message);
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
        argMsgLog = null;
        if (bRaiseError)
        {
            throw new Exception(sRaiseErrorText);
        }

        return;
    }

    public void StartAcousticRecorder()
    {
        Attribute(StartAcousticRecorder.VB_Description == "Start the BK Connect recorder. The Data Recorder Service must have been previously started.");
        const object PROC_NAME = "StartAcousticRecorder";
        bool bRaiseError = false;
        string sRaiseErrorText = "";
        int lRaiseErrorNumber = 0;
        try
        {
            bRaiseError = false;
            mobjController.Interlocks(ILCK_GENERAL).Check();
            mobjController.Commands.Clear(PROC_NAME, mobjController.GetInvalidOpcodes(ocStartRecord.ToString()));
            mobjController.Commands.Append(ocStartRecord, PROC_NAME);
        }
        catch (Exception ex)
        {
            bRaiseError = true;
            lRaiseErrorNumber = CCommandError(ex.HResult);
            sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, PROC_NAME, ex.Message);
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
            throw new Exception(sRaiseErrorText);
        }

        return;
    }

    public void StopAcousticLogging()
    {
        Attribute(StopAcousticLogging.VB_Description == "Stop logging communications to the diagnostic panel");
        const object PROC_NAME = "StopAcousticLogging";
        bool bRaiseError = false;
        string sRaiseErrorText = "";
        int lRaiseErrorNumber = 0;
        try
        {
            bRaiseError = false;
            mobjController.Commands.Insert(ocStopLogging, PROC_NAME);
        }
        catch (Exception ex)
        {
            bRaiseError = true;
            lRaiseErrorNumber = CCommandError(ex.HResult);
            sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, PROC_NAME, ex.Message);
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
            throw new Exception(sRaiseErrorText);
        }

        return;
    }

    public void StopAcousticRecorder()
    {
        Attribute(StopAcousticRecorder.VB_Description == "Stop the BK Connect recorder. The Data Recorder Service must have been previously started.");
        const object PROC_NAME = "StopAcousticRecorder";
        bool bRaiseError = false;
        string sRaiseErrorText = "";
        int lRaiseErrorNumber = 0;
        try
        {
            bRaiseError = false;
            mobjController.Interlocks(ILCK_GENERAL).Check();
            mobjController.Commands.Clear(PROC_NAME, mobjController.GetInvalidOpcodes(ocStopRecord.ToString()));
            mobjController.Commands.Append(ocStopRecord, PROC_NAME);
        }
        catch (Exception ex)
        {
            bRaiseError = true;
            lRaiseErrorNumber = CCommandError(ex.HResult);
            sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, PROC_NAME, ex.Message);
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
            throw new Exception(sRaiseErrorText);
        }

        return;
    }

    public void SuspendAcoustic()
    {
        Attribute(SuspendAcoustic.VB_Description == "Suspend the control program");
        const object PROC_NAME = "SuspendAcoustic";
        bool bRaiseError = false;
        string sRaiseErrorText = "";
        int lRaiseErrorNumber = 0;
        try
        {
            bRaiseError = false;
            mobjController.Commands.Clear(PROC_NAME);
            mobjController.Commands.Append(ocSuspend, PROC_NAME);
        }
        catch (Exception ex)
        {
            bRaiseError = true;
            lRaiseErrorNumber = CCommandError(ex.HResult);
            sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, PROC_NAME, ex.Message);
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
            throw new Exception(sRaiseErrorText);
        }

        return;
    }
}