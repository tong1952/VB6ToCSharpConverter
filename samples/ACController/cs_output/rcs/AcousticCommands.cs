// Converted from AcousticCommands.cls by vb6cs
// Date: 2026-05-11 03:41

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
    public void StartAcousticRecord()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        mobjController.Interlocks(ILCK_GENERAL).Check();
        mobjController.Commands.Clear(PROC_NAME, mobjController.GetInvalidOpcodes(ocStartRecord.ToString()));
        mobjController.Commands.Append(ocStartRecord, PROC_NAME);
        goto Done;
        Error();
        bRaiseError = true;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, PROC_NAME, ex.Message);
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
            throw new Exception(sRaiseErrorText);
        }

        return;
    }

    public void SetAcousticTrigger()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        mobjController.Interlocks(ILCK_GENERAL).Check();
        mobjController.Commands.Clear(PROC_NAME, mobjController.GetInvalidOpcodes(ocSetTrigger.ToString()));
        mobjController.Commands.Append(ocSetTrigger, PROC_NAME);
        goto Done;
        Error();
        bRaiseError = true;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, PROC_NAME, ex.Message);
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
            throw new Exception(sRaiseErrorText);
        }

        return;
    }

    public void StopAcousticRecord()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        mobjController.Interlocks(ILCK_GENERAL).Check();
        mobjController.Commands.Clear(PROC_NAME, mobjController.GetInvalidOpcodes(ocCancelRecord.ToString()));
        mobjController.Commands.Append(ocCancelRecord, PROC_NAME);
        goto Done;
        Error();
        bRaiseError = true;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, PROC_NAME, ex.Message);
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
            throw new Exception(sRaiseErrorText);
        }

        return;
    }

    public void SendAcousticParameters()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        mobjController.CheckInterlock(ILCK_GENERAL);
        mobjController.Commands.Clear(PROC_NAME, mobjController.GetInvalidOpcodes(ocSendPars.ToString()));
        mobjController.Commands.Append(ocSendPars, PROC_NAME);
        goto Done;
        Error();
        bRaiseError = true;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, PROC_NAME, ex.Message);
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
            throw new Exception(sRaiseErrorText);
        }

        return;
    }

    public void TakeAcousticControl()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        mobjController.CheckInterlock(ILCK_GENERAL);
        mobjController.Commands.Clear(PROC_NAME, mobjController.GetInvalidOpcodes(ocTakeControl.ToString()));
        mobjController.Commands.Append(ocTakeControl, PROC_NAME);
        goto Done;
        Error();
        bRaiseError = true;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, PROC_NAME, ex.Message);
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
            throw new Exception(sRaiseErrorText);
        }

        return;
    }

    public void ReleaseAcousticControl()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        mobjController.CheckInterlock(ILCK_GENERAL);
        mobjController.Commands.Clear(PROC_NAME, mobjController.GetInvalidOpcodes(ocReleaseControl.ToString()));
        mobjController.Commands.Append(ocReleaseControl, PROC_NAME);
        goto Done;
        Error();
        bRaiseError = true;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, PROC_NAME, ex.Message);
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
            throw new Exception(sRaiseErrorText);
        }

        return;
    }

    public void ResetAcousticCommunications()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        mobjController.Commands.Clear(PROC_NAME, mobjController.GetInvalidOpcodes(ocResetCommunications.ToString()));
        mobjController.Commands.Append(ocResetCommunications, PROC_NAME);
        goto Done;
        Error();
        bRaiseError = true;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, PROC_NAME, ex.Message);
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
            throw new Exception(sRaiseErrorText);
        }

        return;
    }

    public void StartAcousticLogging(object LogToMsgLog = null)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        argMsgLog = BooleanArgData(LogToMsgLog, true);
        mobjController.Commands.Insert(ocStartLogging, PROC_NAME, argMsgLog);
        goto Done;
        Error();
        bRaiseError = true;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, PROC_NAME, ex.Message);
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
        argMsgLog = null;
        if (bRaiseError)
        {
            throw new Exception(sRaiseErrorText);
        }

        return;
    }

    public void StopAcousticLogging()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        mobjController.Commands.Insert(ocStopLogging, PROC_NAME);
        goto Done;
        Error();
        bRaiseError = true;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, PROC_NAME, ex.Message);
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
            throw new Exception(sRaiseErrorText);
        }

        return;
    }

    public void InitializeAcoustic()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        mobjController.Commands.Append(Opcodes.ocInitialize, PROC_NAME);
        mobjStatuses.Reset(STS_GFAI_INITIALIZED, false);
        goto Done;
        Error();
        bRaiseError = true;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, PROC_NAME, ex.Message);
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
            throw new Exception(sRaiseErrorText);
        }

        return;
    }

    public void SuspendAcoustic()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        mobjController.Commands.Clear(PROC_NAME);
        mobjController.Commands.Append(ocSuspend, PROC_NAME);
        goto Done;
        Error();
        bRaiseError = true;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, PROC_NAME, ex.Message);
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
            throw new Exception(sRaiseErrorText);
        }

        return;
    }

    public bool AcousticInitialized
    {
        get
        {
            bool __result = false;
            __result = mobjStatuses.GetValue(STS_GFAI_INITIALIZED);
            return __result;
        }
    }

    private void Class_Initialize()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
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
        goto Done;
        Error();
        if (ex.HResult & vbObjectError != vbObjectError)
        {
            EH.Module = TypeName(this);
            EH.Procedure = "Class_Initialize";
            EH_MsgBox();
        }

        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        objSOM = null;
        objSPGM = null;
        return;
    }

    private float GetCC(ref CCNames CC)
    {
        float __result = 0;
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        __result = mobjConstants.PropertyByPosition(CC).Value;
        goto Done;
        Error();
        sRaiseErrorText = string.Concat(string.Concat(string.Concat(string.Concat("Can't get the value of control constant '", CC), "': "), ex.Message), ". Make sure the controller is initialized");
        bRaiseError = true;
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

    private void Class_Terminate()
    {
        mobjController = null;
        mobjConstants = null;
    }
}