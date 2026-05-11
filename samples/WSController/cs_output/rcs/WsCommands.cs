// Converted from WsCommands.cls by vb6cs
// Date: 2026-05-11 03:43

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace Converted;
public class WsCommands
{
    private Controller mobjController;
    private SharedPropertyGroup mobjConstants;
    private Statuses mobjStatuses;
    public void InitializeWindSpeed(ref object CommandObject, ref string MethodName)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        objCommand = GetController.Commands.Append(Opcodes.ocInitialize, MethodName);
        mobjStatuses.Reset(STS_WINDSPEED_INITIALIZED, false);
        goto Done;
        Error();
        if (ex.HResult == cmderrTimedout)
        {
            goto Done;
        }

        bRaiseError = true;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, ex.Message);
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
            if (objCommand.SuppressLogging == false)
            {
            }
            else
            {
                lRaiseErrorNumber = lRaiseErrorNumber | tceSuppressLogging;
            }

            throw new Exception(sRaiseErrorText);
        }

        return;
    }

    public void SuspendWindSpeed(ref object CommandsObject, ref string MethodName)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        GetController.Commands.Clear(MethodName);
        GetController.Commands.Append(ocSuspend, MethodName);
        goto Done;
        Error();
        bRaiseError = true;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, ex.Message);
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

    public void StartMainFan(ref object CommandsObject, ref string MethodName)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        GetController.CheckInterlock(ocInitialize.ToString())();
        GetController.CheckInterlock(Opcodes.ocStartMainFan.ToString());
        GetController.Commands.Clear(MethodName, GetController.InvalidOpcodes(ocStartMainFan.ToString()));
        GetController.Commands.Append(ocStartMainFan, MethodName);
        mobjStatuses.Reset(STS_MAINFAN_ON, false);
        goto Done;
        Error();
        bRaiseError = true;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, ex.Message);
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

    public void StopMainFan(ref object CommandsObject, ref string MethodName)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        GetController.CheckInterlock(ocInitialize.ToString());
        GetController.CheckInterlock(Opcodes.ocStopMainFan.ToString());
        GetController.Commands.Clear(MethodName, GetController.InvalidOpcodes(ocStopMainFan.ToString()));
        GetController.Commands.Append(ocStopMainFan, MethodName);
        mobjStatuses.Reset(STS_MAINFAN_OFF, false);
        goto Done;
        Error();
        bRaiseError = true;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, ex.Message);
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

    public void StartMainFanAux(ref object CommandsObject, ref string MethodName)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        GetController.CheckInterlock(ocInitialize.ToString());
        GetController.CheckInterlock(Opcodes.ocStartMainFanAux.ToString());
        GetController.Commands.Clear(MethodName, GetController.InvalidOpcodes(ocStartMainFanAux.ToString()));
        GetController.Commands.Append(ocStartMainFanAux, MethodName);
        mobjStatuses.Reset(STS_MAINFAN_AUX_ON, false);
        goto Done;
        Error();
        bRaiseError = true;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, ex.Message);
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

    public void StopMainFanAux(ref object CommandsObject, ref string MethodName)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        GetController.CheckInterlock(ocInitialize.ToString());
        GetController.CheckInterlock(Opcodes.ocStopMainFanAux.ToString());
        GetController.Commands.Clear(MethodName, GetController.InvalidOpcodes(ocStopMainFanAux.ToString()));
        GetController.Commands.Append(ocStopMainFanAux, MethodName);
        mobjStatuses.Reset(STS_MAINFAN_AUX_OFF, false);
        goto Done;
        Error();
        bRaiseError = true;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, ex.Message);
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

    public void ResetMainFanFaults(ref object CommandsObject, ref string MethodName)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        GetController.CheckInterlock(ocInitialize.ToString())();
        GetController.CheckInterlock(Opcodes.ocResetFaults.ToString());
        GetController.Commands.Clear(MethodName, GetController.InvalidOpcodes(ocResetFaults.ToString()));
        GetController.Commands.Append(ocResetFaults, MethodName);
        goto Done;
        Error();
        bRaiseError = true;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, ex.Message);
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

    public void StartPurge(ref object CommandsObject, ref string MethodName)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        GetController.CheckInterlock(ocStartPurge.ToString())();
        GetController.Commands.Clear(MethodName, GetController.InvalidOpcodes(ocStartPurge.ToString()));
        GetController.Commands.Append(ocStartPurge, MethodName);
        goto Done;
        Error();
        bRaiseError = true;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, ex.Message);
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

    public void StopPurge(ref object CommandsObject, ref string MethodName)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        GetController.CheckInterlock(ocStopPurge.ToString())();
        GetController.Commands.Clear(MethodName, GetController.InvalidOpcodes(ocStopPurge.ToString()));
        GetController.Commands.Append(ocStopPurge, MethodName);
        goto Done;
        Error();
        bRaiseError = true;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, ex.Message);
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

    public void StartIdle(ref object CommandsObject, ref string MethodName)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        GetController.CheckInterlock(ocInitialize.ToString());
        GetController.CheckInterlock(Opcodes.ocStartIdle.ToString());
        GetController.Commands.Clear(MethodName, GetController.InvalidOpcodes(ocStartIdle.ToString()));
        GetController.Commands.Append(ocStartIdle, MethodName);
        mobjStatuses.Reset(STS_IDLE_STARTED, false);
        goto Done;
        Error();
        bRaiseError = true;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, ex.Message);
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

    public void StopIdle(ref object CommandsObject, ref string MethodName)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        GetController.CheckInterlock(ocInitialize.ToString());
        GetController.CheckInterlock(Opcodes.ocStopIdle.ToString());
        GetController.Commands.Clear(MethodName, GetController.InvalidOpcodes(ocStopIdle.ToString()));
        GetController.Commands.Append(ocStopIdle, MethodName);
        mobjStatuses.Reset(STS_IDLE_STOPPED, false);
        goto Done;
        Error();
        bRaiseError = true;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, ex.Message);
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

    public void StopIdleMotion(ref object CommandsObject, ref string MethodName)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        GetController.CheckInterlock(ocInitialize.ToString());
        GetController.CheckInterlock(Opcodes.ocCancelIdle.ToString());
        GetController.Commands.Clear(MethodName, GetController.InvalidOpcodes(ocCancelIdle.ToString()));
        GetController.Commands.Insert(ocCancelIdle, MethodName);
        goto Done;
        Error();
        bRaiseError = true;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, ex.Message);
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

    public void SetNozzleSize(ref object CommandsObject, ref string MethodName, ref object NozzleSize)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        argSize = StringArgData(NozzleSize);
        argSize.AddValidValues("Small", "Medium", "Large");
        CheckMethodArgs(CommandsObject, MethodName, argSize);
        if (argSize.Value == "Large")
        {
            lNozzleSize = nsLarge;
        }
        else if (argSize.Value == "Medium")
        {
            lNozzleSize = nsMedium;
        }
        else
        {
            lNozzleSize = nsSmall;
        }

        GetController.CheckInterlock(ocInitialize.ToString());
        GetController.CheckInterlock(Opcodes.ocSetNozzleSize.ToString());
        GetController.Commands.Clear(MethodName, GetController.InvalidOpcodes(ocSetNozzleSize.ToString()));
        GetController.Commands.Append(ocSetNozzleSize, MethodName, lNozzleSize);
        goto Done;
        Error();
        bRaiseError = true;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, ex.Message);
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
        argSize = null;
        if (bRaiseError)
        {
            throw new Exception(sRaiseErrorText);
        }

        return;
    }

    public void StopNozzleMotion(ref object CommandsObject, ref string MethodName)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        GetController.CheckInterlock(ocInitialize.ToString());
        GetController.CheckInterlock(Opcodes.ocStopNozzleMotion.ToString());
        GetController.Commands.Clear(MethodName, GetController.InvalidOpcodes(ocStopNozzleMotion.ToString()));
        GetController.Commands.Insert(ocStopNozzleMotion, MethodName);
        goto Done;
        Error();
        bRaiseError = true;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, ex.Message);
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

    public void SetWindSpeed(ref object CommandsObject, ref string MethodName, ref object Target, object RampRate = null)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        GetController.CheckInterlock(ocInitialize.ToString());
        argTarget = ValueArgData(Target, GetCC(CC_WSMin), GetCC(CC_WSMax));
        argRampRate = ValueArgData(RampRate, null, GetCC(CC_WSRampMax), GetCC(CC_WSRampDef));
        CheckMethodArgs(CommandsObject, MethodName, argTarget, argRampRate);
        GetController.CheckInterlock(Array(ocSetWindSpeed.ToString(), argTarget.Value));
        GetController.Commands.Clear(MethodName, GetController.InvalidOpcodes(ocSetWindSpeed.ToString()));
        GetController.Commands.Append(ocSetWindSpeed, MethodName, argTarget, argRampRate);
        mobjStatuses.Reset(STS_WINDSPEED_ACHIEVED, false);
        goto Done;
        Error();
        bRaiseError = true;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, ex.Message);
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
        argRampRate = null;
        argTarget = null;
        if (bRaiseError)
        {
            throw new Exception(sRaiseErrorText);
        }

        return;
    }

    public void SetFanSpeed(ref object CommandsObject, ref string MethodName, ref object Target, object RampRate = null)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        GetController.CheckInterlock(ocInitialize.ToString());
        argTarget = ValueArgData(Target, GetCC(CC_FSMin), GetCC(CC_FSMax));
        argRampRate = ValueArgData(RampRate, null, GetCC(CC_FSRampMax), GetCC(CC_FSRampDef));
        CheckMethodArgs(CommandsObject, MethodName, argTarget, argRampRate);
        GetController.CheckInterlock(Array(ocSetFanSpeed.ToString(), argTarget.Value));
        GetController.Commands.Clear(MethodName, GetController.InvalidOpcodes(ocSetFanSpeed.ToString()));
        GetController.Commands.Append(ocSetFanSpeed, MethodName, argTarget, argRampRate);
        mobjStatuses.Reset(STS_FANSPEED_ACHIEVED, false);
        goto Done;
        Error();
        bRaiseError = true;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, ex.Message);
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
        argRampRate = null;
        argTarget = null;
        if (bRaiseError)
        {
            throw new Exception(sRaiseErrorText);
        }

        return;
    }

    public void TrackWindSpeed(ref object CommandsObject, ref string MethodName, ref object TrackPar, object AutoIdle = null, object Headwind = null)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        GetController.CheckInterlock(ocInitialize.ToString());
        if (WindTunnel == wtCWT)
        {
            argAutoIdle = BooleanArgData(AutoIdle, GetCC(CC_DefAutoIdle));
        }

        argTrackPar = StringArgData(TrackPar, GetCCString(CC_DefTrackParName));
        if (argTrackPar.Value == "")
        {
            argTrackPar.Value = GetCCString(CC_DefTrackParName);
        }

        argHeadWind = ValueArgData(Headwind, GetCC(CC_HeadwindMin), GetCC(CC_HeadwindMax), 0);
        if (WindTunnel == wtCWT)
        {
            CheckMethodArgs(CommandsObject, MethodName, argTrackPar, argAutoIdle, argHeadWind);
        }
        else
        {
            CheckMethodArgs(CommandsObject, MethodName, argTrackPar, argHeadWind);
        }

        objPar = GetController.GetParameter(argTrackPar.Value);
        if (!IsCompatibleUnits(objPar.Units, "kph"))
        {
            throw new Exception(tceInvalidArgs);
            LoadResString(ridTrackParInvalidUnits, objPar.Name, objPar.Units, "kph")();
        }

        GetController.CheckInterlock(ocTrackWindSpeed.ToString());
        GetController.Commands.Clear(MethodName, GetController.InvalidOpcodes(ocTrackWindSpeed.ToString()));
        GetController.Commands.Append(ocTrackWindSpeed, MethodName, argTrackPar, argHeadWind, argAutoIdle);
        goto Done;
        Error();
        bRaiseError = true;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, ex.Message);
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
        argTrackPar = null;
        argHeadWind = null;
        argAutoIdle = null;
        objPar = null;
        if (bRaiseError)
        {
            throw new Exception(sRaiseErrorText);
        }

        return;
    }

    public void SetCFlapLeft1(ref object CommandsObject, ref string MethodName, ref object Target)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        argTarget = ValueArgData(Target, GetCC(CC_CFLeft1Min), GetCC(CC_CFLeft1Max));
        CheckMethodArgs(CommandsObject, MethodName, argTarget);
        GetController.CheckInterlock(ocInitialize.ToString());
        GetController.CheckInterlock(Opcodes.ocSetCFlapLeft1.ToString());
        GetController.Commands.Clear(MethodName, GetController.InvalidOpcodes(ocSetCFlapLeft1.ToString()));
        GetController.Commands.Append(ocSetCFlapLeft1, MethodName, argTarget);
        goto Done;
        Error();
        bRaiseError = true;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, ex.Message);
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
        argTarget = null;
        if (bRaiseError)
        {
            throw new Exception(sRaiseErrorText);
        }

        return;
    }

    public void SetCFlapLeft2(ref object CommandsObject, ref string MethodName, ref object Target)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        argTarget = ValueArgData(Target, GetCC(CC_CFLeft2Min), GetCC(CC_CFLeft2Max));
        CheckMethodArgs(CommandsObject, MethodName, argTarget);
        GetController.CheckInterlock(ocInitialize.ToString());
        GetController.CheckInterlock(Opcodes.ocSetCFlapLeft2.ToString());
        GetController.Commands.Clear(MethodName, GetController.InvalidOpcodes(ocSetCFlapLeft2.ToString()));
        GetController.Commands.Append(ocSetCFlapLeft2, MethodName, argTarget);
        goto Done;
        Error();
        bRaiseError = true;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, ex.Message);
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
        argTarget = null;
        if (bRaiseError)
        {
            throw new Exception(sRaiseErrorText);
        }

        return;
    }

    public void SetCFlapRight1(ref object CommandsObject, ref string MethodName, ref object Target)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        argTarget = ValueArgData(Target, GetCC(CC_CFRight1Min), GetCC(CC_CFRight1Max));
        CheckMethodArgs(CommandsObject, MethodName, argTarget);
        GetController.CheckInterlock(ocInitialize.ToString());
        GetController.CheckInterlock(Opcodes.ocSetCFlapRight1.ToString());
        GetController.Commands.Clear(MethodName, GetController.InvalidOpcodes(ocSetCFlapRight1.ToString()));
        GetController.Commands.Append(ocSetCFlapRight1, MethodName, argTarget);
        goto Done;
        Error();
        bRaiseError = true;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, ex.Message);
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
        argTarget = null;
        if (bRaiseError)
        {
            throw new Exception(sRaiseErrorText);
        }

        return;
    }

    public void SetCFlapRight2(ref object CommandsObject, ref string MethodName, ref object Target)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        argTarget = ValueArgData(Target, GetCC(CC_CFRight2Min), GetCC(CC_CFRight2Max));
        CheckMethodArgs(CommandsObject, MethodName, argTarget);
        GetController.CheckInterlock(ocInitialize.ToString());
        GetController.CheckInterlock(Opcodes.ocSetCFlapRight2.ToString());
        GetController.Commands.Clear(MethodName, GetController.InvalidOpcodes(ocSetCFlapRight2.ToString()));
        GetController.Commands.Append(ocSetCFlapRight2, MethodName, argTarget);
        goto Done;
        Error();
        bRaiseError = true;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, ex.Message);
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
        argTarget = null;
        if (bRaiseError)
        {
            throw new Exception(sRaiseErrorText);
        }

        return;
    }

    public void SetCFlapTop(ref object CommandsObject, ref string MethodName, ref object Target)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        argTarget = ValueArgData(Target, GetCC(CC_CFTopMin), GetCC(CC_CFTopMax));
        CheckMethodArgs(CommandsObject, MethodName, argTarget);
        GetController.CheckInterlock(ocInitialize.ToString());
        GetController.CheckInterlock(Opcodes.ocSetCFlapTop.ToString());
        GetController.Commands.Clear(MethodName, GetController.InvalidOpcodes(ocSetCFlapTop.ToString()));
        GetController.Commands.Append(ocSetCFlapTop, MethodName, argTarget);
        goto Done;
        Error();
        bRaiseError = true;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, ex.Message);
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
        argTarget = null;
        if (bRaiseError)
        {
            throw new Exception(sRaiseErrorText);
        }

        return;
    }

    public void StopCFlap(ref object CommandsObject, ref string MethodName)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        GetController.CheckInterlock(ocInitialize.ToString());
        GetController.CheckInterlock(Opcodes.ocStopCF.ToString());
        GetController.Commands.Clear(MethodName, GetController.InvalidOpcodes(ocStopCF.ToString()));
        GetController.Commands.Insert(ocStopCF, MethodName);
        goto Done;
        Error();
        bRaiseError = true;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, ex.Message);
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

    public void ShowWindspeedStatus(ref object CommandObject, ref string MethodName, object Sender = null)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        // On Error GoTo Error_Interlocks — nested handler not restructured
        GetController.CheckInterlock(ocInitialize.ToString());
        GetController.CheckInterlock(ocShowStatus.ToString());
        // On Error GoTo Error — nested handler not restructured
        sMessage = LoadResString(stsNoFaultsDetected);
        LogMessage(sMessage, mtInformation);
        switch (true)
        {
            case IsMissing(Sender):
            case Sender == null:
                break;
            default:
                // On Error GoTo Error_Message — nested handler not restructured
                ;
                Sender.ShowMessageDialog(sMessage, guiInformation, false, PROCESS_DESCRIPTION, 10);
                break;
        }

        goto Done;
    Error_Interlocks:
        ;
        sMessage = ex.Message;
        goto Interlocked;
    Interlocked:
        ;
        // On Error GoTo Error — nested handler not restructured
        objEH = new EH();
        objEH.AppError = LoadResString(stsEHTheFollowingFaultsExist, sMessage);
        objEH.Component = App;
        LogMessage2(objEH);
        switch (true)
        {
            case IsMissing(Sender):
            case Sender == null:
                break;
            default:
                sMessage = LoadResString(stsTheFollowingFaultsExist, sMessage);
                // On Error GoTo Error_Message — nested handler not restructured
                ;
                Sender.ShowMessageDialog(sMessage, guiWarning, false, PROCESS_DESCRIPTION, 30);
                break;
        }

        goto Done;
        Error();
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, PROC_NAME, ex.Message);
        bRaiseError = true;
        if (ex.HResult & vbObjectError != vbObjectError)
        {
            EH.Module = TypeName(this);
            EH.Procedure = PROC_NAME;
            EH_MsgBox();
        }

        goto Done;
    Error_Message:
        ;
        lRaiseErrorNumber = tceUnexpectedError;
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, PROC_NAME, LoadResString(ridErrDisplayingMessage, ex.Message));
        bRaiseError = true;
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

    public void UseNozzleMethod(ref object CommandObject, ref string MethodName)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        GetController.CheckInterlock(ocInitialize.ToString());
        GetController.Commands.Append(ocSetNozzleMethod, MethodName);
        goto Done;
        Error();
        bRaiseError = true;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, ex.Message);
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

    public void UsePlenumMethod(ref object CommandObject, ref string MethodName)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        GetController.CheckInterlock(ocInitialize.ToString());
        GetController.Commands.Append(ocSetPlenumMethod, MethodName);
        goto Done;
        Error();
        bRaiseError = true;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, ex.Message);
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

    public void ShowIdleStatus(ref object CommandObject, ref string MethodName, object Sender = null)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        // On Error GoTo Error_Interlocks — nested handler not restructured
        GetController.CheckInterlock(ocInitialize.ToString());
        GetController.CheckInterlock(ocShowIdleStatus.ToString());
        // On Error GoTo Error — nested handler not restructured
        sMessage = LoadResString(stsNoFaultsDetected);
        LogMessage(sMessage, mtInformation);
        switch (true)
        {
            case IsMissing(Sender):
            case Sender == null:
                break;
            default:
                // On Error GoTo Error_Message — nested handler not restructured
                ;
                Sender.ShowMessageDialog(sMessage, guiInformation, false, PROCESS_DESCRIPTION, 10);
                break;
        }

        goto Done;
    Error_Interlocks:
        ;
        sMessage = ex.Message;
        goto Interlocked;
    Interlocked:
        ;
        // On Error GoTo Error — nested handler not restructured
        objEH = new EH();
        objEH.AppError = LoadResString(stsEHTheFollowingFaultsExist, sMessage);
        objEH.Component = App;
        LogMessage2(objEH);
        switch (true)
        {
            case IsMissing(Sender):
            case Sender == null:
                break;
            default:
                sMessage = LoadResString(stsTheFollowingFaultsExist, sMessage);
                // On Error GoTo Error_Message — nested handler not restructured
                ;
                Sender.ShowMessageDialog(sMessage, guiWarning, false, PROCESS_DESCRIPTION, 30);
                break;
        }

        goto Done;
        Error();
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, PROC_NAME, ex.Message);
        bRaiseError = true;
        if (ex.HResult & vbObjectError != vbObjectError)
        {
            EH.Module = TypeName(this);
            EH.Procedure = PROC_NAME;
            EH_MsgBox();
        }

        goto Done;
    Error_Message:
        ;
        lRaiseErrorNumber = tceUnexpectedError;
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, PROC_NAME, LoadResString(ridErrDisplayingMessage, ex.Message));
        bRaiseError = true;
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

    internal bool GetStatus(ref string StatusName)
    {
        bool __result = false;
        __result = mobjStatuses.GetValue(StatusName);
        return __result;
    }

    private Controller GetController
    {
        get
        {
            Controller __result = 0;
            // TODO: On Error GoTo Error — handler label not found in this scope
            if (mobjController == null)
            {
                objSOM = new SOM();
                mobjController = objSOM.GetObject(SOM_CONTROLLER);
            }

            __result = mobjController;
            goto Done;
            Error();
            EH.Module = TypeName(this);
            EH.Procedure = "GetController";
            EH_MsgBox();
            goto Done;
        Done:
            ;
            // On Error Resume Next — nested handler; see enclosing try/catch
            objSOM = null;
            return __result;
        }
    }

    private void Class_Initialize()
    {
        mobjStatuses = new Statuses();
        mobjStatuses.UseSPG(SPG_STATUSES);
    }

    private void Class_Terminate()
    {
        mobjController = null;
        mobjConstants = null;
        mobjStatuses = null;
    }

    private float GetCC(ref SPGConstants CC)
    {
        float __result = 0;
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        if (mobjConstants == null)
        {
            objSPGM = new SharedPropertyGroupManager();
            mobjConstants = objSPGM.CreatePropertyGroup(SPG_CONSTANTS, 0, 0, bExists);
        }

        __result = mobjConstants.PropertyByPosition(CC).Value;
        goto Done;
        Error();
        sRaiseErrorText = LoadResString(stsCantGetCC, CC, ex.Message);
        bRaiseError = true;
        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        objSPGM = null;
        if (bRaiseError)
        {
            throw new Exception(sRaiseErrorText);
        }

        return __result;
    }

    private string GetCCString(ref SPGConstants CC)
    {
        string __result = "";
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        if (mobjConstants == null)
        {
            objSPGM = new SharedPropertyGroupManager();
            mobjConstants = objSPGM.CreatePropertyGroup(SPG_CONSTANTS, 0, 0, bExists);
        }

        __result = mobjConstants.PropertyByPosition(CC).Value;
        goto Done;
        Error();
        sRaiseErrorText = LoadResString(stsCantGetCC, CC, ex.Message);
        bRaiseError = true;
        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        objSPGM = null;
        if (bRaiseError)
        {
            throw new Exception(sRaiseErrorText);
        }

        return __result;
    }

    public void SetBlockage(ref object CommandsObject, ref string MethodName, ref object Blockage)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        if (Blockage == BLOCKAGE_1)
        {
            fBlockage = USE_BLOCKAGE1;
        }
        else if (Blockage == BLOCKAGE_2)
        {
            fBlockage = USE_BLOCKAGE2;
        }
        else
        {
            argBlockage = ValueArgData(Blockage, GetCC(CC_MinBlockage), GetCC(CC_MaxBlockage));
            CheckMethodArgs(CommandsObject, MethodName, argBlockage);
            fBlockage = argBlockage.Value;
        }

        GetController.CheckInterlock(ocInitialize.ToString());
        GetController.CheckInterlock(ocSetBlockage.ToString());
        GetController.Commands.Clear(MethodName, GetController.InvalidOpcodes(ocSetBlockage.ToString()));
        GetController.Commands.Append(ocSetBlockage, MethodName, fBlockage);
        goto Done;
        Error();
        bRaiseError = true;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, ex.Message);
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
        argBlockage = null;
        if (bRaiseError)
        {
            throw new Exception(sRaiseErrorText);
        }

        return;
    }

    public void ResetIdle(ref object CommandsObject, ref string MethodName)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        GetController.CheckInterlock(ocInitialize.ToString());
        GetController.CheckInterlock(Opcodes.ocResetIdle.ToString());
        GetController.Commands.Clear(MethodName, GetController.InvalidOpcodes(ocResetIdle.ToString()));
        GetController.Commands.Append(ocResetIdle, MethodName);
        goto Done;
        Error();
        bRaiseError = true;
        lRaiseErrorNumber = CCommandError(ex.HResult);
        sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, ex.Message);
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
}