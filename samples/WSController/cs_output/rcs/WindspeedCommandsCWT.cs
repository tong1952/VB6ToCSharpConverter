// Converted from WindspeedCommandsCWT.cls by vb6cs
// Date: 2026-05-11 03:43

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace Converted;
public class WindspeedCommandsCWT : TalentCommandClassEx1
{
    public WsCommands mobjCommands;
    public void SetWindSpeed(ref object Target, object RampRate = null)
    {
        mobjCommands.SetWindSpeed(this, "SetWindSpeed", Target, RampRate);
    }

    public void TrackWindSpeed(ref object TrackParameter, object AutoIdle = null, object Headwind = null)
    {
        mobjCommands.TrackWindSpeed(this, "TrackWindSpeed", TrackParameter, AutoIdle, Headwind);
    }

    public void StartMainFanAux()
    {
        mobjCommands.StartMainFanAux(this, "StartMainFanAux");
    }

    public void StopMainFanAux()
    {
        mobjCommands.StopMainFanAux(this, "StopMainFanAux");
    }

    public void StartMainFan()
    {
        mobjCommands.StartMainFan(this, "StartMainFan");
    }

    public void StopMainFan()
    {
        mobjCommands.StopMainFan(this, "StopMainFan");
    }

    public void ResetMainFanFaults()
    {
        mobjCommands.ResetMainFanFaults(this, "ResetMainFanFaults");
    }

    public void SetFanSpeed(ref object Target, object RampRate = null)
    {
        mobjCommands.SetFanSpeed(this, "SetFanSpeed", Target, RampRate);
    }

    public void StartPurge()
    {
        mobjCommands.StartPurge(this, "StartPurge");
    }

    public void StopPurge()
    {
        mobjCommands.StopPurge(this, "StopPurge");
    }

    public void StartIdle()
    {
        mobjCommands.StartIdle(this, "StartIdle");
    }

    public void StopIdle()
    {
        mobjCommands.StopIdle(this, "StopIdle");
    }

    public void StopIdleMotion()
    {
        mobjCommands.StopIdleMotion(this, "StopIdleMotion");
    }

    public void SetNozzleSize(ref object NozzleSize)
    {
        mobjCommands.SetNozzleSize(this, "SetNozzleSize", NozzleSize);
    }

    public void StopNozzleMotion()
    {
        mobjCommands.StopNozzleMotion(this, "StopNozzleMotion");
    }

    public void SetCFlapLeft1(ref object Target)
    {
        mobjCommands.SetCFlapLeft1(this, "SetCFlapLeft1", Target);
    }

    public void SetCFlapLeft2(ref object Target)
    {
        mobjCommands.SetCFlapLeft2(this, "SetCFlapLeft2", Target);
    }

    public void SetCFlapRight1(ref object Target)
    {
        mobjCommands.SetCFlapRight1(this, "SetCFlapRight1", Target);
    }

    public void SetCFlapRight2(ref object Target)
    {
        mobjCommands.SetCFlapRight2(this, "SetCFlapRight2", Target);
    }

    public void SetCFlapTop(ref object Target)
    {
        mobjCommands.SetCFlapTop(this, "SetCFlapTop", Target);
    }

    public void StopCFlap()
    {
        mobjCommands.StopCFlap(this, "StopCFlap");
    }

    public void UseNozzleMethod()
    {
        mobjCommands.UseNozzleMethod(this, "UseNozzleMethod");
    }

    public void UsePlenumMethod()
    {
        mobjCommands.UsePlenumMethod(this, "UsePlenumMethod");
    }

    public void ShowWindspeedStatus(object Sender = null)
    {
        mobjCommands.ShowWindspeedStatus(this, "ShowWindSpeedStatus", Sender);
    }

    public void InitializeWindSpeed()
    {
        mobjCommands.InitializeWindSpeed(this, "InitializeWindSpeed");
    }

    public void SuspendWindSpeed()
    {
        mobjCommands.SuspendWindSpeed(this, "SuspendWindSpeed");
    }

    public bool WindspeedAchieved
    {
        get
        {
            bool __result = false;
            __result = mobjCommands.GetStatus(STS_WINDSPEED_ACHIEVED);
            return __result;
        }
    }

    public bool FanSpeedAchieved
    {
        get
        {
            bool __result = false;
            __result = mobjCommands.GetStatus(STS_FANSPEED_ACHIEVED);
            return __result;
        }
    }

    public bool IdleStarted
    {
        get
        {
            bool __result = false;
            __result = mobjCommands.GetStatus(STS_IDLE_STARTED);
            return __result;
        }
    }

    public bool IdleStopped
    {
        get
        {
            bool __result = false;
            __result = mobjCommands.GetStatus(STS_IDLE_STOPPED);
            return __result;
        }
    }

    public bool MainFanAuxStarted
    {
        get
        {
            bool __result = false;
            __result = mobjCommands.GetStatus(STS_MAINFAN_AUX_ON);
            return __result;
        }
    }

    public bool MainFanAuxStopped
    {
        get
        {
            bool __result = false;
            __result = !mobjCommands.GetStatus(STS_MAINFAN_AUX_ON);
            return __result;
        }
    }

    public bool MainFanStarted
    {
        get
        {
            bool __result = false;
            __result = mobjCommands.GetStatus(STS_MAINFAN_ON);
            return __result;
        }
    }

    public bool MainFanStopped
    {
        get
        {
            bool __result = false;
            __result = !mobjCommands.GetStatus(STS_MAINFAN_ON);
            return __result;
        }
    }

    public bool WindspeedInitialized
    {
        get
        {
            bool __result = false;
            __result = mobjCommands.GetStatus(STS_WINDSPEED_INITIALIZED);
            return __result;
        }
    }

    private void Class_Initialize()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        mobjCommands = new WsCommands();
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
            EH_RaiseError(OLEERR_METHOD_FAILED, TypeName(this), PROC_NAME, sRaiseErrorText);
        }

        return;
    }

    private void Class_Terminate()
    {
        mobjCommands = null;
    }

    private object TalentCommandClassEx1_Caption
    {
        get
        {
            object __result = null;
            __result = LoadResString(ridCommandsClassCaption);
            return __result;
        }
    }

    public void SetBlockage(ref object Blockage)
    {
        mobjCommands.SetBlockage(this, "SetBlockage", Blockage);
    }

    public void ResetIdle()
    {
        mobjCommands.ResetIdle(this, "ResetIdle");
    }

    public void StopIdleMotion()
    {
        mobjCommands.StopIdleMotion(this, "StopIdleMotion");
    }

    public void ShowIdleStatus(object Sender = null)
    {
        mobjCommands.ShowIdleStatus(this, "ShowIdleStatus", Sender);
    }
}