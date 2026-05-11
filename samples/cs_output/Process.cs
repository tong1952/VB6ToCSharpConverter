// Converted from Process.cls by vb6cs
// Date: 2026-05-11 03:23

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace Converted;
public class Process : TalentProcessLib
{
    private Controller mobjController;
    private void Class_Initialize()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        objSOM = new SOM();
        // On Error Resume Next — nested handler; see enclosing try/catch
        mobjController = objSOM.GetObject(SOM_CONTROLLER);
        if (Err)
        {
            // On Error GoTo Error — nested handler not restructured
            mobjController = new Controller();
            objSOM.AddObject(mobjController, SOM_CONTROLLER);
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
            EH_RaiseError(OLEERR_METHOD_FAILED, TypeName(this), PROC_NAME, sRaiseErrorText)();
        }

        return;
    }

    private void Class_Terminate()
    {
        mobjController = null;
    }

    public object CommandInterface
    {
        get
        {
            object __result = null;
            // TODO: On Error GoTo Error — handler label not found in this scope
            bRaiseError = false;
            __result = new SimCommands();
            goto Done;
            Error();
            sRaiseErrorText = ex.Message;
            bRaiseError = true;
            if (ex.HResult & vbObjectError != vbObjectError)
            {
                EH.Module = TypeName(this);
                EH.Procedure = PROC_NAME;
                EH.Item(1) = "NoCommandInterface";
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

    public Controller Controller
    {
        get
        {
            Controller __result = 0;
            // TODO: On Error GoTo Error — handler label not found in this scope
            bRaiseError = false;
            __result = mobjController;
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

    public void Controller_Set(ref Controller vNewValue)
    {
        Attribute(Controller.VB_Description == "Returns the controller object");
        const object PROC_NAME = "Controller (Set)";
        bool bRaiseError = false;
        string sRaiseErrorText = "";
        // On Error GoTo Error — nested handler not restructured
        bRaiseError = false;
        if (mobjController == null)
        {
            mobjController = vNewValue;
        }
        else if (vNewValue == null)
        {
            mobjController = vNewValue;
        }
        else
        {
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
            EH_RaiseError(OLEERR_PROPERTY_SET_FAILED, TypeName(this), PROC_NAME, sRaiseErrorText)();
        }

        return;
    }

    public string Description
    {
        get
        {
            string __result = "";
            __result = PROCESS_DESCRIPTION;
            return __result;
        }
    }

    public Events Events
    {
        get
        {
            Events __result = 0;
            // TODO: On Error GoTo Error — handler label not found in this scope
            bRaiseError = false;
            __result = mobjController.Events;
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

    public void Hide()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
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
            EH_RaiseError(OLEERR_METHOD_FAILED, TypeName(this), PROC_NAME, sRaiseErrorText)();
        }

        return;
    }

    public string Name
    {
        get
        {
            string __result = "";
            __result = PROCESS_NAME;
            return __result;
        }
    }

    public int PID
    {
        get
        {
            int __result = 0;
            // TODO: On Error GoTo Error — handler label not found in this scope
            bRaiseError = false;
            __result = GetCurrentProcessId();
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

    public void Quit()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        mobjController.Commands.Insert(ocShutdown, "Shutdown");
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

        goto ErrorExit;
    ErrorExit:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        mobjController.Shutdown();
        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        if (bRaiseError)
        {
            EH_RaiseError(OLEERR_METHOD_FAILED, TypeName(this), PROC_NAME, sRaiseErrorText)();
        }

        return;
    }

    public void Show()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
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
            EH_RaiseError(OLEERR_METHOD_FAILED, TypeName(this), PROC_NAME, sRaiseErrorText)();
        }

        return;
    }

    public string StatusAsText
    {
        get
        {
            string __result = "";
            // TODO: On Error GoTo Error — handler label not found in this scope
            bRaiseError = false;
            if (mobjController.Initialized == false)
            {
                __result = "Not Initialized";
            }
            else
            {
                __result = "Running";
            }

            goto Done;
            Error();
            __result = "Running";
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

    private object TalentProcess_Application
    {
        get
        {
            object __result = null;
            __result = Controller;
            return __result;
        }
    }

    private object TalentProcess_CommandInterface
    {
        get
        {
            object __result = null;
            __result = CommandInterface;
            return __result;
        }
    }

    private string TalentProcess_Computer
    {
        get
        {
            string __result = "";
            __result = GetComputerName();
            return __result;
        }
    }

    private string TalentProcess_Description
    {
        get
        {
            string __result = "";
            __result = Description;
            return __result;
        }
    }

    private TalentProcessLib.Events TalentProcess_Events
    {
        get
        {
            TalentProcessLib.Events __result = 0;
            __result = Events;
            return __result;
        }
    }

    private void TalentProcess_Hide()
    {
        Hide();
    }

    private string TalentProcess_Name
    {
        get
        {
            string __result = "";
            __result = Name;
            return __result;
        }
    }

    private int TalentProcess_PID
    {
        get
        {
            int __result = 0;
            __result = PID;
            return __result;
        }
    }

    private void TalentProcess_Quit()
    {
        Quit();
    }

    private bool TalentProcess_Ready
    {
        get
        {
            bool __result = false;
            if (mobjController == null)
            {
                __result = false;
            }
            else
            {
                __result = mobjController.Ready;
            }

            return __result;
        }
    }

    private void TalentProcess_Show()
    {
        Show();
    }

    private string TalentProcess_StatusAsText
    {
        get
        {
            string __result = "";
            __result = StatusAsText;
            return __result;
        }
    }

    private bool TalentProcess_Visible
    {
        get
        {
            bool __result = false;
            __result = Visible;
            return __result;
        }
    }

    private void TalentProcess_Visible_Let(ref bool RHS)
    {
        Visible = RHS;
    }

    public bool Visible
    {
        get
        {
            bool __result = false;
            // TODO: On Error GoTo Error — handler label not found in this scope
            bRaiseError = false;
            __result = false;
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

    public void Visible_Let(ref bool vNewValue)
    {
        const object PROC_NAME = "Visible (Let)";
        bool bRaiseError = false;
        string sRaiseErrorText = "";
        // On Error GoTo Error — nested handler not restructured
        bRaiseError = false;
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
            EH_RaiseError(OLEERR_PROPERTY_SET_FAILED, TypeName(this), PROC_NAME, sRaiseErrorText)();
        }

        return;
    }
}