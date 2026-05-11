// Converted from Process.cls by vb6cs
// Date: 2026-05-11 03:41

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
            EH.Module = "Process";
            EH.Procedure = "Hide";
            EH_MsgBox();
        }

        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        if (bRaiseError)
        {
            EH_RaiseError(OLEERR_METHOD_FAILED, "Process", "Hide", sRaiseErrorText)();
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
            EH.Module = "Process";
            EH.Procedure = "Show";
            EH_MsgBox();
        }

        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        if (bRaiseError)
        {
            EH_RaiseError(OLEERR_METHOD_FAILED, "Process", "Show", sRaiseErrorText)();
        }

        return;
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
                EH.Module = "Process";
                EH.Procedure = "Visible";
                EH_MsgBox();
            }

            goto Done;
        Done:
            ;
            // On Error Resume Next — nested handler; see enclosing try/catch
            if (bRaiseError)
            {
                EH_RaiseError(OLEERR_PROPERTY_GET_FAILED, "Process", "Visible", sRaiseErrorText)();
            }

            return __result;
        }
    }

    public void Visible_Let(ref bool vNewValue)
    {
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
            EH.Module = "Process";
            EH.Procedure = "Visible";
            EH_MsgBox();
        }

        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        if (bRaiseError)
        {
            EH_RaiseError(OLEERR_PROPERTY_SET_FAILED, "Process", "Visible", sRaiseErrorText)();
        }

        return;
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
                EH.Module = "Process";
                EH.Procedure = "Events";
                EH_MsgBox();
            }

            goto Done;
        Done:
            ;
            // On Error Resume Next — nested handler; see enclosing try/catch
            if (bRaiseError)
            {
                EH_RaiseError(OLEERR_PROPERTY_GET_FAILED, "Process", "Events", sRaiseErrorText)();
            }

            return __result;
        }
    }

    public void Controller_Set(ref Controller vNewValue)
    {
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
            EH.Module = "Process";
            EH.Procedure = "Controller";
            EH_MsgBox();
        }

        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        if (bRaiseError)
        {
            EH_RaiseError(OLEERR_PROPERTY_SET_FAILED, "Processes", "Controller", sRaiseErrorText)();
        }

        return;
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
                EH.Module = "Process";
                EH.Procedure = "Controller";
                EH_MsgBox();
            }

            goto Done;
        Done:
            ;
            // On Error Resume Next — nested handler; see enclosing try/catch
            if (bRaiseError)
            {
                EH_RaiseError(OLEERR_PROPERTY_GET_FAILED, "Process", "Controller", sRaiseErrorText)();
            }

            return __result;
        }
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

    public string Description
    {
        get
        {
            string __result = "";
            __result = PROCESS_DESCRIPTION;
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
            EH.Module = "Process";
            EH.Procedure = "Quit";
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
            EH_RaiseError(OLEERR_METHOD_FAILED, "Process", "Quit", sRaiseErrorText)();
        }

        return;
    }

    public AcousticCommands CommandInterface
    {
        get
        {
            AcousticCommands __result = 0;
            // TODO: On Error GoTo Error — handler label not found in this scope
            bRaiseError = false;
            __result = new AcousticCommands();
            goto Done;
            Error();
            sRaiseErrorText = ex.Message;
            bRaiseError = true;
            if (ex.HResult & vbObjectError != vbObjectError)
            {
                EH.Module = "Process";
                EH.Procedure = "CommandInterface";
                EH.Item(1) = "DYController.AcousticCommands";
                EH_MsgBox();
            }

            goto Done;
        Done:
            ;
            // On Error Resume Next — nested handler; see enclosing try/catch
            if (bRaiseError)
            {
                EH_RaiseError(OLEERR_PROPERTY_GET_FAILED, "Application", "CommandInterface", sRaiseErrorText)();
            }

            return __result;
        }
    }

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
            EH.Module = "Process";
            EH.Procedure = "Class_Initialize";
            EH_MsgBox();
        }

        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        if (bRaiseError)
        {
            EH_RaiseError(OLEERR_METHOD_FAILED, "Process", "Class_Initialize", sRaiseErrorText)();
        }

        return;
    }

    private void Class_Terminate()
    {
        mobjController = null;
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

    private void TalentProcess_Visible_Let(ref bool RHS)
    {
        Visible = RHS;
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
}