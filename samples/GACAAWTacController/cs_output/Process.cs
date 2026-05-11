// Converted from Process.cls by vb6cs
// Date: 2026-05-11 05:50

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
        const object PROC_NAME = "Class_Initialize";
        bool bRaiseError = false;
        string sRaiseErrorText = "";
        SOM objSOM = 0;
        try
        {
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
            EH_RaiseError(OLEERR_METHOD_FAILED, TypeName(this), PROC_NAME, sRaiseErrorText)();
        }

        return;
    }

    private void Class_Terminate()
    {
        mobjController = null;
    }

    public AcousticCommands CommandInterface
    {
        get
        {
            AcousticCommands __result = 0;
            Attribute(CommandInterface.VB_Description == "Returns a new Interface object to allow external interfaces with this application");
            const object PROC_NAME = "CommandInterface (Get)";
            bool bRaiseError = false;
            string sRaiseErrorText = "";
            try
            {
                bRaiseError = false;
                __result = new AcousticCommands();
            }
            catch (Exception ex)
            {
                sRaiseErrorText = ex.Message;
                bRaiseError = true;
                if (ex.HResult & vbObjectError != vbObjectError)
                {
                    EH.Module = TypeName(this);
                    EH.Procedure = PROC_NAME;
                    EH.Item(1) = "DYController.AcousticCommands";
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

    public Controller Controller
    {
        get
        {
            Controller __result = 0;
            const object PROC_NAME = "Controller (Get)";
            bool bRaiseError = false;
            string sRaiseErrorText = "";
            try
            {
                bRaiseError = false;
                __result = mobjController;
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
    Error:
        ;
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
            Attribute(Events.VB_Description == "Sets the events object");
            Attribute(Events.VB_MemberFlags == "40");
            const object PROC_NAME = "Events (Get)";
            bool bRaiseError = false;
            string sRaiseErrorText = "";
            try
            {
                bRaiseError = false;
                __result = mobjController.Events;
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

    public void Hide()
    {
        Attribute(Hide.VB_Description == "Hides the program form(s)");
        const object PROC_NAME = "Hide";
        bool bRaiseError = false;
        string sRaiseErrorText = "";
        try
        {
            bRaiseError = false;
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
            Attribute(PID.VB_Description == "Returns the PID of this process");
            const object PROC_NAME = "PID (Get)";
            bool bRaiseError = false;
            string sRaiseErrorText = "";
            try
            {
                bRaiseError = false;
                __result = GetCurrentProcessId();
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

    public void Quit()
    {
        Attribute(Quit.VB_Description == "Stops this program");
        const object PROC_NAME = "Quit";
        bool bRaiseError = false;
        string sRaiseErrorText = "";
        try
        {
            bRaiseError = false;
            mobjController.Commands.Insert(ocShutdown, "Shutdown");
            goto Done;
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
        Attribute(Show.VB_Description == "Displays the program form");
        const object PROC_NAME = "Show";
        bool bRaiseError = false;
        string sRaiseErrorText = "";
        try
        {
            bRaiseError = false;
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
            EH_RaiseError(OLEERR_METHOD_FAILED, TypeName(this), PROC_NAME, sRaiseErrorText)();
        }

        return;
    }

    public string StatusAsText
    {
        get
        {
            string __result = "";
            Attribute(StatusAsText.VB_Description == "Returns a status information string");
            const object PROC_NAME = "StatusAsText (Get)";
            bool bRaiseError = false;
            string sRaiseErrorText = "";
            float fMean = 0;
            float fMax = 0;
            float fMin = 0;
            float fStdDev = 0;
            float fFreq = 0;
            try
            {
                bRaiseError = false;
                if (mobjController.Initialized == false)
                {
                    __result = "Not Initialized";
                }
                else
                {
                    __result = "Running";
                }
            }
            catch (Exception ex)
            {
                __result = "Running";
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
            Attribute(Visible.VB_Description == "Sets/returns the visibility");
            const object PROC_NAME = "Visible (Get)";
            bool bRaiseError = false;
            string sRaiseErrorText = "";
            try
            {
                bRaiseError = false;
                __result = false;
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

    public void Visible_Let(ref bool vNewValue)
    {
        const object PROC_NAME = "Visible (Let)";
        bool bRaiseError = false;
        string sRaiseErrorText = "";
        // On Error GoTo Error — nested handler not restructured
        bRaiseError = false;
        goto Done;
    Error:
        ;
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