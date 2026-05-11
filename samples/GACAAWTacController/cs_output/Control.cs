// Converted from Control.frm by vb6cs
// Date: 2026-05-11 05:50

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace Converted;
public class Control
{
    private Controller mobjController;
    private void cmdHide_Click()
    {
        try
        {
            SetVisible(false);
            return;
        }
        catch
        {
            // On Error Resume Next: exceptions suppressed
        }
    }

    private void cmdQuit_Click()
    {
        try
        {
            if (Console.WriteLine("Are you sure you want to stop this control program?") == vbNo)
            {
                goto Done;
            }

            Unload(this);
        }
        catch (Exception ex)
        {
        }

    Done:
        ;
        return;
    }

    private void Form_Load()
    {
        const object PROC_NAME = "Form_Load";
        SOM objSOM = 0;
        try
        {
            objSOM = new SOM();
            mobjController = objSOM.GetObject(SOM_CONTROLLER);
            this.Caption = PROCESS_NAME;
            lblProjectDesc = App.FileDescription;
            goto Done;
        }
        catch (Exception ex)
        {
            EH.Procedure = PROC_NAME;
            EH.Module = TypeName(this);
            EH_MsgBox();
        }

    Error_Exit:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        Unload(this);
        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        objSOM = null;
        return;
    }

    private void Form_QueryUnload(ref short Cancel, ref short UnloadMode)
    {
        try
        {
            switch (UnloadMode)
            {
                case vbFormControlMenu:
                    SetVisible(false);
                    Cancel = true;
                    break;
            }

            return;
        }
        catch
        {
            // On Error Resume Next: exceptions suppressed
        }
    }

    private void Form_Resize()
    {
        try
        {
            switch (this.WindowState)
            {
                case vbMinimized:
                    cmdHide_Click();
                    this.WindowState = vbNormal;
                    break;
            }

            return;
        }
        catch
        {
            // On Error Resume Next: exceptions suppressed
        }
    }

    private void Form_Unload(ref short Cancel)
    {
        try
        {
            mobjController = null;
            return;
        }
        catch
        {
            // On Error Resume Next: exceptions suppressed
        }
    }

    public void SetVisible(ref bool Visible)
    {
        Attribute(SetVisible.VB_Description == "Sets the visibility");
        const object PROC_NAME = "SetVisible";
        bool bRaiseError = false;
        string sRaiseErrorText = "";
        try
        {
            bRaiseError = false;
            switch (Visible)
            {
                case true:
                    this.WindowState = vbNormal;
                    this.Show();
                    Process.EHMode = ehDisplayAndLog;
                    break;
                case false:
                    this.Hide();
                    Process.EHMode = ehLogOnly;
                    break;
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

    private void tmrDesignTime_Timer()
    {
        try
        {
            mobjController.ControlLogic();
            return;
        }
        catch
        {
            // On Error Resume Next: exceptions suppressed
        }
    }
}