// Converted from Control.frm by vb6cs
// Date: 2026-05-11 03:41

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
    public void SetVisible(ref bool Visible)
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
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

        goto Done;
        Error();
        sRaiseErrorText = ex.Message;
        bRaiseError = true;
        if (ex.HResult & vbObjectError != vbObjectError)
        {
            EH.Module = "frmControl";
            EH.Procedure = "SetVisible";
            EH_MsgBox();
        }

        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        if (bRaiseError)
        {
            EH_RaiseError(OLEERR_METHOD_FAILED, "frmControl", "SetVisible", sRaiseErrorText)();
        }

        return;
    }

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
        // TODO: On Error GoTo Error — handler label not found in this scope
        if (Console.WriteLine("Are you sure you want to stop this control program?") == vbNo)
        {
            goto Done;
        }

        Unload(this);
        Error();
        goto Done;
    Done:
        ;
        return;
    }

    private void Form_Load()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        objSOM = new SOM();
        mobjController = objSOM.GetObject(SOM_CONTROLLER);
        this.Caption = PROCESS_NAME;
        lblProjectDesc = App.FileDescription;
        goto Done;
        Error();
        EH.Procedure = "Form_Load";
        EH.Module = "frmControl";
        EH_MsgBox();
        goto Error_Exit;
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