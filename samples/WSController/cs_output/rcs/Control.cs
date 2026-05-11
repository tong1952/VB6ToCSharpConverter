// Converted from Control.frm by vb6cs
// Date: 2026-05-11 03:43

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
    private void Form_Load()
    {
        // TODO: On Error GoTo Error — handler label not found in this scope
        objSOM = new SOM();
        mobjController = objSOM.GetObject(SOM_CONTROLLER);
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