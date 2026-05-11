// Converted from Derivative.cls by vb6cs
// Date: 2026-05-11 03:43

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace Converted;
public class Derivative
{
    private const object Fuss = 1E-10;
    private float mfLastError;
    private float mfLastOutput;
    public float Output(float FeedbackError, float UpdateTime, float Td, float Alpha)
    {
        float __result = 0;
        // TODO: On Error GoTo Error — handler label not found in this scope
        bRaiseError = false;
        if (Td < Fuss * UpdateTime)
        {
            Td = Fuss * UpdateTime;
        }

        fTransition = Math.Exp(-UpdateTime * Alpha / Td);
        __result = fTransition * mfLastOutput + Td / UpdateTime * 1 - fTransition * FeedbackError - mfLastError;
        mfLastError = FeedbackError;
        mfLastOutput = Output;
        goto Done;
        Error();
        sRaiseErrorText = ex.Message;
        bRaiseError = true;
        goto Done;
    Done:
        ;
        // On Error Resume Next — nested handler; see enclosing try/catch
        if (bRaiseError)
        {
            EH_RaiseError(OLEERR_METHOD_FAILED, "Derivative", "Output", sRaiseErrorText)();
        }

        return __result;
    }
}