// Converted from Filter1.cls by vb6cs
// Date: 2026-05-11 03:23

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace Converted;
public class Filter1
{
    private float mfLastOutput;
    private float mfLastTime_s;
    public float Filter1Control(ref object CurrentValue, ref object TimeConstant_s)
    {
        float __result = 0;
        Attribute(Filter1Control.VB_Description == "IIR (infinite impulse response) single pole filter. Time constant is in seconds.");
        Attribute(Filter1Control.VB_UserMemId == 0);
        float fTime_s = 0;
        float fDeltaTime_s = 0;
        float fAlpha = 0;
        float fOutput = 0;
        fTime_s = Timer;
        if (fTime_s < mfLastTime_s)
        {
            fDeltaTime_s = 86400 - mfLastTime_s + fTime_s;
        }
        else
        {
            fDeltaTime_s = fTime_s - mfLastTime_s;
        }

        mfLastTime_s = fTime_s;
        fAlpha = fDeltaTime_s / TimeConstant_s + fDeltaTime_s;
        fOutput = fAlpha * CurrentValue + 1 - fAlpha * mfLastOutput;
        __result = fOutput;
        mfLastOutput = fOutput;
        return __result;
    }
}