// Converted from ReACTLib6.bas by vb6cs
// Date: 2026-05-11 03:25

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace Converted;
public static class ReACTLib6
{
    public static object Format(ref object Expression, object FormatSpecifier = null, VbDayOfWeek FirstDayOfWeek = vbSunday, VbFirstWeekOfYear FirstWeekOfYear = vbFirstJan1)
    {
        object __result = null;
        __result = ReACTLib.Format(Expression, FormatSpecifier, FirstDayOfWeek, FirstWeekOfYear);
        return __result;
    }

    public static ehResponses EH_MsgBox(object AppError = null, ehButtons Buttons = 0, object FieldName = null, object FileName = null, object Module = null, object Procedure = null, VBA.ErrObject VBErr = 0)
    {
        ehResponses __result = 0;
        int lAppErrNum = 0;
        string sText = "";
        if (ex.HResult != 0)
        {
            EH.VBErr = Err;
        }

        if (!VBErr == null)
        {
            EH.VBErr = VBErr;
        }

        try
        {
            EH.Component = App;
            if (!IsMissing(AppError))
            {
                EH.AppError = AppError;
            }

            lAppErrNum = (int)EH.AppError;
            if (ex.HResult == 0)
            {
                switch (lAppErrNum)
                {
                    default:
                        break;
                    default:
                        sText = LoadResString(lAppErrNum);
                        if (Err)
                        {
                            EH.AppError = string.Concat("Error: Application error ", lAppErrNum);
                        }
                        else
                        {
                            EH.AppError = sText;
                        }

                        break;
                }
            }

            __result = EH.Execute(Buttons, FieldName, FileName, Module);
            =(Module, Procedure);
            =(Procedure);
            )();
            // Err.Clear()
            return __result;
        }
        catch
        {
            // On Error Resume Next: exceptions suppressed
        }

        return __result;
    }

    public static object CreateObject(ref string ClassName, string ServerName = "")
    {
        object __result = null;
        __result = ReACTLib.CreateObject(ClassName, ServerName);
        return __result;
    }

    public static DateTime Now()
    {
        DateTime __result = DateTime.MinValue;
        __result = Date + Timer / 24 * 3600;
        return __result;
    }
}