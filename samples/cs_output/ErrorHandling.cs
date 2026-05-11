// Converted from ErrorHandling.bas by vb6cs
// Date: 2026-05-11 03:23

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace Converted;
public static class ErrorHandling
{
    public static void WithPreCode(short x)
    {
        string result = "";
        result = "starting";
        try
        {
            result = (int)100 / (int)x.ToString();
            Console.WriteLine(result);
            return;
        }
        catch (Exception ex)
        {
            Console.WriteLine(string.Concat("Error: ", ex.Message));
        }
    }

    public static double SafeDivide(double a, double b)
    {
        double __result = 0;
        try
        {
            __result = a / b;
            return __result;
        }
        catch (Exception ex)
        {
            Console.WriteLine(string.Concat(string.Concat(string.Concat("Division error #", ex.HResult.ToString()), ": "), ex.Message));
            __result = 0;
        }

        return __result;
    }

    public static void NoHandler()
    {
        short x = 0;
        x = 10;
    }

    public static void BestEffort(string path)
    {
        try
        {
            Console.WriteLine("Trying...");
            Console.WriteLine("Done");
        }
        catch
        {
            // On Error Resume Next: exceptions suppressed
        }
    }

    public static void ThrowCustomError(string msg)
    {
        try
        {
            throw new Exception(msg);
            return;
        }
        catch (Exception ex)
        {
            Console.WriteLine(string.Concat("Caught: ", ex.Message));
            // Err.Clear()
        }
    }

    public static void MultiOp(short n)
    {
        int total = 0;
        total = 0;
        try
        {
            total = total + n;
            total = total * 2;
            return;
        }
        catch (Exception ex)
        {
            Console.WriteLine(string.Concat("Op failed: ", ex.Message));
        }
    }
}