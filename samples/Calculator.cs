// Converted from Calculator.bas by vb6cs
// Date: 2026-05-07 23:23

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace Converted;
public static class Calculator
{
    public const string APP_NAME = "VB6 Calc";
    private int mRunCount;
    public double Add(double a, double b)
    {
        double __result = default;
        __result = a + b;
        return __result;
    }

    public double Divide(double a, double b)
    {
        double __result = default;
        // On Error GoTo ErrHandler — see catch block
        ;
        if (b == 0)
        {
            Console.WriteLine("Division by zero!");
            __result = 0;
            return __result;
        }

        __result = a / b;
        return __result;
        ErrHandler:
            ;
        Console.WriteLine(string.Concat("Error: ", Err.Description));
        __result = 0;
        return __result;
    }

    public void PrintTable(short n)
    {
        short i = 0;
        double result = 0;
        for (i = 1; i <= n; i++)
        {
            result = i * i;
            switch (result)
            {
                case 1:
                    Debug.Print("One");
                    break;
                case 4:
                case 9:
                    Debug.Print(string.Concat("Small square: ", result.ToString()));
                    break;
                case 50:
                    Debug.Print(string.Concat("Large: ", result.ToString()));
                    break;
                default:
                    Debug.Print(string.Concat("Square: ", result.ToString()));
                    break;
            }
        }
    }

    public double SumArray(double[] arr)
    {
        double __result = default;
        double total = 0;
        object item = null;
        total = 0;
        foreach (var item in arr)
        {
            total = total + (double)item;
        }

        __result = total;
        return __result;
    }

    public string BuildReport(string name, short score)
    {
        string __result = default;
        string report = "";
        report = string.Concat("Name: ", name.ToUpper());
        report = string.Concat(string.Concat(report, " | Score: "), score.ToString());
        report = string.Concat(report, " | Grade: ");
        if (score >= 90)
        {
            report = string.Concat(report, "A");
        }
        else if (score >= 80)
        {
            report = string.Concat(report, "B");
        }
        else if (score >= 70)
        {
            report = string.Concat(report, "C");
        }
        else
        {
            report = string.Concat(report, "F");
        }

        __result = report;
        return __result;
    }

    public int Factorial(int n)
    {
        int __result = default;
        int result = 0;
        int i = 0;
        result = 1;
        i = n;
        while (i > 1)
        {
            result = result * i;
            i = i - 1;
        }

        __result = result;
        return __result;
    }

    public int IncrementCount()
    {
        int __result = default;
        static int count = 0;
        count = count + 1;
        mRunCount = count;
        __result = count;
        return __result;
    }
}