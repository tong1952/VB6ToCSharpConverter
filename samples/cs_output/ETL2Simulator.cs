// Converted from ETL2Simulator.bas by vb6cs
// Date: 2026-05-11 03:23

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace Converted;
public static class ETL2Simulator
{
    private const object MOD_NAME = "modCTRCSimulator";
    public enum ResIDs
    {
        ridCantCreateObject = 101,
        ridUnrecognizedOpcode = 102,
        ridControllerShuttingDown = 103,
        ridRunning = 104,
        ridRunningAt = 105,
        ridErrDisplayingMessage = 106
    }

    public const object SOM_CONTROLLER = "Controller";
    public const object SPG_CONSTANTS = "Constants";
    public const object SPG_DATA = "Data";
    public enum CCNames
    {
        ccUpdateTime = 0,
        ccTCPIPAddr = 1,
        ccTCPIPPort = 2,
        ccUDPPort = 3,
        ccUDPLocalPort = 4,
        ccUDPAddr = 5
    }

    [DllImport("kernel32", EntryPoint = "GetCurrentProcessId")]
    public static extern int GetCurrentProcessId();
    [DllImport("kernel32", EntryPoint = "GetCurrentProcess")]
    public static extern int GetCurrentProcess();
    [DllImport("kernel32", EntryPoint = "TerminateProcess")]
    public static extern int TerminateProcess(int hProcess, int uExitCode);
    [DllImport("kernel32", EntryPoint = "RtlMoveMemory")]
    public static extern void CopyMemory(IntPtr pDest, IntPtr pSource, int ByteLen);
    public enum AE
    {
        InvalidActiveXLaunch = 101,
        CantLoadConstants = 104
    }

    public enum Opcodes
    {
        ocInitialize = 1,
        ocSuspend = 2,
        ocShutdown = 1000
    }

    public const object NOFAULT = 1;
    public const object FAULT = 0;
    public const object UseDefaultValue = -10000000000;
    public const object SPG_STATUSES = "Statuses";
    public const object STS_SIM_INITIALIZED = "SIM_INITIALIZED";
    public const object PROCESS_NAME = "Sim";
    public const object PROCESS_DESCRIPTION = "ETL2 Simulator";
    public const object REG_PROCESSES = "Software\\ReACT Technologies\\Talent\\4.0\\Processes";
    public static string LoadResString(ref int ID, params object[] ItemTextArgs)
    {
        string __result = "";
        if (ID & STD_TALENT_STRING_BASE == STD_TALENT_STRING_BASE)
        {
            __result = LoadStdTalentString(ID, ItemTextArgs);
        }
        else
        {
            __result = VB.LoadResString(ID);
            __result = ReplaceItemText(LoadResString, ItemTextArgs);
        }

        return __result;
    }
}