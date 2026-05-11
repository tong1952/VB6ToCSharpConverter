// Converted from GACAAWTacControl.Bas by vb6cs
// Date: 2026-05-11 05:50

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace Converted;
public static class GACAAWTacControl
{
    private const object MOD_NAME = "modACControl";
    public enum ResIDs
    {
        ridNoResponseFromBKConnect = 120,
        ridBKConnectCmdInvalidArg = 121,
        ridBKConnectCmdError = 122,
        ridBKConnectRejectedCmd = 123,
        ridTCPCommsFaulted = 124
    }

    public const object SOM_CONTROLLER = "Controller";
    public const object SPG_CONSTANTS = "Constants";
    public const object SPG_DATA = "Data";
    public enum CCNames
    {
        CC_xxx = 0
    }

    [DllImport("kernel32", EntryPoint = "GetCurrentProcessId")]
    public static extern int GetCurrentProcessId();
    [DllImport("kernel32", EntryPoint = "GetCurrentProcess")]
    public static extern int GetCurrentProcess();
    [DllImport("kernel32", EntryPoint = "TerminateProcess")]
    public static extern int TerminateProcess(int hProcess, int uExitCode);
    [DllImport("kernel32", EntryPoint = "RtlMoveMemory")]
    public static extern void CopyMemory(IntPtr pDest, IntPtr pSource, int ByteLen);
    [DllImport("kernel32", EntryPoint = "RtlMoveMemory")]
    public static extern void CopyMemory2(ref IntPtr Destination, ref IntPtr Source, int Length);
    public enum AE
    {
        InvalidActiveXLaunch = 101,
        CantCreateConfigurations = 102,
        InitializeFailedParsMissing = 103,
        CantLoadConstants = 104,
        BKConnectResponseError = 105,
        CommunicationsFaulted = 106
    }

    public enum Opcodes
    {
        ocInitialize = 1,
        ocSuspend = 2,
        ocShowStatus = 3,
        ocResetCommunications = 4,
        ocStartLogging = 5,
        ocStopLogging = 6,
        ocSendCommand = 7,
        ocSetIPAddr = 8,
        ocStartRecord = 20,
        ocStopRecord = 22,
        ocShutdown = 1000
    }

    public enum BKConnectStates
    {
        gsUninitialized = 0,
        gsReady = 1,
        gsRecording = 4,
        gsRecordingStopped = 8
    }

    public const object NOFAULT = 1;
    public const object FAULT = 0;
    public const string ILCK_GENERAL = "General";
    public const object ILCK_COM_OK = "ComOK";
    public const object ILCK_OPERATION_OK = "OperationOK";
    public const object SPG_STATUSES = "Statuses";
    public const object STS_ACOUSTIC_INITIALIZED = "ACOUSTIC_INITIALIZED";
    public const object STS_ACOUSTIC_FAULTED = "ACOUSTIC_FAULTED";
    public const object PROCESS_NAME = "Acoustic System";
    public const object PROCESS_DESCRIPTION = "BKConnect Communications Program";
    public const object REG_PROCESSES = "Software\\ReACT Technologies\\Talent\\4.0\\Processes";
    public const object REG_ACOUSTIC = "Software\\ReACT Technologies\\Talent\\4.0\\Acoustic System";
    public static object ArgValue(ref object Arg, VbVarType TargetVariableType = vbVariant)
    {
        object __result = null;
        object vValue = null;
        if (Arg is CommandListLib.Argument)
        {
            vValue = Arg.Value;
        }
        else
        {
            vValue = Arg;
        }

        if (TargetVariableType == vbVariant)
        {
            __result = vValue;
        }
        else
        {
            __result = Convert(vValue, TargetVariableType);
        }

        return __result;
    }

    public static string ArgValueAsString(ref object Expression)
    {
        string __result = "";
        try
        {
            switch (VarType(Expression))
            {
                case vbString:
                    __result = string.Concat(string.Concat("\"", Expression.ToString()), "\"");
                    break;
                default:
                    __result = CStrUS(Expression);
                    break;
            }

            return __result;
        }
        catch
        {
            // On Error Resume Next: exceptions suppressed
        }

        return __result;
    }

    public static string ByteArrayToString(ref object InputByteArray)
    {
        string __result = "";
        byte[] xBytes;
        int i = 0;
        string sHex = "";
        try
        {
            for (i = 0; i <= InputByteArray.Length - 1; i++)
            {
                if (ByteArrayToString != "")
                {
                    __result = string.Concat(ByteArrayToString, " ");
                }

                sHex = Convert.ToString(InputByteArray(i), 16);
                if (sHex.Length < 2)
                {
                    sHex = string.Concat("0", sHex);
                }

                __result = string.Concat(ByteArrayToString, sHex);
                if (i >= 255)
                {
                    __result = string.Concat(ByteArrayToString, " ...");
                    break;
                }
            }
        }
        catch (Exception ex)
        {
        }

    Done:
        ;
        return __result;
    }

    public static string GetCmdArgs(ref CommandListLib.Arguments Arguments)
    {
        string __result = "";
        short i = 0;
        try
        {
            for (i = 1; i <= Arguments.Count; i++)
            {
                if (GetCmdArgs != "")
                {
                    __result = string.Concat(GetCmdArgs, ", ");
                }

                __result = string.Concat(GetCmdArgs, ArgValueAsString(Arguments(i).Value));
            }

            if (GetCmdArgs != "")
            {
                __result = string.Concat(" ", GetCmdArgs);
            }

        Done:
            ;
            return __result;
        }
        catch
        {
            // On Error Resume Next: exceptions suppressed
        }

        return __result;
    }

    public static object IsSameCommand(ref Command LastCommand, ref string ReceivedMemnomic)
    {
        object __result = null;
        try
        {
            if (LastCommand.TagVariant.LastCommand.CommandText.Trim() == ReceivedMemnomic.Trim())
            {
                __result = true;
            }
        }
        catch (Exception ex)
        {
        }

    Done:
        ;
        return __result;
    }

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

    public static void LogCommMessage(string MessageText, ref object DataPacket, TalentProcessLib.MessageTypes MessageType = mtInformation)
    {
        string sText = "";
        int i = 0;
        try
        {
            sText = MessageText;
            if (IsMissing(DataPacket))
            {
            }
            else
            {
                if (sText.Substring(sText.Length - 1) != " ")
                {
                    sText = string.Concat(sText, " ");
                }

                if (VarType(DataPacket) == vbString)
                {
                    sText = string.Concat(sText, "\"");
                    for (i = 1; i <= DataPacket.Length; i++)
                    {
                        switch ((short)Microsoft.VisualBasic.Strings.Mid(DataPacket, i, 1).FirstOrDefault())
                        {
                            case 2:
                                sText = string.Concat(sText, "{STX}");
                                break;
                            case 3:
                                sText = string.Concat(sText, "{ETX}");
                                break;
                            case 10:
                                sText = string.Concat(sText, "{LF}");
                                break;
                            case 13:
                                sText = string.Concat(sText, "{CR}");
                                break;
                            default:
                                sText = string.Concat(sText, Microsoft.VisualBasic.Strings.Mid(DataPacket, i, 1));
                                break;
                        }

                        if (i >= 200)
                        {
                            sText = string.Concat(string.Concat(string.Concat(sText, ".... "), DataPacket.Length - 200), " more characters");
                            break;
                        }
                    }

                    sText = string.Concat(sText, "\"");
                }
                else
                {
                    sText = string.Concat(string.Concat(string.Concat(sText, "\"<"), LenB(DataPacket)), " bytes binary data>\"");
                }
            }

            LogMessage(sText, MessageType);
            return;
        }
        catch
        {
            // On Error Resume Next: exceptions suppressed
        }
    }

    public static void Main()
    {
        const object PROC_NAME = "Main";
        bool bRunning = false;
        // On Error Resume Next — nested handler; see enclosing try/catch
        if (App.PrevInstance)
        {
            if (App.StartMode != vbSModeStandalone)
            {
                EH.AppError = AE.InvalidActiveXLaunch;
                EH.Item(1) = PROCESS_NAME;
                EH_MsgBox();
            }

            End();
        }

        try
        {
            App.StartLogging("", vbLogAuto);
            App.OleServerBusyRaiseError = true;
            App.OleServerBusyTimeout = 1000;
            Process.Title = PROCESS_NAME;
            Process.EHCaption = PROCESS_NAME;
            Process.EHMode = ehLogOnly;
            Process.EHMsgsToLog = ehCriticalExclamation;
            Process.LogMode = logToNT;
            goto Done;
        }
        catch (Exception ex)
        {
            if (Err != 364)
            {
                EH.Module = MOD_NAME;
                EH.Procedure = PROC_NAME;
                EH_MsgBox();
            }
        }

    ErrorExit:
        ;
        EH.AppError = ERR_CANT_START;
        EH.Module = MOD_NAME;
        EH.Procedure = PROC_NAME;
        EH_MsgBox();
        // On Error Resume Next — nested handler; see enclosing try/catch
        goto Done;
    Done:
        ;
        return;
    }

    public static string QTrim(ref object Expression)
    {
        string __result = "";
        try
        {
            __result = Expression;
            __result = Expression.Trim();
            if (QTrim.Substring(0, 1) == "\"")
            {
                __result = QTrim.Substring(2 - 1);
            }

            if (QTrim.Substring(QTrim.Length - 1) == "\"")
            {
                __result = Microsoft.VisualBasic.Strings.Mid(QTrim, 1, QTrim.Length - 1);
            }

            return __result;
        }
        catch
        {
            // On Error Resume Next: exceptions suppressed
        }

        return __result;
    }
}