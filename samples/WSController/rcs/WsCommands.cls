head 1.9;
branch ;
access ;
symbols CD%204%2E00%2E0021:1.9 CD%204%2E00%2E0020:1.9 CD%204%2E00%2E0019:1.9
        CD%204%2E00%2E0018:1.9 CD%204%2E00%2E0017:1.9 CD%204%2E00%2E0016:1.9 4%2E00%2E0082:1.9
        CD%204%2E00%2E0015:1.8 CD%204%2E00%2E0014:1.8 CD%204%2E00%2E0013:1.8 4%2E00%2E0065:1.8
        4%2E00%2E0062:1.7 4%2E00%2E0049:1.6 4%2E00%2E0044:1.5 4%2E00%2E0041:1.4 4%2E00%2E0039:1.3
        4%2E00%2E0034:1.2;
locks ; strict;
comment @@;


1.9
date 2019.11.20.23.27.15; author Kavanagh; state Develop;
branches ;
next 1.8;

1.8
date 2019.09.16.01.26.52; author Kavanagh; state Develop;
branches ;
next 1.7;

1.7
date 2019.09.11.21.31.25; author Kavanagh; state Develop;
branches ;
next 1.6;

1.6
date 2019.09.06.23.48.17; author Kavanagh; state Develop;
branches ;
next 1.5;

1.5
date 2019.07.04.05.31.27; author Kavanagh; state Develop;
branches ;
next 1.4;

1.4
date 2019.07.04.00.43.49; author Kavanagh; state Develop;
branches ;
next 1.3;

1.3
date 2019.04.26.02.44.36; author Kavanagh; state Develop;
branches ;
next 1.2;

1.2
date 2019.04.17.03.31.37; author Kavanagh; state Develop;
branches ;
next 1.1;

1.1
date 2019.04.16.23.31.54; author Kavanagh; state Develop;
branches ;
next ;

ext
@project Z:/Client Facilities/CAERI/Source Code/Applications/WSController/CWTCSWSController.pj;
@

desc
@@


1.9
log
@Added basic collector flaps commands@
text
@VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WsCommands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private mobjController As Controller
Private mobjConstants As SharedPropertyGroup
Private mobjStatuses As Statuses


Public Sub InitializeWindSpeed(CommandObject As Object, MethodName As String)

' Initialize the program

    Const PROC_NAME = "InitializeWindSpeed"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    Dim objCommand As Command
    
    On Error GoTo Error

    bRaiseError = False

    Set objCommand = GetController.Commands.Append(Opcodes.ocInitialize, MethodName)
    mobjStatuses.Reset STS_WINDSPEED_INITIALIZED, False

    GoTo Done

Error:
    If Err.Number = cmderrTimedout Then
        Resume Done
    End If
    bRaiseError = True
    lRaiseErrorNumber = CCommandError(Err.Number)
    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, Err.Description)
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        If objCommand.SuppressLogging = False Then           ' or error
        Else
            lRaiseErrorNumber = lRaiseErrorNumber Or tceSuppressLogging
        End If
        On Error GoTo 0
        Err.Raise lRaiseErrorNumber, PROCESS_NAME, sRaiseErrorText
    End If

    Exit Sub

End Sub

Public Sub SuspendWindSpeed(CommandsObject As Object, MethodName As String)

' Suspend the control program

    Const PROC_NAME = "SuspendWindSpeed"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    
    On Error GoTo Error

    bRaiseError = False
    
    With GetController
        .Commands.Clear MethodName ' Clear all   , .InvalidOpcodes(CStr(ocSuspend))
        .Commands.Append ocSuspend, MethodName
    End With
    
    GoTo Done

Error:
    bRaiseError = True
    lRaiseErrorNumber = CCommandError(Err.Number)
    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, Err.Description)
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Err.Raise lRaiseErrorNumber, PROCESS_NAME, sRaiseErrorText
    End If

    Exit Sub
End Sub

Public Sub StartMainFan(CommandsObject As Object, MethodName As String)

' Starts the Main Fan

    Const PROC_NAME = "StartMainFan"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    
    On Error GoTo Error

    bRaiseError = False
    
    With GetController
        .CheckInterlock (CStr(ocInitialize))
        .CheckInterlock CStr(Opcodes.ocStartMainFan)
        .Commands.Clear MethodName, .InvalidOpcodes(CStr(ocStartMainFan))
        .Commands.Append ocStartMainFan, MethodName
    End With
    mobjStatuses.Reset STS_MAINFAN_ON, False
    
    GoTo Done

Error:
    bRaiseError = True
    lRaiseErrorNumber = CCommandError(Err.Number)
    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, Err.Description)
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Err.Raise lRaiseErrorNumber, PROCESS_NAME, sRaiseErrorText
    End If

    Exit Sub

End Sub

Public Sub StopMainFan(CommandsObject As Object, MethodName As String)

' Stops the main fan

    Const PROC_NAME = "StopMainFan"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    
    On Error GoTo Error

    bRaiseError = False

    With GetController
        .CheckInterlock CStr(ocInitialize)
        .CheckInterlock CStr(Opcodes.ocStopMainFan)
        .Commands.Clear MethodName, .InvalidOpcodes(CStr(ocStopMainFan))
        .Commands.Append ocStopMainFan, MethodName
    End With
    mobjStatuses.Reset STS_MAINFAN_OFF, False
    
    GoTo Done

Error:
    bRaiseError = True
    lRaiseErrorNumber = CCommandError(Err.Number)
    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, Err.Description)
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Err.Raise lRaiseErrorNumber, PROCESS_NAME, sRaiseErrorText
    End If

    Exit Sub

End Sub

Public Sub StartMainFanAux(CommandsObject As Object, MethodName As String)

' Starts the Main Fan Auxiliary

    Const PROC_NAME = "StartMainFanAux"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    
    On Error GoTo Error

    bRaiseError = False
    
    With GetController
        .CheckInterlock CStr(ocInitialize)
        .CheckInterlock CStr(Opcodes.ocStartMainFanAux)
        .Commands.Clear MethodName, .InvalidOpcodes(CStr(ocStartMainFanAux))
        .Commands.Append ocStartMainFanAux, MethodName
    End With
    mobjStatuses.Reset STS_MAINFAN_AUX_ON, False
    
    GoTo Done

Error:
    bRaiseError = True
    lRaiseErrorNumber = CCommandError(Err.Number)
    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, Err.Description)
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Err.Raise lRaiseErrorNumber, PROCESS_NAME, sRaiseErrorText
    End If

    Exit Sub

End Sub

Public Sub StopMainFanAux(CommandsObject As Object, MethodName As String)

' Stops the main fan

    Const PROC_NAME = "StopMainFanAux"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    
    On Error GoTo Error

    bRaiseError = False

    With GetController
        .CheckInterlock CStr(ocInitialize)
        .CheckInterlock CStr(Opcodes.ocStopMainFanAux)
        .Commands.Clear MethodName, .InvalidOpcodes(CStr(ocStopMainFanAux))
        .Commands.Append ocStopMainFanAux, MethodName
    End With
    mobjStatuses.Reset STS_MAINFAN_AUX_OFF, False
    
    GoTo Done

Error:
    bRaiseError = True
    lRaiseErrorNumber = CCommandError(Err.Number)
    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, Err.Description)
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Err.Raise lRaiseErrorNumber, PROCESS_NAME, sRaiseErrorText
    End If

    Exit Sub

End Sub
Public Sub ResetMainFanFaults(CommandsObject As Object, MethodName As String)

' Starts the Main Fan

    Const PROC_NAME = "ResetMainFanFaults"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    
    On Error GoTo Error

    bRaiseError = False
    
    With GetController
        .CheckInterlock (CStr(ocInitialize))
        .CheckInterlock CStr(Opcodes.ocResetFaults)
        .Commands.Clear MethodName, .InvalidOpcodes(CStr(ocResetFaults))
        .Commands.Append ocResetFaults, MethodName
    End With
    
    GoTo Done

Error:
    bRaiseError = True
    lRaiseErrorNumber = CCommandError(Err.Number)
    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, Err.Description)
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Err.Raise lRaiseErrorNumber, PROCESS_NAME, sRaiseErrorText
    End If

    Exit Sub

End Sub
Public Sub StartPurge(CommandsObject As Object, MethodName As String)

' Starts the Main Fan

    Const PROC_NAME = "StartPurge"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    
    On Error GoTo Error

    bRaiseError = False
    
    With GetController
        .CheckInterlock (CStr(ocStartPurge))
        .Commands.Clear MethodName, .InvalidOpcodes(CStr(ocStartPurge))
        .Commands.Append ocStartPurge, MethodName
    End With
    
    GoTo Done

Error:
    bRaiseError = True
    lRaiseErrorNumber = CCommandError(Err.Number)
    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, Err.Description)
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Err.Raise lRaiseErrorNumber, PROCESS_NAME, sRaiseErrorText
    End If

    Exit Sub

End Sub
Public Sub StopPurge(CommandsObject As Object, MethodName As String)

' Stops the Main Fan

    Const PROC_NAME = "StopPurge"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    
    On Error GoTo Error

    bRaiseError = False
    
    With GetController
        .CheckInterlock (CStr(ocStopPurge))
        .Commands.Clear MethodName, .InvalidOpcodes(CStr(ocStopPurge))
        .Commands.Append ocStopPurge, MethodName
    End With
    
    GoTo Done

Error:
    bRaiseError = True
    lRaiseErrorNumber = CCommandError(Err.Number)
    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, Err.Description)
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Err.Raise lRaiseErrorNumber, PROCESS_NAME, sRaiseErrorText
    End If

    Exit Sub

End Sub

Public Sub StartIdle(CommandsObject As Object, MethodName As String)

' Starts the Main Fan

    Const PROC_NAME = "StartIdle"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    
    On Error GoTo Error

    bRaiseError = False
        
    With GetController
        .CheckInterlock CStr(ocInitialize)
        .CheckInterlock CStr(Opcodes.ocStartIdle)
        .Commands.Clear MethodName, .InvalidOpcodes(CStr(ocStartIdle))
        .Commands.Append ocStartIdle, MethodName
    End With
    mobjStatuses.Reset STS_IDLE_STARTED, False
    
    GoTo Done

Error:
    bRaiseError = True
    lRaiseErrorNumber = CCommandError(Err.Number)
    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, Err.Description)
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Err.Raise lRaiseErrorNumber, PROCESS_NAME, sRaiseErrorText
    End If

    Exit Sub

End Sub

Public Sub StopIdle(CommandsObject As Object, MethodName As String)

' Stops the main fan

    Const PROC_NAME = "StopIdle"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    
    On Error GoTo Error

    bRaiseError = False

    With GetController
        .CheckInterlock CStr(ocInitialize)
        .CheckInterlock CStr(Opcodes.ocStopIdle)
        .Commands.Clear MethodName, .InvalidOpcodes(CStr(ocStopIdle))
        .Commands.Append ocStopIdle, MethodName
    End With
    mobjStatuses.Reset STS_IDLE_STOPPED, False
    
    GoTo Done

Error:
    bRaiseError = True
    lRaiseErrorNumber = CCommandError(Err.Number)
    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, Err.Description)
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Err.Raise lRaiseErrorNumber, PROCESS_NAME, sRaiseErrorText
    End If

    Exit Sub

End Sub
Public Sub StopIdleMotion(CommandsObject As Object, MethodName As String)

' Starts the Main Fan

    Const PROC_NAME = "StopIdleMotion"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    
    On Error GoTo Error

    bRaiseError = False
        
    With GetController
        .CheckInterlock CStr(ocInitialize)
        .CheckInterlock CStr(Opcodes.ocCancelIdle)
        .Commands.Clear MethodName, .InvalidOpcodes(CStr(ocCancelIdle))
        .Commands.Insert ocCancelIdle, MethodName
    End With
    
    GoTo Done

Error:
    bRaiseError = True
    lRaiseErrorNumber = CCommandError(Err.Number)
    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, Err.Description)
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Err.Raise lRaiseErrorNumber, PROCESS_NAME, sRaiseErrorText
    End If

    Exit Sub

End Sub
Public Sub SetNozzleSize(CommandsObject As Object, MethodName As String, NozzleSize)

' Starts the Main Fan

    Const PROC_NAME = "StopNozzleMotion"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    Dim argSize As ArgData
    Dim lNozzleSize As NozzleSizes
    
    On Error GoTo Error

    bRaiseError = False
        
    Set argSize = StringArgData(NozzleSize)
    argSize.AddValidValues "Small", "Medium", "Large"
    CheckMethodArgs CommandsObject, MethodName, argSize
    
    If argSize.Value = "Large" Then
        lNozzleSize = nsLarge
    ElseIf argSize.Value = "Medium" Then
        lNozzleSize = nsMedium
    Else
        lNozzleSize = nsSmall
    End If
    
    With GetController
        .CheckInterlock CStr(ocInitialize)
        .CheckInterlock CStr(Opcodes.ocSetNozzleSize)
        .Commands.Clear MethodName, .InvalidOpcodes(CStr(ocSetNozzleSize))
        .Commands.Append ocSetNozzleSize, MethodName, lNozzleSize
    End With
    
    GoTo Done

Error:
    bRaiseError = True
    lRaiseErrorNumber = CCommandError(Err.Number)
    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, Err.Description)
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next
    Set argSize = Nothing
    If bRaiseError Then
        On Error GoTo 0
        Err.Raise lRaiseErrorNumber, PROCESS_NAME, sRaiseErrorText
    End If

    Exit Sub

End Sub
Public Sub StopNozzleMotion(CommandsObject As Object, MethodName As String)

' Starts the Main Fan

    Const PROC_NAME = "StopNozzleMotion"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    
    On Error GoTo Error

    bRaiseError = False
        
    With GetController
        .CheckInterlock CStr(ocInitialize)
        .CheckInterlock CStr(Opcodes.ocStopNozzleMotion)
        .Commands.Clear MethodName, .InvalidOpcodes(CStr(ocStopNozzleMotion))
        .Commands.Insert ocStopNozzleMotion, MethodName
    End With
    
    GoTo Done

Error:
    bRaiseError = True
    lRaiseErrorNumber = CCommandError(Err.Number)
    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, Err.Description)
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Err.Raise lRaiseErrorNumber, PROCESS_NAME, sRaiseErrorText
    End If

    Exit Sub

End Sub

Public Sub SetWindSpeed(CommandsObject As Object, MethodName As String, Target, Optional RampRate)

' Sets the wind speed

' Part          Description
' ----------------------------------------------------------------------
' Target        Desired target value
' RampRate      Optional ramp rate

    Const PROC_NAME = "SetWindSpeed"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    Dim argTarget As ArgData
    Dim argRampRate As ArgData
    
    On Error GoTo Error

    bRaiseError = False
    
    
    With GetController
        .CheckInterlock CStr(ocInitialize)
        Set argTarget = ValueArgData(Target, GetCC(CC_WSMin), GetCC(CC_WSMax))
        Set argRampRate = ValueArgData(RampRate, , GetCC(CC_WSRampMax), GetCC(CC_WSRampDef))
        CheckMethodArgs CommandsObject, MethodName, argTarget, argRampRate
        .CheckInterlock Array(CStr(ocSetWindSpeed), argTarget.Value)
        .Commands.Clear MethodName, .InvalidOpcodes(CStr(ocSetWindSpeed))
        .Commands.Append ocSetWindSpeed, MethodName, argTarget, argRampRate
    End With
    mobjStatuses.Reset STS_WINDSPEED_ACHIEVED, False
                
    GoTo Done

Error:
    bRaiseError = True
    lRaiseErrorNumber = CCommandError(Err.Number)
    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, Err.Description)
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next
    Set argRampRate = Nothing
    Set argTarget = Nothing
    If bRaiseError Then
        On Error GoTo 0
        Err.Raise lRaiseErrorNumber, PROCESS_NAME, sRaiseErrorText
    End If

    Exit Sub

End Sub

Public Sub SetFanSpeed(CommandsObject As Object, MethodName As String, _
        Target, Optional RampRate)

' Sets the Main Fan speed

' Part          Description
' ----------------------------------------------------------------------
' Target        Desired target value
' RampRate      Optional ramp rate

    Const PROC_NAME = "SetFanSpeed"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    Dim argTarget As ArgData
    Dim argRampRate As ArgData
    
    On Error GoTo Error

    bRaiseError = False
    
    
    With GetController
        .CheckInterlock CStr(ocInitialize)
        Set argTarget = ValueArgData(Target, GetCC(CC_FSMin), GetCC(CC_FSMax))
        Set argRampRate = ValueArgData(RampRate, , GetCC(CC_FSRampMax), GetCC(CC_FSRampDef))
        CheckMethodArgs CommandsObject, MethodName, argTarget, argRampRate
        .CheckInterlock Array(CStr(ocSetFanSpeed), argTarget.Value)
        .Commands.Clear MethodName, .InvalidOpcodes(CStr(ocSetFanSpeed))
        .Commands.Append ocSetFanSpeed, MethodName, argTarget, argRampRate
    End With
    mobjStatuses.Reset STS_FANSPEED_ACHIEVED, False
                
    GoTo Done

Error:
    bRaiseError = True
    lRaiseErrorNumber = CCommandError(Err.Number)
    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, Err.Description)
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next
    Set argRampRate = Nothing
    Set argTarget = Nothing
    If bRaiseError Then
        On Error GoTo 0
        Err.Raise lRaiseErrorNumber, PROCESS_NAME, sRaiseErrorText
    End If

    Exit Sub

End Sub

Public Sub TrackWindSpeed(CommandsObject As Object, MethodName As String, _
        TrackPar, Optional AutoIdle, Optional Headwind)

' Tracks the wind speed to a parameter (usually dyno) with an optional Headwind

' Part          Description
' ----------------------------------------------------------------------
' TrackPar      Track parameter name (string)
' HeadWind      Headwind (Headwind) to apply

    Const PROC_NAME = "TrackWindSpeed"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    Dim argTrackPar As ArgData
    Dim argHeadWind As ArgData
    Dim argAutoIdle As ArgData
    Dim objPar As Parameter
    
    On Error GoTo Error

    bRaiseError = False
    
    With GetController
        .CheckInterlock CStr(ocInitialize)
        If WindTunnel = wtCWT Then
            Set argAutoIdle = BooleanArgData(AutoIdle, GetCC(CC_DefAutoIdle))
        End If
        Set argTrackPar = StringArgData(TrackPar, GetCCString(CC_DefTrackParName))
        If argTrackPar.Value = "" Then
            argTrackPar.Value = GetCCString(CC_DefTrackParName)
        End If
        Set argHeadWind = ValueArgData(Headwind, GetCC(CC_HeadwindMin), GetCC(CC_HeadwindMax), 0)
        If WindTunnel = wtCWT Then
            CheckMethodArgs CommandsObject, MethodName, argTrackPar, argAutoIdle, argHeadWind
        Else
            CheckMethodArgs CommandsObject, MethodName, argTrackPar, argHeadWind
        End If
        Set objPar = .GetParameter(argTrackPar.Value)
        If Not IsCompatibleUnits(objPar.Units, "kph") Then
            Err.Raise tceInvalidArgs, , LoadResString(ridTrackParInvalidUnits, objPar.Name, objPar.Units, "kph")
        End If
        .CheckInterlock CStr(ocTrackWindSpeed)
        .Commands.Clear MethodName, .InvalidOpcodes(CStr(ocTrackWindSpeed))
        .Commands.Append ocTrackWindSpeed, MethodName, argTrackPar, argHeadWind, argAutoIdle
    End With
                
    GoTo Done

Error:
    bRaiseError = True
    lRaiseErrorNumber = CCommandError(Err.Number)
    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, Err.Description)
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next
    Set argTrackPar = Nothing
    Set argHeadWind = Nothing
    Set argAutoIdle = Nothing
    Set objPar = Nothing
    If bRaiseError Then
        On Error GoTo 0
        Err.Raise lRaiseErrorNumber, PROCESS_NAME, sRaiseErrorText
    End If

    Exit Sub

End Sub
Public Sub SetCFlapLeft1(CommandsObject As Object, MethodName As String, Target)

' Starts the Main Fan

    Const PROC_NAME = "SetCFlapLeft1"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    Dim argTarget As ArgData
    
    On Error GoTo Error

    bRaiseError = False
        
    Set argTarget = ValueArgData(Target, GetCC(CC_CFLeft1Min), GetCC(CC_CFLeft1Max))
    CheckMethodArgs CommandsObject, MethodName, argTarget
        
    With GetController
        .CheckInterlock CStr(ocInitialize)
        .CheckInterlock CStr(Opcodes.ocSetCFlapLeft1)
        .Commands.Clear MethodName, .InvalidOpcodes(CStr(ocSetCFlapLeft1))
        .Commands.Append ocSetCFlapLeft1, MethodName, argTarget
    End With
    
    GoTo Done

Error:
    bRaiseError = True
    lRaiseErrorNumber = CCommandError(Err.Number)
    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, Err.Description)
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next
    Set argTarget = Nothing
    If bRaiseError Then
        On Error GoTo 0
        Err.Raise lRaiseErrorNumber, PROCESS_NAME, sRaiseErrorText
    End If

    Exit Sub

End Sub
Public Sub SetCFlapLeft2(CommandsObject As Object, MethodName As String, Target)

' Starts the Main Fan

    Const PROC_NAME = "SetCFlapLeft2"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    Dim argTarget As ArgData
    
    On Error GoTo Error

    bRaiseError = False
        
    Set argTarget = ValueArgData(Target, GetCC(CC_CFLeft2Min), GetCC(CC_CFLeft2Max))
    CheckMethodArgs CommandsObject, MethodName, argTarget
        
    With GetController
        .CheckInterlock CStr(ocInitialize)
        .CheckInterlock CStr(Opcodes.ocSetCFlapLeft2)
        .Commands.Clear MethodName, .InvalidOpcodes(CStr(ocSetCFlapLeft2))
        .Commands.Append ocSetCFlapLeft2, MethodName, argTarget
    End With
    
    GoTo Done

Error:
    bRaiseError = True
    lRaiseErrorNumber = CCommandError(Err.Number)
    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, Err.Description)
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next
    Set argTarget = Nothing
    If bRaiseError Then
        On Error GoTo 0
        Err.Raise lRaiseErrorNumber, PROCESS_NAME, sRaiseErrorText
    End If

    Exit Sub

End Sub
Public Sub SetCFlapRight1(CommandsObject As Object, MethodName As String, Target)

' Starts the Main Fan

    Const PROC_NAME = "SetCFlapRight1"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    Dim argTarget As ArgData
    
    On Error GoTo Error

    bRaiseError = False
        
    Set argTarget = ValueArgData(Target, GetCC(CC_CFRight1Min), GetCC(CC_CFRight1Max))
    CheckMethodArgs CommandsObject, MethodName, argTarget
        
    With GetController
        .CheckInterlock CStr(ocInitialize)
        .CheckInterlock CStr(Opcodes.ocSetCFlapRight1)
        .Commands.Clear MethodName, .InvalidOpcodes(CStr(ocSetCFlapRight1))
        .Commands.Append ocSetCFlapRight1, MethodName, argTarget
    End With
    
    GoTo Done

Error:
    bRaiseError = True
    lRaiseErrorNumber = CCommandError(Err.Number)
    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, Err.Description)
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next
    Set argTarget = Nothing
    If bRaiseError Then
        On Error GoTo 0
        Err.Raise lRaiseErrorNumber, PROCESS_NAME, sRaiseErrorText
    End If

    Exit Sub

End Sub
Public Sub SetCFlapRight2(CommandsObject As Object, MethodName As String, Target)

' Starts the Main Fan

    Const PROC_NAME = "SetCFlapRight2"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    Dim argTarget As ArgData
    
    On Error GoTo Error

    bRaiseError = False
        
    Set argTarget = ValueArgData(Target, GetCC(CC_CFRight2Min), GetCC(CC_CFRight2Max))
    CheckMethodArgs CommandsObject, MethodName, argTarget
        
    With GetController
        .CheckInterlock CStr(ocInitialize)
        .CheckInterlock CStr(Opcodes.ocSetCFlapRight2)
        .Commands.Clear MethodName, .InvalidOpcodes(CStr(ocSetCFlapRight2))
        .Commands.Append ocSetCFlapRight2, MethodName, argTarget
    End With
    
    GoTo Done

Error:
    bRaiseError = True
    lRaiseErrorNumber = CCommandError(Err.Number)
    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, Err.Description)
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next
    Set argTarget = Nothing
    If bRaiseError Then
        On Error GoTo 0
        Err.Raise lRaiseErrorNumber, PROCESS_NAME, sRaiseErrorText
    End If

    Exit Sub

End Sub
Public Sub SetCFlapTop(CommandsObject As Object, MethodName As String, Target)

' Starts the Main Fan

    Const PROC_NAME = "SetCFlapTop"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    Dim argTarget As ArgData
    
    On Error GoTo Error

    bRaiseError = False
        
    Set argTarget = ValueArgData(Target, GetCC(CC_CFTopMin), GetCC(CC_CFTopMax))
    CheckMethodArgs CommandsObject, MethodName, argTarget
        
    With GetController
        .CheckInterlock CStr(ocInitialize)
        .CheckInterlock CStr(Opcodes.ocSetCFlapTop)
        .Commands.Clear MethodName, .InvalidOpcodes(CStr(ocSetCFlapTop))
        .Commands.Append ocSetCFlapTop, MethodName, argTarget
    End With
    
    GoTo Done

Error:
    bRaiseError = True
    lRaiseErrorNumber = CCommandError(Err.Number)
    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, Err.Description)
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next
    Set argTarget = Nothing
    If bRaiseError Then
        On Error GoTo 0
        Err.Raise lRaiseErrorNumber, PROCESS_NAME, sRaiseErrorText
    End If

    Exit Sub

End Sub
Public Sub StopCFlap(CommandsObject As Object, MethodName As String)

' Starts the Main Fan

    Const PROC_NAME = "StopCFlap"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    
    On Error GoTo Error

    bRaiseError = False
        
    With GetController
        .CheckInterlock CStr(ocInitialize)
        .CheckInterlock CStr(Opcodes.ocStopCF)
        .Commands.Clear MethodName, .InvalidOpcodes(CStr(ocStopCF))
        .Commands.Insert ocStopCF, MethodName
    End With
    
    GoTo Done

Error:
    bRaiseError = True
    lRaiseErrorNumber = CCommandError(Err.Number)
    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, Err.Description)
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Err.Raise lRaiseErrorNumber, PROCESS_NAME, sRaiseErrorText
    End If

    Exit Sub

End Sub

Public Sub ShowWindspeedStatus(CommandObject As Object, MethodName As String, Optional Sender)

' Display current fault conditions

    Const PROC_NAME = "ShowWindSpeedStatus"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    Dim sMessage As String
    Dim objEH As EH
    
    On Error GoTo Error

    bRaiseError = False
    
    On Error GoTo Error_Interlocks
    
    GetController.CheckInterlock CStr(ocInitialize)
    GetController.CheckInterlock CStr(ocShowStatus)
        
    On Error GoTo Error
    sMessage = LoadResString(stsNoFaultsDetected)
    LogMessage sMessage, mtInformation
    Select Case True
    Case IsMissing(Sender), Sender Is Nothing
    Case Else
        On Error GoTo Error_Message
        Sender.ShowMessageDialog sMessage, guiInformation, False, PROCESS_DESCRIPTION, 10#
    End Select
    
    GoTo Done
    
Error_Interlocks:
    sMessage = Err.Description
    Resume Interlocked
    
Interlocked:
    On Error GoTo Error
    Set objEH = New EH
    With objEH
        .AppError = LoadResString(stsEHTheFollowingFaultsExist, sMessage)
        Set .Component = App
    End With
    LogMessage2 objEH
    Select Case True
    Case IsMissing(Sender), Sender Is Nothing
    Case Else
        sMessage = LoadResString(stsTheFollowingFaultsExist, sMessage)
        On Error GoTo Error_Message
        Sender.ShowMessageDialog sMessage, guiWarning, False, PROCESS_DESCRIPTION, 30#
    End Select
        
    GoTo Done

Error:
    lRaiseErrorNumber = CCommandError(Err.Number)
    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, PROC_NAME, Err.Description)
    bRaiseError = True
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume Done

Error_Message:
    lRaiseErrorNumber = tceUnexpectedError
    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, PROC_NAME, LoadResString(ridErrDisplayingMessage, Err.Description))
'    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, PROC_NAME, "Error displaying message on the GUI: " & Err.Description)
    bRaiseError = True
    Resume Done
    
Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Err.Raise lRaiseErrorNumber, PROCESS_NAME, sRaiseErrorText
    End If

    Exit Sub

End Sub
Public Sub UseNozzleMethod(CommandObject As Object, MethodName As String)

    Const PROC_NAME = "UseNozzleMethod"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    
    On Error GoTo Error

    bRaiseError = False
        
    With GetController
        .CheckInterlock CStr(ocInitialize)
        .Commands.Append ocSetNozzleMethod, MethodName
    End With
    
    GoTo Done

Error:
    bRaiseError = True
    lRaiseErrorNumber = CCommandError(Err.Number)
    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, Err.Description)
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Err.Raise lRaiseErrorNumber, PROCESS_NAME, sRaiseErrorText
    End If

    Exit Sub

End Sub
Public Sub UsePlenumMethod(CommandObject As Object, MethodName As String)

    Const PROC_NAME = "UsePlenumMethod"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    
    On Error GoTo Error

    bRaiseError = False
        
    With GetController
        .CheckInterlock CStr(ocInitialize)
        .Commands.Append ocSetPlenumMethod, MethodName
    End With
    
    GoTo Done

Error:
    bRaiseError = True
    lRaiseErrorNumber = CCommandError(Err.Number)
    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, Err.Description)
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Err.Raise lRaiseErrorNumber, PROCESS_NAME, sRaiseErrorText
    End If

    Exit Sub

End Sub
Public Sub ShowIdleStatus(CommandObject As Object, MethodName As String, Optional Sender)

' Display current fault conditions

    Const PROC_NAME = "ShowIdleStatus"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    Dim sMessage As String
    Dim objEH As EH
    
    On Error GoTo Error

    bRaiseError = False
    
    On Error GoTo Error_Interlocks
    
    GetController.CheckInterlock CStr(ocInitialize)
    GetController.CheckInterlock CStr(ocShowIdleStatus)
        
    On Error GoTo Error
    sMessage = LoadResString(stsNoFaultsDetected)
    LogMessage sMessage, mtInformation
    Select Case True
    Case IsMissing(Sender), Sender Is Nothing
    Case Else
        On Error GoTo Error_Message
        Sender.ShowMessageDialog sMessage, guiInformation, False, PROCESS_DESCRIPTION, 10#
    End Select
    
    GoTo Done
    
Error_Interlocks:
    sMessage = Err.Description
    Resume Interlocked
    
Interlocked:
    On Error GoTo Error
    Set objEH = New EH
    With objEH
        .AppError = LoadResString(stsEHTheFollowingFaultsExist, sMessage)
        Set .Component = App
    End With
    LogMessage2 objEH
    Select Case True
    Case IsMissing(Sender), Sender Is Nothing
    Case Else
        sMessage = LoadResString(stsTheFollowingFaultsExist, sMessage)
        On Error GoTo Error_Message
        Sender.ShowMessageDialog sMessage, guiWarning, False, PROCESS_DESCRIPTION, 30#
    End Select
        
    GoTo Done

Error:
    lRaiseErrorNumber = CCommandError(Err.Number)
    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, PROC_NAME, Err.Description)
    bRaiseError = True
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume Done

Error_Message:
    lRaiseErrorNumber = tceUnexpectedError
    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, PROC_NAME, LoadResString(ridErrDisplayingMessage, Err.Description))
    bRaiseError = True
    Resume Done
    
Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Err.Raise lRaiseErrorNumber, PROCESS_NAME, sRaiseErrorText
    End If

    Exit Sub

End Sub


Friend Function GetStatus(StatusName As String) As Boolean
    GetStatus = mobjStatuses.GetValue(StatusName)
End Function

Private Property Get GetController() As Controller

'
    Dim objSOM As SOM
    Dim bExists As Boolean
    
    On Error GoTo Error

    If mobjController Is Nothing Then
        Set objSOM = New SOM
        Set mobjController = objSOM.GetObject(SOM_CONTROLLER)
    End If
    
    Set GetController = mobjController

    GoTo Done

Error:
    EH.Module = TypeName(Me)
    EH.Procedure = "GetController"
    Call EH_MsgBox
    Resume Done

Done:
    On Error Resume Next
    Set objSOM = Nothing
    Exit Property

End Property

Private Sub Class_Initialize()

    Set mobjStatuses = New Statuses
    mobjStatuses.UseSPG SPG_STATUSES
    
End Sub

Private Sub Class_Terminate()

    Set mobjController = Nothing
    Set mobjConstants = Nothing
    Set mobjStatuses = Nothing
    
End Sub
Private Function GetCC(CC As SPGConstants) As Single

' Returns a specific control constant

' Returns
' --------------------------------------------------------
' The return value is a ...

    Const PROC_NAME = "GetCC"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim objSPGM As SharedPropertyGroupManager
    Dim bExists As Boolean
    
    On Error GoTo Error

    bRaiseError = False

    If mobjConstants Is Nothing Then
        Set objSPGM = New SharedPropertyGroupManager
        Set mobjConstants = objSPGM.CreatePropertyGroup(SPG_CONSTANTS, 0, 0, bExists)
    End If
    GetCC = mobjConstants.PropertyByPosition(CC).Value

    GoTo Done

Error:
    sRaiseErrorText = LoadResString(stsCantGetCC, CC, Err.Description)
    bRaiseError = True
    Resume Done

Done:
    On Error Resume Next
    Set objSPGM = Nothing
    If bRaiseError Then
        On Error GoTo 0
        Err.Raise 5, TypeName(Me) & "." & PROC_NAME, sRaiseErrorText
    End If

    Exit Function

End Function

Private Function GetCCString(CC As SPGConstants) As String

' Returns a specific control constant

' Returns
' --------------------------------------------------------
' The return value is a ...

    Const PROC_NAME = "GetCCString"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim objSPGM As SharedPropertyGroupManager
    Dim bExists As Boolean
    
    On Error GoTo Error

    bRaiseError = False

    If mobjConstants Is Nothing Then
        Set objSPGM = New SharedPropertyGroupManager
        Set mobjConstants = objSPGM.CreatePropertyGroup(SPG_CONSTANTS, 0, 0, bExists)
    End If
    GetCCString = mobjConstants.PropertyByPosition(CC).Value

    GoTo Done

Error:
    sRaiseErrorText = LoadResString(stsCantGetCC, CC, Err.Description)
    bRaiseError = True
    Resume Done

Done:
    On Error Resume Next
    Set objSPGM = Nothing
    If bRaiseError Then
        On Error GoTo 0
        Err.Raise 5, TypeName(Me) & "." & PROC_NAME, sRaiseErrorText
    End If

    Exit Function

End Function

@


1.8
log
@@
text
@d843 283
@


1.7
log
@Added PUrge commands. Added control of overflow flaps (AAWT). Changed Relhumidity par name (AAWT) to make it consistent with CWT@
text
@d524 102
a625 1
        .Commands.Append ocCancelIdle, MethodName
@


1.6
log
@@
text
@d333 84
@


1.5
log
@Completed for initial testing at CWT. Added Nozzle/Plenum commands@
text
@a606 1
    
d609 4
a612 2
        Set argAutoIdle = BooleanArgData(AutoIdle, GetCC(CC_DefAutoIdle))
        Set argTrackPar = StringArgData(TrackPar, CStr(mobjConstants.PropertyByPosition(CC_DefTrackParName).Value))
d614 1
a614 1
            argTrackPar.Value = CStr(mobjConstants.PropertyByPosition(CC_DefTrackParName).Value)
d617 5
a621 1
        CheckMethodArgs CommandsObject, MethodName, argTrackPar, argAutoIdle, argHeadWind
d973 43
@


1.4
log
@Initial version for WS tuning at CWT@
text
@d736 78
@


1.3
log
@Added ResetMainFanFaults command@
text
@a110 63

Public Sub SetBlockage(CommandsObject As Object, MethodName As String, Blockage)

' Sets the wind speed

' Part          Description
' ----------------------------------------------------------------------
' Target        Desired target value
' RampRate      Optional ramp rate

    Const PROC_NAME = "SetBlockage"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    Dim argBlockage As ArgData
    Dim fBlockage As Single
    
    On Error GoTo Error

    bRaiseError = False

    If Blockage = BLOCKAGE_1 Then
        fBlockage = USE_BLOCKAGE1
    ElseIf Blockage = BLOCKAGE_2 Then
        fBlockage = USE_BLOCKAGE2
    Else
        Set argBlockage = ValueArgData(Blockage, GetCC(CC_MinBlockage), GetCC(CC_MaxBlockage))
        CheckMethodArgs CommandsObject, MethodName, argBlockage
        fBlockage = argBlockage.Value
    End If
    
    With GetController
        .CheckInterlock CStr(ocInitialize)
        .CheckInterlock CStr(ocSetBlockage) ' set mode to manual instead of rejecting cmd
        .Commands.Clear MethodName, .InvalidOpcodes(CStr(ocSetBlockage))
        .Commands.Append ocSetBlockage, MethodName, fBlockage
    End With

    GoTo Done

Error:
    bRaiseError = True
    lRaiseErrorNumber = CCommandError(Err.Number)
    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, Err.Description)
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next
    Set argBlockage = Nothing
    If bRaiseError Then
        On Error GoTo 0
        Err.Raise lRaiseErrorNumber, PROCESS_NAME, sRaiseErrorText
    End If

    Exit Sub

End Sub

@


1.2
log
@Initial version for CWT@
text
@d353 43
@


1.1
log
@Initial revision@
text
@a442 43
Public Sub ResetIdle(CommandsObject As Object, MethodName As String)

' Starts the Main Fan

    Const PROC_NAME = "ResetIdle"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    Dim lRaiseErrorNumber As Long
    
    On Error GoTo Error

    bRaiseError = False
        
    With GetController
        .CheckInterlock CStr(ocInitialize)
        .CheckInterlock CStr(Opcodes.ocResetIdle)
        .Commands.Clear MethodName, .InvalidOpcodes(CStr(ocResetIdle))
        .Commands.Append ocResetIdle, MethodName
    End With
    
    GoTo Done

Error:
    bRaiseError = True
    lRaiseErrorNumber = CCommandError(Err.Number)
    sRaiseErrorText = RaiseErrorText(lRaiseErrorNumber, MethodName, Err.Description)
    If (Err.Number And vbObjectError) <> vbObjectError Then
        EH.Module = TypeName(Me)
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Err.Raise lRaiseErrorNumber, PROCESS_NAME, sRaiseErrorText
    End If

    Exit Sub

End Sub
@
