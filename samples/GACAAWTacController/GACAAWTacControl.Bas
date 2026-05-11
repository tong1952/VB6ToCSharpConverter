Attribute VB_Name = "modACControl"
Option Explicit
Option Compare Text

Private Const MOD_NAME = "modACControl"

Public Enum ResIDs
    ridNoResponseFromBKConnect = 120
    ridBKConnectCmdInvalidArg = 121
    ridBKConnectCmdError = 122
    ridBKConnectRejectedCmd = 123
    ridTCPCommsFaulted = 124
End Enum

' Object, Shared property Group, and Shared Property names
Public Const SOM_CONTROLLER = "Controller"
Public Const SPG_CONSTANTS = "Constants"
Public Const SPG_DATA = "Data"

Public Enum CCNames
    CC_xxx = 0
End Enum

Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal pDest As Any, ByVal pSource As Any, ByVal ByteLen As Long)
Public Declare Sub CopyMemory2 Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Enum AE
    InvalidActiveXLaunch = 101
    CantCreateConfigurations = 102
    InitializeFailedParsMissing = 103
    CantLoadConstants = 104
    BKConnectResponseError = 105
    CommunicationsFaulted = 106
End Enum

Public Enum Opcodes
    ocInitialize = 1
    ocSuspend = 2
    ocShowStatus = 3
    ocResetCommunications = 4
    ocStartLogging = 5
    ocStopLogging = 6
    ocSendCommand = 7
    ocSetIPAddr = 8
'    ocSendPars = 8
'    ocTakeControl = 9
'    ocReleaseControl = 10
'    ocGetCurrentControl = 11
'    ocGetCurrentStatus = 12
    ocStartRecord = 20
'    ocSetTrigger = 21
    ocStopRecord = 22
    ocShutdown = 1000
End Enum

Public Enum BKConnectStates
    gsUninitialized = 0
'    gsPreview = 1
    gsReady = 1
'    gsPreTrigger = 2
'    gsTriggerReady = 3
    gsRecording = 4
'    gsDataTransfer = 5
'    gsCancelRecord = 6
'    gsPostProcessing = 7
'    gsMeasurementFinished = 8
    gsRecordingStopped = 8
End Enum

Public Const NOFAULT = 1#
Public Const FAULT = 0#

Public Const ILCK_GENERAL As String = "General"
Public Const ILCK_COM_OK = "ComOK"
Public Const ILCK_OPERATION_OK = "OperationOK"

' Status Names
Public Const SPG_STATUSES = "Statuses"
Public Const STS_ACOUSTIC_INITIALIZED = "ACOUSTIC_INITIALIZED"
Public Const STS_ACOUSTIC_FAULTED = "ACOUSTIC_FAULTED"
    
Public Const PROCESS_NAME = "Acoustic System"
Public Const PROCESS_DESCRIPTION = "BKConnect Communications Program"
Public Const REG_PROCESSES = "Software\ReACT Technologies\Talent\4.0\Processes"
Public Const REG_ACOUSTIC = "Software\ReACT Technologies\Talent\4.0\Acoustic System"

Public Function ArgValue(Arg As Variant, Optional TargetVariableType As VbVarType = vbVariant) As Variant

    Dim vValue As Variant
    
    If TypeOf Arg Is CommandListLib.Argument Then
        vValue = Arg.Value
    Else
        vValue = Arg
    End If

    If TargetVariableType = vbVariant Then
        ArgValue = vValue
    Else
        ArgValue = Convert(vValue, TargetVariableType)
    End If
        
    Exit Function
        
End Function

Public Function ArgValueAsString(Expression As Variant) As String

    On Error Resume Next
    Select Case VarType(Expression)
    Case vbString
        ArgValueAsString = """" & CStr(Expression) & """"
    Case Else
        ArgValueAsString = CStrUS(Expression)
    End Select
    
    Exit Function
    
End Function

Public Function ByteArrayToString(InputByteArray As Variant) As String

    Dim xBytes() As Byte
    Dim i As Long
    Dim sHex As String
    
    On Error GoTo Error

    For i = LBound(InputByteArray) To UBound(InputByteArray)
        If ByteArrayToString <> "" Then
            ByteArrayToString = ByteArrayToString & " "
        End If
        sHex = Hex(InputByteArray((i)))
        If Len(sHex) < 2 Then
            sHex = "0" & sHex
        End If
        ByteArrayToString = ByteArrayToString & sHex
        If i >= 255 Then
            ByteArrayToString = ByteArrayToString & " ..."
            Exit For
        End If
    Next i
    GoTo Done
    
Error:
    Resume Done
    
Done:
    Exit Function
    
End Function

Public Function GetCmdArgs(Arguments As CommandListLib.Arguments) As String

    Dim i As Integer
    
    On Error Resume Next
    For i = 1 To Arguments.Count
        If GetCmdArgs <> "" Then
            GetCmdArgs = GetCmdArgs & ", "
        End If
        GetCmdArgs = GetCmdArgs & ArgValueAsString(Arguments(i).Value)
    Next i
    If GetCmdArgs <> "" Then
        GetCmdArgs = " " & GetCmdArgs
    End If
    
Done:
    Exit Function
    
End Function

Public Function IsSameCommand(LastCommand As Command, ReceivedMemnomic As String)

    On Error GoTo Error
    If Trim(LastCommand.TagVariant.LastCommand.CommandText) = Trim(ReceivedMemnomic) Then
        IsSameCommand = True
    End If
    GoTo Done
Error:
    Resume Done
Done:
    Exit Function
    
End Function

Public Function LoadResString(ID As Long, ParamArray ItemTextArgs() As Variant) As String

' Replace standard VB LoadResString function with this enhanced version that
' provides string substitutions and better language support.

' Look at the ID value to see which resource file to use:
'
' If the &H2000 bit is set, then get the string from the standard Talent
' set of strings

' If you are using a separate, application specific resource information only
' type DLL to store localization strings, then make sure those ID's have the
' &H1000 bit set. If this bit is set, then call the application specific
' LoadxxResString method to load the string

' If neither of these bits are set, then load the string from this
' projects resource file, but use GetResString instead of the
' built in VB function. GetResString provides better langauge support
' than LoadResString. For example, if the resource file contains German
' (Standard) and English (US) strings, and the computer is running
' German (Austrian) language, GetResString will return the German string
' whereas LoadResString will return the English string.

    If (ID And STD_TALENT_STRING_BASE) = STD_TALENT_STRING_BASE Then
        LoadResString = LoadStdTalentString(ID, ItemTextArgs)
'    ElseIf (ID And SEPARATE_RES_FILE_BASE) = SEPARATE_RES_FILE_BASE Then
'        LoadResString = LoadxxResString(ID, ItemTextArgs)
    Else 'If IsWithinIDE() Then
        LoadResString = VB.LoadResString(ID)
        LoadResString = ReplaceItemText(LoadResString, ItemTextArgs)
'    Else
'        LoadResString = GetResString(App.hInstance, ID, ItemTextArgs)
    End If
    
End Function

Public Sub LogCommMessage(ByVal MessageText As String, DataPacket, Optional MessageType As TalentProcessLib.MessageTypes = mtInformation)

'

    Dim sText As String
    Dim i As Long
    
    On Error Resume Next
    
    sText = MessageText
        
    If IsMissing(DataPacket) Then
    Else
        If Right(sText, 1) <> " " Then
            sText = sText & " "
        End If
        If VarType(DataPacket) = vbString Then
            sText = sText & """"
            For i = 1 To Len(DataPacket)
                Select Case Asc(Mid(DataPacket, i, 1))
                Case 2
                    sText = sText & "{STX}"
                Case 3
                    sText = sText & "{ETX}"
                Case 10
                    sText = sText & "{LF}"
                Case 13
                    sText = sText & "{CR}"
                Case Else
                    sText = sText & Mid(DataPacket, i, 1)
                End Select
                If i >= 200 Then
                    sText = sText & ".... " & Len(DataPacket) - 200 & " more characters"
                    Exit For
                End If
            Next i
            sText = sText & """"
        Else
            sText = sText & """<" & LenB(DataPacket) & " bytes binary data>"""
        End If
    End If

    LogMessage sText, MessageType
    
    Exit Sub

End Sub

Public Sub Main()

' This is the main startup routine

    Const PROC_NAME = "Main"
    Dim bRunning As Boolean
    
    On Error Resume Next
' Stop this instance if already running
    If App.PrevInstance Then
        If App.StartMode <> vbSModeStandalone Then
            EH.AppError = AE.InvalidActiveXLaunch
            EH.Item(1) = PROCESS_NAME
            Call EH_MsgBox
        End If
        End
    End If
    
' Create the Main Controller object
    On Error GoTo Error
        
    App.StartLogging "", vbLogAuto
    App.OleServerBusyRaiseError = True
    App.OleServerBusyTimeout = 1000
    
    Process.Title = PROCESS_NAME
    Process.EHCaption = PROCESS_NAME
    'If Not IsWithinIDE() Then
        Process.EHMode = ehLogOnly
    'Else
    '    Process.EHMode = ehDisplayAndLog
    'End If
    Process.EHMsgsToLog = ehCriticalExclamation
    Process.LogMode = logToNT
                
    GoTo Done
    
Error:
    If Err <> 364 Then
        EH.Module = MOD_NAME
        EH.Procedure = PROC_NAME
        Call EH_MsgBox
    End If
    Resume ErrorExit

ErrorExit:
    EH.AppError = ERR_CANT_START
    EH.Module = MOD_NAME
    EH.Procedure = PROC_NAME
    Call EH_MsgBox
    On Error Resume Next
    GoTo Done

Done:
    Exit Sub
    
End Sub

Public Function QTrim(Expression) As String

    On Error Resume Next
    QTrim = Expression
    QTrim = Trim(Expression)
    If Left(QTrim, 1) = """" Then
        QTrim = Mid(QTrim, 2)
    End If
    If Right(QTrim, 1) = """" Then
        QTrim = Mid(QTrim, 1, Len(QTrim) - 1)
    End If
    Exit Function
    
End Function
