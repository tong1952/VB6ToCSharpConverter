Attribute VB_Name = "modCTRCSimulator"
Option Explicit
Option Compare Text

Private Const MOD_NAME = "modCTRCSimulator"

Public Enum ResIDs
    ridCantCreateObject = 101
    ridUnrecognizedOpcode = 102
    ridControllerShuttingDown = 103
    ridRunning = 104
    ridRunningAt = 105
    ridErrDisplayingMessage = 106
End Enum
'
' Object, Shared property Group, and Shared Property names
Public Const SOM_CONTROLLER = "Controller"
Public Const SPG_CONSTANTS = "Constants"
Public Const SPG_DATA = "Data"

Public Enum CCNames
    ccUpdateTime = 0
    ccTCPIPAddr = 1
    ccTCPIPPort = 2
    ccUDPPort = 3
    ccUDPLocalPort = 4
    ccUDPAddr = 5
End Enum

'Public Enum SharedDataItems
'    sdDyMode = 0
'    sdThrMode = 1
'    sdDyConfig = 2
'End Enum

Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
            (ByVal pDest As Any, ByVal pSource As Any, ByVal ByteLen As Long)

Public Enum AE
    InvalidActiveXLaunch = 101
'    CantCreateConfigurations = 102
'    InitializeFailedParsMissing = 103
    CantLoadConstants = 104
'    DataPacketError = 105
'    CommunicationsFaulted = 158
End Enum

Public Enum Opcodes
    ocInitialize = 1
    ocSuspend = 2
    ocShutdown = 1000
End Enum

Public Const NOFAULT = 1#
Public Const FAULT = 0#
Public Const UseDefaultValue = -10000000000#

'' Status Names
Public Const SPG_STATUSES = "Statuses"
Public Const STS_SIM_INITIALIZED = "SIM_INITIALIZED"
Public Const PROCESS_NAME = "Sim"
Public Const PROCESS_DESCRIPTION = "ETL2 Simulator"
Public Const REG_PROCESSES = "Software\ReACT Technologies\Talent\4.0\Processes"

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
