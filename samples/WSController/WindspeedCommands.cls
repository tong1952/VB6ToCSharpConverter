VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WindspeedCommands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim mobjCommands As WsCommands

Implements TalentCommandClassEx1

Public Sub SetWindSpeed(Target, Optional RampRate)
    mobjCommands.SetWindSpeed Me, "SetWindSpeed", Target, RampRate
End Sub
Public Sub TrackWindSpeed(TrackParameter, Optional AutoIdle, Optional Headwind)
    mobjCommands.TrackWindSpeed Me, "TrackWindSpeed", TrackParameter, AutoIdle, Headwind
End Sub

Public Sub StartMainFanAux()
    mobjCommands.StartMainFanAux Me, "StartMainFanAux"
End Sub

Public Sub StopMainFanAux()
    mobjCommands.StopMainFanAux Me, "StopMainFanAux"
End Sub

Public Sub StartMainFan()
    mobjCommands.StartMainFan Me, "StartMainFan"
End Sub

Public Sub StopMainFan()
    mobjCommands.StopMainFan Me, "StopMainFan"
End Sub
Public Sub SetFanSpeed(Target, Optional RampRate)
    mobjCommands.SetFanSpeed Me, "SetFanSpeed", Target, RampRate
End Sub

Public Sub SetBlockage(Blockage)
    mobjCommands.SetBlockage Me, "SetBlockage", Blockage
End Sub

Public Sub StartIdle()
    mobjCommands.StartIdle Me, "StartIdle"
End Sub
Public Sub StopIdle()
    mobjCommands.StopIdle Me, "StopIdle"
End Sub
Public Sub ResetIdle()
    mobjCommands.ResetIdle Me, "ResetIdle"
End Sub
Public Sub StopIdleMotion()
    mobjCommands.StopIdleMotion Me, "StopIdleMotion"
End Sub

Public Sub ShowWindspeedStatus(Optional Sender)
    mobjCommands.ShowWindspeedStatus Me, "ShowWindSpeedStatus", Sender
End Sub
Public Sub ShowIdleStatus(Optional Sender)
    mobjCommands.ShowIdleStatus Me, "ShowIdleStatus", Sender
End Sub

Public Sub InitializeWindSpeed()
    mobjCommands.InitializeWindSpeed Me, "InitializeWindSpeed"
End Sub

Public Sub SuspendWindSpeed()
    mobjCommands.SuspendWindSpeed Me, "SuspendWindSpeed"
End Sub

' Status Properties
Public Property Get WindspeedAchieved() As Boolean
    WindspeedAchieved = mobjCommands.GetStatus(STS_WINDSPEED_ACHIEVED)
End Property
Public Property Get FanSpeedAchieved() As Boolean
    FanSpeedAchieved = mobjCommands.GetStatus(STS_FANSPEED_ACHIEVED)
End Property
Public Property Get IdleStarted() As Boolean
    IdleStarted = mobjCommands.GetStatus(STS_IDLE_STARTED)
End Property
Public Property Get IdleStopped() As Boolean
    IdleStopped = mobjCommands.GetStatus(STS_IDLE_STOPPED)
End Property
Public Property Get MainFanAuxStarted() As Boolean
    MainFanAuxStarted = mobjCommands.GetStatus(STS_MAINFAN_AUX_ON)
End Property
Public Property Get MainFanAuxStopped() As Boolean
    MainFanAuxStopped = Not mobjCommands.GetStatus(STS_MAINFAN_AUX_ON)
End Property
Public Property Get MainFanStarted() As Boolean
    MainFanStarted = mobjCommands.GetStatus(STS_MAINFAN_ON)
End Property
Public Property Get MainFanStopped() As Boolean
    MainFanStopped = Not mobjCommands.GetStatus(STS_MAINFAN_ON)
End Property
Public Property Get WindspeedInitialized() As Boolean
    WindspeedInitialized = mobjCommands.GetStatus(STS_WINDSPEED_INITIALIZED)
End Property

Private Sub Class_Initialize()

'

    Const PROC_NAME = "Class_Initialize"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String

    On Error GoTo Error

    bRaiseError = False

    Set mobjCommands = New WsCommands
    
    GoTo Done

Error:
    sRaiseErrorText = Err.Description
    bRaiseError = True
    EH.Module = TypeName(Me)
    EH.Procedure = PROC_NAME
    Call EH_MsgBox
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        EH_RaiseError OLEERR_METHOD_FAILED, TypeName(Me), PROC_NAME, sRaiseErrorText
    End If

    Exit Sub

   
End Sub

Private Sub Class_Terminate()

    Set mobjCommands = Nothing
    
End Sub

Private Property Get TalentCommandClassEx1_Caption() As Variant

    TalentCommandClassEx1_Caption = LoadResString(ridCommandsClassCaption)
    
End Property
