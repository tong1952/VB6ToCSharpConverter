head 1.5;
branch ;
access ;
symbols CD%204%2E00%2E0021:1.5 CD%204%2E00%2E0020:1.5 CD%204%2E00%2E0019:1.5
        CD%204%2E00%2E0018:1.5 CD%204%2E00%2E0017:1.5 CD%204%2E00%2E0016:1.5 CD%204%2E00%2E0015:1.5
        CD%204%2E00%2E0014:1.5 CD%204%2E00%2E0013:1.5 4%2E00%2E0062:1.5 4%2E00%2E0049:1.4
        4%2E00%2E0039:1.3 4%2E00%2E0034:1.2;
locks ; strict;
comment @@;


1.5
date 2019.09.11.21.31.25; author Kavanagh; state Develop;
branches ;
next 1.4;

1.4
date 2019.09.06.23.48.17; author Kavanagh; state Develop;
branches ;
next 1.3;

1.3
date 2019.04.26.02.44.36; author Kavanagh; state Develop;
branches ;
next 1.2;

1.2
date 2019.04.17.03.31.42; author Kavanagh; state Develop;
branches ;
next 1.1;

1.1
date 2019.04.16.23.31.53; author Kavanagh; state Develop;
branches ;
next ;

ext
@project Z:/Client Facilities/CAERI/Source Code/Applications/WSController/CWTCSWSController.pj;
@

desc
@@


1.5
log
@Added PUrge commands. Added control of overflow flaps (AAWT). Changed Relhumidity par name (AAWT) to make it consistent with CWT@
text
@VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WindspeedCommandsAAWT"
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
Public Sub TrackWindSpeed(TrackParameter, Optional Headwind)
    mobjCommands.TrackWindSpeed Me, "TrackWindSpeed", TrackParameter, , Headwind
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
Public Sub ResetMainFanFaults()
    mobjCommands.ResetMainFanFaults Me, "ResetMainFanFaults"
End Sub

Public Sub SetFanSpeed(Target, Optional RampRate)
    mobjCommands.SetFanSpeed Me, "SetFanSpeed", Target, RampRate
End Sub
Public Sub UseNozzleMethod()
    mobjCommands.UseNozzleMethod Me, "UseNozzleMethod"
End Sub
Public Sub UsePlenumMethod()
    mobjCommands.UsePlenumMethod Me, "UsePlenumMethod"
End Sub

Public Sub ShowWindspeedStatus(Optional Sender)
    mobjCommands.ShowWindspeedStatus Me, "ShowWindSpeedStatus", Sender
End Sub
Public Sub StartPurge()
    mobjCommands.StartPurge Me, "StartPurge"
End Sub

Public Sub StopPurge()
    mobjCommands.StopPurge Me, "StopPurge"
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
@


1.4
log
@@
text
@d58 7
@


1.3
log
@Added ResetMainFanFaults command@
text
@d22 2
a23 2
Public Sub TrackWindSpeed(TrackParameter, Optional AutoIdle, Optional Headwind)
    mobjCommands.TrackWindSpeed Me, "TrackWindSpeed", TrackParameter, AutoIdle, Headwind
d47 6
@


1.2
log
@Initial version for CWT@
text
@d41 4
@


1.1
log
@Initial revision@
text
@@
