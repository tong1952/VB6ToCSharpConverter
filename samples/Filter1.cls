VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Filter1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mfLastOutput As Single
Private mfLastTime_s As Single

Public Function Filter1Control(CurrentValue, TimeConstant_s) As Single
Attribute Filter1Control.VB_Description = "IIR (infinite impulse response) single pole filter. Time constant is in seconds."
Attribute Filter1Control.VB_UserMemId = 0

' IIR (infinite impulse response) single pole filter. Time constant is in seconds.

    Dim fTime_s As Single
    Dim fDeltaTime_s As Single
    Dim fAlpha As Single
    Dim fOutput As Single
    
    fTime_s = Timer
    If fTime_s < mfLastTime_s Then
        fDeltaTime_s = 86400# - mfLastTime_s + fTime_s
    Else
        fDeltaTime_s = fTime_s - mfLastTime_s
    End If
    mfLastTime_s = fTime_s
    
    fAlpha = fDeltaTime_s / (TimeConstant_s + fDeltaTime_s)
    
    fOutput = fAlpha * CurrentValue + (1# - fAlpha) * mfLastOutput
    
    Filter1Control = fOutput
    mfLastOutput = fOutput
    
End Function

