head 1.2;
branch ;
access ;
symbols 4%2E00%2E0034:1.2;
locks ; strict;
comment @@;


1.2
date 2019.04.17.03.31.41; author Kavanagh; state Develop;
branches ;
next 1.1;

1.1
date 2019.04.16.23.31.56; author Kavanagh; state Develop;
branches ;
next ;


desc
@@


1.2
log
@Initial version for CWT@
text
@VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Derivative"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const Fuss = 0.0000000001   'Td fuss value

Private mfLastError As Single
Private mfLastOutput As Single

Public Function Output(ByVal FeedbackError As Single, ByVal UpdateTime As Single, ByVal Td As Single, ByVal Alpha As Single) As Single

'This is an interacting algorithm implementation of a generic
'PID controller :
'                                  -dt*alpha/Td
'       DERIV(T) = DERIV(T - 1) * e
'                               -dt*alpha/Td
'               + Td/dt * (1 - e            ) * (ERR(t) - ERR(t-1))
'
'
' where:
'
'   Kp     is the proportional gain
'   Ti     is the integral time constant
'   Td     is the derivative time constant
'   alpha  is the derivative filter multiplier
'   DERIV  is the derivative error component
'   INTEG  is the integral error component
'   ERR    is the error between the actual and setpoint signals
'   OUTPUT is the calculated change in setpoint
'   (t)    signifies the current value
'   (t-1)  signifies the previous value
'   dt     is the time period between calls to this routine

' Part          Description
' ----------------------------------------------------------------------
' Error         is the difference between the setpoint command signal
'               and the actual (feedback signal) measurement signal

' Returns
' --------------------------------------------------------
' The return value is a normalized feedback control signal as Single.

    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String
    
    Dim fTransition As Single           'transition integral
    Dim fOutput As Single
    
    On Error GoTo Error

    bRaiseError = False
    
    If Td < Fuss * UpdateTime Then
        Td = Fuss * UpdateTime   'avoid numerical error
    End If
    
    fTransition = Exp(-UpdateTime * Alpha / Td)
                                        'compute transition
    
    Output = fTransition * mfLastOutput _
                + (Td / UpdateTime) * (1# - fTransition) _
                * (FeedbackError - mfLastError)
    
    mfLastError = FeedbackError
    mfLastOutput = Output
    
    GoTo Done

Error:
    sRaiseErrorText = Err.Description
    bRaiseError = True
    Resume Done

Done:
    On Error Resume Next
    If bRaiseError Then
        On Error GoTo 0
        Call EH_RaiseError(OLEERR_METHOD_FAILED, "Derivative", "Output", sRaiseErrorText)
    End If

    Exit Function

End Function


@


1.1
log
@Initial revision@
text
@@
