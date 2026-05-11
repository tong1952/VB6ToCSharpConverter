head 1.2;
branch ;
access ;
symbols CD%204%2E00%2E0021:1.2 CD%204%2E00%2E0020:1.2 CD%204%2E00%2E0019:1.2
        CD%204%2E00%2E0018:1.2 CD%204%2E00%2E0017:1.2 CD%204%2E00%2E0016:1.2 CD%204%2E00%2E0015:1.2
        CD%204%2E00%2E0014:1.2 CD%204%2E00%2E0013:1.2 4%2E00%2E0034:1.2;
locks ; strict;
comment @@;


1.2
date 2019.04.17.03.31.40; author Kavanagh; state Develop;
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


1.2
log
@Initial version for CWT@
text
@VERSION 5.00
Begin VB.Form frmControl 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   1215
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   5160
   Icon            =   "Control.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrDesignTime 
      Enabled         =   0   'False
      Left            =   240
      Top             =   120
   End
   Begin VB.Label Label1 
      Caption         =   "This form is only used when running the program within VB."
      Height          =   732
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   3252
   End
End
Attribute VB_Name = "frmControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjController As Controller

Private Sub Form_Load()

    Dim I As Long
    Dim objSOM As SOM
    
    On Error GoTo Error
    
    Set objSOM = New SOM
    Set mobjController = objSOM.GetObject(SOM_CONTROLLER)
    
    GoTo Done
    
Error:
    EH.Procedure = "Form_Load"
    EH.Module = "frmControl"
    Call EH_MsgBox
    Resume Error_Exit
    
Error_Exit:
    On Error Resume Next
    Unload Me
    GoTo Done
    
Done:
    On Error Resume Next
    Set objSOM = Nothing
    Exit Sub
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    Set mobjController = Nothing
    Exit Sub
    
End Sub




Private Sub tmrDesignTime_Timer()

' This routine is used during design time to time the control loop
' execution in a manner that is easier to use within the IDE

    On Error Resume Next

    mobjController.ControlLogic

    Exit Sub

End Sub
@


1.1
log
@Initial revision@
text
@@
