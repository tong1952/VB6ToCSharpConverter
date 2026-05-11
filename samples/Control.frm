VERSION 5.00
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
      Left            =   3000
      Top             =   0
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "Minimize"
      Default         =   -1  'True
      Height          =   372
      Left            =   3960
      TabIndex        =   0
      Top             =   600
      Width           =   972
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   372
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Width           =   972
   End
   Begin VB.Label lblProjectDesc 
      Caption         =   "Label1"
      Height          =   432
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   2532
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mobjController As Controller

Private Sub cmdHide_Click()

    On Error Resume Next
    SetVisible False
    Exit Sub
    
End Sub

Private Sub cmdQuit_Click()

    On Error GoTo Error
    If MsgBox("Are you sure you want to stop this control program?", vbQuestion + vbYesNo) = vbNo Then
        GoTo Done
    End If
    
    Unload Me
'    End
    
Error:
    Resume Done
    
Done:
    Exit Sub
    
End Sub

Private Sub Form_Load()

    Const PROC_NAME = "Form_Load"
    Dim objSOM As SOM
    
    On Error GoTo Error
    
    Set objSOM = New SOM
    Set mobjController = objSOM.GetObject(SOM_CONTROLLER)
    
    Me.Caption = PROCESS_NAME
    lblProjectDesc = App.FileDescription
    GoTo Done
    
Error:
    EH.Procedure = PROC_NAME
    EH.Module = TypeName(Me)
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    On Error Resume Next
    Select Case UnloadMode
    Case vbFormControlMenu
        SetVisible False
        Cancel = True
    End Select
    Exit Sub
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    Select Case Me.WindowState
    Case vbMinimized
        Call cmdHide_Click
        Me.WindowState = vbNormal
    End Select
    Exit Sub
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    Set mobjController = Nothing
    Exit Sub
    
End Sub

Public Sub SetVisible(Visible As Boolean)
Attribute SetVisible.VB_Description = "Sets the visibility"

' Sets the visibility

' Part          Description
' ----------------------------------------------------------------------
' Visible       Desired visibilty

    Const PROC_NAME = "SetVisible"
    Dim bRaiseError As Boolean
    Dim sRaiseErrorText As String

    On Error GoTo Error

    bRaiseError = False

    Select Case Visible
    Case True
        Me.WindowState = vbNormal
        Me.Show
        Process.EHMode = ehDisplayAndLog
    
    Case False
        Me.Hide
        Process.EHMode = ehLogOnly
        
    End Select

    GoTo Done

Error:
    sRaiseErrorText = Err.Description
    bRaiseError = True
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
        Call EH_RaiseError(OLEERR_METHOD_FAILED, TypeName(Me), PROC_NAME, sRaiseErrorText)
    End If

    Exit Sub

End Sub


Private Sub tmrDesignTime_Timer()
' This routine is used during design time to time the control loop
' execution in a manner that is easier to use within the IDE

    On Error Resume Next

    mobjController.ControlLogic

    Exit Sub

End Sub
