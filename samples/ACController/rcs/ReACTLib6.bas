head 1.1;
branch ;
access ;
symbols CD%204%2E00%2E0021:1.1 CD%204%2E00%2E0020:1.1 CD%204%2E00%2E0019:1.1
        4%2E00%2E0001:1.1;
locks ; strict;
comment @@;


1.1
date 2020.01.20.00.44.03; author Kavanagh; state Develop;
branches ;
next ;

ext
@project Z:/Client Facilities/CAERI/Source Code/Applications/ACController/CWTCSACController.pj;
@

desc
@@


1.1
log
@Initial revision@
text
@Attribute VB_Name = "modReACTLib"
Option Explicit
Public Function Format(Expression, Optional FormatSpecifier, _
            Optional FirstDayOfWeek As VbDayOfWeek = vbSunday, _
            Optional FirstWeekOfYear As VbFirstWeekOfYear = vbFirstJan1)

' Replace the Format function
    Format = ReACTLib.Format(Expression, FormatSpecifier, FirstDayOfWeek, FirstWeekOfYear)
    
End Function
Public Function EH_MsgBox(Optional AppError, Optional Buttons As ehButtons, _
        Optional FieldName, Optional FileName, Optional Module, _
        Optional Procedure, Optional VBErr As VBA.ErrObject _
        ) As ehResponses

' Display the error message dialog, after setting some properties

    Dim lAppErrNum As Long
    Dim sText As String
    
' Save current error values, before any On Error statements
    If Err.Number <> 0 Then
        Set EH.VBErr = Err
    End If
    If Not (VBErr Is Nothing) Then
        Set EH.VBErr = VBErr
    End If
    On Error Resume Next
    
' Save the calling application (component) information
    Set EH.Component = App

' Load the local error message, if AppError is numeric and not in the range
' reserved for common errors
    If Not IsMissing(AppError) Then
        EH.AppError = AppError
    End If
    lAppErrNum = CLng(EH.AppError)
    If Err.Number = 0 Then
        Select Case lAppErrNum
        Case 1001 To 2000       ' Common error range. Pass on to error handler
        
        Case Else
            sText = LoadResString(lAppErrNum)
            If Err Then
                EH.AppError = "Error: Application error " & lAppErrNum
            Else
                EH.AppError = sText
            End If
        End Select
    End If
    
' Show the error message
    EH_MsgBox = EH.Execute(Buttons:=Buttons, FieldName:=FieldName, _
        FileName:=FileName, Module:=Module, Procedure:=Procedure)
    Err.Clear
    
    Exit Function
    
End Function
Public Function CreateObject(ClassName As String, _
            Optional ServerName As String) As Variant

' Creates an ActiveX object (replacement for VB standard routine)

' Part          Description
' ----------------------------------------------------------------------
' ClassName     Class name to create
' ServerName    The name of the server to create the object on

' Returns
' --------------------------------------------------------
' The return value is an object ...

    Set CreateObject = ReACTLib.CreateObject(ClassName, ServerName)


End Function


Public Function Now() As Date
    Now = Date + (Timer / (24# * 3600#))
End Function

@
