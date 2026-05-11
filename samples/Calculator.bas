Option Explicit

' Sample VB6 module — covers most language constructs

Public Const APP_NAME As String = "VB6 Calc"
Private mRunCount As Long

' Simple arithmetic
Public Function Add(ByVal a As Double, ByVal b As Double) As Double
    Add = a + b
End Function

Public Function Divide(ByVal a As Double, ByVal b As Double) As Double
    On Error GoTo ErrHandler
    If b = 0 Then
        MsgBox "Division by zero!"
        Divide = 0
        Exit Function
    End If
    Divide = a / b
    Exit Function
ErrHandler:
    MsgBox "Error: " & Err.Description
    Divide = 0
End Function

' Demonstrates For, While, Select Case
Public Sub PrintTable(ByVal n As Integer)
    Dim i As Integer
    Dim result As Double

    For i = 1 To n
        result = i * i
        Select Case result
            Case 1
                Debug.Print "One"
            Case 4, 9
                Debug.Print "Small square: " & CStr(result)
            Case Is > 50
                Debug.Print "Large: " & CStr(result)
            Case Else
                Debug.Print "Square: " & CStr(result)
        End Select
    Next i
End Sub

' Array usage and For Each
Public Function SumArray(ByVal arr() As Double) As Double
    Dim total As Double
    Dim item As Variant
    total = 0
    For Each item In arr
        total = total + CDbl(item)
    Next item
    SumArray = total
End Function

' String operations
Public Function BuildReport(ByVal name As String, ByVal score As Integer) As String
    Dim report As String
    report = "Name: " & UCase(name)
    report = report & " | Score: " & CStr(score)
    report = report & " | Grade: "
    If score >= 90 Then
        report = report & "A"
    ElseIf score >= 80 Then
        report = report & "B"
    ElseIf score >= 70 Then
        report = report & "C"
    Else
        report = report & "F"
    End If
    BuildReport = report
End Function

' Do loop
Public Function Factorial(ByVal n As Long) As Long
    Dim result As Long
    Dim i As Long
    result = 1
    i = n
    Do While i > 1
        result = result * i
        i = i - 1
    Loop
    Factorial = result
End Function

' Static variable
Public Function IncrementCount() As Long
    Static count As Long
    count = count + 1
    mRunCount = count
    IncrementCount = count
End Function
