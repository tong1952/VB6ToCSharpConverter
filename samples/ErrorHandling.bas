Attribute VB_Name = "ErrorHandling"

' Pattern 1: On Error GoTo in a Sub with pre-statements
Public Sub WithPreCode(ByVal x As Integer)
    Dim result As String
    result = "starting"
    On Error GoTo Handler
    result = CStr(100 \ x)
    MsgBox result
    Exit Sub
Handler:
    MsgBox "Error: " & Err.Description
End Sub

' Pattern 2: On Error GoTo in a Function (return value preserved)
Public Function SafeDivide(ByVal a As Double, ByVal b As Double) As Double
    On Error GoTo DivErr
    SafeDivide = a / b
    Exit Function
DivErr:
    MsgBox "Division error #" & CStr(Err.Number) & ": " & Err.Description
    SafeDivide = 0
End Function

' Pattern 3: On Error GoTo 0 (disable, no handler emitted)
Public Sub NoHandler()
    On Error GoTo 0
    Dim x As Integer
    x = 10
End Sub

' Pattern 4: On Error Resume Next (remaining stmts wrapped in try/catch)
Public Sub BestEffort(ByVal path As String)
    On Error Resume Next
    MsgBox "Trying..."
    MsgBox "Done"
End Sub

' Pattern 5: Err.Raise translates to throw, Err.Clear to comment
Public Sub ThrowCustomError(ByVal msg As String)
    On Error GoTo ErrH
    Err.Raise 1001, "MyApp", msg
    Exit Sub
ErrH:
    MsgBox "Caught: " & Err.Description
    Err.Clear
End Sub

' Pattern 6: Multiple operations, Resume Next in handler (suppressed)
Public Sub MultiOp(ByVal n As Integer)
    Dim total As Long
    total = 0
    On Error GoTo OpsErr
    total = total + n
    total = total * 2
    Exit Sub
OpsErr:
    MsgBox "Op failed: " & Err.Description
    Resume Next
End Sub
