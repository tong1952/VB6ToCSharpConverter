VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Person"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Option Explicit

' Private backing fields
Private mName As String
Private mAge As Integer
Private mEmail As String

' Events
Public Event NameChanged(ByVal OldName As String, ByVal NewName As String)
Public Event AgeChanged(ByVal NewAge As Integer)

' Properties
Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(ByVal value As String)
    Dim oldName As String
    oldName = mName
    If Len(value) = 0 Then
        Err.Raise 5, "Person", "Name cannot be empty"
    End If
    mName = value
    RaiseEvent NameChanged(oldName, mName)
End Property

Public Property Get Age() As Integer
    Age = mAge
End Property

Public Property Let Age(ByVal value As Integer)
    If value < 0 Or value > 150 Then
        Err.Raise 5, "Person", "Invalid age"
    End If
    mAge = value
    RaiseEvent AgeChanged(mAge)
End Property

Public Property Get Email() As String
    Email = mEmail
End Property

Public Property Let Email(ByVal value As String)
    mEmail = value
End Property

' Methods
Public Function Greeting() As String
    Greeting = "Hello, my name is " & mName & " and I am " & CStr(mAge) & " years old."
End Function

Public Sub CopyFrom(ByRef other As Person)
    mName = other.Name
    mAge = other.Age
    mEmail = other.Email
End Sub

Public Function IsAdult() As Boolean
    IsAdult = (mAge >= 18)
End Function

Public Function ToString() As String
    Dim sb As String
    sb = "Person{Name=" & mName
    sb = sb & ", Age=" & CStr(mAge)
    sb = sb & ", Email=" & mEmail & "}"
    ToString = sb
End Function
