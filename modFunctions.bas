Attribute VB_Name = "modFunctions"
Option Explicit

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal numBytes As Long)

Public Packets As clsPackets
Public LoggedIn As Boolean

Public Function MakeWORD(ByRef Value As Integer) As String
    Dim Result As String * 2
    CopyMemory ByVal Result, Value, 2
    MakeWORD = Result
End Function

Public Function MakeDWORD(ByRef Value As Long) As String
    Dim Result As String * 4
    CopyMemory ByVal Result, Value, 4
    MakeDWORD = Result
End Function

Public Function GetWORD(ByRef Data As String) As Integer
    Dim intReturn As Integer
    CopyMemory intReturn, ByVal Data, 2
    GetWORD = intReturn
End Function

Public Function GetDWORD(ByRef Data As String) As Long
    Dim lngReturn As Long
    CopyMemory lngReturn, ByVal Data, 4
    GetDWORD = lngReturn
End Function

Public Function lngMIN(ByVal L1 As Long, ByVal L2 As Long) As Long
    If L1 < L2 Then
        lngMIN = L1
    Else
        lngMIN = L2
    End If
End Function
