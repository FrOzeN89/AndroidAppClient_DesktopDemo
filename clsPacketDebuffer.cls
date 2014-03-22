VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPacketDebuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strBuffer As String

Public Sub SetBuffer(ByRef Buffer As String)
    strBuffer = Buffer
End Sub

Public Function RemoveHeader() As Byte
    RemoveHeader = AscW(Left$(strBuffer, 1))
    strBuffer = Mid$(strBuffer, 6)
End Function

Public Function RemoveByte() As Byte
    RemoveByte = AscW(Left$(strBuffer, 1))
    strBuffer = Mid$(strBuffer, 2)
End Function

Public Function RemoveWORD() As Long
    RemoveWORD = GetWORD(Mid$(strBuffer, 1, 2))
    strBuffer = Mid$(strBuffer, 3)
End Function

Public Function RemoveDWORD() As Long
    RemoveDWORD = GetDWORD(Mid$(strBuffer, 1, 4))
    strBuffer = Mid$(strBuffer, 5)
End Function

Public Function RemoveNTString() As String
    Dim intPosition As Integer
    
    If Not InStr(1, strBuffer, ChrW$(0)) > 1 Then
        Exit Function
    Else
        intPosition = InStr(1, strBuffer, ChrW$(0))
        RemoveNTString = Mid$(strBuffer, 1, intPosition - 1)
        strBuffer = Mid$(strBuffer, intPosition + 1)
    End If
End Function


