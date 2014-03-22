VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPacketBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strBuffer As String

Public Sub InsertByte(ByRef Data As Byte)
    strBuffer = strBuffer & ChrW$(Data)
End Sub

Public Sub InsertWORD(ByRef Data As Integer)
    strBuffer = strBuffer & MakeWORD(Data)
End Sub

Public Sub InsertDWORD(ByRef Data As Long)
    strBuffer = strBuffer & MakeDWORD(Data)
End Sub

Public Sub InsertNTString(ByRef Data As String)
    strBuffer = strBuffer & Data & ChrW$(0)
End Sub

Public Sub InsertData(ByRef Data() As Byte)
    Dim strTemp As String
    strTemp = StrConv(Data, vbUnicode)
    
    strBuffer = strBuffer & strTemp
End Sub

Public Sub SendPacket(ByRef PacketId As Byte)
    If frmMain.wsClient.State <> sckClosed Then
        frmMain.wsClient.SendData ChrW$(PacketId) & MakeDWORD(Len(strBuffer) + 5) & strBuffer
    End If
    
    strBuffer = vbNullString
End Sub
