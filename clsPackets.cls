VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPackets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const PKT_USERINFO  As Long = &H1
Private Const PKT_FILE      As Long = &H2

Private PacketBuffer As clsPacketBuffer
Private PacketDebuffer As clsPacketDebuffer

Private FileNum As Integer
Private FileName As String
Private blFileSending As Boolean

Public Property Get FileSending() As Boolean
    FileSending = blFileSending
End Property

Private Sub Class_Initialize()
    Set PacketBuffer = New clsPacketBuffer
    Set PacketDebuffer = New clsPacketDebuffer
    
    blFileSending = False
End Sub

Private Sub Class_Terminate()
    Set PacketBuffer = Nothing
    Set PacketDebuffer = Nothing
End Sub

Public Sub ParsePacket(ByRef Packet As String)

    Dim bytHeader As Byte
    
    PacketDebuffer.SetBuffer Packet
    bytHeader = PacketDebuffer.RemoveHeader
    
    Select Case bytHeader
        Case PKT_USERINFO
            Dim Result As Byte
            Result = PacketDebuffer.RemoveByte
            
            If Result = 1 Then
                'Login successful
                frmMain.LoadCameraScreen
            Else
                'Login failed
                frmMain.LoginFailed
            End If
            
        Case Else
            'Unknown Packet

    End Select

End Sub

Public Sub Send_UserInfo(ByRef Username As String, ByRef Password As String)
    With PacketBuffer
        .InsertNTString Username
        .InsertNTString Password
        .SendPacket PKT_USERINFO
    End With
End Sub

Public Sub Send_File(ByRef FilePath As String)
    
    Dim Buffer() As Byte, FileLength As Long
    
    FileNum = FreeFile
    FileName = StrReverse$(Left$(StrReverse$(FilePath), InStr(1, StrReverse$(FilePath), "\") - 1))
    
    With PacketBuffer
    
        Open FilePath For Binary Access Read As FileNum
        
            FileLength = LOF(FileNum)
        
            If FileLength < 1024 Then
            
                'Send complete file if less than 1024 bytes. Then close file.
            
                ReDim Buffer(FileLength - 1)
                Get FileNum, , Buffer
                
                .InsertNTString FileName
                .InsertDWORD FileLength
                .InsertData Buffer
            
                Close FileNum
                
                .SendPacket PKT_FILE
                
            Else
            
                'Send file in 1024 byte partitions, leave file open until end.
                
                blFileSending = True
                
                frmMain.Send_Total = FileLength \ 1024 + 1
                frmMain.progBar.Visible = True
                
                ReDim Buffer(1023)
                Get FileNum, , Buffer
            
                .InsertNTString FileName
                .InsertDWORD FileLength
                .InsertData Buffer
                .SendPacket PKT_FILE
                
                frmMain.SendUpdate
                
            End If
    
    End With

End Sub

Public Sub Send_FilePart()

    Dim Buffer() As Byte, BufferSize As Long
    
    If FileNum <= 0 Then Exit Sub
    
    If Loc(FileNum) >= LOF(FileNum) Then
        Close FileNum
        FileNum = 0
        Exit Sub
    End If
        
    BufferSize = lngMIN(LOF(FileNum) - Loc(FileNum), 1024)
    ReDim Buffer(BufferSize - 1)
    
    Get FileNum, , Buffer
    
    PacketBuffer.InsertNTString FileName
    PacketBuffer.InsertData Buffer
    PacketBuffer.SendPacket PKT_FILE

    frmMain.SendUpdate

End Sub
