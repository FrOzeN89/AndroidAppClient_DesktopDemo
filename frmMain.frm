VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   ClientHeight    =   9360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   624
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   378
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock wsClient 
      Left            =   4920
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox picPhone 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   9105
      Left            =   0
      Picture         =   "frmMain.frx":8D18A
      ScaleHeight     =   607
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   317
      TabIndex        =   0
      Top             =   0
      Width           =   4755
      Begin MSComctlLib.ProgressBar progBar 
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   7560
         Visible         =   0   'False
         Width           =   4170
         _ExtentX        =   7355
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.PictureBox picImage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF00FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5670
         Left            =   360
         ScaleHeight     =   378
         ScaleMode       =   0  'User
         ScaleWidth      =   276.014
         TabIndex        =   11
         Top             =   7080
         Visible         =   0   'False
         Width           =   4170
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "&Connect"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   9
         Top             =   5880
         Width           =   1935
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   840
         MaxLength       =   32
         PasswordChar    =   "•"
         TabIndex        =   8
         Text            =   "test123"
         Top             =   5280
         Width           =   3015
      End
      Begin VB.TextBox txtUsername 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   840
         MaxLength       =   32
         TabIndex        =   7
         Text            =   "FrOzeN"
         Top             =   4320
         Width           =   3015
      End
      Begin VB.TextBox txtPort 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   840
         MaxLength       =   5
         TabIndex        =   6
         Text            =   "80"
         Top             =   3360
         Width           =   3015
      End
      Begin VB.TextBox txtServer 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   840
         TabIndex        =   5
         Top             =   2400
         Width           =   3015
      End
      Begin VB.Image imgClose 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   240
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   10
         Top             =   6600
         Width           =   3015
      End
      Begin VB.Label lblPassword 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   4
         Top             =   4920
         Width           =   1095
      End
      Begin VB.Label lblUsername 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Username:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   3
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Label lblPort 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Port:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   2
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label lblServer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Server:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   1
         Top             =   2040
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'Form Movement
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Const RGN_OR = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Public Send_Count As Integer
Public Send_Total As Integer

Public Sub SendUpdate()
    Send_Count = Send_Count + 1
    
    If Send_Count < Send_Total Then
        progBar.Value = Send_Count * 100 \ Send_Total
    Else
        Send_Count = 0
        Send_Total = 0
        progBar.Value = 0
        progBar.Visible = False
    End If
End Sub

Public Function MakeRegion(picSkin As PictureBox) As Long
    Dim X As Long, Y As Long, StartLineX As Long
    Dim FullRegion As Long, LineRegion As Long
    Dim TransparentColor As Long
    Dim InFirstRegion As Boolean
    Dim InLine As Boolean  ' Flags whether we are in a non-tranparent pixel sequence
    Dim hdc As Long
    Dim PicWidth As Long
    Dim PicHeight As Long
    
    hdc = picSkin.hdc
    PicWidth = picSkin.ScaleWidth
    PicHeight = picSkin.ScaleHeight
    
    InFirstRegion = True: InLine = False
    X = Y = StartLineX = 0
    TransparentColor = GetPixel(hdc, 0, 0)
    
    For Y = 0 To PicHeight - 1
        For X = 0 To PicWidth - 1
            
            If GetPixel(hdc, X, Y) = TransparentColor Or X = PicWidth Then
                ' We reached a transparent pixel
                If InLine Then
                    InLine = False
                    LineRegion = CreateRectRgn(StartLineX, Y, X, Y + 1)
                    
                    If InFirstRegion Then
                        FullRegion = LineRegion
                        InFirstRegion = False
                    Else
                        CombineRgn FullRegion, FullRegion, LineRegion, RGN_OR
                        DeleteObject LineRegion
                    End If
                End If
            Else
             
             
                If Not InLine Then
                    InLine = True
                    StartLineX = X
                End If
            End If
        Next
    Next
    
    MakeRegion = FullRegion
End Function

Public Sub LoadCameraScreen()

    lblServer.Visible = False
    lblPort.Visible = False
    lblUsername.Visible = False
    lblPassword.Visible = False
    lblStatus.Visible = False
    
    txtServer.Visible = False
    txtPort.Visible = False
    txtUsername.Visible = False
    txtPassword.Visible = False

    cmdConnect.Visible = False
    
    picPhone.Picture = frmMain.Picture

    Dim WindowRegion As Long
    WindowRegion = MakeRegion(picPhone)
    SetWindowRgn Me.hwnd, WindowRegion, True

    LoggedIn = True

End Sub

Public Sub LoginFailed()

    lblServer.Enabled = True
    lblPort.Enabled = True
    lblUsername.Enabled = True
    lblPassword.Enabled = True
    txtServer.Enabled = True
    txtPort.Enabled = True
    txtUsername.Enabled = True
    txtPassword.Enabled = True
    cmdConnect.Enabled = True

    lblStatus.Caption = "Login failed."
    
    wsClient.Close
    
    LoggedIn = False

End Sub

Private Sub cmdConnect_Click()

    lblServer.Enabled = False
    lblPort.Enabled = False
    lblUsername.Enabled = False
    lblPassword.Enabled = False
    txtServer.Enabled = False
    txtPort.Enabled = False
    txtUsername.Enabled = False
    txtPassword.Enabled = False
    cmdConnect.Enabled = False

    lblStatus.Caption = "Connecting.."
    
    wsClient.Connect txtServer.Text, txtPort.Text

End Sub

Private Sub Form_Load()
        
    Set Packets = New clsPackets
    LoggedIn = False
    
    picPhone.Left = 0
    picPhone.Top = 0

    picImage.Top = 120
    picImage.Left = 16

    Me.Width = picPhone.Width * Screen.TwipsPerPixelX
    Me.Height = picPhone.Height * Screen.TwipsPerPixelY
    
    Dim WindowRegion As Long
    WindowRegion = MakeRegion(picPhone)
    SetWindowRgn Me.hwnd, WindowRegion, True
    
    txtServer.Text = wsClient.LocalIP
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Packets = Nothing
End Sub

Private Sub imgClose_Click()
    Dim Result As VbMsgBoxResult
    Result = MsgBox("Are you sure you would like to quit?", vbYesNo)
    
    If Result = vbYes Then
        Unload Me
    End If
End Sub

Private Sub picPhone_KeyPress(KeyAscii As Integer)

    If (KeyAscii = 32) And (LoggedIn) Then

        Dim FilePath As String
        FilePath = TakePicture
        
        Packets.Send_File FilePath
    
    End If

End Sub

Private Sub picPhone_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim RetVal As Long
    If Button = 1 Then
        Call ReleaseCapture
        RetVal = SendMessage(frmMain.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    End If

End Sub

Private Sub wsClient_Close()
    LoggedIn = False
End Sub

Private Sub wsClient_Connect()
    lblStatus.Caption = "Connected, Logging in.."
    
    Dim Username As String, Password As String
    Username = Left$(txtUsername.Text, 32)
    Password = Left$(txtPassword.Text, 32)
    
    Packets.Send_UserInfo Username, Password
End Sub

Private Sub wsClient_DataArrival(ByVal bytesTotal As Long)

    Dim strData As String
    wsClient.GetData strData
    
    Dim intLength As Long

    Do
        intLength = GetDWORD(Mid$(strData, 2, 4))
        Packets.ParsePacket Left$(strData, intLength)
        strData = Mid$(strData, intLength + 1)
    Loop Until LenB(strData) = 0

End Sub

Private Sub wsClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    LoginFailed
    
    lblStatus.Caption = "Error occured (" & CStr(Number) & ")"
    LoggedIn = False
End Sub

Private Sub wsClient_SendComplete()
    'If File is sending, then call Send_FilePart each time Winsock confirms send completed
    If Packets.FileSending Then
        Packets.Send_FilePart
    End If
End Sub
