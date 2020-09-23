VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Ping IP Address"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H008080FF&
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   1095
   End
   Begin VB.ListBox lstResponse 
      Height          =   1620
      Left            =   1080
      TabIndex        =   3
      Top             =   720
      Width           =   3615
   End
   Begin VB.CommandButton cmdPing 
      Caption         =   "Ping"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtIPAddress 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Text            =   "google.com"
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblResponse 
      BackStyle       =   0  'Transparent
      Caption         =   "Response:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4800
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label lblAuthor 
      BackStyle       =   0  'Transparent
      Caption         =   "Eng. Usama El-Mokadem"
      BeginProperty Font 
         Name            =   "Mistral"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "http://musama.tripod.com"
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label lblIP 
      BackStyle       =   0  'Transparent
      Caption         =   "IP Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' VBPing: Visual Basic ping IP
'
' Name: VBPing
' Description: Ping IP Address
' Version: 1.10
' Date: 14 March 2007
' Last update: 15 March 2007
' Author: Eng. Usama El-Mokadem: musama@hotmail.com - ©1996-2007
'
' CONTACT INFORMATION:
' Eng. Usama El-Mokadem
' Email: musama@hotmail.com
' Web: http://musama.tripod.com
' Mobile: 0020 10 1289308
' Egypt
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Const ICMP_SUCCESS As Long = 0
Private Const WS_VERSION_REQD As Long = &H101
Private Const INADDR_NONE = &HFFFF

Private Type WSADATA
   wVersion         As Integer
   wHighVersion     As Integer
   szDescription(0 To 256) As Byte
   szSystemStatus(0 To 128) As Byte
   iMaxSockets      As Long
   iMaxUDPDG        As Long
   lpVendorInfo     As Long
End Type

Private Type IP_OPTION_INFORMATION
   Ttl              As Byte
   Tos              As Byte
   Flags            As Byte
   OptionsSize      As Byte
   OptionsData      As Long
End Type

Private Type ICMP_ECHO_REPLY
   address          As Long
   Status           As Long
   RoundTripTime    As Long
   DataSize         As Long
   Reserved         As Integer
   ptrData          As Long
   Options          As IP_OPTION_INFORMATION
   Data             As String * 250
End Type

Private Type HOSTENT
    hName           As Long
    hAliases        As Long
    hAddrType       As Integer
    hLength         As Integer
    hAddrList       As Long
End Type

Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Private Declare Function inet_addr Lib "WSOCK32.DLL" (ByVal cp As String) As Long
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpSendEcho Lib "icmp.dll" _
   (ByVal IcmpHandle As Long, _
    ByVal DestinationAddress As Long, _
    ByVal RequestData As String, _
    ByVal RequestSize As Long, _
    ByVal RequestOptions As Long, _
    ReplyBuffer As ICMP_ECHO_REPLY, _
    ByVal ReplySize As Long, _
    ByVal Timeout As Long) As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long
Private Declare Function gethostbyname Lib "ws2_32.dll" (ByVal host_name As String) As Long
Private Declare Function inet_ntoa Lib "ws2_32.dll" (ByVal inn As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)


Private Function ping(IPAddress As String, Response As ICMP_ECHO_REPLY) As Long
    Dim hIcmp As Long
    Dim lAddress As Long
    Dim PingString As String
    Dim PtrToHost As Long
    Dim addList As Long
    Dim Host As HOSTENT

    PingString = "hello"

    lAddress = inet_addr(IPAddress)
    If lAddress = INADDR_NONE Then
        PtrToHost = gethostbyname(IPAddress)
        If PtrToHost <> 0 Then
            RtlMoveMemory Host, ByVal PtrToHost, Len(Host)
            RtlMoveMemory addList, ByVal Host.hAddrList, 4
            RtlMoveMemory lAddress, ByVal addList, Host.hLength
        End If
    End If

    If (lAddress <> -1) And (lAddress <> 0) Then
        hIcmp = IcmpCreateFile()

        If hIcmp Then
            Call IcmpSendEcho(hIcmp, lAddress, PingString, Len(PingString), 0, Response, Len(Response), 1000)
            ping = Response.Status
            Call IcmpCloseHandle(hIcmp)
        Else
            ping = -1
        End If
    Else
        ping = -1
    End If
End Function

Private Sub cmdPing_Click()
    Dim WSAD As WSADATA
    Dim Response As ICMP_ECHO_REPLY

    lstResponse.AddItem ""
    lstResponse.AddItem "Starting Ping: " & txtIPAddress.Text

    If WSAStartup(WS_VERSION_REQD, WSAD) = ICMP_SUCCESS Then
        If ping(txtIPAddress.Text, Response) = 0 Then
            lstResponse.AddItem " --- Success ---"
        Else
            lstResponse.AddItem " --- Fail ---"
        End If
        WSACleanup
    Else
        lstResponse.AddItem " --- ERROR ---"
    End If
    lstResponse.AddItem " --- ----- ---"
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()

End Sub
