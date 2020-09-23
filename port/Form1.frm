VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PortScanner    (By logicport)"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   3360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Hostname or IP Address:"
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3135
      Begin VB.TextBox txtIp 
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   4440
      Width           =   1455
   End
   Begin MSWinsockLib.Winsock myInfo 
      Left            =   3480
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Scan"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3960
      Width           =   3135
   End
   Begin VB.Frame Frame3 
      Caption         =   "Open Ports:"
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   3135
      Begin MSWinsockLib.Winsock sckScan 
         Index           =   0
         Left            =   2400
         Top             =   1320
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.ListBox lstOpen 
         Height          =   1620
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Port Options:"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   3135
      Begin VB.TextBox txtTo 
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Text            =   "60000"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtFrom 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Text            =   "0"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "To"
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Top             =   360
         Width           =   255
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const iSockets As Integer = 200
Dim lPort As Long

Private Sub cmdClear_Click()
lstOpen.Clear
End Sub

Private Sub cmdStop_Click()
Dim I As Integer
cmdScan.Enabled = True
cmdStop.Enabled = False
    For I = 1 To iSockets
        Unload sckScan(I)
    Next I
End Sub


Private Sub Form_Load()
txtIp.Text = myInfo.LocalIP
End Sub

Private Sub cmdScan_Click()
cmdScan.Enabled = False
cmdStop.Enabled = True
cmdClear.Enabled = True
Dim I As Integer
   lPort = txtFrom.Text
   For I = 1 To iSockets
      Load sckScan(I)
      sckScan(I).Connect txtIp.Text, lPort
      lPort = lPort + 1
   Next I
End Sub


Private Sub NextPort(Index As Integer)
   sckScan(Index).Close
   If lPort < txtTo.Text Then
      sckScan(Index).Connect txtIp.Text, lPort
   End If
   lPort = lPort + 1
End Sub


Private Sub sckScan_Connect(Index As Integer)
   lstOpen.AddItem "Port " & sckScan(Index).RemotePort & " Open"
   NextPort Index
End Sub

Private Sub sckScan_Error(Index As Integer, ByVal Number As Integer, _
Description As String, ByVal Scode As Long, ByVal Source As String, ByVal _
HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
   NextPort Index
End Sub

