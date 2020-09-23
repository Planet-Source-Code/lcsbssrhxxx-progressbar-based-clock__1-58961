VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmClock 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5190
   Icon            =   "frmClock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrCheck 
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame3 
      Caption         =   "Seconds"
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   4935
      Begin MSComctlLib.ProgressBar pgSec 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         Max             =   60
         Scrolling       =   1
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Minutes"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4935
      Begin MSComctlLib.ProgressBar pgMin 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         Max             =   60
         Scrolling       =   1
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Hours"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      Begin MSComctlLib.ProgressBar pgHr 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         Max             =   24
         Scrolling       =   1
      End
   End
   Begin VB.PictureBox i1 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   4200
      Picture         =   "frmClock.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox i2 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   4200
      Picture         =   "frmClock.frx":0884
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   540
   End
End
Attribute VB_Name = "frmClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'[Type NotifyIconData For Tray Icon]
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
'[Tray Constants]
Const NIM_ADD = &H0 'Add to Tray
Const NIM_MODIFY = &H1 'Modify Details
Const NIM_DELETE = &H2 'Remove From Tray
Const NIF_MESSAGE = &H1 'Message
Const NIF_ICON = &H2 'Icon
Const NIF_TIP = &H4 'TooTipText
Const WM_MOUSEMOVE = &H200 'On Mousemove
Const WM_LBUTTONDBLCLK = &H203 'Left Double Click
Const WM_RBUTTONDOWN = &H204 'Right Button Down
Const WM_RBUTTONUP = &H205 'Right Button Up
Const WM_RBUTTONDBLCLK = &H206 'Right Double Click

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Dim InTray As Boolean

Private Sub Form_Load()
Dim strSec As String
Dim strMin As String
Dim strHr As String
    strSec = Second(Now)
    strMin = Minute(Now)
    strHr = Hour(Now)
    pgHr.Value = strHr
    pgMin.Value = strMin
    pgSec.Value = strSec
    frmClock.Caption = Time
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case InTray
        Case True
            If Button = 1 Then
                Me.WindowState = vbNormal
                Me.Show
            End If
        Case False
            Exit Sub
    End Select
End Sub
Private Sub tmrCheck_Timer()
Dim strSec As String
Dim strMin As String
Dim strHr As String
Dim TrayIco As NOTIFYICONDATA
    strSec = Second(Now)
    strMin = Minute(Now)
    strHr = Hour(Now)
    pgHr.Value = strHr
    pgMin.Value = strMin
    pgSec.Value = strSec
    frmClock.Caption = Time
    If frmClock.Icon = i1.Picture Then
        frmClock.Icon = i2.Picture
    Else
        frmClock.Icon = i1.Picture
    End If
    
    If Me.WindowState = 1 Then
        InTray = True
        Me.Hide
        With TrayIco
            .cbSize = Len(TrayIco)
            .hwnd = Me.hwnd
            .uId = vbNull
            .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
            .uCallBackMessage = WM_MOUSEMOVE
            .hIcon = i1.Picture
            .szTip = Time & vbNullChar
        End With
        Shell_NotifyIcon NIM_ADD, TrayIco
    Else
        InTray = False
        Shell_NotifyIcon NIM_DELETE, TrayIco
    End If
End Sub
