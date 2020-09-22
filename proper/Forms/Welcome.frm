VERSION 5.00
Begin VB.Form Frm_welcome 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2295
   ClientLeft      =   5520
   ClientTop       =   2400
   ClientWidth     =   3015
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Welcome.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   0
      Picture         =   "Welcome.frx":24A2
      ScaleHeight     =   2235
      ScaleWidth      =   2955
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      Begin VB.Label lbl_day 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Today :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   0
         TabIndex        =   3
         Top             =   1680
         Width           =   3015
      End
      Begin VB.Label lbl_time 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Login at :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   0
         TabIndex        =   2
         Top             =   1200
         Width           =   3015
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   3000
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label lbl_user 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "User name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   0
         TabIndex        =   1
         Top             =   720
         Width           =   3015
      End
   End
End
Attribute VB_Name = "Frm_welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40

Private Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim i As Integer
Private Sub popup()
On Error Resume Next
    Picture1.Visible = True
    i = Me.Height
    Me.Height = 0
    While Me.Height < i
        Me.Height = Me.Height + 2
        Me.Top = Me.Top - 2
        DoEvents
    Wend
End Sub
Private Sub popdown()
On Error Resume Next
    i = Me.Height
    While Me.Height > 500
        Me.Height = Me.Height - 2
        Me.Top = Me.Top + 2
        DoEvents
    Wend
End Sub
Private Sub Form_Activate()
On Error Resume Next
    mdi_start.Enabled = False
    lbl_user.Caption = uname
    lbl_time.Caption = "Login at:" & Format$(Now, "hh:mm:ss AM/PM")
    lbl_day.Caption = "Today:" & Format$(Date, "dd-MMM-yy")
    Call popup
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    Sleep welcometime 'Wait for 1 Seconds
    Call popdown
mdi_start.Enabled = True
Unload Me
End Sub
Private Sub Form_Load()
On Error Resume Next
    Me.Left = Screen.Width - (Me.Width + 50)
    Me.Top = Screen.Height - 450 '450 assumed height for taskbar
    Picture1.Visible = False
End Sub

