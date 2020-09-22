VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Frm_settings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administer settings"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   Icon            =   "Frm_settings.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   7695
   Begin VB.CommandButton cmd_finedel 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      MouseIcon       =   "Frm_settings.frx":24A2
      MousePointer    =   99  'Custom
      Picture         =   "Frm_settings.frx":25F4
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Format Fine info. database"
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmd_default 
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      MouseIcon       =   "Frm_settings.frx":2B5F
      MousePointer    =   99  'Custom
      Picture         =   "Frm_settings.frx":2CB1
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Set Default settings"
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmd_apply 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      MouseIcon       =   "Frm_settings.frx":330D
      MousePointer    =   99  'Custom
      Picture         =   "Frm_settings.frx":345F
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Apply settings"
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmd_cancel 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      MouseIcon       =   "Frm_settings.frx":394E
      MousePointer    =   99  'Custom
      Picture         =   "Frm_settings.frx":3AA0
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Ok"
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmd_Change 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      MouseIcon       =   "Frm_settings.frx":4012
      MousePointer    =   99  'Custom
      Picture         =   "Frm_settings.frx":4164
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Click to modify"
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   12960
      TabIndex        =   21
      Text            =   "Text5"
      Top             =   3840
      Width           =   1935
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3015
      Left            =   120
      TabIndex        =   20
      Top             =   840
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   5318
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   538
      TabMaxWidth     =   2999
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Library"
      TabPicture(0)   =   "Frm_settings.frx":4644
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Fra_mem"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Administrator"
      TabPicture(1)   =   "Frm_settings.frx":4660
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Check1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmd_deletea"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txt_welcome"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txt_splash"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "fra_form"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "fra_pass"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label4"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label2"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lbl_time"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "lbl_wl"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "lbl_spl"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).ControlCount=   11
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Show welcome screen at startup"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -74640
         TabIndex        =   47
         Top             =   2520
         Width           =   3975
      End
      Begin VB.CommandButton cmd_deletea 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -69720
         MaskColor       =   &H8000000F&
         MouseIcon       =   "Frm_settings.frx":467C
         MousePointer    =   99  'Custom
         Picture         =   "Frm_settings.frx":47CE
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Format the Database"
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox txt_welcome 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -72240
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox txt_splash 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -72240
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Frame fra_form 
         Caption         =   "Open form"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   1455
         Left            =   -69720
         TabIndex        =   35
         Top             =   480
         Width           =   1815
         Begin VB.OptionButton opt_tl 
            Caption         =   "Top left"
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton opt_ce 
            Caption         =   "Screen center"
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   720
            Width           =   1455
         End
         Begin VB.OptionButton opt_def 
            Caption         =   "Default"
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   1080
            Width           =   975
         End
      End
      Begin VB.Frame fra_pass 
         Caption         =   "Password settings"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   1215
         Left            =   -74760
         TabIndex        =   32
         Top             =   480
         Width           =   4935
         Begin VB.TextBox txt_pass1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1920
            Locked          =   -1  'True
            PasswordChar    =   "*"
            TabIndex        =   7
            Top             =   360
            Width           =   2775
         End
         Begin VB.TextBox txt_pass2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1920
            Locked          =   -1  'True
            PasswordChar    =   "*"
            TabIndex        =   8
            Top             =   720
            Width           =   2775
         End
         Begin VB.Label lbl_p1 
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   240
            TabIndex        =   34
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label lbl_p2 
            BackStyle       =   0  'Transparent
            Caption         =   "Password confirm"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   720
            Width           =   2055
         End
      End
      Begin VB.Frame Fra_mem 
         Caption         =   "Employee"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   1935
         Left            =   3720
         TabIndex        =   27
         Top             =   600
         Width           =   3495
         Begin VB.TextBox txt_per 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   6
            Top             =   1440
            Width           =   1695
         End
         Begin VB.TextBox txt_temp 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   5
            Top             =   1080
            Width           =   1695
         End
         Begin VB.TextBox txt_new 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   4
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label lbl_per 
            BackStyle       =   0  'Transparent
            Caption         =   "Permenent"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   1470
            Width           =   1215
         End
         Begin VB.Label lbl_temp 
            BackStyle       =   0  'Transparent
            Caption         =   "Temporary"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   1110
            Width           =   1095
         End
         Begin VB.Label llbl_new 
            BackStyle       =   0  'Transparent
            Caption         =   "Newly joined"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   750
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Employeees Salary settings"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   360
            TabIndex        =   28
            Top             =   360
            Width           =   2415
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Transactions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   1935
         Left            =   240
         TabIndex        =   22
         Top             =   600
         Width           =   3375
         Begin VB.TextBox txt_maxday 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   2400
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   0
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txt_ref 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   2400
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   3
            Top             =   1440
            Width           =   855
         End
         Begin VB.TextBox txt_fine 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   2400
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   2
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox txt_maxhold 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   2400
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   1
            Top             =   720
            Width           =   855
         End
         Begin VB.Label lbl_daylimit 
            BackStyle       =   0  'Transparent
            Caption         =   "Max. days to hold book"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   390
            Width           =   2055
         End
         Begin VB.Label lbl_refcopy 
            BackStyle       =   0  'Transparent
            Caption         =   "Max.no of refrence"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   1470
            Width           =   1815
         End
         Begin VB.Label lbl_rate 
            BackStyle       =   0  'Transparent
            Caption         =   "Fine charge per day"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   1110
            Width           =   1815
         End
         Begin VB.Label lbl_maxbook 
            BackStyle       =   0  'Transparent
            Caption         =   "Max. Books hold"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   750
            Width           =   1935
         End
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Delete All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69720
         TabIndex        =   41
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "in ms"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -70560
         TabIndex        =   39
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lbl_time 
         BackStyle       =   0  'Transparent
         Caption         =   "in ms"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -70560
         TabIndex        =   38
         Top             =   1845
         Width           =   1095
      End
      Begin VB.Label lbl_wl 
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome screen stay time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -74640
         TabIndex        =   37
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label lbl_spl 
         BackStyle       =   0  'Transparent
         Caption         =   "Splash screen stay time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -74640
         TabIndex        =   36
         Top             =   1800
         Width           =   2175
      End
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Apply"
      Height          =   255
      Left            =   6480
      TabIndex        =   46
      Top             =   4605
      Width           =   1095
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      Height          =   255
      Left            =   5280
      TabIndex        =   45
      Top             =   4605
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Change"
      Height          =   255
      Left            =   4080
      TabIndex        =   44
      Top             =   4605
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Default"
      Height          =   255
      Left            =   2880
      TabIndex        =   43
      Top             =   4605
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Delete fine"
      Height          =   255
      Left            =   120
      TabIndex        =   42
      Top             =   4605
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"Frm_settings.frx":4DEA
      Height          =   615
      Left            =   720
      TabIndex        =   40
      Top             =   120
      Width           =   6855
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   120
      Top             =   120
      Width           =   600
   End
End
Attribute VB_Name = "Frm_settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temp As Integer
Dim str As String
Dim ps As String
Dim rs As ADODB.Recordset
Dim db As ADODB.Connection

Private Sub cmd_apply_Click()
On Error GoTo errlable
If (cheak = True) Then
        If (opt_tl.Value = True) Then
        temp = 1
        ElseIf (opt_ce.Value = True) Then
        temp = 2
        Else
        temp = 3
        End If
                    str = " UPDATE Custom SET "
                    str = str & "Dayslimit = " & CDbl(txt_maxday.Text) & ", "
                    str = str & "Fratepday = " & CDbl(txt_fine.Text) & ", "
                    str = str & "Maxhold = " & CDbl(txt_maxhold.Text) & ", "
                    str = str & "Pass = '" & Trim(txt_pass1.Text) & "', "
                    str = str & "Refcopy = " & CDbl(txt_ref.Text) & ", "
                    str = str & "Salnew = " & CDbl(txt_new.Text) & ", "
                    str = str & "Salper = " & CDbl(txt_per.Text) & ", "
                    str = str & "Saltemp = " & CDbl(txt_temp.Text) & ", "
                    str = str & "Splashtime = " & CDbl(txt_splash.Text) & ", "
                    str = str & "Viewe = " & temp & ", "
                    str = str & "Welcome=" & Check1.Value & ", "
                    str = str & "Welcometime = " & CDbl(txt_welcome.Text) & " WHERE Key=1"
        db.Execute str
        MsgBox "Changes are Applied.", vbInformation, "Save"
                        
                        cmd_Change.Enabled = True
                        cmd_deletea.Enabled = True
                        cmd_apply.Enabled = False
                        Label8.Caption = "OK"
                        Call locktext(True)
'Activate currently running variable with new value
                view = temp
                fratepday = CDbl(txt_fine.Text)
                dayslimit = CDbl(txt_maxday.Text)
                refcopy = CDbl(txt_ref.Text)
                maxhold = CDbl(txt_maxhold.Text)
                salnew = CDbl(txt_new.Text)
                saltemp = CDbl(txt_temp.Text)
                salper = CDbl(txt_per.Text)
                splashtime = CDbl(txt_splash.Text)
                welcometime = CDbl(txt_welcome.Text)
                Welcome = Check1.Value
                If (temp = 1) Then
                opt_tl.Value = True
                ElseIf temp = 2 Then
                opt_ce.Value = True
                Else
                opt_def.Value = True
                End If
End If
Exit Sub
errlable:
MsgBox Err.Number & Err.Description
End Sub
Private Function cheak() As Boolean
Dim flag As Boolean
flag = False
              If (txt_fine.Text = "") Then
              MsgBox "Please enter fine amount.", vbInformation, "Field missing"
              ElseIf txt_maxday.Text = "" Then
              MsgBox "Please enter max. value of days for bookhold.", vbInformation, "Field missing"
              ElseIf txt_ref.Text = "" Then
              MsgBox "Please enter no. for refcopy.", vbInformation, "Field missing"
              ElseIf txt_maxhold.Text = "" Then
              MsgBox "Please enter max no. copy hold by Member.", vbInformation, "Field missing"
              ElseIf txt_new.Text = "" Then
              MsgBox "Please enter Salary for newly joined.", vbInformation, "Field missing"
              ElseIf txt_temp.Text = "" Then
              MsgBox "Please enter Salary for temporarily working.", vbInformation, "Field missing"
              ElseIf txt_per.Text = "" Then
              MsgBox "Please enter Salary for permenently working.", vbInformation, "Field missing"
              ElseIf txt_splash.Text = "" Then
              MsgBox "Please enter splashscreen stay time in ms.", vbInformation, "Field missing"
              ElseIf txt_welcome.Text = "" Then
              MsgBox "Please enter Welcome screen stay time in ms.", vbInformation, "Field missing"
              ElseIf txt_pass1.Text = "" Then
              MsgBox "Please enter Password.", vbInformation, "Field missing"
              ElseIf txt_pass2.Text = "" Then
              MsgBox "Please enter Passwordconfirm.", vbInformation, "Field missing"
              ElseIf Not IsNumeric(txt_fine.Text) Then
              MsgBox "Fine amount mustbe Numeric.", vbInformation, "Improper value"
              ElseIf Not IsNumeric(txt_maxday.Text) Then
              MsgBox "Max. day of bookhold mustbe Numeric.", vbInformation, "Improper value"
              ElseIf Not IsNumeric(txt_ref.Text) Then
              MsgBox "Max no.of refrence copy mustbe Numeric.", vbInformation, "Improper value"
              ElseIf Not IsNumeric(txt_maxhold.Text) Then
              MsgBox "Max no.of bookhold by member mustbe Numeric.", vbInformation, "Improper value"
              ElseIf Not IsNumeric(txt_new.Text) Then
              MsgBox "Salary mustbe Numeric.", vbInformation, "Improper value"
              ElseIf Not IsNumeric(txt_temp.Text) Then
              MsgBox "Salary mustbe Numeric.", vbInformation, "Improper value"
              ElseIf Not IsNumeric(txt_per.Text) Then
              MsgBox "Salary mustbe Numeric.", vbInformation, "Improper value"
              ElseIf Not IsNumeric(txt_splash.Text) Then
              MsgBox "Splash screen stay time mustbe Numeric.", vbInformation, "Improper value"
              ElseIf Not IsNumeric(txt_welcome.Text) Then
              MsgBox "Welcome screen stay time mustbe Numeric.", vbInformation, "Improper value"
              ElseIf txt_pass2.Text <> txt_pass1.Text Then
              MsgBox "May be typing mistake,plese verify the password.", vbCritical, "Invalid password"
              Else
              flag = True
              End If
   cheak = flag
End Function
Private Sub cmd_cancel_Click()
Unload Me
End Sub

Private Sub cmd_Change_Click()
                Call locktext(False)
                cmd_Change.Enabled = False
                cmd_deletea.Enabled = False
                cmd_apply.Enabled = True
                Label8.Caption = "Cancel"
                End Sub
Private Sub locktext(val As Boolean)
                txt_fine.Locked = val
                txt_maxday.Locked = val
                txt_ref.Locked = val
                txt_maxhold.Locked = val
                txt_new.Locked = val
                txt_temp.Locked = val
                txt_per.Locked = val
                txt_splash.Locked = val
                txt_welcome.Locked = val
                txt_pass1.Locked = val
                txt_pass2.Locked = val
                Check1.Enabled = Not val
                opt_tl.Enabled = Not val
                opt_ce.Enabled = Not val
                opt_def.Enabled = Not val
                  
End Sub

Private Sub cmd_default_Click()
On Error GoTo errlable
                    str = " UPDATE Custom SET "
                    str = str & "Dayslimit = 15,"
                    str = str & "Fratepday = 1,"
                    str = str & "Maxhold = 2,"
                    str = str & "Pass = '" & Trim(ps) & "', "
                    str = str & "Refcopy = 2,"
                    str = str & "Salnew = 2000,"
                    str = str & "Salper = 4500,"
                    str = str & "Saltemp = 3000,"
                    str = str & "Splashtime = 2000,"
                    str = str & "Viewe = 3,"
                    str = str & "Welcome = True,"
                    str = str & "Welcometime =1000   WHERE Key=1"
        db.Execute str
        Call showdata
        Check1.Value = 1
        MsgBox "Default Changes are Applied.", vbInformation, "Save"
                        
                        cmd_Change.Enabled = True
                        cmd_deletea.Enabled = True
                        cmd_apply.Enabled = False
                        Label8.Caption = "OK"
                        Call locktext(True)

'Activate currently running variable with new value
                view = 3
                fratepday = CDbl(txt_fine.Text)
                dayslimit = CDbl(txt_maxday.Text)
                refcopy = CDbl(txt_ref.Text)
                maxhold = CDbl(txt_maxhold.Text)
                salnew = CDbl(txt_new.Text)
                saltemp = CDbl(txt_temp.Text)
                salper = CDbl(txt_per.Text)
                splashtime = CDbl(txt_splash.Text)
                welcometime = CDbl(txt_welcome.Text)
'                Welcome = Check1.Value
                If (view = 1) Then
                opt_tl.Value = True
                ElseIf view = 2 Then
                opt_ce.Value = True
                Else
                opt_def.Value = True
                End If

Exit Sub
errlable:
MsgBox Err.Number & Err.Description
End Sub
Private Sub cmd_deletea_Click()
On erro GoTo lable
 Beep
If MsgBox("Execution of command will delete all the information about Library database except admin. settings,Are you sure you wan't to delete Datarecord ?", vbYesNo + vbExclamation, "Confirm Delete") = vbYes Then
If MsgBox("You will never be able to retrive information back,Are you sure you wan't to delete Datarecord ?", vbYesNo + vbCritical, "Warning") = vbYes Then
   str = "DELETE FROM Book"
db.Execute str
   str = "DELETE FROM Member"
db.Execute str
   str = "DELETE FROM Issue"
db.Execute str
   str = "DELETE FROM fine"
db.Execute str
MsgBox "All entry except Administrator settings and employee information are deleted sucessfully.", vbInformation, "Database formatted"
End If
End If
Exit Sub
lable:
MsgBox Err.Number & Err.Description

End Sub
Private Sub cmd_finedel_Click()
On erro GoTo lable
 Beep
If MsgBox("Execution of command will delete all the Fine information,Are you sure you wan't to delete Datarecord ?", vbYesNo + vbExclamation, "Confirm Delete") = vbYes Then
   str = "DELETE FROM fine"
   db.Execute str
MsgBox "Fine database entry deleted successfully.", vbInformation, "Delete"
End If
Exit Sub
lable:
MsgBox Err.Number & Err.Description
End Sub
Private Sub Form_Load()
 On erro GoTo errlable
      If (view = 1) Then
     Me.Top = 50
     Me.Left = 50
     ElseIf (view = 2) Then
     Me.Top = 700
     Me.Left = (Screen.Width - Me.Width) / 2
     End If

Image1.Picture = mdi_start.ImageList1.ListImages(15).Picture
 Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database\Library.mdb;Jet OLEDB:Database Password=Library;"

  Set rs = New ADODB.Recordset
  rs.Open "select Dayslimit,Fratepday,Maxhold,Pass,Refcopy,Salnew,Salper,Saltemp,Splashtime,Viewe,Welcometime,Welcome from Custom", db, adOpenStatic, adLockOptimistic
ps = rs.Fields(3)
Label8.Caption = "OK"
cmd_apply.Enabled = False
Call showdata
Exit Sub
errlable:
MsgBox Err.Number & Err.Description
End Sub
Private Sub showdata()
           If rs.EOF = False And rs.BOF = False Then
                temp = rs.Fields(9)
                txt_fine.Text = rs.Fields(1)
                txt_maxday.Text = rs.Fields(0)
                txt_ref.Text = rs.Fields(4)
                txt_maxhold.Text = rs.Fields(2)
                txt_new.Text = rs.Fields(5)
                txt_temp.Text = rs.Fields(7)
                txt_per.Text = rs.Fields(6)
                txt_splash.Text = rs.Fields(8)
                txt_welcome.Text = rs.Fields(10)
                txt_pass1.Text = rs.Fields(3)
                txt_pass2.Text = rs.Fields(3)
           If (rs.Fields(11) = True) Then
             Check1.Value = 1
           Else
             Check1.Value = 0
           End If
           If (temp = 1) Then
           opt_tl.Value = True
           ElseIf temp = 2 Then
           opt_ce.Value = True
           Else
           opt_def.Value = True
           End If
End If
End Sub

