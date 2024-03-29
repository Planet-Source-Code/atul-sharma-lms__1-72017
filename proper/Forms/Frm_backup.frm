VERSION 5.00
Begin VB.Form Frm_backup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backup Records"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   Icon            =   "Frm_backup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   7110
   Begin VB.Frame fra_back 
      Caption         =   "Current backup"
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
      Height          =   3015
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   6855
      Begin VB.FileListBox File1 
         Height          =   1845
         Left            =   240
         Pattern         =   "*.mdb"
         TabIndex        =   13
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton cmdSave 
         Height          =   735
         Left            =   3960
         MouseIcon       =   "Frm_backup.frx":24A2
         MousePointer    =   99  'Custom
         Picture         =   "Frm_backup.frx":25F4
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Backup"
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Height          =   735
         Left            =   5400
         MouseIcon       =   "Frm_backup.frx":2BAD
         MousePointer    =   99  'Custom
         Picture         =   "Frm_backup.frx":2CFF
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Cancel"
         Top             =   1440
         Width           =   1095
      End
      Begin VB.DirListBox Dir1 
         Height          =   1890
         Left            =   480
         TabIndex        =   1
         ToolTipText     =   "Select folder"
         Top             =   840
         Width           =   3015
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   480
         TabIndex        =   0
         ToolTipText     =   "Select diskdrive"
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "See the details of last backup"
         Height          =   255
         Left            =   3960
         TabIndex        =   16
         Top             =   2520
         Width           =   2535
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
         Height          =   255
         Left            =   5400
         TabIndex        =   15
         Top             =   2220
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Backup"
         Height          =   255
         Left            =   3960
         TabIndex        =   14
         Top             =   2220
         Width           =   1095
      End
      Begin VB.Label lbl_Status 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   3840
         TabIndex        =   9
         Top             =   480
         Width           =   2775
      End
   End
   Begin VB.Frame fra_last 
      Caption         =   "Last backup"
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
      Height          =   1695
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   6855
      Begin VB.Label lbl_path 
         BackStyle       =   0  'Transparent
         Caption         =   "Last backup path"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   495
         Left            =   1080
         TabIndex        =   12
         Top             =   1080
         Width           =   5655
      End
      Begin VB.Label lbl_lasttime 
         BackStyle       =   0  'Transparent
         Caption         =   "Last backup time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   255
         Left            =   1080
         TabIndex        =   11
         Top             =   720
         Width           =   5535
      End
      Begin VB.Label lbl_lastdate 
         BackStyle       =   0  'Transparent
         Caption         =   "Last backup date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   255
         Left            =   1080
         TabIndex        =   10
         Top             =   360
         Width           =   5535
      End
      Begin VB.Label lbl_apath 
         BackStyle       =   0  'Transparent
         Caption         =   "Path   :"
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
         TabIndex        =   8
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lbl_time 
         BackStyle       =   0  'Transparent
         Caption         =   "Time   :"
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
         TabIndex        =   7
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lbl_Date 
         BackStyle       =   0  'Transparent
         Caption         =   "Date   :"
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
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Make backup for your refrence before deleting fine informations at regular intervals and secure your informations"
      Height          =   495
      Left            =   720
      TabIndex        =   17
      Top             =   120
      Width           =   6255
   End
   Begin VB.Image Image1 
      Height          =   465
      Left            =   120
      Top             =   120
      Width           =   600
   End
End
Attribute VB_Name = "Frm_backup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'add ref microsoft script library for file system object
Dim Fsys As New FileSystemObject
Dim bckupFile As File
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdsave_click()
On Error Resume Next
    cmdSave.Enabled = False
    'lbl_Status.Caption = "Please Wait, Backup in Progress..."
    sbStatusBar.Panels(1) = "Please Wait, Backup in Progress..."
    Dim destination As String
    Dim Source As String
    Dim currDate, currTime As String
    currDate = Format$(Now, "dd, mmm, yyyy")
    currTime = Format$(Now, "hh:mm:ss AM/PM")
    
    destination = File1.Path & "\" & "Libraryback.mdb"
    Source = App.Path & "\Database\Library.mdb"
   
    Set bckupFile = Fsys.GetFile(finalpath)
    bckupFile.Attributes = Compressed
    Fsys.CopyFile Source, destination, True
    SaveSetting App.Title, "Settings", "BackupPath", destination
    SaveSetting App.Title, "Settings", "BackupDate", currDate
    SaveSetting App.Title, "Settings", "BackupTime", currTime
    lbl_status.Caption = "Backup sucessfull"
    cmdSave.Enabled = True
    MsgBox "All data BackUp Succcessfully on disk", vbInformation, "Backup"
    Unload Me
End Sub
Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub
Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Form_Load()
    Dim lastPath As String
    Dim lastDate As String
    Dim lastTime As String
    File1.Visible = False
Image1.Picture = mdi_start.ImageList1.ListImages(17).Picture
      If (view = 1) Then
     Me.Top = 50
     Me.Left = 50
     ElseIf (view = 2) Then
     Me.Top = 700
     Me.Left = (Screen.Width - Me.Width) / 2
     End If
    'Read Registry for previous settings stored
    lastPath = GetSetting(App.Title, "Settings", "BackupPath")
    lastDate = GetSetting(App.Title, "Settings", "BackupDate")
    lastTime = GetSetting(App.Title, "Settings", "BackupTime")
    
  lbl_status.Caption = "Select path and press Backup."

    If lastPath = "" Then
        lbl_path.Caption = "No Backup made previously"
        lbl_lastdate.Caption = " "
        lbl_lasttime.Caption = " "
    Else
        lbl_path.Caption = lastPath
        lbl_lastdate.Caption = lastDate
        lbl_lasttime.Caption = lastTime
    End If
End Sub

