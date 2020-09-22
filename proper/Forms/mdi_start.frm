VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdi_start 
   BackColor       =   &H8000000F&
   Caption         =   "Library management system."
   ClientHeight    =   9060
   ClientLeft      =   60
   ClientTop       =   -2175
   ClientWidth     =   14775
   Icon            =   "mdi_start.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Picture         =   "mdi_start.frx":0ECA
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   20
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Books"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Members"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Issue"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Return"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Fine Informations"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Search"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Global"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Reports"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Book report"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Member report"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Issue report"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Calculator"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Notepad"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Keybord shotcuts"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "About"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Log off"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exit"
         EndProperty
      EndProperty
      MousePointer    =   99
      MouseIcon       =   "mdi_start.frx":2E4369
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   8790
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10583
            MinWidth        =   10583
            Text            =   "Library Management System(Created By Atul N Dhiraj)"
            TextSave        =   "Library Management System(Created By Atul N Dhiraj)"
            Object.ToolTipText     =   "Graphics by Atul, Contact :atulsharma06@yahoo.com"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   970
            MinWidth        =   970
            Text            =   "   User"
            TextSave        =   "   User"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3881
            MinWidth        =   3881
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   1349
            MinWidth        =   882
            Text            =   "  Today"
            TextSave        =   "  Today"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   2293
            MinWidth        =   2293
            TextSave        =   "4/16/2009"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "11:46 AM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   5609
            Text            =   "Contact : atulsharma06@yahoo.com"
            TextSave        =   "Contact : atulsharma06@yahoo.com"
            Object.ToolTipText     =   "Created by : Atul N Dhiraj"
         EndProperty
      EndProperty
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "mdi_start.frx":2E44CB
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   360
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_start.frx":2E462D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_start.frx":2E5307
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_start.frx":2E5FE1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_start.frx":2E6CBB
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_start.frx":2E7995
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_start.frx":2E866F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_start.frx":2E9349
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_start.frx":2EA023
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_start.frx":2EACFD
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_start.frx":2EB9D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_start.frx":2EC6B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_start.frx":2ED38B
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_start.frx":2EE065
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_start.frx":2EED3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_start.frx":2EFA19
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_start.frx":2F06F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi_start.frx":2F13CD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnu_database 
      Caption         =   "&Database"
      Begin VB.Menu sm_books 
         Caption         =   "&Books"
         Shortcut        =   ^B
      End
      Begin VB.Menu sm_members 
         Caption         =   "&Members"
         Shortcut        =   ^M
      End
      Begin VB.Menu firstbarfirst 
         Caption         =   "-"
      End
      Begin VB.Menu sm_logoff 
         Caption         =   "&Logoff"
         Shortcut        =   ^L
      End
      Begin VB.Menu firstbarsecond 
         Caption         =   "-"
      End
      Begin VB.Menu sm_exit 
         Caption         =   "&Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnu_tranjection 
      Caption         =   "T&ransaction"
      Begin VB.Menu sm_issue 
         Caption         =   "&Issue"
         Shortcut        =   ^I
      End
      Begin VB.Menu sm_return 
         Caption         =   "&Return"
         Shortcut        =   ^R
      End
      Begin VB.Menu sm_fine 
         Caption         =   "&Fine Informations"
         Shortcut        =   ^F
      End
      Begin VB.Menu secondbarfirst 
         Caption         =   "-"
      End
      Begin VB.Menu sm_search 
         Caption         =   "&Search.."
         Shortcut        =   ^S
      End
      Begin VB.Menu line 
         Caption         =   "-"
      End
      Begin VB.Menu sm_global 
         Caption         =   "&Global"
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu mnu_administer 
      Caption         =   "&Administrator"
      Begin VB.Menu sm_employees 
         Caption         =   "&Employees"
         Shortcut        =   ^E
      End
      Begin VB.Menu thirdbarfirst 
         Caption         =   "-"
      End
      Begin VB.Menu sm_backup 
         Caption         =   "Back up"
         Shortcut        =   ^U
      End
      Begin VB.Menu temp 
         Caption         =   "-"
      End
      Begin VB.Menu sm_settings 
         Caption         =   "Se&ttings"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu Mnu_rep 
      Caption         =   "&Reports"
      Begin VB.Menu sm_bookrpt 
         Caption         =   "Book Report"
      End
      Begin VB.Menu sm_member 
         Caption         =   "Member Report"
      End
      Begin VB.Menu sm_issret 
         Caption         =   "Issue return Report"
      End
   End
   Begin VB.Menu mnu_tools 
      Caption         =   "&Tools"
      Begin VB.Menu sm_notepad 
         Caption         =   "Notepad"
         Shortcut        =   ^N
      End
      Begin VB.Menu sm_calculator 
         Caption         =   "Calculator"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu mnu_help 
      Caption         =   "&Help"
      Begin VB.Menu smnu_keyboard 
         Caption         =   "Keyboard"
         Shortcut        =   ^K
      End
      Begin VB.Menu nwlne 
         Caption         =   "-"
      End
      Begin VB.Menu sm_about 
         Caption         =   "About"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "mdi_start"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OSWinHelp% Lib "User32" Alias "WinHelpA" (ByVal hWnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Sub MDIForm_Load()
Me.Show
Me.Enabled = False
'setting toolbar images
With Toolbar2
Set .ImageList = ImageList1
.Buttons(2).Image = 1
.Buttons(3).Image = 7
.Buttons(5).Image = 5
.Buttons(6).Image = 6
.Buttons(7).Image = 14
.Buttons(8).Image = 2
.Buttons(9).Image = 3
.Buttons(11).Image = 10
.Buttons(13).Image = 8
.Buttons(14).Image = 9
.Buttons(16).Image = 12
.Buttons(17).Image = 13
.Buttons(19).Image = 4
.Buttons(20).Image = 11
End With
sbStatusBar.Panels(3).Text = "Login"
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
On Error Resume Next
If MsgBox("Are You Sure you want to Quit ?", vbExclamation + vbOKCancel, "Library Management System") = vbOK Then
Unload frmLogin
Else
Cancel = True
End If
End Sub
Private Sub sbStatusBar_PanelClick(ByVal Panel As MSComctlLib.Panel)
ShellExecute Me.hWnd, vbNullString, "http://atul.freind4u.googlepages.com/home", vbNullString, vbNullString, SW_SHOWNORMAL
End Sub
Private Sub sm_about_Click()
Load frmAbout
frmAbout.Show
End Sub
Private Sub sm_backup_Click()
Load Frm_backup
Frm_backup.Show
End Sub

Private Sub sm_bookrpt_Click()
Load Frm_bookrpt
Frm_bookrpt.Show
End Sub

Private Sub sm_books_Click()
Load Frm_books
Frm_books.Show
End Sub
Private Sub sm_calculator_Click()
On Error GoTo errHandle
    Dim a As Double
    a = Shell("C:\WINDOWS\System32\calc.exe", vbNormalFocus)
    Exit Sub
errHandle:
    MsgBox "Unable to run Calculator Utility on your computer", vbInformation, "Error in opening!!!"
    Resume Next
End Sub
Private Sub sm_employees_Click()
Load Frm_Employees
Frm_Employees.Show
End Sub
Private Sub sm_exit_Click()
Unload Me
End Sub

Private Sub sm_fine_Click()
Load Frm_Fine
Frm_Fine.Show
End Sub

Private Sub sm_global_Click()
Load Frm_global
Frm_global.Show
End Sub
Private Sub sm_help_Click()
 Dim nRet As Integer
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hWnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub
Private Sub sm_hsearch_Click()
    Dim nRet As Integer
    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hWnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If
End Sub

Private Sub sm_issret_Click()
Load Frm_issretrpt
Frm_issretrpt.Show
End Sub
Private Sub sm_issue_Click()
Load Frm_issue
Frm_issue.Show
End Sub
Private Sub sm_logoff_Click()
If MsgBox("Are You Sure you want to logoff ?", vbExclamation + vbOKCancel, "Library Management System") = vbOK Then
Call logoff
DoEvents
End If
End Sub

Private Sub sm_member_Click()
Load Frm_memrpt
Frm_memrpt.Show
End Sub
Private Sub sm_members_Click()
Load Frm_members
Frm_members.Show
End Sub
Private Sub sm_notepad_Click()
On Error GoTo errcode
    Dim a As Double
    a = Shell("C:\WINDOWS\System32\notepad.exe", vbNormalFocus)
    Exit Sub
errcode:
    MsgBox "Unable to run Notepad Utility on your computer", vbInformation, "Error in opening!!!"
    Resume Next
End Sub
Private Sub sm_return_Click()
Load Frm_return
Frm_return.Show
End Sub
Private Sub sm_search_Click()
Load Frm_search
Frm_search.Show
End Sub
Private Sub sm_settings_Click()
Load Frm_settings
Frm_settings.Show
End Sub
Private Sub smnu_keyboard_Click()
Load Frm_keyboard
Frm_keyboard.Show
End Sub
Private Sub Toolbar2_ButtonClick(ByVal button As MSComctlLib.button)
Select Case button.Index
    Case 2: Call sm_books_Click
    Case 3: Call sm_members_Click
    Case 5: Call sm_issue_Click
    Case 6: Call sm_return_Click
    Case 7: Call sm_fine_Click
    Case 8: Call sm_search_Click
    Case 9: Call sm_global_Click
    Case 11: 'add report
    Case 13: Call sm_calculator_Click
    Case 14: Call sm_notepad_Click
    Case 16: Call smnu_keyboard_Click
    Case 17: Call sm_about_Click
    Case 19: Call sm_logoff_Click
    Case 20: Call sm_exit_Click
End Select
End Sub
Private Sub Toolbar2_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.Index
    Case 1:
         Call sm_bookrpt_Click
    Case 2:
         Call sm_member_Click
    Case 3:
         Call sm_issret_Click
End Select
End Sub
