VERSION 5.00
Begin VB.Form Frm_Employees 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee's details"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8295
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
   Icon            =   "Frm_Employee.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frm_cmd 
      ForeColor       =   &H00400040&
      Height          =   1095
      Left            =   480
      TabIndex        =   27
      Top             =   5040
      Width           =   7455
      Begin VB.CommandButton cmdLast 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   6960
         MouseIcon       =   "Frm_Employee.frx":24A2
         MousePointer    =   99  'Custom
         Picture         =   "Frm_Employee.frx":25F4
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Move Last"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   6600
         MouseIcon       =   "Frm_Employee.frx":2846
         MousePointer    =   99  'Custom
         Picture         =   "Frm_Employee.frx":2998
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Move Next"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   5520
         MouseIcon       =   "Frm_Employee.frx":2BA4
         MousePointer    =   99  'Custom
         Picture         =   "Frm_Employee.frx":2CF6
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Move Previous"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   5160
         MouseIcon       =   "Frm_Employee.frx":2F05
         MousePointer    =   99  'Custom
         Picture         =   "Frm_Employee.frx":3057
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Move First"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmd_new 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   120
         MaskColor       =   &H8000000F&
         MouseIcon       =   "Frm_Employee.frx":32A6
         MousePointer    =   99  'Custom
         Picture         =   "Frm_Employee.frx":33F8
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Add new record"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmd_edit 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   960
         MouseIcon       =   "Frm_Employee.frx":39D6
         MousePointer    =   99  'Custom
         Picture         =   "Frm_Employee.frx":3B28
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Edit record"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmd_delete 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   1800
         MouseIcon       =   "Frm_Employee.frx":40CD
         MousePointer    =   99  'Custom
         Picture         =   "Frm_Employee.frx":421F
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Delete record"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmd_save 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   2640
         MouseIcon       =   "Frm_Employee.frx":4769
         MousePointer    =   99  'Custom
         Picture         =   "Frm_Employee.frx":48BB
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Save record"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmd_cancel 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   3480
         MouseIcon       =   "Frm_Employee.frx":4E53
         MousePointer    =   99  'Custom
         Picture         =   "Frm_Employee.frx":4FA5
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Cancel"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmd_close 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   4320
         MouseIcon       =   "Frm_Employee.frx":5525
         MousePointer    =   99  'Custom
         Picture         =   "Frm_Employee.frx":5677
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
         Height          =   255
         Left            =   3480
         TabIndex        =   46
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Save"
         Height          =   255
         Left            =   2640
         TabIndex        =   45
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Delete"
         Height          =   255
         Left            =   1800
         TabIndex        =   44
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Edit"
         Height          =   255
         Left            =   960
         TabIndex        =   43
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "New"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Close"
         Height          =   255
         Left            =   4320
         TabIndex        =   41
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lbl_total 
         BackColor       =   &H00C8D0D4&
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   6360
         TabIndex        =   40
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lbl_rec 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C8D0D4&
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   5160
         TabIndex        =   39
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C8D0D4&
         BackStyle       =   0  'Transparent
         Caption         =   "of"
         Height          =   255
         Left            =   6120
         TabIndex        =   38
         Top             =   720
         Width           =   255
      End
   End
   Begin VB.Frame fra_personal 
      Caption         =   "Personal info"
      ForeColor       =   &H00000040&
      Height          =   2655
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   8055
      Begin VB.ComboBox cmb_sex 
         DataField       =   "Sex"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         ItemData        =   "Frm_Employee.frx":5BF1
         Left            =   1680
         List            =   "Frm_Employee.frx":5BFB
         Locked          =   -1  'True
         TabIndex        =   19
         Tag             =   "7"
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox txt_note 
         DataField       =   "Note"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   70
         TabIndex        =   18
         Tag             =   "8"
         Top             =   2160
         Width           =   6135
      End
      Begin VB.TextBox txt_phone 
         DataField       =   "Phone"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   17
         Tag             =   "6"
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox txt_mail 
         DataField       =   "Email"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   16
         Tag             =   "5"
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox txt_add 
         DataField       =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   1455
         Left            =   5040
         Locked          =   -1  'True
         MaxLength       =   125
         MultiLine       =   -1  'True
         TabIndex        =   15
         Tag             =   "11"
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox txt_lname 
         DataField       =   "Lname"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   14
         Tag             =   "4"
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txt_fname 
         DataField       =   "Fname"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   13
         Tag             =   "3"
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label lbl_add 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         Left            =   4200
         TabIndex        =   26
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lbl_sex 
         BackStyle       =   0  'Transparent
         Caption         =   "Sex"
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
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lbl_note 
         BackStyle       =   0  'Transparent
         Caption         =   "Special note"
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
         Top             =   2205
         Width           =   1215
      End
      Begin VB.Label lbl_phone 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone no."
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
         Top             =   1485
         Width           =   975
      End
      Begin VB.Label lbl_mail 
         BackStyle       =   0  'Transparent
         Caption         =   "Email address"
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
         TabIndex        =   22
         Top             =   1140
         Width           =   1335
      End
      Begin VB.Label lbl_lname 
         BackStyle       =   0  'Transparent
         Caption         =   "Last name"
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
         TabIndex        =   21
         Top             =   780
         Width           =   975
      End
      Begin VB.Label lbl_fname 
         BackStyle       =   0  'Transparent
         Caption         =   "First name"
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
         TabIndex        =   20
         Top             =   420
         Width           =   1095
      End
   End
   Begin VB.Frame frm_post 
      Caption         =   "Office info."
      ForeColor       =   &H00000040&
      Height          =   1575
      Left            =   5040
      TabIndex        =   7
      Top             =   720
      Width           =   3135
      Begin VB.ComboBox cmb_post 
         DataField       =   "Post"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         ItemData        =   "Frm_Employee.frx":5C0D
         Left            =   1200
         List            =   "Frm_Employee.frx":5C1A
         Locked          =   -1  'True
         TabIndex        =   9
         Tag             =   "9"
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txt_sal 
         Alignment       =   1  'Right Justify
         DataField       =   "Salary"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   8
         Tag             =   "10"
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lbl_salary 
         BackStyle       =   0  'Transparent
         Caption         =   "Salary"
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
         TabIndex        =   11
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lbl_post 
         BackStyle       =   0  'Transparent
         Caption         =   "Post-aid"
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
         TabIndex        =   10
         Top             =   420
         Width           =   855
      End
   End
   Begin VB.Frame fra_log 
      Caption         =   "Login info."
      ForeColor       =   &H00000040&
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   4455
      Begin VB.TextBox txt_pass2 
         DataField       =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   3
         Tag             =   "2"
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox txt_pass1 
         DataField       =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   2
         Tag             =   "1"
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txt_empid 
         DataField       =   "Empid"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   1
         Tag             =   "0"
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label1 
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
         TabIndex        =   6
         Top             =   1095
         Width           =   1575
      End
      Begin VB.Label lbl_pass1 
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
         TabIndex        =   5
         Top             =   750
         Width           =   975
      End
      Begin VB.Label lbl_ID 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee ID"
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
         TabIndex        =   4
         Top             =   420
         Width           =   1095
      End
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Administrator can add new employee or edit existence employee profiles."
      Height          =   375
      Left            =   1080
      TabIndex        =   47
      Top             =   120
      Width           =   6975
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   240
      Picture         =   "Frm_Employee.frx":5C39
      Top             =   0
      Width           =   720
   End
End
Attribute VB_Name = "Frm_Employees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Emprecordset As ADODB.Recordset
Dim Empconnection As ADODB.Connection
Dim saveflag As Boolean
Dim str As String
Dim slct As String
Private Function cheak() As Boolean
Dim flag As Boolean
flag = False
                If txt_empid.Text = "" Then
                  MsgBox ("Please enter EmployeeID."), vbInformation, "Data missing"
                ElseIf txt_pass1.Text = "" Then
                  MsgBox ("Please enter Password."), vbInformation, "Data missing"
                ElseIf txt_pass2.Text = "" Then
                 MsgBox ("Please enter Password as Varifier so that wrong password can be detected."), vbInformation, "Data missing"
                ElseIf txt_fname.Text = "" Then
                 MsgBox ("Please enter Employee first name."), vbInformation, "Data missing"
                ElseIf txt_lname.Text = "" Then
                 MsgBox ("Please enter EmployeeID second name."), vbInformation, "Data missing"
                ElseIf cmb_sex.Text = "" Then
                  MsgBox ("Please select the sex."), vbInformation, "Invalid arguments"
                ElseIf (cmb_sex.Text <> "Male" And cmb_sex.Text <> "Female") Then
                 MsgBox ("Please select the sex."), vbInformation, "Invalid arguments"
                ElseIf cmb_post.Text = "" Then
                 MsgBox ("Please select the post-aid."), vbInformation, "Invalid arguments"
                ElseIf (cmb_post.Text <> "New" And cmb_post.Text <> "Temporary" And cmb_post.Text <> "Permanent") Then
                 MsgBox ("Please select the post-aid."), vbInformation, "Invalid arguments"
                ElseIf txt_add.Text = "" Then
                 MsgBox ("Please enter Employee contact address."), vbInformation, "Data missing"
                ElseIf txt_pass1.Text <> txt_pass2.Text Then
                 MsgBox ("May be typing mistake,Please re-enter the password."), vbInformation, "Invalid password"
                 txt_pass1.Text = ""
                 txt_pass2.Text = ""
                 txt_pass1.SetFocus
                Else
                 flag = True
                End If
cheak = flag
End Function
Private Sub locate()
  lbl_total.Caption = Emprecordset.RecordCount
  lbl_rec.Caption = Emprecordset.AbsolutePosition
End Sub
Private Sub showdata()
  If Emprecordset.EOF = False And Emprecordset.BOF = False Then
                    txt_add.Text = Emprecordset.Fields(0)
                    txt_mail.Text = Emprecordset.Fields(1)
                    txt_empid.Text = Emprecordset.Fields(2)
                    txt_fname.Text = Emprecordset.Fields(3)
                    txt_lname.Text = Emprecordset.Fields(4)
                    txt_phone.Text = Emprecordset.Fields(5)
                    cmb_post.Text = Emprecordset.Fields(6)
                    txt_pass1.Text = Emprecordset.Fields(7)
                    txt_pass2.Text = Emprecordset.Fields(7)
                    txt_sal.Text = Emprecordset.Fields(8)
                    cmb_sex.Text = Emprecordset.Fields(9)
                    txt_note.Text = Emprecordset.Fields(10)
                  
 End If
 Call locate
 End Sub
Private Sub clear()
                    txt_add.Text = ""
                    cmb_post.Text = ""
                    txt_mail.Text = ""
                    txt_empid.Text = ""
                    txt_fname.Text = ""
                    txt_lname.Text = ""
                    txt_note.Text = ""
                    txt_pass1.Text = ""
                    txt_pass2.Text = ""
                    txt_phone.Text = ""
                    txt_sal.Text = ""
                    cmb_sex.Text = ""
End Sub
Private Sub setlock(val As Boolean)
                    txt_add.Locked = val
                    cmb_post.Locked = val
                    txt_mail.Locked = val
                    txt_empid.Locked = val
                    txt_fname.Locked = val
                    txt_lname.Locked = val
                    txt_note.Locked = val
                    txt_pass1.Locked = val
                    txt_pass2.Locked = val
                    txt_phone.Locked = val
                    cmb_sex.Locked = val
End Sub
Private Sub button(val As Boolean)
                    cmd_new.Enabled = val
                    cmd_edit.Enabled = val
                    cmd_delete.Enabled = val
                    cmdFirst.Enabled = val
                    cmdLast.Enabled = val
                    cmdNext.Enabled = val
                    cmdPrevious.Enabled = val
                    cmd_save.Enabled = Not val
                    cmd_cancel.Enabled = Not val
                  
End Sub

Private Sub cmb_post_Click()
    If cmb_post.Text = "New" Then
        txt_sal.Text = salnew
     ElseIf cmb_post.Text = "Temporary" Then
        txt_sal.Text = saltemp
     Else
        txt_sal.Text = salper
     End If
   
End Sub

Private Sub cmd_cancel_Click()
On erro GoTo cancelerr
'disablink control
    setlock (True)
 
 If Emprecordset.BOF And Emprecordset.EOF Then
   GoTo newproc
 Else
   Emprecordset.MoveFirst
   Call showdata
 End If

newproc:
  txt_fname.SetFocus
'enable control
    cmdFirst.Enabled = True
    cmdPrevious.Enabled = True
    cmdNext.Enabled = True
    cmdLast.Enabled = True
    cmd_delete.Enabled = True
    cmd_edit.Enabled = True
    cmd_new.Enabled = True
'disable buttons
    cmd_save.Enabled = False
    cmd_cancel.Enabled = False

Exit Sub
cancelerr:
MsgBox Err.Description
End Sub

Private Sub cmd_close_Click()
Unload Me
End Sub


Private Sub Form_Load()
  On Error GoTo errlable
     If (view = 1) Then
     Me.Top = 50
     Me.Left = 50
     ElseIf (view = 2) Then
     Me.Top = 700
     Me.Left = (Screen.Width - Me.Width) / 2
     End If
  
  Set Empconnection = New ADODB.Connection
  Empconnection.CursorLocation = adUseClient
  Empconnection.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Database\Library.mdb;Jet OLEDB:Database Password=Library;"
  
  slct = "select Address,Email,Empid,Fname,Lname,Phone,Pos,Psword,Salary,Sex,Spe from Emptab Order by Fname"
  Set Emprecordset = New ADODB.Recordset
  Emprecordset.Open slct, Empconnection, adOpenStatic, adLockOptimistic
 
 Call showdata
   cmd_save.Enabled = False
   cmd_cancel.Enabled = False
Exit Sub
errlable:
MsgBox Err.Number & Err.Description
End Sub

Private Sub cmd_delete_Click()
On Error GoTo delerr
 Beep
If MsgBox("Execution of command will delete current Datarecord,Are you sure you wan't to delete Datarecord ?", vbYesNo + vbExclamation, "Confirm Delete") = vbYes Then
str = "DELETE FROM Emptab WHERE "
str = str & "Empid='"
str = str & Trim(txt_empid.Text) & "'"
'MsgBox str
Empconnection.Execute str
Emprecordset.Requery
MsgBox ("Record deleted Successfully."), vbInformation, "Delete"

        If Emprecordset.BOF And Emprecordset.EOF Then
            Call clear
            MsgBox ("The previous record was last record,Now no record left."), vbInformation, "Last record"
            cmd_delete.Enabled = False
        Else
            Emprecordset.MoveNext
                If Emprecordset.EOF Then
                Emprecordset.MoveLast
                End If
            Call showdata
       End If
End If
Exit Sub
delerr:
MsgBox Err.Description
End Sub
Private Sub cmd_save_Click()
On erro GoTo saver
If cheak = True Then
  'Autocorrection procedure
      If cmb_post.Text = "New" Then
        txt_sal.Text = salnew
     ElseIf cmb_post.Text = "Temporary" Then
        txt_sal.Text = saltemp
     Else
        txt_sal.Text = salper
     End If
   
    If txt_mail.Text = "" Then
    txt_mail.Text = "None"
    End If
    
    If txt_phone.Text = "" Then
    txt_phone.Text = "None"
    End If
    
    If txt_note.Text = "" Then
    txt_note.Text = "None"
    End If
  
      If saveflag = True Then
            'for new record
            str = "INSERT INTO Emptab "
            str = str & "(Address,Email,Empid,Fname,Lname,Phone,Pos,Psword,Salary,Sex,Spe) "
            str = str & "VALUES" & "('" & Trim(txt_add.Text) & "', "
            str = str & "'" & Trim(txt_mail.Text) & "', "
            str = str & "'" & Trim(txt_empid.Text) & "', "
            str = str & "'" & Trim(txt_fname.Text) & "', "
            str = str & "'" & Trim(txt_lname.Text) & "', "
            str = str & "'" & Trim(txt_phone.Text) & "', "
            str = str & "'" & Trim(cmb_post.Text) & "', "
            str = str & "'" & Trim(txt_pass1.Text) & "', "
            str = str & CDbl(txt_sal.Text) & ","
            str = str & "'" & Trim(cmb_sex.Text) & "', "
            str = str & "'" & Trim(txt_note.Text) & "')"
            'MsgBox str
            Empconnection.Execute str, , adCmdText + adExecuteNoRecords
            
      Else
            'for updating current record
            str = "UPDATE Emptab SET "
            str = str & "Address = '" & Trim(txt_add.Text) & "',"
            str = str & " Pos = '" & Trim(cmb_post.Text) & "',"
            str = str & " Email = '" & Trim(txt_mail.Text) & "',"
            str = str & " Empid = '" & Trim(txt_empid.Text) & "',"
            str = str & " Fname = '" & Trim(txt_fname.Text) & "',"
            str = str & " Lname = '" & Trim(txt_lname.Text) & "',"
            str = str & " Spe = '" & Trim(txt_note.Text) & "',"
            str = str & " Psword = '" & Trim(txt_pass1.Text) & "',"
            str = str & " Phone = '" & Trim(txt_phone.Text) & "',"
            str = str & " Salary = " & CDbl(txt_sal.Text) & ","
            str = str & " Sex = '" & Trim(cmb_sex.Text) & "'"
            str = str & " WHERE Empid = '" & Trim(txt_empid.Text) & "'"
            'MsgBox str
            Empconnection.Execute str
            End If

            Emprecordset.Requery
            MsgBox ("Record saved successfully.")
            Call setlock(True)
            Call button(True)
            Call showdata
  
End If
Exit Sub
saver:
MsgBox Err.Description
End Sub

Private Sub cmdFirst_Click()
 On Error GoTo GoFirstError

   Emprecordset.MoveFirst
'show thw current data record
   Call showdata
Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
 On Error GoTo GoLastError

   Emprecordset.MoveLast
'show thw current data record
   Call showdata
Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
On Error GoTo GoNextError
  
  If Not Emprecordset.EOF Then Emprecordset.MoveNext
  If Emprecordset.EOF And Emprecordset.RecordCount > 0 Then
     Beep
     'moved off the end so go back
     Emprecordset.MoveLast
    
  End If
'show thw current data record
     Call showdata
Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
 On Error GoTo GoPrevError

  If Not Emprecordset.BOF Then Emprecordset.MovePrevious
  If Emprecordset.BOF And Emprecordset.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    Emprecordset.MovePrevious
 
  End If
'show thw current data record
    Call showdata
Exit Sub

GoPrevError:
 If Err.Number = 3021 Then
MsgBox ("This is first Record."), vbInformation, "First record"
Emprecordset.MoveNext
ElseIf Err.Number <> 0 Then
MsgBox Err.Number & Err.Description
End If
End Sub

Private Sub cmd_edit_Click()
On Error GoTo editerr
            'Call clear
            Call button(False)
            Call setlock(False)
            'cmd_cancel.Enabled = False
            saveflag = False
            txt_empid.Locked = True
            txt_fname.SetFocus
Exit Sub
editerr:
MsgBox Err.Description
End Sub

Private Sub cmd_new_Click()
On Error GoTo newerr
            Call clear
            Call button(False)
            Call setlock(False)
            saveflag = True
            txt_empid.SetFocus
Exit Sub
newerr:
MsgBox Err.Description
End Sub

