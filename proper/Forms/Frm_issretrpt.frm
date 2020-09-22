VERSION 5.00
Begin VB.Form Frm_issretrpt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Member issue-return Report"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5295
   Icon            =   "Frm_issretrpt.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   5295
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   5055
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Frm_issretrpt.frx":24A2
         Left            =   1440
         List            =   "Frm_issretrpt.frx":24AC
         TabIndex        =   4
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Height          =   495
         Left            =   3720
         MouseIcon       =   "Frm_issretrpt.frx":24C6
         MousePointer    =   99  'Custom
         Picture         =   "Frm_issretrpt.frx":2618
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txt_memid 
         Height          =   285
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   2
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Members"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C8D0D4&
         BackStyle       =   0  'Transparent
         Caption         =   "MemberID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Show"
         Height          =   255
         Left            =   3720
         TabIndex        =   5
         Top             =   795
         Width           =   1215
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Select  specifications for member and press 'Show'  to see report"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "Frm_issretrpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim db As ADODB.Connection
Dim status As Boolean
Dim str As String
Private Sub Combo1_Click()
If (Combo1.Text = "Specific Member") Then
txt_memid.Locked = False
ElseIf (Combo1.Text = "All") Then
txt_memid.Locked = True
End If
txt_memid.Text = ""
End Sub
Private Sub Command1_Click()
If (Combo1.Text <> "Specific Member" And Combo1.Text <> "All") Then
 MsgBox "Please select proper Member specifications.", vbCritical, "Invalid Data"
Exit Sub
End If
If (Combo1.Text = "Specific Member") Then
    If (txt_memid.Text <> "") Then
        If IsNumeric(txt_memid.Text) Then
        str = "select * from Issue where Memid=" & txt_memid.Text
        Else
        MsgBox "Please enter Member ID Numeric.", vbCritical, "Data missing"
        Exit Sub
        End If
    Else
    MsgBox "Please enter Member ID.", vbCritical, "Invalid Data"
    Exit Sub
    End If
Else
str = "select * from Issue"
End If

again:
If (status = False) Then
rs.Open str, db, adOpenStatic, adLockOptimistic
status = True
Else
rs.Close
status = False
GoTo again
End If

Set Rpt_Issret.DataSource = rs
Rpt_Issret.Show vbModal

End Sub

Private Sub Form_Load()
status = False

Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
Set db = New ADODB.Connection
db.CursorLocation = adUseClient
db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Database\Library.mdb;Jet OLEDB:Database Password=Library;"
End Sub

