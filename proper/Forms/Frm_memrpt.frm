VERSION 5.00
Begin VB.Form Frm_memrpt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Member report"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   Icon            =   "Frm_memrpt.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   5175
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4935
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   4
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Frm_memrpt.frx":24A2
         Left            =   1440
         List            =   "Frm_memrpt.frx":24AC
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Height          =   495
         Left            =   3480
         MouseIcon       =   "Frm_memrpt.frx":24C0
         MousePointer    =   99  'Custom
         Picture         =   "Frm_memrpt.frx":2612
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "For"
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
         Left            =   840
         TabIndex        =   7
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Member ID"
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
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Show"
         Height          =   255
         Left            =   3480
         TabIndex        =   5
         Top             =   765
         Width           =   1215
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Select  specifications for member and press 'Show'  to see report"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "Frm_memrpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim db As ADODB.Connection
Dim status As Boolean
Dim str As String
Private Sub Combo1_Click()
If (Combo1.Text = "All") Then
Text1.Text = ""
Text1.Locked = True
ElseIf (Combo1.Text = "Member ID") Then
Text1.Text = ""
Text1.Locked = False
End If
End Sub
Private Sub Command1_Click()
If (Combo1.Text <> "All" And Combo1.Text <> "Member ID") Then
 MsgBox "Please select proper Member specifications.", vbCritical, "Invalid Data"
Exit Sub
End If

If (Combo1.Text = "All") Then
str = "Select * from member"
ElseIf (Combo1.Text = "Member ID") Then
        If (Text1.Text <> "") Then
            If IsNumeric(Text1.Text) Then
            str = "Select * from member where Memid=" & Text1.Text
            Else
            MsgBox ("Please enter member ID Numeric value."), vbExclamation, "Invalid value"
            Exit Sub
            End If
        Else
        MsgBox ("Please enter member ID."), vbExclamation, "Invalid value"
        Exit Sub
        End If
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
Set Rpt_member.DataSource = rs
Rpt_member.Show vbModal
End Sub
Private Sub Form_Load()
Set db = New ADODB.Connection
Set rs = New ADODB.Recordset
db.CursorLocation = adUseClient
db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Database\Library.mdb;Jet OLEDB:Database Password=Library;"
status = False
End Sub

