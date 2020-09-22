VERSION 5.00
Begin VB.Form Frm_bookrpt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Book Report"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   Icon            =   "Frm_bookrpt.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   4710
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4455
      Begin VB.CommandButton Command1 
         Height          =   495
         Left            =   3120
         MouseIcon       =   "Frm_bookrpt.frx":24A2
         MousePointer    =   99  'Custom
         Picture         =   "Frm_bookrpt.frx":25F4
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Frm_bookrpt.frx":2A7F
         Left            =   1080
         List            =   "Frm_bookrpt.frx":2A89
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   2
         Top             =   720
         Width           =   1575
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
         Left            =   480
         TabIndex        =   7
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Book ID"
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
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Show"
         Height          =   255
         Left            =   3120
         TabIndex        =   5
         Top             =   780
         Width           =   1215
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Books options and pess 'Show' to show report"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "Frm_bookrpt"
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
ElseIf (Combo1.Text = "Book ID") Then
Text1.Text = ""
Text1.Locked = False
End If
End Sub
Private Sub Command1_Click()
If (Combo1.Text <> "All" And Combo1.Text <> "Book ID") Then
 MsgBox "Please select proper Book specifications.", vbCritical, "Invalid Data"
Exit Sub
End If

If (Combo1.Text = "All") Then
str = "Select * from Book"
ElseIf (Combo1.Text = "Book ID") Then
        If (Text1.Text <> "") Then
            If IsNumeric(Text1.Text) Then
            str = "Select * from Book where Bookid=" & Text1.Text
            Else
            MsgBox ("Please enter Book ID Numeric value."), vbExclamation, "Invalid value"
            Exit Sub
            End If
        Else
        MsgBox ("Please enter Book ID."), vbExclamation, "Invalid value"
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
Set Rpt_book.DataSource = rs
Rpt_book.Show vbModal
End Sub
Private Sub Form_Load()
Set db = New ADODB.Connection
Set rs = New ADODB.Recordset
db.CursorLocation = adUseClient
db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Database\Library.mdb;Jet OLEDB:Database Password=Library;"
status = False
End Sub

