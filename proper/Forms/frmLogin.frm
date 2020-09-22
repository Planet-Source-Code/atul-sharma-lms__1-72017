VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   3570
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3975
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2109.272
   ScaleMode       =   0  'User
   ScaleWidth      =   3732.31
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   3735
      Begin VB.TextBox txt_pass 
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
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1320
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   9
         ToolTipText     =   "Password"
         Top             =   720
         Width           =   2325
      End
      Begin VB.TextBox txt_uname 
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
         Height          =   345
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   8
         ToolTipText     =   "EmployeeID  for employee"
         Top             =   240
         Width           =   2325
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
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
         Height          =   270
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1080
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:"
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
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
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
      Left            =   2760
      MouseIcon       =   "frmLogin.frx":24A2
      MousePointer    =   99  'Custom
      Picture         =   "frmLogin.frx":25F4
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cancel to Abort"
      Top             =   2640
      Width           =   1020
   End
   Begin VB.CommandButton cmdOK 
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
      Left            =   1680
      MouseIcon       =   "frmLogin.frx":2B74
      MousePointer    =   99  'Custom
      Picture         =   "frmLogin.frx":2CC6
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Click to submit"
      Top             =   2640
      Width           =   1020
   End
   Begin VB.ComboBox cmb_as 
      Height          =   315
      ItemData        =   "frmLogin.frx":3238
      Left            =   1440
      List            =   "frmLogin.frx":3242
      TabIndex        =   0
      ToolTipText     =   "Enter as"
      Top             =   960
      Width           =   2295
   End
   Begin VB.Line Line1 
      X1              =   112.674
      X2              =   3605.552
      Y1              =   496.299
      Y2              =   496.299
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Specify option for enterring in Library system and enter password,Specify User name for Users."
      Height          =   615
      Left            =   1080
      TabIndex        =   6
      Top             =   120
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   240
      Picture         =   "frmLogin.frx":325F
      Top             =   0
      Width           =   720
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   3285
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   3285
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter as"
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
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str As String
Dim Loginrecord As ADODB.Recordset
Dim Loginconnection As ADODB.Connection

Private Sub cmb_as_Click()
If (cmb_as.Text = "Administrator") Then
txt_uname.Enabled = False
Else
txt_uname.Enabled = True
End If
End Sub

Private Sub cmdCancel_Click()
Unload mdi_start
End Sub
Private Sub cmdOK_Click()
If cmb_as.Text = "Administrator" Then
            If (txt_pass.Text = "") Then
            MsgBox "Please enter password.", vbInformation, "Password missing"
            txt_pass.SetFocus
            Exit Sub
            End If
  str = "Select Pass from Custom"
  Loginrecord.Open str, Loginconnection, adOpenStatic, adLockOptimistic
              If (txt_pass.Text = Loginrecord(0)) Then
               Loginrecord.Close
               uname = "Administrator"
               mdi_start.Enabled = True
               'mdi_start.Show
               Me.Hide
               DoEvents
               
               mdi_start.mnu_administer.Enabled = True
             Else
              Loginrecord.Close
              MsgBox "Invalid password.", vbInformation, "Acess Denied"
              txt_pass.Text = ""
              Exit Sub
             End If
ElseIf cmb_as.Text = "Employee" Then
            If (txt_uname.Text = "") Then
               MsgBox "Please enter User name.", vbInformation, "Username missing"
            ElseIf (txt_pass.Text = "") Then
               MsgBox "Please enter password.", vbInformation, "Password missing"
            Else
              str = "Select count(*) from Emptab where Empid = '" & Trim(txt_uname.Text) & "' and Psword = '" & Trim(txt_pass.Text) & "'"
              Loginrecord.Open str, Loginconnection, adOpenStatic, adLockOptimistic
                     If (Loginrecord(0) = 0) Then
                       MsgBox "Invalid password or Id.", vbInformation, "Acess Denied"
                       txt_uname.Text = ""
                       txt_pass.Text = ""
                       txt_uname.SetFocus
                       Loginrecord.Close
                       Exit Sub
                    Else
                Loginrecord.Close
                str = "Select Fname,Lname from Emptab where Empid = '" & Trim(txt_uname.Text) & "' and Psword = '" & Trim(txt_pass.Text) & "'"
                Loginrecord.Open str, Loginconnection, adOpenStatic, adLockOptimistic
                       uname = Loginrecord(0) & " " & Loginrecord(1)
                       mdi_start.Enabled = True
'                       mdi_start.Show
                       Me.Hide
                       DoEvents
                       If (Welcome = True) Then
            
                       End If
                       Loginrecord.Close
                       mdi_start.mnu_administer.Enabled = False
                    End If
            End If
Else
MsgBox "Invalid enter Catagory.", vbCritical, "Invalid catagory"
End If
mdi_start.sbStatusBar.Panels(3).Text = uname
End Sub
Private Sub Form_Activate()
 cmb_as.SetFocus
 mdi_start.Enabled = False

End Sub

Private Sub Form_Load()
  On Error GoTo errlable
  Set Loginconnection = New ADODB.Connection
  Loginconnection.CursorLocation = adUseClient
  Loginconnection.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Database\Library.mdb;Jet OLEDB:Database Password=Library;"
  
  str = "Select count(*) from Emptab"
  Set Loginrecord = New ADODB.Recordset
  Loginrecord.Open str, Loginconnection, adOpenStatic, adLockOptimistic
 If (Loginrecord(0) = 0) Then
 cmb_as.Text = "Administrator"
 cmb_as.Locked = True
 txt_uname.Enabled = False
 End If
Loginrecord.Close
Exit Sub

errlable:
MsgBox Err.Number & Err.Description
End Sub

