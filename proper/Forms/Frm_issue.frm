VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frm_issue 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Issue"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3975
   Icon            =   "Frm_issue.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   3735
      Begin VB.CommandButton cmd_add 
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
         MouseIcon       =   "Frm_issue.frx":24A2
         MousePointer    =   99  'Custom
         Picture         =   "Frm_issue.frx":25F4
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Add new"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_issue 
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
         Left            =   2520
         MouseIcon       =   "Frm_issue.frx":2BD2
         MousePointer    =   99  'Custom
         Picture         =   "Frm_issue.frx":2D24
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Issue book"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_return 
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
         MouseIcon       =   "Frm_issue.frx":331E
         MousePointer    =   99  'Custom
         Picture         =   "Frm_issue.frx":3470
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Switch to Return form"
         Top             =   1200
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
         Left            =   1320
         MouseIcon       =   "Frm_issue.frx":3A79
         MousePointer    =   99  'Custom
         Picture         =   "Frm_issue.frx":3BCB
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Cancel"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdFirst 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   1440
         MouseIcon       =   "Frm_issue.frx":414B
         MousePointer    =   99  'Custom
         Picture         =   "Frm_issue.frx":429D
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Move First"
         Top             =   1200
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   1800
         MouseIcon       =   "Frm_issue.frx":44EC
         MousePointer    =   99  'Custom
         Picture         =   "Frm_issue.frx":463E
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Move Previous"
         Top             =   1200
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   2880
         MouseIcon       =   "Frm_issue.frx":484D
         MousePointer    =   99  'Custom
         Picture         =   "Frm_issue.frx":499F
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Move Next"
         Top             =   1200
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdLast 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   3240
         MouseIcon       =   "Frm_issue.frx":4BAB
         MousePointer    =   99  'Custom
         Picture         =   "Frm_issue.frx":4CFD
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Move Last"
         Top             =   1200
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "New"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   885
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
         Height          =   255
         Left            =   1320
         TabIndex        =   24
         Top             =   885
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Issue"
         Height          =   255
         Left            =   2520
         TabIndex        =   23
         Top             =   885
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Move to Return"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1840
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "of"
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
         Left            =   2400
         TabIndex        =   21
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label lbl_rec 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   1440
         TabIndex        =   20
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lbl_total 
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
         Height          =   255
         Left            =   2640
         TabIndex        =   19
         Top             =   1680
         Width           =   975
      End
   End
   Begin VB.Frame Fra_Date 
      Caption         =   "Date of"
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
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   3735
      Begin MSMask.MaskEdBox msk_return 
         Height          =   285
         Left            =   2160
         TabIndex        =   3
         ToolTipText     =   "Administrator default settings"
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         ForeColor       =   4194304
         MaxLength       =   10
         Format          =   "mm/dd/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox msk_issue 
         Height          =   285
         Left            =   2160
         TabIndex        =   2
         ToolTipText     =   "Administrator default settings"
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         ForeColor       =   4194304
         MaxLength       =   10
         Format          =   "mm/dd/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Doissue 
         BackStyle       =   0  'Transparent
         Caption         =   "Issue (mm/dd/yyyy)"
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
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lbl_Doreturn 
         BackStyle       =   0  'Transparent
         Caption         =   "Return (mm/dd/yyyy)"
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
         Top             =   750
         Width           =   1815
      End
   End
   Begin VB.TextBox txt_bookid 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox txt_memid 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   1335
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3840
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "For Issue book enter MemberID and Bookid. Issuedate is Currentdate and Returndate is date before that book should be retuned."
      Height          =   615
      Left            =   615
      TabIndex        =   9
      Top             =   120
      Width           =   3240
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   120
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lbl_bookid 
      BackStyle       =   0  'Transparent
      Caption         =   "Book ID"
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
      TabIndex        =   5
      Top             =   1335
      Width           =   735
   End
   Begin VB.Label lbl_memberid 
      BackStyle       =   0  'Transparent
      Caption         =   "Member ID"
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
      TabIndex        =   4
      Top             =   1005
      Width           =   1095
   End
End
Attribute VB_Name = "Frm_issue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str As String
Dim rmem As ADODB.Recordset
Dim rbook As ADODB.Recordset
Dim riss As ADODB.Recordset
Dim Issueconnection As ADODB.Connection
Dim Issuerecord As ADODB.Recordset
Private Sub cmd_add_Click()
Call cleartext
Call setbutton(False)
Call locktext(False)
msk_issue.Text = Format$(Now, "mm/dd/yyyy")
'msk_issue.Enabled = False
msk_return.Text = Format$(Now + dayslimit, "mm/dd/yyyy")
'msk_return.Enabled = False
End Sub
Private Sub locate()
  lbl_total.Caption = Issuerecord.RecordCount
  lbl_rec.Caption = Issuerecord.AbsolutePosition
End Sub
Private Sub locktext(val As Boolean)
txt_bookid.Locked = val
msk_issue.Enabled = Not val
msk_return.Enabled = Not val
txt_memid.Locked = val
End Sub
Private Sub setbutton(val As Boolean)
cmd_add.Enabled = val
cmd_Return.Enabled = val
cmdFirst.Enabled = val
cmdLast.Enabled = val
cmdNext.Enabled = val
cmdPrevious.Enabled = val
cmd_issue.Enabled = Not val
cmd_cancel.Enabled = Not val
End Sub
Private Function cheak() As Boolean
Dim flag As Boolean
flag = False
If msk_return.Text = "__/__/____" Then
MsgBox "Please select the date.", vbInformation, "Field missing"
ElseIf msk_issue.Text = "__/__/____" Then
ElseIf txt_bookid.Text = "" Then
MsgBox "Please enter the Bookid.", vbInformation, "Field missing"
ElseIf txt_memid.Text = "" Then
MsgBox "Please enter the Memberid.", vbInformation, "Field missing"
Else
flag = True
End If
cheak = flag
End Function
Private Sub cleartext()
txt_bookid.Text = ""
msk_issue.Text = "__/__/____"
msk_return.Text = "__/__/____"
txt_memid.Text = ""
End Sub
Private Sub cmd_cancel_Click()
Call locktext(True)
Call setbutton(True)
 If Not (Issuerecord.BOF And Issuerecord.EOF) Then
   Issuerecord.MoveFirst
   Call showdata
 End If
End Sub
Private Sub cmd_issue_Click()
On Error GoTo errlable
If (cheak = True) Then

'If member id exists
str = "select count(*) from Member where Memid = " & Trim(txt_memid.Text)
rmem.Open str, Issueconnection, adOpenStatic, adLockOptimistic
If rmem(0) = 0 Then
    MsgBox ("Member with mentioned memberID does not exists."), vbCritical, "Invalid arguments"
    rmem.Close
    Exit Sub
Else
    'Is capable of holding book.
    rmem.Close
    str = "select Bookinhand from Member where Memid = " & Trim(txt_memid.Text)
    rmem.Open str, Issueconnection, adOpenStatic, adLockOptimistic
            If rmem(0) = maxhold Then
            MsgBox ("Members can not hold books greater then " & maxhold & "."), vbCritical, "Invalid arguments"
            rmem.Close
            GoTo recycle
            End If
End If
rmem.Close
'if book is present for specified bookid
str = "select count(*) from Book where Bookid = " & Trim(txt_bookid.Text)
rbook.Open str, Issueconnection, adOpenStatic, adLockOptimistic
If rbook(0) = 0 Then
    MsgBox ("Book with mentioned bookid does not exists."), vbCritical, "Invalid arguments"
    rbook.Close
    Exit Sub
Else
    'is there available copy
    rbook.Close
    str = "select Avano from Book where Bookid = " & Trim(txt_bookid.Text)
    rbook.Open str, Issueconnection, adOpenStatic, adLockOptimistic
            If rbook(0) <= refcopy Then
            MsgBox ("Book contains only refrence copies which cannot be issued."), vbCritical, "Invalid arguments"
            rbook.Close
            GoTo recycle
            End If
End If
rbook.Close
'member has same book or not
 str = "Select count(*) from Issue where Bookid = " & Trim(txt_bookid.Text) & " And Memid = " & Trim(txt_memid.Text)
 riss.Open str, Issueconnection, adOpenStatic, adLockOptimistic
 If (riss(0) <> 0) Then
     MsgBox ("Member has already issue mentioned book copy.member can not take same book again."), vbCritical, "Invalid arguments"
     riss.Close
 Exit Sub
 End If
 Beep
If MsgBox("Issue Info.:MemberId=" & CDbl(txt_memid.Text) & " And  BookId=" & CDbl(txt_bookid.Text), vbYesNo, "Confirm Data") = vbYes Then
            str = "INSERT INTO Issue"
            str = str & " (Areturndate,Bookid,Issuedate,Returndate,Memid) "
            str = str & "VALUES('" & Trim(msk_return.Text) & "', "
            str = str & CDbl(txt_bookid.Text) & ", "
            str = str & "'" & Trim(msk_issue.Text) & "', "
            str = str & "'" & Trim(msk_return.Text) & "', "
            str = str & CDbl(txt_memid.Text) & ")"
            Issueconnection.Execute str
            
            str = "UPDATE Book SET "
            str = str & "Avano = Avano-1,"
            str = str & "Issno = Issno+1 where Bookid = " & Trim(txt_bookid.Text)
            Issueconnection.Execute str
            
            str = "UPDATE Member SET "
            str = str & "Bookinhand = Bookinhand+1 where Memid = " & Trim(txt_memid.Text)
            Issueconnection.Execute str
            
            Issuerecord.Requery
            MsgBox "All entry Updated sucessfully.", vbInformation, "Record saved"
    Call locktext(True)
    Call setbutton(True)
Else
recycle:
    Call locktext(True)
    Call setbutton(True)
    Call cleartext
End If

End If
Exit Sub
errlable:
MsgBox Err.Number & Err.Description
End Sub
Private Sub cmd_Return_Click()
Load Frm_return
Frm_return.Show
Unload Me
End Sub
Private Sub Form_Load()
On Error GoTo lable
     If (view = 1) Then
     Me.Top = 50
     Me.Left = 50
     ElseIf (view = 2) Then
     Me.Top = 700
     Me.Left = (Screen.Width - Me.Width) / 2
     End If
Image1.Picture = mdi_start.ImageList1.ListImages(5).Picture
Set Issueconnection = New ADODB.Connection
Issueconnection.CursorLocation = adUseClient
 Issueconnection.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & "Data source=" & App.Path & "\Database\Library.mdb;Jet OLEDB:Database Password=Library;"

Set Issuerecord = New ADODB.Recordset
Issuerecord.Open "Select Areturndate,Bookid,Issuedate,Returndate,Memid from Issue Order by Memid", Issueconnection, adOpenStatic, adLockOptimistic

Set rmem = New ADODB.Recordset
Set rbook = New ADODB.Recordset
Set riss = New ADODB.Recordset

Call showdata
Call setbutton(True)
Call locktext(True)
Exit Sub

lable:
MsgBox Err.Number & Err.Description
End Sub
Private Sub showdata()
If Issuerecord.EOF = False And Issuerecord.BOF = False Then
'msk_return.Text = Issuerecord.Fields(0)
txt_bookid.Text = Issuerecord.Fields(1)
msk_issue.Text = Format$(Issuerecord.Fields(2), "mm/dd/yyyy")
msk_return.Text = Format$(Issuerecord.Fields(3), "dd/mm/yyyy")
txt_memid.Text = Issuerecord.Fields(4)
End If
Call locate
End Sub
Private Sub cmdFirst_Click()
 On Error GoTo GoFirstError

   Issuerecord.MoveFirst
'show thw current data record
   Call showdata
Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
 On Error GoTo GoLastError
 
   Issuerecord.MoveLast
'show thw current data record
   Call showdata
Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
Dim my As String
On Error GoTo GoNextError
  
  If Not Issuerecord.EOF Then Issuerecord.MoveNext
  If Issuerecord.EOF And Issuerecord.RecordCount > 0 Then
     Beep
     'moved off the end so go back
     Issuerecord.MoveLast
    
  End If
'show thw current data record
     Call showdata
Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
 On Error GoTo GoPrevError
  
  If Not Issuerecord.BOF Then Issuerecord.MovePrevious
  If Issuerecord.BOF And Issuerecord.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    Issuerecord.MovePrevious
 
  End If
'show thw current data record
    Call showdata
Exit Sub

GoPrevError:
   If Err.Number = 3021 Then
MsgBox ("This is first Record."), vbInformation, "First record"
Issuerecord.MoveNext
ElseIf Err.Number <> 0 Then
MsgBox Err.Number & Err.Description
End If
End Sub
