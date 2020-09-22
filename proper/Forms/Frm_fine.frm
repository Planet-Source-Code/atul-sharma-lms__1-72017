VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Frm_Fine 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fine Information"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   Icon            =   "Frm_fine.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   4695
   Begin VB.CommandButton cmdLast 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   1920
      MouseIcon       =   "Frm_fine.frx":24A2
      MousePointer    =   99  'Custom
      Picture         =   "Frm_fine.frx":25F4
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Move Last"
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton cmdNext 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   1560
      MouseIcon       =   "Frm_fine.frx":2846
      MousePointer    =   99  'Custom
      Picture         =   "Frm_fine.frx":2998
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Move Next"
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton cmdPrevious 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   480
      MouseIcon       =   "Frm_fine.frx":2BA4
      MousePointer    =   99  'Custom
      Picture         =   "Frm_fine.frx":2CF6
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Move Previous"
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton cmdFirst 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   120
      MouseIcon       =   "Frm_fine.frx":2F05
      MousePointer    =   99  'Custom
      Picture         =   "Frm_fine.frx":3057
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Move First"
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   345
   End
   Begin VB.CommandButton cmd_back 
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
      Left            =   3480
      MouseIcon       =   "Frm_fine.frx":32A6
      MousePointer    =   99  'Custom
      Picture         =   "Frm_fine.frx":33F8
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Back to Returnform"
      Top             =   3240
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid Datagrid 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "All fine Information"
      Top             =   600
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4471
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483633
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "Memid"
         Caption         =   "MemberID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Bookid"
         Caption         =   "BookID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Fine"
         Caption         =   "Fine Amount"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Areturndate"
         Caption         =   "Return date"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   824.882
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   750.047
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1470.047
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Back to Return"
      Height          =   255
      Left            =   3480
      TabIndex        =   10
      Top             =   3855
      Width           =   1095
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
      Left            =   1320
      TabIndex        =   9
      Top             =   3720
      Width           =   975
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
      Left            =   120
      TabIndex        =   8
      Top             =   3720
      Width           =   855
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
      Left            =   1080
      TabIndex        =   7
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fine Informations can be deleted by administrator from 'Administer settings/Delete fine'."
      Height          =   495
      Left            =   735
      TabIndex        =   2
      Top             =   120
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Top             =   75
      Width           =   600
   End
End
Attribute VB_Name = "Frm_Fine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str As String
Dim Fineconn As ADODB.Connection
Dim Finerecord As ADODB.Recordset
Private Sub cmd_back_Click()
Load Frm_return
Frm_return.Show
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
Image1.Picture = mdi_start.ImageList1.ListImages(14).Picture
Set Fineconn = New ADODB.Connection
Fineconn.CursorLocation = adUseClient
Fineconn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & "Data source=" & App.Path & "\Database\Library.mdb;Jet OLEDB:Database Password=Library;"

str = "Select Memid,Bookid,Fine,Areturndate from Fine order by Memid"
Set Finerecord = New ADODB.Recordset
Finerecord.Open str, Fineconn, adOpenStatic, adLockOptimistic
            Datagrid.Visible = True
            Set Datagrid.DataSource = Finerecord
            Datagrid.ReBind
Call locate
Exit Sub
errlable:
MsgBox Err.Number & Err.Description
End Sub
Private Sub cmdFirst_Click()
 On Error GoTo GoFirstError

   Finerecord.MoveFirst
'show thw current data record
Call locate
Exit Sub
GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
 On Error GoTo GoLastError

   Finerecord.MoveLast
'show thw current data record
Call locate
Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
On Error GoTo GoNextError
  
  If Not Finerecord.EOF Then Finerecord.MoveNext
  If Finerecord.EOF And Finerecord.RecordCount > 0 Then
     Beep
     'moved off the end so go back
     Finerecord.MoveLast
    
  End If
'show thw current data record
 
Call locate
Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub
Private Sub locate()
  lbl_total.Caption = Finerecord.RecordCount
  lbl_rec.Caption = Finerecord.AbsolutePosition
End Sub
Private Sub cmdPrevious_Click()
 On Error GoTo GoPrevError

  If Not Finerecord.BOF Then Finerecord.MovePrevious
  If Finerecord.BOF And Finerecord.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    Finerecord.MovePrevious
 
  End If
'show thw current data record
Call locate
Exit Sub

GoPrevError:
 If Err.Number = 3021 Then
MsgBox ("This is first Record."), vbInformation, "First record"
Finerecord.MoveNext
ElseIf Err.Number <> 0 Then
MsgBox Err.Number & Err.Description
End If
End Sub

