VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Frm_search 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search.."
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9825
   Icon            =   "Frm_search.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   9825
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   4455
      Left            =   120
      TabIndex        =   1
      Top             =   520
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   7858
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   529
      TabMaxWidth     =   2646
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Books"
      TabPicture(0)   =   "Frm_search.frx":24A2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "bdatagrid"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra_search"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "bpbar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Members"
      TabPicture(1)   =   "Frm_search.frx":24BE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "mdatagrid"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fra_msearch"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "mpbar"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin MSComctlLib.ProgressBar mpbar 
         Height          =   375
         Left            =   -68400
         TabIndex        =   15
         Top             =   1200
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar bpbar 
         Height          =   375
         Left            =   6600
         TabIndex        =   14
         Top             =   1200
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Frame fra_msearch 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   1095
         Left            =   -74760
         TabIndex        =   8
         Top             =   600
         Width           =   6135
         Begin VB.TextBox txt_mvalue 
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   1560
            TabIndex        =   11
            ToolTipText     =   "Value for field"
            Top             =   720
            Width           =   3135
         End
         Begin VB.ComboBox cmb_mfield 
            ForeColor       =   &H00400000&
            Height          =   315
            ItemData        =   "Frm_search.frx":24DA
            Left            =   1560
            List            =   "Frm_search.frx":24F9
            TabIndex        =   10
            ToolTipText     =   "Select Member field"
            Top             =   240
            Width           =   3135
         End
         Begin VB.CommandButton txt_msearch 
            Caption         =   "Search"
            Height          =   735
            Left            =   4920
            MouseIcon       =   "Frm_search.frx":2555
            MousePointer    =   99  'Custom
            Picture         =   "Frm_search.frx":26A7
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Search"
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lbl_values 
            BackColor       =   &H00C8D0D4&
            BackStyle       =   0  'Transparent
            Caption         =   "Value"
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
            Left            =   720
            TabIndex        =   13
            Top             =   750
            Width           =   615
         End
         Begin VB.Label lbl_fields 
            BackColor       =   &H00C8D0D4&
            BackStyle       =   0  'Transparent
            Caption         =   "Field"
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
            Left            =   720
            TabIndex        =   12
            Top             =   280
            Width           =   615
         End
      End
      Begin VB.Frame fra_search 
         Caption         =   "Search"
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
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   6135
         Begin VB.CommandButton cmd_bsearch 
            Caption         =   "Search"
            Height          =   735
            Left            =   4920
            MouseIcon       =   "Frm_search.frx":2CB0
            MousePointer    =   99  'Custom
            Picture         =   "Frm_search.frx":2E02
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Search"
            Top             =   240
            Width           =   975
         End
         Begin VB.ComboBox cmb_bfield 
            ForeColor       =   &H00400000&
            Height          =   315
            ItemData        =   "Frm_search.frx":340B
            Left            =   1560
            List            =   "Frm_search.frx":342D
            TabIndex        =   4
            ToolTipText     =   "Select book's field"
            Top             =   240
            Width           =   3135
         End
         Begin VB.TextBox txt_bvalue 
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   1560
            TabIndex        =   3
            ToolTipText     =   "Value for search"
            Top             =   720
            Width           =   3135
         End
         Begin VB.Label lbl_field 
            BackStyle       =   0  'Transparent
            Caption         =   "Field"
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
            Left            =   720
            TabIndex        =   7
            Top             =   280
            Width           =   615
         End
         Begin VB.Label lbl_value 
            BackStyle       =   0  'Transparent
            Caption         =   "Value"
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
            Left            =   720
            TabIndex        =   6
            Top             =   750
            Width           =   615
         End
      End
      Begin MSDataGridLib.DataGrid bdatagrid 
         Height          =   2415
         Left            =   240
         TabIndex        =   16
         ToolTipText     =   "Detail view of books"
         Top             =   1800
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   4260
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         Appearance      =   0
         BackColor       =   -2147483633
         DefColWidth     =   7
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
         Caption         =   "Search result for books"
         ColumnCount     =   14
         BeginProperty Column00 
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
         BeginProperty Column01 
            DataField       =   "Title"
            Caption         =   "Book Title"
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
            DataField       =   "Author1"
            Caption         =   "Author1"
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
            DataField       =   "Author2"
            Caption         =   "Author2"
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
         BeginProperty Column04 
            DataField       =   "Author3"
            Caption         =   "Author3"
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
         BeginProperty Column05 
            DataField       =   "Avano"
            Caption         =   "Available"
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
         BeginProperty Column06 
            DataField       =   "Issno"
            Caption         =   "Issue"
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
         BeginProperty Column07 
            DataField       =   "Totalno"
            Caption         =   "Total"
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
         BeginProperty Column08 
            DataField       =   "Edition"
            Caption         =   "Edition"
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
         BeginProperty Column09 
            DataField       =   "ISBNNumber"
            Caption         =   "ISBNNumber"
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
         BeginProperty Column10 
            DataField       =   "Pages"
            Caption         =   "Pages"
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
         BeginProperty Column11 
            DataField       =   "Price"
            Caption         =   "Price"
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
         BeginProperty Column12 
            DataField       =   "Subject"
            Caption         =   "Subject"
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
         BeginProperty Column13 
            DataField       =   "Publication"
            Caption         =   "Publication"
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
            SizeMode        =   1
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
               ColumnWidth     =   645.165
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2835.213
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1709.858
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1590.236
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   659.906
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   629.858
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1335.118
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   1725.165
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   2624.882
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid mdatagrid 
         Height          =   2415
         Left            =   -74760
         TabIndex        =   17
         Top             =   1800
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   4260
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
         Caption         =   "Search result for Member"
         ColumnCount     =   12
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
            DataField       =   "Fname"
            Caption         =   "First name"
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
            DataField       =   "Lname"
            Caption         =   "Last name"
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
            DataField       =   "Address"
            Caption         =   "Address"
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
         BeginProperty Column04 
            DataField       =   "Phone"
            Caption         =   "Phone"
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
         BeginProperty Column05 
            DataField       =   "Email"
            Caption         =   "Email"
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
         BeginProperty Column06 
            DataField       =   "Deposite"
            Caption         =   "Deposite"
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
         BeginProperty Column07 
            DataField       =   "Birthdate"
            Caption         =   "Birthday"
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
         BeginProperty Column08 
            DataField       =   "Dojoin"
            Caption         =   "Join at"
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
         BeginProperty Column09 
            DataField       =   "Doexpire"
            Caption         =   "Expire at"
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
         BeginProperty Column10 
            DataField       =   "Sex"
            Caption         =   "Sex"
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
         BeginProperty Column11 
            DataField       =   "Noted"
            Caption         =   "Note"
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
               ColumnWidth     =   1604.976
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1620.284
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   3404.977
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1665.071
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   2505.26
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1170.142
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1230.236
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   734.74
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   2429.858
            EndProperty
         EndProperty
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Top             =   0
      Width           =   495
   End
   Begin VB.Label lbl_status 
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
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   8895
   End
End
Attribute VB_Name = "Frm_search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fnd As String
Dim mflag As Boolean
Dim bflag As Boolean
Dim conn As ADODB.Connection
Dim MR As ADODB.Recordset
Dim BR As ADODB.Recordset



Private Sub cmb_bfield_click()
If (cmb_bfield.Text = "All") Then
txt_bvalue.Enabled = False
Else
txt_bvalue.Enabled = True
End If
lbl_status.Caption = " Search for book's Record field."
txt_bvalue.Text = ""
End Sub

Private Sub cmb_mfield_click()
If (cmb_mfield.Text = "All") Then
txt_mvalue.Enabled = False
Else
txt_mvalue.Enabled = True
End If
txt_mvalue.Text = ""
lbl_status.Caption = " Search for Member's Record field."
End Sub
Private Sub cmd_bsearch_Click()
On eror GoTo errlable:
'write code for validity
again:
bpbar.Value = 0
If (cmb_bfield.Text = "All" Or txt_bvalue.Text = "") Then
fnd = "select Author1,Author2,Author3,Avano,Bookid,Edition,ISBNNumber,Issno,Pages,Price,Publication,Subject,Title,Totalno from Book order by Bookid"
lbl_status.Caption = " Search for Book's Record field Alldata."
bpbar.Value = 30
ElseIf (cmb_bfield.Text = "Author") Then
fnd = "select Author1,Author2,Author3,Avano,Bookid,Edition,ISBNNumber,Issno,Pages,Price,Publication,Subject,Title,Totalno from Book where Author1 like'" & Trim(txt_bvalue.Text) & "%' or Author2 like'" & Trim(txt_bvalue.Text) & "%' or Author3 like'" & Trim(txt_bvalue.Text) & "%'"
lbl_status.Caption = " Search for Book's Record field Author."
bpbar.Value = 30
ElseIf (cmb_bfield.Text = "Price" Or cmb_bfield.Text = "Pages" Or cmb_bfield.Text = "Bookid") Then
    If IsNumeric(txt_bvalue.Text) Then
    fnd = "select Author1,Author2,Author3,Avano,Bookid,Edition,ISBNNumber,Issno,Pages,Price,Publication,Subject,Title,Totalno from Book where " & Trim(cmb_bfield) & " = " & Trim(txt_bvalue)
    lbl_status.Caption = " Search for Book's Record field " & Trim(cmb_bfield.Text) & " of book."
    bpbar.Value = 30
    Else
    txt_bvalue.Text = ""
    Exit Sub
    End If
    
Else
fnd = "select Author1,Author2,Author3,Avano,Bookid,Edition,ISBNNumber,Issno,Pages,Price,Publication,Subject,Title,Totalno from Book where " & Trim(cmb_bfield) & " like '" & Trim(txt_bvalue) & "%'"
lbl_status.Caption = " Search for Book's Record field " & Trim(cmb_bfield.Text) & " of book."
bpbar.Value = 30
End If
 If (bflag = False) Then
            BR.Open fnd, conn, adOpenStatic, adLockOptimistic
            bpbar.Value = 50
            bdatagrid.Visible = True
            Set bdatagrid.DataSource = BR
            bpbar.Value = 70
            bdatagrid.ReBind
            bflag = True
            bpbar.Value = 85
            Else
            bflag = False
            BR.Close
            GoTo again
            bpbar.Value = 90
              End If
bpbar.Value = 100
bpbar.Value = 0
Exit Sub
errlable:
bpbar.Value = 0
MsgBox Err.Description
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
Image1.Picture = mdi_start.ImageList1.ListImages(2).Picture
 Set conn = New ADODB.Connection
 conn.CursorLocation = adUseClient
 conn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Database\Library.mdb;Jet OLEDB:Database Password=Library;"
Set BR = New ADODB.Recordset
Set MR = New ADODB.Recordset
 lbl_status.Caption = " Choose the options for Datamember,Field and values for search."
Exit Sub
errlable:
MsgBox Err.Number & "  " & Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub


Private Sub txt_msearch_Click()
'write a code validity
On Error GoTo errlable
again:
 mpbar.Value = 0
 lbl_status.Caption = " Search for Member's Record field " & Trim(cmb_mfield.Text) & " of Member."
If (cmb_mfield.Text = "All" Or txt_mvalue.Text = "") Then
 fnd = "select Address,Birthdate,Bookinhand,Deposite,Doexpire,Dojoin,Email,Fname,Lname,Memid,Noted,Phone,Sex from Member Order by Memid"
 lbl_status.Caption = " Search for Member's Record field Alldata."
mpbar.Value = 40
ElseIf (cmb_mfield.Text = "First name") Then
 fnd = "select Address,Birthdate,Bookinhand,Deposite,Doexpire,Dojoin,Email,Fname,Lname,Memid,Noted,Phone,Sex from Member where Fname like '" & Trim(txt_mvalue.Text) & "%'"
mpbar.Value = 40
ElseIf (cmb_mfield.Text = "Last name") Then
 fnd = "select Address,Birthdate,Bookinhand,Deposite,Doexpire,Dojoin,Email,Fname,Lname,Memid,Noted,Phone,Sex from Member where Lname like '" & Trim(txt_mvalue.Text) & "%'"
mpbar.Value = 40
ElseIf (cmb_mfield.Text = "Member id") Then
 fnd = "select Address,Birthdate,Bookinhand,Deposite,Doexpire,Dojoin,Email,Fname,Lname,Memid,Noted,Phone,Sex from Member where Memid=" & Trim(txt_mvalue.Text)
mpbar.Value = 40
ElseIf (cmb_mfield.Text = "Address") Then
 fnd = "select Address,Birthdate,Bookinhand,Deposite,Doexpire,Dojoin,Email,Fname,Lname,Memid,Noted,Phone,Sex from Member where Address like '" & Trim(txt_mvalue.Text) & "%'"
mpbar.Value = 40
ElseIf (cmb_mfield.Text = "Phone") Then
 fnd = "select Address,Birthdate,Bookinhand,Deposite,Doexpire,Dojoin,Email,Fname,Lname,Memid,Noted,Phone,Sex from Member where Phone like'" & Trim(txt_mvalue.Text) & "%'"
mpbar.Value = 40
ElseIf (cmb_mfield.Text = "Email") Then
 fnd = "select Address,Birthdate,Bookinhand,Deposite,Doexpire,Dojoin,Email,Fname,Lname,Memid,Noted,Phone,Sex from Member where Email like'" & Trim(txt_mvalue.Text) & "%'"
mpbar.Value = 40
ElseIf (cmb_mfield.Text = "Birth date") Then
 fnd = "select Address,Birthdate,Bookinhand,Deposite,Doexpire,Dojoin,Email,Fname,Lname,Memid,Noted,Phone,Sex from Member where Birthdate like'" & Trim(txt_mvalue.Text) & "%'"
mpbar.Value = 40
ElseIf (cmb_mfield.Text = "Date of join") Then
 fnd = "select Address,Birthdate,Bookinhand,Deposite,Doexpire,Dojoin,Email,Fname,Lname,Memid,Noted,Phone,Sex from Member where Dojoin like'" & Trim(txt_mvalue.Text) & "%'"
mpbar.Value = 40
End If
 If (mflag = False) Then
            MR.Open fnd, conn, adOpenStatic, adLockOptimistic
            mpbar.Value = 65
            mdatagrid.Visible = True
            Set mdatagrid.DataSource = MR
            mpbar.Value = 80
            mdatagrid.ReBind
            mflag = True
            mpbar.Value = 90
            Else
            mflag = False
            MR.Close
            GoTo again
              End If
mpbar.Value = 100
mpbar.Value = 0
Exit Sub
errlable:
mpbar.Value = 0
MsgBox Err.Number & "  " & Err.Description
End Sub
