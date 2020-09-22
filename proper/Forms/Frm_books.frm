VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Frm_books 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Books Detail"
   ClientHeight    =   6150
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8175
   Icon            =   "Frm_books.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   9551
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Detail view"
      TabPicture(0)   =   "Frm_books.frx":24A2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Datagrid"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Individual view"
      TabPicture(1)   =   "Frm_books.frx":24BE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Fra_self"
      Tab(1).Control(1)=   "txt_title"
      Tab(1).Control(2)=   "txt_publication"
      Tab(1).Control(3)=   "Fra_Author"
      Tab(1).Control(4)=   "txt_isbn"
      Tab(1).Control(5)=   "txt_price"
      Tab(1).Control(6)=   "txt_subject"
      Tab(1).Control(7)=   "txt_pages"
      Tab(1).Control(8)=   "txt_edition"
      Tab(1).Control(9)=   "txt_Bookid"
      Tab(1).Control(10)=   "frm_cmd"
      Tab(1).Control(11)=   "lbl_title"
      Tab(1).Control(12)=   "lbl_pub"
      Tab(1).Control(13)=   "lbl_isbn"
      Tab(1).Control(14)=   "lbl_price"
      Tab(1).Control(15)=   "lbl_subject"
      Tab(1).Control(16)=   "lbl_Pages"
      Tab(1).Control(17)=   "lbl_edition"
      Tab(1).Control(18)=   "lbl_bookid"
      Tab(1).ControlCount=   19
      Begin VB.Frame Fra_self 
         Caption         =   "Copy info."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   1575
         Left            =   -70440
         TabIndex        =   31
         Top             =   960
         Width           =   3135
         Begin VB.TextBox txt_totalno 
            Alignment       =   1  'Right Justify
            DataField       =   "Totalno"
            DataSource      =   "Adodc"
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   34
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txt_issue 
            Alignment       =   1  'Right Justify
            DataField       =   "Issno"
            DataSource      =   "Adodc"
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   33
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txt_avano 
            Alignment       =   1  'Right Justify
            DataField       =   "Avano"
            DataSource      =   "Adodc"
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   32
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lbl_c3 
            BackStyle       =   0  'Transparent
            Caption         =   "Available"
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
            TabIndex        =   37
            Top             =   405
            Width           =   1095
         End
         Begin VB.Label lbl_c2 
            BackStyle       =   0  'Transparent
            Caption         =   "Issued"
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
            TabIndex        =   36
            Top             =   765
            Width           =   615
         End
         Begin VB.Label lbl_c1 
            BackStyle       =   0  'Transparent
            Caption         =   "Total copy"
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
            TabIndex        =   35
            Top             =   1125
            Width           =   975
         End
      End
      Begin VB.TextBox txt_title 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Title"
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   -73560
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   30
         Top             =   240
         Width           =   6255
      End
      Begin VB.TextBox txt_publication 
         DataField       =   "Publication"
         DataSource      =   "Adodc"
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   -73560
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   29
         Top             =   600
         Width           =   6255
      End
      Begin VB.Frame Fra_Author 
         Caption         =   "Author"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   1575
         Left            =   -74760
         TabIndex        =   22
         Top             =   960
         Width           =   4215
         Begin VB.TextBox txt_author1 
            DataField       =   "Author1"
            DataSource      =   "Adodc"
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   25
            Top             =   360
            Width           =   2895
         End
         Begin VB.TextBox txt_author2 
            DataField       =   "Author2"
            DataSource      =   "Adodc"
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   24
            Top             =   720
            Width           =   2895
         End
         Begin VB.TextBox txt_author3 
            DataField       =   "Author3"
            DataSource      =   "Adodc"
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   23
            Top             =   1080
            Width           =   2895
         End
         Begin VB.Label lbl_a1 
            BackStyle       =   0  'Transparent
            Caption         =   "First"
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
            TabIndex        =   28
            Top             =   405
            Width           =   495
         End
         Begin VB.Label lbl_a2 
            BackStyle       =   0  'Transparent
            Caption         =   "Second"
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
            TabIndex        =   27
            Top             =   765
            Width           =   735
         End
         Begin VB.Label lbl_a3 
            BackStyle       =   0  'Transparent
            Caption         =   "Third"
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
            TabIndex        =   26
            Top             =   1125
            Width           =   615
         End
      End
      Begin VB.TextBox txt_isbn 
         DataField       =   "ISBNNumber"
         DataSource      =   "Adodc"
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   -70320
         Locked          =   -1  'True
         MaxLength       =   13
         TabIndex        =   21
         Top             =   3360
         Width           =   3015
      End
      Begin VB.TextBox txt_price 
         Alignment       =   1  'Right Justify
         DataField       =   "Price"
         DataSource      =   "Adodc"
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   -73560
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   20
         Top             =   3360
         Width           =   1815
      End
      Begin VB.TextBox txt_subject 
         DataField       =   "Subject"
         DataSource      =   "Adodc"
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   -70320
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   19
         Top             =   3000
         Width           =   3015
      End
      Begin VB.TextBox txt_pages 
         Alignment       =   1  'Right Justify
         DataField       =   "Pages"
         DataSource      =   "Adodc"
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   -73560
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   18
         Top             =   3000
         Width           =   1815
      End
      Begin VB.TextBox txt_edition 
         DataField       =   "Edition"
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   -70320
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   17
         Top             =   2640
         Width           =   3015
      End
      Begin VB.TextBox txt_Bookid 
         Alignment       =   1  'Right Justify
         DataField       =   "Bookid"
         DataSource      =   "Adodc"
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   -73560
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Frame frm_cmd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   1095
         Left            =   -74760
         TabIndex        =   1
         Top             =   3720
         Width           =   7455
         Begin VB.CommandButton cmd_close 
            Appearance      =   0  'Flat
            Height          =   615
            Left            =   4320
            MouseIcon       =   "Frm_books.frx":24DA
            MousePointer    =   99  'Custom
            Picture         =   "Frm_books.frx":262C
            Style           =   1  'Graphical
            TabIndex        =   47
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton cmdcancel 
            Appearance      =   0  'Flat
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
            MouseIcon       =   "Frm_books.frx":2BA6
            MousePointer    =   99  'Custom
            Picture         =   "Frm_books.frx":2CF8
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Cancel"
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton cmd_save 
            Appearance      =   0  'Flat
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
            Left            =   2640
            MouseIcon       =   "Frm_books.frx":3278
            MousePointer    =   99  'Custom
            Picture         =   "Frm_books.frx":33CA
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Save record"
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton cmd_delete 
            Appearance      =   0  'Flat
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
            Left            =   1800
            MouseIcon       =   "Frm_books.frx":3962
            MousePointer    =   99  'Custom
            Picture         =   "Frm_books.frx":3AB4
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Delete record"
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton cmd_edit 
            Appearance      =   0  'Flat
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
            Left            =   960
            MouseIcon       =   "Frm_books.frx":3FFE
            MousePointer    =   99  'Custom
            Picture         =   "Frm_books.frx":4150
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Edit record"
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton cmd_new 
            Appearance      =   0  'Flat
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
            MaskColor       =   &H8000000F&
            MouseIcon       =   "Frm_books.frx":46F5
            MousePointer    =   99  'Custom
            Picture         =   "Frm_books.frx":4847
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Add new record"
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton cmdFirst 
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   5160
            MouseIcon       =   "Frm_books.frx":4E25
            MousePointer    =   99  'Custom
            Picture         =   "Frm_books.frx":4F77
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Move First"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   345
         End
         Begin VB.CommandButton cmdPrevious 
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   5520
            MouseIcon       =   "Frm_books.frx":51C6
            MousePointer    =   99  'Custom
            Picture         =   "Frm_books.frx":5318
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Move Previous"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   345
         End
         Begin VB.CommandButton cmdNext 
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   6600
            MouseIcon       =   "Frm_books.frx":5527
            MousePointer    =   99  'Custom
            Picture         =   "Frm_books.frx":5679
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Move Next"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   345
         End
         Begin VB.CommandButton cmdLast 
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   6960
            MouseIcon       =   "Frm_books.frx":5885
            MousePointer    =   99  'Custom
            Picture         =   "Frm_books.frx":59D7
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Move Last"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   345
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
            Left            =   6120
            TabIndex        =   52
            Top             =   720
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
            Left            =   5160
            TabIndex        =   51
            Top             =   720
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
            Left            =   6360
            TabIndex        =   50
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Close"
            Height          =   255
            Left            =   4320
            TabIndex        =   48
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "New"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Edit"
            Height          =   255
            Left            =   960
            TabIndex        =   14
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Delete"
            Height          =   255
            Left            =   1800
            TabIndex        =   13
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Save"
            Height          =   255
            Left            =   2640
            TabIndex        =   12
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Cancel"
            Height          =   255
            Left            =   3480
            TabIndex        =   11
            Top             =   840
            Width           =   735
         End
      End
      Begin MSDataGridLib.DataGrid Datagrid 
         Height          =   4815
         Left            =   120
         TabIndex        =   38
         ToolTipText     =   "Detail view of books"
         Top             =   120
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   8493
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
         Caption         =   "Detail view for books"
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
      Begin VB.Label lbl_title 
         BackStyle       =   0  'Transparent
         Caption         =   "Book title "
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
         Left            =   -74640
         TabIndex        =   46
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lbl_pub 
         BackStyle       =   0  'Transparent
         Caption         =   "Publication"
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
         Left            =   -74640
         TabIndex        =   45
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lbl_isbn 
         BackStyle       =   0  'Transparent
         Caption         =   "ISBN no"
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
         Left            =   -71280
         TabIndex        =   44
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label lbl_price 
         BackStyle       =   0  'Transparent
         Caption         =   "Price Rs."
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
         Left            =   -74520
         TabIndex        =   43
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label lbl_subject 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject"
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
         Left            =   -71280
         TabIndex        =   42
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label lbl_Pages 
         BackStyle       =   0  'Transparent
         Caption         =   "Pages"
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
         Left            =   -74520
         TabIndex        =   41
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label lbl_edition 
         BackStyle       =   0  'Transparent
         Caption         =   "Edition"
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
         Left            =   -71280
         TabIndex        =   40
         Top             =   2640
         Width           =   975
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
         Left            =   -74520
         TabIndex        =   39
         Top             =   2640
         Width           =   735
      End
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   495
      Left            =   3600
      TabIndex        =   53
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C8D0D4&
      BackStyle       =   0  'Transparent
      Caption         =   $"Frm_books.frx":5C29
      Height          =   495
      Left            =   720
      TabIndex        =   49
      Top             =   120
      Width           =   7215
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Top             =   90
      Width           =   600
   End
End
Attribute VB_Name = "Frm_books"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Bookrecord As ADODB.Recordset
Dim Bookconnection As ADODB.Connection
Dim str As String
Dim slct As String
Dim saveflag As Boolean
'Function cheaking validity of textbox
Private Function cheak() As Boolean
'declaring variable
   Dim status As Boolean
   status = False
   
      If txt_title.Text = "" Then
          MsgBox ("Please enter the Title."), vbInformation, "Information required"
        ElseIf txt_publication.Text = "" Then
          MsgBox ("Please enter the Publications."), vbInformation, "Information required"
        ElseIf txt_author1.Text = "" Then
          MsgBox ("Please enter the  First Authors name."), vbInformation, "Information required"
        ElseIf txt_bookid.Text = "" Then
          MsgBox ("Please enter bookid distinct from other"), vbInformation, "Information required"
        ElseIf txt_pages.Text = "" Then
          MsgBox ("Please enter no of pages of book."), vbInformation, "Information required"
        ElseIf txt_price.Text = "" Then
          MsgBox ("Please enter the price."), vbInformation, "Information required"
        ElseIf txt_totalno.Text = "" Then
          MsgBox ("Please enter no of copies."), vbInformation, "Information required"
         ElseIf txt_issue.Text = "" Then
          MsgBox ("Please enter no of copies issued."), vbInformation, "Information required"
        ElseIf txt_avano.Text = "" Then
          MsgBox ("Please enter no of copies available."), vbInformation, "Information required"
        ElseIf txt_edition = "" Then
          MsgBox ("Please enter the detail about edition of book."), vbInformation, "Information required"
        ElseIf txt_subject.Text = "" Then
          MsgBox ("Please enter subject related to the book."), vbInformation, "Information required"
        ElseIf txt_isbn.Text = "" Then
          MsgBox ("Please enter ISBN no. for book."), vbInformation, "Information required"
        ElseIf IsNumeric(txt_author1.Text) Then
          MsgBox ("Enter the valid author name."), vbInformation, "Invalid information"
        ElseIf IsNumeric(txt_author2.Text) Then
          MsgBox ("Enter the valid author name."), vbInformation, "Invalid information"
        ElseIf IsNumeric(txt_author3.Text) Then
          MsgBox ("Enter the valid author name."), vbInformation, "Invalid information"
        ElseIf Not IsNumeric(txt_bookid.Text) Then
          MsgBox ("Bookid must be numeric."), vbInformation, "Invalid information"
        ElseIf Not IsNumeric(txt_pages.Text) Then
          MsgBox ("Enter page as in form of string of digits."), vbInformation, "Invalid information"
        ElseIf Not IsNumeric(txt_price.Text) Then
          MsgBox ("Price must be digit form,enter valid price."), vbInformation, "Invalid information"
         ElseIf IsNumeric(txt_edition.Text) Then
          MsgBox ("Enter the valid string for edition."), vbInformation, "Invalid information"
         ElseIf IsNumeric(txt_subject.Text) Then
          MsgBox ("Subject name can not be Numeric."), vbInformation, "Invalid information"
        ElseIf Not IsNumeric(txt_totalno.Text) Then
         MsgBox ("Total no of copy must be Numeric."), vbInformation, "Invalid information"
        ElseIf Not IsNumeric(txt_avano.Text) Then
         MsgBox ("Available no of copy must be Numeric."), vbInformation, "Invalid information"
        ElseIf Not IsNumeric(txt_issue.Text) Then
         MsgBox ("Issue no of copy must be Numeric."), vbInformation, "Invalid information"
        ElseIf Not (CDbl(txt_totalno.Text) = (CDbl(txt_avano.Text) + CDbl(txt_issue.Text))) Then
          MsgBox ("Possibly incorrect data in copy info. frame."), vbInformation, "Invalid information"
        Else
        status = True
        End If
   cheak = status
End Function
'subroutin for setting text box mode
Private Sub setlock(val As Boolean)
     txt_title.Locked = val
     txt_publication.Locked = val
     txt_author1.Locked = val
     txt_author2.Locked = val
     txt_author3.Locked = val
     txt_price.Locked = val
     txt_pages.Locked = val
     txt_subject.Locked = val
     txt_isbn.Locked = val
     txt_totalno.Locked = val
     txt_edition.Locked = val
     txt_bookid.Locked = val
     txt_issue.Locked = val
     txt_avano.Locked = val

End Sub
'make blank the text box
Private Sub clear()
            txt_title.Text = ""
            txt_publication.Text = ""
            txt_author1.Text = ""
            txt_author2.Text = ""
            txt_author3.Text = ""
            txt_price.Text = ""
            txt_subject.Text = ""
            txt_isbn.Text = ""
            txt_pages.Text = ""
            txt_totalno.Text = ""
            txt_avano.Text = ""
            txt_issue.Text = ""
            txt_edition.Text = ""
            txt_bookid.Text = ""

'set focus to fiRSt textbox
            txt_title.SetFocus
End Sub
Private Sub showdata()
  If Bookrecord.EOF = False And Bookrecord.BOF = False Then
          txt_author1.Text = Bookrecord.Fields(0)
          txt_author2.Text = Bookrecord.Fields(1)
          txt_author3.Text = Bookrecord.Fields(2)
          txt_avano.Text = Bookrecord.Fields(3)
          txt_bookid.Text = Bookrecord.Fields(4)
          txt_edition.Text = Bookrecord.Fields(5)
          txt_isbn.Text = Bookrecord.Fields(6)
          txt_issue.Text = Bookrecord.Fields(7)
          txt_pages.Text = Bookrecord.Fields(8)
          txt_price.Text = Bookrecord.Fields(9)
          txt_publication.Text = Bookrecord.Fields(10)
          txt_subject.Text = Bookrecord.Fields(11)
          txt_title.Text = Bookrecord.Fields(12)
          txt_totalno.Text = Bookrecord.Fields(13)
 End If
 Call locate
 End Sub
Private Sub setbutton(val As Boolean)
   cmdFirst.Enabled = val
    cmdPrevious.Enabled = val
    cmdNext.Enabled = val
    cmdLast.Enabled = val
    cmd_delete.Enabled = val
    cmd_edit.Enabled = val
    cmd_new.Enabled = val
    cmd_save.Enabled = Not val
    cmdCancel.Enabled = Not val
End Sub
Private Sub cmd_close_Click()
Unload Me
End Sub
Private Sub cmdCancel_Click()
On erro GoTo cancelerr
'disablink control
    setlock (True)
 
 If Bookrecord.BOF And Bookrecord.EOF Then
   GoTo newproc
 Else
   Bookrecord.MoveFirst
   Call showdata
 End If

newproc:
  txt_title.SetFocus
  Call setbutton(True)
Exit Sub
cancelerr:
MsgBox Err.Description
End Sub
Private Sub cmdFirst_Click()
 On Error GoTo GoFirstError

   Bookrecord.MoveFirst
'show thw current data record
   Call showdata
 
Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub
Private Sub cmdLast_Click()
 On Error GoTo GoLastError
  'lblStatus.Caption = "               Move       >>"

   Bookrecord.MoveLast
'show thw current data record
   Call showdata
Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub
Private Sub cmdNext_Click()
On Error GoTo GoNextError
 'lblStatus.Caption = "               Move       >"
  
  If Not Bookrecord.EOF Then Bookrecord.MoveNext
  If Bookrecord.EOF And Bookrecord.RecordCount > 0 Then
     Beep
     'moved off the end so go back
     Bookrecord.MoveLast
    
  End If
'show thw current data record
     Call showdata
  
Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
 On Error GoTo GoPrevError
 '  lblStatus.Caption = "      <       Move"

  If Not Bookrecord.BOF Then Bookrecord.MovePrevious
  If Bookrecord.BOF And Bookrecord.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    Bookrecord.MovePrevious
 
  End If
'show thw current data record
    Call showdata
 Exit Sub

GoPrevError:
If Err.Number = 3021 Then
MsgBox ("This is first Record."), vbInformation, "First record"
Bookrecord.MoveNext
ElseIf Err.Number <> 0 Then
MsgBox Err.Number & Err.Description
End If
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
   
Image1.Picture = mdi_start.ImageList1.ListImages(1).Picture
   Set Bookconnection = New ADODB.Connection
   Bookconnection.CursorLocation = adUseClient
   Bookconnection.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Database\Library.mdb;Jet OLEDB:Database Password=Library;"
   slct = "select Author1,Author2,Author3,Avano,Bookid,Edition,ISBNNumber,Issno,Pages,Price,Publication,Subject,Title,Totalno from Book Order by Bookid"
     Set Bookrecord = New ADODB.Recordset
   Bookrecord.Open slct, Bookconnection, adOpenStatic, adLockOptimistic

'show current record
Call showdata
Set Datagrid.DataSource = Bookrecord
Datagrid.ReBind
'disable buttons
  cmd_save.Enabled = False
  cmdCancel.Enabled = False
Exit Sub
errlable:
MsgBox Err.Number & Err.Description
End Sub
Private Sub locate()
  lbl_total.Caption = Bookrecord.RecordCount
  lbl_rec.Caption = Bookrecord.AbsolutePosition
End Sub
Private Sub cmd_delete_Click()
On erro GoTo lable
 Beep
If MsgBox("Execution of command will delete current Datarecord,Are you sure you wan't to delete Datarecord ?", vbYesNo + vbExclamation, "Confirm Delete") = vbYes Then
   str = "DELETE FROM Book WHERE "
   str = str & "Bookid = "
   str = str & CDbl(txt_bookid.Text)
   Bookconnection.Execute str
   Bookrecord.Requery
   MsgBox "Record deleted sucessfully.", vbinformayion, "Delete"

If Bookrecord.BOF And Bookrecord.EOF Then
    Call clear
    MsgBox ("The previous record was last record,Now no record left."), vbInformation, "Last record"
    cmd_delete.Enabled = False
Else
   Bookrecord.MoveNext
      If Bookrecord.EOF Then
       Bookrecord.MoveLast
      End If
   Call showdata
End If
'message for status of mode
'lblStatus.Caption = "Record deleted."
End If
Exit Sub
lable:
MsgBox Err.Number & Err.Description
End Sub
Private Sub cmd_edit_Click()
On Error GoTo lable

'Make all entries in input mode
     Call setlock(False)
     saveflag = False
'message for status of mode
     '      lblStatus.Caption = " Edit record"
   Call setbutton(False)
  ' cmdcancel.Enabled = False
'set focus
            txt_title.SetFocus
Exit Sub
lable:
MsgBox Err.Description
End Sub
Private Sub cmd_new_Click()
On Error GoTo lable

'Make all entries in input mode enable
     Call setlock(False)
 'clear the text field
     Call clear
    saveflag = True
    'lblStatus.Caption = " Add new record."
       
Call setbutton(False)
Exit Sub
lable:
'Error handling statement
MsgBox Err.Description
End Sub
Private Sub cmd_save_Click()
On Error GoTo lable
'Make all entries in input mode enable
          Call setlock(False)
 'cheaking for validity condition
            If cheak = True Then
              If txt_author2.Text = "" Then
                txt_author2.Text = "None"
               End If
              If txt_author3.Text = "" Then
                txt_author3.Text = "None"
               End If

'saving new record
If saveflag = True Then
str = "INSERT INTO Book"
str = str & "(Author1, Author2, Author3, Avano, Bookid, Edition, ISBNNumber, Issno, Pages, Price, Publication, Subject, Title, Totalno) "
str = str & "VALUES('" & Trim(txt_author1.Text) & "', "
str = str & "'" & Trim(txt_author2.Text) & "', "
str = str & "'" & Trim(txt_author3.Text) & "', "
str = str & CDbl(txt_avano.Text) & ", "
str = str & CDbl(txt_bookid.Text) & ", "
str = str & "'" & Trim(txt_edition.Text) & "', "
str = str & "'" & Trim(txt_isbn.Text) & "', "
str = str & CDbl(txt_issue.Text) & ", "
str = str & CDbl(txt_pages.Text) & ", "
str = str & CDbl(txt_price.Text) & ", "
str = str & "'" & Trim(txt_publication.Text) & "', "
str = str & "'" & Trim(txt_subject.Text) & "', "
str = str & "'" & Trim(txt_title.Text) & "', "
str = str & CDbl(txt_totalno.Text) & ")"
Bookconnection.Execute str
Else 'for editing the record
str = "UPDATE Book SET "
str = str & "Author1='" & Trim(txt_author1.Text) & "',"
str = str & "Author2='" & Trim(txt_author2.Text) & "',"
str = str & "Author3='" & Trim(txt_author3.Text) & "',"
str = str & "Avano=" & CDbl(txt_avano.Text) & ","
str = str & "Bookid=" & CDbl(txt_bookid.Text) & ","
str = str & "Edition='" & Trim(txt_edition.Text) & "',"
str = str & "ISBNNumber='" & Trim(txt_isbn.Text) & "',"
str = str & "Issno=" & CDbl(txt_issue.Text) & ","
str = str & "Pages=" & CDbl(txt_pages.Text) & ","
str = str & "Price=" & CDbl(txt_price.Text) & ","
str = str & "Publication='" & Trim(txt_publication.Text) & "',"
str = str & "Subject='" & Trim(txt_subject.Text) & "',"
str = str & "Title='" & Trim(txt_title.Text) & "',"
str = str & "Totalno=" & CDbl(txt_totalno.Text)
str = str & " WHERE Bookid=" & CDbl(txt_bookid.Text)
Bookconnection.Execute str
End If
'Make all entries input mode disable
Call setlock(True)

Bookrecord.Requery
Bookrecord.MoveLast
'show thw current data record
Call showdata
 'message for status of mode
           'lblStatus.Caption = " New record Saved."
           MsgBox ("Record has been suceefully saved."), vbInformation, "Saving Record"
Call setbutton(True)
End If
Exit Sub
lable:
If Err.Number = -2147467259 Then
MsgBox ("BookID already exist,please enter anothe ID."), vbCritical, "BookID exist"
txt_bookid.SetFocus
ElseIf Err.Number <> 0 Then
MsgBox Err.Number & Err.Description
End If
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
Call locate
Call showdata
End Sub
