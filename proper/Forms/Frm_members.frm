VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Frm_members 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Member's Details"
   ClientHeight    =   6525
   ClientLeft      =   3765
   ClientTop       =   1845
   ClientWidth     =   9135
   Icon            =   "Frm_members.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_book2 
      Height          =   375
      Left            =   8520
      MouseIcon       =   "Frm_members.frx":24A2
      MousePointer    =   99  'Custom
      Picture         =   "Frm_members.frx":25F4
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   "Members Issued books"
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton cmd_book1 
      Height          =   375
      Left            =   8520
      Picture         =   "Frm_members.frx":29F9
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   6000
      Width           =   495
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   9975
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Detail view"
      TabPicture(0)   =   "Frm_members.frx":2DFD
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DataGrid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Individual view"
      TabPicture(1)   =   "Frm_members.frx":2E19
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Fra_library"
      Tab(1).Control(1)=   "fra_personal"
      Tab(1).Control(2)=   "frm_cmd"
      Tab(1).ControlCount=   3
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   5055
         Left            =   120
         TabIndex        =   49
         Top             =   120
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   8916
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
         Caption         =   "Detail view for Members detail"
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
         Left            =   -74280
         TabIndex        =   29
         Top             =   3960
         Width           =   7455
         Begin VB.CommandButton cmdLast 
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   6960
            MouseIcon       =   "Frm_members.frx":2E35
            MousePointer    =   99  'Custom
            Picture         =   "Frm_members.frx":2F87
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Move Last"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   345
         End
         Begin VB.CommandButton cmdNext 
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   6600
            MouseIcon       =   "Frm_members.frx":31D9
            MousePointer    =   99  'Custom
            Picture         =   "Frm_members.frx":332B
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Move Next"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   345
         End
         Begin VB.CommandButton cmdPrevious 
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   5520
            MouseIcon       =   "Frm_members.frx":3537
            MousePointer    =   99  'Custom
            Picture         =   "Frm_members.frx":3689
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Move Previous"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   345
         End
         Begin VB.CommandButton cmdFirst 
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   5160
            MouseIcon       =   "Frm_members.frx":3898
            MousePointer    =   99  'Custom
            Picture         =   "Frm_members.frx":39EA
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Move First"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   345
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
            MouseIcon       =   "Frm_members.frx":3C39
            MousePointer    =   99  'Custom
            Picture         =   "Frm_members.frx":3D8B
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Add new record"
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
            MouseIcon       =   "Frm_members.frx":4369
            MousePointer    =   99  'Custom
            Picture         =   "Frm_members.frx":44BB
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Edit record"
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
            MouseIcon       =   "Frm_members.frx":4A60
            MousePointer    =   99  'Custom
            Picture         =   "Frm_members.frx":4BB2
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Delete record"
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
            MouseIcon       =   "Frm_members.frx":50FC
            MousePointer    =   99  'Custom
            Picture         =   "Frm_members.frx":524E
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Save record"
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton cmd_cancel 
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
            MouseIcon       =   "Frm_members.frx":57E6
            MousePointer    =   99  'Custom
            Picture         =   "Frm_members.frx":5938
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Cancel"
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton cmd_close 
            Appearance      =   0  'Flat
            Height          =   615
            Left            =   4320
            MouseIcon       =   "Frm_members.frx":5EB8
            MousePointer    =   99  'Custom
            Picture         =   "Frm_members.frx":600A
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Cancel"
            Height          =   255
            Left            =   3480
            TabIndex        =   48
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Save"
            Height          =   255
            Left            =   2640
            TabIndex        =   47
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Delete"
            Height          =   255
            Left            =   1800
            TabIndex        =   46
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Edit"
            Height          =   255
            Left            =   960
            TabIndex        =   45
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "New"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Close"
            Height          =   255
            Left            =   4320
            TabIndex        =   43
            Top             =   840
            Width           =   735
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
            TabIndex        =   42
            Top             =   720
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
            Left            =   5160
            TabIndex        =   41
            Top             =   720
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
            Left            =   6120
            TabIndex        =   40
            Top             =   720
            Width           =   255
         End
      End
      Begin VB.Frame fra_personal 
         Caption         =   "Personal info."
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
         Height          =   2295
         Left            =   -74760
         TabIndex        =   14
         Top             =   120
         Width           =   8415
         Begin VB.ComboBox cmb_sex 
            DataField       =   "Sex"
            ForeColor       =   &H00400000&
            Height          =   315
            ItemData        =   "Frm_members.frx":6584
            Left            =   1680
            List            =   "Frm_members.frx":658E
            Locked          =   -1  'True
            TabIndex        =   21
            Tag             =   "7"
            Top             =   1800
            Width           =   2655
         End
         Begin VB.TextBox txt_phone 
            DataField       =   "Phone"
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            MaxLength       =   25
            TabIndex        =   20
            Tag             =   "6"
            Top             =   1440
            Width           =   2655
         End
         Begin VB.TextBox txt_mail 
            DataField       =   "Email"
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   19
            Tag             =   "5"
            Top             =   1080
            Width           =   2655
         End
         Begin VB.TextBox txt_add 
            DataField       =   "Address"
            ForeColor       =   &H00400000&
            Height          =   1335
            Left            =   5400
            Locked          =   -1  'True
            MaxLength       =   125
            MultiLine       =   -1  'True
            TabIndex        =   18
            Tag             =   "11"
            Top             =   360
            Width           =   2775
         End
         Begin VB.TextBox txt_lname 
            DataField       =   "Lname"
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            MaxLength       =   25
            TabIndex        =   17
            Tag             =   "4"
            Top             =   720
            Width           =   2655
         End
         Begin VB.TextBox txt_fname 
            DataField       =   "Fname"
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            MaxLength       =   25
            TabIndex        =   16
            Tag             =   "3"
            Top             =   360
            Width           =   2655
         End
         Begin MSMask.MaskEdBox msk_bdate 
            Height          =   285
            Left            =   6600
            TabIndex        =   15
            Top             =   1800
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   503
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   10
            Format          =   "mm/dd/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
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
            Left            =   4560
            TabIndex        =   28
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
            TabIndex        =   27
            Top             =   1845
            Width           =   495
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
            TabIndex        =   26
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
            TabIndex        =   25
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
            TabIndex        =   24
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
            TabIndex        =   23
            Top             =   420
            Width           =   1095
         End
         Begin VB.Label lbl_birth 
            BackStyle       =   0  'Transparent
            Caption         =   "Birthdate(mm/dd/yyyy)"
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
            Left            =   4560
            TabIndex        =   22
            Top             =   1840
            Width           =   1935
         End
      End
      Begin VB.Frame Fra_library 
         Caption         =   "Library info."
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
         Height          =   1455
         Left            =   -74760
         TabIndex        =   1
         Top             =   2520
         Width           =   8415
         Begin VB.TextBox txt_deposite 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   6480
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txt_bookhnd 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   6480
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txt_note 
            DataField       =   "Note"
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   2640
            Locked          =   -1  'True
            MaxLength       =   70
            TabIndex        =   5
            Tag             =   "8"
            Top             =   960
            Width           =   2655
         End
         Begin VB.TextBox txt_memid 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   6960
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   960
            Width           =   1215
         End
         Begin MSMask.MaskEdBox msk_expr 
            Height          =   285
            Left            =   2640
            TabIndex        =   2
            Top             =   600
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Format          =   "mm/dd/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox msk_join 
            Height          =   285
            Left            =   2640
            TabIndex        =   3
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   503
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   10
            Format          =   "mm/dd/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
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
            TabIndex        =   13
            Top             =   1000
            Width           =   1215
         End
         Begin VB.Label lbl_join 
            BackStyle       =   0  'Transparent
            Caption         =   "Date of join(mm/dd/yyyy)"
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
            TabIndex        =   12
            Top             =   300
            Width           =   2295
         End
         Begin VB.Label lbl_expr 
            BackStyle       =   0  'Transparent
            Caption         =   "Date of expire(mm/dd/yyyy)"
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
            Top             =   645
            Width           =   2535
         End
         Begin VB.Label lbl_depo 
            BackStyle       =   0  'Transparent
            Caption         =   "Deposits"
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
            Left            =   5160
            TabIndex        =   10
            Top             =   270
            Width           =   855
         End
         Begin VB.Label lbl_bookin 
            BackStyle       =   0  'Transparent
            Caption         =   "Book in hand"
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
            Left            =   5160
            TabIndex        =   9
            Top             =   630
            Width           =   1215
         End
         Begin VB.Label lbl_mid 
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
            Left            =   5760
            TabIndex        =   8
            Top             =   1020
            Width           =   1095
         End
      End
   End
   Begin MSDataGridLib.DataGrid Datagrid 
      Height          =   1335
      Left            =   120
      TabIndex        =   53
      ToolTipText     =   "Detail view of books"
      Top             =   6600
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   2355
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
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
      Caption         =   "Members Issued books"
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
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   $"Frm_members.frx":65A0
      Height          =   495
      Left            =   645
      TabIndex        =   52
      Top             =   120
      Width           =   8295
   End
   Begin VB.Image Image1 
      Height          =   555
      Left            =   120
      Top             =   90
      Width           =   480
   End
End
Attribute VB_Name = "Frm_members"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Memconnection As ADODB.Connection
Dim Memrecordset As ADODB.Recordset
Dim Flexgridset As ADODB.Recordset
Dim temp As ADODB.Recordset
Dim bookshow As Boolean
Dim saveflag As Boolean
Dim lodbook As Boolean
Dim slct As String
Dim str As String
Private Sub clear()
                    txt_add.Text = ""
                    msk_bdate.Text = "__/__/____"
                    txt_bookhnd.Text = ""
                    txt_deposite.Text = ""
                    msk_expr.Text = "__/__/____"
                    msk_join.Text = "__/__/____"
                    txt_mail.Text = ""
                    txt_fname.Text = ""
                    txt_lname.Text = ""
                    txt_memid.Text = ""
                    txt_note.Text = ""
                    txt_phone.Text = ""
                    cmb_sex.Text = ""
End Sub
Private Sub locktext(val As Boolean)
                    txt_add.Locked = val
                    msk_bdate.Enabled = Not val
                    'txt_bookhnd.Locked = val
                    txt_deposite.Locked = val
                    msk_expr.Enabled = Not val
                    msk_join.Enabled = Not val
                    txt_mail.Locked = val
                    txt_fname.Locked = val
                    txt_lname.Locked = val
                    txt_memid.Locked = val
                    txt_note.Locked = val
                    txt_phone.Locked = val
                    cmb_sex.Locked = val
End Sub
Private Sub setbutton(val As Boolean)
               cmd_new.Enabled = val
               cmd_edit.Enabled = val
               cmd_delete.Enabled = val
               cmdFirst.Enabled = val
               cmdLast.Enabled = val
               cmdNext.Enabled = val
               cmdPrevious.Enabled = val
               cmd_cancel.Enabled = Not val
               cmd_save.Enabled = Not val
   
End Sub
Private Function cheak() As Boolean
    Dim flag As Boolean
    flag = False
                 If txt_add.Text = "" Then
                MsgBox "Please enter member's address.", vbInformation, "Information required"
                 ElseIf msk_bdate.Text = "__/__/____" Then
                MsgBox "Please enter member's date of birth.", vbInformation, "Information required"
               '  ElseIf txt_bookhnd.Text = "" Then
               ' MsgBox "Please enter no of books contain by member.", vbInformation, "Information required"
                 ElseIf txt_deposite.Text = "" Then
                MsgBox "Please enter deposite amount.", vbInformation, "Information required"
                 ElseIf msk_expr.Text = "__/__/____" Then
                 MsgBox "Please enter date of account expire.", vbInformation, "Information required"
                ElseIf msk_join.Text = "__/__/____" Then
                MsgBox "Please enter date of join.", vbInformation, "Information required"
                 ElseIf txt_fname.Text = "" Then
                MsgBox "Please enter member's first name.", vbInformation, "Information required"
                 ElseIf txt_lname.Text = "" Then
                MsgBox "Please enter member's last name or family name.", vbInformation, "Information required"
                 ElseIf txt_memid.Text = "" Then
                MsgBox "Please enter member ID no.", vbInformation, "Information required"
                 ElseIf cmb_sex.Text = "" Then
                MsgBox "Please select sex.", vbInformation, "Information required"
                 ElseIf (cmb_sex.Text <> "Male" And cmb_sex.Text <> "Female") Then
                 MsgBox ("Please select the sex."), vbInformation, "Invalid arguments"
                 ElseIf Not IsNumeric(txt_deposite.Text) Then
                 MsgBox ("Deposite must be Numeric value."), vbInformation, "Invalid arguments"
              '   ElseIf Not IsNumeric(txt_bookhnd.Text) Then
              '   MsgBox ("Book in hand must be Numeric."), vbInformation, "Invalid arguments"
                 ElseIf Not IsNumeric(txt_memid.Text) Then
                 MsgBox ("MemberID must be Numeric."), vbInformation, "Invalid arguments"
                 Else
                 flag = True
                End If
cheak = flag
End Function
Private Sub cmd_books_Click()
If (bookshow = True) Then
Me.Height = 6900
cmd_book1.Visible = False
cmd_book2.Visible = True
Else
Me.Height = 8445
cmd_book1.Visible = True
cmd_book2.Visible = False
End If
bookshow = Not bookshow
End Sub
Private Sub cmd_book1_Click()
Call cmd_books_Click
End Sub
Private Sub cmd_book2_Click()
Call cmd_books_Click
End Sub
Private Sub cmd_close_Click()
Unload Me
End Sub
Private Sub cmd_cancel_Click()
On erro GoTo cancelerr
'disablink control
    Call locktext(True)
'    lblStatus.Caption = " Cancel."
 
 If Memrecordset.BOF And Memrecordset.EOF Then
   GoTo newproc
 Else
   Memrecordset.MoveFirst
   Call showdata
 End If

newproc:
  txt_fname.SetFocus
Call setbutton(True)
Exit Sub
cancelerr:
MsgBox Err.Description
End Sub

Private Sub cmd_delete_Click()
On erro GoTo lable
 Beep
str = "select Bookinhand from Member where Memid = " & CDbl(txt_memid.Text)
temp.Open str, Memconnection, adOpenStatic, adLockOptimistic
If temp(0) <> 0 Then
MsgBox "Member account cannot be deleeted because member has not returned books.", vbInformation, "Books not returned"
temp.Close
Exit Sub
End If
temp.Close
If MsgBox("Execution of command will delete current Datarecord,Are you sure you wan't to delete Datarecord ?", vbYesNo + vbExclamation, "Confirm Delete") = vbYes Then
   str = "DELETE FROM Member WHERE "
   str = str & "Memid = "
   str = str & CDbl(txt_memid.Text)
   Memconnection.Execute str
   Memrecordset.Requery
   MsgBox "Record deleted sucessfully.", vbinformayion, "Delete"

If Memrecordset.BOF And Memrecordset.EOF Then
    Call clear
    MsgBox ("The previous record was last record,Now no record left."), vbInformation, "Last record"
    cmd_delete.Enabled = False
Else
   Memrecordset.MoveNext
      If Memrecordset.EOF Then
       Memrecordset.MoveLast
      End If
Call showdata
End If

'message for status of mode
'lblStatus.Caption = " Record deleted."
End If
Exit Sub
lable:
MsgBox Err.Number & Err.Description
End Sub
Private Sub cmd_edit_Click()
Call locktext(False)
Call setbutton(False)
msk_bdate.Enabled = True
msk_expr.Enabled = True
msk_join.Enabled = True
txt_bookhnd.Locked = True
'cmd_cancel.Enabled = False
txt_fname.SetFocus
saveflag = False
'lblStatus.Caption = " Edit record."
End Sub
Private Sub cmd_new_Click()
Call locktext(False)
Call clear
Call setbutton(False)
msk_bdate.Enabled = True
msk_expr.Enabled = True
msk_join.Enabled = True
txt_bookhnd.Text = 0
txt_fname.SetFocus
saveflag = True
'lblStatus.Caption = " Add new record."
End Sub
Private Sub cmd_save_Click()
'error cheaking and autocorrection handle
On Error GoTo errlable
If (cheak = True) Then
    If (txt_note.Text = "") Then
    txt_note.Text = "None"
    End If
    If (txt_phone.Text = "") Then
    txt_phone.Text = "None"
    End If
    If (txt_mail.Text = "") Then
    txt_mail.Text = "None"
    End If
    
If (saveflag = True) Then
           txt_bookhnd.Text = 0
            str = "INSERT INTO Member "
            str = str & "(Address, Birthdate, Bookinhand, Deposite, Doexpire, Dojoin, Email, Fname, Lname, Memid, Noted, Phone, Sex) "
            str = str & "VALUES('" & Trim(txt_add.Text) & "', "
            str = str & "'" & Trim(msk_bdate.Text) & "', "
            str = str & CDbl(txt_bookhnd.Text) & ", "
            str = str & CDbl(Trim(txt_deposite.Text)) & ", "
            str = str & "'" & Trim(msk_expr.Text) & "', "
            str = str & "'" & Trim(msk_join.Text) & "', "
            str = str & "'" & Trim(txt_mail.Text) & "', "
            str = str & "'" & Trim(txt_fname.Text) & "', "
            str = str & "'" & Trim(txt_lname.Text) & "', "
            str = str & CDbl(Trim(txt_memid.Text)) & ", "
            str = str & "'" & Trim(txt_note.Text) & "', "
            str = str & "'" & Trim(txt_phone.Text) & "', "
            str = str & "'" & Trim(cmb_sex.Text) & "' )"
            'MsgBox str
Memconnection.Execute str
Else
            str = "UPDATE Member SET "
            str = str & " Address = '" & Trim(txt_add.Text) & "',"
            str = str & " Birthdate  = '" & Trim(msk_bdate.Text) & "',"
            str = str & " Bookinhand = '" & Trim(txt_bookhnd.Text) & "',"
            str = str & " Deposite = " & CDbl(txt_deposite.Text) & ","
            str = str & " Doexpire = '" & Trim(msk_expr.Text) & "',"
            str = str & " Dojoin = '" & Trim(msk_join.Text) & "',"
            str = str & " Email = '" & Trim(txt_mail.Text) & "',"
            str = str & " Fname = '" & Trim(txt_fname.Text) & "',"
            str = str & " Lname = '" & Trim(txt_lname.Text) & "',"
            str = str & " Memid = " & CDbl(txt_memid.Text) & ","
            str = str & " Noted = '" & Trim(txt_note.Text) & "',"
            str = str & " Phone = '" & Trim(txt_phone.Text) & "',"
            str = str & " Sex = '" & Trim(cmb_sex.Text) & "'"
            str = str & " WHERE Memid= " & CDbl(txt_memid.Text)
            'MsgBox str
Memconnection.Execute str
End If

        Memrecordset.Requery
        Memrecordset.MoveFirst
        MsgBox ("Record saved successfully."), vbInformation, "Save"
        Call locktext(True)
        Call setbutton(True)
        Call showdata
End If
Exit Sub
errlable:
If (Err.Number = -2147467259) Then
MsgBox ("Member ID already exist,please enter anothe ID."), vbCritical, "MemberID exist"
txt_memid.SetFocus
ElseIf (Err.Number = -2147217913) Then
MsgBox ("May be date field pattern wrong."), vbCritical, "Date"
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
Image1.Picture = mdi_start.ImageList1.ListImages(7).Picture
  Set Memconnection = New ADODB.Connection
  Memconnection.CursorLocation = adUseClient
  Memconnection.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Database\Library.mdb;Jet OLEDB:Database Password=Library;"

Set Memrecordset = New ADODB.Recordset
  Memrecordset.Open "select Address,Birthdate,Bookinhand,Deposite,Doexpire,Dojoin,Email,Fname,Lname,Memid,Noted,Phone,Sex from Member Order by Memid", Memconnection, adOpenStatic, adLockOptimistic
Set DataGrid1.DataSource = Memrecordset
  DataGrid1.ReBind

bookshow = False
lodbook = False

Set Flexgridset = New ADODB.Recordset
Set temp = New ADODB.Recordset
  Call showdata
  Call setbutton(True)
msk_bdate.Enabled = False
msk_expr.Enabled = False
msk_join.Enabled = False
cmd_book1.Visible = False
Exit Sub
errlable:
MsgBox Err.Number & Err.Description
End Sub
Private Sub loadbook()
        If Memrecordset.EOF = False And Memrecordset.BOF = False Then
again:
                If (lodbook = False) Then
                Flexgridset.Open "select Author1,Author2,Author3,Bookid,Edition,ISBNNumber,Pages,Price,Publication,Subject,Title,Avano,Issno,Totalno from Book where Bookid in(select Bookid from Issue where Memid=" & Trim(txt_memid.Text) & ")", Memconnection, adOpenStatic, adLockOptimistic
                lodbook = True
                        Set Datagrid.DataSource = Flexgridset
                        Datagrid.ReBind
                Else
                Flexgridset.Close
                lodbook = False
                GoTo again
                End If
        End If
End Sub
Private Sub locate()
  lbl_total.Caption = Memrecordset.RecordCount
  lbl_rec.Caption = Memrecordset.AbsolutePosition
End Sub
Private Sub showdata()
  If Memrecordset.EOF = False And Memrecordset.BOF = False Then
                    txt_add.Text = Memrecordset.Fields(0)
                    msk_bdate.Text = Format$(Memrecordset.Fields(1), "MM/dd/yyyy")
                    txt_bookhnd.Text = Memrecordset.Fields(2)
                    txt_deposite.Text = Memrecordset.Fields(3)
                    msk_expr.Text = Format$(Memrecordset.Fields(4), "MM/dd/yyyy")
                    msk_join.Text = Format$(Memrecordset.Fields(5), "MM/dd/yyyy")
                    txt_mail.Text = Memrecordset.Fields(6)
                    txt_fname.Text = Memrecordset.Fields(7)
                    txt_lname.Text = Memrecordset.Fields(8)
                    txt_memid.Text = Memrecordset.Fields(9)
                    txt_note.Text = Memrecordset.Fields(10)
                    txt_phone.Text = Memrecordset.Fields(11)
                    cmb_sex.Text = Memrecordset.Fields(12)
 End If
 Call locate
 If bookshow Then
    Call loadbook
 End If
 End Sub
Private Sub cmdFirst_Click()
 On Error GoTo GoFirstError

   Memrecordset.MoveFirst
'   lblStatus.Caption = "      <<     Move"
'show thw current data record
   Call showdata
Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
 On Error GoTo GoLastError
 ' lblStatus.Caption = "               Move       >>"

   Memrecordset.MoveLast
'show thw current data record
   Call showdata
Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
Dim my As String
On Error GoTo GoNextError
 'lblStatus.Caption = "               Move       >"
  
  If Not Memrecordset.EOF Then Memrecordset.MoveNext
  If Memrecordset.EOF And Memrecordset.RecordCount > 0 Then
     Beep
     'moved off the end so go back
     Memrecordset.MoveLast
    
  End If
'show thw current data record
     Call showdata
Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
 On Error GoTo GoPrevError
  ' lblStatus.Caption = "      <       Move"

  If Not Memrecordset.BOF Then Memrecordset.MovePrevious
  If Memrecordset.BOF And Memrecordset.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    Memrecordset.MovePrevious
 
  End If
'show thw current data record
    Call showdata
Exit Sub

GoPrevError:
  If Err.Number = 3021 Then
MsgBox ("This is first Record."), vbInformation, "First record"
Memrecordset.MoveNext
ElseIf Err.Number <> 0 Then
MsgBox Err.Number & Err.Description
End If
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
Call locate
Call showdata
End Sub
