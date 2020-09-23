VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRecon 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Reconciliation"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12855
   Icon            =   "frmRecon.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   12855
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   9375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14775
      Begin VB.TextBox TxtClr 
         Height          =   285
         Left            =   0
         TabIndex        =   31
         Text            =   "Yes"
         Top             =   4920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TxtUnc 
         Height          =   285
         Left            =   0
         TabIndex        =   28
         Text            =   "No"
         Top             =   5880
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox CboBankName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2880
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   720
         Width           =   4695
      End
      Begin VB.ComboBox CboAccountName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2880
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   240
         Width           =   4695
      End
      Begin VB.Frame FrmConfirm 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2415
         Left            =   7680
         TabIndex        =   9
         Top             =   240
         Visible         =   0   'False
         Width           =   6735
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   4080
            TabIndex        =   37
            Top             =   1320
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   64880641
            CurrentDate     =   39567
         End
         Begin LVbuttons.LaVolpeButton CmdConfirm 
            Height          =   495
            Left            =   3600
            TabIndex        =   34
            Top             =   1800
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "Clear Transaction"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            BCOL            =   14215660
            FCOL            =   0
            FCOLO           =   0
            EMBOSSM         =   12632256
            EMBOSSS         =   16777215
            MPTR            =   0
            MICON           =   "frmRecon.frx":0442
            ALIGN           =   1
            IMGLST          =   "(None)"
            IMGICON         =   "(None)"
            ICONAlign       =   0
            ORIENT          =   0
            STYLE           =   0
            IconSize        =   2
            SHOWF           =   -1  'True
            BSTYLE          =   0
         End
         Begin VB.TextBox TxtSrNo 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3000
            TabIndex        =   16
            Top             =   480
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox TxtDate 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   15
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox TxtChequeNo 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   14
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox TxtParticulars 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   1440
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   13
            Top             =   960
            Width           =   2535
         End
         Begin VB.TextBox TxtDebit 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   12
            Top             =   1560
            Width           =   1335
         End
         Begin VB.TextBox TxtCredit 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   11
            Top             =   1920
            Width           =   1335
         End
         Begin VB.TextBox TxtRemarks 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   4080
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Top             =   480
            Width           =   2535
         End
         Begin LVbuttons.LaVolpeButton CmdCancel 
            Height          =   495
            Left            =   5160
            TabIndex        =   35
            Top             =   1800
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "&Cancel"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            BCOL            =   14215660
            FCOL            =   0
            FCOLO           =   0
            EMBOSSM         =   12632256
            EMBOSSS         =   16777215
            MPTR            =   0
            MICON           =   "frmRecon.frx":045E
            ALIGN           =   1
            IMGLST          =   "(None)"
            IMGICON         =   "(None)"
            ICONAlign       =   0
            ORIENT          =   0
            STYLE           =   0
            IconSize        =   2
            SHOWF           =   -1  'True
            BSTYLE          =   0
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cleared Date"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   4080
            TabIndex        =   36
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Date : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   915
            TabIndex        =   22
            Top             =   240
            Width           =   540
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cheque No. : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   390
            TabIndex        =   21
            Top             =   600
            Width           =   1065
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Particulars : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   405
            TabIndex        =   20
            Top             =   960
            Width           =   1050
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Debit Amount : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   150
            TabIndex        =   19
            Top             =   1560
            Width           =   1305
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Credit Amount : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   4
            Left            =   90
            TabIndex        =   18
            Top             =   1920
            Width           =   1365
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Remarks : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   5
            Left            =   4080
            TabIndex        =   17
            Top             =   240
            Width           =   900
         End
      End
      Begin VB.TextBox TxtAccountNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1215
         Width           =   4695
      End
      Begin VB.TextBox TxtBankBalance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1710
         Width           =   2415
      End
      Begin VB.TextBox TxtBookBalance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2205
         Width           =   2415
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   240
         Top             =   7800
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRecon.frx":047A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3135
         Left            =   240
         TabIndex        =   23
         Top             =   2880
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   5530
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "###"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Cheque No."
            Object.Width           =   2558
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Particulars"
            Object.Width           =   6526
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Debit"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Credit"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Closing"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Remarks"
            Object.Width           =   5644
         EndProperty
      End
      Begin MSComctlLib.ListView UnClearedList 
         Height          =   2655
         Left            =   240
         TabIndex        =   24
         Top             =   6480
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "###"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Cheque No."
            Object.Width           =   2558
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Particulars"
            Object.Width           =   6526
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Debit"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Credit"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Remarks"
            Object.Width           =   8114
         EndProperty
      End
      Begin VB.Label LblDebitBank 
         Caption         =   "Label6"
         Height          =   255
         Left            =   7320
         TabIndex        =   33
         Top             =   1800
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label LblCreditBank 
         Caption         =   "Label6"
         Height          =   255
         Left            =   7320
         TabIndex        =   32
         Top             =   2160
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label LblCredit 
         Caption         =   "Label6"
         Height          =   255
         Left            =   5400
         TabIndex        =   30
         Top             =   2160
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label LblDebit 
         Caption         =   "Label6"
         Height          =   255
         Left            =   5400
         TabIndex        =   29
         Top             =   1800
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         Caption         =   "- : Uncleared Transactions : -"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   240
         TabIndex        =   25
         Top             =   6120
         Width           =   14295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000080&
         Caption         =   "Balance as per Book :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   2280
         Width           =   2655
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000080&
         Caption         =   "Balance as per Bank :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1770
         Width           =   2655
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000080&
         Caption         =   "Account No :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1260
         Width           =   2655
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000080&
         Caption         =   "Account Name :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000080&
         Caption         =   "Bank :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   750
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmRecon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TempRs As New ADODB.Recordset
Dim TempRs1 As New ADODB.Recordset
Dim TempRs3 As New ADODB.Recordset
Dim TempRs4 As New ADODB.Recordset

Private Sub CboAccountName_Click()
    GetBankName
End Sub

Private Sub CboBankName_Click()
    GetAccNo
    FillDatainList
    FillUncleared
    BookBalance
    BankBalance
End Sub

Private Sub CmdCancel_Click()
    FrmConfirm.Visible = False
End Sub

Private Sub CmdConfirm_Click()
    If MsgBox("Are you sure you want to Clear this Transaction ?", vbQuestion + vbYesNo) = vbYes Then
    str9 = "Update Transactions set Clear='" & TxtClr.Text & "',ClearDate='" & DTPicker1.Value & "' where SrNo=" & Val(TxtSrNo.Text) & ""
    GConn.Execute str9
    MsgBox "Transaction Cleared Successfully", vbInformation
    FrmConfirm.Visible = False
    BankBalance
    BookBalance
    FillUncleared
    End If
End Sub

Private Sub Form_Load()
    Frame1.Left = (Screen.Width - Frame1.Width) / 2
    Frame1.Top = (Screen.Height - Frame1.Height) / 12
    
    Dim strd As String
    strd = "Select SrNo,AccountName From Accounts"
    Call FillCombo(CboAccountName, strd)
    
    strd = "Select SrNo,BankName From Accounts"
    Call FillCombo(CboBankName, strd)

End Sub

Public Sub FillDatainList()
    Dim Rs1 As New ADODB.Recordset
    If Rs1.State = 1 Then Rs1.Close
    Rs1.Open "Select SrNo,Date,ChequeNo,Particulars,Debit,Credit,Closing,Remarks from Transactions where AccountNo='" & TxtAccountNo.Text & "'", GConn, 3, 4
    'Rs1.Open "Select * from Transactions where AccountNo='" & TxtAccountNo.Text & "'", GConn, 3, 4

    Call FillListView(ListView1, Rs1, 8, 1, False, False)
End Sub

Public Sub GetBankName()
Dim rsF As New ADODB.Recordset
    
    If rsF.State = 1 Then rsF.Close
    rsF.Open "Select SrNo,BankName From Accounts Where AccountName='" & CboAccountName.Text & "'", GConn, adOpenKeyset, adLockReadOnly
    With CboBankName
    .Clear
    While rsF.EOF = False
        .AddItem IIf(IsNull(rsF(1)), "", rsF(1))
        .ItemData(.NewIndex) = IIf(IsNull(rsF(0)), 0, rsF(0))
        rsF.MoveNext
    Wend
    End With
End Sub

Public Sub GetAccNo()
    Dim Rs2 As New ADODB.Recordset
    If Rs2.State = 1 Then Rs2.Close
    Rs2.Open "Select * from Accounts where AccountName='" & CboAccountName.Text & "' and BankName='" & CboBankName.Text & "'", GConn, 3, 4
    TxtAccountNo.Text = Rs2.Fields(3).Value
End Sub

Public Sub FillUncleared()
    Dim RsU As New ADODB.Recordset
    If RsU.State = 1 Then RsU.Close
    RsU.Open "Select SrNo,Date,ChequeNo,Particulars,Debit,Credit,Remarks from Transactions where AccountNo='" & TxtAccountNo.Text & "' and Clear='" & TxtUnc.Text & "'", GConn, 3, 4
    'Rs1.Open "Select * from Transactions where AccountNo='" & TxtAccountNo.Text & "'", GConn, 3, 4

    Call FillListView(UnClearedList, RsU, 7, 1, False, False)
End Sub

Public Sub BookBalance()
On Error Resume Next
    If TempRs.State = 1 Then TempRs.Close
TempRs.Open "Select Sum(Debit) as ToTBalance from Transactions where AccountNo='" & TxtAccountNo.Text & "'", GConn, adOpenKeyset, adLockOptimistic
                    LblDebit.Caption = TempRs.Fields("ToTBalance")

    If TempRs1.State = 1 Then TempRs1.Close
TempRs1.Open "Select Sum(Credit) as ToTBalance1 from Transactions where AccountNo='" & TxtAccountNo.Text & "'", GConn, adOpenKeyset, adLockOptimistic
                    LblCredit.Caption = TempRs1.Fields("ToTBalance1")

TxtBookBalance.Text = Format$((LblCredit.Caption) - Val(LblDebit.Caption), "#,##0.00")

End Sub

Public Sub BankBalance()
On Error Resume Next
    If TempRs3.State = 1 Then TempRs3.Close
TempRs3.Open "Select Sum(Debit) as ToTBalance from Transactions where AccountNo='" & TxtAccountNo.Text & "' and Clear='" & TxtClr.Text & "'", GConn, adOpenKeyset, adLockOptimistic
                    LblDebitBank.Caption = TempRs3.Fields("ToTBalance")

    If TempRs4.State = 1 Then TempRs4.Close
TempRs4.Open "Select Sum(Credit) as ToTBalance1 from Transactions where AccountNo='" & TxtAccountNo.Text & "' and Clear='" & TxtClr.Text & "'", GConn, adOpenKeyset, adLockOptimistic
                    LblCreditBank.Caption = TempRs4.Fields("ToTBalance1")

TxtBankBalance.Text = Format$((LblCreditBank.Caption) - Val(LblDebitBank.Caption), "#,##0.00")

End Sub

Public Sub LoadData()
On Error Resume Next
Dim RsQ As New ADODB.Recordset
If RsQ.State = 1 Then RsQ.Close
RsQ.Open "select * from Transactions where SrNo=" & Val(TxtSrNo.Text) & "", GConn, 3, 4

TxtDate.Text = RsQ.Fields(2).Value
TxtChequeNo.Text = RsQ.Fields(3).Value
TxtParticulars.Text = RsQ.Fields(4).Value
TxtDebit.Text = RsQ.Fields(5).Value
TxtCredit.Text = RsQ.Fields(6).Value
TxtRemarks.Text = RsQ.Fields(8).Value
End Sub

Private Sub UnClearedList_DblClick()
On Error GoTo ErrMod
    If CboAccountName.Text = "" Or CboBankName.Text = "" Then
    MsgBox "Nothing to Display", vbInformation
    Else
    TxtSrNo.Text = UnClearedList.SelectedItem.Text
    LoadData
    FrmConfirm.Visible = True
    End If
Exit Sub
ErrMod:
MsgBox Err.Description
End Sub
