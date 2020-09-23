VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTransaction 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Transactions"
   ClientHeight    =   9810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13950
   Icon            =   "frmTransaction.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9810
   ScaleWidth      =   13950
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   360
      Top             =   9000
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
            Picture         =   "frmTransaction.frx":0442
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3495
      Left            =   120
      TabIndex        =   29
      Top             =   5760
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   6165
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
   Begin LVbuttons.LaVolpeButton CmdNew 
      Height          =   495
      Left            =   120
      TabIndex        =   25
      Top             =   5040
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&New"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
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
      MICON           =   "frmTransaction.frx":0894
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
   Begin VB.TextBox TxtClosing 
      Height          =   375
      Left            =   5760
      TabIndex        =   22
      Top             =   4920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   615
      Left            =   -240
      TabIndex        =   21
      Top             =   5280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox TxtSrNo 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   240
      TabIndex        =   20
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   120
      TabIndex        =   9
      Top             =   3360
      Width           =   14295
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   20578305
         CurrentDate     =   39564
      End
      Begin VB.TextBox TxtCredit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9360
         TabIndex        =   3
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox TxtRemarks 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   10920
         TabIndex        =   4
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox TxtDebit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7800
         TabIndex        =   2
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox TxtParticulars 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   720
         Width           =   4335
      End
      Begin VB.TextBox TxtChequeNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   0
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         Caption         =   "Credit Amt."
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
         Left            =   9360
         TabIndex        =   19
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         Caption         =   "Remarks"
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
         Left            =   10920
         TabIndex        =   14
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         Caption         =   "Debit Amt."
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
         Left            =   7800
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         Caption         =   "Particulars"
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
         Left            =   3360
         TabIndex        =   12
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         Caption         =   "Cheque No."
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
         Left            =   1560
         TabIndex        =   11
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         Caption         =   "Date"
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
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   14295
      Begin VB.ComboBox CboBankName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   3600
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   960
         Width           =   9735
      End
      Begin VB.ComboBox CboAccountName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   3600
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   240
         Width           =   9735
      End
      Begin VB.TextBox TxtBalance 
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
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9840
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1680
         Width           =   3495
      End
      Begin VB.TextBox TxtAccountNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1680
         Width           =   3975
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000080&
         Caption         =   "Balance :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   7680
         TabIndex        =   17
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000080&
         Caption         =   "Account No :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   15
         Top             =   1680
         Width           =   3375
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000080&
         Caption         =   "Account Name :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000080&
         Caption         =   "Bank :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   3375
      End
   End
   Begin LVbuttons.LaVolpeButton CmdModify 
      Height          =   495
      Left            =   1680
      TabIndex        =   30
      Top             =   5040
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Modify"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
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
      MICON           =   "frmTransaction.frx":08B0
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
   Begin LVbuttons.LaVolpeButton CmdModSave 
      Height          =   495
      Left            =   1680
      TabIndex        =   31
      Top             =   5040
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Save"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
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
      MICON           =   "frmTransaction.frx":08CC
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
   Begin VB.Label LblCredit 
      Height          =   375
      Left            =   4080
      TabIndex        =   24
      Top             =   4920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label LblDebit 
      Height          =   375
      Left            =   2520
      TabIndex        =   23
      Top             =   4920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "Transaction Entry"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   14295
   End
End
Attribute VB_Name = "frmTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rs As New ADODB.Recordset
Dim Rs1 As New ADODB.Recordset
Dim Rs2 As New ADODB.Recordset
Dim RsG As New ADODB.Recordset
Dim TempRs As New ADODB.Recordset
Dim TempRs1 As New ADODB.Recordset

Private Sub CboAccountName_Click()
    TxtAccountNo.Text = ""
    TxtBalance.Text = ""
    GetBankName
End Sub

Private Sub CboBankName_Click()
    GetAccNo
    GR
    FillDatainList
End Sub

'''''Private Sub CmdModify_Click()
'''''    If TxtSrNo.Text = "" Then
'''''    MsgBox "Select Transaction first", vbInformation
'''''    ListView1.SetFocus
'''''    Else
'''''    Frame2.Enabled = True
'''''    CmdModify.Visible = False
'''''    CmdModSave.Visible = True
'''''    End If
'''''End Sub
'''''
'''''Private Sub CmdModSave_Click()
'''''    str4 = "Update Transactions set [Date]='" & DTPicker1.Value & "',ChequeNo='" & TxtChequeNo.Text & "',Particulars='" & TxtParticulars.Text & "',Debit='" & TxtDebit.Text & "',Credit='" & TxtCredit.Text & "',Closing='" & TxtClosing.Text & "',Remarks='" & TxtRemarks.Text & "' where SrNo=" & Val(TxtSrNo.Text) & ""
'''''    GConn.Execute str4
'''''    MsgBox "Updated Successfully", vbInformation
'''''    BlankFields
'''''    GR
'''''    FillDatainList
'''''    ListView1.Refresh
'''''    Frame2.Enabled = False
'''''    CmdModSave.Visible = False
'''''    CmdModify.Visible = True
'''''End Sub

Private Sub CmdNew_Click()
'    CmdAccountOpen.SetFocus
    If CboAccountName.Text = "" Then
    MsgBox "Select Account First", vbInformation
    CboAccountName.SetFocus
    Else
    Frame2.Enabled = True
    BlankFields
    TxtChequeNo.SetFocus
    If RsG.State = 1 Then RsG.Close
    RsG.Open "Select Max(SrNO) From Transactions", GConn, adOpenKeyset, adLockReadOnly
    If RsG.EOF = False Then
        TxtSrNo.Text = IIf(IsNull(RsG(0)), 1, RsG(0) + 1)
    End If
    End If
End Sub

Private Sub Command2_Click()
    If TxtParticulars.Text = "" Then
    MsgBox "Some required field(s) is/are empty" & vbCrLf & vbCrLf & "Check it and fill it", vbInformation
    Else
    str1 = "Insert Into Transactions(SrNo,AccountNo,[Date],ChequeNo,Particulars,Debit,Credit,Closing,Remarks,Clear) values(" & Val(TxtSrNo.Text) & ",'" & TxtAccountNo.Text & "','" & DTPicker1.Value & "','" & TxtChequeNo.Text & "','" & TxtParticulars.Text & "'," & Val(TxtDebit.Text) & "," & Val(TxtCredit.Text) & "," & Val(TxtClosing.Text) & ",'" & TxtRemarks.Text & "','No')"
    GConn.Execute str1
    MsgBox "Record Saved Successfully", vbInformation
    FillDatainList
    GR
    BlankFields
    End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()
    Label1.Left = (Screen.Width - Label1.Width) / 2
    Frame1.Left = (Screen.Width - Frame1.Width) / 2
    Frame1.Top = (Screen.Height - Frame1.Height) / 12
    Frame2.Left = (Screen.Width - Label1.Width) / 2
    ListView1.Left = (Screen.Width - ListView1.Width) / 2
    CmdNew.Left = Frame2.Left
    DTPicker1.Value = Date
    Frame2.Enabled = False
    
    Dim strd As String
    strd = "Select SrNo,AccountName From Accounts"
    Call FillCombo(CboAccountName, strd)
    
    strd = "Select SrNo,BankName From Accounts"
    Call FillCombo(CboBankName, strd)
        
End Sub

Public Sub GetAccNo()
    If Rs2.State = 1 Then Rs2.Close
    Rs2.Open "Select * from Accounts where AccountName='" & CboAccountName.Text & "' and BankName='" & CboBankName.Text & "'", GConn, 3, 4
    TxtAccountNo.Text = Rs2.Fields(3).Value
End Sub

Public Sub GR()
On Error Resume Next
    If TempRs.State = 1 Then TempRs.Close
TempRs.Open "Select Sum(Debit) as ToTBalance from Transactions where AccountNo='" & TxtAccountNo.Text & "'", GConn, adOpenKeyset, adLockOptimistic
                    LblDebit.Caption = TempRs.Fields("ToTBalance")

    If TempRs1.State = 1 Then TempRs1.Close
TempRs1.Open "Select Sum(Credit) as ToTBalance1 from Transactions where AccountNo='" & TxtAccountNo.Text & "'", GConn, adOpenKeyset, adLockOptimistic
                    LblCredit.Caption = TempRs1.Fields("ToTBalance1")

TxtBalance.Text = Format$((LblCredit.Caption) - Val(LblDebit.Caption), "#,##0.00")
TxtClosing.Text = Format$((LblCredit.Caption) - Val(LblDebit.Caption), "#,##0.00")
End Sub

Private Sub ListView1_DblClick()
On Error GoTo ErrMod
    If CboAccountName.Text = "" Or CboBankName.Text = "" Then
    MsgBox "Nothing to Display", vbInformation
    Else
    TxtSrNo.Text = ListView1.SelectedItem.Text
    LoadData
    GR
    End If
Exit Sub
ErrMod:
MsgBox Err.Description
End Sub

Private Sub TxtChequeNo_Change()
'    CmdNew_Click
End Sub

Private Sub TxtCredit_Change()
    Call OnlyNumber(TxtCredit)
    TxtClosing.Text = Val(LblCredit.Caption) - Val(LblDebit.Caption) + Val(TxtCredit.Text)
    TxtBalance.Text = Val(LblCredit.Caption) - Val(LblDebit.Caption) + Val(TxtCredit.Text)
End Sub

Private Sub TxtDebit_Change()
    Call OnlyNumber(TxtDebit)
    TxtClosing.Text = Val(LblCredit.Caption) - Val(LblDebit.Caption) - Val(TxtDebit.Text)
    TxtBalance.Text = Val(LblCredit.Caption) - Val(LblDebit.Caption) - Val(TxtDebit.Text)
End Sub

Private Sub TxtRemarks_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
    If MsgBox("Are you sure you want to Save this record?", vbQuestion + vbYesNo) = vbYes Then
    Command2_Click
    Else
    BlankFields
    End If
    End If
End Sub

Public Sub BlankFields()
    DTPicker1.Value = Date
    TxtChequeNo.Text = ""
    TxtParticulars.Text = ""
    TxtDebit.Text = ""
    TxtCredit.Text = ""
    TxtRemarks.Text = ""
End Sub

Public Sub FillDatainList()
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

Public Sub LoadData()
On Error Resume Next
Dim RsQ As New ADODB.Recordset
If RsQ.State = 1 Then RsQ.Close
RsQ.Open "select * from Transactions where SrNo=" & Val(TxtSrNo.Text) & "", GConn, 3, 4

DTPicker1.Value = RsQ.Fields(2).Value
TxtChequeNo.Text = RsQ.Fields(3).Value
TxtParticulars.Text = RsQ.Fields(4).Value
TxtDebit.Text = RsQ.Fields(5).Value
TxtCredit.Text = RsQ.Fields(6).Value
TxtClosing.Text = RsQ.Fields(7).Value
TxtRemarks.Text = RsQ.Fields(8).Value
End Sub
