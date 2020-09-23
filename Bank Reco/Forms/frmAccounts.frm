VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmAccounts 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Accounts"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5430
   Icon            =   "frmAccounts.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   3120
      Width           =   5175
      Begin LVbuttons.LaVolpeButton CmdAdd 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Add New"
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
         COLTYPE         =   2
         BCOL            =   14215660
         FCOL            =   0
         FCOLO           =   16711680
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmAccounts.frx":0742
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
      Begin LVbuttons.LaVolpeButton CmdModify 
         Height          =   375
         Left            =   1395
         TabIndex        =   12
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Modify"
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
         COLTYPE         =   2
         BCOL            =   14215660
         FCOL            =   0
         FCOLO           =   16711680
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmAccounts.frx":075E
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
      Begin LVbuttons.LaVolpeButton CmdDelete 
         Height          =   375
         Left            =   2685
         TabIndex        =   13
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Delete"
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
         COLTYPE         =   2
         BCOL            =   14215660
         FCOL            =   0
         FCOLO           =   16711680
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmAccounts.frx":077A
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
      Begin LVbuttons.LaVolpeButton CmdList 
         Height          =   375
         Left            =   3960
         TabIndex        =   14
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&List"
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
         COLTYPE         =   2
         BCOL            =   14215660
         FCOL            =   0
         FCOLO           =   16711680
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmAccounts.frx":0796
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
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Account Detail"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5175
      Begin VB.TextBox TxtClear 
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Text            =   "Yes"
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TxtDate 
         Height          =   285
         Left            =   3720
         TabIndex        =   18
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox TxtParticulars 
         Height          =   375
         Left            =   2640
         TabIndex        =   17
         Text            =   "Opening Balance"
         Top             =   120
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox TxtSrX 
         Height          =   375
         Left            =   2520
         TabIndex        =   16
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox TxtAccountName 
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
         Height          =   375
         Left            =   1650
         TabIndex        =   0
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox TxtOpening 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1650
         TabIndex        =   3
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox TxtAccountNo 
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
         Height          =   375
         Left            =   1650
         TabIndex        =   2
         Top             =   1800
         Width           =   2895
      End
      Begin VB.TextBox TxtBankName 
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
         Height          =   375
         Left            =   1650
         TabIndex        =   1
         Top             =   1320
         Width           =   3375
      End
      Begin VB.TextBox TxtSrNo 
         Alignment       =   2  'Center
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
         Height          =   375
         Left            =   1650
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Account Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   300
         TabIndex        =   15
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Opening Balance"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   2400
         Width           =   1395
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Account No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   525
         TabIndex        =   8
         Top             =   1920
         Width           =   990
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Bank Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   570
         TabIndex        =   7
         Top             =   1440
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Sr. No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   990
         TabIndex        =   6
         Top             =   480
         Width           =   525
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   20
      Top             =   3120
      Width           =   5175
      Begin LVbuttons.LaVolpeButton CmdModSave 
         Height          =   375
         Left            =   1320
         TabIndex        =   21
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Update"
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
         COLTYPE         =   2
         BCOL            =   14215660
         FCOL            =   0
         FCOLO           =   16711680
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmAccounts.frx":07B2
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
      Begin LVbuttons.LaVolpeButton CmdSave 
         Height          =   375
         Left            =   1320
         TabIndex        =   22
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Save"
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
         COLTYPE         =   2
         BCOL            =   14215660
         FCOL            =   0
         FCOLO           =   16711680
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmAccounts.frx":07CE
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
      Begin LVbuttons.LaVolpeButton CmdCancel 
         Height          =   375
         Left            =   2760
         TabIndex        =   23
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Cancel"
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
         COLTYPE         =   2
         BCOL            =   14215660
         FCOL            =   0
         FCOLO           =   16711680
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmAccounts.frx":07EA
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
   End
End
Attribute VB_Name = "frmAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsG As New ADODB.Recordset
Dim RsG1 As New ADODB.Recordset

Private Sub CmdAdd_Click()
    Frame1.Enabled = True
    BlankFields
    GetSrTran
    TxtDate.Text = Date
    TxtAccountName.SetFocus
    If RsG.State = 1 Then RsG.Close
    RsG.Open "Select Max(SrNO) From Accounts", GConn, adOpenKeyset, adLockReadOnly
    If RsG.EOF = False Then
        TxtSrNo.Text = IIf(IsNull(RsG(0)), 1, RsG(0) + 1)
    End If
    Frame2.Visible = False
    CmdModSave.Visible = False
    CmdSave.Visible = True
    CmdCancel.Visible = True
    Frame3.Visible = True
End Sub

Private Sub CmdCancel_Click()
    BlankFields
    Frame3.Visible = False
    Frame2.Visible = True
End Sub

Private Sub CmdDelete_Click()
    If TxtSrNo.Text = "" Then
    MsgBox "Select Record first" & vbCrLf & vbCrLf & "Double click on grid to select Record", vbInformation
    Load frmAccountsList
    Unload Me
    Else
    If MsgBox("Are you sure you want to Delete Account " & " : " & TxtAccountName.Text & " ? " & vbCrLf & vbCrLf & "Record No. " & TxtSrNo.Text, vbCritical + vbYesNoCancel) = vbYes Then
    str9 = "Delete * from Transactions where AccountNo='" & TxtAccountNo.Text & "'"
    GConn.Execute str9
    str3 = "Delete * from Accounts where SrNo=" & Val(TxtSrNo.Text) & ""
    GConn.Execute str3
    MsgBox "Record deleted Successfully", vbInformation
    BlankFields
    End If
    End If
End Sub

Private Sub CmdList_Click()
    Load frmAccountsList
    Unload Me
End Sub

Private Sub CmdModify_Click()
    If TxtSrNo.Text = "" Then
    MsgBox "Select Record first" & vbCrLf & vbCrLf & "Double click on grid to select Record", vbInformation
    Load frmAccountsList
    Unload Me
    Else
    Frame2.Visible = False
    CmdSave.Visible = False
    CmdModSave.Visible = True
    CmdCancel.Visible = True
    Frame3.Visible = True
    Frame1.Enabled = True
    End If
End Sub

Private Sub CmdModSave_Click()
    str4 = "Update Accounts set AccountName='" & TxtAccountName.Text & "',BankName='" & TxtBankName.Text & "',AccountNo='" & TxtAccountNo.Text & "',Opening='" & TxtOpening.Text & "' where SrNo=" & Val(TxtSrNo.Text) & ""
    GConn.Execute str4
    MsgBox "Updated Successfully", vbInformation
    BlankFields
    Frame1.Enabled = False
    Frame3.Visible = False
    Frame2.Visible = True
End Sub

Private Sub CmdSave_Click()
If TxtBankName.Text = "" Or TxtAccountNo.Text = "" Or TxtOpening.Text = "" Then
    MsgBox "Some reuiqred field(s) are empty, Check it and fill it", vbInformation
    Else
    str1 = "Insert Into Accounts(SrNo,AccountName,BankName,AccountNo,Opening) values(" & Val(TxtSrNo.Text) & ",'" & TxtAccountName.Text & "','" & TxtBankName.Text & "','" & TxtAccountNo.Text & "','" & TxtOpening.Text & "')"
    str2 = "Insert Into Transactions(SrNo,[Date],AccountNo,Particulars,Debit,Credit,Closing,Clear) values (" & Val(TxtSrX.Text) & ",'" & TxtDate.Text & "','" & TxtAccountNo.Text & "','" & TxtParticulars.Text & "','0'," & Val(TxtOpening.Text) & "," & Val(TxtOpening.Text) & ",'" & TxtClear.Text & "')"
    GConn.Execute str1
    GConn.Execute str2
    MsgBox "Saves Successfully!"
    BlankFields
    Frame1.Enabled = False
    Frame3.Visible = False
    Frame2.Visible = True
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()
    Call gFormCenter(Me)
    Frame1.Enabled = False
End Sub

Public Sub BlankFields()
    TxtSrNo.Text = ""
    TxtAccountName.Text = ""
    TxtBankName.Text = ""
    TxtAccountNo.Text = ""
    TxtOpening.Text = ""
End Sub

Public Sub GetSrTran()
If RsG1.State = 1 Then RsG1.Close
    RsG1.Open "Select Max(SrNO) From Transactions", GConn, adOpenKeyset, adLockReadOnly
    If RsG1.EOF = False Then
        TxtSrX.Text = IIf(IsNull(RsG1(0)), 1, RsG1(0) + 1)
    End If
End Sub

Private Sub TxtOpening_Change()
    Call OnlyNumber(TxtOpening)
End Sub

Public Sub LoadData()
On Error Resume Next
Dim RsQ As New ADODB.Recordset
If RsQ.State = 1 Then RsQ.Close
RsQ.Open "select * from Accounts where SrNo=" & Val(TxtSrNo.Text) & "", GConn, 3, 4

TxtAccountName.Text = RsQ.Fields(1).Value
TxtBankName.Text = RsQ.Fields(2).Value
TxtAccountNo.Text = RsQ.Fields(3).Value
TxtOpening.Text = RsQ.Fields(4).Value

End Sub
