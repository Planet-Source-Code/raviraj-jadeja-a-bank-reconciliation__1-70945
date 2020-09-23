VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Bank Reconciliation"
   ClientHeight    =   8325
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   12825
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   840
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0FE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1438
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12825
      _ExtentX        =   22622
      _ExtentY        =   1005
      ButtonWidth     =   1905
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Accounts"
            Object.ToolTipText     =   "Accounts"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Transactions"
            Object.ToolTipText     =   "Transaction"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reconciliation"
            Object.ToolTipText     =   "Reconciliatoin"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
      MousePointer    =   99
      MouseIcon       =   "MDIForm1.frx":188A
   End
   Begin VB.Menu MnuFile 
      Caption         =   "&File"
      Begin VB.Menu MnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu MnuMaster 
      Caption         =   "&General"
      Begin VB.Menu MnuAccounts 
         Caption         =   "&Accounts"
      End
      Begin VB.Menu MnuTransactions 
         Caption         =   "&Transactions"
      End
      Begin VB.Menu MnuReconciliation 
         Caption         =   "&Reconciliation"
      End
   End
   Begin VB.Menu MnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
Dim hMenu As Long

hMenu = GetSystemMenu(Me.hwnd, 0&)
If hMenu Then
Call DeleteMenu(hMenu, SC_CLOSE, MF_BYCOMMAND)
DrawMenuBar (Me.hwnd)
End If


End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
If MsgBox("Are you sure you want to Exit?", vbQuestion + vbYesNo) = vbYes Then
End
Else
Cancel = 1
End If
End Sub

Private Sub MnuAbout_Click()
    Load frmAbout
End Sub

Private Sub MnuAccounts_Click()
    Load frmAccounts
End Sub

Private Sub MnuExit_Click()
    If MsgBox("Are you sure you want to Exit?", vbQuestion + vbYesNo) = vbYes Then
    End
    Else
    Cancel = 1
    End If
End Sub

Private Sub MnuReconciliation_Click()
    Load frmRecon
End Sub

Private Sub MnuTransactions_Click()
    Load frmTransaction
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Caption

    Case "Accounts"
    Load frmAccounts
    
    Case "Transactions"
    Load frmTransaction
    
    Case "Reconciliation"
    Load frmRecon
    
    Case "Exit"
    MnuExit_Click
    
End Select
End Sub
