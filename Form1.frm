VERSION 5.00
Begin VB.Form FormLogin 
   Caption         =   "Form Login"
   ClientHeight    =   5580
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   5580
   ScaleWidth      =   7035
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      Height          =   495
      Left            =   3600
      TabIndex        =   5
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LOGIN"
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3360
      PasswordChar    =   "X"
      TabIndex        =   3
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3240
      Width           =   1140
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "USERNAME"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   1680
      Width           =   1695
   End
End
Attribute VB_Name = "FormLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
UserName = Text1.Text
Password = Text2.Text
Data1.RecordSource = "Select * From tbl_login Where Username = '" & Text1.Text & "'"
MsgBox "Selamat Datang"
FormMnUtama.Show
FormLogin.Hide
If Trim(Text1 = "" Or Text2 = "") Then
MsgBox "Username atau Password Yang Anda Masukkan Salah", vbInformation, "Informasi"
End If
FormMnUtama.MnTransaksi.Enabled = True
End Sub

Private Sub Command2_Click()
Pesan = MsgBox("Anda Yakin Ingin Keluar ?", vbQuestion + vbYesNo, "Question")
If Pesan = vbYes Then
End
Else
Form1.SetFocus
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
On Error Resume Next
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 And Not Text1.Text = "" Then
Data1.RecordSource = "Select * From tbl_login Where Username = '" & Text1.Text & "'"
MsgBox "Selamat Datang"
FormMenuUtama.Show
FormLogin.Hide
If Trim(Text1 = "" Or Text2 = "") Then
MsgBox "Username atau Password Yang Anda Masukkan Salah", vbInformation, "Informasi"
End If
End If
FormMenuUtama.MnTransaksi.Enabled = True
End Sub
