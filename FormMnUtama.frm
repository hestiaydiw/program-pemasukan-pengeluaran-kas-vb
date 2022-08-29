VERSION 5.00
Begin VB.Form FormMnUtama 
   Caption         =   "Form MnUtama"
   ClientHeight    =   10635
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   Picture         =   "FormMnUtama.frx":0000
   ScaleHeight     =   10635
   ScaleWidth      =   20250
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5040
      Top             =   1440
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   6135
   End
   Begin VB.Menu MnMenu 
      Caption         =   "MENU"
      Begin VB.Menu MnAdmin 
         Caption         =   "ADMIN"
      End
      Begin VB.Menu MnExit 
         Caption         =   "EXIT"
      End
   End
   Begin VB.Menu MnTransaksi 
      Caption         =   "TRANSAKSI"
      Begin VB.Menu MnPemasukan 
         Caption         =   "PEMASUKAN"
      End
      Begin VB.Menu MnPengeluaran 
         Caption         =   "PENGELUARAN"
      End
   End
End
Attribute VB_Name = "FormMnUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
FormMnUtama.MnTransaksi.Enabled = False
End Sub

Private Sub MnExit_Click()
Pesan = MsgBox("Anda yakin ingin keluar?", vbYesNo)
If Pesan = vbYes Then End
End Sub
Private Sub MnAdmin_Click()
FormLogin.Show
End Sub

Private Sub MnPemasukan_Click()
FormPemasukan.Show
End Sub
Private Sub MnPengeluaran_Click()
FormPengeluaran.Show
End Sub
Private Sub Timer1_Timer()
    Label2.Caption = Time
    Label1.Caption = Date
    Label1.Caption = namahari(Date)
End Sub
Function namahari(dtanggal As Date) As String
    Dim SHari As String
    aHari = Array("Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu")
    namahari = aHari(Abs(Weekday(dtanggal) - 1)) & "," & "" & Format(dtanggal, "d mmmm yyyy")
End Function
