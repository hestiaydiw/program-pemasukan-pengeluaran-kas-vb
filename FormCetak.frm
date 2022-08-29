VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FormCetak 
   Caption         =   "Form Cetak"
   ClientHeight    =   4590
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   Picture         =   "FormCetak.frx":0000
   ScaleHeight     =   4590
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   960
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command3 
      Caption         =   "KEMBALI"
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CETAK LAPORAN PENGELUARAN"
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   2880
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CETAK LAPORAN PEMASUKAN"
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      Top             =   2280
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2520
      TabIndex        =   3
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "TAHUN"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "BULAN KE-"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   1080
      Width           =   1575
   End
End
Attribute VB_Name = "FormCetak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Combo1 = "" Or Text1 = "" Then
MsgBox "Silahkan Isi Bulan dan Tahun terlebih dahulu!", vbCritical, "Perhatian !"
Else
    CrystalReport1.SelectionFormula = "Month({tbl_pemasukan.TanggalPemasukan})=" & Combo1.Text & " and Year({tbl_pemasukan.TanggalPemasukan})=" & Text1.Text
    CrystalReport1.ReportFileName = App.Path & "\zaimaarpem.rpt"
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.RetrieveDataFiles
    CrystalReport1.Action = 1
End If
Text1.Text = ""
End Sub

Private Sub Command2_Click()
If Combo1 = "" Or Text1 = "" Then
MsgBox "Silahkan Isi Bulan dan Tahun terlebih dahulu!", vbCritical, "Perhatian !"
Else
    CrystalReport1.SelectionFormula = "Month({tbl_pengeluaran.TanggalPengeluaran})=" & Combo1.Text & " and Year({tbl_pengeluaran.TanggalPengeluaran})=" & Text1.Text
    CrystalReport1.ReportFileName = App.Path & "\zaimaarpeng.rpt"
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.RetrieveDataFiles
    CrystalReport1.Action = 1
End If
Text1.Text = ""
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
Combo1.Clear
Combo1.AddItem "1"
Combo1.AddItem "2"
Combo1.AddItem "3"
Combo1.AddItem "4"
Combo1.AddItem "5"
Combo1.AddItem "6"
Combo1.AddItem "7"
Combo1.AddItem "8"
Combo1.AddItem "9"
Combo1.AddItem "10"
Combo1.AddItem "11"
Combo1.AddItem "12"
End Sub

