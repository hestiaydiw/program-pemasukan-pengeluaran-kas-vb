VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FormPengeluaran 
   Caption         =   "Form Pengeluaran"
   ClientHeight    =   9300
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   12645
   LinkTopic       =   "Form1"
   Picture         =   "FormPengeluaran.frx":0000
   ScaleHeight     =   9300
   ScaleWidth      =   12645
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   10800
      Top             =   840
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      Top             =   2520
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Height          =   1605
      Left            =   5280
      TabIndex        =   5
      Top             =   3090
      Width           =   3495
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   4800
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SAVE"
      Height          =   495
      Left            =   9480
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DELETE"
      Height          =   495
      Left            =   9480
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   11160
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "E:\PROGRAM BARU\zaimaarpeng.rpt"
      PrintFileLinesPerPage=   60
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FormPengeluaran.frx":1811F2
      Height          =   3135
      Left            =   2040
      OleObjectBlob   =   "FormPengeluaran.frx":181206
      TabIndex        =   2
      Top             =   5280
      Width           =   9015
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   1920
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
      _Version        =   393216
      Format          =   133562369
      CurrentDate     =   43463
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "E:\PROGRAM BARU\dbzaimaarpp.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tbl_pengeluaran"
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   12
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   11
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "TANGGAL"
      Height          =   375
      Left            =   2880
      TabIndex        =   10
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "NOMOR TRANSAKSI"
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "KETERANGAN"
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "JUMLAH PENGELUARAN"
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Menu MnMenu 
      Caption         =   "MENU"
      Begin VB.Menu MnHome 
         Caption         =   "HOME"
      End
   End
   Begin VB.Menu MnTransaksi 
      Caption         =   "TRANSAKSI"
      Begin VB.Menu MnPemasukan 
         Caption         =   "PEMASUKAN"
      End
      Begin VB.Menu MnPengeluaran 
         Caption         =   "PEGELUARAN"
      End
   End
   Begin VB.Menu MnCetak 
      Caption         =   "CETAK"
      Begin VB.Menu MnCSemua 
         Caption         =   "CETAK SEMUA"
      End
      Begin VB.Menu MnCBerdasarkan 
         Caption         =   "CETAK BERDASARKAN..."
      End
   End
End
Attribute VB_Name = "FormPengeluaran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next

If Trim(Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "") Then
    MsgBox "Data Belum Dilengkapi", vbInformation, "Informasi"
Else
Data1.Recordset.AddNew
Data1.Recordset!TanggalPengeluaran = DTPicker1.Value
Data1.Recordset!NomorTransaksi = Text1.Text
Data1.Recordset!KeteranganPengeluaran = Text2.Text
Data1.Recordset!JumlahPengeluaran = Text3.Text
Data1.Recordset.Update
Data1.Refresh
End If
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text1.SetFocus
End Sub

Private Sub Command2_Click()
On Error Resume Next
Data1.RecordSource = "Select * From tbl_Pengeluaran Where NomorTransaksi = '" & Text1.Text & "'"
x = MsgBox("Data Akan Dihapus ?", vbYesNo, "Delete")
If x = vbYes Then
If Not Data1.Recordset.EOF Then Data1.Recordset.Delete
End If
End Sub

Private Sub Command4_Click()
FormMenuUtama.Show
FormMenuUtama.MnTransaksi.Enabled = True
FormPengeluaran.Hide
End Sub


Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
DTPicker1.CustomFormat = "dd MMMM yyyy"
End Sub

Private Sub Form_Load()
    DTPicker1.Format = dtpCustom
    DTPicker1.CustomFormat = "dd MMMM yyyy"
    DTPicker1.Value = Date
End Sub

Private Sub MnCBerdasarkan_Click()
FormCetak.Show
End Sub

Private Sub MnCSemua_Click()
CrystalReport1.ReportFileName = App.Path & "\zaimaarpeng.rpt"
CrystalReport1.WindowState = crptMaximized
CrystalReport1.RetrieveDataFiles
CrystalReport1.Action = 0
End Sub

Private Sub MnExit_Click()
Pesan = MsgBox("Anda yakin ingin keluar?", vbYesNo)
If Pesan = vbYes Then End
End Sub

Private Sub MnHome_Click()
FormMnUtama.Show
FormMnUtama.MnTransaksi.Enabled = True
Unload Me
End Sub

Private Sub MnPemasukan_Click()
FormPemasukan.Show
Unload Me
End Sub

Private Sub MnPengeluaran_Click()
FormPengeluaran.Show
End Sub

Private Sub Text1_Change()
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then KeyAscii = 0
Text1.MaxLength = 10
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
