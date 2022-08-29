VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FormPemasukan 
   Caption         =   "Form Pemasukan"
   ClientHeight    =   9225
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   12765
   LinkTopic       =   "Form1"
   Picture         =   "FormPemasukan.frx":0000
   ScaleHeight     =   9225
   ScaleWidth      =   12765
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   11040
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "E:\PROGRAM BARU\zaimaarpem.rpt"
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DELETE"
      Height          =   495
      Left            =   9000
      TabIndex        =   12
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SAVE"
      Height          =   495
      Left            =   9000
      TabIndex        =   11
      Top             =   1680
      Width           =   1215
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FormPemasukan.frx":1806A2
      Height          =   3135
      Left            =   1560
      OleObjectBlob   =   "FormPemasukan.frx":1806B6
      TabIndex        =   10
      Top             =   5760
      Width           =   9015
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "E:\PROGRAM BARU\dbzaimaarpp.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tbl_pemasukan"
      Top             =   5880
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   4800
      TabIndex        =   9
      Top             =   2280
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
      _Version        =   393216
      Format          =   133562369
      CurrentDate     =   43463
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   4800
      TabIndex        =   8
      Top             =   5160
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Height          =   1605
      Left            =   4800
      TabIndex        =   7
      Top             =   3450
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4800
      TabIndex        =   6
      Top             =   2880
      Width           =   3495
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   11040
      Top             =   1800
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "JUMLAH PEMASUKAN"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   5160
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "KETERANGAN"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "NOMOR TRANSAKSI"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "TANGGAL"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   2280
      Width           =   2175
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
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   3015
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
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   3495
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
         Caption         =   "PENGELUARAN"
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
Attribute VB_Name = "FormPemasukan"
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
Data1.Recordset!TanggalPemasukan = DTPicker1.Value
Data1.Recordset!NomorTransaksi = Text1.Text
Data1.Recordset!KeteranganPemasukan = Text2.Text
Data1.Recordset!JumlahPemasukan = Text3.Text
Data1.Recordset.Update
Data1.Refresh
End If
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text1.SetFocus
End Sub

Private Sub Command4_Click()
FormMnUtama.Show
FormMnUtama.MnTransaksi.Enabled = True
FormPemasukan.Hide
End Sub

Private Sub Command2_Click()
On Error Resume Next
Data1.RecordSource = "Select * From tbl_pemasukan Where NomorTransaksi = '" & Text1.Text & "'"
x = MsgBox("Data Akan Dihapus ?", vbYesNo, "Delete")
If x = vbYes Then
If Not Data1.Recordset.EOF Then Data1.Recordset.Delete
End If
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
CrystalReport1.ReportFileName = App.Path & "\zaimaarpem.rpt"
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

Private Sub MnPengeluaran_Click()
FormPengeluaran.Show
Unload Me
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

