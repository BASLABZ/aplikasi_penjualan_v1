VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form FormMenuUtama 
   Caption         =   "Menu Utama"
   ClientHeight    =   8280
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   12945
   LinkTopic       =   "Form1"
   Picture         =   "FormMenuUtama.frx":0000
   ScaleHeight     =   8280
   ScaleWidth      =   12945
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   840
      Top             =   5880
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7905
      Width           =   12945
      _ExtentX        =   22834
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Visible         =   0   'False
            Object.Width           =   4410
            MinWidth        =   4410
            Text            =   "10"
            TextSave        =   "10"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu MnFile 
      Caption         =   "File"
      Begin VB.Menu MnSeting 
         Caption         =   "Setting Stok"
      End
      Begin VB.Menu MnLogout 
         Caption         =   "Logout"
      End
      Begin VB.Menu MnExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu MnOlahData 
      Caption         =   "Olah Data"
      Begin VB.Menu MnJenisBarang 
         Caption         =   "Jenis Barang"
      End
      Begin VB.Menu MnBarang 
         Caption         =   "Barang"
      End
      Begin VB.Menu MnSupplier 
         Caption         =   "Supplier"
      End
      Begin VB.Menu MnUser 
         Caption         =   "User"
      End
   End
   Begin VB.Menu MnTransaksi 
      Caption         =   "Transaksi"
      Begin VB.Menu MnPenjualan 
         Caption         =   "Penjualan"
      End
      Begin VB.Menu MnPembelian 
         Caption         =   "Pemebelian"
      End
   End
   Begin VB.Menu mnGrafik 
      Caption         =   "Grafik"
      Begin VB.Menu mnGrafikBrangterjual 
         Caption         =   "Grafik Barang Terjual"
      End
      Begin VB.Menu MnGrafikPenghasilan 
         Caption         =   "Grafik Pendapatan"
      End
   End
   Begin VB.Menu MnLaporan 
      Caption         =   "Laporan"
      Begin VB.Menu MnStok 
         Caption         =   "Lap Stok"
         Begin VB.Menu MnLapStokmenipis 
            Caption         =   "Laporan Stok Menipis"
         End
         Begin VB.Menu MnLapStok 
            Caption         =   "Laporan Stok Barang"
         End
      End
      Begin VB.Menu MnLapPenjualan 
         Caption         =   "Laporan Penjualan"
      End
      Begin VB.Menu MnLapBeli 
         Caption         =   "Laporan Pembelian"
      End
      Begin VB.Menu LapbarangJual 
         Caption         =   "Laporan Barang Terjual"
      End
   End
End
Attribute VB_Name = "FormMenuUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Koneksi
End Sub

Private Sub LapbarangJual_Click()
FormFilterBarangTerjual.Show
End Sub

Private Sub MnBarang_Click()
FormBarang.Show
End Sub

Private Sub MnExit_Click()
Unload Me
End Sub

Private Sub mnGrafikBrangterjual_Click()
Grafikterlaris.Show
End Sub

Private Sub MnGrafikPenghasilan_Click()
GrafikPendapatan.Show
End Sub

Private Sub MnJenisBarang_Click()
FormJenisBarang.Show
End Sub

Private Sub MnLapBeli_Click()
FormFilterpembelian.Show
End Sub

Private Sub MnLapPenjualan_Click()
FormFilterPenjualan.Show
End Sub

Private Sub MnLapStok_Click()
LapStokAll
End Sub

Private Sub MnLapStokmenipis_Click()
LapStokMenipis
End Sub

Private Sub MnLogout_Click()
Unload Me
FormLogin.Show
End Sub

Private Sub MnPembelian_Click()
FormTransaksiPembelian.Show
End Sub

Private Sub MnPenjualan_Click()
FormTransaksiPenjualan.Show
End Sub

Private Sub MnReturPembelian_Click()
FormReturPembelian.Show
End Sub

Private Sub MnSeting_Click()
Stokmenipis.Show
End Sub

Private Sub MnSupplier_Click()
FormSupplier.Show
End Sub

Private Sub MnUser_Click()
FormUser.Show
End Sub

Private Sub Timer1_Timer()
StatusBar1.Panels(4) = "Tanggal" & Format(Date, " dd mmmm yyyy ")
StatusBar1.Panels(3) = "Jam " & Format(Time, " hh:mm:ss Am/Pm ")
End Sub


Sub Kasir()
MnOlahData.Visible = False
MnPembelian.Visible = False
MnLaporan.Visible = False
End Sub

Sub Pemilik()
MnOlahData.Visible = False
MnTransaksi.Visible = False
End Sub

Sub Admin()
MnLaporan.Visible = False
End Sub

Sub LapStokAll()
strsql = "select * from barang"
With LaporanStokBarang
        .DataControl1.ConnectionString = strConn
        .DataControl1.Source = strsql
        '.Field6.Text = RsCek(0)
        '.Label14.Caption = "Laporan Penjualan Bulan : " + Mid(CmbBulan.Text, 6, 20) + " Tahun : " + CmbTahun.Text
        .Show
    End With
End Sub

Sub LapStokMenipis()
strsql = "select * from barang where stok < '" & FormMenuUtama.StatusBar1.Panels(5) & "'"
With LaporanStokBarang
        .DataControl1.ConnectionString = strConn
        .DataControl1.Source = strsql
        '.Field6.Text = RsCek(0)
        .Label14.Caption = "Laporan Stok Menipis Barang"
        .Show
    End With
End Sub
