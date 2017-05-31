VERSION 5.00
Begin VB.Form FormFilterBarangTerjual 
   Caption         =   "Filter Barang Terjual"
   ClientHeight    =   6645
   ClientLeft      =   8355
   ClientTop       =   1470
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FormBarangTerjual.frx":0000
   ScaleHeight     =   6645
   ScaleWidth      =   5460
   Begin VB.ComboBox CmbKategori 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   840
      TabIndex        =   11
      Top             =   840
      Width           =   4215
   End
   Begin VB.OptionButton OptHariIni 
      BackColor       =   &H00C0C000&
      Caption         =   "Hari Ini"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   1920
      Width           =   2055
   End
   Begin VB.OptionButton OptBulan 
      BackColor       =   &H00C0C000&
      Caption         =   "Bulanan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   2520
      Width           =   2175
   End
   Begin VB.OptionButton OptTahun 
      BackColor       =   &H00C0C000&
      Caption         =   "Tahunan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   4320
      Width           =   2415
   End
   Begin VB.ComboBox CmbBulan 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2880
      TabIndex        =   3
      Top             =   3240
      Width           =   2175
   End
   Begin VB.ComboBox CmbTahun 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2880
      TabIndex        =   2
      Top             =   3720
      Width           =   2175
   End
   Begin VB.ComboBox cmbTahun1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2880
      TabIndex        =   1
      Top             =   5040
      Width           =   2175
   End
   Begin VB.CommandButton cmdLaporan 
      Caption         =   "Laporan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      Picture         =   "FormBarangTerjual.frx":548F5
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Kategori"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   10
      Top             =   480
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00C0C000&
      FillStyle       =   0  'Solid
      Height          =   1335
      Left            =   360
      Top             =   240
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bulan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   9
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Tahun"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   8
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Tahun"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C0C000&
      FillStyle       =   0  'Solid
      Height          =   4095
      Left            =   360
      Top             =   1680
      Width           =   4935
   End
End
Attribute VB_Name = "FormFilterBarangTerjual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Koneksi
TampilCombo
End Sub

Sub TampilCombo()
Dim rsjenis As Recordset
Set rsjenis = New Recordset
Dim sqlJenis As String

sqlJenis = "SELECT * FROM JenisBarang"
rsjenis.Open sqlJenis, Conn, adOpenStatic, adLockReadOnly

Do While Not rsjenis.EOF
    CmbKategori.AddItem rsjenis!idJenisBarang + "-" + rsjenis!NamaJenisBarang
    rsjenis.MoveNext
Loop

CmbBulan.AddItem "01 - Januari"
CmbBulan.AddItem "02 - Februari"
CmbBulan.AddItem "03 - Maret"
CmbBulan.AddItem "04 - April"
CmbBulan.AddItem "05 - Mei"
CmbBulan.AddItem "06 - Juni"
CmbBulan.AddItem "07 - Juli"
CmbBulan.AddItem "08 - Agustus"
CmbBulan.AddItem "09 - September"
CmbBulan.AddItem "10 - Oktober"
CmbBulan.AddItem "11 - November"
CmbBulan.AddItem "12 - Desember"

CmbTahun.AddItem "2014"
CmbTahun.AddItem "2015"
CmbTahun.AddItem "2016"
CmbTahun.AddItem "2017"

cmbTahun1.AddItem "2014"
cmbTahun1.AddItem "2015"
cmbTahun1.AddItem "2016"
cmbTahun1.AddItem "2017"
End Sub


Private Sub cmdLaporan_Click()
If OptHariIni.Value = True Then
    strsql = "select *"
    strsql = strsql & " From penjualan p join detailPenjualan dp on dp.idPenjualan=p.idPenjualan"
    strsql = strsql & " Join barang b on b.idbarang=dp.idBarang"
    strsql = strsql & " Join jenisBarang k on b.idJenisBarang=k.idJenisBarang"
    strsql = strsql & " where (tglJual)='" & Format(Now, "YYYY-MM-DD") & "' and "
    strsql = strsql & " b.idJenisBarang='" & Mid(CmbKategori.Text, 1, 4) & "'"

    Set RsLihat = Conn.Execute(strsql)
    If Not RsLihat.EOF Then
    LapHariIni
    Else
        MsgBox "Data Tidak Tersedia", vbInformation, "Informasi"
    End If
    
ElseIf OptBulan.Value = True Then

    strsql = "select *"
    strsql = strsql & " From penjualan p join detailPenjualan dp on dp.idPenjualan=p.idPenjualan"
    strsql = strsql & " Join barang b on b.idbarang=dp.idBarang"
    strsql = strsql & " Join jenisBarang k on b.idJenisBarang=k.idJenisBarang"
    strsql = strsql & " where month(tglJual)='" & Left(CmbBulan.Text, 2) & "'"
    strsql = strsql & " and year(tglJual)='" & CmbTahun.Text & "'"
    strsql = strsql & " and b.idJenisBarang='" & Mid(CmbKategori.Text, 1, 4) & "'"
    
    Set RsLihat = Conn.Execute(strsql)
    If Not RsLihat.EOF Then
    LapBulan
    Else
        MsgBox "Data Tidak Tersedia", vbInformation, "Informasi"
    End If
ElseIf OptTahun.Value = True Then
    strsql = "select *"
    strsql = strsql & " From penjualan p join detailPenjualan dp on dp.idPenjualan=p.idPenjualan"
    strsql = strsql & " Join barang b on b.idbarang=dp.idBarang"
    strsql = strsql & " Join jenisBarang k on b.idJenisBarang=k.idJenisBarang"
    strsql = strsql & " where year(tglJual)='" & cmbTahun1.Text & "'"
    strsql = strsql & " and b.idJenisBarang='" & Mid(CmbKategori.Text, 1, 4) & "'"
    
    Set RsLihat = Conn.Execute(strsql)
    If Not RsLihat.EOF Then
    LapTahun
    Else
        MsgBox "Data Tidak Tersedia", vbInformation, "Informasi"
    End If
End If
End Sub


Sub LapBulan()
    strsql = "select namaBarang,sum(jumlah) as tot"
    strsql = strsql & " From penjualan p join detailPenjualan dp on dp.idPenjualan=p.idPenjualan"
    strsql = strsql & " Join barang b on b.idbarang=dp.idBarang"
    strsql = strsql & " Join JenisBarang k on b.idJenisBarang=k.idJenisBarang"
    strsql = strsql & " where month(tglJual)='" & Left(CmbBulan.Text, 2) & "'"
    strsql = strsql & " and year(tglJual)='" & CmbTahun.Text & "'"
    strsql = strsql & " and b.idJenisBarang='" & Mid(CmbKategori.Text, 1, 4) & "' "
    strsql = strsql & " Group by namaBarang order by sum(jumlah) desc"

With LaporanBarangTerjual
        .DataControl1.ConnectionString = strConn
        .DataControl1.Source = strsql
        .LabelPeriode.Caption = "Laporan Penjualan Bulan : " + Mid(CmbBulan.Text, 6, 20) + " Tahun : " + CmbTahun.Text
        .LabelKategori.Caption = Mid(CmbKategori, 6, 30)
        .Show
    End With
End Sub

Sub LapTahun()
    strsql = "select namaBarang,sum(jumlah) as tot"
    strsql = strsql & " From penjualan p join detailPenjualan dp on dp.idPenjualan=p.idPenjualan"
    strsql = strsql & " Join barang b on b.idbarang=dp.idBarang"
    strsql = strsql & " Join JenisBarang k on b.idJenisBarang=k.idJenisBarang"
    strsql = strsql & " where year(tglJual)='" & cmbTahun1.Text & "'"
    strsql = strsql & " and b.idJenisBarang='" & Mid(CmbKategori.Text, 1, 4) & "'"
    strsql = strsql & " Group by namaBarang order by sum(jumlah) desc"

With LaporanBarangTerjual
        .DataControl1.ConnectionString = strConn
        .DataControl1.Source = strsql
        .LabelPeriode.Caption = "Laporan Penjualan Tahun : " + cmbTahun1.Text
        .LabelKategori.Caption = Mid(CmbKategori, 6, 30)
        .Show
    End With
End Sub

Sub LapHariIni()

    strsql = "select namaBarang,sum(jumlah) as tot"
    strsql = strsql & " From penjualan p join detailPenjualan dp on dp.idPenjualan=p.idPenjualan"
    strsql = strsql & " Join barang b on b.idbarang=dp.idBarang"
    strsql = strsql & " Join JenisBarang k on b.idJenisBarang=k.idJenisBarang"
    strsql = strsql & " where (tglJual)='" & Format(Now, "YYYY-MM-DD") & "' and "
    strsql = strsql & " b.idJenisBarang='" & Mid(CmbKategori.Text, 1, 4) & "'"
    strsql = strsql & " Group by namaBarang order by sum(jumlah) desc"
    
With LaporanBarangTerjual
        .DataControl1.ConnectionString = strConn
        .DataControl1.Source = strsql
        .LabelPeriode.Caption = "Laporan Penjualan  : " + Format(Now, "DD MM YYYY")
        .LabelKategori.Caption = Mid(CmbKategori, 6, 30)
        .Show
    End With
End Sub

