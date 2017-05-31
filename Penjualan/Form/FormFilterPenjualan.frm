VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormFilterPenjualan 
   Caption         =   "Pemilihan Laporan Penjualan"
   ClientHeight    =   5280
   ClientLeft      =   7845
   ClientTop       =   1980
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FormFilterPenjualan.frx":0000
   ScaleHeight     =   5280
   ScaleWidth      =   5370
   Begin VB.OptionButton OptPeriode 
      BackColor       =   &H00C0C000&
      Caption         =   "Penjualan Periode"
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
      Left            =   480
      TabIndex        =   9
      Top             =   360
      Width           =   2055
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
      Left            =   3960
      Picture         =   "FormFilterPenjualan.frx":548F5
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4440
      Width           =   1215
   End
   Begin VB.ComboBox cmbTahun1 
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
      Left            =   2760
      TabIndex        =   6
      Top             =   3600
      Width           =   2175
   End
   Begin VB.ComboBox CmbTahun 
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
      Left            =   2760
      TabIndex        =   3
      Top             =   2280
      Width           =   2175
   End
   Begin VB.ComboBox CmbBulan 
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
      Left            =   2760
      TabIndex        =   2
      Top             =   1800
      Width           =   2175
   End
   Begin VB.OptionButton OptTahun 
      BackColor       =   &H00C0C000&
      Caption         =   "Total Penjualan Tahunan"
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
      Left            =   600
      TabIndex        =   1
      Top             =   2880
      Width           =   2415
   End
   Begin VB.OptionButton OptBulan 
      BackColor       =   &H00C0C000&
      Caption         =   "Total Penjualan Bulananan"
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
      Left            =   600
      TabIndex        =   0
      Top             =   1200
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   840
      TabIndex        =   10
      Top             =   720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   20643841
      CurrentDate     =   41810
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   3240
      TabIndex        =   11
      Top             =   720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   20643841
      CurrentDate     =   41810
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   " s/d"
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
      Left            =   2640
      TabIndex        =   12
      Top             =   720
      Width           =   495
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
      Left            =   960
      TabIndex        =   7
      Top             =   3600
      Width           =   1695
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
      Left            =   960
      TabIndex        =   5
      Top             =   2280
      Width           =   1695
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
      Left            =   960
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C0C000&
      FillStyle       =   0  'Solid
      Height          =   4095
      Left            =   240
      Top             =   240
      Width           =   4935
   End
End
Attribute VB_Name = "FormFilterPenjualan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strsql As String
Dim RsLihat As Recordset
Private Sub cmdLaporan_Click()
If OptPeriode.Value = True Then
    strsql = "select p.idPenjualan,namaBarang,jumlah,dp.hargaJual,(dp.hargaJual*jumlah) as subtotal from"
    strsql = strsql & " penjualan p join detailPenjualan dp"
    strsql = strsql & " on dp.idPenjualan=p.idPenjualan"
    strsql = strsql & " join barang b on b.idBarang=dp.idBarang"
    strsql = strsql & " where tglJual between '" & Format(DTPicker1.Value, "YYYY-MM-DD") & "' and '" & Format(DTPicker2.Value, "YYYY-MM-DD") & "'"

    Set RsLihat = Conn.Execute(strsql)
    If Not RsLihat.EOF Then
    LapPeriode
    Else
        MsgBox "Data Tidak Tersedia", vbInformation, "Informasi"
    End If
    
ElseIf OptBulan.Value = True Then
    strsql = "select * "
    strsql = strsql & " From penjualan"
    strsql = strsql & " where month(tglJual)='" & Left(CmbBulan.Text, 2) & "'"
    strsql = strsql & " and year(tglJual)='" & CmbTahun.Text & "'"

    Set RsLihat = Conn.Execute(strsql)
    If Not RsLihat.EOF Then
    LapBulan
    Else
        MsgBox "Data Tidak Tersedia", vbInformation, "Informasi"
    End If
ElseIf OptTahun.Value = True Then
    strsql = "select * "
    strsql = strsql & " From penjualan"
    strsql = strsql & " where year(tglJual)='" & cmbTahun1.Text & "'"

    Set RsLihat = Conn.Execute(strsql)
    If Not RsLihat.EOF Then
    LapTahun
    Else
        MsgBox "Data Tidak Tersedia", vbInformation, "Informasi"
    End If
End If
End Sub

Private Sub Form_Load()
Koneksi
TampilCombo
End Sub

Sub TampilCombo()
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


Sub LapBulan()
Dim RsCek As Recordset

strsql = "select sum(total) as tot"
strsql = strsql & " From penjualan"
strsql = strsql & " where month(tglJual)='" & Left(CmbBulan.Text, 2) & "'"
strsql = strsql & " and year(tglJual)='" & CmbTahun.Text & "'"
Set RsCek = Conn.Execute(strsql)

strsql = "select tglJual,sum(total) as tot"
strsql = strsql & " From penjualan"
strsql = strsql & " where month(tglJual)='" & Left(CmbBulan.Text, 2) & "'"
strsql = strsql & " and year(tglJual)='" & CmbTahun.Text & "'"
strsql = strsql & " group by tglJual"

With LaporanjualPer
        .DataControl1.ConnectionString = strConn
        .DataControl1.Source = strsql
        .Field6.Text = FormatCurrency(RsCek(0), 2)
        .Label14.Caption = "Laporan Penjualan Bulan : " + Mid(CmbBulan.Text, 6, 20) + " Tahun : " + CmbTahun.Text
        .Show
    End With
End Sub

Sub LapTahun()
Dim RsCek As Recordset

strsql = "select sum(total) as tot"
strsql = strsql & " From penjualan"
strsql = strsql & " where year(tglJual)='" & cmbTahun1.Text & "'"
Set RsCek = Conn.Execute(strsql)

strsql = "select month(tglJual),sum(total) as tot"
strsql = strsql & " From penjualan"
strsql = strsql & " where year(tglJual)='" & cmbTahun1.Text & "'"
strsql = strsql & " group by month(tglJual)"

With LaporanjualTahun
        .DataControl1.ConnectionString = strConn
        .DataControl1.Source = strsql
        .Label14.Caption = "Laporan Penjualan Tahun : " + cmbTahun1.Text
        .Field6.Text = FormatCurrency(RsCek(0), 2)
        .Show
    End With
End Sub

Sub LapPeriode()
Dim RsCek As Recordset

strsql = "select sum(total) "
strsql = strsql & " From penjualan"
strsql = strsql & " where tglJual between '" & Format(DTPicker1.Value, "YYYY-MM-DD") & "' and '" & Format(DTPicker2.Value, "YYYY-MM-DD") & "'"
Set RsCek = Conn.Execute(strsql)

strsql = "select p.idPenjualan,namaBarang,jumlah,dp.hargaJual,(dp.hargaJual*jumlah) as subtotal from"
    strsql = strsql & " penjualan p join detailPenjualan dp"
    strsql = strsql & " on dp.idPenjualan=p.idPenjualan"
    strsql = strsql & " join barang b on b.idBarang=dp.idBarang"
    strsql = strsql & " where tglJual between '" & Format(DTPicker1.Value, "YYYY/MM/DD") & "' and '" & Format(DTPicker2.Value, "YYYY/MM/DD") & "'"


With LaporanPenjualanPeriode
        .DataControl1.ConnectionString = strConn
        .DataControl1.Source = strsql
        .Label14.Caption = "Laporan Penjualan  Periode : " + Format(DTPicker1.Value, "DD MM YYYY") + " - " + Format(DTPicker2.Value, "DD MM YYYY")
        .Field6.Text = FormatCurrency(RsCek(0), 2)
        .Show
    End With
End Sub
