VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormFilterpembelian 
   Caption         =   "Pemilihian Laporan Pembelian"
   ClientHeight    =   5490
   ClientLeft      =   8100
   ClientTop       =   1980
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FormFilterpembelian.frx":0000
   ScaleHeight     =   5490
   ScaleWidth      =   5475
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   3360
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
      Format          =   20709377
      CurrentDate     =   41810
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   960
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
      Format          =   20709377
      CurrentDate     =   41810
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
      Picture         =   "FormFilterpembelian.frx":548F5
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4560
      Width           =   1215
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
      Left            =   2760
      TabIndex        =   7
      Top             =   3720
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
      Left            =   2760
      TabIndex        =   4
      Top             =   2400
      Width           =   2175
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
      Left            =   2760
      TabIndex        =   3
      Top             =   1920
      Width           =   2175
   End
   Begin VB.OptionButton OptTahun 
      BackColor       =   &H00C0C000&
      Caption         =   "Total Pembelian Tahunan"
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
      TabIndex        =   2
      Top             =   3000
      Width           =   3855
   End
   Begin VB.OptionButton OptBulan 
      BackColor       =   &H00C0C000&
      Caption         =   "Total Pembelian Bulanan"
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
      Top             =   1200
      Width           =   3375
   End
   Begin VB.OptionButton OptPeriode 
      BackColor       =   &H00C0C000&
      Caption         =   "Pembelian Periode"
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
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   2055
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
      Left            =   2760
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
      TabIndex        =   8
      Top             =   3720
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
      TabIndex        =   6
      Top             =   2400
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
      TabIndex        =   5
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C0C000&
      FillStyle       =   0  'Solid
      Height          =   4335
      Left            =   240
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "FormFilterpembelian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strsql As String
Dim RsLihat As Recordset

Private Sub cmdLaporan_Click()
Dim RsCekada As Recordset

If OptPeriode.Value = True Then
   strsql = "select p.idPembelian,namaBarang,jumlah,dp.HargaBeli,(dp.HargaBeli*jumlah) as subtotal from"
    strsql = strsql & " Pembelian p join DetailPembelian dp"
    strsql = strsql & " on dp.idPembelian=p.idPembelian"
    strsql = strsql & " join barang b on b.idBarang=dp.idBarang"
    strsql = strsql & " where TanggalBeli between '" & Format(DTPicker1.Value, "YYYY-MM-DD") & "' and '" & Format(DTPicker2.Value, "YYYY-MM-DD") & "'"

    Set RsLihat = Conn.Execute(strsql)
    If Not RsLihat.EOF Then
    lapPeriode
    Else
        MsgBox "Data Tidak Tersedia", vbInformation, "Informasi"
    End If
ElseIf OptBulan.Value = True Then
    strsql = "select *"
    strsql = strsql & " From Pembelian"
    strsql = strsql & " where month(TanggalBeli)='" & Left(CmbBulan.Text, 2) & "'"
    strsql = strsql & " and year(TanggalBeli)='" & CmbTahun.Text & "'"

    Set RsLihat = Conn.Execute(strsql)
    If Not RsLihat.EOF Then
    LapBulan
    Else
        MsgBox "Data Tidak Tersedia", vbInformation, "Informasi"
    End If
ElseIf OptTahun.Value = True Then
    strsql = "select *"
    strsql = strsql & " From Pembelian"
    strsql = strsql & " where year(TanggalBeli)='" & cmbTahun1.Text & "'"
    
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
strsql = strsql & " From Pembelian"
strsql = strsql & " where month(tanggalBeli)='" & Left(CmbBulan.Text, 2) & "'"
strsql = strsql & " and year(TanggalBeli)='" & CmbTahun.Text & "'"
Set RsCek = Conn.Execute(strsql)

strsql = "select TanggalBeli,sum(total) as tot"
strsql = strsql & " From Pembelian"
strsql = strsql & " where month(TanggalBeli)='" & Left(CmbBulan.Text, 2) & "'"
strsql = strsql & " and year(TanggalBeli)='" & CmbTahun.Text & "'"
strsql = strsql & " group by TanggalBeli"

With LaporanBeliBulan
        .DataControl1.ConnectionString = strConn
        .DataControl1.Source = strsql
        .Field6.Text = FormatCurrency(RsCek(0), 2)
        .Label14.Caption = "Laporan Pembelian Bulan : " + Mid(CmbBulan.Text, 6, 20) + " Tahun : " + CmbTahun.Text
        .Show
    End With
End Sub

Sub LapTahun()
Dim RsCek As Recordset

strsql = "select sum(total) as tot"
strsql = strsql & " From Pembelian"
strsql = strsql & " where year(TanggalBeli)='" & cmbTahun1.Text & "'"
Set RsCek = Conn.Execute(strsql)

strsql = "select month(TanggalBeli),sum(total) as tot"
strsql = strsql & " From Pembelian"
strsql = strsql & " where year(TanggalBeli)='" & cmbTahun1.Text & "'"
strsql = strsql & " group by month(TanggalBeli)"

With LaporanBeliTahun
        .DataControl1.ConnectionString = strConn
        .DataControl1.Source = strsql
        .Label14.Caption = "Laporan Pembelian Tahun : " + cmbTahun1.Text
        .Field6.Text = FormatCurrency(RsCek(0), 2)
        .Show
    End With
End Sub

Sub lapPeriode()
Dim RsCek As Recordset

strsql = "select sum(total) as tot"
strsql = strsql & " From Pembelian"
  strsql = strsql & " where TanggalBeli between '" & Format(DTPicker1.Value, "YYYY/MM/DD") & "' and '" & Format(DTPicker2.Value, "YYYY/MM/DD") & "'"
Set RsCek = Conn.Execute(strsql)

strsql = "select p.idPembelian,namaBarang,jumlah,dp.HargaBeli,(dp.HargaBeli*jumlah) as subtotal from"
strsql = strsql & " Pembelian p join DetailPembelian dp"
strsql = strsql & " on dp.idPembelian=p.idPembelian"
strsql = strsql & " join barang b on b.idBarang=dp.idBarang"
strsql = strsql & " where TanggalBeli between '" & Format(DTPicker1.Value, "YYYY-MM-DD") & "' and '" & Format(DTPicker2.Value, "YYYY-MM-DD") & "'"

With LaporanPembelianPeriode
        .DataControl1.ConnectionString = strConn
        .DataControl1.Source = strsql
        .Label14.Caption = "Laporan Pembelian Periode : " + Format(DTPicker1.Value, "DD MM YYYY") + "-" + Format(DTPicker2.Value, "DD MM YYYY")
        .Field6.Text = FormatCurrency(RsCek(0), 2)
        .Show
    End With
End Sub

