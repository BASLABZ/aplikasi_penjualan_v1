VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Grafikterlaris 
   Caption         =   "Grafik Barang Terlaris"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Grafikterlaris.frx":0000
   ScaleHeight     =   8520
   ScaleWidth      =   9960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLaporan 
      Caption         =   "LihatData"
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
      Left            =   4560
      Picture         =   "Grafikterlaris.frx":548F5
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2160
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
      Left            =   7200
      TabIndex        =   4
      Top             =   960
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
      TabIndex        =   3
      Top             =   1440
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
      TabIndex        =   2
      Top             =   960
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
      Left            =   5400
      TabIndex        =   1
      Top             =   240
      Width           =   2415
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
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   5325
      Left            =   240
      OleObjectBlob   =   "Grafikterlaris.frx":54B80
      TabIndex        =   8
      Top             =   3000
      Width           =   9495
   End
   Begin VB.Line Line1 
      X1              =   5160
      X2              =   5160
      Y1              =   120
      Y2              =   2040
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C0C000&
      FillStyle       =   0  'Solid
      Height          =   1935
      Left            =   240
      Top             =   120
      Width           =   9495
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
      Left            =   5520
      TabIndex        =   7
      Top             =   960
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
      Top             =   1440
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
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "Grafikterlaris"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLaporan_Click()
Dim Rscek As Recordset

If OptBulan.Value = True Then
    If CmbBulan.Text = "" Or CmbTahun.Text = "" Then
        MsgBox "Lengkapai Tahun dan Bulan", vbExclamation, "Peringatan"
    Else
        LihatBulan
    End If

ElseIf OptTahun.Value = True Then
    If cmbTahun1.Text = "" Then
        MsgBox "Lengkapai Tahun", vbExclamation, "Peringatan"
    Else
        LihatTahun
    End If
End If
End Sub

Private Sub Form_Load()
Koneksi
TampilCombo
MSChart1.ColumnCount = 0
End Sub

Sub LihatTahun()
Dim RsLihat As Recordset
Dim RsCount As Recordset

With Me.MSChart1
    Dim i As Integer
    
    strsql = "select count(distinct(namaBarang)) as tot"
    strsql = strsql & " From penjualan p join detailPenjualan dp on dp.idPenjualan=p.idPenjualan"
    strsql = strsql & " Join barang b on b.idbarang=dp.idBarang"
    strsql = strsql & " Join JenisBarang k on b.idJenisBarang=k.idJenisBarang"
    strsql = strsql & " where year(tglJual)='" & cmbTahun1.Text & "'"
    Set RsCount = Conn.Execute(strsql)
    
    .chartType = VtChChartType2dBar
    .ColumnCount = RsCount(0)
    .RowCount = 1
    .RowLabel = "Grafik Tahun " + cmbTahun1.Text
    
    If .ColumnCount = 0 Then
        MsgBox "Tidak Ada data pada periode tersebut"
    End If

    strsql = "select namaBarang,sum(jumlah) as tot"
    strsql = strsql & " From penjualan p join detailPenjualan dp on dp.idPenjualan=p.idPenjualan"
    strsql = strsql & " Join barang b on b.idbarang=dp.idBarang"
    strsql = strsql & " Join JenisBarang k on b.idJenisBarang=k.idJenisBarang"
    strsql = strsql & " where year(tglJual)='" & cmbTahun1.Text & "'"
    strsql = strsql & " Group by namaBarang order by sum(jumlah) desc"
    Set RsLihat = Conn.Execute(strsql)
    
    i = 1
    Do While Not RsLihat.EOF
        .Row = 1
        .Column = i
        .Data = RsLihat(1)
        .ColumnLabel = RsLihat(0)
        .Plot.SeriesCollection(i).DataPoints(-1).DataPointLabel.LocationType = VtChLabelLocationTypeOutside
        RsLihat.MoveNext
        i = i + 1
    Loop
    End With
End Sub


Sub LihatBulan()
Dim RsLihat As Recordset
Dim RsCount As Recordset

With Me.MSChart1
    Dim i As Integer
    
    strsql = "select count(distinct(namaBarang)) as tot"
    strsql = strsql & " From penjualan p join detailPenjualan dp on dp.idPenjualan=p.idPenjualan"
    strsql = strsql & " Join barang b on b.idbarang=dp.idBarang"
    strsql = strsql & " Join JenisBarang k on b.idJenisBarang=k.idJenisBarang"
    strsql = strsql & " where month(tglJual)='" & Left(CmbBulan.Text, 2) & "'"
    strsql = strsql & " and year(tglJual)='" & CmbTahun.Text & "'"
    Set RsCount = Conn.Execute(strsql)
    
    .chartType = VtChChartType2dBar
    .ColumnCount = RsCount(0)
    .RowCount = 1
    .RowLabel = "Grafik Bulan " + Mid(CmbBulan.Text, 6, 30) + " Tahun " + CmbTahun.Text
    
    If .ColumnCount = 0 Then
        MsgBox "Tidak Ada data pada periode tersebut"
    End If
    
    strsql = "select namaBarang,sum(jumlah) as tot"
    strsql = strsql & " From penjualan p join detailPenjualan dp on dp.idPenjualan=p.idPenjualan"
    strsql = strsql & " Join barang b on b.idbarang=dp.idBarang"
    strsql = strsql & " Join JenisBarang k on b.idJenisBarang=k.idJenisBarang"
    strsql = strsql & " where month(tglJual)='" & Left(CmbBulan.Text, 2) & "'"
    strsql = strsql & " and year(tglJual)='" & CmbTahun.Text & "'"
    strsql = strsql & " Group by namaBarang order by sum(jumlah) desc"
    Set RsLihat = Conn.Execute(strsql)
    
    i = 1
    Do While Not RsLihat.EOF
        .Row = 1
        .Column = i
        .Data = RsLihat(1)
        .ColumnLabel = RsLihat(0)
        .Plot.SeriesCollection(i).DataPoints(-1).DataPointLabel.LocationType = VtChLabelLocationTypeOutside
        RsLihat.MoveNext
        i = i + 1
    Loop
    End With
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

