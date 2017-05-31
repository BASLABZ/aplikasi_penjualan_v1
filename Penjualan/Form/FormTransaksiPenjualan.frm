VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FormTransaksiPenjualan 
   Caption         =   "Transaksi Penjualan"
   ClientHeight    =   7290
   ClientLeft      =   5760
   ClientTop       =   2505
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FormTransaksiPenjualan.frx":0000
   ScaleHeight     =   7290
   ScaleWidth      =   9855
   Begin VB.CommandButton CmdBatal 
      Caption         =   "Batal"
      Height          =   495
      Left            =   8400
      Picture         =   "FormTransaksiPenjualan.frx":548F5
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6360
      Width           =   1095
   End
   Begin VB.TextBox TxtStok 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox TxtBayar 
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
      Height          =   375
      Left            =   4080
      TabIndex        =   21
      Top             =   6120
      Width           =   2175
   End
   Begin VB.CommandButton CmdTambah 
      Caption         =   "Tambah"
      Height          =   495
      Left            =   8280
      Picture         =   "FormTransaksiPenjualan.frx":54B60
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox TxtIdBarang 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox TxtHargaBarang 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox TxtNamaBarang 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox TxtJumlah 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   7080
      TabIndex        =   4
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox TxtSubtotal 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "Simpan"
      Height          =   495
      Left            =   7080
      Picture         =   "FormTransaksiPenjualan.frx":54DC8
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton CmdHapus 
      Caption         =   "Hapus"
      Height          =   495
      Left            =   6960
      Picture         =   "FormTransaksiPenjualan.frx":55033
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3240
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListViewBarang 
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   3836
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Id Barang"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nama Barang"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Harga"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Jumlah"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Sub Total"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Stok"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   25
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label LabelKembali 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   23
      Top             =   6720
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Kembali"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   22
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Bayar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   20
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Id Pembelian"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   19
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label NoId 
      BackStyle       =   0  'Transparent
      Caption         =   "No Id"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   18
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Pembelian"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   17
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label LabelTanggal 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Pembelian"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   16
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Id Barang"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   15
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Harga Barang"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   13
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Subtotal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   12
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Barang"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label LabelTotal 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   6720
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C0C000&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   120
      Top             =   120
      Width           =   9495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00C0C000&
      FillStyle       =   0  'Solid
      Height          =   2175
      Left            =   120
      Top             =   960
      Width           =   9495
   End
End
Attribute VB_Name = "FormTransaksiPenjualan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Nodetail As String
Sub GenerateNomorPenjualan()
Dim RsNo As Recordset
Dim t As Integer
Dim No As String

SQL = "select *  from penjualan ORDER BY idPenjualan DESC"
Set RsNo = Conn.Execute(SQL)
If RsNo.EOF = True Then
NoId.Caption = "P" + "001"
Else
t = Val(Mid(RsNo("idPenjualan"), 2, 3))
No = "P" + Format(Str(t + 1), "000")
NoId.Caption = No
End If
End Sub

Sub TambahKeList()
With ListViewBarang.ListItems.Add
    .Text = TxtIdBarang.Text
    .SubItems(1) = TxtNamaBarang.Text
    .SubItems(2) = TxtHargaBarang.Text
    .SubItems(3) = TxtJumlah.Text
    .SubItems(4) = TxtSubtotal.Text
End With
End Sub

Sub FungsiCariTotal()
Dim Total As Double
Total = 0
For i = 1 To ListViewBarang.ListItems.Count
    With ListViewBarang.ListItems.Item(i)
            Total = Total + Val(.SubItems(4))
    End With
Next
LabelTotal.Caption = Total
End Sub

Sub HapusBarang()

If ListViewBarang.ListItems.Count = 0 Then
    MsgBox "Data Barang Belum Dipilih ", vbCritical, "Peringatan"
Else
    ListViewBarang.ListItems.Remove (ListViewBarang.SelectedItem.Index)
End If
End Sub

Sub ClearBarang()
TxtIdBarang.Text = ""
TxtNamaBarang.Text = ""
TxtJumlah.Text = ""
TxtSubtotal.Text = ""
TxtHargaBarang.Text = ""
TxtStok.Text = ""
End Sub

Sub SimpanPenjualan()
strsql = "INSERT INTO Penjualan VALUES("
        strsql = strsql & "'" & NoId.Caption & "',"
        strsql = strsql & "'" & FormMenuUtama.StatusBar1.Panels(2) & "',"
        strsql = strsql & "'" & Format(LabelTanggal.Caption, "YYYY-MM-DD") & "',"
        strsql = strsql & "" & LabelTotal.Caption & ")"
        Conn.Execute (strsql)
        MsgBox "Berhasil Simpan", vbInformation, "Berhasil"
End Sub

Sub SimpanDetail()
For i = 1 To ListViewBarang.ListItems.Count
With ListViewBarang.ListItems.Item(i)
        strsql = "INSERT INTO DetailPenjualan values ('" & NoId.Caption & "','" & .Text & "'," & .SubItems(2) & "," & .SubItems(3) & ")"
        Conn.Execute (strsql)

        strsql = "update barang set stok=Stok -" & Val(.SubItems(3)) & " where idBarang='" & .Text & "'"
        Conn.Execute (strsql)
End With
Next
End Sub

Sub Bersih()
LabelTotal.Caption = "0"
GenerateNomorPenjualan
ListViewBarang.ListItems.Clear
LabelKembali.Caption = "0"
TxtBayar.Text = ""
End Sub

Sub Pilih()
Dim RsCek As Recordset
Dim SQLCek As String

SQLCek = "Select * from barang where idBarang='" & ListViewBarang.SelectedItem.Text & "'"
Set RsCek = Conn.Execute(SQLCek)


If ListViewBarang.ListItems.Count = 0 Then
    MsgBox "Data Masih Belum Tersedia ", vbCritical, "Peringatan"
Else
    TxtIdBarang.Text = ListViewBarang.SelectedItem.Text
    TxtNamaBarang.Text = ListViewBarang.SelectedItem.SubItems(1)
    TxtJumlah.Text = ListViewBarang.SelectedItem.SubItems(3)
    TxtHargaBarang.Text = ListViewBarang.SelectedItem.SubItems(2)
    TxtSubtotal.Text = ListViewBarang.SelectedItem.SubItems(4)
    TxtStok.Text = RsCek!stok
End If
End Sub


Sub CetakLaporan()
strsql = "select dp.idPenjualan,dp.idBarang,namaBarang,dp.hargaJual,dp.jumlah,dp.hargaJual*dp.jumlah as subtotal "
strsql = strsql & "from detailPenjualan dp join Barang p "
strsql = strsql & "on p.idBarang=dp.idBarang "
strsql = strsql & "where dp.idPenjualan='" & NoId.Caption & "'"
With NotaPenjualan
        .DataControl1.ConnectionString = strConn
        .DataControl1.Source = strsql
        .Show
    End With
NotaPenjualan.Field6 = FormatCurrency(LabelTotal.Caption, 2)
NotaPenjualan.Field7 = FormatCurrency(TxtBayar.Text, 2)
NotaPenjualan.Field8 = FormatCurrency(LabelKembali.Caption, 2)
End Sub

Private Sub Cmdbatal_Click()
ClearBarang
Bersih
End Sub

Private Sub CmdHapus_Click()
HapusBarang
ClearBarang
FungsiCariTotal
End Sub

Private Sub CmdSimpan_Click()
If TxtBayar.Text = "" Or _
    ListViewBarang.ListItems.Count = 0 Then
    MsgBox "Silahkan Lakukan Pembayaran", vbExclamation, "Peringatan"
Else
SimpanPenjualan
SimpanDetail
CetakLaporan
Bersih
End If
End Sub

Private Sub CmdTambah_Click()
If Val(TxtJumlah.Text) > Val(TxtStok.Text) Then
    MsgBox "Maaf Stok Tidak mencukupi", vbExclamation, "Peringatan"
    TxtJumlah.Text = ""
Else
Dim status As Boolean
    If TxtIdBarang.Text = "" Or TxtJumlah.Text = "" Then
    MsgBox "Data Barang Belum Lengkap", vbCritical, "Peringatan"
    Else
    status = False
    For i = 1 To ListViewBarang.ListItems.Count
        With ListViewBarang.ListItems.Item(i)
        
        If TxtIdBarang.Text = .Text Then
            status = True
        End If
        End With
    Next
    If status = False Then
            TambahKeList
            ClearBarang
            FungsiCariTotal
            Else
            HapusBarang
            TambahKeList
            ClearBarang
            FungsiCariTotal
        End If
    End If
End If
End Sub

Private Sub Form_Load()
Koneksi
LabelTanggal.Caption = Format(Now, "DD-MM-YYYY")
GenerateNomorPenjualan
End Sub

Private Sub ListViewBarang_Click()
Pilih
End Sub

Private Sub TxtBayar_KeyPress(KeyAscii As Integer)
If InStr("0123456789", Chr(KeyAscii)) = 0 Then
If KeyAscii <> vbKeyBack Then
If KeyAscii <> 13 Then
KeyAscii = 0
End If
End If
End If
If KeyAscii = 13 Then
        If Val(TxtBayar.Text) >= (LabelTotal.Caption) Then
            LabelKembali.Caption = Val(TxtBayar.Text) - LabelTotal.Caption
        Else
            MsgBox "Maaf Uang Kurang", vbExclamation, "Informasi"
        End If
End If
End Sub

Private Sub TxtIdBarang_Click()
BrowseBarangJual.Show
End Sub

Private Sub TxtJumlah_Change()
TxtSubtotal.Text = Val(TxtHargaBarang.Text) * Val(TxtJumlah.Text)
End Sub



Private Sub TxtJumlah_KeyPress(KeyAscii As Integer)
If InStr("0123456789", Chr(KeyAscii)) = 0 Then
If KeyAscii <> vbKeyBack Then
KeyAscii = 0
End If
End If
End Sub
