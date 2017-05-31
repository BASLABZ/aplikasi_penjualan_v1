VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form FormTransaksiPembelian 
   Caption         =   "Transaksi Pembelian"
   ClientHeight    =   8010
   ClientLeft      =   6030
   ClientTop       =   1980
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FormTransaksi.frx":0000
   ScaleHeight     =   8010
   ScaleWidth      =   9960
   Begin VB.CommandButton CmdBatal 
      Caption         =   "Batal"
      Height          =   495
      Left            =   7440
      Picture         =   "FormTransaksi.frx":548F5
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   7320
      Width           =   1095
   End
   Begin VB.CommandButton CmdHapus 
      Caption         =   "Hapus"
      Height          =   495
      Left            =   6960
      Picture         =   "FormTransaksi.frx":54B60
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "Simpan"
      Height          =   495
      Left            =   8640
      Picture         =   "FormTransaksi.frx":54CAE
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7320
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListViewBarang 
      Height          =   2655
      Left            =   240
      TabIndex        =   19
      Top             =   4440
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   4683
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
   Begin VB.TextBox TxtSubtotal 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   2880
      Width           =   2175
   End
   Begin VB.TextBox TxtJumlah 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   6960
      TabIndex        =   17
      Top             =   2400
      Width           =   2175
   End
   Begin VB.TextBox TxtNamaBarang 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox TxtHargaBarang 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   3480
      Width           =   1695
   End
   Begin VB.TextBox TxtIdBarang 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton CmdTambah 
      Caption         =   "Tambah"
      Height          =   495
      Left            =   8160
      Picture         =   "FormTransaksi.frx":54F19
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox TxtNamaSupplier 
      Appearance      =   0  'Flat
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
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1080
      Width           =   2655
   End
   Begin VB.TextBox TXtIdSupplier 
      Appearance      =   0  'Flat
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
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1080
      Width           =   1935
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
      Left            =   1560
      TabIndex        =   23
      Top             =   7320
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
      Left            =   240
      TabIndex        =   22
      Top             =   7320
      Width           =   1095
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
      Left            =   480
      TabIndex        =   16
      Top             =   2880
      Width           =   2055
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
      Left            =   4920
      TabIndex        =   11
      Top             =   2880
      Width           =   1935
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
      Left            =   4920
      TabIndex        =   10
      Top             =   2400
      Width           =   1935
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
      Left            =   480
      TabIndex        =   9
      Top             =   3480
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
      Left            =   480
      TabIndex        =   8
      Top             =   2400
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
      Left            =   6960
      TabIndex        =   6
      Top             =   480
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
      Left            =   4920
      TabIndex        =   5
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Supplier"
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
      Left            =   4920
      TabIndex        =   4
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Id Supplier"
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
      TabIndex        =   2
      Top             =   1080
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
      Left            =   2640
      TabIndex        =   1
      Top             =   480
      Width           =   2055
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
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C0C000&
      FillStyle       =   0  'Solid
      Height          =   1455
      Left            =   240
      Top             =   240
      Width           =   9495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00C0C000&
      FillStyle       =   0  'Solid
      Height          =   2175
      Left            =   240
      Top             =   2040
      Width           =   9495
   End
End
Attribute VB_Name = "FormTransaksiPembelian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Nodetail As String
Sub GenerateNomorPembelian()
Dim RsNo As Recordset
Dim t As Integer
Dim No As String

SQL = "select *  from pembelian ORDER BY idPembelian DESC"
Set RsNo = Conn.Execute(SQL)
If RsNo.EOF = True Then
NoId.Caption = "P" + "001"
Else
t = Val(Mid(RsNo("idPembelian"), 2, 3))
No = "P" + Format(Str(t + 1), "000")
NoId.Caption = No
End If
End Sub


Sub GenerateNomorDetail()
Dim RsNo As Recordset
Dim t As Integer
Dim No As String

SQL = "select *  from detailPembelian ORDER BY idDetailPembelian DESC"
Set RsNo = Conn.Execute(SQL)
If RsNo.EOF = True Then
Nodetail = "DP" + "0001"
Else
t = Val(Mid(RsNo("idDetailPembelian"), 3, 4))
No = "DP" + Format(Str(t + 1), "0000")
Nodetail = No
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
End Sub

Sub SimpanPembelian()
strsql = "INSERT INTO Pembelian VALUES("
        strsql = strsql & "'" & NoId.Caption & "',"
        strsql = strsql & "'" & FormMenuUtama.StatusBar1.Panels(2) & "',"
        strsql = strsql & "'" & TXtIdSupplier.Text & "',"
        strsql = strsql & "'" & Format(LabelTanggal.Caption, "YYYY-MM-DD") & "',"
        strsql = strsql & "" & LabelTotal.Caption & ")"
        Conn.Execute (strsql)
        MsgBox "Berhasil Simpan", vbInformation, "Berhasil"
End Sub

Sub SimpanDetail()
For i = 1 To ListViewBarang.ListItems.Count
GenerateNomorDetail
With ListViewBarang.ListItems.Item(i)
GenerateNomorDetail
        strsql = "INSERT INTO Detailpembelian values ('" & Nodetail & "','" & NoId.Caption & "','" & .Text & "'," & .SubItems(2) & "," & .SubItems(3) & ")"
        Conn.Execute (strsql)

        strsql = "update barang set stok=Stok +" & Val(jumlah + .SubItems(3)) & " where idBarang='" & .Text & "'"
        Conn.Execute (strsql)
End With
Next
End Sub

Sub Bersih()
LabelTotal.Caption = "0"
GenerateNomorPembelian
ListViewBarang.ListItems.Clear
TXtIdSupplier.Text = ""
TxtNamaSupplier.Text = ""
End Sub



Sub Pilih()
If ListViewBarang.ListItems.Count = 0 Then
    MsgBox "Data Masih Belum Tersedia ", vbCritical, "Peringatan"
Else
    TxtIdBarang.Text = ListViewBarang.SelectedItem.Text
    TxtNamaBarang.Text = ListViewBarang.SelectedItem.SubItems(1)
    TxtJumlah.Text = ListViewBarang.SelectedItem.SubItems(3)
    TxtHargaBarang.Text = ListViewBarang.SelectedItem.SubItems(2)
    TxtSubtotal.Text = ListViewBarang.SelectedItem.SubItems(4)
End If
End Sub

Private Sub Cmdbatal_Click()
Bersih
ClearBarang
End Sub

Private Sub CmdHapus_Click()
HapusBarang
End Sub

Private Sub CmdSimpan_Click()
If TXtIdSupplier.Text = "" Or ListViewBarang.ListItems.Count = 0 Then
    MsgBox "Isikan data secara lengkap", vbExclamation, "Peringatan"
Else
SimpanPembelian
SimpanDetail
Bersih
End If
End Sub

Private Sub CmdTambah_Click()
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
End Sub

Private Sub Form_Load()
Koneksi
GenerateNomorPembelian
LabelTanggal.Caption = Format(Now, "DD-MM-YYYY")
End Sub

Private Sub ListViewBarang_Click()
Pilih
End Sub

Private Sub TxtIdBarang_Click()
BrowseBarangBeli.Show
End Sub

Private Sub TXtIdSupplier_Click()
BrowseSuplier.Show
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
