VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FormBarang 
   Caption         =   "Barang"
   ClientHeight    =   6885
   ClientLeft      =   6030
   ClientTop       =   1725
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FormBarang.frx":0000
   ScaleHeight     =   6885
   ScaleWidth      =   9945
   Begin VB.TextBox TXtStok 
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
      Left            =   8040
      TabIndex        =   19
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox TxtHargaJual 
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
      Left            =   8040
      TabIndex        =   18
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox TxtHargaBeli 
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
      Left            =   8040
      TabIndex        =   17
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox TxtNamaBarang 
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
      Left            =   2760
      TabIndex        =   16
      Top             =   1800
      Width           =   2895
   End
   Begin VB.ComboBox ComboJenisBarang 
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
      Height          =   360
      Left            =   2760
      TabIndex        =   15
      Top             =   1200
      Width           =   2895
   End
   Begin VB.TextBox TxtIdBarang 
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
      Left            =   2760
      TabIndex        =   14
      Top             =   600
      Width           =   2895
   End
   Begin VB.TextBox TxtCari 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2400
      TabIndex        =   10
      Top             =   3240
      Width           =   5295
   End
   Begin VB.CommandButton Cmdbatal 
      Caption         =   "Batal"
      Height          =   495
      Left            =   6720
      Picture         =   "FormBarang.frx":548F5
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton CmdHapus 
      Caption         =   "Hapus"
      Height          =   495
      Left            =   5640
      Picture         =   "FormBarang.frx":54B60
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "Edit"
      Height          =   495
      Left            =   4560
      Picture         =   "FormBarang.frx":54CAE
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "Simpan"
      Height          =   495
      Left            =   3480
      Picture         =   "FormBarang.frx":54ED7
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton CmdTmbah 
      Caption         =   "Tambah"
      Height          =   495
      Left            =   2400
      Picture         =   "FormBarang.frx":55142
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid 
      Height          =   2655
      Left            =   600
      TabIndex        =   11
      Top             =   3960
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   4683
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label7 
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
      Left            =   6120
      TabIndex        =   13
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Harga Jual"
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
      Left            =   6120
      TabIndex        =   12
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cari"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   720
      TabIndex        =   9
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Harga Beli"
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
      Left            =   6120
      TabIndex        =   3
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label3 
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
      Left            =   840
      TabIndex        =   2
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Jenis Barang"
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
      Left            =   840
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label1 
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
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C0C000&
      FillStyle       =   0  'Solid
      Height          =   1935
      Left            =   600
      Top             =   480
      Width           =   9135
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00C0C000&
      Height          =   615
      Left            =   600
      Top             =   3120
      Width           =   9135
   End
End
Attribute VB_Name = "FormBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmdbatal_Click()
Bersih
CLoad
End Sub

Private Sub CmdEdit_Click()
Cubah
End Sub

Private Sub CmdHapus_Click()
Hapus
CLoad
End Sub

Private Sub CmdSimpan_Click()
CekSimpan
CLoad
End Sub

Private Sub CmdTmbah_Click()
CTambah
GenerateNomor
End Sub

Private Sub DataGrid_Click()
Cpilih
End Sub

Private Sub Form_Load()
Koneksi
TampilData
TampilCombo
CLoad
End Sub

Private Sub TxtCari_Change()
Pencarian
End Sub


Sub TabelHeader()
With Datagrid
    .RowHeight = 280
    .HeadFont = "Arial"
    .HeadLines = 2
    .Columns(0).Caption = "ID BARANG"
    .Columns(1).Caption = "ID J.BARANG"
    .Columns(2).Caption = "NAMA BARANG"
    .Columns(3).Caption = "HARGA BELI"
    .Columns(4).Caption = "HARGA JUAL"
    .Columns(5).Caption = "STOK"
    .Columns(0).Width = 1500
    .Columns(1).Width = 1500
    .Columns(2).Width = 2000
    .Columns(3).Width = 1000
    .Columns(4).Width = 1000
    .Columns(5).Width = 1000
End With
End Sub

Sub TampilData()
'TabelHeader
Set rs = New Recordset
SQL = "Select * from barang"
rs.Open SQL, Conn, adOpenStatic, adLockReadOnly
Set Datagrid.DataSource = rs
TabelHeader
End Sub

Sub Simpan()
        SQL = "INSERT INTO Barang VALUES("
        SQL = SQL & "'" & TxtIdBarang.Text & "',"
        SQL = SQL & "'" & Mid(ComboJenisBarang.Text, 1, 4) & "',"
        SQL = SQL & "'" & TxtNamaBarang.Text & "',"
        SQL = SQL & "" & TxtHargaBeli.Text & ","
        SQL = SQL & "" & TxtHargaJual.Text & ","
        SQL = SQL & "" & TxtStok.Text & ")"
        Conn.Execute (SQL)
        MsgBox "Berhasil User Simpan", vbInformation, "Informasi"
        TampilData
End Sub

Sub ubah()
    SQL = "UPDATE Barang set "
    SQL = SQL & "NamaBarang='" & TxtNamaBarang.Text & "',"
    SQL = SQL & "hargaJual=" & TxtHargaJual.Text & ","
    SQL = SQL & "hargaBeli=" & TxtHargaBeli.Text & ","
    SQL = SQL & "Stok=" & TxtStok.Text & ","
    SQL = SQL & "IDJenisBarang='" & Mid(ComboJenisBarang.Text, 1, 4) & "'"
    SQL = SQL & "Where idbarang='" & TxtIdBarang.Text & "'"
    Conn.Execute (SQL)
    MsgBox "Data Berhasil Diubah", vbInformation, "Berhasil"
    TampilData
End Sub

Sub CekSimpan()
Dim RsCek As Recordset

SQL = "Select * from barang where idbarang='" & TxtIdBarang.Text & "'"
Set RsCek = Conn.Execute(SQL)
If Not RsCek.EOF Then
    ubah
Else
    Simpan
End If
End Sub

Sub Hapus()
SQL = "DELETE from barang where idbarang ='" & TxtIdBarang.Text & "'"
    Conn.Execute (SQL)
    MsgBox "Berhasil Dihapus", vbInformation, "Berhasil"
    TampilData
End Sub

Sub GenerateNomor()
Dim RsNo As Recordset
Dim t As Integer
Dim No As String

SQL = "select *  from Barang ORDER BY idBarang DESC"
Set RsNo = Conn.Execute(SQL)
If RsNo.EOF = True Then
TxtIdBarang.Text = "B" + "001"
Else
t = Val(Mid(RsNo("Idbarang"), 2, 3))
No = "B" + Format(Str(t + 1), "000")
TxtIdBarang.Text = No
End If
End Sub


Sub Bersih()
TxtIdBarang.Text = ""
TxtNamaBarang.Text = ""
TxtStok.Text = ""
TxtHargaBeli.Text = ""
TxtHargaJual.Text = ""
ComboJenisBarang.Text = ""
End Sub

Sub Pilih()
TxtIdBarang.Text = rs(0)
TxtNamaBarang.Text = rs(2)
TxtStok.Text = rs(5)
TxtHargaBeli.Text = rs(3)
TxtHargaJual.Text = rs(4)
ComboJenisBarang.Text = rs(1)
End Sub

Sub Pencarian()
Set rs = New Recordset
SQL = "Select * from barang"
SQL = SQL & " where namaBarang like '%" & TxtCari.Text & "%'"
rs.Open SQL, Conn, adOpenStatic, adLockReadOnly
Set Datagrid.DataSource = rs
TabelHeader
End Sub

Sub TampilCombo()
Dim rsjenis As Recordset
Set rsjenis = New Recordset
Dim sqlJenis As String

sqlJenis = "SELECT * FROM JenisBarang"
rsjenis.Open sqlJenis, Conn, adOpenStatic, adLockReadOnly

Do While Not rsjenis.EOF
    ComboJenisBarang.AddItem rsjenis!idJenisBarang + "-" + rsjenis!NamaJenisBarang
    rsjenis.MoveNext
Loop
End Sub

Sub Hidup()
TxtIdBarang.Enabled = False
TxtNamaBarang.Enabled = True
ComboJenisBarang.Enabled = True
TxtStok.Enabled = True
TxtHargaBeli.Enabled = True
TxtHargaJual.Enabled = True
ComboJenisBarang.SetFocus
End Sub

Sub Mati()
TxtIdBarang.Enabled = False
TxtNamaBarang.Enabled = False
ComboJenisBarang.Enabled = False
TxtStok.Enabled = False
TxtHargaBeli.Enabled = False
TxtHargaJual.Enabled = False
End Sub

Sub CLoad()
CmdBatal.Enabled = False
CmdEdit.Enabled = False
CmdHapus.Enabled = False
CmdSimpan.Enabled = False
CmdTmbah.Enabled = True
Bersih
Mati
End Sub

Sub CTambah()
CmdBatal.Enabled = True
CmdEdit.Enabled = False
CmdHapus.Enabled = False
CmdSimpan.Enabled = True
CmdTmbah.Enabled = False
Bersih
Hidup
End Sub

Sub Cpilih()
Pilih
CmdBatal.Enabled = True
CmdEdit.Enabled = True
CmdHapus.Enabled = True
CmdSimpan.Enabled = False
CmdTmbah.Enabled = False
Mati
End Sub

Sub Cubah()
CmdBatal.Enabled = True
CmdEdit.Enabled = False
CmdHapus.Enabled = False
CmdSimpan.Enabled = True
CmdTmbah.Enabled = False
Hidup
End Sub

Private Sub TxtHargaBeli_KeyPress(KeyAscii As Integer)
If InStr("0123456789", Chr(KeyAscii)) = 0 Then
If KeyAscii <> vbKeyBack Then
KeyAscii = 0
End If
End If
End Sub

Private Sub TxtHargaJual_KeyPress(KeyAscii As Integer)
If InStr("0123456789", Chr(KeyAscii)) = 0 Then
If KeyAscii <> vbKeyBack Then
KeyAscii = 0
End If
End If
End Sub

Private Sub TxtNamaBarang_Change()
TxtNamaBarang.MaxLength = 40
End Sub

Private Sub TXtStok_KeyPress(KeyAscii As Integer)
If InStr("0123456789", Chr(KeyAscii)) = 0 Then
If KeyAscii <> vbKeyBack Then
KeyAscii = 0
End If
End If
End Sub
