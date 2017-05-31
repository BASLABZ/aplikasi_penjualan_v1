VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FormSupplier 
   Caption         =   "Supplier"
   ClientHeight    =   6315
   ClientLeft      =   6285
   ClientTop       =   2235
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FormSupplier.frx":0000
   ScaleHeight     =   6315
   ScaleWidth      =   9585
   Begin VB.CommandButton CmdTmbah 
      Caption         =   "Tambah"
      Height          =   495
      Left            =   1920
      Picture         =   "FormSupplier.frx":548F5
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "Simpan"
      Height          =   495
      Left            =   3000
      Picture         =   "FormSupplier.frx":54B4B
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "Edit"
      Height          =   495
      Left            =   4080
      Picture         =   "FormSupplier.frx":54DB6
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton CmdHapus 
      Caption         =   "Hapus"
      Height          =   495
      Left            =   5160
      Picture         =   "FormSupplier.frx":54FDF
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Cmdbatal 
      Caption         =   "Batal"
      Height          =   495
      Left            =   6240
      Picture         =   "FormSupplier.frx":5512D
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox TxtCari 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   2760
      Width           =   3255
   End
   Begin VB.TextBox TxtAlamatSupplier 
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
      Height          =   1095
      Left            =   6840
      TabIndex        =   5
      Top             =   600
      Width           =   2295
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
      Left            =   2400
      TabIndex        =   4
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox txtidSupplier 
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
      Left            =   2400
      TabIndex        =   3
      Top             =   600
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid DataGrid 
      Height          =   2655
      Left            =   240
      TabIndex        =   8
      Top             =   3480
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
   Begin VB.Label Label5 
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
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat Supplier"
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
      Left            =   5040
      TabIndex        =   2
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label2 
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
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label1 
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
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00C0C000&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   240
      Top             =   2640
      Width           =   9135
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C0C000&
      FillStyle       =   0  'Solid
      Height          =   1695
      Left            =   240
      Top             =   240
      Width           =   9135
   End
End
Attribute VB_Name = "FormSupplier"
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
CLoad
End Sub

Private Sub TxtAlamatSupplier_Change()
TxtAlamatSupplier.MaxLength = 40
End Sub

Private Sub TxtCari_Change()
Pencarian
End Sub


Sub TabelHeader()
With Datagrid
    .RowHeight = 280
    .HeadFont = "Arial"
    .HeadLines = 2
    .Columns(0).Caption = "ID SUPPLIER"
    .Columns(1).Caption = "NAMA SUPPLIER"
    .Columns(2).Caption = "ALAMAT SUPPLIER"
    .Columns(0).Width = 1500
    .Columns(1).Width = 2000
    .Columns(2).Width = 2500
End With
End Sub

Sub TampilData()
Set rs = New Recordset
SQL = "Select * from Supplier"
rs.Open SQL, Conn, adOpenStatic, adLockReadOnly
Set Datagrid.DataSource = rs
TabelHeader
End Sub

Sub Simpan()
        SQL = "INSERT INTO Supplier VALUES("
        SQL = SQL & "'" & TXtIdSupplier.Text & "',"
        SQL = SQL & "'" & TxtNamaSupplier.Text & "',"
        SQL = SQL & "'" & TxtAlamatSupplier.Text & "')"
        Conn.Execute (SQL)
        MsgBox "Berhasil Jenis Barang Simpan", vbInformation, "Informasi"
        TampilData
End Sub

Sub ubah()
    SQL = "UPDATE Supplier set "
    SQL = SQL & "NamaSupplier='" & TxtNamaSupplier.Text & "',"
    SQL = SQL & "AlamatSupplier='" & TxtAlamatSupplier.Text & "'"
    SQL = SQL & "Where IdSupplier='" & TXtIdSupplier.Text & "'"
    Conn.Execute (SQL)
    MsgBox "Data Berhasil Diubah", vbInformation, "Berhasil"
    TampilData
End Sub

Sub CekSimpan()
Dim RsCek As Recordset

SQL = "Select * from Supplier where idSupplier='" & TXtIdSupplier.Text & "'"
Set RsCek = Conn.Execute(SQL)
If Not RsCek.EOF Then
    ubah
Else
    Simpan
End If
End Sub

Sub Hapus()
SQL = "DELETE from Supplier where idSupplier ='" & TXtIdSupplier.Text & "'"
    Conn.Execute (SQL)
    MsgBox "Berhasil Dihapus", vbInformation, "Berhasil"
    TampilData
End Sub

Sub GenerateNomor()
Dim RsNo As Recordset
Dim t As Integer
Dim No As String

SQL = "select idSupplier  from Supplier ORDER BY idSupplier DESC"
Set RsNo = Conn.Execute(SQL)
If RsNo.EOF = True Then
TXtIdSupplier.Text = "S" + "001"
Else
t = Val(Mid(RsNo("idSupplier"), 2, 3))
No = "S" + Format(Str(t + 1), "000")
TXtIdSupplier.Text = No
End If
End Sub


Sub Bersih()
TXtIdSupplier.Text = ""
TxtNamaSupplier.Text = ""
TxtAlamatSupplier.Text = ""
End Sub

Sub Pilih()
TXtIdSupplier.Text = rs(0)
TxtNamaSupplier.Text = rs(1)
TxtAlamatSupplier.Text = rs(2)
End Sub

Sub Pencarian()
Set rs = New Recordset
SQL = "Select * from Supplier where namaSupplier like '%" & TxtCari.Text & "%'"
rs.Open SQL, Conn, adOpenStatic, adLockReadOnly
Set Datagrid.DataSource = rs
TabelHeader
End Sub

Sub Hidup()
TxtNamaSupplier.Enabled = True
TxtAlamatSupplier.Enabled = True
TXtIdSupplier.Enabled = False
TxtNamaSupplier.SetFocus
End Sub

Sub Mati()
TxtNamaSupplier.Enabled = False
TxtAlamatSupplier.Enabled = False
TXtIdSupplier.Enabled = False
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



Private Sub TxtNamaSupplier_Change()
TxtNamaSupplier.MaxLength = 40
End Sub
