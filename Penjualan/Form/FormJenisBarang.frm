VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FormJenisBarang 
   Caption         =   "Jenis Barang"
   ClientHeight    =   6495
   ClientLeft      =   8625
   ClientTop       =   1725
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FormJenisBarang.frx":0000
   ScaleHeight     =   6495
   ScaleWidth      =   6060
   Begin VB.CommandButton CmdTmbah 
      Caption         =   "Tambah"
      Height          =   495
      Left            =   480
      Picture         =   "FormJenisBarang.frx":548F5
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "Simpan"
      Height          =   495
      Left            =   1560
      Picture         =   "FormJenisBarang.frx":54B4B
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "Edit"
      Height          =   495
      Left            =   2640
      Picture         =   "FormJenisBarang.frx":54DB6
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton CmdHapus 
      Caption         =   "Hapus"
      Height          =   495
      Left            =   3720
      Picture         =   "FormJenisBarang.frx":54FDF
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton Cmdbatal 
      Caption         =   "Batal"
      Height          =   495
      Left            =   4800
      Picture         =   "FormJenisBarang.frx":5512D
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox TxtCari 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   2880
      Width           =   3015
   End
   Begin VB.TextBox TxtNamaJenis 
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
      TabIndex        =   3
      Top             =   1320
      Width           =   2895
   End
   Begin VB.TextBox txtidJenis 
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
      TabIndex        =   2
      Top             =   720
      Width           =   2895
   End
   Begin MSDataGridLib.DataGrid DataGrid 
      Height          =   2655
      Left            =   480
      TabIndex        =   6
      Top             =   3600
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   4683
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   19
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
         Size            =   9.75
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
      Left            =   960
      TabIndex        =   4
      Top             =   2880
      Width           =   1335
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
      Left            =   720
      TabIndex        =   1
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Id Jenis Barang"
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
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C0C000&
      FillStyle       =   0  'Solid
      Height          =   1695
      Left            =   480
      Top             =   360
      Width           =   5415
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00C0C000&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   480
      Top             =   2760
      Width           =   5415
   End
End
Attribute VB_Name = "FormJenisBarang"
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

Private Sub TxtCari_Change()
Pencarian
End Sub

Sub TabelHeader()
With Datagrid
    .RowHeight = 280
    .HeadFont = "Arial"
    .HeadLines = 2
    .Columns(0).Caption = "ID JENIS BARANG"
    .Columns(1).Caption = "NAMA JENIS BARANG"
    .Columns(0).Width = 1500
    .Columns(1).Width = 3500
End With

End Sub

Sub TampilData()
Set rs = New Recordset
SQL = "Select * from JenisBarang"
rs.Open SQL, Conn, adOpenStatic, adLockReadOnly
Set Datagrid.DataSource = rs
TabelHeader
End Sub

Sub Simpan()
        SQL = "INSERT INTO JenisBarang VALUES("
        SQL = SQL & "'" & txtidJenis.Text & "',"
        SQL = SQL & "'" & TxtNamaJenis.Text & "')"
        Conn.Execute (SQL)
        MsgBox "Berhasil Jenis Barang Simpan", vbInformation, "Informasi"
        TampilData
End Sub

Sub ubah()
    SQL = "UPDATE JenisBarang set "
    SQL = SQL & "NamaJenisBarang='" & TxtNamaJenis.Text & "'"
    SQL = SQL & "Where IdJenisBarang='" & txtidJenis.Text & "'"
    Conn.Execute (SQL)
    MsgBox "Data Berhasil Diubah", vbInformation, "Berhasil"
    TampilData
End Sub

Sub CekSimpan()
Dim RsCek As Recordset

SQL = "Select * from JenisBarang where idJenisBarang='" & txtidJenis.Text & "'"
Set RsCek = Conn.Execute(SQL)

If Not RsCek.EOF Then
    ubah
Else
    Simpan
End If
End Sub

Sub Hapus()
SQL = "DELETE from JenisBarang where idjenisBarang ='" & txtidJenis.Text & "'"
    Conn.Execute (SQL)
    MsgBox "Berhasil Dihapus", vbInformation, "Berhasil"
    TampilData
End Sub

Sub GenerateNomor()
Dim RsNo As Recordset
Dim t As Integer
Dim No As String

SQL = "select idJenisBarang  from JenisBarang ORDER BY idJenisBarang DESC"
Set RsNo = Conn.Execute(SQL)
If RsNo.EOF = True Then
txtidJenis.Text = "J" + "001"
Else
t = Val(Mid(RsNo("idJenisBarang"), 2, 3))
No = "J" + Format(Str(t + 1), "000")
txtidJenis.Text = No
End If
End Sub


Sub Bersih()
txtidJenis.Text = ""
TxtNamaJenis.Text = ""
End Sub

Sub Pilih()
txtidJenis.Text = rs(0)
TxtNamaJenis.Text = rs(1)
End Sub

Sub Pencarian()
Set rs = New Recordset
SQL = "Select * from JenisBarang where namaJenisBarang like '%" & TxtCari.Text & "%'"
rs.Open SQL, Conn, adOpenStatic, adLockReadOnly
Set Datagrid.DataSource = rs
TabelHeader
End Sub

Sub Hidup()
TxtNamaJenis.Enabled = True
TxtNamaJenis.SetFocus
End Sub

Sub Mati()
TxtNamaJenis.Enabled = False
txtidJenis.Enabled = False
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

Private Sub TxtNamaJenis_Change()
TxtNamaJenis.MaxLength = 40
End Sub
