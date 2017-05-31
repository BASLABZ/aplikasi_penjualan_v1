VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FormUser 
   Caption         =   "User"
   ClientHeight    =   6390
   ClientLeft      =   5505
   ClientTop       =   3270
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FormUser.frx":0000
   ScaleHeight     =   6390
   ScaleWidth      =   9630
   Begin VB.TextBox TxtUser 
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
      Left            =   2520
      TabIndex        =   15
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton CmdTmbah 
      Caption         =   "Tambah"
      Height          =   495
      Left            =   1920
      Picture         =   "FormUser.frx":548F5
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "Simpan"
      Height          =   495
      Left            =   3000
      Picture         =   "FormUser.frx":54B4B
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "Edit"
      Height          =   495
      Left            =   4080
      Picture         =   "FormUser.frx":54DB6
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton CmdHapus 
      Caption         =   "Hapus"
      Height          =   495
      Left            =   5160
      Picture         =   "FormUser.frx":54FDF
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Cmdbatal 
      Caption         =   "Batal"
      Height          =   495
      Left            =   6240
      Picture         =   "FormUser.frx":5512D
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2040
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid 
      Height          =   2655
      Left            =   360
      TabIndex        =   9
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
   Begin VB.TextBox TxtCari 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   2760
      Width           =   3255
   End
   Begin VB.ComboBox ComboLevel 
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
      Left            =   6960
      TabIndex        =   6
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox TxtPassword 
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
      TabIndex        =   5
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox TxtNama 
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
      Left            =   2520
      TabIndex        =   4
      Top             =   1200
      Width           =   1935
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
      Left            =   480
      TabIndex        =   7
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Level"
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
      Left            =   5160
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Left            =   5160
      TabIndex        =   2
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama User"
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
      Left            =   600
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Id User"
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
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C0C000&
      FillStyle       =   0  'Solid
      Height          =   1695
      Left            =   360
      Top             =   240
      Width           =   9135
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00C0C000&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   360
      Top             =   2640
      Width           =   9135
   End
End
Attribute VB_Name = "FormUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmdbatal_Click()
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
    .Columns(0).Caption = "ID USER"
    .Columns(1).Caption = "NAMA "
    .Columns(2).Caption = "PASSWORD"
    .Columns(3).Caption = "PASSWORD"
    .Columns(0).Width = 1500
    .Columns(1).Width = 2000
    .Columns(2).Width = 2500
    .Columns(3).Width = 2500
End With
End Sub

Sub TampilData()
Set rs = New Recordset
SQL = "Select * from user"
rs.Open SQL, Conn, adOpenStatic, adLockReadOnly
Set Datagrid.DataSource = rs
TabelHeader
End Sub

Sub Simpan()
        SQL = "INSERT INTO User VALUES("
        SQL = SQL & "'" & TxtUser.Text & "',"
        SQL = SQL & "'" & TxtNama.Text & "',"
        SQL = SQL & "'" & TxtPassword.Text & "',"
        SQL = SQL & "'" & ComboLevel.Text & "')"
        Conn.Execute (SQL)
        MsgBox "Berhasil User Simpan", vbInformation, "Informasi"
        TampilData
End Sub

Sub ubah()
    SQL = "UPDATE User set "
    SQL = SQL & "Nama='" & TxtNama.Text & "',"
    SQL = SQL & "Level='" & ComboLevel.Text & "',"
    SQL = SQL & "Password='" & TxtPassword.Text & "'"
    SQL = SQL & "Where idUser='" & TxtUser.Text & "'"
    Conn.Execute (SQL)
    MsgBox "Data Berhasil Diubah", vbInformation, "Berhasil"
    TampilData
End Sub

Sub CekSimpan()
Dim RsCek As Recordset

SQL = "Select * from User where idUser='" & TxtUser.Text & "' or nama='" & TxtNama.Text & "'"
Set RsCek = Conn.Execute(SQL)
If Not RsCek.EOF Then
    ubah
Else
    Simpan
End If
End Sub

Sub Hapus()
SQL = "DELETE from User where idUser ='" & TxtUser.Text & "'"
    Conn.Execute (SQL)
    MsgBox "Berhasil Dihapus", vbInformation, "Berhasil"
    TampilData
End Sub

Sub GenerateNomor()
Dim RsNo As Recordset
Dim t As Integer
Dim No As String

SQL = "select IdUser  from User ORDER BY idUser DESC"
Set RsNo = Conn.Execute(SQL)
If RsNo.EOF = True Then
TxtUser.Text = "U" + "001"
Else
t = Val(Mid(RsNo("IdUser"), 2, 3))
No = "U" + Format(Str(t + 1), "000")
TxtUser.Text = No
End If
End Sub


Sub Bersih()
TxtUser.Text = ""
TxtNama.Text = ""
TxtPassword.Text = ""
ComboLevel.Text = ""
End Sub

Sub Pilih()
TxtUser.Text = rs(0)
TxtNama.Text = rs(1)
TxtPassword.Text = rs(2)
ComboLevel.Text = rs(3)
End Sub

Sub Pencarian()
Set rs = New Recordset
SQL = "Select * from user where nama like '%" & TxtCari.Text & "%'"
rs.Open SQL, Conn, adOpenStatic, adLockReadOnly
Set Datagrid.DataSource = rs
TabelHeader
End Sub

Sub TampilCombo()
ComboLevel.AddItem "Admin"
ComboLevel.AddItem "Kasir"
ComboLevel.AddItem "Pemilik"
End Sub

Sub Hidup()
TxtNama.Enabled = True
TxtPassword.Enabled = True
TxtUser.Enabled = False
ComboLevel.Enabled = True
TxtNama.SetFocus
End Sub

Sub Mati()
TxtNama.Enabled = False
TxtPassword.Enabled = False
TxtUser.Enabled = False
ComboLevel.Enabled = False
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

Private Sub TxtNama_Change()
TxtNama.MaxLength = 40
End Sub

Private Sub TxtPassword_Change()
TxtPassword.MaxLength = 6
End Sub
