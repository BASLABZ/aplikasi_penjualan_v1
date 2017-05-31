VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FormReturPembelian 
   Caption         =   "Retur Pembelian"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FormReturPembelian.frx":0000
   ScaleHeight     =   6855
   ScaleWidth      =   9855
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdBatal 
      Caption         =   "Batal"
      Height          =   495
      Left            =   7320
      Picture         =   "FormReturPembelian.frx":548F5
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton CmdTambah 
      Caption         =   "Tambah"
      Height          =   495
      Left            =   8400
      Picture         =   "FormReturPembelian.frx":54B60
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "Simpan"
      Height          =   495
      Left            =   8520
      Picture         =   "FormReturPembelian.frx":54DC8
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton CmdHapus 
      Caption         =   "Hapus"
      Height          =   495
      Left            =   7200
      Picture         =   "FormReturPembelian.frx":55033
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox TxtBarangRetur 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   7080
      TabIndex        =   11
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox TxtJumlahBarang 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   2280
      Width           =   1575
   End
   Begin VB.ComboBox CmbBarang 
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
      Left            =   2640
      TabIndex        =   7
      Top             =   1800
      Width           =   4335
   End
   Begin VB.TextBox TxtIdNota 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1320
      Width           =   1815
   End
   Begin MSComctlLib.ListView ListViewBarang 
      Height          =   2535
      Left            =   120
      TabIndex        =   12
      Top             =   3480
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   4471
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
         Text            =   "Kode Detail"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nama Barang"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Jumlah Pembelian"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Jumlah Retur"
         Object.Width           =   2999
      EndProperty
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Barang Retur"
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
      Left            =   5400
      TabIndex        =   9
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Pembelian"
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
      TabIndex        =   8
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Barang"
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
      TabIndex        =   6
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Nota Pembelian"
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
      TabIndex        =   5
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label LabelTanggal 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal retur"
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
      TabIndex        =   3
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Retur Pembelian"
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
      TabIndex        =   2
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
      TabIndex        =   1
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Id Retur"
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
      TabIndex        =   0
      Top             =   360
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
      Height          =   1815
      Left            =   120
      Top             =   960
      Width           =   9495
   End
End
Attribute VB_Name = "FormReturPembelian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CmbBarang_Click()
TxtJumlahBarang.Text = ""
TampilJumlah
End Sub

Private Sub CmdBatal_Click()
Bersih
End Sub

Private Sub CmdSimpan_Click()
SimpanRetur
SimpanDetail
Bersih
End Sub

Private Sub CmdTambah_Click()
Dim status As Boolean
    If CmbBarang.Text = "" Or TxtBarangRetur.Text = "" Then
    MsgBox "Data Barang Belum Lengkap", vbCritical, "Peringatan"
    Else
    status = False
    For i = 1 To ListViewBarang.ListItems.Count
        With ListViewBarang.ListItems.Item(i)
        If Mid(CmbBarang.Text, 1, 4) = .Text Then
            status = True
        End If
        End With
    Next
    If status = False Then
            TambahKeList
            ClearBarang
            Else
            HapusBarang
            TambahKeList
            ClearBarang
        End If
    End If
End Sub

Private Sub Form_Load()
Koneksi
LabelTanggal.Caption = Format(Now, "DD-MM-YYYY")
GenerateNomorRetur
End Sub

Sub GenerateNomorRetur()
Dim RsNo As Recordset
Dim t As Integer
Dim No As String

SQL = "select max(right(idreturPembelian,4))  from ReturPembelian"
Set RsNo = Conn.Execute(SQL)
If RsNo.EOF = True Then
NoId.Caption = "RP" + "0001"
Else
t = Val(Mid(RsNo(0), 3, 4))
No = "RP" + Format(Str(t + 1), "0000")
NoId.Caption = No
End If
End Sub


Private Sub TxtBarangRetur_KeyPress(KeyAscii As Integer)
If InStr("0123456789", Chr(KeyAscii)) = 0 Then
If KeyAscii <> vbKeyBack Then
KeyAscii = 0
End If
End If
End Sub

Private Sub TxtIdNota_Click()
CmbBarang.Clear
TxtJumlahBarang.Text = ""
TxtBarangRetur.Text = ""
BrowsePembelian.Show
End Sub

Sub tampilBarang()
Dim rsBarang As Recordset
Set rsBarang = New Recordset
Dim sqlBarang As String

sqlBarang = "SELECT dp.Idbarang,dp.IdDetailPembelian,b.NamaBarang FROM DetailPembelian dp"
sqlBarang = sqlBarang & " Join Barang b on b.idBarang=dp.idBarang"
sqlBarang = sqlBarang & " where idPembelian='" & TxtIdNota.Text & "'"

rsBarang.Open sqlBarang, Conn, adOpenStatic, adLockReadOnly

Do While Not rsBarang.EOF
    CmbBarang.AddItem rsBarang!idBarang + "-" + rsBarang!idDetailPembelian + "-" + rsBarang!namaBarang
    rsBarang.MoveNext
Loop
End Sub

Sub TampilJumlah()
Dim RsCek As Recordset
SQL = "Select * from detailPembelian where idDetailPembelian='" & Mid(CmbBarang, 6, 6) & "'"
Set RsCek = Conn.Execute(SQL)
TxtJumlahBarang.Text = RsCek!jumlah
End Sub

Sub TambahKeList()
With ListViewBarang.ListItems.Add
    .Text = Mid(CmbBarang.Text, 1, 4)
    .SubItems(1) = Mid(CmbBarang.Text, 6, 6)
    .SubItems(2) = Mid(CmbBarang.Text, 13, 25)
    .SubItems(3) = TxtJumlahBarang.Text
    .SubItems(4) = TxtBarangRetur.Text
End With
End Sub

Sub ClearBarang()
CmbBarang.Text = ""
TxtJumlahBarang.Text = ""
TxtBarangRetur.Text = ""
End Sub
Sub HapusBarang()
If ListViewBarang.ListItems.Count = 0 Then
    MsgBox "Data Barang Belum Dipilih ", vbCritical, "Peringatan"
Else
    ListViewBarang.ListItems.Remove (ListViewBarang.SelectedItem.Index)
End If
End Sub

Sub SimpanRetur()
strsql = "INSERT INTO ReturPembelian VALUES("
        strsql = strsql & "'" & NoId.Caption & "',"
        strsql = strsql & "'" & TxtIdNota.Text & "',"
        strsql = strsql & "'User',"
        strsql = strsql & "" & Format(Now, "YYYY-MM-DD") & ")"
        Conn.Execute (strsql)
        MsgBox "Retur Pemebelian Berhasil Simpan", vbInformation, "Berhasil"
End Sub

Sub SimpanDetail()
Dim RsCek As Recordset
Dim SQLCek As String
Dim Total As Double


For i = 1 To ListViewBarang.ListItems.Count
With ListViewBarang.ListItems.Item(i)
        strsql = "INSERT INTO DetailreturPembelian values ('" & NoId.Caption & "','" & .SubItems(1) & "'," & .SubItems(4) & ")"
        Conn.Execute (strsql)

        strsql = "update DetailPembelian set jumlah=jumlah -" & Val(.SubItems(4)) & " where idDetailPembelian='" & .SubItems(1) & "'"
        Conn.Execute (strsql)
        
        SQLCek = "Select sum(jumlah*hargaBeli) from DetailPembelian where IdPembelian='" & TxtIdNota.Text & "'"
        Set RsCek = Conn.Execute(SQLCek)
        Total = RsCek(0)
        
        SQLCek = "Update pembelian set total='" & Total & "' where idPembelian='" & TxtIdNota.Text & "'"
        Conn.Execute (SQLCek)
End With
Next
End Sub


Sub Bersih()
ListViewBarang.ListItems.Clear
GenerateNomorRetur
TxtIdNota.Text = ""
End Sub
