VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form BrowsePembelian 
   Caption         =   "Browse Pembelian"
   ClientHeight    =   3780
   ClientLeft      =   9660
   ClientTop       =   1980
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "BrowsePembelian.frx":0000
   ScaleHeight     =   3780
   ScaleWidth      =   6420
   Begin VB.TextBox TxtCari 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6000
   End
   Begin MSDataGridLib.DataGrid Datagrid 
      Height          =   2415
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   4260
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16761024
      HeadLines       =   1
      RowHeight       =   18
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
         Name            =   "Times New Roman"
         Size            =   9
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
            LCID            =   1033
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
            LCID            =   1033
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "* Klik 2 Kali Untuk Memilih Data Pembelian"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   3240
      Width           =   4815
   End
End
Attribute VB_Name = "BrowsePembelian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Datagrid_DblClick()
Pilih
FormReturPembelian.tampilBarang
End Sub

Private Sub Form_Load()
Koneksi
TampilData
End Sub
Sub TampilData()
Set rs = New Recordset
strsql = "Select * from Pembelian"
rs.Open strsql, Conn, adOpenStatic, adLockReadOnly
Set Datagrid.DataSource = rs
'AturTabel
End Sub
Sub AturTabel()
With Datagrid
    .RowHeight = 280
    .HeadFont = "Arial"
    .HeadLines = 2
    .Columns(0).Caption = "ID BARANG"
    .Columns(1).Caption = "ID JENIS BARANG"
    .Columns(2).Caption = "NAMA BARANG"
    .Columns(3).Caption = "HARGA BELI"
    .Columns(4).Caption = "HARGA JUAL"
    .Columns(5).Caption = "STOK"
    .Columns(0).Width = 1500
    .Columns(1).Width = 1500
    .Columns(2).Width = 1500
    .Columns(3).Width = 1500
    .Columns(4).Width = 1500
    .Columns(5).Width = 1500
End With
End Sub
Private Sub TxtCari_Change()
'Set rs = New Recordset
'strsql = "Select * from BARANG where NamaBarang like '%" & TxtCari.Text & "%'"
'rs.Open strsql, Conn, adOpenStatic, adLockReadOnly
'Set Datagrid.DataSource = rs
'AturTabel
End Sub

Sub Pilih()
FormReturPembelian.TxtIdNota.Text = rs(0)
Unload Me
End Sub


