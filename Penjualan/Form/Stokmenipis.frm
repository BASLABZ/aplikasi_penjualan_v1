VERSION 5.00
Begin VB.Form Stokmenipis 
   Caption         =   "Setting Stok"
   ClientHeight    =   2685
   ClientLeft      =   7605
   ClientTop       =   5160
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Stokmenipis.frx":0000
   ScaleHeight     =   2685
   ScaleWidth      =   4005
   Begin VB.CommandButton CmdUbah 
      Caption         =   "Ubah"
      Height          =   495
      Left            =   2760
      Picture         =   "Stokmenipis.frx":548F5
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox txtStok 
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
      Left            =   2160
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label LabelMinimal 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Stok Minimal"
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
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Stok Minimal"
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
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C0C000&
      FillStyle       =   0  'Solid
      Height          =   1695
      Left            =   120
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "Stokmenipis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdUbah_Click()
FormMenuUtama.StatusBar1.Panels(5) = TXtStok.Text
TXtStok.Text = ""
MsgBox "Data setting berhasil di ubah", vbInformation, "Informasi"
Lihat
End Sub

Sub Lihat()
LabelMinimal.Caption = FormMenuUtama.StatusBar1.Panels(5)
End Sub

Private Sub Form_Load()
Lihat
End Sub
