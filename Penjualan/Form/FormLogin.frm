VERSION 5.00
Begin VB.Form FormLogin 
   Caption         =   "Halaman Login"
   ClientHeight    =   3510
   ClientLeft      =   8625
   ClientTop       =   1980
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FormLogin.frx":0000
   ScaleHeight     =   3510
   ScaleWidth      =   5055
   Begin VB.CommandButton CmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      Picture         =   "FormLogin.frx":3A98C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton CmdLogin 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      Picture         =   "FormLogin.frx":3AD92
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox TxtPass 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1920
      Width           =   2415
   End
   Begin VB.TextBox TxtUsername 
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
      Left            =   2160
      TabIndex        =   2
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Login Area"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   6
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   5295
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FF8080&
      Height          =   1575
      Left            =   -120
      Top             =   960
      Width           =   5415
   End
End
Attribute VB_Name = "FormLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsCek As Recordset

Private Sub CmdClose_Click()
Unload Me
End Sub

Private Sub CmdLogin_Click()
If TxtUsername.Text = "" Or TxtPass.Text = "" Then
    MsgBox "Silahkan Lengkapi data Login", vbExclamation, "Peringatan"
Else
    MasukAdmin
End If
End Sub

Private Sub Form_Load()
Koneksi
End Sub

Sub MasukAdmin()
    SQL = "Select * from user where "
    SQL = SQL & "nama= '" & TxtUsername.Text & "' and "
    SQL = SQL & "Password= '" & TxtPass.Text & "' and "
    SQL = SQL & "Level= 'Admin'"
Set RsCek = Conn.Execute(SQL)
If Not RsCek.EOF Then
    FormMenuUtama.Show
    FormMenuUtama.StatusBar1.Panels(1) = "Status : Admin"
    FormMenuUtama.StatusBar1.Panels(2) = TxtUsername.Text
    FormMenuUtama.StatusBar1.Panels(5) = RsCek!idUser
    FormMenuUtama.Admin
    Unload Me
Else
    MasukKasir
End If
End Sub

Sub MasukKasir()
   SQL = "Select * from user where "
    SQL = SQL & "nama= '" & TxtUsername.Text & "' and "
    SQL = SQL & "Password= '" & TxtPass.Text & "' and "
    SQL = SQL & "Level= 'Kasir'"
Set RsCek = Conn.Execute(SQL)
If Not RsCek.EOF Then
    FormMenuUtama.Show
    FormMenuUtama.StatusBar1.Panels(1) = "Status : Kasir"
    FormMenuUtama.StatusBar1.Panels(2) = TxtUsername.Text
    FormMenuUtama.StatusBar1.Panels(5) = RsCek!idUser
    FormMenuUtama.Kasir
    Unload Me
Else
    MasukPemilik
End If
End Sub

Sub MasukPemilik()
   SQL = "Select * from User where "
    SQL = SQL & "nama= '" & TxtUsername.Text & "' and "
    SQL = SQL & "Password= '" & TxtPass.Text & "' and "
    SQL = SQL & "Level= 'Pemilik'"
Set RsCek = Conn.Execute(SQL)
If Not RsCek.EOF Then
    FormMenuUtama.Show
    FormMenuUtama.StatusBar1.Panels(1) = "Status : Pemilik"
    FormMenuUtama.StatusBar1.Panels(2) = TxtUsername.Text
    FormMenuUtama.StatusBar1.Panels(5) = RsCek!idUser
    FormMenuUtama.Pemilik
    Unload Me
Else
    MsgBox "Maaf Username atau Pasword Belum Terdaftar", vbCritical, "Peringatan"
    Bersih
End If
End Sub

Sub Bersih()
TxtUsername.Text = ""
TxtPass.Text = ""
End Sub

Private Sub TxtPass_Change()
TxtPass.MaxLength = 6
End Sub

Private Sub TxtUsername_Change()
TxtUsername.MaxLength = 40
End Sub
