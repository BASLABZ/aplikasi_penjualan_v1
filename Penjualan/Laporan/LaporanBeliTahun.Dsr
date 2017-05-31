VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} LaporanBeliTahun 
   Caption         =   "ActiveReport1"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20370
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   35930
   _ExtentY        =   19420
   SectionData     =   "LaporanBeliTahun.dsx":0000
End
Attribute VB_Name = "LaporanBeliTahun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_BeforePrint()
Field1.Height = Detail.Height
Field2.Height = Detail.Height
End Sub

Private Sub Detail_Format()
With DataControl1.Recordset
        If Not .EOF Then
            Field1.Text = .Fields("month(TanggalBeli)").Value
            Field2.Text = FormatCurrency(.Fields("tot").Value, 2)
            End If
    End With
End Sub

Private Sub PageFooter_Format()
Label16.Caption = "Yogyakarta, " + Format(Now, "DD MMMM YYYY ")
Label17.Caption = FormMenuUtama.StatusBar1.Panels(2)
End Sub

