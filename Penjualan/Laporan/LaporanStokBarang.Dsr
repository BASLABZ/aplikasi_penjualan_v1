VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} LaporanStokBarang 
   Caption         =   "ActiveReport1"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13110
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   23125
   _ExtentY        =   13705
   SectionData     =   "LaporanStokBarang.dsx":0000
End
Attribute VB_Name = "LaporanStokBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_BeforePrint()
Field1.Height = Detail.Height
Field2.Height = Detail.Height
Field3.Height = Detail.Height
Field4.Height = Detail.Height
Field5.Height = Detail.Height
End Sub

Private Sub Detail_Format()
With DataControl1.Recordset
        If Not .EOF Then
            Field1.Text = .Fields("idBarang").Value
            Field2.Text = .Fields("NamaBarang").Value
            Field3.Text = FormatCurrency(.Fields("HargaBeli").Value, 2)
            Field4.Text = FormatCurrency(.Fields("HargaJual").Value, 2)
            Field5.Text = .Fields("Stok").Value
            End If
    End With
End Sub

Private Sub PageFooter_Format()
Label16.Caption = "Yogyakarta, " + Format(Now, "DD MMMM YYYY ")
Label17.Caption = FormMenuUtama.StatusBar1.Panels(2)
End Sub
