VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} NotaPenjualan 
   Caption         =   "Nota Jual"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11025
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   19447
   _ExtentY        =   14182
   SectionData     =   "NotaPenjualan.dsx":0000
End
Attribute VB_Name = "NotaPenjualan"
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
            Field3.Text = FormatCurrency(.Fields("HargaJual").Value, 2)
            Field4.Text = .Fields("Jumlah").Value
            Field5.Text = FormatCurrency(.Fields("Subtotal").Value, 2)
            End If
    End With
End Sub


