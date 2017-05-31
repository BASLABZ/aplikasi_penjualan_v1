Attribute VB_Name = "ModulKonesi"
Public Conn As New Connection
Public strConn As String
Public rs As ADODB.Recordset
Public SQL As String

Sub Koneksi()

strConn = "DSN=Penjualan"
Conn.CursorLocation = adUseClient

If Conn.State = adStateClosed Then
    Conn.Open strConn
    If Conn.State = adStateClosed Then
        MsgBox "Koneksi ke database gagal !", vbOKOnly + vbCritical, "Kesalahan"
        End
    End If
End If
End Sub

