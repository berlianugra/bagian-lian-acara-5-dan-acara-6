Imports System.Data.OleDb

Module Module1

    Public A As String
    Public CNN As OleDbConnection
    Public da As OleDbDataAdapter
    Public ds As DataSet
    Public dt As DataTable
    Public tx As Label
    Public OLECMD As OleDbCommand
    Public cmd As OleDbCommand
    Public cmd1 As OleDbCommand
    Public Rd As OleDbDataReader

    Public CONN As String =
        "Provider=Microsoft.JET.OLEDB.4.0;Data Source=" & Application.StartupPath & "\Database Acara 3.mdb;"

    Public Sub Koneksi()

        CNN = New OleDbConnection(CONN)

        If CNN.State = ConnectionState.Closed Then
            CNN.Open()
        End If

    End Sub
    Public Class ProdukItem

        Public ID As String
        Public Nama As String

        Public Sub New(idProduk As String, namaProduk As String)

            ID = idProduk
            Nama = namaProduk

        End Sub

        Public Overrides Function ToString() As String

            Return Nama

        End Function

    End Class
End Module