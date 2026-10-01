Imports System.Data.OleDb

Public Class FormRiwayatUsulan

    Private Sub FormRiwayatUsulan_Load(
        sender As Object,
        e As EventArgs
    ) Handles MyBase.Load

        TampilkanDataUsulan()

    End Sub


    Private Sub TampilkanDataUsulan()

        Try

            Koneksi()

            Dim query As String =
                "SELECT " &
                "[ID_Usulan], " &
                "[ID_Kualitas], " &
                "[No_ID], " &
                "[ID_Kerugian], " &
                "[Usulan_Perbaikan], " &
                "[Tanggal_Usulan] " &
                "FROM [Data_Usulan] " &
                "ORDER BY [Tanggal_Usulan] DESC"

            Using cmd As New OleDbCommand(
                query,
                CNN
            )

                Using adapter As New OleDbDataAdapter(cmd)

                    Dim dt As New DataTable()

                    adapter.Fill(dt)

                    dgvUsulan.DataSource = dt

                End Using

            End Using

            CNN.Close()

        Catch ex As Exception

            If CNN IsNot Nothing AndAlso
               CNN.State = ConnectionState.Open Then

                CNN.Close()

            End If

            MessageBox.Show(
                "Gagal menampilkan data usulan." &
                vbCrLf &
                ex.Message,
                "Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            )

        End Try

    End Sub

End Class