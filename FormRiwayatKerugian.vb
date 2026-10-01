Imports System.Data.OleDb

Public Class FormRiwayatKerugian

    Private Sub FormRiwayatKerugian_Load(
        sender As Object,
        e As EventArgs
    ) Handles MyBase.Load

        TampilkanDataKerugian()

    End Sub


    Private Sub TampilkanDataKerugian()

        Try

            Koneksi()

            Dim query As String =
                "SELECT " &
                "[ID_Kerugian], " &
                "[No_ID], " &
                "[Total_Biaya_Keugian], " &
                "[ID_Produk], " &
                "[ID_NG] " &
                "FROM [Data_Pemrosesan_Kerugian]"

            Using cmdKerugian As New OleDbCommand(
                query,
                CNN
            )

                Using adapter As New OleDbDataAdapter(
                    cmdKerugian
                )

                    Dim dtData As New DataTable()

                    adapter.Fill(dtData)

                    dgvKerugian.DataSource = dtData

                End Using

            End Using

            CNN.Close()

        Catch ex As Exception

            If CNN IsNot Nothing AndAlso
               CNN.State = ConnectionState.Open Then

                CNN.Close()

            End If

            MessageBox.Show(
                "Gagal menampilkan data kerugian." &
                vbCrLf &
                ex.Message,
                "Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            )

        End Try

    End Sub
    Private Sub btnKembaliKerugian_Click(
    sender As Object,
    e As EventArgs
) Handles btnKembaliKerugian.Click

        FormKerugian.Show()
        Me.Hide()

    End Sub
    Private Sub btnCariKerugian_Click(
        sender As Object,
        e As EventArgs
    ) Handles btnCariKerugian.Click

        If txtCariKerugian.Text.Trim() = "" Then

            MessageBox.Show(
                "Masukkan ID Kerugian terlebih dahulu.",
                "Peringatan",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            )

            txtCariKerugian.Focus()

            Exit Sub

        End If


        Try

            Koneksi()

            Dim query As String =
                "SELECT " &
                "[ID_Kerugian], " &
                "[No_ID], " &
                "[Total_Biaya_Keugian], " &
                "[ID_Produk], " &
                "[ID_NG] " &
                "FROM [Data_Pemrosesan_Kerugian] " &
                "WHERE [ID_Kerugian] = ?"

            Using cmdCari As New OleDbCommand(
                query,
                CNN
            )

                cmdCari.Parameters.AddWithValue(
                    "@ID_Kerugian",
                    txtCariKerugian.Text.Trim()
                )

                Using adapter As New OleDbDataAdapter(
                    cmdCari
                )

                    Dim dtHasil As New DataTable()

                    adapter.Fill(dtHasil)

                    If dtHasil.Rows.Count > 0 Then

                        dgvKerugian.DataSource = dtHasil

                    Else

                        dgvKerugian.DataSource = Nothing

                        MessageBox.Show(
                            "ID Kerugian tidak ditemukan.",
                            "Informasi",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information
                        )

                    End If

                End Using

            End Using

            CNN.Close()

        Catch ex As Exception

            If CNN IsNot Nothing AndAlso
               CNN.State = ConnectionState.Open Then

                CNN.Close()

            End If

            MessageBox.Show(
                "Gagal mencari data kerugian." &
                vbCrLf &
                ex.Message,
                "Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            )

        End Try

    End Sub
End Class