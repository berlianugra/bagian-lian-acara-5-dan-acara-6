Imports System.Data.OleDb

Public Class FormInputUsulan

    Private usulanSudahDitambahkan As Boolean = False

    Private Sub Form5_Load(
        sender As Object,
        e As EventArgs
    ) Handles MyBase.Load

        dtpTanggal.Value = DateTime.Now

        txtTotalProduksi.ReadOnly = True
        txtJumlahNG.ReadOnly = True

        txtTotalProduksi.BackColor = Color.LightGray
        txtJumlahNG.BackColor = Color.LightGray

        LoadProduk()
        LoadJenisNG()

    End Sub


    Private Sub LoadProduk()

        Try

            Koneksi()

            cmbNamaProduk.Items.Clear()

            Dim query As String =
                "SELECT [ID_Produk], [Nama_Produk] " &
                "FROM [Produk] " &
                "ORDER BY [Nama_Produk]"

            Using cmdProduk As New OleDbCommand(
                query,
                CNN
            )

                Using rdProduk As OleDbDataReader =
                    cmdProduk.ExecuteReader()

                    While rdProduk.Read()

                        cmbNamaProduk.Items.Add(
                            New ProdukItem(
                                rdProduk("ID_Produk").ToString(),
                                rdProduk("Nama_Produk").ToString()
                            )
                        )

                    End While

                End Using

            End Using

            CNN.Close()

        Catch ex As Exception

            If CNN IsNot Nothing AndAlso
               CNN.State = ConnectionState.Open Then

                CNN.Close()

            End If

            MessageBox.Show(
                "Gagal mengambil data produk." &
                vbCrLf &
                ex.Message,
                "Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            )

        End Try

    End Sub


    Private Sub LoadJenisNG()

        Try

            Koneksi()

            cmbJenisNG.Items.Clear()

            Dim query As String =
                "SELECT DISTINCT [Jenis_NG] " &
                "FROM [Data_Produk_NG] " &
                "WHERE [Jenis_NG] IS NOT NULL " &
                "ORDER BY [Jenis_NG]"

            Using cmdJenis As New OleDbCommand(
                query,
                CNN
            )

                Using rdJenis As OleDbDataReader =
                    cmdJenis.ExecuteReader()

                    While rdJenis.Read()

                        cmbJenisNG.Items.Add(
                            rdJenis("Jenis_NG").ToString()
                        )

                    End While

                End Using

            End Using

            CNN.Close()

        Catch ex As Exception

            If CNN IsNot Nothing AndAlso
               CNN.State = ConnectionState.Open Then

                CNN.Close()

            End If

            MessageBox.Show(
                "Gagal mengambil jenis NG." &
                vbCrLf &
                ex.Message,
                "Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            )

        End Try

    End Sub

    Private Sub cmbNamaProduk_SelectedIndexChanged(
        sender As Object,
        e As EventArgs
    ) Handles cmbNamaProduk.SelectedIndexChanged

        If cmbNamaProduk.SelectedIndex = -1 Then
            Exit Sub
        End If

        Dim produk As ProdukItem =
    TryCast(cmbNamaProduk.SelectedItem, ProdukItem)

        If produk Is Nothing Then
            MessageBox.Show(
        "Data produk yang dipilih tidak valid.",
        "Peringatan",
        MessageBoxButtons.OK,
        MessageBoxIcon.Warning
    )
            Exit Sub
        End If

        AmbilTotalProduksi(produk.ID)

        txtJumlahNG.Clear()

        usulanSudahDitambahkan = False

    End Sub

    Private Sub AmbilTotalProduksi(
        idProduk As String
    )

        Try

            Koneksi()

            Dim query As String =
                "SELECT TOP 1 " &
                "[Total_Produksi], " &
                "[Tanggal_Produksi] " &
                "FROM [Data_Pengelolaan_Total_Produksi] " &
                "WHERE [ID_Produk] = ? " &
                "ORDER BY [Tanggal_Produksi] DESC"

            Using cmdProduksi As New OleDbCommand(
                query,
                CNN
            )

                cmdProduksi.Parameters.AddWithValue(
                    "@ID_Produk",
                    idProduk
                )

                Using rdProduksi As OleDbDataReader =
                    cmdProduksi.ExecuteReader()

                    If rdProduksi.Read() Then

                        txtTotalProduksi.Text =
                            rdProduksi("Total_Produksi").ToString()

                    Else

                        txtTotalProduksi.Clear()

                        MessageBox.Show(
                            "Data total produksi untuk produk tersebut tidak ditemukan.",
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
                "Gagal mengambil data total produksi." &
                vbCrLf &
                ex.Message,
                "Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            )

        End Try

    End Sub


    Private Sub cmbJenisNG_SelectedIndexChanged(
        sender As Object,
        e As EventArgs
    ) Handles cmbJenisNG.SelectedIndexChanged

        If cmbNamaProduk.SelectedIndex = -1 Then
            Exit Sub
        End If

        If cmbJenisNG.SelectedIndex = -1 Then
            Exit Sub
        End If

        Dim produk As ProdukItem =
            CType(
                cmbNamaProduk.SelectedItem,
                ProdukItem
            )

        AmbilJumlahNG(
            produk.ID,
            cmbJenisNG.Text
        )

        usulanSudahDitambahkan = False

    End Sub


    Private Sub AmbilJumlahNG(
        idProduk As String,
        jenisNG As String
    )

        Try

            Koneksi()


            Dim queryTotal As String =
                "SELECT TOP 1 [ID_Total_Produk] " &
                "FROM [Data_Pengelolaan_Total_Produksi] " &
                "WHERE [ID_Produk] = ? " &
                "ORDER BY [Tanggal_Produksi] DESC"

            Dim idTotalProduk As Object = Nothing

            Using cmdTotal As New OleDbCommand(
                queryTotal,
                CNN
            )

                cmdTotal.Parameters.AddWithValue(
                    "@ID_Produk",
                    idProduk
                )

                idTotalProduk =
                    cmdTotal.ExecuteScalar()

            End Using


            If idTotalProduk Is Nothing OrElse
               IsDBNull(idTotalProduk) Then

                txtJumlahNG.Clear()

                CNN.Close()

                Exit Sub

            End If


            Dim queryNG As String =
                "SELECT TOP 1 [Jumlah_NG] " &
                "FROM [Data_Produk_NG] " &
                "WHERE [ID_Total_Produk] = ? " &
                "AND [Jenis_NG] = ?"

            Dim jumlahNG As Object = Nothing

            Using cmdNG As New OleDbCommand(
                queryNG,
                CNN
            )

                cmdNG.Parameters.AddWithValue(
                    "@ID_Total_Produk",
                    idTotalProduk
                )

                cmdNG.Parameters.AddWithValue(
                    "@Jenis_NG",
                    jenisNG
                )

                jumlahNG =
                    cmdNG.ExecuteScalar()

            End Using


            If jumlahNG IsNot Nothing AndAlso
               Not IsDBNull(jumlahNG) Then

                txtJumlahNG.Text =
                    jumlahNG.ToString()

            Else

                txtJumlahNG.Clear()

                MessageBox.Show(
                    "Jumlah NG untuk jenis tersebut tidak ditemukan.",
                    "Informasi",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                )

            End If

            CNN.Close()

        Catch ex As Exception

            If CNN IsNot Nothing AndAlso
               CNN.State = ConnectionState.Open Then

                CNN.Close()

            End If

            MessageBox.Show(
                "Gagal mengambil jumlah NG." &
                vbCrLf &
                ex.Message,
                "Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            )

        End Try

    End Sub


    Private Sub btnTambahUsulan_Click(
        sender As Object,
        e As EventArgs
    ) Handles btnTambahUsulan.Click


        If cmbNamaProduk.SelectedIndex = -1 Then

            MessageBox.Show(
                "Silakan pilih nama produk terlebih dahulu.",
                "Peringatan",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            )

            cmbNamaProduk.Focus()

            Exit Sub

        End If


        If cmbJenisNG.SelectedIndex = -1 Then

            MessageBox.Show(
                "Silakan pilih jenis NG terlebih dahulu.",
                "Peringatan",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            )

            cmbJenisNG.Focus()

            Exit Sub

        End If

        If txtTotalProduksi.Text.Trim() = "" Then

            MessageBox.Show(
                "Total produksi belum tersedia.",
                "Peringatan",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            )

            Exit Sub

        End If

        If txtJumlahNG.Text.Trim() = "" Then

            MessageBox.Show(
                "Jumlah NG belum tersedia.",
                "Peringatan",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            )

            Exit Sub

        End If


        If rtbUsulan.Text.Trim() = "" Then

            MessageBox.Show(
                "Silakan masukkan usulan perbaikan terlebih dahulu.",
                "Peringatan",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            )

            rtbUsulan.Focus()

            Exit Sub

        End If

        usulanSudahDitambahkan = True

        MessageBox.Show(
            "Data usulan berhasil ditambahkan. Silakan tekan tombol Simpan.",
            "Berhasil",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information
        )

    End Sub


    Private Sub btnSimpanUsulan_Click(
        sender As Object,
        e As EventArgs
    ) Handles btnSimpanUsulan.Click


        If rtbUsulan.Text.Trim() = "" Then

            MessageBox.Show(
                "Silakan masukkan usulan perbaikan terlebih dahulu.",
                "Peringatan",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            )

            rtbUsulan.Focus()

            Exit Sub

        End If

        If usulanSudahDitambahkan = False Then

            MessageBox.Show(
                "Silakan tekan tombol Tambahkan terlebih dahulu.",
                "Peringatan",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            )

            btnTambahUsulan.Focus()

            Exit Sub

        End If


        Dim pesan As String =
            "Data usulan perbaikan akan disimpan." &
            vbCrLf &
            vbCrLf &
            "ID Usulan       : " &
            txtIDUsulan.Text &
            vbCrLf &
            "Tanggal         : " &
            dtpTanggal.Value.ToString("dd/MM/yyyy") &
            vbCrLf &
            "Nama Produk     : " &
            cmbNamaProduk.Text &
            vbCrLf &
            "Jenis NG        : " &
            cmbJenisNG.Text &
            vbCrLf &
            "Total Produksi  : " &
            txtTotalProduksi.Text &
            vbCrLf &
            "Jumlah NG       : " &
            txtJumlahNG.Text &
            vbCrLf &
            "Usulan           : " &
            rtbUsulan.Text


        Dim hasil As DialogResult =
            MessageBox.Show(
                pesan,
                "Konfirmasi Simpan",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            )


        If hasil <> DialogResult.Yes Then
            Exit Sub
        End If

        Try

            Koneksi()

            Dim query As String =
                "INSERT INTO [Data_Usulan] " &
                "([ID_Usulan], [Usulan_Perbaikan], [Tanggal_Usulan]) " &
                "VALUES (?, ?, ?)"

            Using cmdSimpan As New OleDbCommand(
                query,
                CNN
            )

                cmdSimpan.Parameters.AddWithValue(
                    "@ID_Usulan",
                    txtIDUsulan.Text.Trim()
                )

                cmdSimpan.Parameters.AddWithValue(
                    "@Usulan_Perbaikan",
                    rtbUsulan.Text.Trim()
                )

                cmdSimpan.Parameters.AddWithValue(
                    "@Tanggal_Usulan",
                    dtpTanggal.Value
                )

                cmdSimpan.ExecuteNonQuery()

            End Using

            CNN.Close()


            MessageBox.Show(
                "Data usulan berhasil disimpan.",
                "Berhasil",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            )


            Form6.Show()
            Me.Hide()


        Catch ex As Exception

            If CNN IsNot Nothing AndAlso
               CNN.State = ConnectionState.Open Then

                CNN.Close()

            End If

            MessageBox.Show(
                "Gagal menyimpan data usulan." &
                vbCrLf &
                ex.Message,
                "Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            )

        End Try

    End Sub

    Private Sub btnResetUsulan_Click(
        sender As Object,
        e As EventArgs
    ) Handles btnResetUsulan.Click

        txtIDUsulan.Clear()

        dtpTanggal.Value =
            DateTime.Now

        cmbNamaProduk.SelectedIndex = -1
        cmbJenisNG.SelectedIndex = -1

        txtTotalProduksi.Clear()
        txtJumlahNG.Clear()

        rtbUsulan.Clear()

        usulanSudahDitambahkan = False

        txtIDUsulan.Focus()

    End Sub

End Class