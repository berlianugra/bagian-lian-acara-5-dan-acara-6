Imports System.Data.OleDb

Public Class FormInputKerugian

    Private totalSudahDitambahkan As Boolean = False

    Private idTotalProdukAktif As String = ""

    Private idNGAktif As String = ""

    Private Sub FormInputKerugian_Load(
        sender As Object,
        e As EventArgs
    ) Handles MyBase.Load

        dtpTanggal.Value = DateTime.Now

        txtTotalProduksi.ReadOnly = True
        txtJumlahNG.ReadOnly = True
        txtTotalBiayaKerugian.ReadOnly = True

        txtTotalProduksi.BackColor = Color.LightGray
        txtJumlahNG.BackColor = Color.LightGray
        txtTotalBiayaKerugian.BackColor = Color.LightGray

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

            Using cmdProduk As New OleDbCommand(query, CNN)

                Using rd As OleDbDataReader =
                    cmdProduk.ExecuteReader()

                    While rd.Read()

                        cmbNamaProduk.Items.Add(
                            New ProdukItem(
                                rd("ID_Produk").ToString(),
                                rd("Nama_Produk").ToString()
                            )
                        )

                    End While

                End Using

            End Using

            CNN.Close()

        Catch ex As Exception

            TutupKoneksi()

            MessageBox.Show(
                "Gagal mengambil data produk." &
                vbCrLf & ex.Message,
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

            Using cmdNG As New OleDbCommand(query, CNN)

                Using rd As OleDbDataReader =
                    cmdNG.ExecuteReader()

                    While rd.Read()

                        cmbJenisNG.Items.Add(
                            rd("Jenis_NG").ToString()
                        )

                    End While

                End Using

            End Using

            CNN.Close()

        Catch ex As Exception

            TutupKoneksi()

            MessageBox.Show(
                "Gagal mengambil jenis NG." &
                vbCrLf & ex.Message,
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
            Exit Sub
        End If


        idTotalProdukAktif = ""
        idNGAktif = ""

        txtTotalProduksi.Clear()
        txtJumlahNG.Clear()
        txtTotalBiayaKerugian.Clear()

        totalSudahDitambahkan = False

        AmbilTotalProduksi(produk.ID)

    End Sub

    Private Sub AmbilTotalProduksi(
        idProduk As String
    )

        Try

            Koneksi()

            Dim query As String =
                "SELECT TOP 1 " &
                "[ID_Total_Produk], " &
                "[Total_Produksi] " &
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

                Using rd As OleDbDataReader =
                    cmdProduksi.ExecuteReader()

                    If rd.Read() Then

                        idTotalProdukAktif =
                            rd("ID_Total_Produk").ToString()

                        txtTotalProduksi.Text =
                            rd("Total_Produksi").ToString()

                    Else

                        idTotalProdukAktif = ""

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

            TutupKoneksi()

            MessageBox.Show(
                "Gagal mengambil data total produksi." &
                vbCrLf & ex.Message,
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

        If idTotalProdukAktif = "" Then

            MessageBox.Show(
                "Data total produksi belum ditemukan.",
                "Peringatan",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            )

            Exit Sub

        End If

        txtJumlahNG.Clear()
        txtTotalBiayaKerugian.Clear()

        idNGAktif = ""
        totalSudahDitambahkan = False

        AmbilDataNG(
            idTotalProdukAktif,
            cmbJenisNG.Text
        )

    End Sub

    Private Sub AmbilDataNG(
        idTotalProduk As String,
        jenisNG As String
    )

        Try

            Koneksi()

            Dim query As String =
                "SELECT TOP 1 " &
                "[ID_NG], " &
                "[Jumlah_NG] " &
                "FROM [Data_Produk_NG] " &
                "WHERE [ID_Total_Produk] = ? " &
                "AND [Jenis_NG] = ? " &
                "ORDER BY [ID_NG]"

            Using cmdNG As New OleDbCommand(
                query,
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

                Using rd As OleDbDataReader =
                    cmdNG.ExecuteReader()

                    If rd.Read() Then

                        idNGAktif =
                            rd("ID_NG").ToString()

                        txtJumlahNG.Text =
                            rd("Jumlah_NG").ToString()

                    Else

                        idNGAktif = ""

                        txtJumlahNG.Clear()

                        MessageBox.Show(
                            "Data NG untuk produk dan jenis tersebut tidak ditemukan.",
                            "Informasi",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information
                        )

                    End If

                End Using

            End Using

            CNN.Close()

        Catch ex As Exception

            TutupKoneksi()

            MessageBox.Show(
                "Gagal mengambil data NG." &
                vbCrLf & ex.Message,
                "Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            )

        End Try

    End Sub

    Private Sub btnTambahTotal_Click(
        sender As Object,
        e As EventArgs
    ) Handles btnTambahTotal.Click

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


        If idNGAktif = "" Then

            MessageBox.Show(
                "Data NG belum ditemukan.",
                "Peringatan",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            )

            Exit Sub

        End If

        AmbilTotalBiayaKerugian()

    End Sub

    Private Sub AmbilTotalBiayaKerugian()

        Try

            Koneksi()

            Dim query As String =
                "SELECT TOP 1 [Total_Biaya_Keugian] " &
                "FROM [Data_Pemrosesan_Kerugian] " &
                "WHERE [ID_Produk] = ? " &
                "AND [ID_NG] = ?"

            Using cmdBiaya As New OleDbCommand(
                query,
                CNN
            )

                Dim produk As ProdukItem =
                    TryCast(cmbNamaProduk.SelectedItem, ProdukItem)

                If produk Is Nothing Then
                    CNN.Close()
                    Exit Sub
                End If

                cmdBiaya.Parameters.AddWithValue(
                    "@ID_Produk",
                    produk.ID
                )

                cmdBiaya.Parameters.AddWithValue(
                    "@ID_NG",
                    idNGAktif
                )

                Dim biaya As Object =
                    cmdBiaya.ExecuteScalar()

                If biaya IsNot Nothing AndAlso
                   Not IsDBNull(biaya) Then

                    txtTotalBiayaKerugian.Text =
                        Convert.ToDecimal(biaya).ToString("N2")

                    totalSudahDitambahkan = True

                    MessageBox.Show(
                        "Total biaya kerugian berhasil ditambahkan.",
                        "Berhasil",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    )

                Else

                    txtTotalBiayaKerugian.Clear()
                    totalSudahDitambahkan = False

                    MessageBox.Show(
                        "Total biaya kerugian belum tersedia untuk produk dan jenis NG tersebut.",
                        "Informasi",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    )

                End If

            End Using

            CNN.Close()

        Catch ex As Exception

            totalSudahDitambahkan = False

            TutupKoneksi()

            MessageBox.Show(
                "Gagal mengambil total biaya kerugian." &
                vbCrLf & ex.Message,
                "Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            )

        End Try

    End Sub

    Private Sub btnSimpan_Click(
        sender As Object,
        e As EventArgs
    ) Handles btnSimpan.Click

        If cmbNamaProduk.SelectedIndex = -1 Then

            MessageBox.Show(
                "Silakan pilih nama produk.",
                "Peringatan",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            )

            Exit Sub

        End If


        If cmbJenisNG.SelectedIndex = -1 Then

            MessageBox.Show(
                "Silakan pilih jenis NG.",
                "Peringatan",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            )

            Exit Sub

        End If


        If idNGAktif = "" Then

            MessageBox.Show(
                "Data NG belum ditemukan.",
                "Peringatan",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            )

            Exit Sub

        End If


        If Not totalSudahDitambahkan Then

            MessageBox.Show(
                "Silakan tekan Tambah Total terlebih dahulu.",
                "Peringatan",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            )

            Exit Sub

        End If


        If txtTotalBiayaKerugian.Text.Trim() = "" Then

            MessageBox.Show(
                "Total biaya kerugian belum tersedia.",
                "Peringatan",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            )

            Exit Sub

        End If


        Dim produk As ProdukItem =
            TryCast(cmbNamaProduk.SelectedItem, ProdukItem)

        If produk Is Nothing Then
            Exit Sub
        End If


        Dim pesan As String =
            "Data kerugian akan disimpan." &
            vbCrLf & vbCrLf &
            "Tanggal       : " &
            dtpTanggal.Value.ToString("dd/MM/yyyy") &
            vbCrLf &
            "Nama Produk   : " &
            cmbNamaProduk.Text &
            vbCrLf &
            "Jenis NG      : " &
            cmbJenisNG.Text &
            vbCrLf &
            "Jumlah NG     : " &
            txtJumlahNG.Text &
            vbCrLf &
            "Total Produksi: " &
            txtTotalProduksi.Text &
            vbCrLf &
            "Total Biaya   : Rp " &
            txtTotalBiayaKerugian.Text


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


        If SimpanDataKerugian(
            produk.ID,
            idNGAktif
        ) Then

            MessageBox.Show(
                "Data kerugian berhasil disimpan.",
                "Berhasil",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            )

            FormRiwayatKerugian.Show()
            Me.Hide()

        End If

    End Sub


    Private Function SimpanDataKerugian(
        idProduk As String,
        idNG As String
    ) As Boolean

        Try

            Koneksi()

            Dim idKerugian As String =
                "KRG" &
                DateTime.Now.ToString("yyyyMMddHHmmssfff")


            Dim query As String =
                "INSERT INTO [Data_Pemrosesan_Kerugian] " &
                "([ID_Kerugian], [Total_Biaya_Keugian], [ID_Produk], [ID_NG]) " &
                "VALUES (?, ?, ?, ?)"

            Using cmdSimpan As New OleDbCommand(
                query,
                CNN
            )

                cmdSimpan.Parameters.AddWithValue(
                    "@ID_Kerugian",
                    idKerugian
                )

                cmdSimpan.Parameters.AddWithValue(
                    "@Total_Biaya_Keugian",
                    Convert.ToDecimal(txtTotalBiayaKerugian.Text)
                )

                cmdSimpan.Parameters.AddWithValue(
                    "@ID_Produk",
                    idProduk
                )

                cmdSimpan.Parameters.AddWithValue(
                    "@ID_NG",
                    idNG
                )

                cmdSimpan.ExecuteNonQuery()

            End Using

            CNN.Close()

            Return True

        Catch ex As Exception

            TutupKoneksi()

            MessageBox.Show(
                "Data gagal disimpan." &
                vbCrLf & ex.Message,
                "Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            )

            Return False

        End Try

    End Function

    Private Sub btnReset_Click(
        sender As Object,
        e As EventArgs
    ) Handles btnReset.Click

        dtpTanggal.Value = DateTime.Now

        cmbNamaProduk.SelectedIndex = -1
        cmbJenisNG.SelectedIndex = -1

        txtJumlahNG.Clear()
        txtTotalProduksi.Clear()
        txtTotalBiayaKerugian.Clear()

        idTotalProdukAktif = ""
        idNGAktif = ""

        totalSudahDitambahkan = False

        cmbNamaProduk.Focus()

    End Sub

    Private Sub TutupKoneksi()

        Try

            If CNN IsNot Nothing AndAlso
               CNN.State <> ConnectionState.Closed Then

                CNN.Close()

            End If

        Catch

        End Try

    End Sub

End Class