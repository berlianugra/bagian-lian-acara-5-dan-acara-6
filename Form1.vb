Public Class FormKerugian

    Private Sub FormKerugian_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        lblInputKerugian.Visible = False
        lblRiwayatKerugian.Visible = False

    End Sub


    Private Sub lblKerugian_MouseEnter(sender As Object, e As EventArgs) Handles lblKerugian.MouseEnter

        lblInputKerugian.Visible = True
        lblRiwayatKerugian.Visible = True

    End Sub


    Private Sub lblInputKerugian_MouseEnter(sender As Object, e As EventArgs) Handles lblInputKerugian.MouseEnter

        lblInputKerugian.ForeColor = Color.Orange

    End Sub


    Private Sub lblInputKerugian_MouseLeave(sender As Object, e As EventArgs) Handles lblInputKerugian.MouseLeave

        lblInputKerugian.ForeColor = Color.White

    End Sub


    Private Sub lblInputKerugian_Click(sender As Object, e As EventArgs) Handles lblInputKerugian.Click

        FormInputKerugian.Show()
        Me.Hide()

    End Sub


    Private Sub lblRiwayatKerugian_MouseEnter(sender As Object, e As EventArgs) Handles lblRiwayatKerugian.MouseEnter

        lblRiwayatKerugian.ForeColor = Color.Orange

    End Sub


    Private Sub lblRiwayatKerugian_MouseLeave(sender As Object, e As EventArgs) Handles lblRiwayatKerugian.MouseLeave

        lblRiwayatKerugian.ForeColor = Color.White

    End Sub


    Private Sub lblRiwayatKerugian_Click(sender As Object, e As EventArgs) Handles lblRiwayatKerugian.Click

        FormRiwayatKerugian.Show()
        Me.Hide()

    End Sub


    Private Sub pcbInputKerugian_Click(sender As Object, e As EventArgs) Handles pcbInputKerugian.Click
        FormInputKerugian.Show()
        Me.Hide()

    End Sub


    Private Sub pcbRiwayatKerugian_Click(sender As Object, e As EventArgs) Handles pcbRiwayatKerugian.Click

        FormRiwayatKerugian.Show()
        Me.Hide()

    End Sub

End Class