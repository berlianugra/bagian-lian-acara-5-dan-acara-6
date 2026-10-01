Public Class FormUsulanPerbaikan

    Private Sub FormUsulanPerbaikan_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        lblInputUsulan.Visible = False
        lblRiwayatUsulan.Visible = False

    End Sub


    Private Sub lblUsulanPerbaikan_MouseEnter(sender As Object, e As EventArgs) Handles lblUsulanPerbaikan.MouseEnter

        lblInputUsulan.Visible = True
        lblRiwayatUsulan.Visible = True

    End Sub


    Private Sub lblInputUsulan_MouseEnter(sender As Object, e As EventArgs) Handles lblInputUsulan.MouseEnter

        lblInputUsulan.ForeColor = Color.Orange

    End Sub


    Private Sub lblInputUsulan_MouseLeave(sender As Object, e As EventArgs) Handles lblInputUsulan.MouseLeave

        lblInputUsulan.ForeColor = Color.White

    End Sub


    Private Sub lblInputUsulan_Click(sender As Object, e As EventArgs) Handles lblInputUsulan.Click

        FormInputUsulanPerbaikan.Show()
        Me.Hide()

    End Sub


    Private Sub lblRiwayatUsulan_MouseEnter(sender As Object, e As EventArgs) Handles lblRiwayatUsulan.MouseEnter

        lblRiwayatUsulan.ForeColor = Color.Orange

    End Sub


    Private Sub lblRiwayatUsulan_MouseLeave(sender As Object, e As EventArgs) Handles lblRiwayatUsulan.MouseLeave

        lblRiwayatUsulan.ForeColor = Color.White

    End Sub


    Private Sub lblRiwayatUsulan_Click(sender As Object, e As EventArgs) Handles lblRiwayatUsulan.Click

        FormRiwayatUsulan.Show()
        Me.Hide()

    End Sub


    Private Sub pcbInputUsulan_Click(sender As Object, e As EventArgs) Handles pcbInputUsulan.Click
        FormInputUsulan.Show()
        Me.Hide()

    End Sub


    Private Sub pcbRiwayatUsulan_Click(sender As Object, e As EventArgs) Handles pcbRiwayatUsulan.Click

        FormRiwayatUsulan.Show()
        Me.Hide()

    End Sub

End Class