Imports System.Diagnostics

Public Class frmAbout

    Private Sub frmAbout_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub btnCallToAction_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCallToAction.Click
        Dim url As String = "https://github.com/thegreatkeljb?tab=repositories"
        Process.Start(url)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim url As String = "https://www.facebook.com/kingjeybie/"
        Process.Start(url)
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        MessageBox.Show("Email at: shizune.mjbg@gmail.com", "EMAIL", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Dim url As String = "https://mail.google.com/mail/u/0/"
        Process.Start(url)
    End Sub

    Private Sub btnReturn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReturn.Click
        Me.Hide()
        main.Show()
    End Sub
End Class