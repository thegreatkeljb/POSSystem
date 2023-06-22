Imports System.Data.OleDb

Public Class main

    Dim userProfile As Boolean = False

    Sub enableCashier(ByVal choice)
        btnCashier2.Enabled = choice
        btnCashier.Enabled = choice
        If choice = False Then
            btnCashier2.BackColor = Color.Transparent
        End If
    End Sub

    Sub enableInventory(ByVal choice)
        btnInventory.Enabled = choice
        btnInventory2.Enabled = choice
        If choice = False Then
            btnInventory2.BackColor = Color.Transparent
        End If
    End Sub

    Sub enableAdmin(ByVal choice)
        btnAdmin.Enabled = choice
        btnAdmin2.Enabled = choice
        If choice = False Then
            btnAdmin2.BackColor = Color.Transparent
        End If
    End Sub

    Sub enableReport(ByVal choice)
        btnReports2.Enabled = choice
        btnReport.Enabled = choice
        If choice = False Then
            btnReports2.BackColor = Color.Transparent
        End If
    End Sub

    Private Sub btnAdmin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdmin.Click
        Me.Hide()
        frmMain.Show()
        frmMain.defaultMenuStrip()
        frmMain.defaultUserReg()
    End Sub

    Private Sub btnInventory_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInventory.Click
        Me.Hide()
        InventoryForm.Show()
        InventoryForm.defaultInventoryForm()
        InventoryForm.defaultMenuStrip()
        InventoryForm.clearInventoryForm()
    End Sub

    Private Sub btnCashier_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCashier.Click
        Me.Hide()
        Cashier.Show()
        Cashier.btnNewOrder.PerformClick()
    End Sub

    Private Sub btnReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Hide()
        reports.Show()
        reports.resetForm()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdmin2.Click
        Me.Hide()
        frmMain.Show()
        frmMain.defaultMenuStrip()
        frmMain.defaultUserReg()
    End Sub

    Private Sub btnInventory2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInventory2.Click
        Me.Hide()
        InventoryForm.Show()
        InventoryForm.defaultInventoryForm()
        InventoryForm.defaultMenuStrip()
        InventoryForm.clearInventoryForm()
    End Sub

    Private Sub btnCashier2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCashier2.Click
        Me.Hide()
        Cashier.Show()
        Cashier.btnNewOrder.PerformClick()
    End Sub

    Private Sub btnReports2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReports2.Click
        Me.Hide()
        reports.Show()
        reports.resetForm()
    End Sub


    Sub loadMain()
        btnUser.Text = currentUser
        If isCashier = True Then
            enableCashier(True)
            enableInventory(False)
            enableAdmin(False)
            enableReport(False)
        ElseIf isAdmin = True Then
            enableCashier(True)
            enableInventory(True)
            enableAdmin(True)
            enableReport(True)
        ElseIf isInventory = True Then
            enableCashier(False)
            enableInventory(True)
            enableAdmin(False)
            enableReport(False)
        Else
        End If
    End Sub
    Private Sub main_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadMain()

        picUser.SizeMode = PictureBoxSizeMode.Zoom
        btnAdmin.Focus()

       
    End Sub

    'Did not separate the two functionality. For some reason, sometimes objects don't fall in their rightful position. Just to be safe, Im just doing this instead.
    'For clearing the password
    Sub isDefPass(ByVal x)
        btnSave.Enabled = x
        lblPass1.Visible = x
        lblPass2.Visible = x
        lblPass3.Visible = x
        txtConfirmNewPass.Visible = x
        txtCurrentPass.Visible = x
        txtNewPass.Visible = x
        btnCancel.Visible = x
        btnSeePW.Visible = x
        If x = False Then
            btnChangePass.Enabled = True
            btnEditQuestion.Enabled = True
        Else
            btnChangePass.Enabled = False
            btnEditQuestion.Enabled = False
        End If
    End Sub

    'For clearing the security question
    Sub isDefSec(ByVal x)
        btnSave.Enabled = x
        lblQuestion1.Visible = x
        lblQuestion2.Visible = x
        txtSecurityQuestion.Visible = x
        cmbSecurityQuestion.Visible = x
        btnCancel.Visible = x

        txtConfirmNewPass.Clear()
        txtCurrentPass.Clear()
        txtNewPass.Clear()

        If x = False Then
            btnChangePass.Enabled = True
            btnEditQuestion.Enabled = True
        Else
            btnChangePass.Enabled = False
            btnEditQuestion.Enabled = False
        End If
    End Sub
    'End ######

    Dim tempPass As String

    Private Sub btnUser_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUser.Click
        If userProfile = False Then
            userProfile = True
            pnlUser.Visible = True

            'Edit Acc Objects
            isDefPass(False)
            isDefSec(False)
            'End Edit Acc Objects

            lblUserName.Text = ""
            lblPrivilege.Text = ""
            lblPosition.Text = ""
            tempPass = ""

            Try
                picUser.ImageLocation = ""
                picUser.Refresh()
                picUser.ImageLocation = Application.StartupPath & "\images\" & currentUser & ".jpg"
                picUser.Load()
            Catch ex As Exception

            End Try

            Try
                db = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\posdb.mdb")
                db.Open()
                dbcmd = New OleDbCommand("SELECT [PASSWORD], [USER NAME], [POSITION], [PRIVILEGE] FROM tbluser WHERE [USER ID] like '" & currentUser & "'", db)
                Dim reader As OleDbDataReader = dbcmd.ExecuteReader()
                While reader.Read
                    lblUserName.Text &= reader("USER NAME").ToString()
                    lblPrivilege.Text &= reader("PRIVILEGE").ToString()
                    lblPosition.Text &= reader("POSITION").ToString()
                    tempPass = reader("PASSWORD").ToString
                End While
                reader.Close()
            Catch ex As Exception

            Finally
                db.Close()
            End Try
        Else
            userProfile = False
            pnlUser.Visible = False
        End If

    End Sub


    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        userProfile = False
        btnUser.PerformClick()
    End Sub

    Private Sub btnChangePass_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnChangePass.Click
        isDefPass(True)
        btnChangePass.Enabled = True
    End Sub

    Private Sub btnEditQuestion_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEditQuestion.Click
        isDefSec(True)
        btnEditQuestion.Enabled = True
    End Sub

    Dim userFound As Boolean = False

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If btnChangePass.Enabled = True Then
            If txtConfirmNewPass.Text.Trim = txtNewPass.Text.Trim And txtConfirmNewPass.Text.Trim IsNot String.Empty Then
                If tempPass = txtCurrentPass.Text Then
                    Try
                        db.Open()
                        dbcmd = New OleDb.OleDbCommand("UPDATE tbluser SET [PASSWORD] = @val1 WHERE [USER ID] = @id", db)
                        dbcmd.Parameters.AddWithValue("@val1", txtNewPass.Text)
                        dbcmd.Parameters.AddWithValue("@id", currentUser)
                        dbcmd.ExecuteNonQuery()
                        txtConfirmNewPass.Clear()
                        txtCurrentPass.Clear()
                        txtNewPass.Clear()
                        MessageBox.Show("Password has been successfully updated!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    Finally
                        db.Close()
                    End Try
                Else
                    MessageBox.Show("The password you enter did not match. Please try again", "Invalid Password", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    txtConfirmNewPass.Clear()
                    txtCurrentPass.Clear()
                    txtNewPass.Clear()
                End If

            Else
                txtConfirmNewPass.Clear()
                txtCurrentPass.Clear()
                txtNewPass.Clear()
                MessageBox.Show("The password you enter did not match. Please try again", "Invalid Password", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

        ElseIf btnEditQuestion.Enabled = True Then
            Dim pr As DialogResult = MessageBox.Show("Are you sure you want to make this change?", "Modify Security Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If pr = DialogResult.Yes Then
                If txtSecurityQuestion.Text = String.Empty Or cmbSecurityQuestion.SelectedIndex = -1 Then
                    MessageBox.Show("Invalid input. Please do not leave anything blank", "Unsuccessful", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                Else
                    Try
                        db.Open()
                        dbcmd = New OleDb.OleDbCommand("UPDATE tbluser SET [QUESTION] = @val1, [ANSWER] = @val2 WHERE [USER ID] = @id", db)
                        dbcmd.Parameters.AddWithValue("@val1", cmbSecurityQuestion.SelectedItem)
                        dbcmd.Parameters.AddWithValue("@val2", txtSecurityQuestion.Text)
                        dbcmd.Parameters.AddWithValue("@id", currentUser)
                        dbcmd.ExecuteNonQuery()
                        MessageBox.Show("Security question has been successfully updated!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        cmbSecurityQuestion.SelectedIndex = -1
                        txtSecurityQuestion.Clear()
                    Catch ex As Exception

                    Finally
                        db.Close()
                    End Try

                End If
            End If
        End If
    End Sub

    Dim passChar As Boolean = True
    Private Sub btnSeePW_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSeePW.Click
        If passChar = True Then
            txtConfirmNewPass.PasswordChar = ""
            txtCurrentPass.PasswordChar = ""
            txtNewPass.PasswordChar = ""
            passChar = False
        Else
            txtConfirmNewPass.PasswordChar = "●"
            txtCurrentPass.PasswordChar = "●"
            txtNewPass.PasswordChar = "●"
            passChar = True
        End If
    End Sub

    Private Sub btnLogout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogout.Click
        Dim res As DialogResult = MessageBox.Show("Are you sure you want to logout?", "LOGOUT", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If res = DialogResult.Yes Then
            logout()
            Me.Hide()
            login.Show()
        End If
    End Sub


End Class