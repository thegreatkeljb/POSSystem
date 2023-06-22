Imports System.Data.OleDb

Public Class login
    Dim loginCtr As Integer = 4
    Dim passChar As Boolean = True

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Dim res As DialogResult = MessageBox.Show("Wait! Are you sure you want to leave?", "Exit", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If res = DialogResult.Yes Then
            Me.Close()
        Else
            txtUserID.Focus()
        End If
    End Sub

    Private Sub login_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtUserID.Focus()
        txtTempPW.Visible = False
    End Sub

    Private Sub btnLogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogin.Click
        txtTempPW.Clear()
        txtTempPriv.Clear()
        db = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\posdb.mdb")
        Try
            db.Open()
            dbcmd = New OleDbCommand("SELECT [PASSWORD], [POSITION], [USER NAME] FROM [tbluser] WHERE [USER ID]= '" & txtUserID.Text & "'", db)
            Dim reader As OleDbDataReader = dbcmd.ExecuteReader()
            While reader.Read
                txtTempPW.Text &= reader("PASSWORD").ToString()
                txtTempPriv.Text &= reader("POSITION").ToString
                currentUserName = reader("USER NAME").ToString
            End While
            reader.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        db.Close()


        If txtPassword.Text <> "" Then
            If txtPassword.Text = txtTempPW.Text Then
                currentUser = txtUserID.Text
                If txtTempPriv.Text.Trim = "Manager" Then
                    isAdmin = True
                ElseIf txtTempPriv.Text.Trim = "Inventory Clerk" Then
                    isInventory = True
                ElseIf txtTempPriv.Text.Trim = "Cashier" Then
                    isCashier = True
                Else
                End If

                Try
                    Dim currentTime As String = currentDateAndTime.ToString("h:mm tt")
                    db.Open()
                    dbcmd = New OleDb.OleDbCommand("INSERT INTO tbl_dtr([user_id], [username], [login_date], [login_time]) VALUES ('" & currentUser & "', '" & currentUserName & "','" & currentDate & "','" & currentTime & "')", db)
                    dbcmd.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.ToString)
                Finally
                    db.Close()
                End Try

                Me.Hide()
                main.loadMain()
                main.Show()
                txtUserID.Clear()
                txtPassword.Clear()
            Else
                loginCtr -= 1
                MessageBox.Show("The User ID or Password you entered is incorrect. You only have " & loginCtr & " attempts left")
                txtPassword.Clear()
                txtUserID.Clear()
            End If
        End If

        If loginCtr = 0 Then
            Me.Close()
        End If

    End Sub

    Private Sub txtUserID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtUserID.KeyDown
        If e.KeyCode = Keys.Enter Then
            btnLogin.PerformClick()
        End If
    End Sub

    Private Sub txtPassword_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPassword.KeyDown
        If e.KeyCode = Keys.Enter Then
            btnLogin.PerformClick()
        End If
    End Sub

    Private Sub btnSeePW_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSeePW.Click
        If passChar = True Then
            txtPassword.PasswordChar = ""
            passChar = False
        Else
            txtPassword.PasswordChar = "●"
            passChar = True
        End If
    End Sub

    Private Sub linkForgotPW_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles linkForgotPW.LinkClicked
        pnlForgotPass.Visible = True
        pnlUserID.Visible = True
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        pnlForgotPass.Visible = False
    End Sub

    Private Sub btnCancelUserID_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancelUserID.Click
        pnlForgotPass.Visible = False
    End Sub

    Dim secQuestion As String
    Dim secAns As String

    Private Sub btnEnterID_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEnterID.Click
        secQuestion = ""
        db = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\posdb.mdb")
        Try
            db.Open()
            dbcmd = New OleDbCommand("SELECT [QUESTION], [ANSWER] FROM [tbluser] WHERE [USER ID]= '" & txtEnterID.Text.Trim & "'", db)
            Dim reader As OleDbDataReader = dbcmd.ExecuteReader()
            While reader.Read
                secQuestion = reader("QUESTION").ToString()
                secAns = reader("ANSWER").ToString
            End While
            reader.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        db.Close()

        If secQuestion = String.Empty Then
            MessageBox.Show("Invalid User ID. Please try again", "Invalid ID", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            lblSecQuestion.Text = secQuestion
            pnlUserID.Visible = False
        End If
    End Sub

    Dim tryCtr As Integer = 3
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If secAns.ToUpper = txtSecAnswer.Text.ToUpper Then
            If txtNewPass.Text = txtConfirmPass.Text And txtConfirmPass.Text IsNot String.Empty Then

                Try
                    db.Open()
                    dbcmd = New OleDb.OleDbCommand("UPDATE tbluser SET [PASSWORD] = @val1 WHERE [USER ID] = @id", db)
                    dbcmd.Parameters.AddWithValue("@val1", txtConfirmPass.Text)
                    dbcmd.Parameters.AddWithValue("@id", txtEnterID.Text)
                    dbcmd.ExecuteNonQuery()
                Catch ex As Exception
                Finally
                    db.Close()
                End Try
                MessageBox.Show("Password change successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                pnlForgotPass.Visible = False
                txtConfirmPass.Clear()
                txtNewPass.Clear()
                txtEnterID.Clear()

            Else
                MessageBox.Show("Passwords do not match. Please try again", "Invalid Password", MessageBoxButtons.OK, MessageBoxIcon.Information)
                txtConfirmPass.Clear()
                txtNewPass.Clear()
            End If

        Else
            tryCtr -= 1
            MessageBox.Show("Wrong answer to the question. You only try for " & tryCtr & " more time", "Security Question", MessageBoxButtons.OK, MessageBoxIcon.Information)
            txtSecAnswer.Clear()
        End If

        If tryCtr = 0 Then
            Me.Close()
        End If

    End Sub

    Dim quesPWChar As Boolean = True
    Private Sub btnQuesPW_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnQuesPW.Click
        If quesPWChar = True Then
            txtConfirmPass.PasswordChar = ""
            txtNewPass.PasswordChar = ""
            quesPWChar = False
        Else
            txtConfirmPass.PasswordChar = "●"
            txtNewPass.PasswordChar = "●"
            quesPWChar = True
        End If
    End Sub
End Class