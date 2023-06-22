Imports System.Data.OleDb
Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel
Imports CrystalDecisions.CrystalReports.Engine

Public Class frmMain
    'Variables

    'For checking if creating or updating of data is being done.
    Dim isUpdating As Boolean = False

    Dim canNavigate As Boolean = True

    Dim rowCtr As Integer = 0
    Dim max As Integer = 0

    'End of Variables

    'Functions
    Sub clearMenuStrip()
        tsmCreate.Enabled = False
        tsmDelete.Enabled = False
        tsmEdit.Enabled = False
        tsmSave.Enabled = False
        tsmSearch.Enabled = False
        tsmCancel.Enabled = False
    End Sub

    'Default form of menustrip
    Sub defaultMenuStrip()
        clearMenuStrip()
        tsmCreate.Enabled = True
        tsmDelete.Enabled = True
        tsmEdit.Enabled = True
        tsmSearch.Enabled = True    
    End Sub

    'Clears the fill-up form
    Sub clearUserReg()
        txtUsername.Clear()
        txtPassword.Clear()
        txtConfirmPass.Clear()
        cmbPosition.SelectedIndex = -1
        cmbPrivilege.SelectedIndex = -1
    End Sub

    Sub isTextBoxEnabled(ByVal isBool)
        txtUserID.Enabled = isBool
        txtUsername.Enabled = isBool
        txtPassword.Enabled = isBool
        txtConfirmPass.Enabled = isBool
        cmbPosition.Enabled = isBool
        cmbPrivilege.Enabled = isBool
        txtposition.Enabled = isBool
        txtprivilege.Enabled = isBool
    End Sub

    'Default form of the user registration/admin form
    Sub defaultUserReg()

        'Check if creating or updating of user information is being done
        linkUploadPhoto.Visible = False
        isUpdating = False

        Try
            db = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\posdb.mdb")
            db.Open()
            dba = New OleDb.OleDbDataAdapter("SELECT [USER ID],[USER NAME],[POSITION],[PRIVILEGE] FROM tbluser", db)
            dbds = New DataSet
            dba.Fill(dbds, "tbluser")
            dgUserForm.DataSource = dbds.Tables("tbluser")
        Catch ex As Exception
            MsgBox("Error! Please contact an administrator")
        Finally
            db.Close()
        End Try

        autoID()


        max = dgUserForm.RowCount - 2

        cmbPosition.Visible = True
        cmbPrivilege.Visible = True
        pnlStockReport.Visible = False

    End Sub

    'Generates automatic ID number
    Sub autoID()
        txtUserID.Clear()
        'Check the number of rows and use it to indentify the latest USER ID added
        If dgUserForm.RowCount > 1 Then
            Dim num As Integer = dgUserForm.RowCount - 2
            Dim temp As String = dbds.Tables("tbluser").Rows(num).Item("USER ID")

            'Fetch the latest primary key from database and increment it
            db.Open()
            dbcmd = New OleDbCommand("SELECT [NUMBER] FROM [tbluser] WHERE [USER ID]= '" & temp & "'", db)
            Dim reader As OleDbDataReader = dbcmd.ExecuteReader()
            lblTempNo.Text = ""
            While reader.Read
                lblTempNo.Text &= reader("NUMBER").ToString()
            End While
            reader.Close()
            db.Close()

            txtUserID.Text = "BTLRN-2023-" & (1 + Val(lblTempNo.Text))
        End If
    End Sub

    Sub nav(ByVal selectedRow)
        If canNavigate = True Then
            cmbPosition.Visible = False
            cmbPrivilege.Visible = False
                Try
                    db.Open()
                    db = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\posdb.mdb")
                    dba = New OleDb.OleDbDataAdapter("SELECT [USER ID], [USER NAME], [POSITION], [PRIVILEGE] FROM tbluser", db)
                    dbds = New DataSet
                    dba.Fill(dbds, "tbluser")
                    dgUserForm.DataSource = dbds.Tables("tbluser")
                    db.Close()
                If dgUserForm.RowCount > 1 Then
                    dgUserForm.Rows(selectedRow).DefaultCellStyle.BackColor = Color.Gold
                    txtUserID.Text = dbds.Tables("tbluser").Rows(selectedRow).Item("USER ID")
                    txtUsername.Text = dbds.Tables("tbluser").Rows(selectedRow).Item("USER NAME")
                    txtposition.Text = dbds.Tables("tbluser").Rows(selectedRow).Item("POSITION")
                    txtprivilege.Text = dbds.Tables("tbluser").Rows(selectedRow).Item("PRIVILEGE")
                End If
                Catch ex As Exception
                    ex.ToString()
                Finally
                    db.Close()
                End Try
            End If
    End Sub
    'End of Functions

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        defaultMenuStrip()
        defaultUserReg()
        isTextBoxEnabled(False)

        txtUserID.Focus()
        dgUserForm.ForeColor = Color.Black
        lblTempNo.Visible = False
        tsmPrint.Enabled = True

        dgUserForm.DefaultCellStyle.SelectionBackColor = Color.White
        dgUserForm.DefaultCellStyle.SelectionForeColor = Color.Black

        picUser.SizeMode = PictureBoxSizeMode.Zoom

        currUser.Text = currentUser
    End Sub

    Private Sub CreateAccountToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmCreate.Click
        clearUserReg()
        defaultUserReg()
        isTextBoxEnabled(True)
        txtUserID.Enabled = False
        clearMenuStrip()

        canNavigate = False
        tsmCancel.Enabled = True
        tsmSave.Enabled = True
        tsmCreate.Enabled = True
        linkUploadPhoto.Visible = True
    End Sub

    Private Sub EditAccountToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmEdit.Click
        clearUserReg()
        defaultUserReg()
        clearMenuStrip()
        tsmCancel.Enabled = True
        tsmSave.Enabled = True
        tsmEdit.Enabled = True
        isTextBoxEnabled(False)
        txtPassword.Enabled = True
        txtConfirmPass.Enabled = True
        cmbPosition.Visible = False
        cmbPrivilege.Visible = False
        isUpdating = True
    End Sub

    Private Sub SearchToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmSearch.Click
        clearMenuStrip()
        tsmCancel.Enabled = True
        tsmSearch.Enabled = True
        dgUserForm.ReadOnly = False
        Dim inp As String = InputBox("Enter the name of the user you want to search:", "Search User")
        If inp.Length < 1 Then
        Else
            Try
                db.Open()
                dba = New OleDb.OleDbDataAdapter("SELECT [USER ID], [USER NAME], [POSITION], [PRIVILEGE] FROM tbluser WHERE [USER NAME] like '%" & inp.Trim & "%'", db)
                dbds = New DataSet
                dba.Fill(dbds, "tbluser")
                dgUserForm.DataSource = dbds.Tables("tbluser")
            Catch ex As Exception
                MsgBox(ex.ToString)
            Finally
                db.Close()
            End Try
        End If
    End Sub

    Private Sub DeleteAccountToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmDelete.Click
        clearMenuStrip()
        tsmCancel.Enabled = True
        tsmSave.Enabled = True
        tsmDelete.Enabled = True
        isTextBoxEnabled(False)
        txtUserID.Enabled = True
        isUpdating = True
    End Sub

    Private Sub InventoryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        InventoryForm.Show()
        Me.Hide()
    End Sub

    Private Sub tsmExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmExit.Click
        Dim res As DialogResult = MessageBox.Show("Are you sure you want to exit?", "Exit", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If res = DialogResult.Yes Then
            logout()
            Me.Close()
        End If
    End Sub

    Private Sub tsmSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmSave.Click
        Dim msgConf As DialogResult

        If tsmCreate.Enabled = True And canNavigate = False Then
            msgConf = MessageBox.Show("Are you sure you want to register this user?", "Registration", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If msgConf = DialogResult.Yes Then
                If txtPassword.Text.Trim = txtConfirmPass.Text.Trim Then
                    If txtPassword.Text.Trim <> "" Then
                        db.Open()
                        dbcmd = New OleDb.OleDbCommand("INSERT INTO tbluser([USER ID], [USER NAME], [PASSWORD], [POSITION], [PRIVILEGE]) VALUES ('" & txtUserID.Text.Trim & "', '" & txtUsername.Text.Trim & "', '" & txtPassword.Text.Trim & "', '" & cmbPosition.Text.Trim & "', '" & cmbPrivilege.Text.Trim & "')", db)
                        dbcmd.ExecuteNonQuery()
                        MessageBox.Show("Registration Successful", "Registered", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        defaultUserReg()
                        clearUserReg()
                    Else
                        MessageBox.Show("Invalid password. The textbox is blank", "Failed", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        txtPassword.Clear()
                        txtConfirmPass.Clear()
                    End If
                Else
                    MessageBox.Show("Password do not match. Please try again", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    txtPassword.Clear()
                    txtConfirmPass.Clear()
                End If
                db.Close()
                defaultUserReg()
            End If
        End If

        If tsmDelete.Enabled = True Then
            msgConf = MessageBox.Show("Are you sure you want to delete this user? This action cannot be undone!", "Delete User", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
            If msgConf = DialogResult.Yes Then
                Try
                    db.Open()
                    dbcmd = New OleDb.OleDbCommand("INSERT INTO tblArchive([USER ID], [USER NAME], [PASSWORD], [POSITION], [PRIVILEGE], [DATE ARCHIVE]) VALUES ('" & txtUserID.Text.Trim & "', '" & txtUsername.Text.Trim & "', '" & txtPassword.Text.Trim & "', '" & txtposition.Text.Trim & "', '" & txtprivilege.Text.Trim & "', '" & currentDate & "')", db)
                    dbcmd.ExecuteNonQuery()
                    dbcmd = New OleDb.OleDbCommand("DELETE FROM tbluser WHERE [USER ID] = '" & txtUserID.Text & "'", db)
                    dbcmd.ExecuteNonQuery()
                    MessageBox.Show("User Deleted", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    defaultUserReg()
                    clearUserReg()
                Catch ex As Exception
                    MsgBox(ex.ToString)
                Finally
                    db.Close()
                End Try
                defaultUserReg()
            End If

        ElseIf tsmEdit.Enabled = True Then
            msgConf = MessageBox.Show("Are you sure you want to change " & txtUserID.Text & " password?", "Change Password", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If msgConf = DialogResult.Yes Then
                If txtConfirmPass.Text = txtPassword.Text Then
                    If txtPassword.Text.Trim <> "" Then
                        Try
                            db.Open()
                            dbcmd = New OleDb.OleDbCommand("UPDATE tbluser SET [PASSWORD] = @val1 WHERE [USER ID] = @id", db)
                            dbcmd.Parameters.AddWithValue("@val1", txtPassword.Text)
                            dbcmd.Parameters.AddWithValue("@id", txtUserID.Text)
                            dbcmd.ExecuteNonQuery()
                            defaultUserReg()
                            clearUserReg()
                            MessageBox.Show("Password has been successfully updated!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Catch ex As Exception
                            MsgBox(ex.ToString)
                        Finally
                            db.Close()
                        End Try
                    Else
                        MessageBox.Show("Invalid password. The textbox is blank", "Failed", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        txtPassword.Clear()
                        txtConfirmPass.Clear()
                    End If
                Else
                    MessageBox.Show("The password you entered does not match. Please try again!", "Failed", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    txtPassword.Clear()
                    txtConfirmPass.Clear()
                End If
            End If
            defaultUserReg()
        End If

    End Sub

    Private Sub tsmCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmCancel.Click
        clearUserReg()
        defaultUserReg()
        isTextBoxEnabled(False)
        defaultMenuStrip()

        canNavigate = True
    End Sub

    Private Sub txtUserID_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUserID.TextChanged
        Try
            picUser.ImageLocation = ""
            picUser.Refresh()
            picUser.ImageLocation = Application.StartupPath & "\images\" & txtUserID.Text.Trim & ".jpg"
            picUser.Load()
        Catch ex As Exception
            'Sana di mag-error hehe
        End Try

        If isUpdating = True Then
            If txtUserID.TextLength <> 0 Then
                cmbPosition.Visible = False
                cmbPrivilege.Visible = False
                db.Open()
                db = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\posdb.mdb")
                dba = New OleDb.OleDbDataAdapter("SELECT [USER ID], [USER NAME], [POSITION], [PRIVILEGE] FROM tbluser WHERE [USER ID] = '" & txtUserID.Text & "'", db)
                dbds = New DataSet
                dba.Fill(dbds, "tbluser")
                dgUserForm.DataSource = dbds.Tables("tbluser")
                If dgUserForm.RowCount > 1 Then
                    txtUsername.Text = dbds.Tables("tbluser").Rows(0).Item("USER NAME")
                    txtposition.Text = dbds.Tables("tbluser").Rows(0).Item("POSITION")
                    txtprivilege.Text = dbds.Tables("tbluser").Rows(0).Item("PRIVILEGE")
                End If
                db.Close()
            Else
                defaultUserReg()
                clearUserReg()
            End If
        End If

    End Sub

    Private Sub ReturnToMainMenuToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReturnToMainMenuToolStripMenuItem.Click
        Me.Hide()
        main.Show()
    End Sub

    Private Sub tsmBegin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmBegin.Click
        rowCtr = 0
        nav(rowCtr)
    End Sub

    Private Sub tsmNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmNext.Click
        If rowCtr = max Then
            rowCtr = 0
        Else
            rowCtr += 1
        End If
        nav(rowCtr)
    End Sub

    Private Sub tsmPrev_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmPrev.Click
        If rowCtr = 0 Then
            rowCtr = max
        Else
            rowCtr -= 1
        End If
        nav(rowCtr)
    End Sub

    Private Sub tsmLast_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmLast.Click
        rowCtr = max
        nav(rowCtr)
    End Sub

    Private Sub linkUploadPhoto_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles linkUploadPhoto.LinkClicked
        If txtuserid.Text.Trim.Length > 0 Then
            Dim ofdpic As New OpenFileDialog()

            ofdpic.FileName = ""
            If ofdpic.ShowDialog() = DialogResult.OK Then
                Try
                    Dim destFilePath As String = Application.StartupPath & "\images\" & Trim(txtUserID.Text) & ".jpg"
                    My.Computer.FileSystem.CopyFile(ofdpic.FileName, destFilePath, True)
                    picUser.ImageLocation = destFilePath
                    picUser.Load()
                Catch err As Exception
                    MsgBox(err.ToString())
                End Try
            End If
        Else
            MsgBox("Enter your user ID", MsgBoxStyle.Information, "Warning")
        End If
    End Sub


    Private Sub ExportToExcel()
        If dgUserForm.RowCount > 0 Then
            Dim excelApp As New Excel.Application()

            Dim excelWorkbook As Excel.Workbook = excelApp.Workbooks.Add(Type.Missing)
            Dim excelWorksheet As Excel.Worksheet = excelWorkbook.Sheets(1)

            ' Set the column headers in Excel
            For i As Integer = 0 To dgUserForm.Columns.Count - 1
                excelWorksheet.Cells(1, i + 1) = dgUserForm.Columns(i).HeaderText
            Next

            ' Export data from DataGrid to Excel
            For i As Integer = 0 To dgUserForm.Rows.Count - 1
                For j As Integer = 0 To dgUserForm.Columns.Count - 1
                    If dgUserForm.Rows(i).Cells(j).Value IsNot Nothing Then
                        excelWorksheet.Cells(i + 2, j + 1) = dgUserForm.Rows(i).Cells(j).Value.ToString()
                    End If
                Next
            Next

            ' Save the Excel file
            Dim saveFileDialog As New SaveFileDialog()
            saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
            saveFileDialog.FileName = "The ISO Team Enterprise User Report"
            If saveFileDialog.ShowDialog() = DialogResult.OK Then
                excelWorkbook.SaveAs(saveFileDialog.FileName)
                MessageBox.Show("Data exported successfully.")
            End If

            ' Clean up Excel objects
            excelWorkbook.Close()
            excelApp.Quit()
            releaseObject(excelWorksheet)
            releaseObject(excelWorkbook)
            releaseObject(excelApp)
        Else
            MessageBox.Show("No data available to export", "Invalid Action", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    ' Release COM objects to avoid memory leaks
    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub tsmPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmPrint.Click
        ExportToExcel()
    End Sub

    Private Sub tsmReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmReport.Click
        pnlStockReport.Visible = True
        pnlStockReport.Dock = DockStyle.Fill

        Dim userReport As New ReportDocument

        userReport.Load(Application.StartupPath & "\reports\userReport.rpt")
        crepViewerStocks.ReportSource = userReport
        crepViewerStocks.Refresh()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        pnlStockReport.Visible = False
    End Sub

    Dim passChar = True
    Private Sub btnPassChar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPassChar.Click
        If passChar = True Then
            txtPassword.PasswordChar = ""
            txtConfirmPass.PasswordChar = ""
            passChar = False
        Else
            txtPassword.PasswordChar = "●"
            txtConfirmPass.PasswordChar = "●"
            passChar = True
        End If
    End Sub

    Private Sub SoftwareToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SoftwareToolStripMenuItem.Click
        Me.Hide()
        frmAbout.Show()
    End Sub
End Class
