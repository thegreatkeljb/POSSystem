Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel
Imports CrystalDecisions.CrystalReports.Engine

Public Class InventoryForm

    'Var
    Dim isTrue As Boolean = True
    Dim isUpdating As Boolean = False
    Dim rowCtr As Integer = 0
    Dim max As Integer = 0
    Dim isCritValue As Boolean = False
    'End Var

    'Functions
    Sub defaultInventoryForm()
        pnlStockReport.Visible = False
        linkUploadPhoto.Visible = False
        db = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\inventorydb.mdb")
        Try
            dba = New OleDb.OleDbDataAdapter("SELECT [Barcode Number], [Item Name], [Item Description], [Selling Price], [Buying Price], [Quantity], [Critical Value] FROM tbl_items", db)
            dbds = New DataSet
            dba.Fill(dbds, "tbl_items")
        Catch ex As Exception
            MsgBox(ex.ToString)
        Finally
            db.Close()
        End Try

        dgInventoryForm.DataSource = dbds.Tables("tbl_items")

        For Each row As DataGridViewRow In dgInventoryForm.Rows
            If row.Cells("Quantity").Value IsNot Nothing And row.Cells("Critical Value").Value IsNot Nothing Then
                If Integer.Parse(row.Cells("Quantity").Value) <= Integer.Parse(row.Cells("Critical Value").Value) Then
                    If Integer.Parse(row.Cells("Quantity").Value) = 0 Then
                        row.DefaultCellStyle.BackColor = Color.Red
                    Else
                        row.DefaultCellStyle.BackColor = Color.Gold
                    End If

                End If
            Else

            End If
        Next
        isUpdating = False
        max = dgInventoryForm.RowCount - 2
        lblNumItem.Text = max + 1
    End Sub

    Sub isTextBoxEnabled(ByVal isTrue)
        txtBarcodeNumber.Enabled = isTrue
        txtItemName.Enabled = isTrue
        txtItemDescription.Enabled = isTrue
        txtBuyingPrice.Enabled = isTrue
        txtSellingPrice.Enabled = isTrue
        txtQuantity.Enabled = isTrue
        txtCriticalValue.Enabled = isTrue
    End Sub

    Sub clearInventoryForm()
        txtBarcodeNumber.Clear()
        txtItemName.Clear()
        txtItemDescription.Clear()
        txtBuyingPrice.Clear()
        txtSellingPrice.Clear()
        txtQuantity.Clear()
        txtCriticalValue.Clear()
        lblNumItem.Text = max + 1
    End Sub

    Sub clearMenuStrip()
        tsmCreate.Enabled = False
        tsmDelete.Enabled = False
        tsmEdit.Enabled = False
        tsmSave.Enabled = False
        tsmSearch.Enabled = False
        tsmCancel.Enabled = False
    End Sub

    Sub defaultMenuStrip()
        clearMenuStrip()
        lblDescription.Text = "Welcome to the shop"
        isTextBoxEnabled(False)
        tsmCreate.Enabled = True
        tsmDelete.Enabled = True
        tsmEdit.Enabled = True
        tsmSearch.Enabled = True
    End Sub

    Sub nav(ByVal selectedRow)
        Try
            db.Open()
            db = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\inventorydb.mdb")
            dba = New OleDb.OleDbDataAdapter("SELECT [Barcode Number], [Item Name], [Item Description], [Selling Price], [Buying Price], [Quantity] FROM tbl_items", db)
            dbds = New DataSet
            dba.Fill(dbds, "tbl_items")
            dgInventoryForm.DataSource = dbds.Tables("tbl_items")
            db.Close()
            If dgInventoryForm.RowCount > 1 Then
                txtBarcodeNumber.Text = dbds.Tables("tbl_items").Rows(selectedRow).Item("Barcode Number")
                txtItemName.Text = dbds.Tables("tbl_items").Rows(selectedRow).Item("Item Name")
                txtItemDescription.Text = dbds.Tables("tbl_items").Rows(selectedRow).Item("Item Description")
                txtBuyingPrice.Text = dbds.Tables("tbl_items").Rows(selectedRow).Item("Buying Price")
                txtSellingPrice.Text = dbds.Tables("tbl_items").Rows(selectedRow).Item("Selling Price")
                txtQuantity.Text = dbds.Tables("tbl_items").Rows(selectedRow).Item("Quantity")
                txtCriticalValue.Text = dbds.Tables("tbl_items").Rows(selectedRow).Item("Critical Value")
            End If
        Catch ex As Exception
            ex.ToString()
        Finally
            db.Close()
        End Try
    End Sub


    Sub critVal()
        Dim quanti As Integer = 0
        Dim crit As Integer = 0
        lblCrit.Text = ""
        lblCrit.Text = "Stocks on Critical Level:"

        If max > 0 Then
            db = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\inventorydb.mdb")
            db.Open()
            dbcmd = New OleDb.OleDbCommand("SELECT * FROM tbl_items", db)

            Using reader As OleDbDataReader = dbcmd.ExecuteReader()
                While reader.Read()
                    Dim itemName As Object = reader("Item Name")
                    Dim objOne As Object = reader("Quantity")
                    quanti = Convert.ToInt32(objOne)
                    Dim objTwo As Object = reader("Critical Value")
                    crit = Convert.ToInt32(objTwo)

                    If quanti <= crit Then
                        lblCrit.Text &= vbCrLf & itemName & " is in critical value. Stocks left: " & quanti
                        btnWarning.Visible = True
                    End If
                End While
            End Using
        End If
    End Sub

    Function checkInput()
        Dim d As Double
        Dim i As Integer
        If Not Double.TryParse(txtBarcodeNumber.Text.Trim, d) Then
            MessageBox.Show("Invalid Input. Please input numbers in BARCODE NUMBER", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtBarcodeNumber.Clear()
            Return False
        ElseIf Not Double.TryParse(txtSellingPrice.Text.Trim, d) Then
            MessageBox.Show("Invalid Input. Please input numbers in SELLING PRICE", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtSellingPrice.Clear()
            Return False
        ElseIf Not Double.TryParse(txtBuyingPrice.Text.Trim, d) Then
            MessageBox.Show("Invalid Input. Please input numbers in BUYING PRICE", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtBuyingPrice.Clear()
            Return False
        ElseIf Not Double.TryParse(txtQuantity.Text.Trim, i) Then
            MessageBox.Show("Invalid Input. Please input numbers in QUANTITY", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtBuyingPrice.Clear()
            Return False
        ElseIf Not Double.TryParse(txtCriticalValue.Text.Trim, i) Then
            MessageBox.Show("Invalid Input. Please input numbers in CRITICAL VALUE", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtCriticalValue.Clear()
            Return False
        Else
            Return True
        End If
    End Function

    Function isTextboxFull()
        Dim allTextBoxesFilled As Boolean = True

        For Each textBox As TextBox In Me.Controls.OfType(Of TextBox)()
            If String.IsNullOrWhiteSpace(textBox.Text) Then
                allTextBoxesFilled = False
                Exit For
            End If
        Next

        If allTextBoxesFilled Then
            Return True
        Else
            MessageBox.Show("Please fill all the input box", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return True
        End If
    End Function
    'End Functions


    Private Sub InventoryForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        currUser.Text = currentUser
        defaultInventoryForm()
        defaultMenuStrip()
        dgInventoryForm.ForeColor = Color.Black
        dgInventoryForm.ReadOnly = True
        dgInventoryForm.DefaultCellStyle.Font = New Font("Century Gothic", 9)
        dgInventoryForm.ColumnHeadersDefaultCellStyle.Font = New Font("Century Gothic", 9)
        tsmPrint.Enabled = True
        picItem.SizeMode = PictureBoxSizeMode.Zoom

        dgInventoryForm.DefaultCellStyle.SelectionBackColor = Color.White
        dgInventoryForm.DefaultCellStyle.SelectionForeColor = Color.Black


        critVal()
        MessageBox.Show(lblCrit.Text, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
    End Sub

    Private Sub tsmExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmExit.Click
        Dim res As DialogResult = MessageBox.Show("Are you sure you want to exit?", "Exit", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If res = DialogResult.Yes Then
            logout()
            Me.Close()
        End If
    End Sub

    Private Sub tsmSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmSave.Click
        db = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\inventorydb.mdb")
        Dim confMsg As DialogResult
        If tsmCreate.Enabled = True And isUpdating = True Then
            confMsg = MessageBox.Show("Are you sure you want to add this item?", "Add Item", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If confMsg = DialogResult.Yes Then
                If checkInput() = True And isTextboxFull() = True Then
                    Dim tempItem As String = String.Empty
                    Try
                        db.Open()
                        dbcmd = New OleDb.OleDbCommand("SELECT [Barcode Number] FROM tbl_items WHERE [Barcode Number] like '" & txtBarcodeNumber.Text.Trim & "'", db)
                        Dim reader As OleDbDataReader = dbcmd.ExecuteReader
                        While reader.Read
                            tempItem = reader("Barcode Number").ToString
                        End While
                        If tempItem = String.Empty Then
                            dbcmd = New OleDb.OleDbCommand("INSERT INTO tbl_items([Barcode Number], [Item Name], [Item Description], [Selling Price], [Buying Price], [Quantity], [Critical Value]) VALUES ('" & txtBarcodeNumber.Text.Trim & "', '" & txtItemName.Text.Trim & "', '" & txtItemDescription.Text.Trim & "', '" & txtSellingPrice.Text.Trim & "', '" & txtBuyingPrice.Text.Trim & "', '" & txtQuantity.Text.Trim & "', '" & txtCriticalValue.Text.Trim & "')", db)
                            dbcmd.ExecuteNonQuery()
                            MessageBox.Show("Registration Successful", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            defaultInventoryForm()
                            clearInventoryForm()
                        Else
                            MessageBox.Show("Barcode Number is already registered to an item. Please enter another one.", "Invalid Barcode Number", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            txtBarcodeNumber.Clear()
                        End If
                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    Finally
                        db.Close()
                    End Try
                End If
            End If
        End If

        If tsmDelete.Enabled = True Then
            confMsg = MessageBox.Show("Are you sure you want to delete this item? This action cannot be undone!", "Delete Item", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If confMsg = DialogResult.Yes Then
                Try
                        db.Open()
                        Dim tempCmd As OleDb.OleDbCommand
                    tempCmd = New OleDb.OleDbCommand("INSERT INTO tbl_archiveItems([Barcode Number], [Item Name], [Item Description], [Selling Price], [Buying Price], [Quantity], [ArchDate]) VALUES ('" & txtBarcodeNumber.Text.Trim & "', '" & txtItemName.Text.Trim & "', '" & txtItemDescription.Text.Trim & "', '" & txtSellingPrice.Text.Trim & "', '" & txtBuyingPrice.Text.Trim & "', '" & txtQuantity.Text.Trim & "', '" & currentDate & "')", db)
                        tempCmd.ExecuteNonQuery()

                    dbcmd = New OleDb.OleDbCommand("DELETE FROM tbl_items WHERE [Barcode Number] = '" & txtBarcodeNumber.Text & "'", db)
                    dbcmd.ExecuteNonQuery()
                    MessageBox.Show("Item Deleted", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    clearInventoryForm()
                    defaultInventoryForm()
                Catch ex As Exception
                    MsgBox(ex.ToString)
                Finally
                    db.Close()
                End Try
            End If

        ElseIf tsmEdit.Enabled = True Then
            confMsg = MessageBox.Show("This action will update the item details. Do you want to continue?", "Update Item Details", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If confMsg = DialogResult.Yes Then
                If isTextboxFull() = True Then
                    Try
                        db.Open()
                        dbcmd = New OleDb.OleDbCommand("SELECT [Barcode Number] FROM tbl_items", db)
                        Dim reader As OleDbDataReader = dbcmd.ExecuteReader
                        Dim bcFound As Boolean = False
                        While reader.Read
                            Dim tempBC As String = reader("Barcode Number").ToString
                            If tempBC = txtBarcodeNumber.Text.Trim Then
                                bcFound = True
                            End If
                        End While

                        If bcFound = True Then
                            dbcmd = New OleDb.OleDbCommand("UPDATE tbl_items SET [Quantity] = @val1 WHERE [Barcode Number] = @id", db)
                            dbcmd.Parameters.AddWithValue("@val1", txtQuantity.Text)
                            dbcmd.Parameters.AddWithValue("@id", txtBarcodeNumber.Text)
                            dbcmd.ExecuteNonQuery()
                            clearInventoryForm()
                            defaultInventoryForm()
                            isTextBoxEnabled(False)
                            txtBarcodeNumber.Enabled = True
                            MessageBox.Show("Details Updated!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Else
                            MessageBox.Show("Barcode number not found!", "Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End If

                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    Finally
                        db.Close()
                    End Try
                End If
                defaultInventoryForm()
                clearInventoryForm()
                isTextBoxEnabled(False)
            End If
        End If
    End Sub

    Private Sub tsmCreate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmCreate.Click
        clearInventoryForm()
        lblDescription.Text = "Add Item"
        linkUploadPhoto.Visible = True
        clearMenuStrip()
        tsmCreate.Enabled = True
        tsmSave.Enabled = True
        tsmCancel.Enabled = True
        isTextBoxEnabled(True)
        isUpdating = True
    End Sub

    Private Sub tsmEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmEdit.Click
        lblDescription.Text = "Edit Item Details"
        clearMenuStrip()
        tsmCancel.Enabled = True
        tsmSave.Enabled = True
        tsmEdit.Enabled = True
        isTextBoxEnabled(False)
        txtBarcodeNumber.Enabled = True
        txtQuantity.Enabled = True
        isUpdating = True
    End Sub

    Private Sub tsmSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmSearch.Click
        lblDescription.Text = "Search Item"
        clearMenuStrip()
        tsmSearch.Enabled = True
        tsmCancel.Enabled = True
        isTextBoxEnabled(False)
        txtBarcodeNumber.Enabled = True
        Dim inp As String = InputBox("Enter the name of the item you want to search:", "Search Item")
        If inp.Length < 1 Then
        Else
            Try
                db.Open()
                dba = New OleDb.OleDbDataAdapter("SELECT * FROM tbl_items WHERE [Item Name] like '%" & inp.Trim & "%'", db)
                dbds = New DataSet
                dba.Fill(dbds, "tbl_items")
                dgInventoryForm.DataSource = dbds.Tables("tbl_items")
            Catch ex As Exception
                MsgBox(ex.ToString)
            Finally
                db.Close()
            End Try
        End If

    End Sub

    Private Sub tsmDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmDelete.Click
        lblDescription.Text = "Delete Item"
        clearMenuStrip()
        tsmCancel.Enabled = True
        tsmSave.Enabled = True
        tsmDelete.Enabled = True
        isTextBoxEnabled(False)
        txtBarcodeNumber.Enabled = True
        isUpdating = True
    End Sub

    Private Sub tsmCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmCancel.Click
        clearInventoryForm()
        defaultMenuStrip()
        defaultInventoryForm()

    End Sub


    Private Sub txtBarcodeNumber_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBarcodeNumber.TextChanged
        Try
            picItem.ImageLocation = ""
            picItem.Refresh()
            picItem.ImageLocation = Application.StartupPath & "\images\" & txtBarcodeNumber.Text.Trim & ".jpg"
            picItem.Load()
        Catch ex As Exception
            'Sana di mag-error hehe
        End Try

        If isUpdating = True Then
            If tsmSearch.Enabled = True Or tsmDelete.Enabled = True Or tsmEdit.Enabled = True Then
                If txtBarcodeNumber.TextLength <> 0 Then
                    Try
                        db = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\inventorydb.mdb")
                        db.Open()
                        dbcmd = New OleDb.OleDbCommand("SELECT * FROM tbl_items WHERE [Barcode Number] = '" & txtBarcodeNumber.Text & "'", db)

                        Dim reader As OleDb.OleDbDataReader = dbcmd.ExecuteReader
                        While reader.Read
                            txtItemName.Text = reader("Item Name").ToString
                            txtItemDescription.Text = reader("Item Description").ToString
                            txtBuyingPrice.Text = reader("Buying Price").ToString
                            txtSellingPrice.Text = reader("Selling Price").ToString
                            txtQuantity.Text = reader("Quantity").ToString
                            txtCriticalValue.Text = reader("Critical Value").ToString
                        End While

                    Catch ex As Exception

                    Finally
                        db.Close()
                    End Try

                Else

                    defaultInventoryForm()
                    clearInventoryForm()
                End If
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

    Private Sub btnWarning_mousehover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnWarning.MouseHover
        lblCrit.Visible = True
    End Sub

    Private Sub btnWarning_mouseleave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnWarning.MouseLeave
        lblCrit.Visible = False
    End Sub

    Private Sub btnWarning_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnWarning.Click
        MessageBox.Show(lblCrit.Text, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
    End Sub

    Private Sub linkUploadPhoto_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles linkUploadPhoto.LinkClicked
        If txtBarcodeNumber.Text.Trim.Length > 0 Then
            Dim ofdpic As New OpenFileDialog()
            ofdpic.FileName = ""
            If ofdpic.ShowDialog() = DialogResult.OK Then
                Try
                    Dim destFilePath As String = Application.StartupPath & "\images\" & Trim(txtBarcodeNumber.Text) & ".jpg"
                    My.Computer.FileSystem.CopyFile(ofdpic.FileName, destFilePath, True)
                    picItem.ImageLocation = destFilePath
                    picItem.Load()
                Catch err As Exception
                    MsgBox(err.ToString())
                End Try
            End If
        Else
            MsgBox("Enter your user ID", MsgBoxStyle.Information, "Warning")
        End If

    End Sub

    Private Sub ExportToExcel()
        If dgInventoryForm.RowCount > 0 Then
            Dim excelApp As New Excel.Application()

            Dim excelWorkbook As Excel.Workbook = excelApp.Workbooks.Add(Type.Missing)
            Dim excelWorksheet As Excel.Worksheet = excelWorkbook.Sheets(1)

            ' Set the column headers in Excel
            For i As Integer = 0 To dgInventoryForm.Columns.Count - 1
                excelWorksheet.Cells(1, i + 1) = dgInventoryForm.Columns(i).HeaderText
            Next

            ' Export data from DataGrid to Excel
            For i As Integer = 0 To dgInventoryForm.Rows.Count - 1
                For j As Integer = 0 To dgInventoryForm.Columns.Count - 1
                    If dgInventoryForm.Rows(i).Cells(j).Value IsNot Nothing Then
                        excelWorksheet.Cells(i + 2, j + 1) = dgInventoryForm.Rows(i).Cells(j).Value.ToString()
                    End If
                Next
            Next

            ' Save the Excel file
            Dim saveFileDialog As New SaveFileDialog()
            saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
            saveFileDialog.FileName = "The ISO Team Enterprise Inventory Report"
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


    Private Sub tsmExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmExcel.Click
        pnlStockReport.Visible = True
        pnlStockReport.Dock = DockStyle.Fill

        Dim stockReport As New ReportDocument

        stockReport.Load(Application.StartupPath & "\reports\stockReport.rpt")
        crepViewerStocks.ReportSource = stockReport
        crepViewerStocks.Refresh()

    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        pnlStockReport.Visible = False
    End Sub

    Private Sub SoftwareToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SoftwareToolStripMenuItem.Click
        Me.Hide()
        frmAbout.Show()
    End Sub
End Class