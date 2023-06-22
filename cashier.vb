Imports System.Data.OleDb
Imports CrystalDecisions.CrystalReports.Engine
Imports System.Windows.Forms

Public Class Cashier

    Public necc As New OleDbConnection
    Public neccAdap As New OleDbDataAdapter
    Public neccDS As New DataSet


    Dim subTotal As Double = 0
    Dim amountDue As Double = 0
    Dim tempAmtDue As Double = 0
    Dim discount As Double = 0
    Dim discountAmount As Double = 0
    Dim VAT As Double = 0
    Dim isDCash As Boolean = True


    Dim itemList As String
    Dim itemQuanti As String


    Sub loadDataGrid()

        Try
            necc = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\inventorydb.mdb")
            necc.Open()
            neccAdap = New OleDb.OleDbDataAdapter("SELECT [ITEM], [DESCRIPTION], [QUANTITY], [UNIT PRICE], [TOTAL] FROM tbl_temp", necc)
            neccDS = New DataSet
            neccAdap.Fill(neccDS, "tbl_temp")

            dbcmd = New OleDb.OleDbCommand("SELECT * FROM tbl_misc", necc)
            dbcmd.ExecuteNonQuery()

            Dim reader As OleDbDataReader = dbcmd.ExecuteReader()
            While reader.Read
                VAT = reader("VAT").ToString
            End While
                VAT = "0." & VAT
                VAT = Double.Parse(VAT)
            

        Catch ex As Exception
        Finally
            necc.Close()
        End Try

        dgSale.DataSource = neccDS.Tables("tbl_temp")

        Try
            necc = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\inventorydb.mdb")
            necc.Open()
            neccAdap = New OleDb.OleDbDataAdapter("SELECT * FROM tbl_temp", necc)
            neccDS = New DataSet
            neccAdap.Fill(neccDS, "tbl_temp")
        Catch ex As Exception
        Finally
            necc.Close()
        End Try
        dgTemp.DataSource = neccDS.Tables("tbl_temp")
    End Sub

    Sub dropTheHammer()
        Try
            necc = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\inventorydb.mdb")
            necc.Open()
            dbcmd = New OleDb.OleDbCommand("DELETE * FROM tbl_temp", necc)
            dbcmd.ExecuteNonQuery()

        Catch ex As Exception
        Finally
            necc.Close()
        End Try

        dgSale.DataSource = neccDS.Tables("tbl_temp")
    End Sub

    Sub buttonReset(ByVal bool)
        btnPaymentOption.Visible = bool
        btnDiscount.Visible = bool
        btnVoidTransaction.Visible = bool
        btnTaxOverride.Visible = bool
        btnMisc.Visible = bool
        btnProducts.Visible = bool
        btnHold.Enabled = bool
        btnNewOrder.Enabled = bool
        btnRefund.Enabled = bool
        btnSearchItem.Enabled = bool
    End Sub

    Private Sub Cashier_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtBarcode.Select()
        lblCurrUser.Text = currentUser
        Timer1.Start()

        dgSale.AllowUserToResizeColumns = False
        dgSale.AllowUserToResizeRows = False
        dgSale.DefaultCellStyle.SelectionBackColor = Color.Gold
        dgSale.DefaultCellStyle.SelectionForeColor = Color.Black

        dropTheHammer()
        loadDataGrid()

        dgSale.Columns(0).Width = 250
        dgSale.Columns(1).Width = 150
        dgSale.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgSale.Columns(2).Width = 100
        dgSale.Columns(3).Width = 100
        dgSale.Columns(4).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        dgSale.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

    End Sub

    Dim itemName As String
    Dim description As String
    Dim price As String
    Dim capital As String
    Dim quanti As String
    Dim barcode As String

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddItem.Click
        Dim itemFound As Boolean = False
        Dim enStocks As Boolean = False
        db = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\inventorydb.mdb")

        Try
            db.Open()
            dbcmd = New OleDb.OleDbCommand("SELECT [Barcode Number], [Quantity] FROM tbl_items", db)
            Dim reader As OleDbDataReader = dbcmd.ExecuteReader()
            While reader.Read
                barcode = reader("Barcode Number").ToString
                Dim qty As String = reader("Quantity").ToString
                If barcode = txtBarcode.Text.Trim Then
                    itemFound = True
                    If Val(txtQuantity.Text.Trim) <= Val(qty) Then
                        enStocks = True
                    Else
                        MessageBox.Show("The item with the barcode number " & barcode & " only have " & qty & " stocks left", "Failed", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If

                End If
            End While
        Catch ex As Exception

        Finally
            db.Close()
        End Try

        If itemFound = True And enStocks = True Then
            If Integer.TryParse(txtQuantity.Text, Nothing) Then
                Dim bCodeTemp As String = ""
                Dim newQty As String = ""
                Try
                    db.Open()
                    dbcmd = New OleDb.OleDbCommand("SELECT [BARCODE NUMBER] FROM tbl_temp", db)
                    Dim reader5 As OleDbDataReader = dbcmd.ExecuteReader()
                    While reader5.Read
                        Dim rak As String = ""
                        rak = reader5("BARCODE NUMBER").ToString
                        If rak = txtBarcode.Text Then
                            bCodeTemp = txtBarcode.Text
                        End If
                    End While
                Catch ex As Exception
                    MsgBox(ex.ToString)
                Finally
                    db.Close()
                End Try

                Try
                    db.Open()
                    dbcmd = New OleDb.OleDbCommand("SELECT [Item Name], [Item Description], [Selling Price], [Buying Price], [Quantity] FROM tbl_items WHERE [Barcode Number] = '" & txtBarcode.Text.Trim & "'", db)
                    Dim reader As OleDbDataReader = dbcmd.ExecuteReader()
                    While reader.Read
                        itemName = reader("Item Name").ToString()
                        description = reader("Item Description").ToString
                        price = reader("Selling Price").ToString()
                        quanti = reader("Quantity").ToString
                        capital = reader("Buying Price").ToString
                    End While
                    reader.Close()

                    Dim total As Double = Val(txtQuantity.Text) * Val(price)
                    total = Math.Round(total, 2, MidpointRounding.AwayFromZero)
                    Dim bPrice As Double = Val(txtQuantity.Text) * Val(capital)
                    bPrice = Math.Round(bPrice, 2, MidpointRounding.AwayFromZero)
                    Dim profit As Double = total - bPrice
                    profit = Math.Round(profit, 2, MidpointRounding.AwayFromZero)


                    If Val(quanti) > 0 Then

                        If bCodeTemp IsNot String.Empty And bCodeTemp = txtBarcode.Text.Trim Then
                            Dim res1 As DialogResult = MessageBox.Show("The item is already listed, do you want to update the quantity", "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

                            If res1 = DialogResult.Yes Then
                                Dim newBP As String = ""
                                Dim newPrft As String = ""
                                Dim newTotal As String = ""

                                dbcmd = New OleDb.OleDbCommand("SELECT [QUANTITY], [BUYING PRICE], [PROFIT], [TOTAL] FROM tbl_temp WHERE [BARCODE NUMBER] = '" & txtBarcode.Text.Trim & "'", db)
                                Dim reader1 As OleDbDataReader = dbcmd.ExecuteReader()
                                While reader1.Read
                                    newQty = reader1("QUANTITY").ToString
                                    newBP = reader1("BUYING PRICE").ToString
                                    newPrft = reader1("PROFIT").ToString
                                    newTotal = reader1("TOTAL").ToString
                                End While

                                newPrft = Val(profit) + Val(newPrft)
                                newBP = Val(bPrice) + Val(newBP)
                                newTotal = Val(total) + Val(newTotal)
                                newQty = Val(newQty) + Val(txtQuantity.Text.Trim)

                                dbcmd = New OleDb.OleDbCommand("UPDATE tbl_temp SET [QUANTITY] = '" & newQty & "', [BUYING PRICE] = '" & newBP & "', [PROFIT] = '" & newPrft & "', [TOTAL] = '" & newTotal & "' WHERE [BARCODE NUMBER] = '" & txtBarcode.Text.Trim & "'", db)
                                dbcmd.ExecuteNonQuery()
                            End If

                        Else
                            dbcmd = New OleDb.OleDbCommand("INSERT INTO tbl_temp([ITEM], [DESCRIPTION], [QUANTITY], [UNIT PRICE], [TOTAL], [BARCODE NUMBER], [BUYING PRICE], [PROFIT]) VALUES ('" & itemName & "', '" & description & "', '" & txtQuantity.Text & "', '" & price & "', '" & total & "', '" & txtBarcode.Text.Trim & "', '" & bPrice & "', '" & profit & "')", db)
                            dbcmd.ExecuteNonQuery()
                        End If

                        loadDataGrid()
                        itemList &= itemName & " | "
                        itemQuanti &= txtQuantity.Text & " | "

                        If isDCash = False Then
                            tempAmtDue = tempAmtDue + total
                            discountAmount = discountAmount + (total * discount)

                        Else
                            tempAmtDue = tempAmtDue + total
                            discountAmount = discount
                        End If

                        amountDue = tempAmtDue - discountAmount

                        VAT = amountDue * VAT
                        VAT = Math.Round(VAT, 2, MidpointRounding.AwayFromZero)
                        If taxExempt = True Then
                            subTotal = amountDue - VAT
                            subTotal = Math.Round(subTotal, 2, MidpointRounding.AwayFromZero)
                            amountDue = subTotal
                        Else
                            subTotal = amountDue - VAT
                            subTotal = Math.Round(subTotal, 2, MidpointRounding.AwayFromZero)
                        End If

                        lblSubTotal.Text = subTotal
                        lblSubTotal.Text = lblSubTotal.Text

                        lblTax.Text = VAT
                        lblTax.Text = lblTax.Text
                        lblAmountDue.Text = amountDue
                        lblAmountDue.Text = lblAmountDue.Text
                        lblDiscount.Text = discountAmount
                        lblDiscount.Text = lblDiscount.Text

                        txtBarcode.Clear()
                        txtQuantity.Text = 1
                        txtItemName.Clear()
                    Else
                        MessageBox.Show("Item with barcode number " & txtBarcode.Text & " is out of stock", "Out of Stock", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If

                Catch ex As Exception
                    MsgBox(ex.ToString)
                Finally
                    db.Close()
                End Try

            Else
                MessageBox.Show("You can only input numbers on the quantity textbox", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtQuantity.Text = 1
            End If
        ElseIf itemFound = False And enStocks = False Then
            MessageBox.Show("There is no item with the barcode number: " & txtBarcode.Text & ". Please try again and make sure to avoid any typing error", "No Item", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtBarcode.Clear()
            txtQuantity.Text = "1"
        End If
        loadDataGrid()
    End Sub

    Private Sub btnNewOrder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNewOrder.Click
        txtItemName.Clear()
        txtBarcode.Clear()
        txtPayment.Clear()
        txtQuantity.Text = 1

        subTotal = 0
        amountDue = 0
        tempAmtDue = 0
        discount = 0
        discountAmount = 0
        VAT = 0

        lblItemName.Text = ""
        lblPrice.Text = ""

        subTotal = 0
        lblSubTotal.Text = subTotal
        amountDue = 0
        lblAmountDue.Text = amountDue
        discount = 0
        lblDiscount.Text = discount
        VAT = 0
        lblTax.Text = VAT

        taxExempt = False
        discountClicked = False
        isPaymentCash = True

        itemList = ""
        itemQuanti = ""

        pnlMsg.Visible = False
        dropTheHammer()
        loadDataGrid()

    End Sub

    Private Sub btnIncrement_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnIncrement.Click
        txtQuantity.Text = Val(txtQuantity.Text) + 1
    End Sub

    Private Sub btnDecrement_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDecrement.Click
        If Val(txtQuantity.Text) > 1 Then
            txtQuantity.Text = Val(txtQuantity.Text) - 1
        End If
    End Sub


    Dim discountClicked As Boolean = False
    Private Sub btnDiscountEnable_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDiscountEnable.Click

        Dim res As DialogResult = MessageBox.Show("Are you sure you want to apply this discount?", "Discount", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If res = DialogResult.Yes Then
            If radCash.Checked = True Then
                discount = Double.Parse(txtDiscountAmount.Text)
                isDCash = True
            ElseIf radPercent.Checked = True Then
                txtDiscountAmount.Text = "0." & txtDiscountAmount.Text
                discount = Double.Parse(txtDiscountAmount.Text)
                isDCash = False
            End If
            buttonReset(True)
            discountClicked = True
        Else
            isDCash = False
            pnlDiscount.Visible = False
        End If
        pnlDiscount.Visible = False
    End Sub


    Private Sub btnDiscount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDiscount.Click
        If discountClicked = True Then
            MessageBox.Show("You already applied a discount!", "Failed", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            radCash.Checked = True
            txtDiscountAmount.Clear()
            pnlDiscount.Visible = True
            pnlDiscount.Dock = DockStyle.Bottom
            buttonReset(False)
        End If
    End Sub

    Private Sub btnDiscountCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDiscountCancel.Click
        pnlDiscount.Visible = False
        buttonReset(True)
    End Sub

    Dim isPaymentCash = True
    Private Sub btnPaymentOption_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPaymentOption.Click
        radPaymentCash.Checked = True
        pnlPaymentOption.Visible = True
        pnlPaymentOption.Dock = DockStyle.Bottom
        buttonReset(False)
    End Sub

    Private Sub btnPOCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPOCancel.Click
        pnlPaymentOption.Visible = False
        buttonReset(True)
    End Sub

    Private Sub radPaymentCash_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles radPaymentCash.CheckedChanged
        isPaymentCash = True
    End Sub

    Private Sub radPaymentGCash_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles radPaymentGCash.CheckedChanged
        isPaymentCash = False
    End Sub

    Private Sub btnChangeVAT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnChangeVAT.Click
        If isAdmin = True Then
            Dim res As DialogResult = MessageBox.Show("Are you sure you want to continue? This action might have penalty from the Philippine Law", "Change VAT", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If res = DialogResult.Yes Then
                db = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\inventorydb.mdb")
                db.Open()
                Dim s As String = "PESO"
                dbcmd = New OleDb.OleDbCommand("UPDATE tbl_misc SET [VAT] = '" & txtVAT.Text & "' WHERE [SPC DISC TYPE] = '" & s & "'", db)
                dbcmd.ExecuteNonQuery()
                MessageBox.Show("The Value Added Tax has been changed to " & txtVAT.Text, "Change VAT Success", MessageBoxButtons.OK, MessageBoxIcon.Question)
                btnTaxExemptCancel.PerformClick()
                txtVAT.Clear()
            End If
        Else
            MessageBox.Show("This action can only be done by an administrator", "Change VAT", MessageBoxButtons.OK, MessageBoxIcon.Question)
        End If
        db.Close()
    End Sub

    Dim taxExempt As Boolean = False
    Private Sub btnTaxExemption_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim res As DialogResult = MessageBox.Show("Are you sure you want to continue? This action might have penalty to the Philippine Law", "Tax Exemption", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If res = DialogResult.Yes Then
            taxExempt = True
        End If
    End Sub

    Private Sub btnTaxOverride_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTaxOverride.Click
        pnlTaxExemption.Visible = True
        pnlTaxExemption.Dock = DockStyle.Bottom
        buttonReset(False)
    End Sub

    Private Sub btnTaxExemptCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTaxExemptCancel.Click
        pnlTaxExemption.Visible = False
        buttonReset(True)
    End Sub


    Private Sub btnProducts_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnProducts.Click
        pnlProducts.Visible = True
        buttonReset(False)
        pnlProducts.Dock = DockStyle.Bottom

        Try
            db = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\inventorydb.mdb")
            db.Open()
            dba = New OleDb.OleDbDataAdapter("SELECT [Item Name], [Item Description] FROM tbl_items", db)
            dbds = New DataSet
            dba.Fill(dbds, "tbl_items")
            dgProducts.DataSource = dbds.Tables("tbl_items")
            dgProducts.DefaultCellStyle.Font = New Font("Century Gothic", 8)
        Catch ex As Exception

        Finally
            db.Close()
        End Try
    End Sub

    Private Sub btnProductCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnProductCancel.Click
        pnlProducts.Visible = False
        buttonReset(True)
    End Sub

    Dim isSearching As Boolean = True
    Private Sub btnSearchItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearchItem.Click
        If isSearching = True Then
            pnlSearchItem.Visible = True
            pnlSearchItem.Dock = DockStyle.Bottom
            btnSearchItem.Text = "CANCEL SEARCH"
            btnSearchItem.BackColor = Color.Gainsboro
            isSearching = False
            buttonReset(False)
            btnSearchItem.Enabled = True
        Else
            pnlSearchItem.Visible = False
            btnSearchItem.Text = "SEARCH ITEM"
            btnSearchItem.BackColor = Color.Gold
            isSearching = True
            buttonReset(True)
        End If
    End Sub

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        db = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\inventorydb.mdb")
        Try
            db.Open()
            If txtSearchBarcode.Text.Trim IsNot String.Empty Then
                dbcmd = New OleDb.OleDbCommand("SELECT [Barcode Number], [Item Name], [Item Description], [Selling Price] FROM tbl_items WHERE [Barcode Number] = '" & txtSearchBarcode.Text.Trim & "'", db)

            ElseIf txtSearchName.Text.Trim IsNot String.Empty Then
                dbcmd = New OleDb.OleDbCommand("SELECT [Barcode Number], [Item Name], [Item Description], [Selling Price] FROM tbl_items WHERE [Item Name] = '" & txtSearchName.Text.Trim & "'", db)
            End If

            Dim reader As OleDbDataReader = dbcmd.ExecuteReader()
            While reader.Read
                txtSearchName.Text = reader("Item Name").ToString()
                txtSearchDescrip.Text = reader("Item Description").ToString
                txtSearchPrice.Text = reader("Selling Price").ToString()
                txtSearchBarcode.Text = reader("Barcode Number").ToString
            End While
            reader.Close()

        Catch ex As Exception
            MessageBox.Show("Item not found", "Invalid input", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Finally
            db.Close()
        End Try
    End Sub

    Private Sub btnLogout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogout.Click
        Dim res As DialogResult = MessageBox.Show("You want to logout?", "LOGOUT", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If res = DialogResult.Yes Then
            logout()
            Me.Hide()
            login.Show()
            currentUser = ""
        End If
    End Sub

    Dim refund As Boolean = False
    Private Sub btnRefund_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefund.Click
        If refund = False Then
            pnlRefund.Visible = True
            pnlRefund.Dock = DockStyle.Bottom
            btnRefund.Text = "CANCEL REFUND"
            btnRefund.BackColor = Color.Gainsboro
            refund = True
            buttonReset(False)
            btnRefund.Enabled = True
        Else
            pnlRefund.Visible = False
            btnRefund.Text = "REFUND"
            btnRefund.BackColor = Color.Gold
            refund = False
            txtRefundBarcode.Clear()
            txtRefundName.Clear()
            txtRefundQuantity.Clear()
            txtRefundAmount.Clear()
            buttonReset(True)
        End If
    End Sub

    Private Sub btnRefundRefund_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefundRefund.Click
        If txtRefundName.Text IsNot String.Empty Then
            Dim res As DialogResult = MessageBox.Show("You are about to refund " & txtRefundName.Text & ". Are you sure you want to continue?", "Refund", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If res = DialogResult.Yes Then
                Dim currentTime As String = currentDateAndTime.ToString("h:mm tt")
                db = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\inventorydb.mdb")
                Try
                    db.Open()
                    dbcmd = New OleDb.OleDbCommand("INSERT INTO tbl_refund([Barcode Number], [Refund Amount], [Item Name], [Refund Date], [Refund Time], [Cashier], [Quantity]) VALUES ('" & txtRefundBarcode.Text.Trim & "', '" & txtRefundAmount.Text.Trim & "', '" & txtRefundName.Text.Trim & "', '" & currentDate & "', '" & currentTime & "','" & currentUserName & "', '" & txtRefundQuantity.Text.Trim & "')", db)
                    dbcmd.ExecuteNonQuery()
                    MessageBox.Show("Refund successful", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As Exception
                    MsgBox(ex.ToString)
                Finally
                    db.Close()
                End Try
                btnRefund.PerformClick()
            End If
        Else
            MessageBox.Show("Invalid Barcode Number. Please enter a valid one", "Invalid Barcode", MessageBoxButtons.OK, MessageBoxIcon.Error)
            btnRefund.PerformClick()
        End If
    End Sub


    Private Sub txtRefundBarcode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRefundBarcode.TextChanged
        txtRefundName.Clear()
        db = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\inventorydb.mdb")
        Try
            db.Open()
            dbcmd = New OleDb.OleDbCommand("SELECT [Item Name] FROM tbl_items WHERE [Barcode Number] = '" & txtRefundBarcode.Text.Trim & "'", db)
            Dim reader As OleDb.OleDbDataReader = dbcmd.ExecuteReader
            While reader.Read
                txtRefundName.Text = reader("Item Name").ToString
            End While
            reader.Close()
        Catch ex As Exception
        Finally
            db.Close()
        End Try
    End Sub

    Private Sub btnPayment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPayment.Click
        Dim number As Double
        Dim currentTime As String = currentDateAndTime.ToString("h:mm tt")

        lblMsgSubtotal.Text = subTotal
        lblMsgDiscount.Text = discountAmount
        If amountDue = 0 Then
            lblMsgTax.Text = "0"
        Else
            lblMsgTax.Text = lblTax.Text
        End If

        lblMsgAmountDue.Text = amountDue
        lblMsgPaymentDue.Text = txtPayment.Text
        lblMsgChangeDue.Text = Val(lblMsgPaymentDue.Text.Trim) - amountDue
        lblMsgChangeDue.Text = Math.Round(Val(lblMsgChangeDue.Text), 2, MidpointRounding.AwayFromZero)

        lblMsgCashier.Text = currentUserName
        lblMsgDate.Text = currentDate
        lblMsgTime.Text = currentTime
        pnlMsg.Visible = True

        If Val(txtPayment.Text) < amountDue Or Not Double.TryParse(txtPayment.Text, number) Then
            pnlMsg.Visible = False
            MessageBox.Show("Invalid payment input. Please recheck and try again", "Invalid Payment", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub btnMsgClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMsgClose.Click

    End Sub

    Private Sub btnMsgConfirm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMsgConfirm.Click
        Dim currentTime As String = currentDateAndTime.ToString("h:mm tt")
        db = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\inventorydb.mdb")

        Try
            db.Open()
            dbcmd = New OleDb.OleDbCommand("INSERT INTO tbl_transactions([Item Purchased], [Quantity], [Subtotal], [Discount], [Tax], [Total], [Cashier], [Trans Date], [Trans Time]) VALUES ('" & itemList & "', '" & itemQuanti & "', '" & lblMsgSubtotal.Text & "', '" & lblMsgDiscount.Text & "', '" & lblMsgTax.Text & "', '" & lblAmountDue.Text & "', '" & currentUserName & "', '" & currentDate & "', '" & currentTime & "')", db)
            dbcmd.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox(ex.ToString)
        Finally
            db.Close()
        End Try

        Try
            db.Open()
            For Each row As DataGridViewRow In dgTemp.Rows
                If currentUser = String.Empty Then
                    currentUser = "NA"
                End If

                Dim q1 As String = row.Cells("BARCODE NUMBER").Value.ToString()
                Dim q2 As String = row.Cells("ITEM").Value.ToString()
                Dim q3 As String = row.Cells("DESCRIPTION").Value.ToString()
                Dim q4 As String = row.Cells("QUANTITY").Value.ToString()
                Dim q5 As String = row.Cells("TOTAL").Value.ToString()
                Dim q6 As String = row.Cells("BUYING PRICE").Value.ToString()
                Dim q7 As String = row.Cells("PROFIT").Value.ToString()
                Dim q8 As String = currentTime
                Dim q9 As String = currentDate
                Dim q10 As String = currentUser

                Dim insertQuery As String = "INSERT INTO tbl_sales([Barcode Number], [Item Name], [Item Description], [Quantity], [Price], [Capital], [Profit], [Sales Time], [Sales Date], [Employee]) VALUES (q1, q2, q3, q4, q5, q6, q7, q8, q9, q10)"
                dbcmd = New OleDbCommand(insertQuery, db)

                dbcmd.Parameters.AddWithValue("q1", q1)
                dbcmd.Parameters.AddWithValue("q2", q2)
                dbcmd.Parameters.AddWithValue("q3", q3)
                dbcmd.Parameters.AddWithValue("q4", q4)
                dbcmd.Parameters.AddWithValue("q5", q5)
                dbcmd.Parameters.AddWithValue("q6", q6)
                dbcmd.Parameters.AddWithValue("q7", q7)
                dbcmd.Parameters.AddWithValue("q8", q8)
                dbcmd.Parameters.AddWithValue("q9", q9)
                dbcmd.Parameters.AddWithValue("q10", q10)
                dbcmd.ExecuteNonQuery()

            Next

        Catch ex As Exception
        Finally
            db.Close()
        End Try

        Try
            db.Open()
            dbcmd = New OleDb.OleDbCommand("SELECT [BARCODE NUMBER], [QUANTITY] FROM tbl_temp", db)
            Dim reader As OleDbDataReader = dbcmd.ExecuteReader

            While reader.Read()
                Dim bcode As String = reader("BARCODE NUMBER").ToString
                Dim itemQuantity As String = reader("QUANTITY").ToString

                Dim cmd As OleDb.OleDbCommand
                cmd = New OleDb.OleDbCommand("SELECT [Quantity] FROM tbl_items WHERE [Barcode Number] = '" & bcode & "'", db)

                Dim reader2 As OleDbDataReader = cmd.ExecuteReader
                While reader2.Read

                    Dim stockQuantity As String = reader2("Quantity").ToString
                    Dim newQuantity As Integer = Val(stockQuantity) - Val(itemQuantity)
                    Dim tempcmd As OleDb.OleDbCommand
                    tempcmd = New OleDb.OleDbCommand("UPDATE tbl_items SET [Quantity] = @val1 WHERE [Barcode Number] = @id", db)
                    tempcmd.Parameters.AddWithValue("@val1", newQuantity)
                    tempcmd.Parameters.AddWithValue("@id", bcode)
                    tempcmd.ExecuteNonQuery()

                End While
            End While

        Catch ex As Exception
            MsgBox(ex.ToString)
        Finally
            db.Close()
        End Try

        Dim res As DialogResult = MessageBox.Show("Do you want to print receipt?", "Print receipt", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If res = DialogResult.Yes Then
            Timer1.Stop()
            Dim currentReceiptNumber As String = 0
            db = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\inventorydb.mdb")
            Try
                db.Open()
                dbcmd = New OleDb.OleDbCommand("SELECT [Receipt Number] FROM tbl_transactions", db)
                Dim reader As OleDbDataReader = dbcmd.ExecuteReader
                While reader.Read
                    currentReceiptNumber = reader("Receipt Number").ToString
                End While
            Catch ex As Exception
                MsgBox(ex.ToString)
            Finally
                db.Close()
            End Try

            If currentUserName = String.Empty Then
                currentUserName = "DEFAULT"
            Else

            End If

            Try
                c.Load(Application.StartupPath & "\reports\receiptF.rpt")
                c.SetParameterValue("rcSubtotal", lblMsgSubtotal.Text)
                c.SetParameterValue("rcDiscount", lblMsgDiscount.Text)
                c.SetParameterValue("rcVAT", lblMsgTax.Text)
                c.SetParameterValue("rcAmount", lblMsgAmountDue.Text)
                c.SetParameterValue("rcCash", lblMsgPaymentDue.Text)
                c.SetParameterValue("rcChange", lblMsgChangeDue.Text)
                c.SetParameterValue("rcRecNumber", currentReceiptNumber)
                c.SetParameterValue("rcSalesPerson", currentUserName)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

            frmReceipt.Show()

        Else
            btnNewOrder.PerformClick()
        End If
        Timer1.Start()
    End Sub

  
    Private Sub txtVoidConfirm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtVoidConfirm.Click
        Dim res As DialogResult = MessageBox.Show("You are about to void the transaction with receipt number: " & txtVoidReceipt.Text, "Void Transaction", MessageBoxButtons.OKCancel, MessageBoxIcon.Information)
        If res = DialogResult.OK Then
            Try
                db = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\inventorydb.mdb")
                db.Open()

                Dim fquery As String = "SELECT * FROM tbl_transactions WHERE [Receipt Number] = @KeyValue"

                Using command As New OleDbCommand(fquery, db)
                    command.Parameters.AddWithValue("@KeyValue", Val(txtVoidReceipt.Text))

                    Dim reader As OleDbDataReader = command.ExecuteReader()
                    While reader.Read()
                        Dim a As String = reader("Receipt Number").ToString
                        Dim b As String = reader("Item Purchased").ToString
                        Dim c As String = reader("Total").ToString
                        Dim d As String = reader("Trans Date").ToString

                        Dim cmdtemp As OleDb.OleDbCommand
                        cmdtemp = New OleDb.OleDbCommand("INSERT INTO tbl_voidsales ([Receipt], [Items], [Total], [Date Purchased], [Reason], [Date Voided], [Employee]) VALUES ('" & a & "', '" & b & "',  '" & c & "',  '" & d & "', '" & txtVoidReason.Text & "', '" & currentDate & "', '" & currentUser & "')", db)
                        cmdtemp.ExecuteNonQuery()
                        MsgBox("X")
                    End While
                End Using

                Dim dquery As String = "DELETE * FROM tbl_transactions WHERE [Receipt Number] = @KeyValue"

                Using command As New OleDbCommand(dquery, db)
                    command.Parameters.AddWithValue("@KeyValue", Val(txtVoidReceipt.Text))
                    command.ExecuteNonQuery()
                End Using

                buttonReset(True)
                MessageBox.Show("Transaction Voided", "Successful", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                MsgBox(ex.ToString)
                MessageBox.Show("Receipt Number not found!", "Unsuccessful", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End Try
        End If
    End Sub

    Private Sub btnVoidTransaction_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnVoidTransaction.Click
        pnlVoidTransaction.Visible = True
        pnlVoidTransaction.Dock = DockStyle.Bottom
        buttonReset(False)
    End Sub

    Private Sub txtVoidCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtVoidCancel.Click
        pnlVoidTransaction.Visible = False
        buttonReset(True)
    End Sub

    Private Sub btnMisc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMisc.Click
        pnlMisc.Visible = True
        pnlMisc.Dock = DockStyle.Bottom
        buttonReset(False)
    End Sub

    Private Sub txtMiscCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMiscCancel.Click
        pnlMisc.Visible = False
        buttonReset(True)
    End Sub

    Private Sub btnMiscConfirm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMiscConfirm.Click
        db = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\inventorydb.mdb")
        Dim inp As String = txtMiscAmount.Text.Trim
        Dim d As Double
        Dim currentTime As String = currentDateAndTime.ToString("h:mm tt")
        If Double.TryParse(inp, d) Then
            Dim res As DialogResult = MessageBox.Show("You are about to take " & inp & " pesos for " & txtMiscReason.Text.Trim & " reason. Are you sure you want to continue?", "Input error", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If res = DialogResult.Yes Then
                Try
                    db.Open()
                    dbcmd = New OleDb.OleDbCommand("INSERT INTO tbl_miscfee([Employee], [Amount], [Reason], [Misc_date], [Misc_time]) VALUES ('" & currentUserName & "', '" & txtMiscAmount.Text.Trim & "', '" & txtMiscReason.Text.Trim & "', '" & currentDate & "', '" & currentTime & "')", db)
                    dbcmd.ExecuteNonQuery()
                    MessageBox.Show("Action Successful", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    buttonReset(True)
                    txtMiscAmount.Clear()
                    txtMiscReason.Clear()
                    pnlMisc.Visible = False
                Catch ex As Exception
                    MsgBox(ex.ToString)
                Finally
                    db.Close()
            End Try
        Else
            txtMiscAmount.Clear()
            txtMiscReason.Clear()
            pnlMisc.Visible = False
        End If
        Else
        MessageBox.Show("Please enter proper amount", "Input error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End If
    End Sub

    Private Sub txtBarcode_PreviewKeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.PreviewKeyDownEventArgs) Handles txtBarcode.PreviewKeyDown
        If e.KeyCode = Keys.Enter Then
            btnAddItem.PerformClick()
        else If e.KeyCode = Keys.Tab Then
        e.IsInputKey = True
        txtPayment.Focus()
        ElseIf e.KeyCode = Keys.Up Then
        btnIncrement.PerformClick()
        ElseIf e.KeyCode = Keys.Down Then
        btnDecrement.PerformClick()
        ElseIf e.KeyCode = Keys.Back Then
        lblItemName.Text = ""
        lblPrice.Text = ""
        End If
    End Sub

    Private Sub txtPayment_PreviewKeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.PreviewKeyDownEventArgs) Handles txtPayment.PreviewKeyDown
        If pnlMsg.Visible = False Then
            If e.KeyCode = Keys.Tab Then
                e.IsInputKey = True
                txtBarcode.Focus()
            ElseIf e.KeyCode = Keys.Enter Then
                btnPayment.PerformClick()
            End If
        End If

        If e.KeyCode = Keys.Space And pnlMsg.Visible = True Then
            btnMsgConfirm.PerformClick()
        End If

    End Sub

    Private Sub tsmDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmDelete.Click
        Dim res As DialogResult = MessageBox.Show("Are you sure you want to logout?", "LOGOUT", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If res = DialogResult.Yes Then
            logout()
            Me.Hide()
            login.Show()
        End If
    End Sub

    Private Sub ReturnToMainMenuToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReturnToMainMenuToolStripMenuItem.Click
        Me.Hide()
        main.Show()
    End Sub

    Private Sub tsmExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tsmExit.Click
        Dim res As DialogResult = MessageBox.Show("Are you sure you want to exit?", "EXIT", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If res = DialogResult.Yes Then
            logout()
            Me.Close()
        End If
    End Sub

    Private Sub txtBarcode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBarcode.TextChanged

        lblPrice.Text = ""
        lblItemName.Text = ""

        db = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\inventorydb.mdb")
        Try
            db.Open()
            dbcmd = New OleDb.OleDbCommand("SELECT [Item Name], [Selling Price] FROM tbl_items WHERE [Barcode Number] = '" & txtBarcode.Text.Trim & "'", db)
            Dim reader As OleDb.OleDbDataReader = dbcmd.ExecuteReader


            While reader.Read
                lblItemName.Text = reader("Item Name").ToString
                lblPrice.Text = reader("Selling Price").ToString
            End While

        Catch ex As Exception

        Finally
            db.Close()
        End Try

        loadDataGrid()
    End Sub

    Private Sub SoftwareToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SoftwareToolStripMenuItem.Click
        Me.Hide()
        frmAbout.Show()
    End Sub

    Dim isHolding As Boolean = False
    Private Sub btnHold_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHold.Click

        If isHolding = False Then
            If dgSale.RowCount > 1 Then
                Dim res As DialogResult = MessageBox.Show("Are you sure you want to hold this transaction?", "Hold", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

                If res = DialogResult.Yes Then
                    db = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\inventorydb.mdb")
                    Try
                        db.Open()

                        dbcmd = New OleDb.OleDbCommand("DELETE * FROM tbl_hold", db)
                        dbcmd.ExecuteNonQuery()

                        dbcmd = New OleDb.OleDbCommand("INSERT INTO tbl_hold SELECT * FROM tbl_temp", db)
                        dbcmd.ExecuteNonQuery()

                        dropTheHammer()
                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    Finally
                        db.Close()
                    End Try

                    btnHold.Text = "CONTINUE TRANSACTION"
                    btnHold.BackColor = Color.Goldenrod
                    isHolding = True
                End If
            Else
                MessageBox.Show("Transaction is empty. Action failed.", "Invalid", MessageBoxButtons.OK, MessageBoxIcon.Question)
            End If
        Else
            Dim res As DialogResult = MessageBox.Show("Do you want to continue that transaction on hold?", "Hold", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If res = DialogResult.Yes Then
                dropTheHammer()
                Try
                    db.Open()

                    dbcmd = New OleDb.OleDbCommand("INSERT INTO tbl_temp SELECT * FROM tbl_hold", db)
                    dbcmd.ExecuteNonQuery()

                    dbcmd = New OleDb.OleDbCommand("DELETE * FROM tbl_hold", db)
                    dbcmd.ExecuteNonQuery()

                Catch ex As Exception
                    MsgBox(ex.ToString)
                Finally
                    db.Close()
                End Try

                btnHold.Text = "HOLD TRANSACTION"
                btnHold.BackColor = Color.Gold
                isHolding = False
            End If
        End If
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        lblTime.Text = DateTime.Now.ToString("hh:mm tt")
        loadDataGrid()
    End Sub

End Class