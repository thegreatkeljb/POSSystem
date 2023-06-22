Imports System.Data.OleDb
Imports System.Windows.Forms.DataVisualization.Charting
Imports Excel = Microsoft.Office.Interop.Excel
Imports CrystalDecisions.CrystalReports.Engine

Public Class reports

    Dim cmd As New OleDb.OleDbCommand


    'Sales Report
    Dim salesReport As Boolean = False
    Dim salesSpecify As Boolean = False
    'Payment Report
    Dim paymentReport As Boolean = False
    'Employee Report
    Dim employeeReport As Boolean = False
    Dim employeeReportSelector As Integer = 0
    'Iventory Report
    Dim inventoryReport As Boolean = False
    'Trans Log
    Dim transactionLog As Boolean = False



    Sub resetForm()
        dropTheAkronHammer()
        dgReports.Visible = False

        'Sales Report
        pnlPrintSalesReport.Visible = False
        pnlReturn.Visible = False
        pnlSalesReport.Visible = False
        pnlMenuSalesReport.Visible = False
        salesReport = False
        pnlRevenue.Visible = False
        btnMenuSales.BackColor = Color.Beige
        btnMenuGrossPr.BackColor = Color.Beige
        btnSalesReport.BackColor = Color.LightGray
        'End Sales Report

        'Payment Report
        pnlPaymentRep.Visible = False
        pnlPaymentReport.Visible = False
        paymentReport = False
        btnPaymentReport.BackColor = Color.LightGray
        'End Payment Report

        'Employee Report
        pnlEmployeeReport.Visible = False
        employeeReport = False
        btnEmployeeReport.BackColor = Color.LightGray
        'End Employee Report

        'Inventory Report
        inventoryReport = False
        pnlInventory.Visible = False
        btnInventoryReport.BackColor = Color.LightGray
        isInventoryArchive = False
        btnInventoryArchive.BackColor = Color.Gold
        btnInventoryArchive.Text = "INVENTORY ARCHIVE"
        'End Inventory Report

        'Transaction Log
        pnlPrintTransaction.Visible = False
        pnlPrintDTR.Visible = False
        transactionLog = False
        pnlTransactionLog.Visible = False
        btnTransactionLog.BackColor = Color.LightGray
        'End Transaction Log

    End Sub

    Sub dropTheAkronHammer()
        db = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\inventorydb.mdb")
        Try
            db.Open()
            dbcmd = New OleDb.OleDbCommand("DELETE * FROM tbl_reportsTemp", db)
            dbcmd.ExecuteNonQuery()
            dbcmd = New OleDb.OleDbCommand("DELETE * FROM tbl_temp", db)
            dbcmd.ExecuteNonQuery()
        Catch ex As Exception
        Finally
            db.Close()
        End Try
    End Sub

    Sub kwak(ByVal bcode, ByVal name, ByVal desc, ByVal quanti, ByVal pr, ByVal cap, ByVal prof, ByVal st, ByVal sd, ByVal e, ByVal tax, ByVal receipt, ByVal discount)
        If name = 1 Then
            Dim c1 As DataGridViewColumn = dgReports.Columns("Item Name")
            c1.HeaderText = "NAME"
        End If
        If desc = 1 Then
            Dim c2 As DataGridViewColumn = dgReports.Columns("Item Description")
            c2.HeaderText = "DESCRIPTION"
        End If
        If quanti = 1 Then
            Dim c3 As DataGridViewColumn = dgReports.Columns("Quantity")
            c3.HeaderText = "QTY"
        End If
        If pr = 1 Then
            Dim c4 As DataGridViewColumn = dgReports.Columns("Price")
            c4.HeaderText = "REVENUE"
        End If
        If cap = 1 Then
            Dim c5 As DataGridViewColumn = dgReports.Columns("Capital")
            c5.HeaderText = "COST"
        End If
        If prof = 1 Then
            Dim c6 As DataGridViewColumn = dgReports.Columns("Profit")
            c6.HeaderText = "PROFIT"
        End If
        If st = 1 Then
            Dim c7 As DataGridViewColumn = dgReports.Columns("Sales Time")
            c7.HeaderText = "TIME"
        End If
        If sd = 1 Then
            Dim c8 As DataGridViewColumn = dgReports.Columns("Sales Date")
            c8.HeaderText = "DATE"
        End If
        If e = 1 Then
            Dim c9 As DataGridViewColumn = dgReports.Columns("Employee")
            c9.HeaderText = "EMPLOYEE"
        End If
        If bcode = 1 Then
            Dim c10 As DataGridViewColumn = dgReports.Columns("Barcode Number")
            c10.HeaderText = "BARCODE"
        End If
        If tax = 1 Then
            Dim c11 As DataGridViewColumn = dgReports.Columns("Tax")
            c11.HeaderText = "TAX AMOUNT"
        End If
        If receipt = 1 Then
            Dim c12 As DataGridViewColumn = dgReports.Columns("Receipt")
            c12.HeaderText = "RECEIPT NUMBER"
        End If
        If discount = 1 Then
            Dim c13 As DataGridViewColumn = dgReports.Columns("Discount")
            c13.HeaderText = "DISCOUNT"
        End If


    End Sub
    Private Sub reports_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        resetForm()
        releaseTheReports()

        dgReports.AllowUserToResizeColumns = False
        dgReports.AllowUserToResizeRows = False
        dgReports.DefaultCellStyle.SelectionBackColor = Color.Transparent
        dgReports.DefaultCellStyle.SelectionForeColor = Color.Black
        dgReports.DefaultCellStyle.Font = New Font("Century Gothic", 10)
        dgReports.ColumnHeadersDefaultCellStyle.Font = New Font("Century Gothic", 10, FontStyle.Bold)


        dropTheAkronHammer()
    End Sub

    Private Sub btnSalesReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSalesReport.Click
        If salesReport = False Then
            resetForm()

            pnlMenuSalesReport.Visible = True
            btnSalesReport.BackColor = Color.Gold
            salesReport = True
        Else
            resetForm()
            btnSalesReport.BackColor = Color.LightGray
            salesReport = False
        End If

        menuSales = False

    End Sub

    Private Sub btnSalesDate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSalesDate.Click
        If salesSpecify = False Or cmbSpecifyMonth.Enabled = True Then
            btnSalesDate.BackColor = Color.Gold
            dtpSpecify.Enabled = True
            salesSpecify = True
            cmbSpecifyMonth.Enabled = False
            btnSalesMonth.BackColor = Color.White
        Else
            btnSalesDate.BackColor = Color.White
            cmbSpecifyMonth.Enabled = False
            btnSalesMonth.BackColor = Color.White
            dtpSpecify.Enabled = False
            salesSpecify = False
        End If
    End Sub

    Private Sub btnSalesMonth_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSalesMonth.Click
        If salesSpecify = False Or dtpSpecify.Enabled = True Then
            btnSalesMonth.BackColor = Color.Gold
            cmbSpecifyMonth.Enabled = True
            salesSpecify = True
            dtpSpecify.Enabled = False
            btnSalesDate.BackColor = Color.White
        Else
            btnSalesDate.BackColor = Color.White
            cmbSpecifyMonth.Enabled = False
            btnSalesMonth.BackColor = Color.White
            dtpSpecify.Enabled = False
            salesSpecify = False
        End If
    End Sub

    Dim menuSales As Boolean = False
    Dim isSales As Boolean = False
    Dim isRev As Boolean = False
    Private Sub btnMenuSales_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMenuSales.Click
        pnlPrintSalesReport.Visible = False
        If menuSales = False Or pnlRevenue.Visible = True Then
            dropTheAkronHammer()

            dgReports.Visible = True
            pnlRevenue.Visible = False
            lblSort.Visible = True
            btnSalesTop.Visible = True
            btnSalesWorst.Visible = True
            isRev = False
            isSales = True

            cmbSpecifyMonth.Enabled = False
            dtpSpecify.Enabled = False

            pnlSalesReport.Dock = DockStyle.Top
            pnlSalesReport.Visible = True

            btnMenuSales.BackColor = Color.Gold
            btnMenuGrossPr.BackColor = Color.Beige
            menuSales = True

            Try
                db = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\inventorydb.mdb")
                db.Open()
                dbcmd = New OleDb.OleDbCommand("SELECT [Barcode Number], [Item Name], [Item Description] FROM tbl_items", db)

                Dim reader As OleDbDataReader = dbcmd.ExecuteReader()

                While reader.Read
                    Dim barcode As String = reader("Barcode Number").ToString
                    Dim name As String = reader("Item Name").ToString
                    Dim description As String = reader("Item Description").ToString
                    Dim totQuanti As Integer = 0
                    Dim totPrice As Double = 0
                    Dim totCost As Double = 0
                    Dim totProfit As Double = 0

                    cmd = New OleDb.OleDbCommand("SELECT * FROM tbl_sales WHERE [Barcode Number] = @key", db)
                    cmd.Parameters.AddWithValue("@key", barcode)

                    Dim tempcmd As New OleDb.OleDbCommand
                    Dim reader2 As OleDbDataReader = cmd.ExecuteReader()

                    While reader2.Read
                        Dim quan As String = reader2("Quantity").ToString
                        totQuanti = totQuanti + Val(quan)
                        Dim price As String = reader2("Price").ToString
                        totPrice = totPrice + Val(price)
                        totPrice = Math.Round(totPrice, 2, MidpointRounding.AwayFromZero)
                        Dim cost As String = reader2("Capital").ToString
                        totCost = totCost + Val(cost)
                        totCost = Math.Round(totCost, 2, MidpointRounding.AwayFromZero)
                        Dim profit As String = reader2("Profit").ToString
                        totProfit = totProfit + Val(profit)
                        totProfit = Math.Round(totProfit, 2, MidpointRounding.AwayFromZero)
                    End While
                    reader2.Close()

                    tempcmd = New OleDb.OleDbCommand("INSERT INTO tbl_reportsTemp([Barcode Number], [Item Name], [Item Description], [Quantity], [Price], [Profit]) VALUES ('" & barcode & "', '" & name & "', '" & description & "', '" & totQuanti & "', '" & totPrice & "', '" & totProfit & "')", db)
                    tempcmd.ExecuteNonQuery()

                End While
                reader.Close()
            Catch ex As Exception
                MsgBox(ex.ToString)
            Finally
                db.Close()
            End Try

            Try
                db.Open()
                dba = New OleDb.OleDbDataAdapter("SELECT [Item Name], [Item Description], [Quantity], [Price], [Profit] FROM tbl_reportsTemp", db)
                dbds = New DataSet
                dba.Fill(dbds, "tbl_reportsTemp")

            Catch ex As Exception
            Finally
                db.Close()
            End Try

            dgReports.DataSource = dbds.Tables("tbl_reportsTemp")
            dgReports.Columns(0).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            dgReports.Columns(1).Width = 150
            kwak(0, 1, 1, 1, 1, 0, 1, 0, 0, 0, 0, 0, 0)
        Else
            isRev = False
            isSales = False

            btnMenuSales.BackColor = Color.Beige
            pnlSalesReport.Visible = False
            menuSales = False
            dgReports.Visible = False
        End If

    End Sub

    Private Sub btnMenuGrossPr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMenuGrossPr.Click        
        pnlPrintSalesReport.Visible = False

        If menuSales = False Or btnSalesTop.Visible = True Then

            dropTheAkronHammer()

            lblSort.Visible = False
            btnSalesTop.Visible = False
            btnSalesWorst.Visible = False
            isRev = True
            isSales = False

            cmbSpecifyMonth.Enabled = False
            dtpSpecify.Enabled = False

            pnlSalesReport.Dock = DockStyle.Top
            pnlSalesReport.Visible = True

            btnMenuSales.BackColor = Color.Beige
            btnMenuGrossPr.BackColor = Color.Gold
            menuSales = True

            txtGrossRev.Text = globalTotalRevenue
            txtGrossTax.Text = globalTotalTax
            txtGrossRefund.Text = globalTotalRefund
            txtGrossDiscount.Text = globalTotalDiscount
            txtGrossIncome.Text = globalNetInc
            txtGrossMisc.Text = globalTotalMisc
            txtGrossVoid.Text = globalTotalVoid

            Try
                db.Open()
                dbcmd = New OleDb.OleDbCommand("SELECT [Trans Date] FROM tbl_transactions", db)
                Dim reader As OleDbDataReader = dbcmd.ExecuteReader

                Dim cString As String = ""
                While reader.Read

                    Dim tempDate As String = reader("Trans Date").ToString

                    If cString <> tempDate Then
                        cString = tempDate
                        cmd = New OleDb.OleDbCommand("INSERT INTO tbl_reportsTemp([Sales Date]) VALUES ('" & tempDate & "')", db)
                        cmd.ExecuteNonQuery()
                    End If

                End While

                dbcmd = New OleDb.OleDbCommand("SELECT [Sales Date] FROM tbl_reportsTemp", db)
                reader = dbcmd.ExecuteReader

                While reader.Read

                    Dim theDate As String = reader("Sales Date").ToString

                    cmd = New OleDb.OleDbCommand("SELECT [Subtotal], [Total] FROM tbl_transactions WHERE [Trans Date] = '" & theDate & "'", db)
                    Dim reader2 As OleDbDataReader = cmd.ExecuteReader

                    Dim tempTotSub As Double = 0
                    Dim tempTotTot As Double = 0
                    While reader2.Read

                        Dim tempSub As String = reader2("Subtotal").ToString
                        tempTotSub = tempTotSub + Val(tempSub)
                        tempTotSub = Math.Round(tempTotSub, 2, MidpointRounding.AwayFromZero)
                        Dim tempTotal As String = reader2("Total").ToString
                        tempTotTot = tempTotTot + Val(tempTotal)
                        tempTotTot = Math.Round(tempTotTot, 2, MidpointRounding.AwayFromZero)

                    End While

                    Dim tempcmd As OleDb.OleDbCommand = New OleDb.OleDbCommand("INSERT INTO tbl_temp([ITEM], [UNIT PRICE], [TOTAL]) VALUES ('" & theDate & "', '" & tempTotSub & "', '" & tempTotTot & "')", db)
                    tempcmd.ExecuteNonQuery()

                End While

                chartGross.Series.Clear()
                chartGross.Series.Add("TOTAL REVENUE")
                chartGross.Series.Add("NET REVENUE")
                chartGross.Series("NET REVENUE").Color = Color.LimeGreen
                chartGross.Series("NET REVENUE").ChartType = SeriesChartType.FastLine

                chartGross.ChartAreas("ChartArea1").AxisX.Title = "Time"
                chartGross.ChartAreas("ChartArea1").AxisY.Title = "Amount"
                chartGross.Update()

                Dim series1 As Series = chartGross.Series("TOTAL REVENUE")
                Dim series2 As Series = chartGross.Series("NET REVENUE")

                dbcmd = New OleDb.OleDbCommand("SELECT [ITEM], [UNIT PRICE], [TOTAL] FROM tbl_temp", db)
                reader = dbcmd.ExecuteReader
                While reader.Read
                    Dim xValueString As String = reader("ITEM").ToString()
                    Dim xValue As DateTime

                    If DateTime.TryParse(xValueString, xValue) Then
                        Dim y1Value As Object = reader("UNIT PRICE")
                        Dim y2Value As Object = reader("TOTAL")

                        series1.Points.AddXY(xValue, y1Value)
                        series2.Points.AddXY(xValue, y2Value)
                    End If

                    Dim chartArea As ChartArea = chartGross.ChartAreas(0)
                    chartArea.AxisX.MajorGrid.Enabled = False
                    chartArea.AxisY.MajorGrid.Enabled = False
                    chartArea.AxisX.LabelStyle.Enabled = False

                End While

            Catch ex As Exception
                MsgBox(ex.ToString)
            Finally
                db.Close()
            End Try

            pnlRevenue.Dock = DockStyle.Fill
            pnlRevenue.Visible = True

        Else
            btnMenuGrossPr.BackColor = Color.Beige
            pnlSalesReport.Visible = False
            menuSales = False
            dgReports.Visible = False
            pnlRevenue.Visible = False

            isRev = False
            isSales = False
        End If
        
    End Sub


    Private Sub btnPaymentReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPaymentReport.Click
        If paymentReport = False Then
            dropTheAkronHammer()
            resetForm()

            paymentButtonsReset()
            btnPaymentReport.BackColor = Color.Gold
            pnlPaymentReport.Visible = True
            pnlPaymentReport.Dock = DockStyle.Top
            paymentReport = True
            pnlPaymentRep.Visible = True
            pnlPaymentRep.Dock = DockStyle.Fill
            pnlPaymentReport.BringToFront()

            db = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\inventorydb.mdb")
            Try
                db.Open()


                txtTotTax.Text = globalTotalTax
                txtTotDisc.Text = globalTotalDiscount
                txtTotRefund.Text = globalTotalRefund
                txtTotMisc.Text = globalTotalMisc
                txtTotVoid.Text = globalTotalVoid
                txtPaymentTotalPayment.Text = globalTotalTax + globalTotalDiscount + globalTotalRefund + globalTotalVoid + globalTotalMisc

                Dim srsTP As Series = chartTotalPayment.Series("srsTotalPayment")
                srsTP.ChartType = SeriesChartType.Pie
                srsTP.Name = "Distribution of Payment"
                srsTP.Points.AddXY("TAX", globalTotalTax)
                srsTP.Points.AddXY("DISCOUNT", globalTotalDiscount)
                srsTP.Points.AddXY("REFUND", globalTotalRefund)
                srsTP.Points.AddXY("VOID SALES", globalTotalVoid)
                srsTP.Points.AddXY("MISCELLANEOUS", globalTotalMisc)

            Catch ex As Exception

            Finally
                db.Close()
            End Try

        Else

            pnlPaymentRep.Visible = False
            btnPaymentReport.BackColor = Color.LightGray
            pnlPaymentReport.Visible = False
            paymentReport = False
        End If

    End Sub



    Private Sub btnSalesTop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSalesTop.Click
        Dim columnName As String = "Price" ' Replace with the actual column name

        If dgReports.Columns.Contains(columnName) Then
            dgReports.Sort(dgReports.Columns(columnName), System.ComponentModel.ListSortDirection.Descending)
            dgReports.Focus()
        End If

        btnSalesTop.BackColor = Color.Gold
        btnSalesWorst.BackColor = Color.WhiteSmoke
    End Sub

    Private Sub btnSalesWorst_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSalesWorst.Click
        Dim columnName As String = "Price" ' Replace with the actual column name

        If dgReports.Columns.Contains(columnName) Then
            dgReports.Sort(dgReports.Columns(columnName), System.ComponentModel.ListSortDirection.Ascending)
        End If

        btnSalesTop.BackColor = Color.WhiteSmoke
        btnSalesWorst.BackColor = Color.Gold
    End Sub

    Sub paymentButtonsReset()
        btnTaxReport.BackColor = Color.WhiteSmoke
        btnDiscountReport.BackColor = Color.WhiteSmoke
        btnRefundReport.BackColor = Color.WhiteSmoke
        btnVoidedSalesReport.BackColor = Color.WhiteSmoke
        btnMiscellanousReport.BackColor = Color.WhiteSmoke
    End Sub

    Private Sub btnTaxReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTaxReport.Click
        db = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\inventorydb.mdb")

        dgReports.Visible = True
        paymentButtonsReset()
        dropTheAkronHammer()
        btnTaxReport.BackColor = Color.Gold
        pnlPaymentRep.Visible = False
        Try
            db.Open()
            dbcmd = New OleDb.OleDbCommand("SELECT [Receipt Number], [Tax], [Trans Date], [Trans Time], [Cashier] FROM tbl_transactions", db)

            Dim reader As OleDbDataReader = dbcmd.ExecuteReader()
            While reader.Read
                Dim receipt As String = reader("Receipt Number").ToString
                Dim tempTime As String = reader("Trans Time").ToString
                Dim tax As String = reader("Tax").ToString
                Dim tempDate As String = reader("Trans Date").ToString
                Dim employee As String = reader("Cashier").ToString

                cmd = New OleDb.OleDbCommand("INSERT INTO tbl_reportsTemp([Receipt], [Tax], [Sales Date], [Sales Time], [Employee]) VALUES ('" & receipt & "', '" & tax & "', '" & tempDate & "', '" & tempTime & "', '" & employee & "')", db)
                cmd.ExecuteNonQuery()

            End While
            reader.Close()

            dba = New OleDb.OleDbDataAdapter("SELECT [Receipt], [Tax], [Sales Date], [Sales Time], [Employee] FROM tbl_reportsTemp", db)
            dbds = New DataSet
            dba.Fill(dbds, "tbl_reportsTemp")

        Catch ex As Exception
            MsgBox(ex.ToString)
        Finally
            db.Close()
        End Try

        dgReports.DataSource = dbds.Tables("tbl_reportsTemp")
        kwak(0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 0)
        dgReports.Columns.Cast(Of DataGridViewColumn)().ToList().ForEach(Sub(column) column.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill)
    End Sub

    Private Sub btnRefundReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefundReport.Click
        db = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\inventorydb.mdb")

        dgReports.Visible = True
        paymentButtonsReset()
        dropTheAkronHammer()
        btnRefundReport.BackColor = Color.Gold
        pnlPaymentRep.Visible = False

        Try
            db.Open()
            dbcmd = New OleDb.OleDbCommand("SELECT * FROM tbl_refund", db)
            dbcmd.ExecuteNonQuery()

            Dim reader As OleDb.OleDbDataReader = dbcmd.ExecuteReader
            While reader.Read
                Dim barcode As String = reader("Barcode Number").ToString
                Dim amount As String = reader("Refund Amount").ToString
                Dim qty As String = reader("Quantity").ToString
                Dim name As String = reader("Item Name").ToString
                Dim employee As String = reader("Cashier").ToString
                Dim tdate As String = reader("Refund Date").ToString
                Dim ttime As String = reader("Refund Time").ToString

                cmd = New OleDb.OleDbCommand("INSERT INTO tbl_reportsTemp([Barcode Number], [Price], [Quantity], [Item Name], [Employee], [Sales Date], [Sales Time]) VALUES ('" & barcode & "', '" & amount & "', '" & qty & "', '" & name & "', '" & employee & "', '" & tdate & "', '" & ttime & "')", db)
                cmd.ExecuteNonQuery()
            End While
            reader.Close()

            dba = New OleDb.OleDbDataAdapter("SELECT [Barcode Number], [Item Name], [Price], [Quantity], [Employee], [Sales Date], [Sales Time] FROM tbl_reportsTemp", db)
            dbds = New DataSet
            dba.Fill(dbds, "tbl_reportsTemp")

        Catch ex As Exception
        Finally
            db.Close()
        End Try

        dgReports.DataSource = dbds.Tables("tbl_reportsTemp")
        kwak(1, 1, 0, 1, 1, 0, 0, 1, 1, 1, 0, 0, 0)
        dgReports.Columns.Cast(Of DataGridViewColumn)().ToList().ForEach(Sub(column) column.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill)
        dgReports.Columns(0).Width = 150
    End Sub

    Private Sub btnDiscountReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDiscountReport.Click
        db = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\inventorydb.mdb")

        dgReports.Visible = True
        paymentButtonsReset()
        dropTheAkronHammer()
        btnDiscountReport.BackColor = Color.Gold
        pnlPaymentRep.Visible = False
        Try
            db.Open()
            dbcmd = New OleDb.OleDbCommand("SELECT [Receipt Number], [Discount], [Trans Date], [Trans Time], [Cashier] FROM tbl_transactions", db)

            Dim reader As OleDbDataReader = dbcmd.ExecuteReader()
            While reader.Read
                Dim receipt As String = reader("Receipt Number").ToString
                Dim tempTime As String = reader("Trans Time").ToString
                Dim discount As String = reader("Discount").ToString
                Dim tempDate As String = reader("Trans Date").ToString
                Dim employee As String = reader("Cashier").ToString

                If Val(discount) <> 0 Then
                    cmd = New OleDb.OleDbCommand("INSERT INTO tbl_reportsTemp([Receipt], [Discount], [Sales Date], [Sales Time], [Employee]) VALUES ('" & receipt & "', '" & discount & "', '" & tempDate & "', '" & tempTime & "', '" & employee & "')", db)
                    cmd.ExecuteNonQuery()
                End If

            End While
            reader.Close()

            dba = New OleDb.OleDbDataAdapter("SELECT [Receipt], [Discount], [Sales Date], [Sales Time], [Employee] FROM tbl_reportsTemp", db)
            dbds = New DataSet
            dba.Fill(dbds, "tbl_reportsTemp")

        Catch ex As Exception
            MsgBox(ex.ToString)
        Finally
            db.Close()
        End Try

        dgReports.DataSource = dbds.Tables("tbl_reportsTemp")
        kwak(0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 0, 1, 1)
        dgReports.Columns.Cast(Of DataGridViewColumn)().ToList().ForEach(Sub(column) column.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill)
    End Sub

    Private Sub btnVoidedSalesReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnVoidedSalesReport.Click
        db = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\inventorydb.mdb")

        dgReports.Visible = True
        paymentButtonsReset()
        dropTheAkronHammer()
        btnVoidedSalesReport.BackColor = Color.Gold
        pnlPaymentRep.Visible = False

        Try
            db.Open()
            dba = New OleDb.OleDbDataAdapter("SELECT [Receipt], [Total], [Date Purchased], [Date Voided], [Employee] from tbl_voidsales", db)
            dbds = New DataSet
            dba.Fill(dbds, "tbl_voidsales")
        Catch ex As Exception
            MsgBox(ex.ToString)
        Finally
            db.Close()
        End Try
        dgReports.DataSource = dbds.Tables("tbl_voidsales")
        kwak(0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 1, 0)
        Dim c1 As DataGridViewColumn = dgReports.Columns("Total")
        c1.HeaderText = "TOTAL"
        Dim c2 As DataGridViewColumn = dgReports.Columns("Date Voided")
        c2.HeaderText = "DATE VOIDED"
        Dim c3 As DataGridViewColumn = dgReports.Columns("Date Purchased")
        c3.HeaderText = "DATE PURCHASED"
        dgReports.Columns.Cast(Of DataGridViewColumn)().ToList().ForEach(Sub(column) column.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill)
    End Sub

    Private Sub btnMiscellanousReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMiscellanousReport.Click
        db = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\inventorydb.mdb")

        dgReports.Visible = True
        paymentButtonsReset()
        dropTheAkronHammer()
        btnMiscellanousReport.BackColor = Color.Gold
        pnlPaymentRep.Visible = False

        Try
            db.Open()
            dba = New OleDb.OleDbDataAdapter("SELECT [Amount], [Employee], [Misc_date], [Misc_time] from tbl_miscfee", db)
            dbds = New DataSet
            dba.Fill(dbds, "tbl_miscfee")
        Catch ex As Exception
            MsgBox(ex.ToString)
        Finally
            db.Close()
        End Try
        dgReports.DataSource = dbds.Tables("tbl_miscfee")
        Dim c0 As DataGridViewColumn = dgReports.Columns("Amount")
        c0.HeaderText = "AMOUNT"
        Dim c1 As DataGridViewColumn = dgReports.Columns("Employee")
        c1.HeaderText = "EMPLOYEE"
        Dim c2 As DataGridViewColumn = dgReports.Columns("Misc_date")
        c2.HeaderText = "DATE"
        Dim c3 As DataGridViewColumn = dgReports.Columns("Misc_time")
        c3.HeaderText = "TIME"
        dgReports.Columns.Cast(Of DataGridViewColumn)().ToList().ForEach(Sub(column) column.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill)
    End Sub

    Private Sub ExportToExcel(ByVal report)
        If dgReports.RowCount > 0 Then
            Dim excelApp As New Excel.Application()

            Dim excelWorkbook As Excel.Workbook = excelApp.Workbooks.Add(Type.Missing)
            Dim excelWorksheet As Excel.Worksheet = excelWorkbook.Sheets(1)

            ' Set the column headers in Excel
            For i As Integer = 0 To dgReports.Columns.Count - 1
                excelWorksheet.Cells(1, i + 1) = dgReports.Columns(i).HeaderText
            Next

            ' Export data from DataGrid to Excel
            For i As Integer = 0 To dgReports.Rows.Count - 1
                For j As Integer = 0 To dgReports.Columns.Count - 1
                    If dgReports.Rows(i).Cells(j).Value IsNot Nothing Then
                        excelWorksheet.Cells(i + 2, j + 1) = dgReports.Rows(i).Cells(j).Value.ToString()
                    End If
                Next
            Next

            ' Save the Excel file
            Dim saveFileDialog As New SaveFileDialog()
            saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
            saveFileDialog.FileName = "The ISO Team Enterprise " & report & " Report"
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

    Private Sub btnSalesExtract_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSalesExtract.Click
        Using fd As SaveFileDialog = New SaveFileDialog() With {.Filter = "Excel Workbook|*.xlsx"}
            Try
                ExportToExcel("Sales")
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Using
    End Sub

   
    Private Sub btnPaymentExtractR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPaymentExtractR.Click
        ExportToExcel("Payment")
    End Sub

    Private Sub btnEmployeeReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEmployeeReport.Click
        If employeeReport = False Then
            dropTheAkronHammer()
            resetForm()

            btnEmployeeReport.BackColor = Color.Gold
            pnlEmployeeReport.Dock = DockStyle.Top
            pnlEmployeeReport.Visible = True
            employeeReport = True
            dgReports.Visible = True

            btnEmployeeList.PerformClick()

        Else
            resetForm()
        End If
    End Sub

    Sub employeeButtonReset()
        btnEmployeeList.BackColor = Color.WhiteSmoke
        btnEmployeeArchive.BackColor = Color.WhiteSmoke
        btnEmployeeDTR.BackColor = Color.WhiteSmoke
        btnPrintDTR.Visible = False
        pnlPrintDTR.Visible = False
        employeeReportSelector = 0
    End Sub

    Private Sub btnEmployeeList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEmployeeList.Click
        employeeButtonReset()
        btnPrintDTR.Text = "PRINT EMPLOYEES"
        btnPrintDTR.Visible = True
        btnEmployeeList.BackColor = Color.Gold
        employeeReportSelector = 2

        db = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\posdb.mdb")
        Try
            db.Open()
            dba = New OleDb.OleDbDataAdapter("SELECT [USER ID],[USER NAME],[POSITION],[PRIVILEGE] FROM tbluser", db)
            dbds = New DataSet
            dba.Fill(dbds, "tbluser")
        Catch ex As Exception
            MsgBox(ex.ToString)
        Finally
            db.Close()
        End Try

        dgReports.DataSource = dbds.Tables("tbluser")
        Dim c2 As DataGridViewColumn = dgReports.Columns("USER NAME")
        c2.HeaderText = "EMPLOYEE NAME"
        Dim c3 As DataGridViewColumn = dgReports.Columns("POSITION")
        c3.HeaderText = "EMPLOYEE POSITION"
        dgReports.Columns.Cast(Of DataGridViewColumn)().ToList().ForEach(Sub(column) column.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill)
    End Sub

    Private Sub btnEmployeeDTR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEmployeeDTR.Click
        employeeButtonReset()
        btnPrintDTR.Text = "PRINT DTR"
        btnPrintDTR.Visible = True
        btnEmployeeDTR.BackColor = Color.Gold
        employeeReportSelector = 1

        db = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\posdb.mdb")
        Try
            db.Open()
            dba = New OleDb.OleDbDataAdapter("SELECT [user_id], [username], [login_date], [login_time], [logout_date], [logout_time] FROM tbl_dtr", db)
            dbds = New DataSet
            dba.Fill(dbds, "tbl_dtr")
        Catch ex As Exception
            MsgBox(ex.ToString)
        Finally
            db.Close()
        End Try

        dgReports.DataSource = dbds.Tables("tbl_dtr")
        Dim c1 As DataGridViewColumn = dgReports.Columns("user_id")
        c1.HeaderText = "USER ID"
        Dim c2 As DataGridViewColumn = dgReports.Columns("username")
        c2.HeaderText = "EMPLOYEE NAME"
        Dim c3 As DataGridViewColumn = dgReports.Columns("login_date")
        c3.HeaderText = "LOGIN DATE"
        Dim c4 As DataGridViewColumn = dgReports.Columns("login_time")
        c4.HeaderText = "LOGIN DATE"
        Dim c5 As DataGridViewColumn = dgReports.Columns("logout_date")
        c5.HeaderText = "LOGOUT DATE"
        Dim c6 As DataGridViewColumn = dgReports.Columns("logout_time")
        c6.HeaderText = "LOGOUT TIME"
        dgReports.Columns.Cast(Of DataGridViewColumn)().ToList().ForEach(Sub(column) column.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill)
    End Sub

    Private Sub btnEmployeeArchive_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEmployeeArchive.Click
        employeeButtonReset()
        btnPrintDTR.Text = "PRINT ARCHIVE"
        btnPrintDTR.Visible = True
        btnEmployeeArchive.BackColor = Color.Gold
        employeeReportSelector = 3

        db = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\posdb.mdb")
        Try
            db.Open()
            dba = New OleDb.OleDbDataAdapter("SELECT [USER ID],[USER NAME],[POSITION],[PRIVILEGE], [DATE ARCHIVE] FROM tblArchive", db)
            dbds = New DataSet
            dba.Fill(dbds, "tblArchive")
        Catch ex As Exception
            MsgBox(ex.ToString)
        Finally
            db.Close()
        End Try

        dgReports.DataSource = dbds.Tables("tblArchive")
        Dim c2 As DataGridViewColumn = dgReports.Columns("USER NAME")
        c2.HeaderText = "EMPLOYEE NAME"
        Dim c3 As DataGridViewColumn = dgReports.Columns("POSITION")
        c3.HeaderText = "EMPLOYEE POSITION"
        Dim c5 As DataGridViewColumn = dgReports.Columns("DATE ARCHIVE")
        c5.HeaderText = "DATE ARCHIVE"
        dgReports.Columns.Cast(Of DataGridViewColumn)().ToList().ForEach(Sub(column) column.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill)
    End Sub

    Private Sub btnTransactionLog_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTransactionLog.Click
        If transactionLog = False Then
            dropTheAkronHammer()
            resetForm()

            btnTransactionLog.BackColor = Color.Gold
            pnlTransactionLog.Dock = DockStyle.Top
            pnlTransactionLog.Visible = True
            transactionLog = True
            dgReports.Visible = True

            db = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\inventorydb.mdb")
            Try
                db.Open()
                dba = New OleDb.OleDbDataAdapter("SELECT * FROM tbl_transactions", db)
                dbds = New DataSet
                dba.Fill(dbds, "tbl_transactions")
            Catch ex As Exception
                MsgBox(ex.ToString)
            Finally
                db.Close()
            End Try

            dgReports.DataSource = dbds.Tables("tbl_transactions")
            kwak(0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 1, 0, 1)
            Dim c1 As DataGridViewColumn = dgReports.Columns("Receipt Number")
            c1.HeaderText = "RECEIPT"
            Dim c2 As DataGridViewColumn = dgReports.Columns("Item Purchased")
            c2.HeaderText = "ITEM PURCHASED"
            Dim c3 As DataGridViewColumn = dgReports.Columns("Subtotal")
            c3.HeaderText = "SUBTOTAL"
            Dim c5 As DataGridViewColumn = dgReports.Columns("Total")
            c5.HeaderText = "TOTAL"
            Dim c6 As DataGridViewColumn = dgReports.Columns("Cashier")
            c6.HeaderText = "EMPLOYEE"
            Dim c7 As DataGridViewColumn = dgReports.Columns("Trans Date")
            c7.HeaderText = "DATE"
            Dim c8 As DataGridViewColumn = dgReports.Columns("Trans Time")
            c8.HeaderText = "TIME"
        Else
            resetForm()
        End If
    End Sub

    Private Sub btnInventoryReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInventoryReport.Click
        If inventoryReport = False Then
                dropTheAkronHammer()
                resetForm()

            btnInventoryReport.BackColor = Color.Gold
            pnlInventory.Dock = DockStyle.Top
            pnlInventory.Visible = True
            inventoryReport = True
            dgReports.Visible = True

                db = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\inventorydb.mdb")
                Try
                    db.Open()
                dba = New OleDb.OleDbDataAdapter("SELECT * FROM tbl_items", db)
                    dbds = New DataSet
                dba.Fill(dbds, "tbl_items")
                Catch ex As Exception
                    MsgBox(ex.ToString)
                Finally
                    db.Close()
                End Try

            dgReports.DataSource = dbds.Tables("tbl_items")
            kwak(1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Dim c5 As DataGridViewColumn = dgReports.Columns("Refund")
            c5.HeaderText = "REFUND QTY"
            Dim c6 As DataGridViewColumn = dgReports.Columns("Selling Price")
            c6.HeaderText = "SELLING PRICE"
            Dim c7 As DataGridViewColumn = dgReports.Columns("Buying Price")
            c7.HeaderText = "BUYING PRICE"
            Dim c8 As DataGridViewColumn = dgReports.Columns("Critical Value")
            c8.HeaderText = "CRITICAL VAL"
            dgReports.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Else
                resetForm()
            End If
    End Sub

    Private Sub btnExtractInventoryReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExtractInventoryReport.Click
        ExportToExcel("Inventory")
    End Sub

    Private Sub btnExtractTransactionLog_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExtractTransactionLog.Click
        ExportToExcel("Transaction Log")
    End Sub

    Private Sub btnExtractEmployeeReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExtractEmployeeReport.Click
        ExportToExcel("Employee")
    End Sub


    Private Sub btnReturn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReturn.Click
        resetForm()
        pnlReturn.Visible = True
    End Sub

    Private Sub btnReturnToMain_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReturnToMain.Click
        Me.Hide()
        main.Show()
    End Sub

    Private Sub btnLogout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogout.Click
        Dim res As DialogResult = MessageBox.Show("Are you sure you want to logout?", "LOGOUT", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If res = DialogResult.Yes Then
            logout()
            Me.Hide()
            login.Show()
            currentUser = ""
        End If
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Dim res As DialogResult = MessageBox.Show("Are you sure you want to exit?", "EXIT", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If res = DialogResult.Yes Then
            logout()
            Me.Close()
        End If
    End Sub

    Private Sub btnPrintTransactionLog_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrintTransactionLog.Click
        pnlPrintTransaction.Visible = True
        pnlPrintTransaction.Dock = DockStyle.Fill

        Dim crysTransaction As New ReportDocument
        crysTransaction.Load(Application.StartupPath & "\reports\transactionReport.rpt")
        crepViewerTransaction.ReportSource = crysTransaction
        crepViewerTransaction.Refresh()
    End Sub

    Private Sub btnPrintDTR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrintDTR.Click
        pnlPrintDTR.Visible = True
        pnlPrintDTR.Dock = DockStyle.Fill

        Dim crysDTR As New ReportDocument
        If employeeReportSelector = 1 Then
            crysDTR.Load(Application.StartupPath & "\reports\dtrReport.rpt")
        ElseIf employeeReportSelector = 2 Then
            crysDTR.Load(Application.StartupPath & "\reports\userReport.rpt")
        ElseIf employeeReportSelector = 3 Then
            crysDTR.Load(Application.StartupPath & "\reports\userArchiveReport.rpt")
        End If
        crepViewerDTR.ReportSource = crysDTR
        crepViewerDTR.Refresh()
    End Sub

    Private Sub btnPrintSales_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrintSales.Click
        pnlPrintSalesReport.Visible = True
        pnlPrintSalesReport.Dock = DockStyle.Fill

        Dim salesReport As New ReportDocument

        If isSales = True Then
            salesReport.Load(Application.StartupPath & "\reports\salesReport.rpt")
            crepViewerSalesReport.ReportSource = salesReport
        ElseIf isRev = True Then
            salesReport.Load(Application.StartupPath & "\reports\revenueReport.rpt")

            salesReport.SetParameterValue("totalRev", txtGrossIncome.Text)
            salesReport.SetParameterValue("misc", txtGrossMisc.Text)
            salesReport.SetParameterValue("refund", txtGrossRefund.Text)
            salesReport.SetParameterValue("tax", txtGrossTax.Text)
            salesReport.SetParameterValue("netRev", txtGrossRev.Text)
            salesReport.SetParameterValue("void", txtGrossVoid.Text)
            salesReport.SetParameterValue("discount", txtGrossDiscount.Text)

            crepViewerSalesReport.ReportSource = salesReport
        End If

        crepViewerSalesReport.Refresh()
    End Sub

    Private Sub btnPrintInventory_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrintInventory.Click
        pnlPrintSalesReport.Visible = True
        pnlPrintSalesReport.Dock = DockStyle.Fill

        Dim inventoryReport As New ReportDocument
        If isInventoryArchive = False Then
            inventoryReport.Load(Application.StartupPath & "\reports\stockReport.rpt")
        Else
            inventoryReport.Load(Application.StartupPath & "\reports\inventoryArchiveReport.rpt")
        End If
        crepViewerSalesReport.ReportSource = inventoryReport
        crepViewerSalesReport.Refresh()
    End Sub

    Dim isInventoryArchive As Boolean = False
    Private Sub btnInventoryArchive_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInventoryArchive.Click
        If isInventoryArchive = False Then
            pnlPrintSalesReport.Visible = False
            btnInventoryArchive.BackColor = Color.WhiteSmoke
            btnInventoryArchive.Text = "RETURN"

            db = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\inventorydb.mdb")
            Try
                db.Open()
                dba = New OleDb.OleDbDataAdapter("SELECT * FROM tbl_archiveItems", db)
                dbds = New DataSet
                dba.Fill(dbds, "tbl_archiveItems")
            Catch ex As Exception
                MsgBox(ex.ToString)
            Finally
                db.Close()
            End Try
            dgReports.Visible = True
            dgReports.DataSource = dbds.Tables("tbl_archiveItems")
            isInventoryArchive = True
        Else
            btnInventoryArchive.BackColor = Color.Gold
            btnInventoryArchive.Text = "INVENTORY ARCHIVE"
            isInventoryArchive = False
            resetForm()
            btnInventoryReport.PerformClick()
        End If
    End Sub
End Class
