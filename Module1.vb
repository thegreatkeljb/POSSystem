Imports System.Data.OleDb
Imports CrystalDecisions.CrystalReports.Engine


Module Module1
    Public db As New OleDbConnection
    Public dba As New OleDbDataAdapter
    Public dbds As New DataSet
    Public dbcmd As New OleDb.OleDbCommand

    Public isAdmin As Boolean
    Public isCashier As Boolean
    Public isInventory As Boolean

    Public currentUser As String
    Public currentUserName As String

    Public currentDateAndTime As DateTime = DateTime.Now
    Public currentDate As String = currentDateAndTime.ToString("MM/dd/yy")

    Public crysReceipt As New ReportDocument
    Public c As New ReportDocument

    Sub userCon()
        Try
            db = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\posdb.mdb")
            dba = New OleDb.OleDbDataAdapter("SELECT [USER ID],[USER NAME],[POSITION],[PRIVILEGE] FROM tbluser", db)
            dbds = New DataSet
            dba.Fill(dbds, "tbluser")

        Catch ex As Exception
            MsgBox(ex.ToString)
        Finally
            db.Close()
        End Try
    End Sub

    Sub inventoryCon()
        Try
            db = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\inventorydb.mdb")
            dba = New OleDb.OleDbDataAdapter("SELECT * FROM tbl_items", db)
            dbds = New DataSet
            dba.Fill(dbds, "tbl_items")

        Catch ex As Exception
            MsgBox(ex.ToString)
        Finally
            db.Close()
        End Try
    End Sub

    Public globalTotalRevenue As Double = 0
    Public globalTotalDiscount As Double = 0
    Public globalTotalTax As Double = 0
    Public globalTotalRefund As Double = 0
    Public globalTotalMisc As Double = 0
    Public globalTotalVoid As Double = 0
    Public globalNetInc As Double = 0

    Sub releaseTheReports()
        db = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\inventorydb.mdb")
        Try
            db.Open()
            dbcmd = New OleDb.OleDbCommand("SELECT [Total], [Discount], [Tax] FROM tbl_transactions", db)

            Dim reader As OleDbDataReader = dbcmd.ExecuteReader
            While reader.Read
                Dim tempTotal As String = reader("Total").ToString
                globalTotalRevenue = globalTotalRevenue + Val(tempTotal)
                globalTotalRevenue = Math.Round(globalTotalRevenue, 2, MidpointRounding.AwayFromZero)

                Dim tempDiscount As String = reader("Discount").ToString
                globalTotalDiscount = globalTotalDiscount + Val(tempDiscount)
                globalTotalDiscount = Math.Round(globalTotalDiscount, 2, MidpointRounding.AwayFromZero)

                Dim tempTax As String = reader("Tax").ToString
                globalTotalTax = globalTotalTax + Val(tempTax)
                globalTotalTax = Math.Round(globalTotalTax, 2, MidpointRounding.AwayFromZero)

            End While
            reader.Close()

            dbcmd = New OleDb.OleDbCommand("SELECT [Refund Amount] FROM tbl_refund", db)

            reader = dbcmd.ExecuteReader
            While reader.Read
                Dim tempRefund As String = reader("Refund Amount").ToString

                If tempRefund IsNot String.Empty Then
                    globalTotalRefund = globalTotalRefund + Val(tempRefund)
                    globalTotalRefund = Math.Round(globalTotalRefund, 2, MidpointRounding.AwayFromZero)
                End If

            End While
            reader.Close()

            dbcmd = New OleDb.OleDbCommand("SELECT [Total] FROM tbl_voidsales", db)
            reader = dbcmd.ExecuteReader
            While reader.Read
                Dim tempamt As String = reader("Total").ToString

                If tempamt IsNot String.Empty Then
                    globalTotalVoid = globalTotalVoid + Val(tempamt)
                    globalTotalVoid = Math.Round(globalTotalVoid, 2, MidpointRounding.AwayFromZero)
                End If

            End While
            reader.Close()

            dbcmd = New OleDb.OleDbCommand("SELECT [Amount] FROM tbl_miscfee", db)
            reader = dbcmd.ExecuteReader
            While reader.Read
                Dim tempamt As String = reader("Amount").ToString

                If tempamt IsNot String.Empty Then
                    globalTotalMisc = globalTotalMisc + Val(tempamt)
                    globalTotalMisc = Math.Round(globalTotalMisc, 2, MidpointRounding.AwayFromZero)
                End If

            End While
            reader.Close()


            globalNetInc = globalTotalRevenue - (globalTotalDiscount + globalTotalTax + globalTotalMisc + globalTotalVoid + globalTotalRefund)
            globalNetInc = Math.Round(globalNetInc, 2, MidpointRounding.AwayFromZero)
        Catch ex As Exception
        Finally
            db.Close()
        End Try
    End Sub


    Sub logout()
        db = New OleDb.OleDbConnection("PROVIDER=microsoft.jet.oledb.4.0; data source = " & Application.StartupPath & "\posdb.mdb")
        currentDate = currentDateAndTime.ToString("MM/dd/yy")
        Dim tempTime As DateTime = DateTime.Now
        Dim currentTime As String = tempTime.ToString("h:mm tt")
        Try
            db.Open()
            dbcmd = New OleDb.OleDbCommand("UPDATE tbl_dtr SET [logout_date] = @param1, [logout_time] = @param2 WHERE key = (SELECT MAX(key) FROM tbl_dtr)", db)
            dbcmd.Parameters.AddWithValue("@param1", currentDate.ToString)
            dbcmd.Parameters.AddWithValue("@param2", currentTime)
            dbcmd.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox(ex.ToString)
        Finally
            db.Close()
        End Try
        currentUser = String.Empty
        isCashier = False
        isAdmin = False
        isInventory = False
    End Sub
End Module

