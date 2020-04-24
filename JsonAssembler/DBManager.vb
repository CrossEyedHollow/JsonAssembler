Imports ReportTools

Public Class DBManager
    Inherits DBBase

    Public Sub New()
        Init()
        orderLengths = New Dictionary(Of String, Integer)
    End Sub

    Dim orderLengths As Dictionary(Of String, Integer)

    Public Sub UpdateStatus(index As Integer, errors As Integer, errorArr As String)
        Dim query As String = UpdateStatusQuery(index, errors, errorArr)
        Execute(query)
    End Sub

    Public Function CheckJsonStatus() As DataTable
        Dim query As String = SelectUncheckedJsons()
        Return ReadDatabase(query)
    End Function

    Public Function SelectDeploymentsForInvoice(index As Integer) As DataTable
        Dim query As String = SelectDeploymentsForInvoiceQuery(index)
        Return ReadDatabase(query)
    End Function

    Public Function CheckForAggregatedCodes(table As String) As DataTable
        Dim query As String = SelectAggregatedCodesQuery(table)
        Return ReadDatabase(query)
    End Function

    Public Function CheckForPrintedCodes(table As String) As DataTable
        Dim query As String = SelectPrintedCodesQuery(table)
        Return ReadDatabase(query)
    End Function

    Public Function GetPrintedCode(table As String, code As String)
        Dim query = $"SELECT fldCode, fldPrintCode from `{DBName}`.`{table}` WHERE fldCode = '{code}';"
        Return ReadDatabase(query)
    End Function

    Public Function CheckForInvoice() As DataTable
        Dim query As String = SelectInvoicesQuery()
        Return ReadDatabase(query)
    End Function

    Public Function CheckForPayments() As DataTable
        Dim query As String = SelectPaymentsQuery()
        Return ReadDatabase(query)
    End Function

    Public Function CheckForArivals() As DataTable
        Dim query As String = SelectArivalsQuery()
        Return ReadDatabase(query)
    End Function

    Public Function CheckForDispatchEvent() As DataTable
        Dim query As String = SelectDispatchEvent()
        Return ReadDatabase(query)
    End Function

    Public Function GetCodesForIDs(IDs As Integer()) As DataTable
        Dim query As String = SelectCodesForIdQuery(IDs)
        Return ReadDatabase(query)
    End Function

    Public Sub ConfirmPrintedCodes(table As String, codes() As String, jsonID As Integer)
        Dim query As String = UpdatePrintedCodeQuery(table, codes, jsonID)
        Execute(query)
    End Sub

    Public Function GetHumanReadableCodeLength(orderID As Integer) As Integer
        'Check if it exists
        If orderLengths.Count > 0 Then
            If orderLengths.Keys.Contains(orderID) Then Return orderLengths(orderID)
        End If
        'Get it from the database
        Dim query As String = SelectHumanReadableLengthQuery(orderID)
        Dim dtResult = ReadDatabase(query)
        If dtResult.Rows.Count = 1 Then
            Dim intResult As Integer = Convert.ToInt32(dtResult.Rows(0)("fldPrimaryCodeLength"))
            orderLengths.Add(orderID, intResult)
            Return Convert.ToInt32(dtResult.Rows(0)("fldPrimaryCodeLength"))
        End If
        Throw New Exception($"Unexpected rows count for orderID: {orderID}, expected rows count: 1")
    End Function

    Public Function GetOrderProducts(orderID As String) As DataTable
        Dim query As String = SelectInvoiceProductQuery(orderID)
        Return ReadDatabase(query)
    End Function

    Public Function CheckForRecalls() As DataTable
        Dim query As String = SelectRecallQuery()
        Return ReadDatabase(query)
    End Function

    Public Sub ConfirmAggregatedCodes(table As String, codes As String(), jsonID As Integer)
        Dim query As String = UpdateAggregatedCoderQuery(table, codes, jsonID)
        Execute(query)
    End Sub

    Public Sub ConfirmInvoice(index As Integer, jsonID As Integer)
        Dim query As String = ConfirmInvoiceQuery(index, jsonID)
        Execute(query)
    End Sub

    Public Sub ConfirmDispatchEvent(index As Integer, jsonID As Integer)
        Dim query As String = ConfirmDispatchQuery(index, jsonID)
        Execute(query)
    End Sub

    Public Sub ConfirmRecall(index As Integer, jsonID As Integer)
        Dim query As String = ConfirmRecallQuery(index, jsonID)
        Execute(query)
    End Sub

    Public Sub ConfirmArrival(index As Integer, jsonID As Integer)
        Dim query As String = ConfirmArrivalQuery(index, jsonID)
        Execute(query)
    End Sub

    Public Sub ConfirmPayment(index As Integer, jsonID As Integer)
        Dim query As String = ConfirmPaymentQuery(index, jsonID)
        Execute(query)
    End Sub

    Public Function GetDispatchedCodes(table As String, dispatchID As Integer) As String()
        'Assemble query
        Dim query As String = SelectDispatchedCodesQuery(table, dispatchID)
        'Read db
        Dim result As DataTable = ReadDatabase(query)
        'If no rows are returned, return empty array of string, else return the codes
        If result.Rows.Count < 1 Then Return New String() {}
        Return GetAllCodes(result, "fldPrintCode")
    End Function

    Public Function InsertJson(body As String, type As String, recallCode As String) As Integer
        Dim query As String = InsertNewJsonQuery(body, type, recallCode)
        Return ExecuteReturnIndex(query)
    End Function

    Public Sub ClearDispatchment(uiType As Integer, upUIs As String(), aUIs As String())
        Dim query As String = ClearDispatchedQuery(uiType, upUIs, aUIs)
        Execute(query)
    End Sub

#Region "Queries"
    Private Function UpdateStatusQuery(index As Integer, errors As Integer, errorArr As String) As String
        Dim array As String = If(errorArr.IsNullOrEmpty(), "null", $"'{errorArr}'")
        Return $"UPDATE `{DBName}`.`tbljson` SET fldStatus = 1, fldError = {errors}, fldErrorArr = {array} WHERE fldIndex = {index};"
    End Function

    Private Function InsertNewJsonQuery(body As String, type As String, recallCode As String) As String
        Return $"INSERT INTO `{DBName}`.`tbljson` (fldType, fldJson, fldRecallCode, fldStatus) VALUES ('{type}', '{body}', '{recallCode}', 0);"
    End Function

    Private Function SelectDeploymentsForInvoiceQuery(index As Integer)
        Return $"SELECT * FROM `{DBName}`.`tbldeployment` where fldInvoiceID = {index};"
    End Function

    Private Function ConfirmDispatchQuery(index As Integer, jsonID As Integer)
        Return $"UPDATE `{DBName}`.`tbldeployment` SET fldRep = NOW(), fldJsonID = {jsonID} WHERE fldIndex = {index};"
    End Function

    Private Function SelectDispatchEvent() As String
        Dim output As String = ""
        output += "SELECT i.*, o.fldUI_Type "
        output += $"FROM `{DBName}`.`tbldeployment` AS i "
        output += $"LEFT JOIN (`{DBName}`.`tblorders` AS o) "
        output += "ON o.fldIndex = i.fldOrderID "
        output += "WHERE i.fldConfirmDate IS NOT NULL AND i.fldRep IS NULL;"
        Return output
    End Function

    Private Function ConfirmInvoiceQuery(index As Integer, jsonID As Integer)
        Return $"UPDATE `{DBName}`.`tblinvoices` SET fldRep = NOW(), fldJsonID = {jsonID} WHERE fldIndex = {index};"
    End Function

    Private Function ConfirmRecallQuery(index As Integer, jsonID As Integer)
        Return $"UPDATE `{DBName}`.`tblrecall` SET fldRep = NOW(), fldJsonID = {jsonID} WHERE fldIndex = {index};"
    End Function

    Private Function ConfirmArrivalQuery(index As Integer, jsonID As Integer)
        Return $"UPDATE `{DBName}`.`tblarrival` SET fldRep = NOW(), fldJsonID = {jsonID} WHERE fldIndex = {index};"
    End Function

    Private Function ClearDispatchedQuery(uiType As Integer, upUIs As String(), aUIs As String())
        Dim output As String = ""
        Select Case uiType
            Case 1
                output += $"UPDATE `{DBName}`.`tblprimarycodes` SET fldDipatchDate = NULL, fldDispatchID = NULL WHERE fldPrintCode in ({String.Join(",", upUIs)})"
            Case 2, 3
                Dim aggCodes As String = String.Join(",", aUIs)
                output += $"UPDATE `{DBBase.DBName}`.`tblboxcodes` AS B, `{DBBase.DBName}`.``tblstackcodes`` AS S "
                output += $"SET B.fldDipatchDate = NULL, B.fldDispatchID = NULL, S.fldDipatchDate = NULL, S.fldDispatchID = NULL "
                output += $"WHERE B.fldPrintCode in ({aggCodes}) "
                output += $"AND S.fldPrintCode in ({aggCodes});"
            Case Else
                Throw New NotImplementedException("ClearDispatchedQuery ui_type must range between 1-3")
        End Select

        Return output
    End Function

    Private Function ConfirmPaymentQuery(index As Integer, jsonID As Integer)
        Return $"UPDATE `{DBName}`.`tblpayments` SET fldRep = NOW(), fldJsonID = {jsonID} WHERE fldIndex = {index};"
    End Function

    Private Function SelectCodesForIdQuery(IDs As Integer())
        Dim strIDs As String = $"'{String.Join("', '", IDs)}'"
        Return $"SELECT * FROM `{DBName}`.`tblboxcodes` WHERE fldDispatchID in ({strIDs});"
    End Function

    Private Function SelectArivalsQuery() As String
        Return $"SELECT * FROM `{DBName}`.`tblarrival` WHERE fldRep IS NULL;"
    End Function

    Private Function SelectAggregatedCodesQuery(table As String)
        Return $"SELECT * FROM `{DBName}`.`{table}` USE INDEX (AgregatedRep_AgregatedDate_idx) WHERE fldAgregatedDate IS NOT NULL AND fldAggregatedRep IS NULL;"
    End Function

    Private Function SelectPrintedCodesQuery(table As String)
        Return $"SELECT * FROM `{DBName}`.`{table}` USE INDEX (primarycodes_printdate_rep) WHERE fldPrintedDate IS NOT NULL AND fldPrintRep IS NULL LIMIT 5000;"
    End Function

    Private Function SelectDispatchedCodesQuery(table As String, dispatchID As Integer)
        Return $"SELECT * FROM `{DBName}`.`{table}` WHERE fldDispatchID = {dispatchID};"
    End Function

    Private Function SelectRecallQuery() As String
        Dim output As String = ""
        output += "SELECT r.*, j.fldRecallCode "
        output += $"FROM `{DBName}`.`tblrecall` AS r "
        output += $"LEFT JOIN (`{DBName}`.`tbljson` as j) "
        output += "ON j.fldIndex = r.fldTargetID "
        output += "WHERE r.fldRep IS NULL;"
        Return output
    End Function

    Private Function SelectHumanReadableLengthQuery(orderID As Integer) As String
        Dim output = "SELECT tblidissuers.fldPrimaryCodeLength " &
        $"FROM {DBName}.tblorders, {DBName}.tblidissuers " &
        $"WHERE tblorders.fldIndex = {orderID} " &
        "AND tblorders.fldIdIssuerUI = tblidissuers.fldUI;"
        Return output
    End Function

    Private Function UpdatePrintedCodeQuery(table As String, codes() As String, jsonID As Integer) As String
        Dim strCodes As String = "'" & String.Join("', '", codes) & "'"
        Return $"UPDATE `{DBName}`.`{table}` SET fldPrintRep = NOW(), fldPrintRepID = {jsonID}  WHERE fldPrintCode in ({strCodes});"
    End Function

    Private Function UpdateAggregatedCoderQuery(table As String, codes() As String, jsonID As Integer) As String
        Dim strCodes As String = "'" & String.Join("', '", codes) & "'"
        Return $"UPDATE `{DBName}`.`{table}` Set fldAggregatedRep = NOW(), fldAggRepID = {jsonID} WHERE fldPrintCode In ({strCodes});"
    End Function

    Private Function SelectInvoicesQuery() As String
        Dim output As String = ""
        output += "Select i.fldIndex, i.fldEventTime, i.fldtype, i.fldOtherType, i.fldInvoiceNumber, i.fldDate, "
        output += "i.fldSellerID, i.fldBuyerID, "
        output += "EO.fldEO_Name1 As fldBuyer_Name, EO.fldEO_Address As fldBuyer_Address, EO.fldEO_AddressStreet1 as fldBuyerStreet1, EO.fldEO_AddressStreet2 as fldBuyerStreet2, EO.fldEO_AddressCity as fldBuyerCity, EO.fldEO_AddressPostCode As fldBuyerPostCode, EO.fldEO_CountryReg As fldBuyer_CountryReg, EO.fldVAT_TAX_N As fldBuyer_Tax_N, "
        output += "i.fldEUBuyer, i.fldFirstSellerEU, i.fldOrderID ,i.fldValue, i.fldCurrency "
        output += $"FROM `{DBName}`.`tblinvoices` As i "
        output += $"LEFT JOIN (`{DBName}`.`tbleo` As EO) "
        output += "On (EO.fldEO_ID = i.fldBuyerID) "
        output += "WHERE i.fldRep IS NULL;"
        Return output
    End Function

    Private Function SelectPaymentsQuery() As String
        Dim output As String = ""
        output += "SELECT P.fldIndex, P.fldEventTime, P.fldPaymentDate, P.fldPaymentType, P.fldPaymentAmount, P.fldPaymentCurrency, "
        output += "I.fldEUBuyer, EO.fldEO_ID, "
        output += "EO.fldEO_Name1 As fldBuyer_Name, EO.fldEO_Address As fldBuyer_Address, EO.fldEO_AddressStreet1 as fldBuyerStreet1, EO.fldEO_AddressStreet2 as fldBuyerStreet2, EO.fldEO_AddressCity as fldBuyerCity, EO.fldEO_AddressPostCode As fldBuyerPostCode, EO.fldEO_CountryReg As fldBuyer_CountryReg, EO.fldVAT_TAX_N As fldBuyer_Tax_N, "
        output += "P.fldPaymentRecipient, P.fldPaymentInvoice, P.fldInvoicePaid, I.fldOrderID, P.fldComment "
        output += $"FROM `{DBName}`.`tblpayments` as P "
        output += $"LEFT JOIN (`{DBName}`.`tbleo` AS EO, `{DBName}`.`tblinvoices` as I) "
        output += "ON (I.`fldinvoicenumber` = P.`fldInvoicePaid` AND EO.`fldEO_ID` = I.fldBuyerID) "
        output += "WHERE P.fldRep IS NULL;"
        Return output
    End Function

    Private Function SelectInvoiceProductQuery(orderID As String) As String
        Dim output As String = ""
        output += "Select a.fldIndex, a.fldOrderID, a.fldProductID, a.fldCount, b.fldPrice "
        output += $"FROM `{DBName}`.`tblorderproducts` As a "
        output += "LEFT JOIN tblproducts As b "
        output += "On b.fldFCode = a.fldProductID "
        output += $"WHERE fldOrderID = {orderID};"
        Return output
    End Function

    Private Function SelectUncheckedJsons() As String
        Return $"SELECT * FROM `{DBName}`.`tbljson` WHERE fldStatus = 0;"
    End Function
#End Region

#Region "Direct access"
    Public Function ReadDatabase(query As String) As DataTable
        cmd.CommandText = query
        adapter.SelectCommand = cmd
        Dim output As New DataTable

        Try
            If conn.State <> ConnectionState.Open Then conn.Open()
            adapter.Fill(output)
        Catch ex As Exception
            ReportTools.Output.Report($"Exception occured While reading from database: '{ex.Message}'")
        End Try

        Return output
    End Function

    Public Function Execute(query As String) As Boolean
        If query = String.Empty Then Return False
        Dim output As Boolean = False

        'Execute the query
        cmd.CommandText = query
        Try
            If conn.State <> ConnectionState.Open Then conn.Open()
            cmd.ExecuteNonQuery()
            output = True
        Catch ex As Exception
            ReportTools.Output.Report($"Exception occured while writing to Database: '{ex.Message}'; {Environment.NewLine}Query: {query}")
        End Try

        'Close connection and return the result
        Return output
    End Function

    Public Function ExecuteReturnIndex(query As String) As Integer
        If query = String.Empty Then Return -1
        Dim output As Integer = -1

        'Execute the query
        cmd.CommandText = query
        Try
            If conn.State <> ConnectionState.Open Then conn.Open()
            cmd.ExecuteNonQuery()
            output = cmd.LastInsertedId
        Catch ex As Exception
            ReportTools.Output.Report($"Exception occured while writing to Database: '{ex.Message}'; {Environment.NewLine}Query: {query}")
        End Try

        'Close connection and return the result
        Return output
    End Function

    Public Sub Disconnect()
        Try
            If conn.State <> ConnectionState.Closed Then conn.Close()
        Catch
        End Try
    End Sub
#End Region

End Class
