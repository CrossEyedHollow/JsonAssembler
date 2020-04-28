﻿Imports System.IO
Imports ReportTools

Module Main

    Dim db As DBManager
    Private jMan As JsonManager

    Dim statusManager As StatusManager

    Sub Main()

        Initialize()

        Dim stopWatch As Stopwatch = New Stopwatch()
        statusManager.Start()

        While True
            Dim eCount As Integer = 0
            Try
                stopWatch.Restart()

                'Cycle trough the tables in order
                eCount += RunPrimaryCodesTable()
                'ReportTime("Primary table check", stopWatch)
                eCount += RunStackCodesTable()
                'ReportTime("Stack table check", stopWatch)
                eCount += RunBoxCodesTable()
                'ReportTime("Box table check", stopWatch)
            Catch ex As Exception
                Output.Report($"Unexpected exception occured: {ex.Message}")
            End Try

            Try
                eCount += RunDispatchEvents()
                'ReportTime("Dispatch events check", stopWatch)
            Catch ex As Exception
                Output.Report($"Unexpected exception occured while processing dispatch event: {ex.Message}")
            End Try

            Try
                eCount += RunInvoices()
                'ReportTime("Dispatch events check", stopWatch)
            Catch ex As Exception
                Output.Report($"Unexpected exception occured while processing invoice event: {ex.Message}")
            End Try

            Try
                eCount += RunPayments()
                'ReportTime("Payment events check", stopWatch)
            Catch ex As Exception
                Output.Report($"Unexpected exception occured while processing payment event: {ex.Message}")
            End Try

            Try
                eCount += RunArrival()
            Catch ex As Exception
                Output.Report($"Unexpected exception occured while processing ERP event: {ex.Message}")

            End Try

            Try
                eCount += RunRecalls()
                'ReportTime("Recall events check", stopWatch)
            Catch ex As Exception
                Output.Report($"Unexpected exception occured while processing recall event: {ex.Message}")
            End Try
            db.Disconnect()

            Dim sleepTime As Integer = 30
            If eCount = 0 Then Output.ToConsole($"No new events (Search elapsed in {stopWatch.Elapsed.TotalSeconds}s), sleeping for {sleepTime}s")
            stopWatch.Stop()

            Threading.Thread.Sleep(TimeSpan.FromSeconds(sleepTime))
        End While
    End Sub

    Private Function RunArrival() As Integer
        Dim events As DataTable = db.CheckForArivals()
        If events.Rows.Count < 1 Then Return 0

        For Each row As DataRow In events.Rows
            Try
                Dim fldIndex As Integer = Convert.ToInt32(row("fldIndex"))
                Dim fldEventTime As Date = Convert.ToDateTime(row("fldEventTime"))
                Dim fldReturnType As Integer = Convert.ToInt32(row("fldReturnType"))
                Dim upUIs As String() = If(IsDBNull(row("fldUpUIs")), Nothing, CStr(row("fldUpUIs")).Split(","))
                Dim aUIs As String() = If(IsDBNull(row("fldAUIs")), Nothing, CStr(row("fldAUIs")).Split(","))
                Dim fldComment As String = If(IsDBNull(row("fldComment")), "", row("fldComment"))

                Dim uiType As Integer = -1

                If Not upUIs.IsNullOrEmpty() And aUIs.IsNullOrEmpty() Then
                    'Unit level only arrival
                    uiType = 1
                ElseIf upUIs.IsNullOrEmpty() And Not aUIs.IsNullOrEmpty() Then
                    'Aggregated level UIs only
                    uiType = 2
                ElseIf Not upUIs.IsNullOrEmpty() And Not aUIs.IsNullOrEmpty() Then
                    'Both types
                    uiType = 3
                Else
                    Throw New Exception("Both upUIs and aUIs columns are empty.")
                End If

                'Assemble the json
                Dim recallCode As String = Guid.NewGuid().ToString()
                Dim jsonBody As String = JsonAssembler.ERP(fldEventTime, fldReturnType, uiType, upUIs, aUIs, recallCode, fldComment)


                'Send json to the primary
                jMan.Post(jsonBody)
                Output.ToConsole("New arrival event sent to the Primary repository. Updating database...")

                Dim jsonIndex As Integer = db.InsertJson(jsonBody, "ERP", recallCode)
                db.ConfirmArrival(fldIndex, jsonIndex)
                'db.ClearDispatchment(uiType, upUIs, aUIs)
            Catch ex As Exception
                Output.Report($"Failed to process ERP event: {ex.Message}")
            End Try
        Next
        Return 1
    End Function

    Private Function RunPayments() As Integer
        Dim events As DataTable = db.CheckForPayments()
        If events.Rows.Count < 1 Then Return 0

        For Each row As DataRow In events.Rows
            Try
                Dim fldIndex As Integer = Convert.ToInt32(row("fldIndex"))
                Dim fldEventTime As Date = Convert.ToDateTime(row("fldEventTime"))
                Dim fldPaymentDate As Date = Convert.ToDateTime(row("fldPaymentDate"))
                Dim fldPaymentType As Integer = Convert.ToInt32(row("fldPaymentType"))
                Dim fldPaymentAmount As Decimal = Convert.ToDecimal(row("fldPaymentAmount"))
                Dim fldPaymentCurrency As String = row("fldPaymentCurrency")
                Dim fldEUBuyer As Integer = Convert.ToInt32(row("fldEUBuyer"))
                Dim fldEO_ID As String = row("fldEO_ID")
                Dim fldBuyer_Name As String = CStr(row("fldBuyer_Name"))
                Dim fldBuyer_Address As String = CStr(row("fldBuyer_Address"))
                Dim fldBuyer_Street1 As String = row("fldBuyerStreet1")
                Dim fldBuyer_Street2 As String = row("fldBuyerStreet2")
                Dim fldBuyer_City As String = row("fldBuyerCity")
                Dim fldBuyer_PostCode As String = row("fldBuyerPostCode")
                Dim fldBuyer_CountryReg As String = row("fldBuyer_CountryReg")
                Dim fldBuyer_Tax_N As String = row("fldBuyer_Tax_N")
                Dim fldPaymentRecipient As String = row("fldPaymentRecipient")
                Dim fldPaymentInvoice As Integer = Convert.ToInt32(row("fldPaymentInvoice"))
                Dim fldInvoicePaid As String = row("fldInvoicePaid")
                Dim fldComment As String = If(IsDBNull(row("fldComment")), "", row("fldComment"))
                Dim fldOrderID As String = row("fldOrderID")

                'To get the codes, first get all deployment rolls = fldIndex
                Dim deploymentsInInvoice As DataTable = db.SelectDeploymentsForInvoice(fldInvoicePaid)
                If deploymentsInInvoice.Rows.Count < 1 Then Throw New Exception($"There are no matches for fldInvoiceID = {fldPaymentInvoice} in tbldeployment.")

                'Take only the dispatchID out of the deployments
                Dim arrDeployments As Integer() = deploymentsInInvoice.Rows.OfType(Of DataRow).Select(Function(dr) dr.Field(Of Integer)("fldIndex")).ToArray()

                'Get the HIGHEST aggregation level codes matching dispatchID
                Dim codes As DataTable = db.GetCodesForIDs(arrDeployments)
                If codes.Rows.Count < 1 Then Throw New Exception($"There are no matches for fldOrderID in ('{String.Join("', '", arrDeployments)}') in tblboxcodes.")

                Dim codesArray = GetAllCodes(codes, "fldPrintCode")

                'Assemble the json
                Dim recallCode As String = Guid.NewGuid().ToString()
                Dim jsonBody As String = JsonAssembler.EPR(fldEventTime, fldPaymentDate, fldPaymentType, fldPaymentAmount, fldPaymentCurrency, fldEUBuyer,
                                                           fldEO_ID, fldBuyer_Name, fldBuyer_Address, fldBuyer_Street1, fldBuyer_Street2, fldBuyer_City, fldBuyer_PostCode, fldBuyer_CountryReg, fldBuyer_Tax_N,
                                                           fldPaymentRecipient, fldPaymentInvoice, fldInvoicePaid, recallCode,
                                                           AggregationType.Aggregated_Only, Nothing, codesArray, fldComment)

                'Send json to the primary
                jMan.Post(jsonBody)
                Output.ToConsole("New invoice sent to the Primary repository. Updating database...")

                Dim jsonIndex As Integer = db.InsertJson(jsonBody, "EPR", recallCode)
                db.ConfirmPayment(fldIndex, jsonIndex)
            Catch ex As Exception
                Output.Report($"Failed to process EPR event: {ex.Message}")
            End Try
        Next

        Return 1
    End Function

    Private Function RunRecalls() As Integer
        Dim recalls As DataTable = db.CheckForRecalls()
        If recalls.Rows.Count < 1 Then Return 0

        For Each rcl As DataRow In recalls.Rows
            Try
                Dim fldIndex As Integer = Convert.ToInt32(rcl("fldIndex"))
                Dim fldTargetCode As String = Convert.ToInt32(rcl("fldTargetID"))
                Dim fldRecallReason1 As Integer = Convert.ToInt32(rcl("fldRecallReason1"))
                Dim fldRecallReason2 As String = Convert.ToString(rcl("fldRecallReason2"))
                Dim fldRecallReason3 As String = Convert.ToString(rcl("fldRecallReason3"))
                Dim recallCode As String = Guid.NewGuid().ToString()

                Dim jsonBody As String = JsonAssembler.RCL(fldTargetCode, fldRecallReason1, recallCode, fldRecallReason2, fldRecallReason3)
                jMan.Post(jsonBody)
                Output.ToConsole($"New recall event sent, updating database.")

                'Update the db
                Dim jsonIndex As Integer = db.InsertJson(jsonBody, "RCL", recallCode)
                db.ConfirmRecall(fldIndex, jsonIndex)
            Catch ex As Exception
                Output.Report($"Failed to process RCL event: {ex.Message}")
            End Try
        Next
        Return 1
    End Function

    Private Function RunDispatchEvents() As Integer
        'Check for new events and return if none are found
        Dim events As DataTable = db.CheckForDispatchEvent()
        If events.Rows.Count < 1 Then Return 0

        For Each row As DataRow In events.Rows
            Try
                'Gather variables
                Dim fldIndex As Integer = Convert.ToInt32(row("fldIndex"))
                Dim fldEventTime As Date = Convert.ToDateTime(row("fldEventTime"))
                Dim fldDestID1 As Integer = Convert.ToInt32(row("fldDestID1"))
                Dim fldDestID2 As String = Convert.ToString(row("fldDestID2"))
                Dim fldDestID3 As String = Convert.ToString(row("fldDestID3"))
                Dim fldDestID4 As String = Convert.ToString(row("fldDestID4"))
                Dim fldDestAddress As String = Convert.ToString(row("fldDestAddress"))
                Dim fldDestinationStreet1 As String = row("fldDestinationStreet1")
                Dim fldDestinationStreet2 As String = row("fldDestinationStreet2")
                Dim fldDestinationCity As String = row("fldDestinationCity")
                Dim fldDestinationPostCode As String = row("fldDestPostCode")
                Dim fldTransportMode As Integer = Convert.ToInt32(row("fldTransportMode"))
                Dim fldTransportVehicle As String = Convert.ToString(row("fldTransportVehicle"))
                Dim fldTransportCont1 As Integer = Convert.ToInt32(row("fldTransportCont1"))
                Dim fldTransporCont2 As String = Convert.ToString(row("fldTransporCont2"))
                Dim fldTransportS1 As Integer = Convert.ToInt32(row("fldTransportS1"))
                Dim fldTransportS2 As String = Convert.ToString(row("fldTransportS2"))
                Dim fldEMCS As Integer = Convert.ToInt32(row("fldEMCS"))
                Dim fldEMCS_ARC As String = Convert.ToString(row("fldEMCS_ARC"))
                Dim fldSAAD As Integer = Convert.ToInt32(row("fldSAAD"))
                Dim fldSAAD_Num As String = Convert.ToString(row("fldSAAD_Num"))
                Dim fldExpDeclaration As Integer = Convert.ToInt32(row("fldExpDeclaration"))
                Dim fldExpDeclNumber As String = Convert.ToString(row("fldExpDeclNumber"))
                Dim fldUI_Type As Integer = Convert.ToInt32(row("fldUI_Type"))
                Dim fldComment As String = Convert.ToString(row("fldComment"))

                Dim aUIs As List(Of String) = New List(Of String)
                Dim upUIs As String() = Nothing
                Dim stackCodes As String()
                Dim boxCodes As String()

                Select Case CType(fldUI_Type, AggregationType)
                    Case AggregationType.Unit_Packets_Only
                        'Get the codes
                        upUIs = db.GetDispatchedCodes("tblprimarycodes", fldIndex)
                    Case AggregationType.Aggregated_Only
                        'Get the codes
                        boxCodes = db.GetDispatchedCodes("tblboxcodes", fldIndex)
                        stackCodes = db.GetDispatchedCodes("tblstackcodes", fldIndex)
                        'Add them to the stack
                        aUIs.AddRange(boxCodes)
                        aUIs.AddRange(stackCodes)
                    Case AggregationType.Both
                        'Get the codes
                        boxCodes = db.GetDispatchedCodes("tblboxcodes", fldIndex)
                        stackCodes = db.GetDispatchedCodes("tblstackcodes", fldIndex)
                        upUIs = db.GetDispatchedCodes("tblprimarycodes", fldIndex)
                        'Add them to the stack
                        aUIs.AddRange(boxCodes)
                        aUIs.AddRange(stackCodes)
                    Case Else
                        Throw New NotImplementedException($"UI_Type: {fldUI_Type} does not exist. Ui_Type range [1-3]")
                End Select

                'Assemble json
                Dim recallCode As String = Guid.NewGuid().ToString()
                Dim jsonBody As String = EDP(fldEventTime, fldDestID1, fldDestID2, StringToArray(fldDestID3), StringToArray(fldDestID4), fldDestAddress, fldDestinationStreet1, fldDestinationStreet2, fldDestinationCity, fldDestinationPostCode,
                                             fldTransportMode, fldTransportVehicle, fldTransportCont1, fldTransporCont2, fldTransportS1, fldTransportS2,
                                              fldEMCS, fldEMCS_ARC, fldSAAD, fldSAAD_Num, fldExpDeclaration, fldExpDeclNumber, fldUI_Type, recallCode, upUIs, aUIs.ToArray(), fldComment)

                'Send json to the primary
                jMan.Post(jsonBody)
                Output.ToConsole("New dispatch event sent to the Primary repository. Updating database...")

                'Update db
                Dim jsonIndex As Integer = db.InsertJson(jsonBody, "EDP", recallCode)
                db.ConfirmDispatchEvent(fldIndex, jsonIndex)
            Catch ex As Exception
                Output.Report($"Failed to dispatch event: {ex.Message}")
            End Try
        Next
        Return 1
    End Function

    Private Function RunInvoices() As Integer
        'Check for NEW invoices and return if none are found
        Dim invoices As DataTable = db.CheckForInvoice()
        If invoices.Rows.Count < 1 Then Return 0

        For Each invoice As DataRow In invoices.Rows
            Try
                'Get the variables
                Dim fldIndex As Integer = Convert.ToInt32(invoice("fldIndex"))
                Dim fldEventTime As Date = Convert.ToDateTime(invoice("fldEventTime"))
                Dim fldType As Integer = Convert.ToInt32(invoice("fldType"))
                Dim fldOtherType As String = invoice("fldOtherType")
                Dim fldInvoiceNumber As String = invoice("fldInvoiceNumber")
                Dim fldDate As Date = Convert.ToDateTime(CStr(invoice("fldDate")))
                Dim fldSellerID As String = invoice("fldSellerID")
                Dim fldBuyerID As String = invoice("fldBuyerID")
                Dim fldBuyer_Name As String = invoice("fldBuyer_Name")
                Dim fldBuyer_Address As String = invoice("fldBuyer_Address")
                Dim fldBuyer_Street1 As String = invoice("fldBuyerStreet1")
                Dim fldBuyer_Street2 As String = invoice("fldBuyerStreet2")
                Dim fldBuyer_City As String = invoice("fldBuyerCity")
                Dim fldBuyer_PostCode As String = invoice("fldBuyerPostCode")
                Dim fldBuyer_CountryReg As String = invoice("fldBuyer_CountryReg")
                Dim fldBuyer_Tax_N As String = invoice("fldBuyer_Tax_N")
                Dim fldEUBuyer As Integer = Convert.ToInt32(CStr(invoice("fldEUBuyer")))
                Dim fldFirstSellerEU As Boolean = Convert.ToBoolean(invoice("fldFirstSellerEU"))
                Dim fldValue As String = (CStr(invoice("fldValue"))).Replace(",", ".")
                Dim fldCurrency As String = invoice("fldCurrency")
                Dim fldOrderID As String = invoice("fldOrderID")
                Dim fldProductIDs As String() = Nothing
                Dim fldProductCounts As Integer() = Nothing
                Dim fldProductPrices As Decimal() = Nothing

                'To get the codes, first get all deployment rolls = fldIndex
                Dim deploymentsInInvoice As DataTable = db.SelectDeploymentsForInvoice(fldInvoiceNumber)
                If deploymentsInInvoice.Rows.Count < 1 Then Throw New Exception($"There are no matches for fldInvoiceID = {fldInvoiceNumber} in tbldeployment.")

                'Take only the dispatchID out of the deployments
                Dim arrDeployments As Integer() = deploymentsInInvoice.Rows.OfType(Of DataRow).Select(Function(dr) dr.Field(Of Integer)("fldIndex")).ToArray()

                'Get the HIGHEST aggregation level codes matching dispatchID
                Dim codes As DataTable = db.GetCodesForIDs(arrDeployments)
                If codes.Rows.Count < 1 Then Throw New Exception($"There are no matches for fldOrderID in ('{String.Join("', '", arrDeployments)}') in tblboxcodes.")

                'If the seller IS in the EU, get the product TPIDs, ProductNumbers and prices
                If fldFirstSellerEU Then
                    Throw New NotImplementedException("fldFirstSellerEU = true, code not ready yet...")
                    'TODO, productIDs and numbers must be taken from the codes, not the orderproducts table, CHANGE ASAP
                    'Get the productIDs, count and price for this OrderID
                    Dim products As DataTable = db.GetOrderProducts(fldOrderID)
                    If products.Rows.Count < 1 Then Throw New Exception($"There are no matches for fldOrderID = '{fldOrderID}' in tblorderproducts.")

                    'LINQ Magic, converts All values in a column into an array 
                    fldProductIDs = products.ColumnToArray("fldProductID")
                    fldProductCounts = products.Rows.OfType(Of DataRow).Select(Function(dr) dr.Field(Of Integer)("fldCount")).ToArray()
                    fldProductPrices = products.Rows.OfType(Of DataRow).Select(Function(dr) dr.Field(Of Decimal)("fldPrice")).ToArray()
                End If

                Dim codesArray = GetAllCodes(codes, "fldPrintCode")

                'Assemble the json
                Dim recallCode As String = Guid.NewGuid().ToString()
                Dim jsonBody As String = JsonAssembler.EIV(fldEventTime, fldType, fldOtherType, fldInvoiceNumber, fldDate,
                                                           fldSellerID, fldEUBuyer, fldBuyerID, fldBuyer_Name,
                                                           fldBuyer_Address, fldBuyer_Street1, fldBuyer_Street2, fldBuyer_City, fldBuyer_PostCode, fldBuyer_CountryReg, fldBuyer_Tax_N, fldFirstSellerEU,
                                                           fldProductIDs, fldProductCounts, fldProductPrices, fldValue, fldCurrency,
                                                           AggregationType.Aggregated_Only, recallCode, Nothing, codesArray)

                'Send json to the primary
                jMan.Post(jsonBody)
                Output.ToConsole("New invoice sent to the Primary repository. Updating database...")

                Dim jsonIndex As Integer = db.InsertJson(jsonBody, "EIV", recallCode)
                db.ConfirmInvoice(fldIndex, jsonIndex)
            Catch ex As Exception
                Output.Report($"Invoice proccesing failed: {ex.Message}")
            End Try
        Next
        Return 1
    End Function

    Private Function RunPrimaryCodesTable() As Integer
        Dim table As String = "tblprimarycodes"
        Dim result As Integer = 0
        'Check for printed first
        Dim dtResult As DataTable = db.CheckForPrintedCodes(table)
        'If there are any
        If dtResult.Rows.Count > 0 Then
            'Get the codes
            Dim longUIs As String() = GetAllCodes(dtResult, "fldPrintCode")
            Dim shortUIs As String() = GetAllCodes(dtResult, "fldCode")

            Try
                'Assemble JSON
                Dim fldEventTime As Date = Convert.ToDateTime(dtResult("fldPrintedDate"))

                Dim recallCode As String = Guid.NewGuid().ToString()
                Dim jsonBody As String = JsonAssembler.EUA(fldEventTime, longUIs, shortUIs, recallCode)

                'Send report
                jMan.Post(jsonBody)
                Output.ToConsole("Application of unit level UIs on unit packets event sent... updating DB.")

                'Update database
                Dim jsonIndex As Integer = db.InsertJson(jsonBody, "EUA", recallCode)
                db.ConfirmPrintedCodes(table, longUIs, jsonIndex)
            Catch ex As Exception
                Output.Report($"Exception occured while posting JSON: {ex.Message}")
            End Try
            result = 1
        Else
            result = 0
        End If

        result += DoAggregationEvent(table, AggregationType.Unit_Packets_Only)
        Return result
    End Function

    Private Function RunBoxCodesTable() As Integer
        Dim table As String = "tblboxcodes"
        'DoPrintingEvent(table)

        'Aggregation is not yet implemented for Boxes
        Return DoAggregationEvent(table, AggregationType.Aggregated_Only)
    End Function

    Private Function RunStackCodesTable() As Integer
        Dim table As String = "tblstackcodes"
        'DoPrintingEvent(table)

        Return DoAggregationEvent(table, AggregationType.Aggregated_Only)
    End Function

    Private Sub DoPrintingEvent(table As String)
        Throw New Exception("DoPrintingEvent function is obsolete")
        ''Check for printed first
        'Dim printedCodes As DataTable = db.CheckForPrintedCodes(table)
        ''If there are any
        'If printedCodes.Rows.Count > 0 Then
        '    'Get the codes
        '    Dim onlyCodes As String() = GetAllCodes(printedCodes, "fldPrintCode")

        '    Try

        '        'Assemble JSON
        '        Dim fldEventTime As Date = Convert.ToDateTime(printedCodes("fldPrintedDate"))
        '        Dim recallCode As String = Guid.NewGuid().ToString()
        '        Dim jsonBody As String = JsonAssembler.EUA(fldEventTime, onlyCodes, onlyCodes, recallCode)

        '        'Send report
        '        jMan.Post(jsonBody)
        '        Output.ToConsole("Application of unit level UIs on unit packets event sent... updating DB.")

        '        'Update database
        '        Dim jsonIndex As Integer = db.InsertJson(jsonBody, "EUA", recallCode)
        '        db.ConfirmPrintedCodes(table, onlyCodes, jsonIndex)
        '    Catch ex As Exception
        '        Output.Report($"Exception occured while posting JSON: {ex.Message}")
        '    End Try
        'Else
        '    Output.ToConsole($"No new printing events in {table}.")
        'End If
    End Sub

    Private Function DoAggregationEvent(table As String, aggType As AggregationType) As Integer
        'After that check for aggregated codes
        Dim aggregatedCodes As DataTable = db.CheckForAggregatedCodes(table)
        If aggregatedCodes.Rows.Count > 0 Then

            'Get the DISTINCT aUIs
            Dim distinctParents = aggregatedCodes.DefaultView.ToTable(True, "fldParentCode")

            'For each aUI
            For Each row As DataRow In distinctParents.Rows
                Try

                    Dim parent As String = row("fldParentCode")
                    Dim fldEventTime As Date = Convert.ToDateTime(aggregatedCodes("fldAgregatedDate"))
                    Dim upUIs As String() = Nothing
                    Dim aUIs As String() = Nothing
                    Dim aUI As String

                    Select Case table
                        Case "tblprimarycodes"
                            aUI = GetPrintedCode("tblstackcodes", parent)
                        Case "tblstackcodes"
                            aUI = GetPrintedCode("tblboxcodes", parent)
                        Case "tblboxcodes"
                            aUI = parent
                        Case Else
                            Throw New NotImplementedException($"'{table}' is not a correct value for function 'DoAggregationEvent'.")
                    End Select

                    'Get a list of the children
                    Dim children As List(Of String) = GetChildren(parent, aggregatedCodes)


                    Select Case aggType
                        Case AggregationType.Aggregated_Only
                            aUIs = children.ToArray()
                        Case AggregationType.Unit_Packets_Only
                            upUIs = children.ToArray()
                        Case Else
                            Throw New NotImplementedException($"{AggregationType.Both.ToString()} not implemented.")
                    End Select

                    'Assemble JSON
                    Dim recallCode As String = Guid.NewGuid().ToString()
                    Dim jsonBody As String = JsonAssembler.EPA(fldEventTime, aUI, aggType, recallCode, upUIs, aUIs)

                    'Send report
                    jMan.Post(jsonBody)
                    Output.ToConsole("Message to report an aggregation event sent... updating DB.")

                    'Update database
                    Dim jsonIndex As Integer = db.InsertJson(jsonBody, "EPA", recallCode)
                    db.ConfirmAggregatedCodes(table, children.ToArray(), jsonIndex)
                Catch ex As Exception
                    Output.Report($"Exception occured while posting JSON: {ex.Message}")
                End Try
            Next
        Else
            Return 0
        End If
        Return 1
    End Function

    Private Function GetPrintedCode(table As String, parent As String) As String
        Dim dtResult As DataTable = db.GetPrintedCode("table", parent)
        If dtResult.Rows.Count <> 1 Then
            Throw New Exception($"Unexpected result from database for parent code: '{parent}', expected 1 row, returned rows: {dtResult.Rows.Count}")
        Else
            Return dtResult(0)("fldPrintCode")
        End If
    End Function

    Private Sub Initialize()
        Settings = New DataSet()
        Settings.ReadXml($"{AppDomain.CurrentDomain.BaseDirectory}Settings.xml")

        'Initialize the DBManager objects
        Dim dbSetting As DataRow = Settings.Tables("tblDBSettings").Rows(0)
        DBBase.DBName = dbSetting("fldDBName")
        DBBase.DBIP = dbSetting("fldServer")
        DBBase.DBUser = dbSetting("fldAccount")
        DBBase.DBPass = dbSetting("fldPassword")
        db = New DBManager() ' The constructor calls the DBBase.Init()

        'Init the JsonManager
        Dim jsonSetting As DataRow = Settings.Tables("tblJSONServer").Rows(0)
        Dim url As String = jsonSetting("fldPostAddress")

        jMan = New JsonManager(url)
        statusManager = New StatusManager(url)

        'Init the general settings
        Dim generalSettings As DataRow = Settings.Tables("tblGeneral").Rows(0)
        Dim eoID = generalSettings("fldEO_ID")
        Dim fID = generalSettings("fldF_ID")
        JsonAssembler.EO_ID = eoID
        JsonAssembler.F_ID = fID
    End Sub

#Region "Helpers"
    Private Sub ReportTime(text As String, wt As Stopwatch)
        Console.WriteLine($"{text} elapsed in {wt.Elapsed.TotalSeconds}s")
        wt.Restart()
    End Sub

    Private Function StringToArray(str As String)
        Return str.Replace(" ", "").Split(",")
    End Function

    ''' <summary>
    ''' Uses some LINQ magic to extract a list consisting of the children of the parent from a datatable
    ''' </summary>
    ''' <returns></returns>
    Public Function GetChildren(parent As String, haystack As DataTable) As List(Of String)
        Dim view = haystack.DefaultView
        view.RowFilter = $"fldParentCode = '{parent}'"
        Return view.ToTable().ColumnToArray("fldPrintCode").ToList()
    End Function

    Private Function GetRawCodes(codes As DataTable) As String()
        Dim output As List(Of String) = New List(Of String)
        For Each row As DataRow In codes.Rows
            'Save the variables
            Dim orderID As Integer = row("fldOrderID")
            Dim printedCode As String = row("fldPrintCode")
            'Find the length of the human readable part
            Dim hrLength As Integer = db.GetHumanReadableCodeLength(orderID)
            'Get the human readable part of the code
            Dim hrPart = printedCode.Substring(0, hrLength)
            'Add to the output
            output.Add(hrPart)
        Next
        'Return as array
        Return output.ToArray()
    End Function

    Public Function GetAllCodes(table As DataTable, column As String) As String()
        Return table.ColumnToArray(column)
    End Function
#End Region
End Module