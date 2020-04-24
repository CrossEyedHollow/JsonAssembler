Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Module JsonOperationals

    Public Property EO_ID As String = ""
    Public Property F_ID As String = ""

#Region "Operationals"
    ''' <summary>
    ''' Application of unit level UIs on unit packets event
    ''' </summary>
    ''' <returns></returns>
    Public Function EUA(eventTime As Date, upUI_1() As String, upUI_2() As String, code As String, Optional comment As String = "none") As String
        Dim json As JObject = New JObject()

        json("EO_ID") = EO_ID
        json("F_ID") = F_ID
        json("Event_Time") = GetTime(eventTime)
        json("Message_Time_long") = GetTimeLong()
        json("upUI_1") = JArray.FromObject(upUI_1)
        json("upUI_2") = JArray.FromObject(upUI_2)
        json("upUI_comment") = comment
        json("Message_Type") = "EUA"
        json("Code") = code

        Return json.ToString(Formatting.Indented)
    End Function

    ''' <summary>
    ''' Message to report an aggregation event
    ''' </summary>
    ''' <returns></returns>
    Public Function EPA(eventTime As Date,
                        aUI As String,
                        AggregationType As AggregationType,
                        code As String,
                        Optional aggregatedUIs1() As String = Nothing,
                        Optional aggregatedUIs2() As String = Nothing,
                        Optional comment As String = "none") As String

        'Check validity of input
        CheckUITypeValidity(AggregationType, aggregatedUIs1, aggregatedUIs2)

        'Assemble
        Dim json As JObject = New JObject()
        json("EO_ID") = EO_ID
        json("F_ID") = F_ID
        json("Event_Time") = GetTime(eventTime)
        json("Message_Time_long") = GetTimeLong()
        json("aUI") = aUI
        json("Aggregation_Type") = CInt(AggregationType)
        json("Aggregated_UIs1") = JArray.FromObject(aggregatedUIs1)
        json("Aggregated_UIs2") = JArray.FromObject(aggregatedUIs2)
        json("aUI_comment") = comment
        json("Message_Type") = "EPA"
        json("Code") = code

        Return json.ToString(Formatting.Indented)
    End Function

    ''' <summary>
    '''  Message to report a dispatch event
    ''' </summary>
    ''' <returns></returns>
    Public Function EDP(eventTime As Date,
                        destinationID1 As Integer,
                        destinationID2 As String,
                        destinationID3() As String,
                        destinationID4() As String,
                        destinationAddress As String,
                        destinationStreet1 As String,
                        destinationStreet2 As String,
                        destinationCity As String,
                        destinationPostCode As String,
                        transport_mode As TransportMode,
                        transportVehicle As String,
                        transport_cont1 As Integer,
                        transport_cont2 As String,
                        transport_s1 As Integer,
                        transport_s2 As String,
                        EMCS As Integer,
                        EMCS_ARC As String,
                        SAAD As Integer,
                        SAAD_number As String,
                        Exp_Declaration As Integer,
                        Exp_DeclarationNumber As String,
                        ui_type As AggregationType,
                        code As String,
                        Optional upUIs() As String = Nothing,
                        Optional aUIs() As String = Nothing,
                        Optional comment As String = "none") As String

        'Check input validity
        CheckUITypeValidity(ui_type, upUIs, aUIs)
        CheckDestinationValidity(destinationID1, destinationID2, destinationID3, destinationID4, destinationAddress)
        CheckBooleanType(transport_cont1, transport_cont2, "Transport_cont1", "Transport_cont2")
        CheckBooleanType(transport_s1, transport_s2, "Transport_s1", "Transport_s2")
        CheckBooleanType(EMCS, EMCS_ARC, "EMCS", "EMCS_ARC")
        CheckBooleanType(SAAD, SAAD_number, "SAAD", "SAAD_number")
        CheckBooleanType(Exp_Declaration, Exp_DeclarationNumber, "Exp_Declaration", "Exp_DeclarationNumber")



        'Assemble
        Dim json As JObject = New JObject()
        json("EO_ID") = EO_ID
        json("F_ID") = F_ID
        json("Event_Time") = GetTime(eventTime)
        json("Message_Time_long") = GetTimeLong()
        json("Destination_ID1") = destinationID1
        json("Destination_ID2") = destinationID2
        json("Destination_ID3") = JArray.FromObject(destinationID3)
        json("Destination_ID4") = JArray.FromObject(destinationID4)
        json("Destination_ID5") = destinationAddress
        json("Destination_ID5_Address_StreetOne") = destinationStreet1
        json("Destination_ID5_Address_StreetTwo") = destinationStreet2
        json("Destination_ID5_Address_City") = destinationCity
        json("Destination_ID5_Address_PostCode") = destinationPostCode
        json("Transport_mode") = CInt(transport_mode)
        json("Transport_vehicle") = transportVehicle
        json("Transport_cont1") = transport_cont1
        json("Transport_cont2") = transport_cont2
        json("Transport_s1") = transport_s1
        json("Transport_s2") = transport_s2
        json("EMCS") = EMCS
        json("EMCS_ARC") = EMCS_ARC
        json("SAAD") = SAAD
        json("SAAD_number") = SAAD_number
        json("Exp_Declaration") = Exp_Declaration
        json("Exp_DeclarationNumber") = Exp_DeclarationNumber
        json("UI_Type") = CInt(ui_type)
        json("upUIs") = JArray.FromObject(upUIs)
        json("aUIs") = JArray.FromObject(aUIs)
        json("Dispatch_comment") = comment
        json("Message_Type") = "EDP"
        json("Code") = code

        ''Input the input
        'Dim output As String = "{" & vbLf &
        '    vbTab & $"""EO_ID"": ""{EO_ID}""," & vbLf &
        '    vbTab & $"""F_ID"": ""{F_ID}""," & vbLf &
        '    vbTab & $"""Event_Time"": ""{GetTime(eventTime)}""," & vbLf &
        '    vbTab & $"""Message_Time_long"": ""{GetTimeLong()}""," & vbLf &
        '    vbTab & $"""Destination_ID1"": {destinationID1}," & vbLf &
        '    vbTab & $"""Destination_ID2"": ""{destinationID2}""," & vbLf &
        '    vbTab & $"""Destination_ID3"": {StringArrayToJsonArray(destinationID3)}," & vbLf &
        '    vbTab & $"""Destination_ID4"": {StringArrayToJsonArray(destinationID4)}," & vbLf &
        '    vbTab & $"""Destination_ID5"": ""{destinationAddress}""," & vbLf &
        '    vbTab & $"""Destination_ID5_Address_StreetOne"": ""{destinationStreet1.ToJSON()}""," & vbLf &
        '    vbTab & $"""Destination_ID5_Address_StreetTwo"": {Format(destinationStreet2)}," & vbLf &
        '    vbTab & $"""Destination_ID5_Address_City"": ""{destinationCity}""," & vbLf &
        '    vbTab & $"""Destination_ID5_Address_PostCode"": {Format(destinationPostCode)}," & vbLf &
        '    vbTab & $"""Transport_mode"": {CInt(transport_mode)}," & vbLf &
        '    vbTab & $"""Transport_vehicle"": {Format(transportVehicle)}," & vbLf &
        '    vbTab & $"""Transport_cont1"": {transport_cont1}," & vbLf &
        '    vbTab & $"""Transport_cont2"": {Format(transport_cont2)}," & vbLf &
        '    vbTab & $"""Transport_s1"": {transport_s1.ToString()}," & vbLf &
        '    vbTab & $"""Transport_s2"": {Format(transport_s2)}," & vbLf &
        '    vbTab & $"""EMCS"": {EMCS.ToString()}," & vbLf &
        '    vbTab & $"""EMCS_ARC"": {Format(EMCS_ARC)}," & vbLf &
        '    vbTab & $"""SAAD"": {SAAD.ToString()}," & vbLf &
        '    vbTab & $"""SAAD_number"": {Format(SAAD_number)}," & vbLf &
        '    vbTab & $"""Exp_Declaration"": {Exp_Declaration.ToString()}," & vbLf &
        '    vbTab & $"""Exp_DeclarationNumber"": {Format(Exp_DeclarationNumber)}," & vbLf &
        '    vbTab & $"""UI_Type"": {CInt(ui_type)}," & vbLf &
        '    vbTab & $"""upUIs"": {StringArrayToJsonArray(upUIs)}," & vbLf &
        '    vbTab & $"""aUIs"": {StringArrayToJsonArray(aUIs)}," & vbLf &
        '    vbTab & $"""Dispatch_comment"": ""{comment}""," & vbLf &
        '    vbTab & """Message_Type"": ""EDP""," & vbLf &
        '    vbTab & $"""Code"": ""{code}""" & vbLf & "}"
        Return json.ToString(Formatting.Indented)
    End Function

    ''' <summary>
    ''' Message to report a reception event
    ''' </summary>
    ''' <returns></returns>
    Public Function ERP(eventTime As Date,
                        productType As Integer,
                        ui_type As AggregationType,
                        upUIs() As String,
                        aUIs() As String,
                        code As String,
                        Optional comment As String = "none") As String
        'Check input validity
        CheckUITypeValidity(ui_type, upUIs, aUIs)

        'Assemble
        Dim json As JObject = New JObject()
        json("EO_ID") = EO_ID
        json("F_ID") = F_ID
        json("Event_Time") = GetTime(eventTime)
        json("Message_Time_long") = GetTimeLong()
        json("Product_Return") = productType
        json("UI_Type") = CInt(ui_type)
        json("upUIs") = JArray.FromObject(upUIs)
        json("aUIs") = JArray.FromObject(aUIs)
        json("Arrival_comment") = comment
        json("Message_Type") = "ERP"
        json("Code") = code

        ''Apply the input 
        'Dim output As String = "{" & vbLf &
        '    vbTab & $"""EO_ID"": ""{EO_ID}""," & vbLf &
        '    vbTab & $"""F_ID"": ""{F_ID}""," & vbLf &
        '    vbTab & $"""Event_Time"": ""{GetTime(eventTime)}""," & vbLf &
        '    vbTab & $"""Message_Time_long"": ""{GetTimeLong()}""," & vbLf &
        '    vbTab & $"""Product_Return"": ""{productType.ToString()}""," & vbLf &
        '    vbTab & $"""UI_Type"": ""{CInt(ui_type)}""," & vbLf &
        '    vbTab & $"""upUIs"": {StringArrayToJsonArray(upUIs)}," & vbLf &
        '    vbTab & $"""aUIs"": {StringArrayToJsonArray(aUIs)}," & vbLf &
        '    vbTab & $"""Arrival_comment"": ""{comment}""," & vbLf &
        '    vbTab & """Message_Type"": ""ERP""," & vbLf &
        '    vbTab & $"""Code"": ""{code}""" & vbLf & "}"
        Return json.ToString(Formatting.Indented)
    End Function

    ''' <summary>
    ''' Message to report a trans-loading event
    ''' </summary>
    ''' <returns></returns>
    Public Function ETL(eventTime As Date,
                        destinationID1 As Integer,
                        destinationID2 As String,
                        destinationAddress As String,
                        destinationStreet1 As String,
                        destinationStreet2 As String,
                        destinationCity As String,
                        destinationPostCode As String,
                        transport_Mode As Integer,
                        transport_Vehicle As String,
                        transport_cont1 As Integer,
                        transport_cont2 As String,
                        EMCS As Integer,
                        ui_type As AggregationType,
                        code As String,
                        Optional upUIs() As String = Nothing,
                        Optional aUIs() As String = Nothing,
                        Optional EMCS_ARC As String = "null",
                        Optional comment As String = "none") As String

        'Check validity
        CheckUITypeValidity(ui_type, upUIs, aUIs)
        CheckDestinationValidity(destinationID1, destinationID2, destinationAddress)
        CheckBooleanType(transport_cont1, transport_cont2, "Transport_cont1", "Transport_cont2")
        CheckBooleanType(EMCS, EMCS_ARC, "EMCS", "EMCS_ARC")

        'Assemble
        Dim json As JObject = New JObject()
        json("EO_ID") = EO_ID
        json("Event_Time") = GetTime(eventTime)
        json("Message_Time_long") = GetTimeLong()
        json("Destination_ID1") = destinationID1
        json("Destination_ID2") = destinationID2
        json("Destination_ID3") = destinationAddress
        json("Destination_ID3_Address_StreetOne") = destinationStreet1
        json("Destination_ID3_Address_StreetTwo") = destinationStreet2
        json("Destination_ID3_Address_City") = destinationCity
        json("Destination_ID3_Address_PostCode") = destinationPostCode
        json("Transport_mode") = transport_Mode
        json("Transport_vehicle") = transport_Vehicle
        json("Transport_cont1") = transport_cont1
        json("Transport_cont2") = transport_cont2
        json("EMCS") = EMCS
        json("EMCS_ARC") = EMCS_ARC
        json("UI_Type") = CInt(ui_type)
        json("upUIs") = JArray.FromObject(upUIs)
        json("aUIs") = JArray.FromObject(aUIs)
        json("Transloading_comment") = comment
        json("Message_Type") = "ETL"
        json("Code") = code

        ''Input the input
        'Dim output As String = "{" & vbLf &
        '    vbTab & $"""EO_ID"": ""{EO_ID}""," & vbLf &
        '    vbTab & $"""Event_Time"": ""{GetTime(eventTime)}""," & vbLf &
        '    vbTab & $"""Message_Time_long"": ""{GetTimeLong()}""," & vbLf &
        '    vbTab & $"""Destination_ID1"": {destinationID1}," & vbLf &
        '    vbTab & $"""Destination_ID2"": ""{destinationID2}""," & vbLf &
        '    vbTab & $"""Destination_ID3"": ""{destinationAddress}""," & vbLf &
        '    vbTab & $"""Destination_ID3_Address_StreetOne"": ""{destinationStreet1.ToJSON()}""," & vbLf &
        '    vbTab & $"""Destination_ID3_Address_StreetTwo"": {Format(destinationStreet2.ToJSON())}," & vbLf &
        '    vbTab & $"""Destination_ID3_Address_City"": ""{destinationCity}""," & vbLf &
        '    vbTab & $"""Destination_ID3_Address_PostCode"": {Format(destinationPostCode)}," & vbLf &
        '    vbTab & $"""Transport_mode"": {transport_Mode}," & vbLf &
        '    vbTab & $"""Transport_vehicle"": {transport_Vehicle}," & vbLf &
        '    vbTab & $"""Transport_cont1"": {transport_cont1.ToString()}," & vbLf &
        '    vbTab & $"""Transport_cont2"": {Format(transport_cont2)}," & vbLf &
        '    vbTab & $"""EMCS"": {EMCS.ToString()}," & vbLf &
        '    vbTab & $"""EMCS_ARC"": {Format(EMCS_ARC)}," & vbLf &
        '    vbTab & $"""UI_Type"": {CInt(ui_type)}," & vbLf &
        '    vbTab & $"""upUIs"": {StringArrayToJsonArray(upUIs)}," & vbLf &
        '    vbTab & $"""aUIs"": {StringArrayToJsonArray(aUIs)}," & vbLf &
        '    vbTab & $"""Transloading_comment"": ""{comment}""," & vbLf &
        '    vbTab & """Message_Type"": ""ETL""," & vbLf &
        '    vbTab & $"""Code"": ""{code}""" & vbLf & "}"
        Return json.ToString(Formatting.Indented)
    End Function

    ''' <summary>
    ''' Message to report an UID disaggregation
    ''' </summary>
    ''' <returns></returns>
    Public Function EUD(eventTime As Date, aUI As String, code As String, Optional comment As String = "none") As String
        'Assemble
        Dim json As JObject = New JObject()
        json("EO_ID") = EO_ID
        json("F_ID") = F_ID
        json("Event_Time") = GetTime(eventTime)
        json("Message_Time_long") = GetTimeLong()
        json("aUI") = aUI
        json("disaUI_Comment") = comment
        json("Message_Type") = "EUD"
        json("Code") = code

        'Dim output As String = "{" & vbLf &
        '    vbTab & $"""EO_ID"": ""{EO_ID}""," & vbLf &
        '    vbTab & $"""F_ID"": ""{F_ID}""," & vbLf &
        '    vbTab & $"""Event_Time"": ""{GetTime(eventTime)}""," & vbLf &
        '    vbTab & $"""Message_Time_long"": ""{GetTimeLong()}""," & vbLf &
        '    vbTab & $"""aUI"": ""{aUI}""," & vbLf &
        '    vbTab & $"""disaUI_Comment"": ""{comment}""," & vbLf &
        '    vbTab & """Message_Type"": ""EUD""," & vbLf &
        '    vbTab & $"""Code"": ""{code}""" & vbLf & "}"
        Return json.ToString(Formatting.Indented)
    End Function

    ''' <summary>
    ''' Message to report the delivery carried out with a vending van to retail outlet
    ''' </summary>
    ''' <returns></returns>
    Public Function EVR(eventTime As Date,
                        ui_type As AggregationType,
                        upUIs() As String,
                        aUIs() As String,
                        code As String,
                        Optional comment As String = "none") As String
        'Check validity
        CheckUITypeValidity(ui_type, upUIs, aUIs)

        'Assemble
        Dim json As JObject = New JObject()
        json("EO_ID") = EO_ID
        json("F_ID") = F_ID
        json("Event_Time") = GetTime(eventTime)
        json("Message_Time_long") = GetTimeLong()
        json("UI_Type") = CInt(ui_type)
        json("upUIs") = JArray.FromObject(upUIs)
        json("aUIs") = JArray.FromObject(aUIs)
        json("Delivery_comment") = comment
        json("Message_Type") = "EVR"
        json("Code") = code

        ''Input the input
        'Dim output As String = "{" & vbLf &
        '    vbTab & $"""EO_ID"": ""{EO_ID}""," & vbLf &
        '    vbTab & $"""F_ID"": ""{F_ID}""," & vbLf &
        '    vbTab & $"""Event_Time"": ""{GetTime(eventTime)}""," & vbLf &
        '    vbTab & $"""Message_Time_long"": ""{GetTimeLong()}""," & vbLf &
        '    vbTab & $"""UI_Type"": {CInt(ui_type)}," & vbLf &
        '    vbTab & $"""upUIs"": {StringArrayToJsonArray(upUIs)}," & vbLf &
        '    vbTab & $"""aUIs"": [""{StringArrayToJsonArray(aUIs)}""]," & vbLf &
        '    vbTab & $"""Delivery_comment"": ""{comment}""," & vbLf &
        '    vbTab & """Message_Type"": ""EVR""," & vbLf &
        '    vbTab & $"""Code"": ""{code}""" & vbLf & "}"
        Return json.ToString(Formatting.Indented)
    End Function

    ''' <summary>
    ''' Message to request a UID deactivation
    ''' </summary>
    ''' <returns></returns>
    Public Function IDA(eventTime As Date,
                        deact_type As AggregationType,
                        deact_reason1 As DeactivationType,
                        upUI() As String,
                        code As String,
                        Optional aUI() As String = Nothing,
                        Optional deact_reason2 As String = "none",
                        Optional deact_reason3 As String = "none") As String
        'Check validity of input
        CheckDeactivationReason(deact_reason1, deact_reason2)
        CheckUITypeValidity(deact_type, upUI, aUI)

        'Assemble
        Dim json As JObject = New JObject()
        json("EO_ID") = EO_ID
        json("Event_Time") = GetTime(eventTime)
        json("Message_Time_long") = GetTimeLong()
        json("Deact_Type") = CInt(deact_type)
        json("Deact_Reason1") = CInt(deact_reason1)
        json("Deact_Reason2") = deact_reason2
        json("Deact_Reason3") = deact_reason3
        json("Deact_upUI") = JArray.FromObject(upUI)
        json("Deact_aUI") = JArray.FromObject(aUI)
        json("Code") = code
        json("Message_Type") = "IDA"

        Return json.ToString(Formatting.Indented)
    End Function

#End Region
#Region "Transactionals"
    Public Function EIV(eventTime As Date,
                        invoice_type1 As InvoiceType,
                        invoice_type2 As String,
                        invoice_number As String,
                        invoice_date As Date,
                        invoice_seller As String,
                        invoice_buyer1 As Integer,
                        invoice_buyer2 As String,
                        buyer_name As String,
                        buyer_address As String,
                        buyer_street1 As String,
                        buyer_street2 As String,
                        buyer_city As String,
                        buyer_PostCode As String,
                        buyer_countryreg As String,
                        buyer_tax_n As String,
                        first_seller_eu As Integer,
                        product_items1() As String,
                        product_items2() As Integer,
                        product_price() As Decimal,
                        invoice_net As String,
                        invoice_currency As String,
                        ui_type As AggregationType,
                        code As String,
                        Optional upUIs() As String = Nothing,
                        Optional aUIs() As String = Nothing,
                        Optional comment As String = "none") As String

        'Check validity of input
        CheckUITypeValidity(ui_type, upUIs, aUIs)
        CheckInvoiceType(invoice_type1, invoice_type2)
        CheckBooleanType(invoice_buyer1, invoice_buyer2, "Invoice_Buyer1", "Invoice_Buyer2")
        CheckBooleanTypeReverse(invoice_buyer1, buyer_name, "Invoice_Buyer1", "Buyer_Name")
        CheckBooleanTypeReverse(invoice_buyer1, buyer_address, "Invoice_Buyer1", "Buyer_Address")
        CheckBooleanTypeReverse(invoice_buyer1, buyer_countryreg, "Invoice_Buyer1", "Buyer_CountryReg")
        CheckBooleanTypeReverse(invoice_buyer1, buyer_tax_n, "Invoice_Buyer1", "Buyer_TAX_N")
        CheckBooleanType(first_seller_eu, product_items1, "First_Seller_EU", "Product_Items_1")
        CheckBooleanType(first_seller_eu, product_items2, "First_Seller_EU", "Product_Items_2")
        CheckBooleanType(first_seller_eu, product_price, "First_Seller_EU", "Product_Price")

        'Assemble
        Dim json As JObject = New JObject()
        json("EO_ID") = EO_ID
        json("Event_Time") = GetTime(eventTime)
        json("Message_Time_long") = GetTimeLong()
        json("Invoice_Type1") = CInt(invoice_type1)
        json("Invoice_Type2") = invoice_type2
        json("Invoice_Number") = invoice_number
        json("Invoice_Date") = GetDate(invoice_date)
        json("Invoice_Seller") = invoice_seller
        json("Invoice_Buyer1") = invoice_buyer1
        json("Invoice_Buyer2") = invoice_buyer2
        json("Buyer_Name") = buyer_name
        json("Buyer_CountryReg") = buyer_countryreg
        json("Buyer_Address") = buyer_address
        json("Buyer_Address_StreetOne") = buyer_street1
        json("Buyer_Address_StreetTwo") = buyer_street2
        json("Buyer_Address_City") = buyer_city
        json("Buyer_Address_PostCode") = buyer_PostCode
        json("Buyer_TAX_N") = buyer_tax_n
        json("First_Seller_EU") = first_seller_eu
        json("Product_Items_1") = JArray.FromObject(product_items1)
        json("Product_Items_2") = JArray.FromObject(product_items2)
        json("Product_Price") = JArray.FromObject(product_price)
        json("Invoice_Net") = invoice_net
        json("Invoice_Currency") = invoice_currency
        json("UI_Type") = ui_type
        json("upUIs") = JArray.FromObject(upUIs)
        json("aUIs") = JArray.FromObject(aUIs)
        json("Invoice_comment") = comment
        json("Message_Type") = "EIV"
        json("Code") = code

        ''Input the input
        'Dim output As String = "{" & vbLf &
        '    vbTab & $"""EO_ID"": ""{EO_ID}""," & vbLf &
        '    vbTab & $"""Event_Time"": ""{GetTime(eventTime)}""," & vbLf &
        '    vbTab & $"""Message_Time_long"": ""{GetTimeLong()}""," & vbLf &
        '    vbTab & $"""Invoice_Type1"": {CInt(invoice_type1)}," & vbLf &
        '    vbTab & $"""Invoice_Type2"": {Format(invoice_type2)}," & vbLf &
        '    vbTab & $"""Invoice_Number"": ""{invoice_number}""," & vbLf &
        '    vbTab & $"""Invoice_Date"": ""{GetDate(invoice_date)}""," & vbLf &
        '    vbTab & $"""Invoice_Seller"": ""{invoice_seller}""," & vbLf &
        '    vbTab & $"""Invoice_Buyer1"": {invoice_buyer1.ToString()}," & vbLf &
        '    vbTab & $"""Invoice_Buyer2"": {Format(invoice_buyer2)}," & vbLf &
        '    vbTab & $"""Buyer_Name"": {Format(buyer_name.ToJSON())}," & vbLf &
        '    vbTab & $"""Buyer_CountryReg"": ""{buyer_countryreg}""," & vbLf &
        '    vbTab & $"""Buyer_Address"": ""{buyer_address.ToJSON()}""," & vbLf &
        '    vbTab & $"""Buyer_Address_StreetOne"": ""{buyer_street1.ToJSON()}""," & vbLf &
        '    vbTab & $"""Buyer_Address_StreetTwo"": {Format(buyer_street2.ToJSON())}," & vbLf &
        '    vbTab & $"""Buyer_Address_City"": ""{buyer_city}""," & vbLf &
        '    vbTab & $"""Buyer_Address_PostCode"": {Format(buyer_PostCode)}," & vbLf &
        '    vbTab & $"""Buyer_TAX_N"": ""{buyer_tax_n}""," & vbLf &
        '    vbTab & $"""First_Seller_EU"": {first_seller_eu.ToString()}," & vbLf &
        '    vbTab & $"""Product_Items_1"": {StringArrayToJsonArray(product_items1)}," & vbLf &
        '    vbTab & $"""Product_Items_2"": {StringArrayToJsonArray(product_items2)}," & vbLf &
        '    vbTab & $"""Product_Price"": {StringArrayToJsonArray(product_price)}," & vbLf &
        '    vbTab & $"""Invoice_Net"": {invoice_net}," & vbLf &
        '    vbTab & $"""Invoice_Currency"": ""{invoice_currency}""," & vbLf &
        '    vbTab & $"""UI_Type"": {CInt(ui_type)}," & vbLf &
        '    vbTab & $"""upUIs"": {StringArrayToJsonArray(upUIs)}," & vbLf &
        '    vbTab & $"""aUIs"": {StringArrayToJsonArray(aUIs)}," & vbLf &
        '    vbTab & $"""Invoice_comment"": ""{comment}""," & vbLf &
        '    vbTab & """Message_Type"": ""EIV""," & vbLf &
        '    vbTab & $"""Code"": ""{code}""" & vbLf & "}"
        Return json.ToString(Formatting.Indented)
    End Function

    Public Function EPR(eventTime As Date,
                        payment_date As Date,
                        payment_type As Integer,
                        payment_amount As Double,
                        payment_currency As String,
                        payment_payer1 As Integer,
                        payment_payer2 As String,
                        payer_name As String,
                        payer_address As String,
                        payer_street1 As String,
                        payer_street2 As String,
                        payer_city As String,
                        payer_PostCode As String,
                        payer_countryreg As String,
                        payer_tax_n As String,
                        payment_recipient As String,
                        payment_invoice As Integer,
                        invoice_paid As String,
                        code As String,
                        Optional ui_type As AggregationType = Nothing,
                        Optional upUIs() As String = Nothing,
                        Optional aUIs() As String = Nothing,
                        Optional comment As String = "none") As String

        'Check input
        CheckBooleanType(payment_payer1, payment_payer2, "Payment_Payer1", "Payment_Payer2")
        CheckBooleanTypeReverse(payment_payer1, payer_name, "Payment_Payer1", "Payer_Name")
        CheckBooleanTypeReverse(payment_payer1, payer_address, "Payment_Payer1", "Payer_Address")
        CheckBooleanTypeReverse(payment_payer1, payer_countryreg, "Payment_Payer1", "Payer_CountryReg")
        CheckBooleanTypeReverse(payment_payer1, payer_tax_n, "Payment_Payer1", "Payer_TAX_N")
        CheckBooleanType(payment_invoice, invoice_paid, "Payment_Invoice", "Invoice_Paid")
        CheckBooleanTypeReverse(payment_invoice, ui_type, "Payment_Invoice", "UI_Type")
        If Not (IsDBNull(ui_type)) Then CheckUITypeValidity(ui_type, upUIs, aUIs)

        'Assemble
        Dim json As JObject = New JObject()
        json("EO_ID") = EO_ID
        json("Event_Time") = GetTime(eventTime)
        json("Message_Time_long") = GetTimeLong()
        json("Payment_Date") = GetDate(payment_date)
        json("Payment_Type") = payment_type
        json("Payment_Amount") = payment_amount
        json("Payment_Currency") = payment_currency
        json("Payment_Payer1") = payment_payer1
        json("Payment_Payer2") = payment_payer2
        json("Payer_Name") = payer_name
        json("Payer_Address") = payer_address
        json("Payer_Address_StreetOne") = payer_street1
        json("Payer_Address_StreetTwo") = payer_street2
        json("Payer_Address_City") = payer_city
        json("Payer_Address_PostCode") = payer_PostCode
        json("Payer_CountryReg") = payer_countryreg
        json("Payer_TAX_N") = payer_tax_n
        json("Payment_Recipient") = payment_recipient
        json("Payment_Invoice") = payment_invoice
        json("Invoice_Paid") = invoice_paid
        json("UI_Type") = CInt(ui_type)
        json("upUIs") = JArray.FromObject(upUIs)
        json("aUIs") = JArray.FromObject(aUIs)
        json("Payment_comment") = comment
        json("Message_Type") = "EPR"
        json("Code") = code

        ''Input the imput
        'Dim output As String = "{" & vbLf &
        '    vbTab & $"""EO_ID"": ""{EO_ID}""," & vbLf &
        '    vbTab & $"""Event_Time"": ""{GetTime(eventTime)}""," & vbLf &
        '    vbTab & $"""Message_Time_long"": ""{GetTimeLong()}""," & vbLf &
        '    vbTab & $"""Payment_Date"": ""{GetDate(payment_date)}""," & vbLf &
        '    vbTab & $"""Payment_Type"": {payment_type}," & vbLf &
        '    vbTab & $"""Payment_Amount"": {payment_amount}," & vbLf &
        '    vbTab & $"""Payment_Currency"": ""{payment_currency}""," & vbLf &
        '    vbTab & $"""Payment_Payer1"": {payment_payer1.ToString()}," & vbLf &
        '    vbTab & $"""Payment_Payer2"": {Format(payment_payer2)}," & vbLf &
        '    vbTab & $"""Payer_Name"": {Format(payer_name.ToJSON())}," & vbLf &
        '    vbTab & $"""Payer_Address"": {Format(payer_address.ToJSON())}," & vbLf &
        '    vbTab & $"""Payer_Address_StreetOne"": {Format(payer_street1.ToJSON())}," & vbLf &
        '    vbTab & $"""Payer_Address_StreetTwo"": {Format(payer_street2.ToJSON())}," & vbLf &
        '    vbTab & $"""Payer_Address_City"": {Format(payer_city.ToJSON())}," & vbLf &
        '    vbTab & $"""Payer_Address_PostCode"": {Format(payer_PostCode.ToJSON())}," & vbLf &
        '    vbTab & $"""Payer_CountryReg"": {Format(payer_countryreg)}," & vbLf &
        '    vbTab & $"""Payer_TAX_N"": {Format(payer_tax_n)}," & vbLf &
        '    vbTab & $"""Payment_Recipient"": ""{payment_recipient}""," & vbLf &
        '    vbTab & $"""Payment_Invoice"": {payment_invoice.ToString()}," & vbLf &
        '    vbTab & $"""Invoice_Paid"": {Format(invoice_paid)}," & vbLf &
        '    vbTab & $"""UI_Type"": {CInt(ui_type)}," & vbLf &
        '    vbTab & $"""upUIs"": {StringArrayToJsonArray(upUIs)}," & vbLf &
        '    vbTab & $"""aUIs"": {StringArrayToJsonArray(aUIs)}," & vbLf &
        '    vbTab & $"""Payment_comment"": ""{comment}""," & vbLf &
        '    vbTab & """Message_Type"": ""EPR""," & vbLf &
        '    vbTab & $"""Code"": ""{code}""" & vbLf & "}"
        Return json.ToString(Formatting.Indented)
    End Function

    Public Function EPO(eventTime As Date,
                        order_number As Integer,
                        order_date As Date,
                        ui_type As AggregationType,
                        code As String,
                        Optional upUIs() As String = Nothing,
                        Optional aUIs() As String = Nothing,
                        Optional comment As String = "none") As String
        'Check validity
        CheckUITypeValidity(ui_type, upUIs, aUIs)

        'Assemble
        Dim json As JObject = New JObject()
        json("EO_ID") = EO_ID
        json("Event_Time") = GetTime(eventTime)
        json("Message_Time_long") = GetTimeLong()
        json("Order_Number") = order_number.ToString()
        json("Order_Date") = GetDate(order_date)
        json("UI_Type") = CInt(ui_type)
        json("upUIs") = JArray.FromObject(upUIs)
        json("aUIs") = JArray.FromObject(aUIs)
        json("Comments") = comment
        json("Message_Type") = "EPO"
        json("Code") = code

        ''Input the input
        'Dim output As String = "{" & vbLf &
        '    vbTab & $"""EO_ID"": ""{EO_ID}""," & vbLf &
        '    vbTab & $"""Event_Time"": ""{GetTime(eventTime)}""," & vbLf &
        '    vbTab & $"""Message_Time_long"": ""{GetTimeLong()}""," & vbLf &
        '    vbTab & $"""Order_Number"": ""{order_number}""," & vbLf &
        '    vbTab & $"""Order_Date"": ""{GetTime(order_date)}""," & vbLf &
        '    vbTab & $"""UI_Type"": {CInt(ui_type)}," & vbLf &
        '    vbTab & $"""upUIs"": {StringArrayToJsonArray(upUIs)}," & vbLf &
        '    vbTab & $"""aUIs"": {StringArrayToJsonArray(aUIs)}," & vbLf &
        '    vbTab & $"""Comments"": ""{comment}""," & vbLf &
        '    vbTab & """Message_Type"": ""EPO""," & vbLf &
        '    vbTab & $"""Code"": ""{code}""" & vbLf & "}"

        Return json.ToString(Formatting.Indented)
    End Function

    Public Function RCL(recall_code As String,
                        recall_reason1 As RecallReasonType,
                        code As String,
                        Optional recall_reason2 As String = "none",
                        Optional comment As String = "none") As String

        CheckRecallReason(recall_reason1, recall_reason2)

        'Assemble
        Dim json As JObject = New JObject()
        json("EO_ID") = EO_ID
        json("Message_Time_long") = GetTimeLong()
        json("Recall_CODE") = recall_code
        json("Recall_Reason1") = recall_reason1
        json("Recall_Reason2") = recall_reason2
        json("Recall_Reason3") = comment
        json("Message_Type") = "RCL"
        json("Code") = code

        'Dim output As String = "{" & vbCrLf &
        '    vbTab & $"""EO_ID"": ""{EO_ID}""," & vbCrLf &
        '    vbTab & $"""Message_Time_long"": ""{GetTimeLong()}""," & vbLf &
        '    vbTab & $"""Recall_CODE"": ""{recall_code}""," & vbCrLf &
        '    vbTab & $"""Recall_Reason1"": {recall_reason1}," & vbCrLf &
        '    vbTab & $"""Recall_Reason2"": ""{recall_reason2}""," & vbCrLf &
        '    vbTab & $"""Recall_Reason3"": ""{comment}""," & vbCrLf &
        '    vbTab & """Message_Type"": ""RCL""," & vbCrLf &
        '    vbTab & $"""Code"": ""{code}""" & vbCrLf & "}"

        Return json.ToString(Formatting.Indented)
    End Function
#End Region
#Region "Custom"
    Public Function STA(recallCode As String) As String
        'Assemble
        Dim json As JObject = New JObject()
        json("Message_Type") = "RCL"
        json("Message_Time_long") = GetTimeLong()
        json("Code") = recallCode
        Return json.ToString(Formatting.Indented)
    End Function
#End Region
#Region "Helpers"
    Private Function StringArrayToJsonArray(array As Array) As String
        If IsDBNull(array) Then Return "null"
        If array Is Nothing Then Return "null"
        If array.Length = 0 Then Return "null"

        'Open the array
        Dim output As String = "["

        'Add all elements to the main string enclosed in double quotes, and followed by a comma
        Dim strArray As String() = array.Cast(Of String)
        For Each item As String In strArray
            output &= $"""{item}"", "
        Next

        'Remove the last comma and whitespace
        output = output.Remove(output.Length - 2, 2)
        'Close the array
        output &= "]"
        Return output
    End Function

    Private Function GetTime() As String
        Return Date.UtcNow.ToString("yyMMddHH")
    End Function
    Private Function GetTime(time As Date) As String
        Return time.ToString("yyMMddHH")
    End Function

    Public Function GetTimeLong() As String
        Return Date.UtcNow.ToString("yyyy-MM-ddThh:mm:ssZ")
    End Function
    Public Function GetTimeLong(t As Date) As String
        Return t.ToString("yyyy-MM-ddThh:mm:ssZ")
    End Function

    Private Function GetDate(dt As Date) As String
        Return dt.ToString("yyyy-MM-dd")
    End Function

    Private Sub CheckUITypeValidity(ui_type As AggregationType, upUIs() As String, aUIs() As String)
        Select Case ui_type
            Case 1
                If upUIs.IsNullOrEmpty() Then Throw New Exception($"ui_type = 1, expected an array Of upUIs.")
            Case 2
                If aUIs.IsNullOrEmpty() Then Throw New Exception($"ui_type = 2, expected an array Of aUIs.")
            Case 3
                If upUIs.IsNullOrEmpty() OrElse aUIs.IsNullOrEmpty() Then Throw New Exception($"ui_type = 3, expexted valid upUIs And aUIs arrays.")
            Case Else
                Throw New Exception("Invalid ui_type value:  " & ui_type)
        End Select
    End Sub

    ''' <summary>
    ''' Checks dependency for null or empty, if type is TRUE
    ''' </summary>
    ''' <param name="type"></param>
    ''' <param name="dependency"></param>
    ''' <param name="typeName"></param>
    ''' <param name="dependencyName"></param>
    Private Sub CheckBooleanType(type As Integer, dependency As String, typeName As String, dependencyName As String)
        'If the type is false, test passes
        If type = 0 Then Return
        If dependency.IsNullOrEmpty() Then Throw New Exception($"{dependencyName} must have valid value when {typeName} is 'true'")
    End Sub

    Private Sub CheckBooleanType(type As Boolean, dependency() As String, typeName As String, dependencyName As String)
        'If the type is false, test passes
        If Not type Then Return
        If dependency.IsNullOrEmpty() Then Throw New Exception($"{dependencyName} must have valid value when {typeName} is 'true'")
    End Sub
    Private Sub CheckBooleanType(type As Boolean, dependency() As Integer, typeName As String, dependencyName As String)
        'If the type is false, test passes
        If Not type Then Return
        If dependency.IsNullOrEmpty() Then Throw New Exception($"{dependencyName} must have valid value when {typeName} is 'true'")
    End Sub
    Private Sub CheckBooleanType(type As Boolean, dependency() As Decimal, typeName As String, dependencyName As String)
        'If the type is false, test passes
        If Not type Then Return
        If dependency.IsNullOrEmpty() Then Throw New Exception($"{dependencyName} must have valid value when {typeName} is 'true'")
    End Sub

    ''' <summary>
    ''' Check dependency for null or empty if type is FALSE
    ''' </summary>
    ''' <param name="type"></param>
    ''' <param name="dependency"></param>
    ''' <param name="typeName"></param>
    ''' <param name="dependencyName"></param>
    Private Sub CheckBooleanTypeReverse(type As Integer, dependency As String, typeName As String, dependencyName As String)
        If type = 1 Then Return
        If dependency.IsNullOrEmpty() Then Throw New Exception($"{dependencyName} must have valid value when {typeName} is 'false'")
    End Sub

    'Private Sub CheckBooleanTypeReverse(type As Boolean, dependency As AggregationType, typeName As String, dependencyName As String)
    '    If type Then Return
    '    If IsDBNull(dependency) OrElse dependency = Nothing Then Throw New Exception($"{dependencyName} must have valid value when {typeName} is 'false'")
    'End Sub

    Private Sub CheckDestinationValidity(destID1 As DestinationType, destID2 As String, destID3() As String, destID4() As String, destAddress As String)
        Select Case destID1
            Case 1
                If destAddress.IsNullOrEmpty() Then Throw New Exception($"Valid address expected for destination_id1 = {destID1}.")
            Case 2
                If destID2.IsNullOrEmpty() Then Throw New Exception($"Valid destination_id2 expected for destination_id1 = {destID1}.")
            Case 3
                If destID3.IsNullOrEmpty() Then Throw New Exception($"Valid destination_id3 expected for destination_id1 = {destID1}.")
            Case 4
                If destID4.IsNullOrEmpty() Then Throw New Exception($"Valid destination_id4 expected for destination_id1 = {destID1}.")
            Case Else
                Throw New NotImplementedException($"Wrong input for destination_id1: {destID1}, expected range: 1-4")
        End Select
    End Sub

    Private Sub CheckDeactivationReason(deact_reason1 As DeactivationType, deact_reason2 As String)
        If deact_reason1 > 6 OrElse deact_reason1 < 1 Then Throw New Exception("Deact_reason1 expected range is 1-6")
        If deact_reason1 = 6 AndAlso deact_reason2.IsNullOrEmpty() Then Throw New Exception("deact_reason2 is mandatory when deact_reason1 = 6 (Other)")
    End Sub

    Private Sub CheckDestinationValidity(destID1 As Integer, destID2 As String, destAddress As String)
        Select Case destID1
            Case 0
                If destAddress.IsNullOrEmpty() Then Throw New Exception($"Valid address expected for destination_id1 = {destID1}.")
            Case 1
                If destID2.IsNullOrEmpty() Then Throw New Exception($"Valid destination_id2 expected for destination_id1 = {destID1}.")
            Case Else
                Throw New NotImplementedException($"Wrong input for destination_id1: {destID1}, expected range: 0-1")
        End Select
    End Sub

    Private Sub CheckInvoiceType(invoice_type1 As InvoiceType, invoice_type2 As String)
        If invoice_type1 = InvoiceType.Other AndAlso invoice_type2.IsNullOrEmpty Then Throw New Exception("Invoice_type2 is mandatory when invoice_type1 = 3 (Other)")
    End Sub

    Private Sub CheckRecallReason(recall_resaon1 As RecallReasonType, recall_reason2 As String)
        If recall_resaon1 = RecallReasonType.Other AndAlso recall_reason2.IsNullOrEmpty() Then Throw New Exception("recall_reason2 is mandatory when recall_reason1 = 3 (Other)")
    End Sub

    ''' <summary>
    ''' If null or emty returns null as string, else returns the string in double quotes "string" 
    ''' </summary>
    ''' <param name="str"></param>
    ''' <returns></returns>
    Private Function Format(str As String) As String
        If str.IsNullOrEmpty() OrElse str Is Nothing Then Return "null"
        Return $"""{str}"""
    End Function
#End Region
#Region "Enums"
    Public Enum AggregationType
        Unit_Packets_Only = 1
        Aggregated_Only = 2
        Both = 3
    End Enum

    Public Enum DestinationType
        Non_EU_dest = 1
        EU_Dest_Other_Than_VM_Fixed_Quantity = 2
        EU_VMs = 3
        Eu_Dest_Other_Than_VM_VV_Delivery = 4
    End Enum

    Public Enum DeactivationType
        Product_destroyed = 1
        Product_stolen = 2
        UI_destroyed = 3
        UI_stolen = 4
        UI_unused = 5
        Other = 6
    End Enum

    Public Enum InvoiceType
        Original = 1
        Correction = 2
        Other = 3
    End Enum

    Public Enum RecallReasonType
        Reported_Event_did_Not_materialise
        Message_contained_erroneous_information
        Other
    End Enum

    Public Enum TransportMode
        Other = 0
        Sea_Transport = 1
        Rail_transport = 2
        Road_transport = 3
        Air_transport = 4
        Postal_consignment = 5
        Fixed_transport_installations = 6
        Inland_waterway_transport = 7
    End Enum
#End Region
End Module
