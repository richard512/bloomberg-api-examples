Option Explicit
'
' public enumerator for request type
Public Enum ENUM_REQUEST_TYPE
    REFERENCE_DATA = 1
    HISTORICAL_DATA = 2
    BULK_REFERENCE_DATA = 3
End Enum
'
' constants
Private Const CONST_SERVICE_TYPE As String = "//blp/refdata"
Private Const CONST_REQUEST_TYPE_REFERENCE As String = "ReferenceDataRequest"
Private Const CONST_REQUEST_TYPE_BULK_REFERENCE As String = "ReferenceDataRequest"
Private Const CONST_REQUEST_TYPE_HISTORICAL As String = "HistoricalDataRequest"
'
' private data structures
Private bInputSecurityArray() As String
Private bInputFieldArray() As String
Private bOutputArray As Variant
Private bOverrideFieldArray() As String
Private bOverrideValueArray() As String
'
' BCOM objects
Private bSession As blpapicomLib2.session
Private bService As blpapicomLib2.Service
Private bRequest As blpapicomLib2.request
Private bSecurityArray As blpapicomLib2.element
Private bFieldArray As blpapicomLib2.element
Private bEvent As blpapicomLib2.Event
Private bIterator As blpapicomLib2.MessageIterator
Private bIteratorData As blpapicomLib2.Message
Private bSecurities As blpapicomLib2.element
Private bSecurity As blpapicomLib2.element
Private bSecurityName As blpapicomLib2.element
Private bSecurityField As blpapicomLib2.element
Private bFieldValue As blpapicomLib2.element
Private bSequenceNumber As blpapicomLib2.element
Private bFields As blpapicomLib2.element
Private bField As blpapicomLib2.element
Private bDataPoint As blpapicomLib2.element
Private bOverrides As blpapicomLib2.element
Private bOverrideArray() As blpapicomLib2.element
'
' class non-object data members
Private bStartDate As String
Private bEndDate As String
Private bRequestType As ENUM_REQUEST_TYPE
Private nSecurities As Long
Private nSecurity As Long
Private bCalendarCodeOverride As String
Private bCurrencyCode As String
Private bNonTradingDayFillOption As String
Private bNonTradingDayFillMethod As String
Private bPeriodicityAdjustment As String
Private bPeriodicitySelection As String
Private bMaxDataPoints As Integer
Private bPricingOption As String
'
Public Function referenceData(ByRef securities As Variant, _
    ByRef fields As Variant, _
    Optional ByRef overrideFields As Variant, _
    Optional ByRef overrideValues As Variant) As Variant
    '
    ' mandatory user input parameters
    bRequestType = REFERENCE_DATA
    bInputSecurityArray = securities
    bInputFieldArray = fields
    '
    ' field names and values for overrides
    If Not (VBA.IsMissing(overrideFields)) Then bOverrideFieldArray = overrideFields
    If Not (VBA.IsMissing(overrideValues)) Then bOverrideValueArray = overrideValues
    '
    processDataRequest
    referenceData = bOutputArray
End Function
'
Public Function bulkReferenceData(ByRef securities As Variant, _
    ByRef fields As Variant, _
    Optional ByRef overrideFields As Variant, _
    Optional ByRef overrideValues As Variant) As Variant
    '
    ' mandatory user input parameters
    bRequestType = BULK_REFERENCE_DATA
    bInputSecurityArray = securities
    bInputFieldArray = fields
    '
    ' field names and values for overrides
    If Not (VBA.IsMissing(overrideFields)) Then bOverrideFieldArray = overrideFields
    If Not (VBA.IsMissing(overrideValues)) Then bOverrideValueArray = overrideValues
    '
    processDataRequest
    bulkReferenceData = bOutputArray
End Function
'
Public Function historicalData(ByRef securities As Variant, _
    ByRef fields As Variant, _
    ByVal startDate As Date, _
    ByVal endDate As Date, _
    Optional ByVal calendarCodeOverride As String, _
    Optional ByVal currencyCode As String, _
    Optional ByVal nonTradingDayFillOption As String, _
    Optional ByVal nonTradingDayFillMethod As String, _
    Optional ByVal periodicityAdjustment As String, _
    Optional ByVal periodicitySelection As String, _
    Optional ByVal maxDataPoints As Integer, _
    Optional ByVal pricingOption As String, _
    Optional ByRef overrideFields As Variant, _
    Optional ByRef overrideValues As Variant) As Variant
    '
    ' mandatory user input parameters
    bRequestType = HISTORICAL_DATA
    bInputSecurityArray = securities
    bInputFieldArray = fields
    bStartDate = startDate
    bEndDate = endDate
    '
    ' checks and conversions for user-defined dates
    If ((startDate = CDate(0)) Or (endDate = CDate(0))) Then _
        Err.Raise vbObjectError, "Bloomberg API", "Date parameters missing for historical data query"
    '
    If (startDate > endDate) Then _
        Err.Raise vbObjectError, "Bloomberg API", "Incorrect date parameters for historical data query"
    '
    bStartDate = convertDateToBloombergString(startDate)
    bEndDate = convertDateToBloombergString(endDate)
    '
    ' optional user input parameters
    bCalendarCodeOverride = calendarCodeOverride
    bCurrencyCode = currencyCode
    bNonTradingDayFillOption = nonTradingDayFillOption
    bNonTradingDayFillMethod = nonTradingDayFillMethod
    bPeriodicityAdjustment = periodicityAdjustment
    bPeriodicitySelection = periodicitySelection
    bMaxDataPoints = maxDataPoints
    bPricingOption = pricingOption
    '
    ' field names and values for overrides
    If Not (VBA.IsMissing(overrideFields)) Then bOverrideFieldArray = overrideFields
    If Not (VBA.IsMissing(overrideValues)) Then bOverrideValueArray = overrideValues
    '
    processDataRequest
    historicalData = bOutputArray
End Function
'
Private Function processDataRequest()
    '
    openSession
    sendRequest
    catchServerEvent
    releaseObjects
End Function
'
Private Function openSession()
    '
    Set bSession = New blpapicomLib2.session
    bSession.Start
    bSession.OpenService CONST_SERVICE_TYPE
    Set bService = bSession.GetService(CONST_SERVICE_TYPE)
End Function
'
Private Function sendRequest()
    '
    Select Case bRequestType
        Case ENUM_REQUEST_TYPE.HISTORICAL_DATA
            ReDim bOutputArray(0 To UBound(bInputSecurityArray, 1), 0 To 0)
            Set bRequest = bService.CreateRequest(CONST_REQUEST_TYPE_HISTORICAL)
            '
            ' set mandatory user input parameter
            bRequest.Set "startDate", bStartDate
            bRequest.Set "endDate", bEndDate
            '
            ' set optional user input parameter
            If (bNonTradingDayFillOption <> "") Then bRequest.Set "nonTradingDayFillOption", bNonTradingDayFillOption
            If (bNonTradingDayFillMethod <> "") Then bRequest.Set "nonTradingDayFillMethod", bNonTradingDayFillMethod
            If (bPeriodicityAdjustment <> "") Then bRequest.Set "periodicityAdjustment", bPeriodicityAdjustment
            If (bPeriodicitySelection <> "") Then bRequest.Set "periodicitySelection", bPeriodicitySelection
            If (bCalendarCodeOverride <> "") Then bRequest.Set "calendarCodeOverride", bCalendarCodeOverride
            If (bCurrencyCode <> "") Then bRequest.Set "currency", bCurrencyCode
            If (bMaxDataPoints <> 0) Then bRequest.Set "maxDataPoints", bMaxDataPoints
            If (bPricingOption <> "") Then bRequest.Set "pricingOption ", bPricingOption
            '
        Case ENUM_REQUEST_TYPE.REFERENCE_DATA
            Dim nSecurities As Long: nSecurities = UBound(bInputSecurityArray)
            Dim nFields As Long: nFields = UBound(bInputFieldArray)
            ReDim bOutputArray(0 To nSecurities, 0 To nFields)
            Set bRequest = bService.CreateRequest(CONST_REQUEST_TYPE_REFERENCE)
            '
        Case ENUM_REQUEST_TYPE.BULK_REFERENCE_DATA
            ReDim bOutputArray(0 To UBound(bInputSecurityArray, 1), 0 To 0)
            Set bRequest = bService.CreateRequest(CONST_REQUEST_TYPE_BULK_REFERENCE)
            '
    End Select
    '
    Set bSecurityArray = bRequest.GetElement("securities")
    Set bFieldArray = bRequest.GetElement("fields")
    appendRequestItems
    setOverrides
    bSession.sendRequest bRequest
End Function
'
Private Function setOverrides()
    '
    On Error GoTo errorHandler
    '
    If (UBound(bOverrideFieldArray) <> UBound(bOverrideValueArray)) Then Exit Function
    Set bOverrides = bRequest.GetElement("overrides")
    '
    ReDim bOverrideArray(LBound(bOverrideFieldArray) To UBound(bOverrideFieldArray))
    Dim i As Integer
    For i = 0 To UBound(bOverrideFieldArray)
        '
        If ((Len(bOverrideFieldArray(i)) > 0) And (Len(bOverrideValueArray(i)) > 0)) Then
            '
            Set bOverrideArray(i) = bOverrides.AppendElment()
            bOverrideArray(i).SetElement "fieldId", bOverrideFieldArray(i)
            bOverrideArray(i).SetElement "value", bOverrideValueArray(i)
        End If
    Next i
    Exit Function
    '
errorHandler:
    Exit Function
End Function
'
Private Function appendRequestItems()
    '
    Dim nSecurities As Long: nSecurities = UBound(bInputSecurityArray)
    Dim nFields As Long: nFields = UBound(bInputFieldArray)
    Dim i As Long
    Dim nItems As Integer: nItems = getMax(nSecurities, nFields)
    For i = 0 To nItems
        If (i <= nSecurities) Then bSecurityArray.AppendValue CStr(bInputSecurityArray(i))
        If (i <= nFields) Then bFieldArray.AppendValue CStr(bInputFieldArray(i))
    Next i
End Function
'
Private Function catchServerEvent()
    '
    Dim bExit As Boolean
    Do While (bExit = False)
        Set bEvent = bSession.NextEvent
        If (bEvent.EventType = PARTIAL_RESPONSE Or bEvent.EventType = RESPONSE) Then
            '
            Select Case bRequestType
                Case ENUM_REQUEST_TYPE.REFERENCE_DATA: getServerData_reference
                Case ENUM_REQUEST_TYPE.HISTORICAL_DATA: getServerData_historical
                Case ENUM_REQUEST_TYPE.BULK_REFERENCE_DATA: getServerData_bulkReference
            End Select
            '
            If (bEvent.EventType = RESPONSE) Then bExit = True
        End If
    Loop
End Function
'
Private Function getServerData_reference()
    '
    Set bIterator = bEvent.CreateMessageIterator
    Do While (bIterator.Next)
        Set bIteratorData = bIterator.Message
        Set bSecurities = bIteratorData.GetElement("securityData")
        Dim offsetNumber As Long, i As Long, j As Long
        nSecurities = bSecurities.Count
        '
        For i = 0 To (nSecurities - 1)
            Set bSecurity = bSecurities.GetValue(i)
            Set bSecurityName = bSecurity.GetElement("security")
            Set bSecurityField = bSecurity.GetElement("fieldData")
            Set bSequenceNumber = bSecurity.GetElement("sequenceNumber")
            offsetNumber = CInt(bSequenceNumber.Value)
            '
            For j = 0 To UBound(bInputFieldArray)
                If (bSecurityField.HasElement(bInputFieldArray(j))) Then
                    Set bFieldValue = bSecurityField.GetElement(bInputFieldArray(j))
                    bOutputArray(offsetNumber, j) = bFieldValue.Value
                End If
            Next j
        Next i
    Loop
End Function
'
Private Function getServerData_bulkReference()
    '
    Set bIterator = bEvent.CreateMessageIterator
    nSecurity = nSecurity + 1
    '
    Do While (bIterator.Next)
        Set bIteratorData = bIterator.Message
        Set bSecurities = bIteratorData.GetElement("securityData")
        Dim offsetNumber As Long, i As Long, j As Long
        Dim nSecurities As Long: nSecurities = bSecurities.Count
        '
        Set bSecurity = bSecurities.GetValue(0)
        Set bSecurityField = bSecurity.GetElement("fieldData")
        '
        If (bSecurityField.HasElement(bInputFieldArray(0))) Then
            Set bFieldValue = bSecurityField.GetElement(bInputFieldArray(0))
            '
            If ((bFieldValue.numValues - 1) > UBound(bOutputArray, 2)) Then _
                ReDim Preserve bOutputArray(0 To UBound(bOutputArray, 1), 0 To bFieldValue.numValues - 1)
            '
            For i = 0 To bFieldValue.numValues - 1
                Set bDataPoint = bFieldValue.GetValue(i)
                bOutputArray(nSecurity - 1, i) = bDataPoint.GetElement(0).Value
            Next i
        End If
    Loop
End Function
'
Private Function getServerData_historical()
    '
    Set bIterator = bEvent.CreateMessageIterator
    Do While (bIterator.Next)
        Set bIteratorData = bIterator.Message
        Set bSecurities = bIteratorData.GetElement("securityData")
        Dim nSecurities As Long: nSecurities = bSecurityArray.Count
        Set bSecurityField = bSecurities.GetElement("fieldData")
        Dim nItems As Long, offsetNumber As Long, nFields As Long, i As Long, j As Long
        nItems = bSecurityField.numValues
        If (nItems = 0) Then Exit Function
        If ((nItems > UBound(bOutputArray, 2))) Then _
            ReDim Preserve bOutputArray(0 To nSecurities - 1, 0 To nItems - 1)
        '
        Set bSequenceNumber = bSecurities.GetElement("sequenceNumber")
        offsetNumber = CInt(bSequenceNumber.Value)
        '
        If (bSecurityField.Count > 0) Then
            For i = 0 To (nItems - 1)
                '
                If (bSecurityField.Count > i) Then
                    Set bFields = bSecurityField.GetValue(i)
                    If (bFields.HasElement(bFieldArray(0))) Then
                        '
                        Dim d As Variant: ReDim d(0 To bFields.NumElements - 1)
                        For j = 0 To bFields.NumElements - 1
                            d(j) = bFields.GetElement(j).GetValue(0)
                        Next j
                        '
                        bOutputArray(offsetNumber, i) = d
                    End If
                End If
            Next i
        End If
    Loop
End Function
'
Private Function releaseObjects()
    '
    nSecurity = 0
    Set bDataPoint = Nothing
    Set bFieldValue = Nothing
    Set bSequenceNumber = Nothing
    Set bSecurityField = Nothing
    Set bSecurityName = Nothing
    Set bSecurity = Nothing
    Set bOverrides = Nothing
    Set bSecurities = Nothing
    Set bIteratorData = Nothing
    Set bIterator = Nothing
    Set bEvent = Nothing
    Set bFieldArray = Nothing
    Set bSecurityArray = Nothing
    Set bRequest = Nothing
    Set bService = Nothing
    bSession.Stop
    Set bSession = Nothing
End Function
'
Private Function convertDateToBloombergString(ByVal d As Date) As String
    '
    Dim dayString As String: dayString = VBA.CStr(VBA.Day(d)): If (VBA.Day(d) < 10) Then dayString = "0" + dayString
    Dim MonthString As String: MonthString = VBA.CStr(VBA.Month(d)): If (VBA.Month(d) < 10) Then MonthString = "0" + MonthString
    Dim yearString As String: yearString = VBA.Year(d)
    convertDateToBloombergString = yearString + MonthString + dayString
End Function
'
Private Function getMax(ByVal a As Long, ByVal b As Long) As Long
    '
    getMax = a: If (b > a) Then getMax = b
End Function