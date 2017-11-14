'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'End
'Attribute VB_Name = "BloomSync"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
Option Explicit

Private session As blpapicomLib.session
Dim refdataservice As blpapicomLib.Service

Dim origSecurities As Variant
Dim orgiFields As Variant

Dim return_format_mode As Integer

Public Enum output_format
    of_vec_with_header = 0
    of_vec_without_header = 1
    of_matrix_with_header = 2
    of_matrix_without_header = 3
End Enum


'vec_securities_id
    '/[Topic Prefix]/SYMBOLOGY@Pricing Source][Exchange]
        'Topic Prefix:ticker (Default), cusip, buid, bbgid, wpk, isin, sedol1, sedol2, sicovam, common, svm, cins, cats


Private Sub Class_Initialize()

Set session = New blpapicomLib.session
session.QueueEvents = True
session.Start

End Sub

Private Sub Class_Terminate()

    Set session = Nothing

End Sub


Public Function bdp_without_error_handling(ByVal vec_securities_id As Variant, vec_fields As Variant, Optional ByVal return_format = output_format.of_vec_with_header, Optional ByVal vec_override_fields As Variant, Optional ByVal vec_override_values As Variant) As Variant

    Dim debug_test As Variant, debug_test_2 As Variant
    
    Dim i As Long, j As Long, k As Long, m As Long, n As Long, p As Long, q As Long
    
    session.OpenService ("//blp/refdata")
    Set refdataservice = session.GetService("//blp/refdata")
    
    Dim req As blpapicomLib.REQUEST
    Set req = refdataservice.CreateRequest("ReferenceDataRequest")
    
    Dim count_securities As Long
    count_securities = 0
    For i = LBound(vec_securities_id, 1) To UBound(vec_securities_id, 1)
        req.GetElement("securities").AppendValue vec_securities_id(i)
        count_securities = count_securities + 1
    Next i
    
    Dim count_fields As Long
    count_fields = 0
    For i = LBound(vec_fields, 1) To UBound(vec_fields, 1)
        req.GetElement("fields").AppendValue vec_fields(i)
        count_fields = count_fields + 1
    Next i
    
    If IsMissing(vec_override_fields) = False And IsMissing(vec_override_values) = False Then
        If UBound(vec_override_fields, 1) = UBound(vec_override_values, 1) Then
            Dim overrides As blpapicomLib.Element
            Set overrides = req.GetElement("overrides")
            
            Dim override As blpapicomLib.Element
            For i = LBound(vec_override_fields, 1) To UBound(vec_override_fields, 1)
                Set override = overrides.AppendElment()
                override.SetElement "fieldId", vec_override_fields(i)
                override.SetElement "value", vec_override_values(i)
            Next i
        End If
    End If
    
    session.sendRequest req
    
    k = 1 'compteur de ticker
    Dim tmp_vec_one_line() As Variant, tmp_vec_header() As Variant
    Dim vec_output() As Variant
        ReDim Preserve vec_output(0) 'pour le header
    Dim already_store_headers As Boolean
        already_store_headers = False
    
    Dim eventObj As blpapicomLib.Event
    Dim msg As blpapicomLib.Message
    Dim it As blpapicomLib.MessageIterator
    
    Dim numSecurities As Long
    Dim security As blpapicomLib.Element
    
    Dim fields As blpapicomLib.Element
    Dim field As blpapicomLib.Element
    Dim numFields As Long
    
    Dim bulkElement As blpapicomLib.Element
    Dim bulkSubElement As blpapicomLib.Element
    Dim numBulkElements As Long
    
    Dim elem As blpapicomLib.Element
    
    Dim vec_field_bulk()
    
    Do
        Set eventObj = session.NextEvent()
        
        If eventObj.EventType = PARTIAL_RESPONSE Or eventObj.EventType = Response Then
            
            Set it = eventObj.CreateMessageIterator()
            
            Do While it.Next()
                
                Set msg = it.Message
                
                numSecurities = msg.GetElement("securityData").NumValues
                
                For i = 0 To numSecurities - 1 'boucle ticker
                    
                    Set security = msg.GetElement("securityData").GetValue(i)
                    
                    Set fields = security.GetElement("fieldData")
                    numFields = fields.NumElements
                    
                    If fields.NumElements = 0 Then 'erreur, security_id invalide
                        
                        ReDim tmp_vec_one_line(count_fields)
                        
                        tmp_vec_one_line(0) = vec_securities_id(i)
                        
                        For j = 1 To count_fields
                            tmp_vec_one_line(j) = "#N/A Invalid security"
                        Next j
                        
                        ReDim Preserve vec_output(k)
                        vec_output(k) = tmp_vec_one_line
                        k = k + 1
                        
                    Else
                    
                        ReDim Preserve tmp_vec_one_line(numFields)
                        tmp_vec_one_line(0) = security.GetElement("security").Value
                        
                        ReDim Preserve tmp_vec_header(numFields)
                        tmp_vec_header(0) = "security_id"
                        
                        For j = 0 To numFields - 1 'boucle field
                            
                            Set field = fields.GetElement(j)
                            
                            If field.DataType = 15 Then
                                
                                For m = 0 To field.NumValues - 1
                                    Set bulkElement = field.GetValue(m)
                                    numBulkElements = bulkElement.NumElements
                                    
                                    For n = 0 To bulkElement.NumElements - 1
                                        Set elem = bulkElement.GetElement(n)
                                        
                                        ReDim Preserve vec_field_bulk(m)
                                        vec_field_bulk(m) = elem.Value
                                        
                                    Next n
                                    
                                    tmp_vec_one_line(j + 1) = vec_field_bulk
                                    
                                    If already_store_headers = False Then
                                        tmp_vec_header(j + 1) = fields.GetElement(j).Name
                                    End If
                                    
                                Next m
                                
                            Else
                            
                                If already_store_headers = False Then
                                    tmp_vec_header(j + 1) = fields.GetElement(j).Name
                                End If
                                
                                tmp_vec_one_line(j + 1) = fields.GetElement(j).Value
                            End If
                        Next j
                        
                        already_store_headers = True
                        vec_output(0) = tmp_vec_header
                        
                        ReDim Preserve vec_output(k)
                        vec_output(k) = tmp_vec_one_line
                        k = k + 1
                    
                    End If
                Next i
            Loop
            
            If eventObj.EventType = Response Then Exit Do
            
        End If
    Loop
    
    
    'transformation selon l'output choisi
    
    Dim matrix_output() As Variant
    Dim tmp_vec As Variant
    
    If return_format = output_format.of_vec_with_header Then
        
        bdp = vec_output
        
    ElseIf return_format = output_format.of_vec_without_header Then
        tmp_vec = vec_output
        
        m = 0
        For i = 1 To UBound(tmp_vec, 1)
            n = 0
            For j = 1 To UBound(tmp_vec(i), 1)
                ReDim Preserve tmp_vec_one_line(n)
                tmp_vec_one_line(n) = tmp_vec(i)(j)
                n = n + 1
            Next j
            
            ReDim Preserve vec_output(m)
            vec_output(m) = tmp_vec_one_line
            m = m + 1
        Next i
        
        bdp = vec_output
    
    ElseIf return_format = output_format.of_matrix_with_header Then
        
        ReDim matrix_output(UBound(vec_output, 1), UBound(vec_output(0), 1))
        
        For i = 0 To UBound(vec_output, 1)
            For j = 0 To UBound(vec_output(i), 1)
                matrix_output(i, j) = vec_output(i)(j)
            Next j
        Next i
        
        bdp = matrix_output
    
    ElseIf return_format = output_format.of_matrix_without_header Then
        
        ReDim matrix_output(UBound(vec_output, 1) - 1, UBound(vec_output(0), 1) - 1)
        tmp_vec = vec_output
        
        m = 0
        For i = 1 To UBound(tmp_vec, 1)
            n = 0
            For j = 1 To UBound(tmp_vec(i), 1)
                matrix_output(m, n) = tmp_vec(i)(j)
                n = n + 1
            Next j
            m = m + 1
        Next i
        
        bdp = matrix_output
        
    End If

End Function


'with error handling
Public Function bdp(ByVal vec_securities_id As Variant, vec_fields As Variant, Optional ByVal return_format = output_format.of_vec_with_header, Optional ByVal vec_override_fields As Variant, Optional ByVal vec_override_values As Variant) As Variant

    origSecurities = vec_securities_id
    orgiFields = vec_fields
    
    return_format_mode = return_format
    
    Dim debug_test As Variant, debug_test_2 As Variant
    
    Dim i As Long, j As Long, k As Long, m As Long, n As Long, p As Long, q As Long
    
    session.OpenService ("//blp/refdata")
    Set refdataservice = session.GetService("//blp/refdata")
    
    Dim req As blpapicomLib.REQUEST
    Set req = refdataservice.CreateRequest("ReferenceDataRequest")
    
    Dim count_securities As Long
    count_securities = 0
    For i = LBound(vec_securities_id, 1) To UBound(vec_securities_id, 1)
        req.GetElement("securities").AppendValue vec_securities_id(i)
        count_securities = count_securities + 1
    Next i
    
    Dim count_fields As Long
    count_fields = 0
    For i = LBound(vec_fields, 1) To UBound(vec_fields, 1)
        req.GetElement("fields").AppendValue vec_fields(i)
        count_fields = count_fields + 1
    Next i
    
    If IsMissing(vec_override_fields) = False And IsMissing(vec_override_values) = False Then
        If UBound(vec_override_fields, 1) = UBound(vec_override_values, 1) Then
            Dim overrides As blpapicomLib.Element
            Set overrides = req.GetElement("overrides")
            
            Dim override As blpapicomLib.Element
            For i = LBound(vec_override_fields, 1) To UBound(vec_override_fields, 1)
                Set override = overrides.AppendElment()
                override.SetElement "fieldId", vec_override_fields(i)
                override.SetElement "value", vec_override_values(i)
            Next i
        End If
    End If
    
    
    Dim queue As blpapicomLib.EventQueue
    Set queue = session.CreateEventQueue
    
    session.sendRequest req, , queue
    
    Dim done As Boolean
    done = False
    
    Dim eventObj As blpapicomLib.Event
    
    Dim partial_output() As Variant
    Dim vec_output() As Variant
    ReDim vec_output(0)
    Dim last_ubound As Double
    
    Dim start_partial As Integer
    k = 0
    m = 0
    Do While done = False
        Set eventObj = queue.NextEvent()
        
        If eventObj.EventType = PARTIAL_RESPONSE Or eventObj.EventType = Response Then
        
            ReDim Preserve partial_output(k)
            partial_output(k) = ProcessResponseEventBDP(eventObj)
            
            last_ubound = UBound(vec_output, 1)
            If last_ubound = 0 Then
                last_ubound = last_ubound - 1
            End If
            
            start_partial = 0
            If k = 0 Then
                start_partial = 0
            Else
                start_partial = 1 'eviter header
            End If
            
            ReDim Preserve vec_output(last_ubound + 1 + UBound(partial_output(k), 1) - start_partial)
            
            For i = start_partial To UBound(partial_output(k), 1)
                vec_output(m) = partial_output(k)(i)
                m = m + 1
            Next i
            
            k = k + 1
        End If
        
        If eventObj.EventType = Response Then
            done = True
        End If
    Loop
    
    
    'retourne l'output dans le meme sens que security_id receptionne
    Dim tmp_vec_output As Variant
    tmp_vec_output = vec_output
    
    
    For i = 0 To UBound(vec_securities_id, 1)
        For j = 1 To UBound(tmp_vec_output, 1)
            If vec_securities_id(i) = tmp_vec_output(j)(0) Then
                vec_output(i + 1) = tmp_vec_output(j)
                Exit For
            End If
        Next j
    Next i
    
    
    'transformation selon l'output choisi
    
    
    Dim matrix_output() As Variant
    Dim tmp_vec As Variant
    Dim tmp_vec_one_line() As Variant
    
    If return_format_mode = output_format.of_vec_with_header Then
        
        bdp = vec_output
        
    ElseIf return_format_mode = output_format.of_vec_without_header Then
        tmp_vec = vec_output
        
        m = 0
        For i = 1 To UBound(tmp_vec, 1)
            n = 0
            For j = 1 To UBound(tmp_vec(i), 1)
                ReDim Preserve tmp_vec_one_line(n)
                tmp_vec_one_line(n) = tmp_vec(i)(j)
                n = n + 1
            Next j
            
            ReDim Preserve vec_output(m)
            vec_output(m) = tmp_vec_one_line
            m = m + 1
        Next i
        
        bdp = vec_output
    
    ElseIf return_format_mode = output_format.of_matrix_with_header Then
        
        ReDim matrix_output(UBound(vec_output, 1), UBound(vec_output(0), 1))
        
        For i = 0 To UBound(vec_output, 1)
            For j = 0 To UBound(vec_output(i), 1)
                matrix_output(i, j) = vec_output(i)(j)
            Next j
        Next i
        
        bdp = matrix_output
    
    ElseIf return_format_mode = output_format.of_matrix_without_header Then
        
        ReDim matrix_output(UBound(vec_output, 1) - 1, UBound(vec_output(0), 1) - 1)
        tmp_vec = vec_output
        
        m = 0
        For i = 1 To UBound(tmp_vec, 1)
            n = 0
            For j = 1 To UBound(tmp_vec(i), 1)
                matrix_output(m, n) = tmp_vec(i)(j)
                n = n + 1
            Next j
            m = m + 1
        Next i
        
        bdp = matrix_output
        
    End If



End Function



'Public Function ProcessResponseEventBDP(eventObj As blpapicomLib.Event, Optional ByVal return_format = output_format.of_vec_with_header) As Variant
Public Function ProcessResponseEventBDP(eventObj As blpapicomLib.Event) As Variant

    Dim debug_test As Variant
    
    Dim i As Long, j As Long, k As Long, m As Long, n As Long, p As Long, q As Long
    
    Dim it As blpapicomLib.MessageIterator
    Set it = eventObj.CreateMessageIterator()
    
    Dim msg As blpapicomLib.Message
    Dim error As blpapicomLib.Element
    
    Dim securityDataArray As blpapicomLib.Element
    Dim numSecurities As Integer
    Dim securityData As blpapicomLib.Element
    Dim security As String
    Dim sequenceNumber As Integer
    Dim fieldDataArray As blpapicomLib.Element
    Dim field As blpapicomLib.Element
    
    Dim vec_output() As Variant
    Dim vec_output_one_line() As Variant
    
    Dim hasErrorSecurity As Boolean
    
    Dim vec_bulk_array() As Variant, vec_bulk_one_line() As Variant
    Dim bulkElement As blpapicomLib.Element
    Dim numBulkElement As Long
    
    Dim elem As blpapicomLib.Element
    
    
    Do While it.Next()
        Set msg = it.Message
        
        'test si appel refus? (limit etc.)
        If msg.AsElement.HasElement("responseError") Then
            Set error = msg.GetElement("responseError")
            Call MsgBox("responseError: " & error.GetElement("message"))
            ProcessResponseEventBDP = Empty
            Exit Function
        End If
        
        Set securityDataArray = msg.GetElement("securityData")
        numSecurities = securityDataArray.NumValues
        
        'mise en place du header
        n = 0
        ReDim Preserve vec_output(0)
        ReDim vec_output_one_line(UBound(orgiFields, 1) + 1)
        vec_output_one_line(0) = "security_id"
        For k = LBound(orgiFields, 1) To UBound(orgiFields, 1)
            vec_output_one_line(n + 1) = orgiFields(k)
            n = n + 1
        Next k
        
        vec_output(0) = vec_output_one_line
        
        For i = 0 To securityDataArray.NumValues - 1 'boucle security_id
            
            hasErrorSecurity = False
            
            Set securityData = securityDataArray.GetValue(i)
            security = securityData.GetElement("security").Value
            
            sequenceNumber = securityData.GetElement("sequenceNumber").Value
            
            
            'check for securityError
            ReDim Preserve vec_output_one_line(0)
            If securityData.HasElement("securityError") Then
                debug_test = securityData.GetElement("securityError").GetElement("category").Value
                debug_test = securityData.GetElement("securityError").GetElement("message").Value
                
                'vec_output_one_line(0) = "#N/A security"
                vec_output_one_line(0) = securityData.GetElement("security").Value
                
                hasErrorSecurity = True
            Else
                vec_output_one_line(0) = securityData.GetElement("security").Value
            End If
            
            
            'Process fieldData
            If securityData.HasElement("fieldData") Then
                Set fieldDataArray = securityData.GetElement("fieldData")
                
                For k = LBound(orgiFields, 1) To UBound(orgiFields, 1) 'boucle sur field
                    
                    ReDim Preserve vec_output_one_line(k + 1)
                    
                    If fieldDataArray.HasElement(orgiFields(k)) Then
                        
                        Set field = fieldDataArray.GetElement(orgiFields(k))
                        
                        If field.DataType = 15 Then 'bulk
                            
                            For m = 0 To field.NumValues - 1
                                
                                Set bulkElement = field.GetValue(m)
                                
                                For n = 0 To bulkElement.NumElements - 1
                                    Set elem = bulkElement.GetElement(n)
                                    ReDim Preserve vec_bulk_one_line(n)
                                    vec_bulk_one_line(n) = elem.Value
                                    
                                Next n
                                
                                ReDim Preserve vec_bulk_array(m)
                                vec_bulk_array(m) = vec_bulk_one_line
                                
                            Next m
                            
                            vec_output_one_line(k + 1) = vec_bulk_array
                            
                        Else 'donnee normal
                            vec_output_one_line(k + 1) = fieldDataArray.GetElement(orgiFields(k)).Value
                        End If
                        
                        
                    Else
                        'debug_test = "N/A"
                        If hasErrorSecurity = True Then
                            vec_output_one_line(k + 1) = "#N/A security"
                        Else
                            vec_output_one_line(k + 1) = "#N/A field"
                        End If
                    End If
                    
                Next k
            Else
                ReDim vec_output_one_line(UBound(orgiFields, 1) + 1)
                For k = LBound(orgiFields, 1) To UBound(orgiFields, 1)
                    vec_output_one_line(k + 1) = "#N/A field"
                Next k
            End If
            
            ReDim Preserve vec_output(i + 1)
            vec_output(i + 1) = vec_output_one_line
            
        Next i
        
    Loop
    
    
    ProcessResponseEventBDP = vec_output

End Function





Public Function bdh_without_error_handling(ByVal vec_securities_id As Variant, vec_fields As Variant, ByVal date_from As Variant, Optional ByVal date_to As Variant, Optional ByVal periodicity As Variant = "DAILY", Optional ByVal return_format = output_format.of_vec_without_header) As Variant

    Dim debug_test As Variant
    
    Dim i As Long, j As Long, k As Long, m As Long, n As Long, p As Long, q As Long, r As Long, s As Long
    
    session.OpenService ("//blp/refdata")
    Set refdataservice = session.GetService("//blp/refdata")
    
    'traitement sur les dates
    Dim date_from_bbg_str As String
    Dim date_to_bbg_str As String
    
    If IsDate(date_from) Then
        date_from_bbg_str = CStr(Year(date_from)) & CStr(Month(date_from)) & CStr(Day(date_from))
    ElseIf IsNumeric(date_from) Then
        date_from_bbg_str = CStr(date_from)
    End If
    
    
    
    If IsMissing(date_to) Then
        date_to = Date
    End If
    
    
    If IsDate(date_to) Then
        date_to_bbg_str = CStr(Year(date_to)) & CStr(Month(date_to)) & CStr(Day(date_to))
    ElseIf IsNumeric(date_from) Then
        date_to_bbg_str = CStr(date_to)
    End If
    
    Dim req As blpapicomLib.REQUEST
    Set req = refdataservice.CreateRequest("HistoricalDataRequest")
    
    For i = LBound(vec_securities_id, 1) To UBound(vec_securities_id, 1)
        req.GetElement("securities").AppendValue (vec_securities_id(i))
    Next i
    
    For i = LBound(vec_fields, 1) To UBound(vec_fields, 1)
        req.GetElement("fields").AppendValue (vec_fields(i))
    Next i
    
    
    req.Set "periodicitySelection", periodicity
    req.Set "startDate", date_from_bbg_str
    req.Set "endDate", date_to_bbg_str
    
    session.sendRequest req
    
    
    Dim eventObj As blpapicomLib.Event
    Dim it As blpapicomLib.MessageIterator
    
    Dim msg As blpapicomLib.Message
    Dim numSecurities As Long
    
    'Dim security As blpapicomLib.Element
    Dim securityData As blpapicomLib.Element
    Dim securityName As blpapicomLib.Element
    
    Dim numDates As Integer
    Dim numFields As Integer
    Dim fieldData As blpapicomLib.Element
    Dim fields As blpapicomLib.Element
    Dim field As blpapicomLib.Element
    
    
    Dim vec_level_ticker() As Variant
    Dim vec_level_date() As Variant
    Dim vec_level_data() As Variant
    
    Dim vec_level_ticker_with_header() As Variant
    Dim vec_level_date_with_header() As Variant
    Dim vec_level_data_with_header() As Variant
    
    Dim max_ubound_data As Long, max_ubound_date As Long, max_ubound_ticker As Long
        max_ubound_data = 0
        max_ubound_date = 0
        max_ubound_ticker = 0
    
    k = 0
    Do
        Set eventObj = session.NextEvent()
        
        If eventObj.EventType = PARTIAL_RESPONSE Or eventObj.EventType = Response Then
            
            Set it = eventObj.CreateMessageIterator()
            
            Do While it.Next() 'boucle ticker
                
                Set msg = it.Message
                
                Set securityData = msg.GetElement("securityData")
                Set securityName = securityData.GetElement("security")
                Set fieldData = securityData.GetElement("fieldData")
                
                'numDates = fieldData.NumValues 'pour debug
                
                If fieldData.NumValues - 1 > max_ubound_date Then
                    max_ubound_date = fieldData.NumValues - 1
                End If
                
                For m = 0 To fieldData.NumValues - 1 'boucle sur date
                    Set fields = fieldData.GetValue(m)
                    
                    'numFields = fields.NumElements ' pour debug
                    
                    If fields.NumElements > max_ubound_data Then
                        max_ubound_data = fields.NumElements
                    End If
                    
                    ReDim Preserve vec_level_data(0)
                    vec_level_data(0) = securityData.GetElement("security").Value
                    
                    For n = 0 To fields.NumElements - 1
                        Set field = fields.GetElement(n)
                        ReDim Preserve vec_level_data(n + 1)
                        ReDim Preserve vec_level_data_with_header(n + 1)
                        
                        vec_level_data(n + 1) = field.Value
                        vec_level_data_with_header(n + 1) = Array(field.Name, field.Value)
                    Next n
                    
                    ReDim Preserve vec_level_date(m)
                    ReDim Preserve vec_level_date_with_header(m)
                    
                    vec_level_date(m) = vec_level_data
                    vec_level_date_with_header(m) = Array(fields.GetElement(0).Value, vec_level_data)
                    
                Next m
                
                ReDim Preserve vec_level_ticker(k)
                ReDim Preserve vec_level_ticker_with_header(k)
                
                vec_level_ticker(k) = vec_level_date
                vec_level_ticker_with_header(k) = Array(securityName.Value, vec_level_date)
                
                k = k + 1
                
            Loop
            
            If eventObj.EventType = Response Then Exit Do
            
        End If
    Loop
    
    
    Dim matrix_output() As Variant
    
    If return_format = output_format.of_vec_with_header Then
        
        bdh_without_error_handling = vec_level_ticker_with_header
        
    ElseIf return_format = output_format.of_vec_without_header Then
        
        bdh_without_error_handling = vec_level_ticker
    
    ElseIf return_format = output_format.of_matrix_with_header Then
        
        ReDim Preserve matrix_output(UBound(vec_level_ticker, 1), max_ubound_date, max_ubound_data)
        
        For p = 0 To UBound(vec_level_ticker, 1)
            For q = 0 To UBound(vec_level_ticker(p), 1)
                For r = 0 To UBound(vec_level_ticker(p)(q), 1)
                    matrix_output(p, q, r) = vec_level_ticker(p)(q)(r)
                Next r
            Next q
        Next p
        
        bdh_without_error_handling = matrix_output
    
    ElseIf return_format = output_format.of_matrix_without_header Then
        
        ReDim Preserve matrix_output(UBound(vec_level_ticker, 1), max_ubound_date, max_ubound_data)
        
        For p = 0 To UBound(vec_level_ticker, 1)
            For q = 0 To UBound(vec_level_ticker(p), 1)
                For r = 0 To UBound(vec_level_ticker(p)(q), 1)
                    matrix_output(p, q, r) = vec_level_ticker(p)(q)(r)
                Next r
            Next q
        Next p
        
        bdh_without_error_handling = matrix_output
        
    End If


End Function




Public Function bdh(ByVal vec_securities_id As Variant, vec_fields As Variant, ByVal date_from As Variant, Optional ByVal date_to As Variant, Optional ByVal periodicity As Variant = "DAILY", Optional ByVal return_format = output_format.of_vec_without_header, Optional ByVal vec_override_fields As Variant, Optional ByVal vec_override_values As Variant) As Variant

    Dim date_tmp As Date
    
    Dim debug_test As Variant
    
    Dim i As Long, j As Long, k As Long, m As Long, n As Long, p As Long, q As Long, r As Long, s As Long
    
    session.OpenService ("//blp/refdata")
    Set refdataservice = session.GetService("//blp/refdata")
    
    'traitement sur les dates
    Dim date_from_bbg_str As String
    Dim date_to_bbg_str As String
    
    Dim date_from_year_txt As String, date_from_month_txt As String, date_from_day_txt As String
    
    If IsDate(date_from) Then
        
        date_from_year_txt = CStr(Year(date_from))
        date_from_month_txt = CStr(Month(date_from))
            If Len(date_from_month_txt) = 1 Then
                date_from_month_txt = "0" & date_from_month_txt
            End If
        date_from_day_txt = CStr(Day(date_from))
            If Len(date_from_day_txt) = 1 Then
                date_from_day_txt = "0" & date_from_day_txt
            End If
        
        date_from_bbg_str = date_from_year_txt & date_from_month_txt & date_from_day_txt
    ElseIf IsNumeric(date_from) Then
        date_from_bbg_str = CStr(date_from)
    End If
    
    
    
    If IsMissing(date_to) Then
        date_to = Date
    End If
    
    
    Dim date_to_year_txt As String, date_to_month_txt As String, date_to_day_txt As String
    
    If IsDate(date_to) Then
        
        date_to_year_txt = CStr(Year(date_to))
        date_to_month_txt = CStr(Month(date_to))
            If Len(date_to_month_txt) = 1 Then
                date_to_month_txt = "0" & date_to_month_txt
            End If
        date_to_day_txt = CStr(Day(date_to))
            If Len(date_to_day_txt) = 1 Then
                date_to_day_txt = "0" & date_to_day_txt
            End If
        
        date_to_bbg_str = date_to_year_txt & date_to_month_txt & date_to_day_txt
    ElseIf IsNumeric(date_from) Then
        date_to_bbg_str = CStr(date_to)
    End If
    
    Dim req As blpapicomLib.REQUEST
    Set req = refdataservice.CreateRequest("HistoricalDataRequest")
    
    For i = LBound(vec_securities_id, 1) To UBound(vec_securities_id, 1)
        req.GetElement("securities").AppendValue (vec_securities_id(i))
    Next i
    
    For i = LBound(vec_fields, 1) To UBound(vec_fields, 1)
        req.GetElement("fields").AppendValue (vec_fields(i))
    Next i
    
    
    req.Set "periodicitySelection", periodicity
    req.Set "startDate", date_from_bbg_str
    req.Set "endDate", date_to_bbg_str
    
    
    If IsMissing(vec_override_fields) = False And IsMissing(vec_override_values) = False Then
        If UBound(vec_override_fields, 1) = UBound(vec_override_values, 1) Then
            Dim overrides As blpapicomLib.Element
            Set overrides = req.GetElement("overrides")
            
            Dim override As blpapicomLib.Element
            For i = LBound(vec_override_fields, 1) To UBound(vec_override_fields, 1)
                Set override = overrides.AppendElment()
                override.SetElement "fieldId", vec_override_fields(i)
                override.SetElement "value", vec_override_values(i)
            Next i
        End If
    End If
    
    
    Dim queue As blpapicomLib.EventQueue
    Set queue = session.CreateEventQueue
    
    origSecurities = vec_securities_id
    orgiFields = vec_fields
    return_format_mode = return_format
    
    
    session.sendRequest req, , queue
    
    Dim done As Boolean
    done = False
    
    Dim eventObj As blpapicomLib.Event
    
    Dim vec_ticker_level() As Variant
    
    k = 0
    Do While done = False
        Set eventObj = queue.NextEvent()
        
        If eventObj.EventType = PARTIAL_RESPONSE Or eventObj.EventType = Response Then
            ReDim Preserve vec_ticker_level(k)
            vec_ticker_level(k) = ProcessResponseEventBDH(eventObj)
            k = k + 1
        End If
        
        If eventObj.EventType = Response Then
            done = True
        End If
        
    Loop
    
    
    'l'ordre d'input n'est plus forcement respecter
    Dim tmp_vec_reorder As Variant
    Dim vec_ticker_level_order() As Variant
    ReDim vec_ticker_level_order(UBound(vec_ticker_level, 1))
    For i = 0 To UBound(vec_securities_id, 1)
        For j = 0 To UBound(vec_ticker_level, 1)
            If vec_ticker_level(j)(0)(0) = vec_securities_id(i) Then
                vec_ticker_level_order(i) = vec_ticker_level(j)
            End If
        Next j
    Next i
    
    vec_ticker_level = vec_ticker_level_order
    
    'formatage de l'output
    Dim max_ubound_ticker As Long, max_ubound_date As Long, max_ubound_data As Long
        max_ubound_ticker = 0
        max_ubound_date = 0
        max_ubound_data = 0
    Dim matrix_output() As Variant
    
    Dim vec_distinct_date() As Variant
    Dim min_date As Date
    Dim min_pos As Long
    Dim count_date As Long
    Dim last_position_date As Long
    
    If return_format = output_format.of_vec_with_header Or return_format = output_format.of_vec_without_header Then
        bdh = vec_ticker_level
    ElseIf return_format = output_format.of_matrix_with_header Or return_format = output_format.of_matrix_without_header Then
    
        ReDim Preserve vec_distinct_date(0)
        date_tmp = Date + 365
        vec_distinct_date(0) = date_tmp
        
        max_ubound_ticker = UBound(vec_ticker_level, 1)
        max_ubound_data = UBound(vec_ticker_level(0)(0), 1)
        
        count_date = 0
        For i = 0 To UBound(vec_ticker_level, 1)
            For j = 0 To UBound(vec_ticker_level(i), 1)
                If IsDate(vec_ticker_level(i)(j)(1)) Then
                    For k = 0 To UBound(vec_distinct_date, 1)
                        If vec_ticker_level(i)(j)(1) = vec_distinct_date(k) Then
                            Exit For
                        Else
                            If k = UBound(vec_distinct_date, 1) Then
                                ReDim Preserve vec_distinct_date(count_date)
                                vec_distinct_date(count_date) = vec_ticker_level(i)(j)(1)
                                count_date = count_date + 1
                            End If
                        End If
                    Next k
                End If
            Next j
        Next i
        
        'sort vec_date
        For i = 0 To UBound(vec_distinct_date, 1)
            
            min_date = vec_distinct_date(i)
            min_pos = i
            
            For j = i + 1 To UBound(vec_distinct_date, 1)
                date_tmp = vec_distinct_date(j)
                If date_tmp < min_date Then
                    min_date = date_tmp
                    min_pos = j
                End If
            Next j
            
            If i <> min_pos Then
                date_tmp = vec_distinct_date(i)
                vec_distinct_date(i) = vec_distinct_date(min_pos)
                vec_distinct_date(min_pos) = date_tmp
            End If
            
        Next i
        
        ReDim Preserve matrix_output(UBound(vec_ticker_level, 1), UBound(vec_distinct_date, 1), UBound(vec_ticker_level(0)(0), 1))
        
        For i = 0 To UBound(vec_ticker_level, 1)
            
            last_position_date = 0
            
            For j = 0 To UBound(vec_distinct_date, 1)
                'regarde si la date existe pour le produit courant
                For m = last_position_date To UBound(vec_ticker_level(i), 1)
                    
                    If IsDate(vec_ticker_level(i)(m)(1)) = False Then 'N/A security etc.
                        matrix_output(i, j, 0) = vec_ticker_level(i)(0)(0) 'ticker
                        matrix_output(i, j, 1) = vec_distinct_date(j) 'date
                        
                        For k = 2 To UBound(vec_ticker_level(0)(0), 1)
                            matrix_output(i, j, k) = vec_ticker_level(i)(m)(1)
                        Next k
                        
                        last_position_date = 0
                    Else
                        date_tmp = vec_ticker_level(i)(m)(1)
                        
                        If date_tmp = vec_distinct_date(j) Then
                            For k = 0 To UBound(vec_ticker_level(0)(0), 1)
                                matrix_output(i, j, k) = vec_ticker_level(i)(m)(k)
                            Next k
                            
                            last_position_date = m
                            Exit For
                        ElseIf date_tmp > vec_distinct_date(j) Then
                            
                            'not found
                            matrix_output(i, j, 0) = vec_ticker_level(i)(0)(0) 'ticker
                            matrix_output(i, j, 1) = vec_distinct_date(j) 'date
                            
                            For k = 2 To UBound(vec_ticker_level(0)(0), 1)
                                matrix_output(i, j, k) = "#N/A data"
                            Next k
                            
                            last_position_date = m - 1
                            If last_position_date < 0 Then
                                last_position_date = 0
                            End If
                            
                            Exit For
                        Else
                            If m = UBound(vec_ticker_level(i), 1) Then
                                last_position_date = 0
                                
                                'not found
                                matrix_output(i, j, 0) = vec_ticker_level(i)(0)(0) 'ticker
                                matrix_output(i, j, 1) = vec_distinct_date(j) 'date
                                
                                For k = 2 To UBound(vec_ticker_level(0)(0), 1)
                                    matrix_output(i, j, k) = "#N/A data"
                                Next k
                            
                            End If
                        End If
                    End If
                Next m
            Next j
        Next i
        
        bdh = matrix_output
        
    End If

End Function



Public Function ProcessResponseEventBDH(eventObj As blpapicomLib.Event) As Variant

    Dim debug_test As Variant
    
    Dim i As Long, j As Long, k As Long, m As Long, n As Long, p As Long, q As Long
    
    
    Dim it As blpapicomLib.MessageIterator
    Set it = eventObj.CreateMessageIterator()
    
    Dim msg As blpapicomLib.Message
    
    Dim error As blpapicomLib.Element
    
    Dim hasErrorSecurity As Boolean
    
    Dim securityDataArray As blpapicomLib.Element
    Dim securityData As blpapicomLib.Element
    Dim numSecurities As Long
    Dim security As String
    Dim sequenceNumber As Long
    Dim dateArray As blpapicomLib.Element
    Dim fieldDataArray As blpapicomLib.Element
    Dim fieldData As blpapicomLib.Element
    Dim fields As blpapicomLib.Element
    Dim field As blpapicomLib.Element
    Dim numFields As Long
    
    
    
    Dim vec_level_ticker() As Variant
    Dim vec_level_date() As Variant
    Dim vec_level_data() As Variant
    
    Dim vec_level_ticker_with_header() As Variant
    Dim vec_level_date_with_header() As Variant
    Dim vec_level_data_with_header() As Variant
    
    Dim max_ubound_data As Long, max_ubound_date As Long, max_ubound_ticker As Long
        max_ubound_data = 0
        max_ubound_date = 0
        max_ubound_ticker = 0
    
    
    
    
    m = 0
    Do While it.Next() 'boucle sur les tickers
        
        hasErrorSecurity = False
        
        Set msg = it.Message
        
        'test si appel refus? (limit etc.)
        If msg.AsElement.HasElement("responseError") Then
            Set error = msg.GetElement("responseError")
            Call MsgBox("responseError: " & error.GetElement("message"))
            ProcessResponseEventBDH = Empty
            Exit Function
        End If
        
        Set securityDataArray = msg.GetElement("securityData")
        numSecurities = securityDataArray.NumValues
        
            
        Set securityData = securityDataArray 'boucle deja a travers msg iterrator
        sequenceNumber = securityData.GetElement("sequenceNumber").Value
        
        
        'check invalid security_id
        If securityData.HasElement("securityError") Then
            hasErrorSecurity = True
        End If
        
        
        ReDim Preserve vec_level_data(0)
        vec_level_data(0) = securityData.GetElement("security").Value 'ticker name
        
        
        'Process fieldData
        If securityData.HasElement("fieldData") Then
            Set dateArray = securityData.GetElement("fieldData")
            
            If hasErrorSecurity = True Then
                
                ReDim Preserve vec_level_data(1)
                vec_level_data(1) = "#N/A security"
                
                For k = LBound(orgiFields, 1) To UBound(orgiFields, 1)
                    ReDim Preserve vec_level_data(k + 2)
                    vec_level_data(k + 2) = "#N/A security"
                Next k
                
                ReDim Preserve vec_level_date(0)
                vec_level_date(0) = vec_level_data
                
            Else
                For j = 0 To dateArray.NumValues - 1 'boucle sur date
                    Set fieldDataArray = dateArray.GetValue(j)
                    numFields = fieldDataArray.NumElements
                    
                    'store date
                    ReDim Preserve vec_level_data(1)
                    vec_level_data(1) = fieldDataArray.GetElement("date").Value
                    
                    For k = LBound(orgiFields, 1) To UBound(orgiFields, 1)
                        
                        ReDim Preserve vec_level_data(k + 2)
                        
                        If fieldDataArray.HasElement(orgiFields(k)) Then
                            vec_level_data(k + 2) = fieldDataArray.GetElement(orgiFields(k)).Value
                        Else
                            If hasErrorSecurity = True Then
                                vec_level_data(k + 2) = "#N/A security"
                            Else
                                vec_level_data(k + 2) = "#N/A field"
                            End If
                        End If
                    Next k
                    
                    ReDim Preserve vec_level_date(j)
                    vec_level_date(j) = vec_level_data
                    
                Next j
            End If
        Else
    '        If hasErrorSecurity = True Then
    '            ReDim Preserve vec_level_data(1)
    '            vec_level_data(1) = "#N/A security"
    '
    '            For k = LBound(orgiFields, 1) To UBound(orgiFields, 1)
    '                ReDim Preserve vec_level_data(k + 2)
    '                vec_level_data(k + 2) = "#N/A security"
    '            Next k
    '        End If
        End If
        
        ProcessResponseEventBDH = vec_level_date
        
    Loop

End Function



Public Function eqs(ByVal eqs_screens_name As String, Optional ByVal eqs_screens_type_private_or_global = "PRIVATE", Optional ByVal eqs_folder = "") As Variant

    Dim i As Integer, j As Integer, k As Integer, m As Integer, n As Integer
    
    
    Dim debug_test As Variant
    
    session.OpenService ("//blp/refdata")
    Set refdataservice = session.GetService("//blp/refdata")
    
    Dim req As blpapicomLib.REQUEST
    
    Set req = refdataservice.CreateRequest("BeqsRequest")
    req.Set "screenName", eqs_screens_name
    req.Set "screenType", eqs_screens_type_private_or_global
    req.Set "Group", eqs_folder
    
    
    session.sendRequest req
    
    Dim eventObj As blpapicomLib.Event
    Dim it As blpapicomLib.MessageIterator
    Dim msg As blpapicomLib.Message
    
    Dim data As blpapicomLib.Element
    
    Dim fieldCount As Integer
    Dim numSecurities As Integer
    
    Dim security As blpapicomLib.Element
    Dim fields As blpapicomLib.Element
    Dim field As blpapicomLib.Element
    
    Dim numFields As Integer
    
    
    Dim vec_fields() As Variant
    Dim vec_output() As Variant
    
    Dim tmp_row() As Variant
    
    
    k = 0 'compteur de securities pour le vec output
    Do
        
        Set eventObj = session.NextEvent()
        
        If eventObj.EventType = PARTIAL_RESPONSE Or eventObj.EventType = Response Then
            
            Set it = eventObj.CreateMessageIterator()
            
            Do While it.Next()
                Set msg = it.Message
                
                If msg.AsElement.HasElement("responseError") Then
                    
                    eqs = Empty
                    Exit Function
                    
                Else
                    
                    Set data = msg.GetElement("data")
                    
                    If k = 0 Then
                        
                        For i = 0 To data.GetElement("fieldDisplayUnits").NumElements - 1
                            ReDim Preserve vec_fields(i)
                            vec_fields(i) = data.GetElement("fieldDisplayUnits").GetElement(i).Name
                        Next i
                        
                        ReDim Preserve vec_output(k)
                        vec_output(k) = vec_fields
                        k = k + 1
                        
                    End If
                    
                    
                    For i = 0 To data.GetElement("securityData").NumValues - 1
                        
                        ReDim Preserve vec_output(k)
                        m = 0
                        Set security = data.GetElement("securityData").GetValue(i)
                        
                        'patch du nom du ticker du titre
                        ReDim Preserve tmp_row(m)
                        tmp_row(m) = UCase(Left(security.GetElement("security").Value, InStr(security.GetElement("security").Value, " ") - 1)) & " " & UCase(Right(security.GetElement("security").Value, 2)) & " EQUITY"
                        m = m + 1
                        
                        
                        Set fields = security.GetElement("fieldData")
                        
                        For j = 1 To fields.NumElements - 1
                            
                            Set field = fields.GetElement(j)
                            
                            ReDim Preserve tmp_row(j)
                            tmp_row(j) = field.Value
                            
                        Next j
                        
                        vec_output(k) = tmp_row
                        
                        k = k + 1
                    Next i
                    
                End If
                
            Loop
            
            If eventObj.EventType = Response Then
                Exit Do
            End If
            
        End If
    Loop
    
    eqs = vec_output


End Function


Public Function portfolio(ByVal vec_portfolios As Variant, Optional ByVal vec_fields As Variant, Optional ByVal return_format = output_format.of_vec_without_header) As Variant

    Dim i As Long, j As Long, k As Long, m As Long, n As Long, p As Long, q As Long
    
    If IsMissing(vec_fields) Then
        vec_fields = Array("PORTFOLIO_DATA")
    End If
    
    
    session.OpenService ("//blp/refdata")
    Set refdataservice = session.GetService("//blp/refdata")
    
    Dim req As blpapicomLib.REQUEST
    
    Set req = refdataservice.CreateRequest("PortfolioDataRequest")
    
    Dim tmp_prt As String
    For i = LBound(vec_portfolios, 1) To UBound(vec_portfolios, 1)
        
        tmp_prt = vec_portfolios(i)
        
        If UCase(Right(vec_portfolios(i), 6)) <> "CLIENT" Then
            tmp_prt = tmp_prt & " CLIENT"
        End If
        
        req.GetElement("securities").AppendValue (tmp_prt)
    Next i
    
    For i = LBound(vec_fields, 1) To UBound(vec_fields, 1)
        req.GetElement("fields").AppendValue (vec_fields(i))
    Next i
    
    session.sendRequest req
    
    Dim eventObj As blpapicomLib.Event
    Dim it As blpapicomLib.MessageIterator
    Dim msg As blpapicomLib.Message
    
    Dim numSecurities As Long
    
    Dim security As blpapicomLib.Element
    Dim fieldArray As blpapicomLib.Element
    Dim field As blpapicomLib.Element
    
    Dim numfieldArray As Long
    
    Dim bulkElement As blpapicomLib.Element
    Dim numBulkElement As Long
    Dim elem As blpapicomLib.Element
    
    Dim vec_level_prt() As Variant
    Dim vec_level_field() As Variant
    Dim vec_bulk_array() As Variant
    Dim vec_bulk_one_line() As Variant
    
    k = 1
    
    Dim vec_header() As Variant
    Dim already_store_headers As Boolean
        already_store_headers = False
    
    Do
        Set eventObj = session.NextEvent()
        
        If eventObj.EventType = PARTIAL_RESPONSE Or eventObj.EventType = Response Then
            
            Set it = eventObj.CreateMessageIterator()
            
            Do While it.Next()
                Set msg = it.Message
                numSecurities = msg.GetElement("securityData").NumValues
                
                For i = 0 To msg.GetElement("securityData").NumValues - 1
                    Set security = msg.GetElement("securityData").GetValue(i)
                    
                    Set fieldArray = security.GetElement("fieldData")
                    numfieldArray = fieldArray.NumElements
                    
                    ReDim Preserve vec_header(fieldArray.NumElements)
                    vec_header(0) = "security_id"
                    
                    ReDim Preserve vec_level_field(0)
                    vec_level_field(0) = security.GetElement("security").Value
                    
                    For j = 0 To fieldArray.NumElements - 1
                        
                        Set field = fieldArray.GetElement(j)
                        
                        If field.DataType = 15 Then 'bulk
                            
                            For m = 0 To field.NumValues - 1
                                
                                Set bulkElement = field.GetValue(m)
                                numBulkElement = bulkElement.NumElements
                                
                                For n = 0 To bulkElement.NumElements - 1
                                    Set elem = bulkElement.GetElement(n)
                                    ReDim Preserve vec_bulk_one_line(n)
                                    vec_bulk_one_line(n) = elem.Value
                                    
                                Next n
                                
                                ReDim Preserve vec_bulk_array(m)
                                vec_bulk_array(m) = vec_bulk_one_line
                                
                            Next m
                            
                            ReDim Preserve vec_level_field(j + 1)
                            vec_level_field(j + 1) = vec_bulk_array
                            
                            If already_store_headers = False Then
                                ReDim Preserve vec_header(j + 1)
                                vec_header(j + 1) = field.Name
                            End If
                            
                        Else '1 data
                            ReDim Preserve vec_level_field(j + 1)
                            vec_level_field(j + 1) = field.Value
                            
                            If already_store_headers = False Then
                                ReDim Preserve vec_header(j + 1)
                                vec_header(j + 1) = field.Name
                            End If
                        End If
                        
                    Next j
                    
                    'mise en place du header
                    If already_store_headers = False Then
                        already_store_headers = True
                        ReDim Preserve vec_level_prt(0)
                        vec_level_prt(0) = vec_header
                    End If
                    
                    ReDim Preserve vec_level_prt(k)
                    vec_level_prt(k) = vec_level_field
                    k = k + 1
                    
                Next i
                
            Loop
            
            If eventObj.EventType = Response Then Exit Do
            
        End If
        
    Loop
    
    Dim matrix_output() As Variant
    
    Dim tmp_vec As Variant
    Dim tmp_vec_one_line() As Variant
    
    If return_format = output_format.of_vec_with_header Then
        portfolio = vec_level_prt
    ElseIf return_format = output_format.of_vec_without_header Then
        
        'retire la premiere ligne ainsi que la premiere colonne de chaque ligne
        
        tmp_vec = vec_level_prt
        
        ReDim vec_level_prt(UBound(tmp_vec, 1) - 1)
        
        m = 0
        For i = 1 To UBound(tmp_vec, 1)
            n = 0
            For j = 1 To UBound(tmp_vec(i), 1)
                ReDim Preserve tmp_vec_one_line(n)
                tmp_vec_one_line(n) = tmp_vec(i)(j)
                n = n + 1
            Next j
            
            ReDim Preserve vec_level_prt(m)
            vec_level_prt(m) = tmp_vec_one_line
            m = m + 1
        Next i
        
        portfolio = vec_level_prt
        
    ElseIf return_format = output_format.of_matrix_with_header Then
        
        ReDim matrix_output(UBound(vec_level_prt, 1), UBound(vec_level_prt(0), 1))
        
        For i = 0 To UBound(vec_level_prt, 1)
            For j = 0 To UBound(vec_level_prt(i), 1)
                matrix_output(i, j) = vec_level_prt(i)(j)
            Next j
        Next i
        
        portfolio = matrix_output
        
    ElseIf return_format = output_format.of_matrix_without_header Then
        
        ReDim matrix_output(UBound(vec_level_prt, 1) - 1, UBound(vec_level_prt(0), 1) - 1)
        
        m = 0
        For i = 1 To UBound(vec_level_prt, 1)
            n = 0
            For j = 1 To UBound(vec_level_prt(i), 1)
                matrix_output(m, n) = vec_level_prt(i)(j)
                n = n + 1
            Next j
            
            m = m + 1
        Next i
        
        portfolio = matrix_output
        
    End If



End Function



'vec_fields not necessary
'0 ticker / 1 time / 2 open / 3 high / 4 low / 5 close / 6 volume
Public Function intraday(ByVal vec_securities_id As Variant, ByVal datetime_from As Variant, ByVal datetime_to As Variant, ByVal interval_in_min As Integer) As Variant

    Dim i As Long, j As Long, k As Long, m As Long, n As Long, p As Long, q As Long
    
    Dim eventObj As blpapicomLib.Event
    Dim it As blpapicomLib.MessageIterator
    Dim msg As blpapicomLib.Message
    
    Dim numSecurities As Long
    
    Dim security As blpapicomLib.Element
    Dim fieldArray As blpapicomLib.Element
    Dim field As blpapicomLib.Element
    
    Dim numfieldArray As Long
    
    Dim elem As blpapicomLib.Element
    
    Dim tickData As Element
    Dim row_tick() As Variant
    
    
    session.OpenService ("//blp/refdata")
    Set refdataservice = session.GetService("//blp/refdata")
    
    Dim req As blpapicomLib.REQUEST
    Set req = refdataservice.CreateRequest("IntradayBarRequest")
    
    req.Set "interval", interval_in_min
    req.Set "eventType", "TRADE"
    
    
    Dim startTime As blpapicomLib.DateTime
        Set startTime = session.CreateDatetime
    Dim endTime As blpapicomLib.DateTime
        Set endTime = session.CreateDatetime
    
    
    Dim matrix_output() As Variant
    Dim matrix_time() As Variant
    
    Dim tmp_date_from As Date
    Dim tmp_date_to As Date
    
    For i = LBound(vec_securities_id, 1) To UBound(vec_securities_id, 1)
        
        req.Set "security", vec_securities_id(i)
        
        
        If IsArray(datetime_from) Then
            tmp_date_from = datetime_from(i)
        ElseIf IsDate(datetime_from) Then
            tmp_date_from = datetime_from
        Else
            
        End If
        
        
        If IsArray(datetime_to) Then
            tmp_date_to = datetime_to(i)
        ElseIf IsDate(datetime_to) Then
            tmp_date_to = datetime_to
        Else
            
        End If
        
        startTime.Year = Year(tmp_date_from)
        startTime.Month = Month(tmp_date_from)
        startTime.Day = Day(tmp_date_from)
        startTime.Hour = Hour(tmp_date_from)
        startTime.Minute = Minute(tmp_date_from)
        startTime.Second = Second(tmp_date_from)
    
        endTime.Year = Year(tmp_date_to)
        endTime.Month = Month(tmp_date_to)
        endTime.Day = Day(tmp_date_to)
        endTime.Hour = Hour(tmp_date_to)
        endTime.Minute = Minute(tmp_date_to)
        endTime.Second = Second(tmp_date_to)
        
        req.Set "startDateTime", startTime
        req.Set "endDateTime", endTime
        
        
        session.sendRequest req
        
        Do
            Set eventObj = session.NextEvent()
            
            If eventObj.EventType = PARTIAL_RESPONSE Or eventObj.EventType = Response Then
                
                Set it = eventObj.CreateMessageIterator()
                
                Do While it.Next()
                    Set msg = it.Message
                    
                    If msg.AsElement.HasElement("responseError") Then
                        GoTo store_data_output
                    Else
                    
                        Set tickData = msg.GetElement("barData").GetElement("barTickData")
                        
                        k = 0
                        For j = 0 To tickData.NumValues - 1 'boucle sur les bar
                            ReDim Preserve row_tick(6)
                            
                            For m = 0 To 5
                                row_tick(0) = vec_securities_id(i)
                                row_tick(m + 1) = tickData.GetValue(j).GetElement(m).Value
                            Next m
                            
                            ReDim Preserve matrix_time(k)
                            matrix_time(k) = row_tick
                            k = k + 1
                        Next j
                    End If
                    
                Loop
                
                If eventObj.EventType = Response Then Exit Do
                
            End If
            
        Loop
store_data_output:
        ReDim Preserve matrix_output(i)
        If k = 0 Then
            matrix_output(i) = Empty
        Else
            matrix_output(i) = matrix_time
        End If
        
    Next i
    
    intraday = matrix_output

End Function
