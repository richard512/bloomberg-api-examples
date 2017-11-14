Private Function getportfolio()
    Dim bloom As New BloomSync
    Dim result As Variant
    Dim s As Variant
    Dim f As Variant
    ReDim s(0 To 0): s(0) = "U12096528-104 client"
    ReDim f(0 To 0): f(0) = "portfolio_mposition"
    result = bloom.portfolio(s, f)
    printBulkReferenceData result
End Function

Private Sub GetMyBloomingFields()
    Dim tickers() As Variant
    Dim fields() As Variant
    Dim results As Variant
    
    tickers = Array( _
            "ADBE US Equity", _
            "ADSK US Equity" _
            )
    fields = Array( _
            "Country_Full_Name", _
            "CNTRY_ISSUE_ISO", _
            "CRNCY", _
            "DVD_EX_Dt", _
            "Quantity", _
            "DAY_TO_DAY_TOT_RETURN_NET_DVDS", _
            "EQY_BETA_STDDEV_ERR", _
            "EQY_BETA_STDDEV_ERR", _
            "PX_TO_BOOK_RATIO", _
            "EQY_DVD_YLD_IND", _
            "CUR_MKT_CAP", _
            "Short_Int_Ratio" _
            )
    results = bloomBDP(tickers, fields)
End Sub

Private Function bloomBDP(tickers() As Variant, bbg_field() As Variant) As Variant
    'MsgBox bloomBDP("CMG US Equity", "Country_Full_Name")
    'MsgBox bloomBDP("CMG US Equity", "CNTRY_ISSUE_ISO")
    'MsgBox bloomBDP("CMG US Equity", "CRNCY")
    'MsgBox bloomBDP("CMG US Equity", "DVD_EX_Dt")
    'Range("C2").Value = bloomBDP("CMG US Equity", "Quantity")
    'MsgBox bloomBDP("CMG US Equity", "DAY_TO_DAY_TOT_RETURN_NET_DVDS")
    'MsgBox bloomBDP("CMG US Equity", "EQY_BETA_STDDEV_ERR")
    'MsgBox bloomBDP("CMG US Equity", "PX_TO_BOOK_RATIO")
    'MsgBox bloomBDP("CMG US Equity", "EQY_DVD_YLD_IND")
    'MsgBox bloomBDP("CMG US Equity", "CUR_MKT_CAP")
    'MsgBox bloomBDP("CMG US Equity", "Short_Int_Ratio")
    '"Country_Full_Name", "CNTRY_ISSUE_ISO", "CRNCY", "DVD_EX_Dt", "Quantity", "DAY_TO_DAY_TOT_RETURN_NET_DVDS", "EQY_BETA_STDDEV_ERR", "EQY_BETA_STDDEV_ERR", "PX_TO_BOOK_RATIO", "EQY_DVD_YLD_IND", "CUR_MKT_CAP"
    Dim bloom As New BloomSync
    Dim bdp_data As Variant
    
    bdp_data = bloom.bdp(tickers, bbg_field, output_format.of_vec_without_header)
    
    For i = 0 To UBound(bdp_data)
        For j = 0 To UBound(bbg_field)
            Debug.Print "bdp_data(" & i & "," & j & ")("; tickers(i) & "." & bbg_field(j) & ") = " & bdp_data(i)(j)
            'Debug.Print "bdp_data(" & i & "," & j & ") = " & (bdp_data(i)(j)(0)(0))
        Next j
    Next i
    
    bloomBDP = bdp_data(0)(0)
    'MsgBox bdp_data(0)(CHG_PCT_1D)
'    MsgBox "result = " & TypeName(result) & vbNewLine & _
'        "ubound = " & UBound(result) & vbNewLine & _
'        "empty = " & result(1)
End Function

Private Sub GetBloomingPortfolio()
    '=BDS("U12096528-104 client","portfolio_mposition","cols=2;rows=47")
    Dim tickers() As Variant
    Dim fields() As Variant
    Dim results As Variant
    
    tickers = Array("U12096528-104 client")
    fields = Array("portfolio_mposition")
    results = bloomPortfolio(tickers, fields)
End Sub

Private Function bloomPortfolio(tickers() As Variant, bbg_field() As Variant) As Variant
    Dim bloom As New BloomSync
    Dim ticker As String
    Dim bdp_data As Variant

    
    
    ticker = security
    
    'bdp_data = bloom.bdp(tickers, bbg_field, output_format.of_vec_without_header)
    bdp_data = bloom.portfolio(tickers, bbg_field)
    Dim px_last As Double
    px_last = 0
    
'    If IsNumeric(bdp_data(0)(dim_px_last)) = True And Left(bdp_data(0)(dim_px_last), 1) <> "#" Then
'        TgglBtn_Price_Last.Caption = bdp_data(0)(dim_px_last)
'        px_last = bdp_data(0)(dim_px_last)
'        MsgBox px_last
'    End If
    
    For i = 0 To UBound(bdp_data)
        For j = 0 To UBound(bdp_data(i))
            For k = 0 To UBound(bdp_data(i)(j))
                'Debug.Print "bdp_data(" & i & "," & j & ")("; tickers(i) & "." & bbg_field(j) & ") = " & bdp_data(i)(j)
                Debug.Print "bdp_data(" & i & "," & j & "," & k & ") = " & (bdp_data(i)(j)(k)(0)) & " | " & (bdp_data(i)(j)(k)(1))
            Next k
        Next j
    Next i
    
    bloomPortfolio = bdp_data(0)(0)
    'MsgBox bdp_data(0)(CHG_PCT_1D)
'    MsgBox "result = " & TypeName(result) & vbNewLine & _
'        "ubound = " & UBound(result) & vbNewLine & _
'        "empty = " & result(1)
End Function

Private Function printBulkReferenceData(ByRef data As Variant)
    '
    'Dim rng As Range: Set rng = Sheets("Sheet1").Range("A1")
    'rng.CurrentRegion.ClearContents
    Dim i As Long, j As Long
    Dim datalength As Long
    datalength = UBound(data)
    If datalength < 1 Then
        MsgBox "printBulkReferenceData: No data"
        Exit Function
    End If
    '
    On Error GoTo ErrorInData
    For i = 0 To UBound(data, 1)
        For j = 0 To UBound(data, 2)
            If (data(i, j)) Then
                Debug.Print i & "," & j & " = " & data(i, j)
                'rng(j + 1, i + 1) = data(i, j)
            End If
        Next j
    Next i
    Exit Function
ErrorInData:
    MsgBox "printBulkReferenceData: Error in data"
End Function
