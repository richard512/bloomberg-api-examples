Option Explicit
'
Private b As BCOM_wrapper
Private r As Variant
Private s() As String
Private f() As String
Private overrideFields() As String
Private overrideValues() As String
'
Sub tester_referenceData()
    '
    ' create wrapper object
    Set b = New BCOM_wrapper
    '
    ' create 3 securities and 4 fields
    ReDim s(0 To 2): s(0) = "GS US Equity": s(1) = "DBK GR Equity": s(2) = "JPM US Equity"
    ReDim f(0 To 3): f(0) = "SECURITY_NAME": f(1) = "BEST_EPS": f(2) = "BEST_PE_RATIO": f(3) = "BEST_DIV_YLD"
    '
    ' retrieve result from wrapper into array and print
    r = b.referenceData(s, f)
    printReferenceData r
    '
    ' create 1 override for fields
    ReDim overrideFields(0 To 0): overrideFields(0) = "BEST_FPERIOD_OVERRIDE"
    ReDim overrideValues(0 To 0): overrideValues(0) = "3FY"
    '
    ' retrieve result from wrapper into array and print
    r = b.referenceData(s, f, overrideFields, overrideValues)
    printReferenceData r
    '
    ' release wrapper object
    Set b = Nothing
End Sub
'
Sub tester_bulkReferenceData()
    '
    ' create wrapper object
    Set b = New BCOM_wrapper
    '
    ' create 3 securities and 1 fields
    ReDim s(0 To 2): s(0) = "GS US Equity": s(1) = "DBK GR Equity": s(2) = "JPM US Equity"
    ReDim f(0 To 0): f(0) = "BOND_CHAIN"
    '
    ' retrieve result from wrapper into array and print
    r = b.bulkReferenceData(s, f)
    printBulkReferenceData r
    '
    ' create 2 overrides for chain
    ReDim overrideFields(0 To 1): overrideFields(0) = "CHAIN_CURRENCY": overrideFields(1) = "CHAIN_COUPON_TYPE"
    ReDim overrideValues(0 To 1): overrideValues(0) = "JPY": overrideValues(1) = "FLOATING"
    '
    ' retrieve result from wrapper into array and print
    r = b.bulkReferenceData(s, f, overrideFields, overrideValues)
    printBulkReferenceData r
    '
    ' release wrapper object
    Set b = Nothing
End Sub
'
Sub tester_historicalData()
    '
    ' create wrapper object
    Set b = New BCOM_wrapper
    '
    ' create 3 securities and 4 fields
    ReDim s(0 To 2): s(0) = "GS US Equity": s(1) = "DBK GR Equity": s(2) = "JPM US Equity"
    ReDim f(0 To 3): f(0) = "PX_OPEN": f(1) = "PX_LOW": f(2) = "PX_HIGH": f(3) = "PX_LAST"
    '
    ' retrieve result from wrapper into array
    r = b.historicalData(s, f, CDate("21.8.2008"), CDate("21.8.2013"), , , "ALL_CALENDAR_DAYS", "PREVIOUS_VALUE")
    printHistoricalData r
    '
    ' release wrapper object
    Set b = Nothing
End Sub
'
Private Function printReferenceData(ByRef data As Variant)
    '
    Dim rng As Range: Set rng = Sheets("Sheet1").Range("A1")
    rng.CurrentRegion.ClearContents
    Dim i As Long, j As Long
    '
    On Error Resume Next
    For i = 0 To UBound(data, 1)
        For j = 0 To UBound(data, 2)
            rng(i + 1, j + 1) = data(i, j)
        Next j
    Next i
End Function
'
Private Function printBulkReferenceData(ByRef data As Variant)
    '
    Dim rng As Range: Set rng = Sheets("Sheet1").Range("A1")
    rng.CurrentRegion.ClearContents
    Dim i As Long, j As Long
    '
    On Error Resume Next
    For i = 0 To UBound(data, 1)
        For j = 0 To UBound(data, 2)
            rng(j + 1, i + 1) = data(i, j)
        Next j
    Next i
End Function
'
Private Function printHistoricalData(ByRef data As Variant)
    '
    Dim rng As Range: Set rng = Sheets("Sheet1").Range("A1")
    rng.CurrentRegion.ClearContents
    Dim i As Long, j As Long, k As Long: k = 1
    '
    On Error Resume Next
    For i = 0 To UBound(data, 1)
        For j = 0 To UBound(data, 2)
            rng(j + 1, i + k) = data(i, j)(0)
            rng(j + 1, i + k + 1) = data(i, j)(1)
            rng(j + 1, i + k + 2) = data(i, j)(2)
            rng(j + 1, i + k + 3) = data(i, j)(3)
        Next j
        '
        k = k + 3
    Next i
End Function