# Bloomberg API Examples

## Python Script
### Bloomberg Historical Data Request (BDH)
df = BDH("AAPL US EQUITY", "PX_LAST", "20150101", "20150103")

### Bloomberg Data Point Request (BDP)
df = BDP("AAPL US EQUITY", "PX_LAST")

### Bloomberg Data Set Request (BDS)
df = BDS("AAPL US EQUITY", "DVD_HIST_ALL")

## Excel VBA

BCOM_wrapper is a wrapper of Bloomberg COM API

BCOM_test demonstrates example usages of BCOM_wrapper

tester_referenceData is like BDP

tester_bulkReferenceData is like BDS

tester_historicalData is like BDH
