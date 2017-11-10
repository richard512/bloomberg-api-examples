# bloomberg

## Python Script
### Bloomberg Historical Data Request (BDH)
df = BDH("AAPL US EQUITY", "PX_LAST", "20150101", "20150103")

### Bloomberg Data Point Request (BDP)
df = BDP("AAPL US EQUITY", "PX_LAST")

### Bloomberg Data Set Request (BDS)
df = BDS("AAPL US EQUITY", "DVD_HIST_ALL")
