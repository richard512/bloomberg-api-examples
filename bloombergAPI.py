"""
    A wrapper for the Bloomberg API.
    BDH(): BDH()
    Excel BDP: BDP()
    Excel BDS: BDS()
"""
# Load Bloomberg Python API blpapi
# Download from https://www.bloomberglabs.com/api/libraries/

import blpapi
import pandas as pd
import numpy as np
from datetime import datetime

class RequestError(Exception):
    """A RequestError is raised when there is a problem with a Bloomberg API response."""
    def __init__ (self, value, description):
        self.value = value
        self.description = description
        
    def __str__ (self):
        return self.description + '\n\n' + str(self.value)

class bloombergAPI:
    """ A wrapper for the Bloomberg API that returns DataFrames.  This class
        manages a //BLP/refdata service and therefore does not handle event
        subscriptions.
    
        All calls are blocking and responses are parsed and returned as 
        DataFrames where appropriate. 
    
        A RequestError is raised when an invalid security is queried.  Invalid
        fields will fail silently and may result in an empty DataFrame.
    """ 
    def __init__ (self, host='localhost', port=8194, open=True):
        self.active = False
        self.host = host
        self.port = port
        if open:
            self.open()
        
    def open (self):
        if not self.active:
            sessionOptions = blpapi.SessionOptions()
            sessionOptions.setServerHost(self.host)
            sessionOptions.setServerPort(self.port)
            self.session = blpapi.Session(sessionOptions)
            self.session.start()
            self.session.openService('//blp/refdata')
            self.refDataService = self.session.getService('//blp/refdata')
            self.active = True
    
    def close (self):
        if self.active:
            self.session.stop()
            self.active = False

    def BDH (self, securities, fields, startDate, endDate, **kwargs):
        """ Equivalent to the Excel BDH Function.
        
            If securities are provided as a list, the returned DataFrame will
            have a MultiIndex.
        """
        defaults = {'startDate'       : startDate,
            'endDate'                 : endDate,
            'periodicityAdjustment'   : 'ACTUAL',
            'periodicitySelection'    : 'DAILY',
            'nonTradingDayFillOption' : 'ACTIVE_DAYS_ONLY',
            'adjustmentNormal'        : True,
            'adjustmentAbnormal'      : True,
            'adjustmentSplit'         : True,
            'adjustmentFollowDPDF'    : False}   
        defaults.update(kwargs)

        response = self.sendRequest('HistoricalData', securities, fields, defaults)
        
        data = []
        keys = []
        
        for msg in response:
            securityData = msg.getElement('securityData')
            fieldData = securityData.getElement('fieldData')
            fieldDataList = [fieldData.getValueAsElement(i) for i in range(fieldData.numValues())]
            
            df = pd.DataFrame()
            
            for fld in fieldDataList:
                for v in [fld.getElement(i) for i in range(fld.numElements()) if fld.getElement(i).name() != 'date']:
                    df.ix[fld.getElementAsDatetime('date'), str(v.name())] = v.getValue()

            df.index = df.index.to_datetime()
            df.replace('#N/A History', np.nan, inplace=True)
            
            keys.append(securityData.getElementAsString('security'))
            data.append(df)
        
        if len(data) == 0:
            return pd.DataFrame()
        if type(securities) == str:
            data = pd.concat(data, axis=1)
            data.columns.name = 'Field'
        else:
            data = pd.concat(data, keys=keys, axis=1, names=['Security','Field'])
            
        data.index.name = 'Date'
        return data
        
    def BDP (self, securities, fields, **kwargs):
        """ Equivalent to the Excel BDP Function.
        
            If either securities or fields are provided as lists, a DataFrame
            will be returned.
        """
        response = self.sendRequest('ReferenceData', securities, fields, kwargs)
        
        data = pd.DataFrame()
        
        for msg in response:
            securityData = msg.getElement('securityData')
            securityDataList = [securityData.getValueAsElement(i) for i in range(securityData.numValues())]
            
            for sec in securityDataList:
                fieldData = sec.getElement('fieldData')
                fieldDataList = [fieldData.getElement(i) for i in range(fieldData.numElements())]
                
                for fld in fieldDataList:
                    data.ix[sec.getElementAsString('security'), str(fld.name())] = fld.getValue()
        
        if data.empty:
            return data
        else: 
            data.index.name = 'Security'
            data.columns.name = 'Field'
            return data.iloc[0,0] if ((type(securities) == str) and (type(fields) == str)) else data
        
    def BDS (self, securities, fields, **kwargs):
        """ Equivalent to the Excel BDS Function.
        
            If securities are provided as a list, the returned DataFrame will
            have a MultiIndex.
            
            You may pass a list of fields to a bulkRequest.  An appropriate
            Index will be generated, however such a DataFrame is unlikely to
            be useful unless the bulk data fields contain overlapping columns.
        """
        defaults = { 'strOverrideField' : 'END_DATE_OVERRIDE',
                     'strOverrideValue' : '20170131'}
        defaults.update(kwargs)

        response = self.sendRequest('ReferenceData', securities, fields, defaults)

       
        for msg in response:
            securityData = msg.getElement('securityData')
            securityDataList = [securityData.getValueAsElement(i) for i in range(securityData.numValues())]
            
            for sec in securityDataList:
                fieldData = sec.getElement('fieldData')
                fieldDataList = [fieldData.getElement(i) for i in range(fieldData.numElements())]
                
                df = pd.DataFrame()
                
                for fld in fieldDataList:
                    for v in [fld.getValueAsElement(i) for i in range(fld.numValues())]:
                        s = pd.Series()
                        for d in [v.getElement(i) for i in range(v.numElements())]:
                            try:
                                s[str(d.name())] = d.getValue()
                            except:
                                s[str(d.name())] = np.nan                                
                        df = df.append(s, ignore_index=True)            
        return df
        
    def sendRequest (self, requestType, securities, fields, elements):
        """ Prepares and sends a request then blocks until it can return 
            the complete response.
            
            Depending on the complexity of your request, incomplete and/or
            unrelated messages may be returned as part of the response.
        """
        request = self.refDataService.createRequest(requestType + 'Request')
        
        if type(securities) == str:
            securities = [securities]
        if type(fields) == str:
            fields = [fields]
        
        for s in securities:
            request.getElement("securities").appendValue(s)
        for f in fields:
            request.getElement("fields").appendValue(f)

        if 'strOverrideField' in elements:
            o = request.getElement('overrides').appendElement()
            o.setElement('fieldId', elements['strOverrideField'])
            o.setElement('value', elements['strOverrideValue'])     
        else:
            for k, v in elements.items():
                if type(v) == datetime:
                    v = v.strftime('%Y%m%d')
                request.set(k, v)
       
        self.session.sendRequest(request)

        response = []
        while True:
            event = self.session.nextEvent()
            for msg in event:
                if msg.hasElement('responseError'):
                    raise RequestError(msg.getElement('responseError'), 'Response Error')
                if msg.hasElement('securityData'):
                    if msg.getElement('securityData').hasElement('fieldExceptions') and (msg.getElement('securityData').getElement('fieldExceptions').numValues() > 0):
                        raise RequestError(msg.getElement('securityData').getElement('fieldExceptions'), 'Field Error')
                    if msg.getElement('securityData').hasElement('securityError'):
                        raise RequestError(msg.getElement('securityData').getElement('securityError'), 'Security Error')
                
                if msg.messageType() == requestType + 'Response':
                    response.append(msg)
                
            if event.eventType() == blpapi.Event.RESPONSE:
                break
                
        return response

    def __enter__ (self):
        self.open()
        return self
        
    def __exit__ (self, exc_type, exc_val, exc_tb):
        self.close()

    def __del__ (self):
        self.close()

if __name__ == '__main__':
    try:
        blp = bbAPI()
        blp.BDH("AAPL US EQUITY","PX_LAST", "20150101", "20150102"))
    except Exception as msg:
        print(msg)
