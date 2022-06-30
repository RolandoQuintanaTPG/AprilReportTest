from datetime import date
import datetime
import itertools
import numpy
from openpyxl import Workbook
import pandas as pd
import numpy as np
from IPython.display import display

import os

# cwd = os.getcwd()
# print(cwd)
returnsWB = pd.read_excel('main/DailyAccountingFile_TravelPass Group_2022-06-28.xlsx')
EverythingWB = pd.read_excel('main/EverythingFile2.xlsx')

reportOrderID = []
reportDate = []
reportExternalConf = []

reportCheckInDate = []
reportCheckOutDate = []

reportExpectedNetRate = []
reportActualNetRate = []

reportMerchantedAmount = []
reportRefundedAmount = []

#print(returnsWB['Description'].iloc[0])

refundIDs = []
IDsDates = {}
IDsNetRateRefund = {}

###Get Net Rate Refund
netRateDF = returnsWB[(returnsWB["Description"] == ("Purchase - Virtual Card Purchase")) | (returnsWB["Description"] == ("Purchase - Virtual Card Return"))]#
netRateDF = pd.pivot_table(netRateDF, index="OrderNumber", values="Amount", aggfunc=np.sum)
#display(netRateDF)
for row in netRateDF.itertuples():
    IDsNetRateRefund[row[0]] = row[1]




###Get Return Order IDs and Refunded Amounts
for row in returnsWB.itertuples():
    if(row[3] == 'Sale - Credit Card Return'): #check description column
        IDsDates[row[5]] = [row[4], row[10]]

###append report lists
print("Searching for ID's")
for refundID in IDsDates:
    df = EverythingWB[EverythingWB.values == refundID]
    if(len(df['OrderNumber'].values) != 0):
        reportOrderID.append(df['OrderNumber'].values[0])
        reportDate.append(IDsDates.get(refundID)[0])
        reportExternalConf.append(df['ExternalConf'].values[0])

        reportCheckInDate.append(df['Arrival'].values[0])
        reportCheckOutDate.append(df['Departure'].values[0])

        reportExpectedNetRate.append(df['Net Rate'].values[0])
        if refundID in IDsNetRateRefund.keys():
            reportActualNetRate.append(IDsNetRateRefund.get(refundID))
        else:
            reportActualNetRate.append('')

        reportMerchantedAmount.append(df['Merchanted Amount'].values[0])
        reportRefundedAmount.append(IDsDates.get(refundID)[1])
    else:
        reportOrderID.append(refundID)
        reportDate.append(IDsDates.get(refundID)[0])
        reportExternalConf.append('')

        reportCheckInDate.append('')
        reportCheckOutDate.append('')

        reportMerchantedAmount.append('')
        reportRefundedAmount.append(IDsDates.get(refundID)[1])

        reportExpectedNetRate.append('')

        if refundID in IDsNetRateRefund.keys():
            reportActualNetRate.append(IDsNetRateRefund.get(refundID))
        else:
            reportActualNetRate.append('')

dict = {'OrderID':reportOrderID, 'Date Processed':reportDate, 'External Confirmation':reportExternalConf, 'Check In Date': reportCheckInDate, 
'Check Out Date' : reportCheckOutDate, 'Expected Net Rate':reportExpectedNetRate, 'Net Rate Refund':reportActualNetRate, 'Merchanted Amount':reportMerchantedAmount, 'Refunded Amount':reportRefundedAmount}

###Create data frame from lists
FinalDF = pd.DataFrame(dict)

###Format Dates
FinalDF['Date Processed'] = FinalDF['Date Processed'].dt.strftime('%m/%d/%Y')
FinalDF['Check In Date'] = FinalDF['Check In Date'].dt.strftime('%m/%d/%Y')
FinalDF['Check Out Date'] = FinalDF['Check Out Date'].dt.strftime('%m/%d/%Y')

###Create final report
FinalDF.to_excel(excel_writer='RefundReport({}).xlsx'.format(datetime.now()),sheet_name="Refund report".format(datetime.now()))


