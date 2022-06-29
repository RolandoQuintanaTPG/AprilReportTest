import itertools
import numpy
from openpyxl import Workbook
import pandas as pd
import os

# cwd = os.getcwd()
# print(cwd)
returnsWB = pd.read_excel('main/DailyAccountingFile_TravelPass Group_2022-06-28.xlsx')
EverythingWB = pd.read_excel('main/EverythingFile2.xlsx')

reportOrderID = []
reportExternalConf = []

reportCheckInDate = []
reportCheckOutDate = []

reportExpectedNetRate = []
reportActualNetRate = []

reportMerchantedAmount = []
reportRefundedAmount = []

#print(returnsWB['Description'].iloc[0])

refundIDs = []

###Get Return Order IDs
for row in returnsWB.itertuples():
    if(row[3] == 'Sale - Credit Card Return'): #check description column
        refundIDs.append(row[5])

# for order in EverythingWB.values[:,4]:
#     print(order)

###append report lists
currentIndx = 0
currentRefundID = refundIDs[0]
# for refundID, everythingID in itertools.product(refundIDs, EverythingWB.values[:,4]):
#     if currentRefundID != refundID:
#         currentRefundID = refundID
#         currentIndx = 0
#     #print("RefundID: {}         EverythingID: {}".format(refundID, everythingID))
#     if refundID == everythingID:
#         reportExternalConf.append(EverythingWB['External'].values[currentIndx])
#     currentIndx += 1
print("Searching for ID's")
for refundID in refundIDs:
    df = EverythingWB[EverythingWB.values == refundID]
    try:
        reportOrderID.append(df['OrderNumber'].values[0])
        reportExternalConf.append(df['ExternalConf'].values[0])
        reportCheckInDate.append(df['Arrival'].values[0])
        reportCheckOutDate.append(df['Departure'].values[0])
    except:
        continue

dict = {'OrderID':reportOrderID, 'External Confirmation':reportExternalConf, 'Check In Date': reportCheckInDate, 'Check Out Date' : reportCheckOutDate}

FinalDF = pd.DataFrame(dict)

#for item in reportOrderID:
 #   print(item)
FinalDF.to_excel(excel_writer='report.xlsx',sheet_name="Report")


