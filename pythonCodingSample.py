import whois as whois
import pandas as pd
from xlsxwriter.utility import xl_rowcol_to_cell
import threading
import math

"""Given an excel sheet of URL's we want to learn more about, this file parses
the excel sheet, cleans up the URL's, and sends batches of requests on a timer
to whois. It then creates an excel sheet with each URL and the whois response
"""


def main(filepathToURLSpreadsheet):
    listOfURLs = makeURLList(filepathToURLSpreadsheet)
    whoIsData, errorURLs = getWhoIsData({}, [], listOfURLs, 0, len(listOfURLs)/5)
    makeSpreadsheet(whoIsData, 'whoIsData.xlsx', errorURLs)
    #this length should be len(listOfURLs) minus how many rows are in the spreadsheet
    print len(errorURLs)
    
#return a list of all the urls of the sites we want to look at & clean them up
def makeURLList(filename):
    news_df = pd.read_excel(filename)
    listOfURLS = news_df['Site name']
    #all these urls we just got have a backslash at the end and the http://
    #so whois doesn't work on them, so we have to get rid of that
    for i in range(len(listOfURLS)):
        listOfURLS[i] = str(listOfURLS[i])
        if listOfURLS[i][-1] == '/':
            listOfURLS[i] = listOfURLS[i][:-1]
        if listOfURLS[i][0:7] == 'http://':
            listOfURLS[i] = listOfURLS[i][7:]
    return listOfURLS

def getWhoIsData(dataDict, errorURLList, sitesList, startIndex, divisionFactor):
    #divisionFactor is the number of chunks sitesList will be split in
    if startIndex >= len(sitesList):
        #if we've gone through all the data
        print "finished getting data"
        return dataDict, errorURLList
    else:
        stopIndex = startIndex + int(math.ceil(float(len(sitesList)/divisionFactor)))
        for index in range(startIndex, stopIndex):
            #for each website in this section of the list, get its data
            try:
                dataDict[sitesList[index]] = whois.query(sitesList[index]).__dict__
            except IndexError:
                pass
                '''
                this program correctly computes how many chunks it needs to 
                break the list up into but it won't correctly stop in the middle
                of the last chunk if the last chunk isn't full.  this means that 
                this error will only happen on the last chunk so it's ok to 
                catch it and not worry
                '''
            except:
                #if this site isn't on whois:
                    errorURLList.append(sitesList[index])
        #recursive call on a timer with stopIndex as the new startIndex
        return threading.Timer(1, lambda: getWhoIsData(dataDict, errorURLList, \
        sitesList, stopIndex, divisionFactor)).start()
 
def makeSpreadsheet(dictOfData, filename, errorURLs):
    df = pd.DataFrame.from_dict(dictOfData)
    writer = pd.ExcelWriter(filename, engine='xlsxwriter', 
                            datetime_format='mmm d yyyy hh:mm:ss',
                            date_format='mmmm dd yyyy')
    #switching the rows and columns:
    df = df.T
    #make an excel sheet:
    df.to_excel(writer, index=False, sheet_name='Sheet 1')
    #format columns:
    worksheet = writer.sheets['Sheet 1']
    worksheet.set_column('A:D', 20)
    worksheet.set_column('E:F', 40)
