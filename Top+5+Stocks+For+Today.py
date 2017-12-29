
# coding: utf-8

# In[123]:


import urllib2


# In[124]:


urlOfFileName = "http://www.nseindia.com/content/historical/EQUITIES/2017/DEC/cm28DEC2017bhav.csv.zip"


# In[125]:


urlOfFileName


# In[126]:


localZipFilePath = "C:/Users/Tejaswi/Desktop/cm28DEC2017bhav2.csv.zip"


# In[127]:


localZipFilePath


# In[128]:


#BoilerPlate code to cross the barricade and download the file from NSE because NSE blocks the Automated programs.
hdr = {'User-Agent':'Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.84 Safari/537.36'}


# In[129]:


hdr


# In[130]:


webRequest = urllib2.Request(urlOfFileName,headers=hdr)


# In[131]:


try:
    page = urllib2.urlopen(webRequest)
    content = page.read()
    output = open(localZipFilePath,"wb")
    output.write(bytearray(content))
    output.close()
except urllib2.HTTPError,e:
    print e.fp.read()
    print "Looks like the file did'nt go through"
    


# In[132]:


import zipfile,os


# In[133]:


localExtractFilePath = "C:/Users/Tejaswi/Desktop/"


# In[134]:


if os.path.exists(localZipFilePath):
    print  "Cool!" + localZipFilePath + "Exist"
    listOfFiles = []
    fh = open(localZipFilePath,'rb')
    zipFileHandler = zipfile.ZipFile(fh)
    for fileName in zipFileHandler.namelist():
        zipFileHandler.extract(fileName,localExtractFilePath)
        listOfFiles.append(localExtractFilePath + fileName)
        print "Extracted" + fileName
        print "We have extracted " , str(len(listOfFiles))
        fh.close()


# In[135]:


import csv


# In[136]:


oneFileName = listOfFiles[0]


# In[137]:


lineNum = 0


# In[138]:


listOfLists = []


# In[139]:


with open(oneFileName,'rb') as csvfile:
    lineReader = csv.reader(csvfile,delimiter=",",quotechar="\"")
    for row in lineReader:
        lineNum = lineNum +1
        if lineNum ==1:
            print "skipping the header row"
            continue
        symbol = row[0]
        close = row[5]
        prevClose = row[7]
        tradedQty = row[9]
        pctChange = float(close)/float(prevClose)-1
        oneResultRow = [symbol,pctChange,float(tradedQty)]
        listOfLists.append(oneResultRow)
        print symbol,"{:,.1f}".format(float(tradedQty)/1e6) + "M INR","{:,.1f}".format(pctChange*100)+"%"
        


# In[140]:


listOfListsSortedByQty = sorted(listOfLists,key=lambda x:x[2],reverse=True)


# In[141]:


listOfListsSortedByQty = sorted(listOfLists,key=lambda x:x[1],reverse=True)


# In[142]:


listOfListsSortedByQty


# In[143]:


import xlsxwriter


# In[144]:


excelFileName = "C:/Users/Tejaswi/Desktop/28DEC2017.xlsx"


# In[145]:


workbook = xlsxwriter.Workbook(excelFileName)


# In[146]:


worksheet = workbook.add_worksheet("Summary")


# In[147]:


worksheet.write_row("A1",["Top Traded Stocks"])
worksheet.write_row("A2",["Stocks","%Change","Total Traded Value"])
for rowNum in range(10):
    oneRowToWrite = listOfListsSortedByQty[rowNum]
    worksheet.write_row("A" + str(rowNum+3),oneRowToWrite)
workbook.close()    

