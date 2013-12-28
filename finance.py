from bs4 import BeautifulSoup
import urllib2
import urllib
import sys

sys.path.append("D:\\PythonLib")
from cvsConversion import cvsConversionExcel
from ExcelOp import *
                
class Crawlwer:
    def __init__(self,reqString,quota):
        #self.url=value
        #self.targetstring=tstring
        #raw_input("please input the string you want to parse:").rstrip()
        self.url_login=reqString
        self.params=quota
 

    def loginin(self):
       request_url=self.url_login
       
       data=self.params
       
       url_params=urllib.urlencode(data)
      # print "data:",data
      # print "url_params:",url_params
      # print "request_url:",request_url
     
       final_url = request_url + "?" + url_params+"+Historical+Prices"
       print final_url
       return final_url
        
def readFinalString(urlstring,tarString,paramData):
        
        #urlstring=self.url
        #tarString=self.targetstring
        #print urlstring
        #print tarString
        print "read final dict:" ,paramData
        filename=paramData.get('s')
        #print filename
        
        fileopen=urllib2.urlopen(urlstring)
        content=fileopen.read()
        soup=BeautifulSoup(content)
        for links in soup.find_all("a"):
            ##print(links.get('href'))
            if tarString in links.get('href'):
                print "The csv's link is:", links.get("href")
                u=urllib.urlopen(links.get("href"))
                location="d:\\"+filename[::]+".csv"          
                print location
                localfile=open(location,'w')
                localfile.write(u.read())
                u.close()
                localfile.close()
        return location


    
##        #print req
##        try:
##          req=urllib2.Request(request_url,url_values)
##          response=urllib2.urlopen(req)
##          print req
##        except urllib2.HTTPError,e:
##            print "The server couldnot fulfil this request"
##            print "Error Reason:", e.reason
##        except urllib2.URLError,e:
##            print "We failed to reachh a server."
##            print "Reason:",e.reason        
##        #print response.read()
        
    
if __name__=='__main__':
    #inputurl=raw_input("please input the string you want to parse:")
    #inputurl='http://finance.yahoo.com/q/hp?s=TWTR+Historical+Prices'
    #targetstring1=raw_input("please input the target string you want to hit:")
    
    targetString1='table.csv'
    
    ##reqString=raw_input("please input the request string url:")
    
    reqString='http://finance.yahoo.com/q/hp'
    input_quota=raw_input("please input the quota:").rstrip().upper()
    #e.g: input_quota="AAPL"
    paramData=dict(zip(['s'],[input_quota]))
    #print paramData
    
    c=Crawlwer(reqString,paramData)
    finalString=c.loginin()
    print "Final String is :", finalString
    filelocation=readFinalString(finalString,targetString1,paramData)
    print "source file:" ,filelocation
    dest_file=cvsConversionExcel(filelocation)
    print "dest file location:",dest_file
    newSheet=raw_input("please input the sheet name you want to add:").rstrip()
    excelObject=excelOp(dest_file,newSheet)
    excelObject.excelChart()
    
    
    
        
        
        
