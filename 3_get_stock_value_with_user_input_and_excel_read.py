


import xlrd
import urllib
import time
import webbrowser
import sys


book = xlrd.open_workbook("smalllist.xls")

sh = book.sheet_by_index(0)

def get_stock():


    ##### get user to input a ticker that is in the smalllist.xls file

    ticker =input("enter ticker here: ")

             
    ### find the link for that Ticker
    not_found=0
    for row in range(sh.nrows):
        
        if sh.cell(row,0).value == ticker:
              not_found = 1
              ref      = (sh.cell(row,2).value) #cell [x,2] contains the ref_id, should make this a var though
              stockurl = (sh.cell(row,3).value)
              company  = (sh.cell(row,1).value)
              get_stock_price(ref,stockurl,company) 
    if not_found==0:
         print "Sorry but the stock ticker you entered has not been found"
                
def get_stock_price(ref,stockurl,company):            
        page = urllib.urlopen(stockurl).read() #read the google finance page source code into a variable called page
        page = unicode(page, 'utf-8') #in the VZ case an ad on the stockurl source code included the "Registered TM" symbol which caused an error - UnicodeDecodeError: 'ascii' codec can't decode byte 
        n=len(ref)
        start_stock = page.find(ref)       #beginning of stock ticker ...in google finance each stock has a unique ref nunber
        end_stock = page.find('</span>',start_stock) # search end of string containing stock value
        stock_price = page[start_stock+n:end_stock] #+1 to avoid the " ...look at the source code to see what I mean
        
       #print the output to a html file, a= append = add the data rather than overwirte the current data          
        
        sys.stdout=open("index.html","a") # will create a new file called index if not existing in the DIR where the .py is saved
        
        print str(company),str(stock_price) # str to reverse the unicode command earlier,otherwise a U'appears in the outpur
        from time import ctime    #current time
        print ctime()
        print '\n'
        sys.stdout.close()   #needed or else the system holds the file and you get sharing violation when you try to manually close it
        webbrowser.open("file:///C:/myscripts/finished code/index.html")          
         
get_stock()          

         

           
