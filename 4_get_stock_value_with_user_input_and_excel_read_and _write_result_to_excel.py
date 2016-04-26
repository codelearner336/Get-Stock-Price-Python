

import xlrd       # read from an excel file
import urllib     # open a web page and copy it source code to a variable (which will then be searched)  
import time       # to get the current time to be added as a timestamp to the printout  
import sys        # to be able to output data to a text file  
import webbrowser # to open the text file created with the data, may remove when writing to an excel
import xlutils    # needed to write to/edit an EXISTING excel file



book = xlrd.open_workbook("smalllist.xls")

sh = book.sheet_by_index(0)

def get_stock():


##### get user to input  ticker values that may or may not be in the smalllist.xls file

    
     ticker =input("enter you ticker choices here, each seperated by a comma, hit ENTER when you have finished \t  ")
     print ticker
     z= len(ticker)
     

##### iterate through the tickers and check each one, allow for inputs that are not in the small list

    ### read the data for a Ticker from the excel 
     
     not_found = 0
     i=-1
     while i < z-1:
      i=i+1                                                      #either it is found, or it is not found , so this is the only place to increment i  
      for row in range(sh.nrows):
        
            if sh.cell(row,0).value == ticker[i]:
                  not_found = 1
                  company  = (sh.cell(row,1).value)
                  ref      = (sh.cell(row,2).value)              #cell [x,2] contains the ref_id, would like to make this a var though
                  stockurl = (sh.cell(row,3).value)
                  get_stock_price(ref,stockurl,company,ticker[i])# the stock price is obtained this executes before the following if/else statement
              
      if not_found==0:
         print "Sorry but the stock ticker you entered --" ,ticker[i]   , "--has not been found"
      else:
        print "The price of stock ticker",  ticker[i], "has been written into your excel file. Please check that file to see the result."
        not_found=0                                               # resetting needed here else value is always 1 for the rest of the iterations evern if a value is not found

##### next go and get the stock price for a particular valid Ticker
        
def get_stock_price(ref,stockurl,company,ticker):
        page = urllib.urlopen(stockurl).read()                    #read the google finance page source code into a variable called page
        page = unicode(page, 'utf-8') #in the VZ case an ad on the stockurl source code included the "Registered TM" symbol which caused an error - UnicodeDecodeError: 'ascii' codec can't decode byte 
        n=len(ref)                    # discovered that not all "ref" are the same length so need to be able to handle that with LEN  
        start_stock = page.find(ref)                              #beginning of stock ticker ...in google finance each stock has a unique ref nunber
        end_stock = page.find('</span>',start_stock)              # search end of string containing stock value
        stock_price = page[start_stock+n:end_stock]               
        
##### This section will write the output to the excel smalllist

        from xlutils.copy import copy
        from time import ctime                                      #current time - nice to add to the output
        
        store_copyof_excel = xlrd.open_workbook("smalllist.xls")    # file should initially be closed before this can execute
        workbook = copy(store_copyof_excel)                         #copy of small list xls is created;  when you save it, it  overwrites the original
        sheet_to_write = workbook.get_sheet(1)                      #sheet_to_write is where I will write the output; get_sheet gets the referenced sheet

        # X is the number of nrows, the number of nrows is always one more than the row numbers i.e 3 rows in total implies rows0/1/2
        # so I leverage that idea to write to the next row which will always be row x, despite all references I found in stackoverflow saying you cant write to an excel.
        
        dummy = (store_copyof_excel).sheet_by_index(1)
        x= dummy.nrows
        
        sheet_to_write.write (x,0,company)          # this will write the value of the variable company or the string "company" if you use inverted commas
        sheet_to_write.write (x,1,ticker)           # row x column 1
        sheet_to_write.write (x,2,ctime())
        sheet_to_write.write (x,3,stock_price)
        workbook.save('smalllist.xls')              # if you do not have this command none of this excel write code will bne seen as SAVE  overwrites the original)
        
         
get_stock()                                         # ta dah!

     
    
