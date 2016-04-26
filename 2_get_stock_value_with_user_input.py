import webbrowser  # disp;ay the web page
import urllib         # read the web page to a variable
import sys      # send the ouput from the web page to a text or html file
import time

def get_stock(stockurl,ref,l):
        page = urllib.urlopen(stockurl).read() #read the google finance page source code into a variable called page
        start_stock = page.find(ref)       #beginning of stock ticker ...in google finance each stock has a unique ref nunber
        end_stock = page.find('</span>',start_stock) # search end of string containing stock value
        stock_price = page[start_stock+l:end_stock] #+1 to avoid the " ...look at the source code to see what I mean
        return (stock_price) # program returns the desired stock price
       

def get_all_stocks ():
                     
        link = ["https://www.google.com/finance?q=t&ei=vQwNV7DvKdXCU4CBv9AB","https://www.google.com/finance?q=vz&ei=bxINV9nxFtXCU4CBv9AB"]
        business = ["AT&T", "VZ"]
        source_ref_id = ['ref_33312_l">', '_664887_l">']

        user_input_link = input("Enter the Google Finance link\t ") 
        link.append(user_input_link)
        user_input_business = input("Enter the company name\t: ") 
        business.append(user_input_business)
        user_input_source_ref_id = input ("Enter the ref_id: \t")
        source_ref_id.append(user_input_source_ref_id)
      
        i=0
        n=len(business)
        print n
        while i < n:
            stockurl = link[i]
            ref=source_ref_id[i]
            l=len(source_ref_id[i]) # this account for the fact that the source ref ids can have different lengths  eg 33313 is shorter than 66487
            stock_price = get_stock(stockurl, ref,l)
            company=business[i]
            #print the output to a html file, a= append = add the data rather than overwirte the current data          
            sys.stdout=open("index.html","a") # will create a new file called index if not existing in the DIR where the .py is saved
            print (company,stock_price)
            from time import ctime    #current time
            print ctime()
            print '\n'
            sys.stdout.close()   #needed or else the system holds the file and you get sharing violation when you try to manually close it
            i=i+1
         
            
get_all_stocks()

webbrowser.open("file:///C:/myscripts/finished%20code/index.html")




