from bs4 import BeautifulSoup as soup
from datetime import datetime
import os, ssl
if (not os.environ.get('PYTHONHTTPSVERIFY', '') and
    getattr(ssl, '_create_unverified_context', None)):
    ssl._create_default_https_context = ssl._create_unverified_context
import urllib.request
import xlsxwriter

#Welcome Message
print ("Welcome to the Online Pricing Tool !")
print ("Please wait while I gather your data, fresh off the web.")
print()

##------------------------------------------------ NETWORK SETTINGS -----------------------------------------------------

# Sony Proxy Support - Un-comment when inside Sony network. 
proxy_support = urllib.request.ProxyHandler({'http' : '43.66.8.18:8080',
                                             'https': '43.66.8.18:8080'})
opener = urllib.request.build_opener(proxy_support)
urllib.request.install_opener(opener)

# Pretending to be a browser.
def passport (page_url):        
    req = urllib.request.Request(page_url, data=None, headers={
          'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_9_3) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/35.0.1916.47 Safari/537.36'
          })
    return req

##--------------------------------------------CREATING EXCEL SPREADSEHEET ------------------------------------------------

#Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('Online_TV_Pricing.xlsx')
worksheetNL = workbook.add_worksheet('Noel Leeming')
worksheetJB = workbook.add_worksheet('JB Hifi')
worksheetHN = workbook.add_worksheet('Harvey Norman')

#Setup worksheet formating here
bold = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': 'black'})
brands = 3 # number of brands (zero indexed)
feats = 4 # number of rows for features 

#Create function to set column widths (zero indexed)
def formating (sheet_brand):
    C = 0
    B = brands * feats
    while C <= B:    #Repeat for each brand
        sheet_brand.set_column(C,    C,    15)  #Model
        sheet_brand.set_column(C +1, C +1, 6)   #Size
        sheet_brand.set_column(C +2, C +2, 10)  #Price
        sheet_brand.set_column(C +3, C +3, 3)   #Gap
        C += feats

#Create funtion to set up headers (zero indexed)
def headers (head_brand):
    R = 0
    C = 0
    B = brands * feats
    while C <= B:    #Repeat for each brand
        head_brand.write(R,    C,    ' ',     bold)  #Top Space
        head_brand.write(R,    C +1, ' ',     bold)  #Top space       
        head_brand.write(R,    C +2, ' ',     bold)  #Top space       
        head_brand.write(R +1, C,    'Model', bold)  #Model
        head_brand.write(R +1, C +1, 'Size',  bold)  #Size
        head_brand.write(R +1, C +2, 'Price', bold)  #Price
        C += feats
        
#Setup Spreedsheet formating with above functions
formating (worksheetNL)
formating (worksheetJB)
formating (worksheetHN)
headers (worksheetNL)
headers (worksheetJB)
headers (worksheetHN)

##--------------------------------------NOEL LEEMING DATA COLLECTION ------------------------------------------------------

#Reset Row and Column counters
row = 2 # start under header
col = 0

#Print update message
print("Gathering NOEL LEEMING Data...")
print()

# List of all pages to scrape
def headphone (brand, X):
    part1 = "https://www.noelleeming.co.nz/shop/Sony/televisions/televisions/c10137-cTVs-b"
    part2 = "-p"
    part3 = ".html?sorter=price-desc"
    url = part1 + brand + part2 + X + part3    
    return url

#Generate Sony URL's
sony1 = headphone("sony","1")
sony2 = headphone("sony","2")
samsung1 = headphone("samsung","1")
samsung2 = headphone("samsung","2")
samsung3 = headphone("samsung","3")
panasonic1 = headphone("panasonic","1")
panasonic2 = headphone("panasonic","2")
panasonic3 = headphone("panasonic","3")
lg = headphone("lg","1")

# List pages by brand
sony_list = [sony1, sony2]
samsung_list = [samsung1, samsung2, samsung3]
panasonic_list = [panasonic1, panasonic2, panasonic3]
lg_list = [lg]

# List of Lists of Brands
url_list = [sony_list, samsung_list, panasonic_list, lg_list]

# Loop through each page and find product boxes
for url in url_list:
    # Download page  
    for page in url:
        req = passport(page)
        f = urllib.request.urlopen(req)
        page = (f.read().decode('utf-8'))

    # Making the page into Beautiful Soup
        pagesoup = soup(page, "html.parser")
        model_list = pagesoup.findAll("li", {"class": 'block product-list'})

    # Start of loop to find the data we want for each product box
        for model_box in model_list:
            # Find the model name
            model_strip = model_box.find("h2", {"class": 'product-list__model'})
            model = model_strip.text.strip()
            # Find the price
            price_list = model_box.find("span", {"class": 'price-lockup__pricing-fullprice-wrap'})
            pricestrip = price_list.text.strip()
            price = ''.join(pricestrip.split())
            #finding the make and size
            make_list = model_box.find("h1", {"class": 'product-list__name'})
            make_strip = make_list.text.split()
            make = make_strip[0]
            size = make_strip[1]
            #Find offer timing
            offer_full = model_box.find("div", {"class": 'allcaps offer__details mrs mbs'})
            offer_split = offer_full.text.strip()
            offer_list = offer_split.splitlines()

            if offer_list == [] :
               offer = "Standard Price"
            else:
               offer = offer_list[0]
        
            #Writing Data to Excel file.
            worksheetNL.write(row, col,     model)
            worksheetNL.write(row, col + 1, size)
            worksheetNL.write(row, col + 2, price)
            row += 1
    
    worksheetNL.write(0, col, make, bold)
    
    col += 4    #move over 4 columns
    row = 2     #reset to top data row (2)  


##--------------------------------------JB HIFI DATA COLLECTION ------------------------------------------------------

#Reset Row and Column counters
row = 2 # start under header
col = 0

#Print update message
print("Gathering JB HIFI Data...")
print()

# Function to generate JBHIFI Headphone URL's
def headphone (brand, X):
    part1 = "https://www.jbhifi.co.nz/tvs/all-tvs/?p="
    part2 = "&s=displayPrice&sd=2&fc=brand%3A%3A"
    part3 = "%3B&mf=brand&fm=false"
    url = part1 + X + part2 + brand + part3    
    return url

#Generate Sony URL's
sony = headphone("SONY","1")
samsung = headphone("SAMSUNG", "1")
panasonic = headphone("PANASONIC", "1")
lg = headphone("LG", "1")

page_list = [sony, samsung, panasonic, lg]

for page in page_list:
  
    # Downloads the page listed above for processing.
    req = passport(page)
    f = urllib.request.urlopen(req)
    page = (f.read().decode('utf-8'))

    # Making the page into Beautiful Soup
    pagesoup = soup(page, "html.parser")
    model_list_promo = pagesoup.findAll("div",{"class":"span03 product-tile has-feature"})
    
    # Taking model name from site naming it 'model' ON PROMO
    for model_box in model_list_promo: 
        #Finding model
        model_full = model_box["title"]
        modelbits = model_full.split()
        make = modelbits[0]
        model = modelbits[1]
        size = modelbits[2]
        
        #Finding price
        price_list = model_box.find("span", {"class":"offer cashback"})
        if price_list is None:
            price_list = model_box.find("span", {"class": "amount"})
            price_strip = price_list.text.strip()
            price = ''.join(price_strip.split())
        else:
            price_strip = price_list.text.strip()
            price_cb = ''.join(price_strip.split())
            price = price_cb + " CB"

        #Writing Data to Excel file.
        worksheetJB.write(row, col,     model)
        worksheetJB.write(row, col + 1, size)
        worksheetJB.write(row, col + 2, price)
        row += 1

    # Taking model name from site naming it 'model'
    model_list = pagesoup.findAll("div",{"class":"span03 product-tile"})
    for model_box in model_list: 
        
        #Finding model
        model_full = model_box["title"]
        modelbits = model_full.split()
        make = modelbits[0]
        model = modelbits[1]
        size = modelbits[2]
        
        # Finding price
        price_list = model_box.find("span", {"class":"offer cashback"})
        if price_list is None:
            price_list = model_box.find("span", {"class": "amount"})
            price_strip = price_list.text.strip()
            price = ''.join(price_strip.split())
        else:
            price_strip = price_list.text.strip()
            price_cb = ''.join(price_strip.split())
            price = price_cb + " CB"
 
        # Writing Data to Excel file.
        worksheetJB.write(row, col,     model)
        worksheetJB.write(row, col + 1, size)
        worksheetJB.write(row, col + 2, price)
        row += 1
        
        #Write brand make in header            
        if row <= 8:
            # Write brand at the top
            worksheetJB.write(0, col, make, bold)
        
    # Clean up JB Page Final Record
    row -= 1
    worksheetJB.write(row, col,     " ")
    worksheetJB.write(row, col + 1, " ")
    worksheetJB.write(row, col + 2, " ")    
    
    # Switch to next Brand
    col += 4    #move over 4 columns
    row = 2     #reset to top data row (2)

##-------------------------------------- HARVEY NORMAN DATA COLLECTION -----------------------------------------------

row = 2 # start under header
col = 0

#Print update message
print("Gathering HARVEY NORMAN Data...")
print()

# Function to generate HARVEY NORMAN Headphone URL's
def headphone (brand):
    part1 = "https://www.harveynorman.co.nz/tv-and-audio/televisions/?subcats=Y&features_hash="
    part2 = "&sort_by=price&sort_order=desc&layout=products_without_options"
    url = part1 + brand + part2  
    return url

#Generate URLS for each brand
sony = headphone("V28")
samsung = headphone("V26")
panasonic = headphone("V25")
lg = headphone("V24")

page_list =[sony, samsung, panasonic, lg]

for page in page_list:

    # Download page 
    req = passport(page)
    f = urllib.request.urlopen(req)
    page = (f.read().decode('utf-8'))

    # Making the page into Beautiful Soup
    pagesoup = soup(page, "html.parser")
    # Splitting site into a list of models naming it 'model_list'
    model_list = pagesoup.find_all("div", {"class": 'product-row'})

    # Starting Loop to remove model number and price from each model in the list.
    for model_box in model_list: 

        #Finding model
        make_list = model_box.find("a", {"class": 'product-title'})
        make_strip = make_list.text.split()
        make = make_strip[0]
        feat1 = make_strip[2]
        feat2 = make_strip[3]
        size = make_strip[1]
        model = (feat1 + feat2)
        
        #Finding price
        price_list = model_box.find("div", {"class": 'price-group-wrap'})
        pricesplit = price_list.text.split()
        price = pricesplit[0]

        #Writing Data to Excel file.
        worksheetHN.write(row, col,     model)
        worksheetHN.write(row, col + 1, size)
        worksheetHN.write(row, col + 2, price)
        row += 1
    
    worksheetHN.write(0, col, make, bold)
    col += 4    #move over 4 columns
    row = 2     #reset to top data row (2)

##----------------------------------------  CLOSE APPLICATION   ------------------------------------------------------
#Close and Save Excel file
print ("Creating Excel File...")
print ()
workbook.close()

#Print closing message
print ("File Created Succsessfully !")
print ()
print ()
input ("Please press Enter to close.")