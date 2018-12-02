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

## Sony Proxy Support - (Un-comment when inside Sony network)
proxy_support = urllib.request.ProxyHandler({'http' : '43.66.8.18:8080',
                                             'https': '43.66.8.18:8080'})
opener = urllib.request.build_opener(proxy_support)
urllib.request.install_opener(opener)


## Pretending to be a browser.
def passport (page_url):        
    req = urllib.request.Request(page_url, data=None, headers={
          'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_9_3) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/35.0.1916.47 Safari/537.36'
          })
    return req

##--------------------------------------------CREATING EXCEL SPREADSEHEET ------------------------------------------------

#Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('Online_Soundbar_Pricing.xlsx')
#Create a worksheet for each reatailer.
worksheetNL = workbook.add_worksheet('Noel Leeming')
worksheetJB = workbook.add_worksheet('JB Hifi')
worksheetHN = workbook.add_worksheet('Harvey Norman')

#Setup worksheet formating here
bold = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': 'black'})
brands = 4 # number of brands (zero indexed)
feats = 3
#Create function to set column widths (zero indexed)
def formating (sheet_brand):
    C = 0
    B = brands * feats
    while C <= B:    #Repeat for each brand
        sheet_brand.set_column(C,    C,    20)  #Model
        sheet_brand.set_column(C +1, C +1, 10)  #Price
        sheet_brand.set_column(C +2, C +2, 3)   #Gap
        C += feats
    #Setup Other
    sheet_brand.set_column(C,    C,    10)  #Make
    sheet_brand.set_column(C +1, C +1, 20)  #Model
    sheet_brand.set_column(C +2, C +2, 10)  #Price
    sheet_brand.set_column(C +3, C +3, 3)   #Gap

#Create funtion to set up headers (zero indexed)
def headers (head_brand):
    R = 0
    C = 0
    B = brands * feats
    while C <= B:    #Repeat for each brand
        head_brand.write(R,    C,    ' ',     bold)  #Top Space
        head_brand.write(R,    C +1, ' ',     bold)  #Top space       
        head_brand.write(R +1, C,    'Model', bold)  #Model
        head_brand.write(R +1, C +1, 'Price', bold)  #Price
        C += feats
    
    head_brand.write(R,    C,    ' ',     bold)  #Top Space
    head_brand.write(R,    C +1, ' ',     bold)  #Top space       
    head_brand.write(R,    C +2, ' ',     bold)  #Top space
    head_brand.write(R +1, C,    'Make',  bold)  #Make
    head_brand.write(R +1, C +1, 'Model', bold)  #Model      
    head_brand.write(R +1, C +2, 'Price', bold)  #Price
    
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

# Function to generate Noel Leeming URL's
def soundbar (brand, X):
    part1 = "https://www.noelleeming.co.nz/shop/audio/home-audio/sound-bars/cAudio-chomeaudio-cSoundbars-b"
    part2 = "-p"
    part3 = ".html?sorter=price-desc"
    url = part1 + brand + part2 + X + part3    
    return url

#Generate Sony URL's
sony = soundbar("sony","1")
samsung = soundbar("samsung","1")
panasonic = soundbar("panasonic","1")
lg = soundbar("lg","1")
sonos = soundbar("sonos","1")
bose = soundbar("bose","1")
orbit = soundbar("orbitsound","1")
platinum = soundbar("platinum","1")
denon = soundbar("denon","1")

# List pages listed by brand
sony_list = [sony]
samsung_list = [samsung]
panasonic_list = [panasonic]
lg_list = [lg]
sonos_list = [sonos]
other_list = [bose, orbit, platinum, denon]

# List of Lists of Brands
url_list = [sony_list, samsung_list, panasonic_list, lg_list, sonos_list, other_list]

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
            #Find offer timing
            offer_full = model_box.find("div", {"class": 'allcaps offer__details mrs mbs'})
            offer_split = offer_full.text.strip()
            offer_list = offer_split.splitlines()

            if offer_list == [] :
               offer = "Standard Price"
            else:
               offer = offer_list[0]
        
            #Writing Data to Excel file.
            if col <= 14:
                worksheetNL.write(row, col,     model)
                worksheetNL.write(row, col + 1, price)
                row += 1
            else:
                worksheetNL.write(row, col,     make)
                worksheetNL.write(row, col + 1, model)
                worksheetNL.write(row, col + 2, price)
                row += 1            
    
    worksheetNL.write(0, col, make, bold)
    
    col += 3    #move over 4 columns
    row = 2     #reset to top data row (2)  
worksheetNL.write(0, 15, "Other", bold)

##--------------------------------------JB HIFI DATA COLLECTION ------------------------------------------------------

#Reset Row and Column counters
row = 2 # start under header
col = 0

#Print update message
print("Gathering JB HIFI Data...")
print()

# Function to generate JBHIFI Headphone URL's
def soundbar (brand, X):
    part1 = "https://www.jbhifi.co.nz/headphones-speakers-audio/soundbars/?p="
    part2 = "&s=displayPrice&sd=2&fc=brand%3A%3A"
    part3 = "%3B&mf=brand&fm=false"
    url = part1 + X + part2 + brand + part3    
    return url

#Generate Sony URL's
sony = soundbar("SONY","1")
samsung = soundbar("SAMSUNG","1")
panasonic = soundbar("PANASONIC","1")
lg = soundbar("LG","1")
sonos = soundbar("SONOS","1")
bose = soundbar("BOSE","1")
jbl = soundbar("JBL","1")

# List pages listed by brand 
sony_list = [sony]
samsung_list = [samsung]
panasonic_list = [panasonic]
lg_list = [lg]
sonos_list = [sonos]
other_list = [bose, jbl]

# List of Lists of Brands
url_list = [sony_list, samsung_list, panasonic_list, lg_list, sonos_list, other_list]

# Loop through each page and find product boxes
for url in url_list:
    # Download page 
    for page in url:
      
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
            m1 = modelbits[1]
            m2 = modelbits[2]
            model = m1 + " " + m2
            
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
                if col <= 14:
                    worksheetJB.write(row, col,     model)
                    worksheetJB.write(row, col + 1, price)
                    row += 1
                else:
                    worksheetJB.write(row, col,     make)
                    worksheetJB.write(row, col + 1, model)
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
            if col <= 14:
                worksheetJB.write(row, col,     model)
                worksheetJB.write(row, col + 1, price)
                row += 1
            else:
                worksheetJB.write(row, col,     make)
                worksheetJB.write(row, col + 1, model)
                worksheetJB.write(row, col + 2, price)
                row += 1
              
            #Write brand make in header            
            if row <= 4:
                # Write brand at the top
                worksheetJB.write(0, col, make, bold)  
              
    # Clean up JB Page Final Record
        row -= 1
        worksheetJB.write(row, col,     " ")
        worksheetJB.write(row, col + 1, " ")
        worksheetJB.write(row, col + 2, " ")
        
    # Switch to next Brand
    col += 3    #move over 4 columns
    row = 2     #reset to top data row (2)
worksheetJB.write(0, 15, "Other", bold)

##-------------------------------------- HARVEY NORMAN DATA COLLECTION -----------------------------------------------

row = 2 # start under header
col = 0

#Print update message
print("Gathering HARVEY NORMAN Data...")
print()

# Function to generate HARVEY NORMAN Headphone URL's
def soundbar (brand):
    part1 = "https://www.harveynorman.co.nz/tv-and-audio/audio/blu-ray-and-home-theatre/sound-bars/?subcats=Y&features_hash="
    part2 = "&sort_by=price&sort_order=desc&layout=products_without_options"
    url = part1 + brand + part2  
    return url

#Generate URLS for each brand
sony = soundbar("V28")
samsung = soundbar("V26")
panasonic = soundbar("V25")
lg = soundbar("V24")
sonos = soundbar("V8830")
other = soundbar("V21.V1380.V1316.V1200.V1496")

sony_list = [sony]
samsung_list = [samsung]
pana_list = [panasonic]
lg_list = [lg]
sonos_list = [sonos]
other_list = [other]

# List of Lists of Brands
url_list = [sony_list, samsung_list, pana_list, lg_list, sonos_list, other_list]

#loop through each brand

for url in url_list:
    #loop through all pages for each brand
    for page in url:

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
            feat1 = make_strip[1]
            if len(make_strip) >= 4:
                feat2 = make_strip[3]
            else:
                feat2 = make_strip[2]
            size = make_strip[1]
            model = (feat1 + " " + feat2)
            
            #Finding price
            price_list = model_box.find("div", {"class": 'price-group-wrap'})
            pricesplit = price_list.text.split()
            price = pricesplit[0]

            #Writing Data to Excel file.
            if col <= 14:
                worksheetHN.write(row, col,     model)
                worksheetHN.write(row, col + 1, price)
                row += 1
            else:
                worksheetHN.write(row, col,     make)
                worksheetHN.write(row, col + 1, model)
                worksheetHN.write(row, col + 2, price)
                row += 1
        
        worksheetHN.write(0, col, make, bold)
        col += 3    #move over 4 columns
        row = 2     #reset to top data row (2)
        
worksheetHN.write(0, 15, "Other", bold)

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