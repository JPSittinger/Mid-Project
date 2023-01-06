import os
import re
import openpyxl
import datetime
import time
import unidecode
import googlemaps
import math
import shelve
from bs4 import BeautifulSoup as soup
from urllib.request import urlopen as uReq

# Definition of the output excel file
today = str(datetime.date.today())
os.chdir(r"C:\Users\a\Google Drive\IRONHACK Data Analysis Bootcamp\Mid-Project\Webscraping")
output_file_path = r'C:\Users\a\Google Drive\IRONHACK Data Analysis Bootcamp\Mid-Project\Webscraping\results.xlsx'
output_file = openpyxl.load_workbook(output_file_path)
output_file.create_sheet(index = 0, title = today + '_h')
output_sheet = output_file[today + '_h']
output_sheet['A1'].value = 'Link'
output_sheet['B1'].value = 'RentPrice'
output_sheet['C1'].value = 'Surface'
output_sheet['D1'].value = 'Bedrooms'
output_sheet['E1'].value = 'Bathrooms'
output_sheet['F1'].value = 'LastUpdate'
output_sheet['G1'].value = 'Neighborhood'
output_sheet['H1'].value = 'Location'
output_sheet['I1'].value = 'Seller'
output_sheet['J1'].value = 'PropertySubtype'
output_sheet['K1'].value = 'Description'
output_sheet['L1'].value = 'Latitude'
output_sheet['M1'].value = 'Longitude'
output_sheet['N1'].value = 'DistanceToCenter'
output_sheet['O1'].value = 'Floor'

# Regular Expressions
link_pattern = re.compile(r'href="((\S)*)"') # Defines the pattern for the link of each result
price_pattern = re.compile(r'strong>((\d+.)?\d+) €') # Defines pattern for the rent
surface_pattern = re.compile(r'<strong>(\d+)') # Defines pattern for the surface
bedrooms_pattern = re.compile(r'(\d+)</strong> hab.') # Defines pattern for num of bedrooms
bathrooms_pattern = re.compile(r'(\d+)</strong> baño') # Defines pattern for num of bathrooms
neighborhood_pattern = re.compile(r'''\n([\w\W]*) \r\n\s*(([\w\W]*)\r\n)?''') # Defines pattern for the neighborhood and location
last_update_pattern = re.compile(r'datetime="([\d\/]*)"') # Defines pattern for the last update date
description_pattern = re.compile(r'''description">([\w\s\S,'`´.<>/+-]*)</p>''') # Defines pattern for the text description
floor_pattern = re.compile(r'Planta número (\d)') # Defines pattern for the floor

# Geocoding preparation
maps_apikey = 'your-google-maps-api'
gmaps = googlemaps.Client(key = maps_apikey)
def distance(lat1, lng1, lat2, lng2):
    if (lat1 == lat2) and (lng1 == lng2):
        return 0
    else:
        theta = lng1-lng2
        dist = math.sin(math.radians(lat1)) * math.sin(math.radians(lat2)) + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) * math.cos(math.radians(theta))
        dist = math.acos(dist)
        dist = math.degrees(dist)
        kilometers = dist * 60 * 1.1515 * 1.609344;
        return kilometers
    
backup_locations = shelve.open('backup_locations')

# Create an empty dictionary
loc_dict = {}

# Update the shelve file with the 'locDict' key
backup_locations.update({'locDict': loc_dict})

backup_locations_dict = backup_locations['locDict']

# Start of the functional code
max_page = input("How many pages of results do you want to analyse? ") # Attention: do not surpass the number of actual pages. Check: https://www.habitaclia.com/alquiler-barcelona.htm?ordenar=mas_recientes
max_page = int(max_page)
checked_ads = 0

start = time.time()

for i in range(0,max_page):
    # Sets the result page link
    if i != 0:
        add = str(-i)
        link = "https://www.habitaclia.com/alquiler-barcelona" + add + ".htm?ordenar=mas_recientes"
    else:
        link = "https://www.habitaclia.com/alquiler-barcelona.htm?ordenar=mas_recientes"
    # uClient downloads the result page, then reads and stores it
    while True:
        try:
            uClient = uReq(link)
            resultpage_raw = uClient.read()
            uClient.close()
            break
        except:
            print('\nSomething failed. Possibly Error 403: Forbidden Access. Trying again in 30 seconds.')
            time.sleep(30)

    # soup parses the page
    resultpage_clean = soup(resultpage_raw, "html.parser")
    
    # We find all the links to the acommodations in the result page and loop over them
    for j in resultpage_clean.findAll("h3", {"class": "list-item-title"}):
        link = link_pattern.search(str(j))
        while True:
            try:
                uClient = uReq(link.group(1))
                adpage_raw = uClient.read()
                uClient.close()
                break
            except:
                print('\nSomething failed. Possibly Error 403: Forbidden Access. Trying again in 30 seconds.')
                time.sleep(30)
                
        adpage_clean = soup(adpage_raw, "html.parser")
        try:
            features = adpage_clean.findAll("li", {"class": "feature"})
            features_neigh = unidecode.unidecode(str(features[-1]))
            neigh_loc = neighborhood_pattern.search(features_neigh)
            try:
                neighborhood = neigh_loc.group(1)[0:-1].strip()
            except:
                neighborhood = None
            try:
                location = neigh_loc.group(3)
            except:
                location = None
            
            features = str(features)
            price = price_pattern.search(features).group(1).replace('.', '')
            surface = surface_pattern.search(features).group(1)
            try:
                bedrooms = bedrooms_pattern.search(features).group(1)
            except:
                bedrooms = None
            try:
                bathrooms = bathrooms_pattern.search(features).group(1)
            except:
                bathrooms = None
            try:
                description = adpage_clean.findAll("p", {"class": "detail-description"})
                description = description_pattern.search(str(description)).group(1).replace('<br/>', '\n')
            except:
                description = None

            offer_info = adpage_clean.find("main", {"class": "w-100 curve-top pointer-events-none gtmproductdetail"})
            seller = offer_info.get('data-esparticular')
            prop_subtype = offer_info.get('data-propertysubtype')
        
            last_update = adpage_clean.findAll("p", {"class": "time-tag"})
            last_update = last_update_pattern.search(str(last_update)).group(1)

            floorResults = adpage_clean.findAll('article', {'class': 'has-aside'})
            for k in floorResults:
                if floor_pattern.search(str(k)) != None:
                    floor = floor_pattern.search(str(k)).group(1)
                    break
                else:
                    floor = None
                
            if location == None and neighborhood != None:
                locationString = neighborhood + ' Barcelona'
            elif location != None and neighborhood == None:
                locationString = location + ' Barcelona'
            else:
                locationString = neighborhood + ' ' + location + ' Barcelona'
                # Note that it will fail if both == None. And I agree: it is not worth it to have a non-geocodable record.
 
            geocode_result = gmaps.geocode(locationString)
            if locationString in backup_locations_dict.keys():
                latitude = backup_locations_dict[locationString]['lat']
                longitude = backup_locations_dict[locationString]['lng']
                dist = backup_locations_dict[locationString]['dist']
            else:
                geocode_result = gmaps.geocode(locationString)
                latitude = geocode_result[0]['geometry']['location']['lat']
                longitude = geocode_result[0]['geometry']['location']['lng']
                dist = str(distance(41.382542, 2.177100, latitude, longitude))
                latitude, longitude = str(latitude), str(longitude)

                backup_locations_dict[locationString] = {'lat': latitude,
                                                         'lng': longitude,
                                                         'dist': dist}
                
            # Now the data is added to the excel file
            row_excel = str(checked_ads + 2)
            output_sheet['A' + row_excel].value, output_sheet['B' + row_excel].value, \
                             output_sheet['C' + row_excel].value, output_sheet['D' + row_excel].value, \
                             output_sheet['E' + row_excel].value, output_sheet['F' + row_excel].value, \
                             output_sheet['G' + row_excel].value, output_sheet['H' + row_excel].value, \
                             output_sheet['I' + row_excel].value, output_sheet['J' + row_excel].value, \
                             output_sheet['K' + row_excel].value, output_sheet['L' + row_excel].value, \
                             output_sheet['M' + row_excel].value, output_sheet['N' + row_excel].value, \
                             output_sheet['O' + row_excel].value = \
                             link.group(1), price, surface, bedrooms, bathrooms, last_update, \
                             neighborhood, location, seller, prop_subtype, description, latitude, \
                             longitude, dist, floor

            checked_ads += 1
        except Exception as e:
            print(e)
            print('Scraping failed, link: ' + link.group(1) + '\n')
        

    print(str(i+1) + ' result page(s) analysed, out of ' + str(max_page) + '\n')

backup_locations['locDict'] = backup_locations_dict
backup_locations.close()
output_file.save(output_file_path)

end = time.time()
print('Minutes elapsed: ' + str((end - start)/60))