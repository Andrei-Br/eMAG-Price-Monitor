from bs4 import BeautifulSoup
import requests

def CheckPrice(link):
    #  link to the object we want to get the price
    url = requests.get(link).text

    soup = BeautifulSoup(url, 'lxml')

    title = soup.find('h1', class_='page-title').text.strip()
    format_Title = title.split(',')

    #  formating the oldPrice (string) to a float
    oldPrice = soup.find(class_='product-old-price').text.strip()
    converted_oldPrice = float(oldPrice[0:5])

    #  formating the newPrice (string) to a float
    newPrice = soup.find(class_='product-new-price').text.strip()
    converted_newPrice = float(newPrice[0:6])

    #  store name
    store = soup.find('span','a', class_ ='text-label').next_sibling.strip()

    #  DICT for returning multiple values to write in EXCEL"""
    d = dict()
    d['title'] = format_Title[0]
    d['oldPrice'] = converted_oldPrice
    d['newPrice'] = converted_newPrice
    d['store'] = store

    return d

