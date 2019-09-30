import logging
from bs4 import *
from openpyxl import *
import requests
from requests.exceptions import SSLError

log = logging.getLogger(__name__)
designations = ["President, CEO, Director, Founder"]

def get_person(site_result):
    print(site_result.headers)
    if dict(site_result.headers)['Content-Type'] == "text/html":
        soup = BeautifulSoup(site_result.text, 'html.parser')
        print(soup.pretty())
        for i in soup.find_all('a'):
            print(i)
            if i.value == 'Our Team' or i.value == 'Team' or i.value == 'About':
                print(i.href)

def check_url(company_url):
    if company_url[:3] != "www":
        company_url = "www." + company_url
    # print(company_url)
    try:
        request_result = requests.get("https://" + company_url)
        if request_result.status_code == 200:
            print(company_url+" - "+str(request_result.status_code))
            get_person(request_result)
    except SSLError:
        try:
            request_result = requests.get("http://" + company_url)
            if request_result.status_code == 200:
                print(company_url+" - "+str(request_result.status_code))
                get_person(request_result)
        except:
            print(company_url + "not found")


def start():
    wb = load_workbook(filename='P2PContactScraping.xlsx')
    sheet_ranges = wb["Sheet1"]
    for i in range(16200, 16202):
        check_url(sheet_ranges['C' + str(i)].value)


start()
