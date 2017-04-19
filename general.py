from bs4 import BeautifulSoup
import urllib2, re
import xlwt
import requests


url= 'https://colleges.niche.com/stanford-university/'
page = urllib2.urlopen(url)
soup = BeautifulSoup(page.read(),"lxml")

general = soup.find("section",{"class":re.compile("block--two")})

# School type, Athlectic division , Athletic conference

bucket1 = general.find("div",{"class":re.compile("profile__bucket--1")})
scalars = bucket1.find_all("div",{"class":re.compile("scalar--two")})
for scalar in scalars:
	print scalar.find("div",{"class":re.compile("scalar__value")}).find('span').string

# school address , website

bucket2 = general.find("div",{"class":re.compile("profile__bucket--2")})

address = bucket2.find("div",{"class":re.compile("profile__address")})
Address_value = address.text
print Address_value[7:]

website = bucket2.find("div",{"class":re.compile("profile__website")})
print website.find("div",{"class":re.compile("profile__website__url")}).text

