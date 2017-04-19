from bs4 import BeautifulSoup
import urllib2, re
import xlwt
import requests


rankings = "https://colleges.niche.com/stanford-university/rankings/"
ran_page = urllib2.urlopen(rankings);
ran_soup = BeautifulSoup(ran_page.read(),"lxml")

Tabbled_content = ran_soup.find("div",{"class":re.compile("tabbed__content")})
# Code to get the different categories

for tab in Tabbled_content.find_all("div",{"class":re.compile("tabbed__content__body")}):
	ol_bucket = tab.find("ol",{"class":re.compile("ordered__list__bucket")})
	for element in ol_bucket.find_all("li",{"class":re.compile("ordered__list__bucket__item")}):
		ranking_card = element.find("div",{"class":re.compile("rankings-card")})
		#title
		ranking_card_title =  ranking_card.find("div",{"class":re.compile("rankings-card__link__title")})
		print ranking_card_title.text
		
		
		ranking_card_rank = ranking_card.find("div",{"class":re.compile("rankings-card__link__rank")})
		print ranking_card_rank.text
		
