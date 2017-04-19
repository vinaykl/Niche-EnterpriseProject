from bs4 import BeautifulSoup
import urllib2, re
import xlwt
import requests
import time

from selenium import webdriver
from selenium.webdriver.common.keys import Keys

from xlrd import open_workbook


wb = open_workbook('test.xls')
for sheet in wb.sheets():
	number_of_rows = sheet.nrows
	number_of_columns = sheet.ncols
	for row in range(1 , number_of_rows):
		name = (sheet.cell(row,1).value)
		link = (sheet.cell(row,2).value)
		filename = name.replace(" ", "")
		print  row, name , link , filename	
		
		

		url= link +'reviews/'
		opener = urllib2.build_opener(urllib2.HTTPRedirectHandler)
    		request = opener.open(url)
    		if request.url != url:
			continue


		page = urllib2.urlopen(url)
		soup = BeautifulSoup(page.read(),"lxml")
		flag =0
		try:
			pages = int(soup.select("select.pagination__pages__selector option")[-1].text.split(None, 1)[1])
		except:
			pages =1
			flag = 1

		#filename =  url.split('/')[3] 


		# Open the excel sheet to write the data
		book = xlwt.Workbook(encoding="utf-8")
		sheet1 = book.add_sheet('vkl')
		sheet1.write(0,0,"Serial Number")
		sheet1.write(0,1,"Rating")
		sheet1.write(0,2,"Review")
		sheet1.write(0,3,"User-Type")
		sheet1.write(0,4,"Review-time")
		i=1

		driver = webdriver.Firefox()
		driver.get(url)


		for page in range(1, pages+ 1):
			#print page
			# integrating BeautifulSoup in

			#reviews_page = urllib2.urlopen(reviews)
			reviews_soup = BeautifulSoup(driver.page_source,"lxml")

			reviews_exp = reviews_soup.find("section",{"class":re.compile("reviews-expansion-bucket")})


			reviews_list = reviews_exp.children
			for rev in reviews_list:
				if rev.name == 'div':
					sheet1.write(i,0,i)
					# from rev get all the info I need
					# get the stars given by the user
					review_star = rev.find("div",{"class":re.compile("review__stars")})
					#print review_star
					review_star_children =  review_star.find('span')
					star = review_star_children.get('class',None)
					#print star
					if str(50) in str(star):
						sheet1.write(i,1,"5 star")
						#print "5 Star"
					elif str(40) in str(star):
						sheet1.write(i,1,"4 star")
						#print "4 Star"
					elif str(30) in str(star):
						sheet1.write(i,1,"3 star")
						#print "3 Star"
					elif str(20) in str(star):
						sheet1.write(i,1,"2 star")
						#print "2 Star"
					else:
						sheet1.write(i,1,"1 star")
						#print "1 Star"

					# get the actual Test review
					review_text = rev.find("span",{"class":re.compile("review__text")})
					sheet1.write(i,2,review_text.text)
					#print review_text.text
		
					# get the year of the student
					# get the time of the review. Each child has one of the information.
					review_third_section = rev.find("ul",{"class":re.compile("review__tagline")})
					#print review_third_section
					childrens = review_third_section.children
					j=3
					for child in childrens:
						sheet1.write(i,j,child.string)
						j = j+1
						#print child.string
					#print '\n'
					i=i+1
			if flag != 1:
				mores = driver.find_element_by_class_name('icon-arrowright-thin--pagination')
				mores.click()
				time.sleep(1)
	

	

		driver.quit()

		#Save the output 
		filename = filename + ".xls"  
		book.save(filename)

