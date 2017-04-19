from bs4 import BeautifulSoup
import urllib2, re
import xlwt
import requests
from xlrd import open_workbook

# Open the excel sheet to write the data
book = xlwt.Workbook(encoding="utf-8")
sheet1 = book.add_sheet('rankings')

i = 0

sheet1.write(i,0,"SR Number")
sheet1.write(i,1,"College Name")
sheet1.write(i,2,"Overall Niche Grade")
sheet1.write(i,3,"Academics")
sheet1.write(i,4,"Value")
sheet1.write(i,5,"Diversity")
sheet1.write(i,6,"Campus")
sheet1.write(i,7,"Athletics")
sheet1.write(i,8,"Party Scene")
sheet1.write(i,9,"Professors")
sheet1.write(i,10,"Location")
sheet1.write(i,11,"Dorms")
sheet1.write(i,12,"Campus Food")
sheet1.write(i,13,"Student Life")
sheet1.write(i,14,"Safety")
sheet1.write(i,15,"College type")
sheet1.write(i,16,"Athletic Division")
sheet1.write(i,17,"Athletic Conference")
sheet1.write(i,18,"Address")
sheet1.write(i,19,"Website")

i = i+1


# 1 Grades


wb = open_workbook('test.xls')
for sheet in wb.sheets():
	number_of_rows = sheet.nrows
	number_of_columns = sheet.ncols
	i = 1
	for row in range(1 , number_of_rows):
		cname = (sheet.cell(row,1).value)
		link = (sheet.cell(row,2).value)
		filename = name.replace(" ", "")
		print name , link , filename	
		sheet1.write(i,0,i)
		sheet1.write(i,1,name)
		url= link
		page = urllib2.urlopen(url)
		soup = BeautifulSoup(page.read(),"lxml")
		

		bucket1 = soup.find("section",{"class":re.compile("block--report-card")})
		# for profile__overall__grade__label
		overall_grade = bucket1.find("div",{"class":re.compile("niche__grade niche__grade")}).string
		sheet1.write(i,2,overall_grade)
		j = 3
		bucket2 = bucket1.find("div",{"class":re.compile("profile__bucket--2")})
		for link in bucket2.find_all("li",{"class":re.compile("ordered__list__bucket__item")}):
			values = link.find("div",{"class":re.compile("niche__grade niche__grade--section")}).string
			sheet1.write(i,j,values)
			j = j+1
		
		general = soup.find("section",{"class":re.compile("block--two")})

		# School type, Athlectic division , Athletic conference

		bucket1 = general.find("div",{"class":re.compile("profile__bucket--1")})
		scalars = bucket1.find_all("div",{"class":re.compile("scalar--two")})
		for scalar in scalars:
			try:
				values = scalar.find("div",{"class":re.compile("scalar__value")}).find('span').string
			except:
				values = "No Data Available"
			sheet1.write(i,j,values)
			j = j+1

		# school address , website

		bucket2 = general.find("div",{"class":re.compile("profile__bucket--2")})

		address = bucket2.find("div",{"class":re.compile("profile__address")})
		Address_value = address.text
		values = Address_value[7:]
		sheet1.write(i,j,values)
		j = j+1

		website = bucket2.find("div",{"class":re.compile("profile__website")})
		try: 
			values = website.find("div",{"class":re.compile("profile__website__url")}).text
		except:
			values = "Data not available"
		sheet1.write(i,j,values)
		j = j+1
		i = i+1


#Save the output    
book.save('Rankings.xls')

