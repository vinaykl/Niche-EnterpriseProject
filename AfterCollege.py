# Work on the after college part
from bs4 import BeautifulSoup
import urllib2, re
import xlwt
import requests

from xlrd import open_workbook

# Open the excel sheet to write the data
book = xlwt.Workbook(encoding="utf-8")
sheet1 = book.add_sheet('Campus')

i = 0

sheet1.write(i,0,"SR Number")
sheet1.write(i,1,"College Name")
sheet1.write(i,2,"Overall value")
sheet1.write(i,3,"of students feel like they are getting their money's worth out of their program.")
sheet1.write(i,4,"Graduation Rate")
sheet1.write(i,5,"Median Earnings 2 Years After Graduation")
sheet1.write(i,6,"Median Earnings 6 Years After Graduation")
sheet1.write(i,7,"Employed 2 Years After Graduation")
sheet1.write(i,8,"Employed 6 Years After Graduation")
sheet1.write(i,9,"Average Loan Amount")
sheet1.write(i,10,"Loan Default Rate")


wb = open_workbook('test.xls')
for sheet in wb.sheets():
	number_of_rows = sheet.nrows
	number_of_columns = sheet.ncols
	i = 1
	for row in range(1 , number_of_rows):
		name = (sheet.cell(row,1).value)
		link = (sheet.cell(row,2).value)
		filename = name.replace(" ", "")
		print i, name , link , filename	

		sheet1.write(i,0,i)
		sheet1.write(i,1,name)
		j =2

		url = link + "after-college/"
		opener = urllib2.build_opener(urllib2.HTTPRedirectHandler)
    		request = opener.open(url)
    		if request.url != url:
			i = i +1
			continue
		page = urllib2.urlopen(url)
		soup = BeautifulSoup(page.read(),"lxml")

		# 4 sections are there in general



		# section 1 -  Overall Value of the colege
		# overall Value , Poll - worth money , Graduation rate
		# use block--one-two

		Value = soup.find("section",{"class":re.compile("block--one-two")})
		# Overall Value-grade
		overall_value = Value.find("div",{"class":re.compile("profile__bucket--1")})
		try:
			value = overall_value.find("div",{"class":re.compile("niche__grade")}).text
		except:
			value = "Data not available"
		sheet1.write(i,j,value)
		j = j+1

		# WorthMoney-poll
		Money_Worth =  Value.find("div",{"class":re.compile("profile__bucket--2")})
		try:
			value = Money_Worth.find("div",{"class":re.compile("poll__single__percent__label")}).text
		except:
			value = "Data not available"
		sheet1.write(i,j,value)
		j = j+1

	
		# Graduation Rate
		Graduation_rate = Value.find("div",{"class":re.compile("profile__bucket--3")})
		try:
			value = Graduation_rate.find("div",{"class":re.compile("scalar__value")}).find('span').string
		except:
			value = "Data not available"
		sheet1.write(i,j,value)
		j = j+1		

		#-------------------------
		# section 2 - Earnings
		# Median earnings after 2 years . after 6 years
		# use block--two-two

		Earnings = soup.find("section",{"class":re.compile("block--two-two")})
		# median Salary after 2 years
		Earnings_2years = Earnings.find("div",{"class":re.compile("profile__bucket--1")})
		try:
			value = Earnings_2years.find("div",{"class":re.compile("scalar__value")}).find('span').string
		except:
			value = "Data not available"
		sheet1.write(i,j,value)
		j = j+1

		# Median salary after 6 years
		Earnings_6years = Earnings.find("div",{"class":re.compile("profile__bucket--2")})
		try:
			value = Earnings_6years.find("div",{"class":re.compile("scalar__value")}).find('span').string
		except:
			value = "Data not available"
		sheet1.write(i,j,value)
		j = j+1



		# section 3 - Job Placement
		# Employment after 2 years , Employment after 6 years
		# use block--two-poll

		blocks = soup.find_all("section",{"class":re.compile("block--two")})


		# Employment after 2 years and 6 years
		blank__bucket = blocks[1].find("div",{"class":re.compile("profile__bucket--1")}).find("div",{"class":re.compile("blank__bucket")})
		childrens = blank__bucket.children
		# child 1 = 2 years child 2 = 6 years
		for child in childrens:
			try:
				value = child.find("div",{"class":re.compile("scalar__value")}).find('span').string
			except:
				value = "Data not available"
			sheet1.write(i,j,value)
			j = j+1


		# section 4 - Student Debt
		# Average loan amount , loan default rate
		# use block--expansion-back
 

		#print blocks[1]

		# Average loan Amount
		Average_loan_amount = blocks[2].find("div",{"class":re.compile("profile__bucket--1")})
		try:
			value = Average_loan_amount.find("div",{"class":re.compile("scalar__value")}).find('span').string
		except:
			value = "Data not available"
		sheet1.write(i,j,value)
		j = j+1

		# Loan Default Rate
		Loan_default_rate = blocks[2].find("div",{"class":re.compile("profile__bucket--2")})
		try:
			value = Loan_default_rate.find("div",{"class":re.compile("scalar__value")}).find('span').string
		except:
			value = "Data not available"
		sheet1.write(i,j,value)
		j = j+1
		
		i = i +1

#Save the output    
book.save('After-College.xls')
	



