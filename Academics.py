from bs4 import BeautifulSoup
import urllib2, re
import xlwt
import requests
from xlrd import open_workbook

# Open the excel sheet to write the data
book = xlwt.Workbook(encoding="utf-8")
sheet1 = book.add_sheet('Costs')

i = 0

sheet1.write(i,0,"SR Number")
sheet1.write(i,1,"College Name")
sheet1.write(i,2,"Student Faculty Ratio")
sheet1.write(i,3,"Female Professors")
sheet1.write(i,4,"Male Professors")
sheet1.write(i,5,"Average Professor Salary")
sheet1.write(i,6,"African American")
sheet1.write(i,7,"Asian American")
sheet1.write(i,8,"Hispanic")
sheet1.write(i,9,"International")
sheet1.write(i,10,"Native American")
sheet1.write(i,11,"Pacific Islander")
sheet1.write(i,12,"Unknown")
sheet1.write(i,13,"White")
sheet1.write(i,14,"of students say professors are passionate about the topics they teach.")
sheet1.write(i,15,"of students say professors care about their students' success.")
sheet1.write(i,16,"of students say professors are engaging and easy to understand.")
sheet1.write(i,17,"of students agree professors are approachable and helpful when needed.")
sheet1.write(i,18,"Academics")
sheet1.write(i,19,"Graduation rate")
sheet1.write(i,20,"Full-Time Retention Rate")
sheet1.write(i,21,"Part-Time Retention Rate")
sheet1.write(i,22,"Academic Calendar")
sheet1.write(i,23,"Research Funding per Student")
sheet1.write(i,24,"Evening Degree Programs")
sheet1.write(i,25,"Teacher Certification")
sheet1.write(i,26,"Distance Education")
sheet1.write(i,27,"Study Abroad")
sheet1.write(i,28,"of students say they attend class.")
sheet1.write(i,29,"of students say they take advantage of office hours/study sessions.")
sheet1.write(i,30,"of students say they do all their homework.")


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

		url = link + "academics/"
		page = urllib2.urlopen(url)
		soup = BeautifulSoup(page.read(),"lxml")

		# total of 4 sections
		blocks = soup.find_all("section",{"class":re.compile("block--two")})

		# About the Proffesors
		# Student Professor ratio ,  Female Professor, Male professor , Average Salary , Faculty Racial diversity
		# block--two
		profile1 = blocks[0].find("div",{"class":re.compile("profile__bucket--1")})
		try:
			value = profile1.find("div",{"class":re.compile("scalar__value")}).find("span").string
		except:
			value = "Data not available"
		sheet1.write(i,j,value)
		j = j+1

		childrens = profile1.find_all("div",{"class":re.compile("scalar--three")})
		for child in childrens:
			try:
				value = child.find("div",{"class":re.compile("scalar__value")}).find("span").string
			except:
				value = "Data not available"
			sheet1.write(i,j,value)
			j = j+1

		profile2 =  blocks[0].find("div",{"class":re.compile("profile__bucket--2")})
		lists = profile2.find("ul",{"class":re.compile("breakdown__rows")})
		childrens = lists.find_all("li",{"class":re.compile("fact__table__row")})
		for child in childrens:
			try:
				value = child.find("div",{"class":re.compile("fact__table__row__value")}).text
			except:
				value = "Data not available"
			sheet1.write(i,j,value)
			j = j+1


		# What students say about professors
		# of students say professors are passionate about the topics they teach , of students say professors are engaging and easy 			to understand, of students say professors care about their students' success , of students agree professors are 		approachable and helpful when needed
		# block--two

		profile_bucket1 = blocks[1].find("div",{"class":re.compile("profile__bucket--1")})
		childrens = profile_bucket1.find_all("div",{"class":re.compile("poll__single--piechart")})
		for child in childrens:
			try:
				value = child.find("div",{"class":re.compile("poll__single__percent__label")}).string
			except:
				value = "Data not available"
			sheet1.write(i,j,value)
			j = j+1


		profile_bucket2 = blocks[1].find("div",{"class":re.compile("profile__bucket--2")})
		childrens = profile_bucket2.find_all("div",{"class":re.compile("poll__single--piechart")})
		for child in childrens:
			try:
				value = child.find("div",{"class":re.compile("poll__single__percent__label")}).string
			except:
				value = "Data not available"
			sheet1.write(i,j,value)
			j = j+1
	
		# Quality of Education
		# academics rating , Graduation rate , Full-Time Retention Rate , Part-Time Retention Rate , Academic Calendar,Research 		Funding per Student ,Non-traditional Learning -list
		# block--two-two

		profile_bucket1 = blocks[2].find("div",{"class":re.compile("profile__bucket--1")})
		try:
			value = profile_bucket1.find("div",{"class":re.compile("niche__grade niche__grade--section")}).string
		except:
			value = "Data not available"
		sheet1.write(i,j,value)
		j = j+1

		profile2 = blocks[2].find("div",{"class":re.compile("profile__bucket--2")})
		try:
			value = profile2.find("div",{"class":re.compile("scalar__value")}).find("span").string
		except:
			value = "Data not available"
		sheet1.write(i,j,value)
		j = j+1

		profile3 = blocks[2].find("div",{"class":re.compile("profile__bucket--3")})
		childrens = profile3.find_all("div",{"class":re.compile("scalar--three")})

		for child in childrens:
			try:
				value = child.find("div",{"class":re.compile("scalar__value")}).find("span").string
			except:
				value = "Data not available"
			sheet1.write(i,j,value)
			j = j+1

		profile4 =  blocks[2].find("div",{"class":re.compile("profile__bucket--4")})
		profile_table_rows = profile4.find("ul",{"class":re.compile("profile__table__rows")})
		rows = profile_table_rows.children
		for row in rows:
			try:
				value = row.find("div",{"class":re.compile("fact__table__row__value")}).text
			except:
				value = "Data not available"
			sheet1.write(i,j,value)
			j = j+1

		# What Students Say About Education
		# poll - Which statement(s) best describes the quality of education you're getting? , polls - of students say they attend 			class, of students say they take advantage of office hours/study sessions , of students say they do all their homework.
		# block--two-poll

		#profile bucket 1 - Yet to do ( Not there in most of the colleges )

		profile2 = blocks[3].find("div",{"class":re.compile("profile__bucket--2")})
		childrens = profile2.find_all("div",{"class":re.compile("poll__single__percent__label")})
		for child in childrens:
			try:
				value = child.string
			except:
				value = "Data not available"
			sheet1.write(i,j,value)
			j = j+1
		
		i = i + 1

#Save the output    
book.save('Academics.xls')


