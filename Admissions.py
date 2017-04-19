
from bs4 import BeautifulSoup
import urllib2, re
import xlwt
import requests


from xlrd import open_workbook

# Open the excel sheet to write the data
book = xlwt.Workbook(encoding="utf-8")
sheet1 = book.add_sheet('Admissions')

i = 0

sheet1.write(i,0,"SR Number")
sheet1.write(i,1,"College Name")
sheet1.write(i,2,"SAT Range")
sheet1.write(i,3,"SAT Reading")
sheet1.write(i,4,"SAT Math")
sheet1.write(i,5,"SAT Writing")
sheet1.write(i,6,"Students Submitting SAT")
sheet1.write(i,7,"ACT Range")
sheet1.write(i,8,"ACT English")
sheet1.write(i,9,"ACT Math")
sheet1.write(i,10,"ACT Writing")
sheet1.write(i,11,"Students Submitting ACT")
sheet1.write(i,12,"Acceptance Rate")
sheet1.write(i,13,"Application Fee")
sheet1.write(i,14,"Application Website")
sheet1.write(i,15,"Female Applicants")
sheet1.write(i,16,"Female Acceptance")
sheet1.write(i,17,"Male Applicants")
sheet1.write(i,18,"Male Acceptance")
sheet1.write(i,19,"High School GPA")
sheet1.write(i,20,"High School Rank")
sheet1.write(i,21,"High School Transcript")
sheet1.write(i,22,"College Prep Courses")
sheet1.write(i,23,"SAT/ACT")
sheet1.write(i,24,"Recommendations")
sheet1.write(i,25,"students say the admissions process made them feel like the school cared about them as an applicant")
sheet1.write(i,26,"of students feel the admissions process evaluated them individually as a real person, not just a set of numbers.")




wb = open_workbook('test.xls')
for sheet in wb.sheets():
	number_of_rows = sheet.nrows
	number_of_columns = sheet.ncols
	i = 1
	for row in range(1 , number_of_rows):
		name = (sheet.cell(row,1).value)
		link = (sheet.cell(row,2).value)
		filename = name.replace(" ", "")
		print name , link , filename	

		sheet1.write(i,0,i)
		sheet1.write(i,1,name)
		admissions = link +"admissions/"
		admissions_page = urllib2.urlopen(admissions)
		admissions_soup = BeautifulSoup(admissions_page.read(),"lxml")
		# three major sections     

		blocks = admissions_soup.find_all("section",{"class":re.compile("block--two")})

		# Can I get in?
		# SAT Range , ACT Range , Acceptance Rate, Application Fee, Application Website, Female Applicants , Female Acceptance,Male 			Applicants, Male Acceptance
		# block-two-two

		profile1 = blocks[0].find("div",{"class":re.compile("profile__bucket--1")})
		# SAT Range
		try:
			value= profile1.find("div",{"class":re.compile("scalar")}).find("div",{"class":re.compile("scalar__value")}).find("span").string
		except:
			value = "Data not available"
		sheet1.write(i,2,value)		

		childrens = profile1.find_all("div",{"class":re.compile("scalar--three")})

		j = 3
		for child in childrens:
			try:
				value = child.find("div",{"class":re.compile("scalar__value")}).find("span").string
			except:
				value = "Data Not Available"
			sheet1.write(i,j,value)
			j = j+1	

		profile2 = blocks[0].find("div",{"class":re.compile("profile__bucket--2")})
		#print profile2
		try:
			value = profile2.find("div",{"class":re.compile("scalar")}).find("div",{"class":re.compile("scalar__value")}).find("span").string
		except:
			value = "Data not available"
		sheet1.write(i,j,value)
		j = j+1	

		childrens = profile2.find_all("div",{"class":re.compile("scalar--three")})
		for child in childrens:
			try:
				value = child.find("div",{"class":re.compile("scalar__value")}).find("span").string
				sheet1.write(i,j,value)
				j = j+1	
			except:
				value = "Data not available"
				sheet1.write(i,j,value)
				j = j+1	
				continue


		profile3 = blocks[0].find("div",{"class":re.compile("profile__bucket--3")})
		#print profile3
		try:
			value = profile3.find("div",{"class":re.compile("scalar")}).find("div",{"class":re.compile("scalar__value")}).find("span").string
		except:
			value = "Data not available"
		sheet1.write(i,j,value)
		j = j+1	

		try:
			value = profile3.find("div",{"class":re.compile("scalar--three")}).find("div",{"class":re.compile("scalar__value")}).find("span").string
		except:
			value = "Data not available"
		sheet1.write(i,j,value)
		j = j+1	

		try:
			value = profile3.find("div",{"class":re.compile("profile__website")}).find('a').get('href')
		except:
			value = "Data not available"
		sheet1.write(i,j,value)
		j = j+1	


		profile4 = blocks[0].find("div",{"class":re.compile("profile__bucket--4")})
		#print profile4

		childrens = profile4.find_all("div",{"class":re.compile("scalar--three")})
		for child in childrens:
			try:
				value = child.find("div",{"class":re.compile("scalar__value")}).find("span").string
			except:
				value = "Data not available"
			sheet1.write(i,j,value)
			j = j+1	



		# Admissions Considerations 
		# High School GPA , High School Rank , High School Transcript ,College Prep Courses, SAT/ACT , Recommendations
		# block--one

		sec2 = admissions_soup.find("section",{"class":re.compile("block--one")})
		lists = sec2.find("div",{"class":re.compile("profile__bucket--1")}).find("ul",{"class":re.compile("profile__table__rows")})
		for list1 in lists:
			try:
				value = list1.find("div",{"class":re.compile("fact__table__row__value")}).text
			except:
				value = "Data not available"
			sheet1.write(i,j,value)
			j = j+1	


		# What Students Say
		# poll -students say the admissions process made them feel like the school cared about them as an applicant, poll - of 			students feel the admissions process evaluated them individually as a real person, not just a set of numbers
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
		i = i+1


#Save the output    
book.save('Admissions.xls')
