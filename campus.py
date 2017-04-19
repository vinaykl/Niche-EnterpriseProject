
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
sheet1.write(i,2,"Dorms")
sheet1.write(i,3,"On-Campus Housing Available")
sheet1.write(i,4,"Freshmen Required to Live on Campus")
sheet1.write(i,5,"Average Housing Costs")
sheet1.write(i,6,"% of students say overall dorm quality is great.")
sheet1.write(i,7,"% of students say the dorms' social atmosphere is great.")
sheet1.write(i,8,"Campus Food")
sheet1.write(i,9,"Meal Plan Available")
sheet1.write(i,10,"Average Meal Plan Cost")
sheet1.write(i,11,"POLL - What are the best food options on campus?")
sheet1.write(i,12,"Campus")
sheet1.write(i,13,"Student Life ")
sheet1.write(i,14,"POLL - What is your overall opinion of your school and the campus community? ")
sheet1.write(i,15,"Safety")
sheet1.write(i,16,"% of students say they feel extremely safe and secure on campus.")
sheet1.write(i,17,"POLL - How does peer pressure affect students' use of drugs and alcohol?")
sheet1.write(i,18,"POLL - How visible are the campus police on campus?")
sheet1.write(i,19,"Men's Varsity Sports")
sheet1.write(i,20,"Women's Varsity Sports")
sheet1.write(i,21,"POLL - How popular are varsity sports on campus?")
sheet1.write(i,22,"POLL - What are your favorite campus events or traditions?")
sheet1.write(i,23,"Party Scene?")
sheet1.write(i,24,"POLL - What is the party scene like on campus?")
sheet1.write(i,25,"POLL - What is the biggest party event of the year?")

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


		url = link + "campus-life/"
		opener = urllib2.build_opener(urllib2.HTTPRedirectHandler)
    		request = opener.open(url)
    		if request.url != url:
			i = i +1
			continue

		page = urllib2.urlopen(url)
		soup = BeautifulSoup(page.read(),"lxml")

		# Section 1 - Housing
		# Dorms quality, On Campus Housing available , Freshmen Requied to stay, Average cost
		# Poll - Dorm quality , Dorm Social atmosphere
		# block--two


		buckets = soup.find_all("section",{"class":re.compile("block")})

		bucket1 = soup.find("section",{"class":re.compile("block--two")})
		# profile bucket 1
		bucket1_section1 = buckets[3].find("div",{"class":re.compile("profile__bucket--1")})
		try:
			value = bucket1_section1.find("div",{"class":re.compile("profile__grade")}).find("div",{"class":re.compile("niche__grade niche")}).string
		except:
			value = "Data not available"
		sheet1.write(i,j,value)
		j = j+1

		childrens = bucket1_section1.find_all("div",{"class":re.compile("scalar--three")})
		for child in childrens:
			try:
				value = child.find("div",{"class":re.compile("scalar__value")}).find('span').string
			except:
				value = "Data not available"
			sheet1.write(i,j,value)
			j = j+1

		# profile bucket 2
		bucket1_section2 = buckets[3].find("div",{"class":re.compile("profile__bucket--2")})
		childrens = bucket1_section2.find_all("div",{"class":re.compile("poll__single--piechart")})
		for child in childrens:
			try:
				value = child.find("div",{"class":re.compile("poll__single__percent__label")}).string
			except:
				value = "Data not available"
			sheet1.write(i,j,value)
			j = j+1



		# Section 2 - Food
		# Food quality, Meal Plan Available , Average Meal plan cost , Poll -Best food options
		# block--horiz-poll
		bucket2_section1 = buckets[5].find("div",{"class":re.compile("profile__bucket--1")})
		try:
			value = bucket2_section1.find("div",{"class":re.compile("profile__grade")}).find("div",{"class":re.compile("niche__grade niche")}).string
		except:
			value = "Data not available"
		sheet1.write(i,j,value)
		j = j+1

		childrens = bucket2_section1.find_all("div",{"class":re.compile("scalar--three")})
		for child in childrens:
			try:
				value = child.find("div",{"class":re.compile("scalar__value")}).find('span').string
			except:
				value = "Data not available"
			sheet1.write(i,j,value)
			j = j+1

		
		s = ""
		try:
			profile_bucket2 =  buckets[5].find("div",{"class":re.compile("profile__bucket--2")})
			ul = profile_bucket2.find("ul",{"class":re.compile("poll__table__results")})
			li = ul.find_all("li",{"class":re.compile("poll__table__result__item")})
			for element in li:
				s = s + element.find("div",{"class":re.compile("poll__table__result__label")}).string
				s = s+ "-"+ element.find("div",{"class":re.compile("poll__table__result__percent")}).string
				s = s+ "\n"
		except:
			v ="vinay"
		
		sheet1.write(i,j,s)
		j = j+1



		# Section 3 - Campus Quality
		# Campus , Student Life , poll - What is your overall opinion of your school and the campus community?
		# block--two
		grades = buckets[6].find_all("div",{"class":re.compile("profile__grade")})
		try:
			value = grades[0].find("div",{"class":re.compile("niche__grade niche__")}).string 
		except:
			value = "Data not available"
			sheet1.write(i,j,value)
			j = j+1

		try:
			value = grades[3].find("div",{"class":re.compile("niche__grade niche__")}).string
		except:
			value = "Data not available"
			sheet1.write(i,j,value)
			j = j+1 

		
		try:
			profile_bucket2 =  buckets[6].find("div",{"class":re.compile("profile__bucket--2")})
			ul = profile_bucket2.find("ul",{"class":re.compile("poll__table__results")})
			li = ul.find_all("li",{"class":re.compile("poll__table__result__item")})
			s = ""
			for element in li:
				s = s + element.find("div",{"class":re.compile("poll__table__result__label")}).string
				s = s+ "-"+ element.find("div",{"class":re.compile("poll__table__result__percent")}).string
				s = s+ "\n"
		except:
			v = "vinay"
		sheet1.write(i,j,s)
		j = j+1


		# Section 4 - Safety
		# Safety Rating, Safety student poll extremly safe , poll - peer pressure on drugs , poll - how often police is visible 
		# block-two-two
		try:
			value = buckets[8].find("div",{"class":re.compile("profile__bucket--1")}).find("div",{"class":re.compile("niche__grade niche__")}).string 
		except:
			value = "Data not available"
			sheet1.write(i,j,value)
			j = j+1 

		try:
			value = buckets[8].find("div",{"class":re.compile("profile__bucket--2")}).find("div",{"class":re.compile("poll__single__percent__label")}).string
		except:
			value = "Data not available"
			sheet1.write(i,j,value)
			j = j+1

		s = ""
		try:
			profile_bucket3 =  buckets[8].find("div",{"class":re.compile("profile__bucket--3")})
			ul = profile_bucket3.find("ul",{"class":re.compile("poll__table__results")})
			li = ul.find_all("li",{"class":re.compile("poll__table__result__item")})
			
			for element in li:
				s= s+ element.find("div",{"class":re.compile("poll__table__result__label")}).string
				s =s + " - " + element.find("div",{"class":re.compile("poll__table__result__percent")}).string
				s = s+ "\n"
		except:
			v = "vinay"
		sheet1.write(i,j,s)
		j = j+1


		s=""
		try:
			profile_bucket4 =  buckets[8].find("div",{"class":re.compile("profile__bucket--4")})
			ul = profile_bucket4.find("ul",{"class":re.compile("poll__table__results")})
			li = ul.find_all("li",{"class":re.compile("poll__table__result__item")})
		
			for element in li:
				s= s+ element.find("div",{"class":re.compile("poll__table__result__label")}).string
				s =s +" - " + element.find("div",{"class":re.compile("poll__table__result__percent")}).string
				s = s+ "\n"
		except:
			v= "vinay"
		sheet1.write(i,j,s)
		j = j+1


		# Section 5 - Sports
		# List of varsity men sports, List of Women varsity sports, Poll - popular carsity sports , Poll - Famous college tradition
		# block-two-two

		profile_bucket1 =  buckets[10].find("div",{"class":re.compile("profile__bucket--1")})
		try:
			value = profile_bucket1.find("div",{"class":re.compile("scalar__value")}).find('span').string
		except:
			value = "Data not available"
			sheet1.write(i,j,value)
			j = j+1

		profile_bucket2 =  buckets[10].find("div",{"class":re.compile("profile__bucket--2")})
		try:
			value = profile_bucket2.find("div",{"class":re.compile("scalar__value")}).find('span').string
		except:
			value = "Data not available"
			sheet1.write(i,j,value)
			j = j+1

		s= ""
		try:
			profile_bucket3 =  buckets[10].find("div",{"class":re.compile("profile__bucket--3")})
			ul = profile_bucket3.find("ul",{"class":re.compile("poll__table__results")})
			li = ul.find_all("li",{"class":re.compile("poll__table__result__item")})	
			for element in li:
				s = s+ element.find("div",{"class":re.compile("poll__table__result__label")}).string
				s = s + " - " + element.find("div",{"class":re.compile("poll__table__result__percent")}).string
				s = s+ "\n"
		except:
			v = "vinay"
		sheet1.write(i,j,s)
		j = j+1

		
		try:
			profile_bucket4 =  buckets[10].find("div",{"class":re.compile("profile__bucket--4")})
			ul = profile_bucket4.find("ul",{"class":re.compile("poll__table__results")})
			li = ul.find_all("li",{"class":re.compile("poll__table__result__item")})
			s = ""
			for element in li:
				s = s+ element.find("div",{"class":re.compile("poll__table__result__label")}).string
				s = s+ " - " + element.find("div",{"class":re.compile("poll__table__result__percent")}).string
				s = s+ "\n"
		except:
			v= "vinay"
		sheet1.write(i,j,s)
		j = j+1




		# Section 6 - party Scene
		# party rating , Poll- Party scene in college , Poll- Biggest party event of the year
		# block--one-two

		profile_bucket1 =  buckets[12].find("div",{"class":re.compile("profile__bucket--1")})
		try:
			value = profile_bucket1.find("div",{"class":re.compile("niche__grade niche__")}).string  
		except:
			value = "Data not available"
			sheet1.write(i,j,value)
			j = j+1

		try:
			profile_bucket2 =  buckets[12].find("div",{"class":re.compile("profile__bucket--2")})
			ul = profile_bucket2.find("ul",{"class":re.compile("poll__table__results")})
			li = ul.find_all("li",{"class":re.compile("poll__table__result__item")})
			s=""
			for element in li:
				s = s+ element.find("div",{"class":re.compile("poll__table__result__label")}).string
				s = s + " - " + element.find("div",{"class":re.compile("poll__table__result__percent")}).string
				s = s+ "\n"
		except:
			v = "vinay"
		sheet1.write(i,j,s)
		j = j+1

		s=""
		try:
			profile_bucket3 =  buckets[12].find("div",{"class":re.compile("profile__bucket--3")})
			ul = profile_bucket3.find("ul",{"class":re.compile("poll__table__results")})
			li = ul.find_all("li",{"class":re.compile("poll__table__result__item")})
			
			for element in li:
				s =s + element.find("div",{"class":re.compile("poll__table__result__label")}).string
				s = s + " - "+ element.find("div",{"class":re.compile("poll__table__result__percent")}).string
				s = s+ "\n"
		except:
			v ="vinay"
		sheet1.write(i,j,s)
		j = j+1

		i = i+1

#Save the output    
book.save('Campus.xls')



