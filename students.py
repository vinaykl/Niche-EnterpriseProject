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
sheet1.write(i,2,"Female Undergrads")
sheet1.write(i,3,"Male undergrads")
sheet1.write(i,4,"in-state")
sheet1.write(i,5,"Out-of-state")
sheet1.write(i,6,"International")
sheet1.write(i,7,"Unknown")
sheet1.write(i,8,"Under 18")
sheet1.write(i,9,"18-19")
sheet1.write(i,10,"20-21")
sheet1.write(i,11,"22-24")
sheet1.write(i,12,"Above 25")
sheet1.write(i,13,"Household Income Levels - <$30k ")
sheet1.write(i,14,"Household Income Levels - $30k-$48k")
sheet1.write(i,15,"Household Income Levels - $49k-$75k")
sheet1.write(i,16,"Household Income Levels - $76k-$110k")
sheet1.write(i,17,"Household Income Levels - $110k+")
sheet1.write(i,18,"African American")
sheet1.write(i,19,"Asian")
sheet1.write(i,20,"Hispanic")
sheet1.write(i,21,"International (Non-Citizen)")
sheet1.write(i,22,"Multiracial")
sheet1.write(i,23,"Native American")
sheet1.write(i,24,"Pacific Islander")
sheet1.write(i,25,"Unknown")
sheet1.write(i,26,"White")
sheet1.write(i,27,"Republican")
sheet1.write(i,28,"Democratic")
sheet1.write(i,29,"Independent")
sheet1.write(i,30,"Other party not mentioned")
sheet1.write(i,31,"I don't care about politics")
sheet1.write(i,32,"Progressive/very liberal")
sheet1.write(i,33,"Liberal")
sheet1.write(i,34,"Moderate")
sheet1.write(i,35,"Conservative")
sheet1.write(i,36,"Very conservative")
sheet1.write(i,37,"Libertarian")
sheet1.write(i,38,"Not sure")

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

		url = link + "students/"
		page = urllib2.urlopen(url)
		soup = BeautifulSoup(page.read(),"lxml")

		blocks = soup.find_all("section",{"class":re.compile("block--two")})


		profile_buckets = blocks[0].find("div",{"class":re.compile("profile__buckets")})
		buckets = profile_buckets.find_all("div",{"class":re.compile("profile__bucket--")})

		# buckets[0] buckets[1] buckets[2]

		# Male and Femaile 
		#print buckets[1]
		vkl = buckets[1].find_all("div",{"class":re.compile("scalar--three")})
		try:
			value = vkl[0].find("div",{"class":re.compile("scalar__value")}).find('span').string
		except:
			value = "Data not available"
		sheet1.write(i,j,value)
		j = j+1

		try:
			value = vkl[1].find("div",{"class":re.compile("scalar__value")}).find('span').string
		except:
			value = "Data not available"
		sheet1.write(i,j,value)
		j = j+1

		# in-state out-state distribution


		vkl1 = buckets[1].find("div",{"class":re.compile("breakdown--bar_chart")})
		vkl2 = vkl1.find("ul",{"class":re.compile("breakdown__rows")})
		childrens = vkl2.children
		for child in childrens:
			try:
				value = child.find("div",{"class":re.compile("fact__table__row__value")}).text
			except:
				value = "Data not available"
			sheet1.write(i,j,value)
			j = j+1

		# Age distribution

		Age = buckets[2].find("div",{"class":re.compile("breakdown--bar_chart")})
		vkl2 = Age.find("ul",{"class":re.compile("breakdown__rows")})
		childrens = vkl2.children
		for child in childrens:
			try:
				value = child.find("div",{"class":re.compile("fact__table__row__value")}).text
			except:
				value = "Data not available"
			sheet1.write(i,j,value)
			j = j+1
		
		# Household income distribution

		Household = buckets[3].find("div",{"class":re.compile("breakdown--bar_chart")})
		vkl2 = Household.find("ul",{"class":re.compile("breakdown__rows")})
		childrens = vkl2.children
		for child in childrens:
			try:
				value = child.find("div",{"class":re.compile("fact__table__row__value")}).text
			except:
				value = "Data not available"
			sheet1.write(i,j,value)
			j = j+1


		# ethnic distribution

		profile_buckets = blocks[1].find("div",{"class":re.compile("profile__buckets")})
		buckets = profile_buckets.find_all("div",{"class":re.compile("profile__bucket--")})

		ethnic = buckets[1].find("div",{"class":re.compile("breakdown--bar_chart")})
		vkl2 = ethnic.find("ul",{"class":re.compile("breakdown__rows")})
		childrens = vkl2.children
		for child in childrens:
			try:
				value= child.find("div",{"class":re.compile("fact__table__row__value")}).text
			except:
				value = "Data not available"
			sheet1.write(i,j,value)
			j = j+1


		# Political party

		profile_buckets = blocks[2].find("div",{"class":re.compile("profile__buckets")})
		buckets = profile_buckets.find_all("div",{"class":re.compile("profile__bucket--")})
		party = buckets[0].find("div",{"class":re.compile("toggle__content--profiles-visible--hidden")})
		#print party
		try:
			vkl2 = party.find("ul",{"class":re.compile("poll__table__results")})
			childrens = vkl2.children
			for child in childrens:
				try:
					value = child.find("div",{"class":re.compile("poll__table__result__percent")}).text
				except:
					value = "Data not available"
				sheet1.write(i,j,value)
				j = j+1
		except:
			j = j + 5
		# views 

		profile_buckets = blocks[2].find("div",{"class":re.compile("profile__buckets")})
		buckets = profile_buckets.find_all("div",{"class":re.compile("profile__bucket--")})
		party = buckets[1].find("div",{"class":re.compile("toggle__content--profiles-visible--hidden")})
		#print party
		try:
			vkl2 = party.find("ul",{"class":re.compile("poll__table__results")})
			childrens = vkl2.children
			for child in childrens:
				try:
					value = child.find("div",{"class":re.compile("poll__table__result__percent")}).text
				except:
					value = "Data not available"
				sheet1.write(i,j,value)
				j = j+1
		except:
			j = j + 7
		i = i+1

#Save the output    
book.save('Students.xls')


