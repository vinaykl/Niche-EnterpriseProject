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
sheet1.write(i,2,"Net Price")
sheet1.write(i,3,"In-State Tuition")
sheet1.write(i,4,"Out-of-State tuition")
sheet1.write(i,5,"Average Housing Costs")
sheet1.write(i,6,"Average meal plan cost")
sheet1.write(i,7,"Books & Supplies")
sheet1.write(i,8,"Tuition Guarantee Plan")
sheet1.write(i,9,"Tuition Payment Plan")
sheet1.write(i,10,"Prepaid Tuition Plan")
sheet1.write(i,11,"Students Taking Out Loans")
sheet1.write(i,12,"Average Loan Amount")
sheet1.write(i,13,"Any Financial Aid")
sheet1.write(i,14,"Average Total Aid Awarded")
sheet1.write(i,15,"Student Recieving Aid-Federal grant Aid")
sheet1.write(i,16,"Student Recieving Aid-State Grant Aid")
sheet1.write(i,17,"Student Recieving Aid-Institution Grant Aid")
sheet1.write(i,18,"Student Recieving Aid-Pell Grant")
sheet1.write(i,19,"Average Aid Awarded-Federal grant Aid")
sheet1.write(i,20,"Average Aid Awarded-State Grant Aid")
sheet1.write(i,21,"Average Aid Awarded-Institution Grant Aid")
sheet1.write(i,22,"Average Aid Awarded-Pell Grant")
sheet1.write(i,23,"Overall value")
sheet1.write(i,24,"of students feel like they are getting their money's worth out of their program")
sheet1.write(i,25,"Graduation Rate")


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

		url = link +"cost/"
		page = urllib2.urlopen(url)
		soup = BeautifulSoup(page.read(),"lxml")

		blocks = soup.find_all("section",{"class":re.compile("block--two")})

		# Bucket 1

		
		try:
			cost = blocks[0].find("div",{"class":re.compile("scalar__value")})
			net_cost = cost.find('span')
			value= net_cost.string
		except:
			value = "Data not available"
		sheet1.write(i,j,value)
		j = j+1

		# Bucket 2
		# inner bucket 1 - Instate fees
		bucket2_bucket1 = blocks[1].find("div",{"class":re.compile("profile__bucket--1")})
		
		try:
			instate = bucket2_bucket1.find("div",{"class":re.compile("scalar__value")})
			instate_cost = instate.find('span')
			value = instate_cost.string
		except:
			value = "Data not available"
		sheet1.write(i,j,value)
		j = j+1

		# inner bucket 2 - Out of state feesg
		bucket2_bucket2 = blocks[1].find("div",{"class":re.compile("profile__bucket--2")})
		
		try:
			outofstate = bucket2_bucket2.find("div",{"class":re.compile("scalar__value")})
			outofstate_cost = outofstate.find('span')
			value = outofstate_cost.string
		except:
			value = "Data not available"
		sheet1.write(i,j,value)
		j = j+1
  
		# inner bucket 3 - Housing , Meal and Books cost
		bucket2_bucket3 = blocks[1].find("div",{"class":re.compile("profile__bucket--3")})
		blanks = bucket2_bucket3.find("div",{"class":re.compile("blank__bucket")})

		childrens = blanks.children
		for child in childrens:
			try:
				Housing = child.find("div",{"class":re.compile("scalar__value")})
				Avg_Housing = Housing.find('span')
				value = Avg_Housing.string
			except:
				value = "Data not available"
			sheet1.write(i,j,value)
			j = j+1


		# inner bucket 4 - Tution prepaid plan, tuition guatantee plan ,Prepaid tuition plan
		bucket2_bucket4 = blocks[1].find("div",{"class":re.compile("profile__bucket--4")})
		blanks = bucket2_bucket4.find("div",{"class":re.compile("blank__bucket")})

		childrens = blanks.children
		for child in childrens:
			Tuition = child.find("div",{"class":re.compile("scalar__value")})
			Tuition_plans = Tuition.find('span')
			try:
				value = Tuition_plans.string
			except:
				value = "Data not available"
			sheet1.write(i,j,value)
			j = j+1


		# Bucket_2 --- >>> check whether it is works for all URL ( Fixed it ) 

		profile_bucket1 = blocks[2].find("div",{"class":re.compile("profile__bucket--1")})
		
		try:
			scalar_value = profile_bucket1.find("div",{"class":re.compile("scalar__value")})
			val = scalar_value.find('span')
			value= val.string
		except:
			value = "Data not available"
		sheet1.write(i,j,value)
		j = j+1

		profile_bucket2 = blocks[2].find("div",{"class":re.compile("profile__bucket--2")})
		
		try:
			scalar_value = profile_bucket2.find("div",{"class":re.compile("scalar__value")})
			val = scalar_value.find('span')
			value = val.string
		except:
			value = "Data not available"
		sheet1.write(i,j,value)
		j = j+1


		# bucket_two_two_two


		profile_buckets = blocks[3].find("div",{"class":re.compile("profile__buckets")})

		profile_bucket1 = profile_buckets.find("div",{"class":re.compile("profile__bucket--1")})
		try:
			scalar_value = profile_bucket1.find("div",{"class":re.compile("scalar__value")})
			val = scalar_value.find('span')
			value =  val.string
		except:
			value = "Data not available"
		sheet1.write(i,j,value)
		j = j+1

		profile_bucket2 = profile_buckets.find("div",{"class":re.compile("profile__bucket--2")})
		
		try:
			scalar_value = profile_bucket2.find("div",{"class":re.compile("scalar__value")})
			val = scalar_value.find('span')
			value = val.string
		except:
			value = "Data not available"
		sheet1.write(i,j,value)
		j = j+1
	

		profile_bucket3 = profile_buckets.find("div",{"class":re.compile("profile__bucket--3")})
		profile_table_rows = profile_bucket3.find("ul",{"class":re.compile("profile__table__rows")})
		rows = profile_table_rows.children
		for row in rows:
			try:
				value = row.find("div",{"class":re.compile("fact__table__row__value")}).text
			except:
				value = "Data not available"
			sheet1.write(i,j,value)
			j = j+1

		profile_bucket4 = profile_buckets.find("div",{"class":re.compile("profile__bucket--4")})
		profile_table_rows = profile_bucket4.find("ul",{"class":re.compile("profile__table__rows")})
		rows = profile_table_rows.children
		for row in rows:
			try:
				value = row.find("div",{"class":re.compile("fact__table__row__value")}).text
			except:
				value = "Data not available"
			sheet1.write(i,j,value)
			j = j+1

		# Bucket 4
		bucket4 = soup.find("section",{"class":re.compile("block--one-two")})
		profile_buckets = bucket4.find("div",{"class":re.compile("profile__buckets")})

		profile_bucket1 = profile_buckets.find("div",{"class":re.compile("profile__bucket--1")})
		try:
			value= profile_bucket1.find("div",{"class":re.compile("niche__grade")}).text
		except:
			value = "Data not available"
		sheet1.write(i,j,value)
		j = j+1

		profile_bucket2 = profile_buckets.find("div",{"class":re.compile("profile__bucket--2")})
		try:
			value = profile_bucket2.find("div",{"class":re.compile("poll__single__percent__label")}).text
		except:
			value = "Data not available"
		sheet1.write(i,j,value)
		j = j+1

		profile_bucket3 = profile_buckets.find("div",{"class":re.compile("profile__bucket--3")})
		try:
			value= profile_bucket3.find("div",{"class":re.compile("scalar__value")}).find('span').text
		except:
			value = "Data not available"
		sheet1.write(i,j,value)
		j = j+1
		i = i+1

#Save the output    
book.save('Costs.xls')

