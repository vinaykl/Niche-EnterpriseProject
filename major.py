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
sheet1.write(i,2,"Majors")

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

		url = link +"majors/"
		page = urllib2.urlopen(url)
		soup = BeautifulSoup(page.read(),"lxml")


		majors_expansion = soup.find("div",{"class":re.compile("majors-expansion")})
		#print majors_expansion
		profile_buckets = majors_expansion.find("div",{"class":re.compile("profile__buckets")})
		childrens = profile_buckets.children
		for child in childrens:
			#print child
			test = child.find("div",{"class":re.compile("majors-list__header")})
			if test == None:
				continue
			else:
				s = ""
				s = s + test.find('h3').string
				s = s+ "\n"
				majors_list = child.find("ul",{"class":re.compile("majors-list__list")})
				#print majors_list
				childrens = majors_list.children
				for child in childrens:
					#print child
					s = s + child.find("div",{"class":re.compile("majors-list__list__item__major")}).string
					s = s + " - " 
					s = s+  child.find("div",{"class":re.compile("majors-list__list__item__count")}).string	
					s = s + "\n"


				sheet1.write(i,j,s)
			j = j+1
		
		i = i +1

#Save the output    
book.save('Major.xls')
		
