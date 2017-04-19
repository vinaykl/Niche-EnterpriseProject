from bs4 import BeautifulSoup
import urllib2, re
import xlwt
import requests

# open the home page and get the numbers of pages in the the total 
url= 'https://colleges.niche.com/?degree=4-year&sort=best'
page = urllib2.urlopen(url)
soup = BeautifulSoup(page.read(),"lxml")
pages = int(soup.select("select.pagination__pages__selector option")[-1].text.split(None, 1)[1])

# Open the excel sheet to write the data
book = xlwt.Workbook(encoding="utf-8")
sheet1 = book.add_sheet('vkl')

# for Serial Number and template URL we need
i=1
test = "https://colleges.niche.com/?degree=4-year&sort=best&page="

# append page number to the template URL 
for page in range(1, pages+ 1):
    test1= test+str(page)
    page = urllib2.urlopen(test1)
    soup1 = BeautifulSoup(page.read(),"lxml")
    vinay = soup1.find("div",{"class":re.compile("search")})
    for link in vinay.find_all('a'):
	# remove the advertisements
	if link.get('href') == "https://colleges.niche.com/scholarships/":
		continue
	else:
		sheet1.write(i,0,i)
		sheet1.write(i,1,link.string)
		sheet1.write(i,2,link.get('href'))	
		i=i+1

#Save the output    
book.save('test.xls')



