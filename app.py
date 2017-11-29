import unicodedata
import urllib2
from urllib2 import  urlparse

import justext
import requests
from bs4 import BeautifulSoup
from xlrd import open_workbook

#this function is for converting unicode string to normal string
def normalize_unicode(data):
    return unicodedata.normalize('NFKD', data).encode('ascii', 'ignore')

#function to create industries_detail.html file
def convert_to_html(data):
    strTable = "<html><table border = '1' bordercolor = 'green' ><tr><th>Ticker</th><th>Company Name</th><th>Market Capitalization</th><th>TTM Sales$</th><th>Employees</th><th>Industry</th><th>Perm Id</th><th>Hierarchical Id</th></tr>"
    key = data.keys()
    for item in key:
        strRW = "<tr><td nowrap align ='center'>" + data[item][0] + "</td><td nowrap align ='center'>" + item + "</td>"+"<td nowrap align ='center'>"  + data[item][1]+"</td><td nowrap align ='center'>"  +   data[item][2]+"</td><td nowrap align ='center'>"  +   data[item][3]+"</td><td nowrap align ='center'>"  +   data[item][4]+"</td><td nowrap align ='center'>"  +   str(data[item][5])+"</td><td nowrap align ='center'>"  +   str(data[item][6])+"</td></tr>"
        strTable = strTable + strRW

    strTable = strTable + "</table></html>"

    hs = open("industries_detail.html", 'w')
    hs.write(strTable)
    hs.close()


#function to create html file of employee details of every industry
def convert_to_html2(data):
    strTable = "<br/><html><table border = '1' bordercolor = 'green' width='100%'><tr><th>Name</th><th>Age</th><th>Since</th><th>Current Position</th><th>Description</th></tr>"
    for item in data:
        strRW = "<tr><td nowrap align ='center'>" + str(item.name) + "</td><td nowrap align ='center'>" + str(item.age) + "</td>"+"<td nowrap align ='center'>"  + str(item.since)+"</td><td nowrap align ='center'>"  +   str(item.current_position)+"</td><td align ='left'>"  +   str(item.description)+"</td></tr>"
        strTable = strTable + strRW

    strTable = strTable + "</table></htmsl><br/>"

    hs = open("employee_detail.html", 'a')
    hs.write(strTable)
    hs.close()

#model class of each employee
class Person:
    def __init__(self, name, age, since, current_position, description):
        self.name = name                    #name of the employee
        self.age = age                      #age of the employee
        self.since = since                  #year of joining
        self.current_position = current_position  #current position
        self.description = description             #biography

    def __str__(self):
        return  ("Name: " + str(self.name) + " \n " +
              "Age:" + str(self.age) + " \n " +
              "Since:" + str(self.since) + " \n " +
              "Current Postion: " + str(self.current_position) + " \n " +
              "Description: " + str(self.description) + "\n")

    def __repr__(self):
        return ("Name: " + str(self.name) + " \n " +
              "Age:" + str(self.age) + " \n " +
              "Since:" + str(self.since) + " \n " +
              "Current Postion: " + str(self.current_position) + " \n " +
              "Description: " + str(self.description) + "\n")


#table class for employee details
class Table:
    def __init__(self):
        self.type = ''
        self.header = []
        self.data = {}

    def populate(self, table):
        data = []
        dict1 = {}
        table_body = table.find('tbody')

        rows = table_body.find_all('tr')
        header_row = rows[0]
        self.header = [normalize_unicode(ele.text.strip()) for ele in header_row.find_all('th')]
        rows = rows[1:]
        for row in rows:
            cols = row.find_all('td')
            cols = [ele.text.strip() for ele in cols]
            data.append([normalize_unicode(ele) for ele in cols])

        for row in data:
            dict1[row[0]] = row

        self.data = dict1

#mapping both details and description of employee
        for data in self.header:
            if data == 'Description':
                self.type = 'Description Table'
                break
            elif data == 'Age':
                self.type = 'Person Data Table'
                break

    def __str__(self):
        return "Header: \n" + str(self.header) + "\nData: \n" + str(self.data)


#basic class for crawling, scrapping and cleaning data
class Crawl:
    f = open('employee_detail.html', 'w')
    f.truncate()
    f.close()
    seed = 'http://www.reuters.com/finance/markets/indices'
    all_links = set()
    links = list()

    try:

        r = requests.get(seed)
        if r.status_code == 200:
            print ('Fetching in page links...')
            # print r.status_code
            content = r.content
            soup = BeautifulSoup(content, "lxml")
            tags = soup('a')
            flg = 0
            for a in tags:
                href = a.get("href")
                if href is not None:
                    new_url = urlparse.urljoin(seed, href)
                    if new_url.find("sector") != -1:
                        print new_url
                        links.append(new_url)                   # 'links' contains URLs of all 10 sectors


        elif r.status_code == 403:
            print "Error: 403 Forbidden url"
        elif r.status_code == 404:
            print "Error: 404 URL not found"
        else:
            print "Make sure you have everything correct."

    except requests.exceptions.ConnectionError, e:
        print "Oops! Connection Error. Try again"

    for link in links:
        try:

            req = requests.get(link)
            if req.status_code == 200:
                print ('Fetching in page links...')
                # print r.status_code
                cont = req.content
                soup1 = BeautifulSoup(cont, "lxml")
                tags1 = soup1('a')

                for a in tags1:
                    href1 = a.get("href")
                    if href1 is not None:
                        url = urlparse.urljoin(link, href1)
                        if url.find("industryCode") != -1:
                            print url
                            all_links.add(url)                     # 'all_links' contains URLs of all the industries


            elif req.status_code == 403:
                print "Error: 403 Forbidden url"
            elif req.status_code == 404:
                print "Error: 404 URL not found"
            else:
                print "Make sure you have everything correct."

        except requests.exceptions.ConnectionError, e:
            print "Oops! Connection Error. Try again"

    # print len(all_links)
    ranking = list()
    for link in all_links:
        try:

            req = requests.get(link)
            if req.status_code == 200:
                print ('Fetching in page links...')
                # print r.status_code
                cont = req.content
                soup1 = BeautifulSoup(cont, "lxml")
                tags1 = soup1('a')

                for a in tags1:
                    href1 = a.get("href")
                    if href1 is not None:
                        url = urlparse.urljoin(link, href1)
                        if url.find("ranking") != -1:
                            print url
                            url += "&page=-1"
                            ranking.append(url)                     # 'ranking' has URLs of companies across all the industries


            elif req.status_code == 403:
                print "Error: 403 Forbidden url"
            elif req.status_code == 404:
                print "Error: 404 URL not found"
            else:
                print "Make sure you have everything correct."

        except requests.exceptions.ConnectionError, e:
            print "Oops! Connection Error. Try again"

    company = list()
    company_info = dict()
    def_url = "https://www.reuters.com/finance/stocks/company-officers/"
    r = len(ranking)
    wb = open_workbook('C_TR_Business_Classification_Index.xlsx')           #opening excel sheet for mapping permID and hierarchicalID with industries
    list1 = []                                                              #for storing industry
    list2 = []                                                              #for storing permID
    list3 = []                                                              #for storing hierarchicalID
    diciti={}                                                               #for mapping industries with permId and hierarchicalID
    for sheets in wb.sheets():
        # for col in range(sheets.ncols):
        for rows in range(1, sheets.nrows):
            if str(sheets.cell(rows, 4).value).strip('"') != '':
             list1.append(sheets.cell(rows, 3).value)
             list2.append(int(sheets.cell(rows, 4).value))
             list3.append(int(sheets.cell(rows, 5).value))
    length_of_indus = len(list1)
    for row in range(0, length_of_indus):
        diciti[list1[row]] = [list2[row],list3[row]]

    for link in ranking:
        r -= 1
        print "------------------------------"
        print link
        print r
        response = requests.get(link)
        paragraphs = justext.justext(response.content, justext.get_stoplist("English"))
        response = urllib2.urlopen(link).read()
        soup = BeautifulSoup(response, "lxml")
        name_box = soup.find_all("div", id="sectionTitle")
        # title = name_box[0].h1.text
        if name_box:
            title = name_box[0].h1.text

            i = -1
            j = 0

            for paragraph in paragraphs:
                i += 1
                if paragraph.text == "MktCap Weighted Average":
                    j = i
                    break

            j += 4
            while j:

                if paragraphs[j].text == "1":
                    break

                lis = list()
                lis.append(paragraphs[j].text)
                lis.append(paragraphs[j + 2].text)
                lis.append(paragraphs[j + 3].text)
                lis.append(paragraphs[j + 4].text)
                lis.append(title)
                if(diciti.has_key(title)):
                    lis.append(diciti[title][0])
                    lis.append(diciti[title][1])
                else:
                    lis.append("NA")
                    lis.append("NA")
                company_info[paragraphs[j + 1].text] = lis

                # 'company_info' is a dictionary which maps Company Names to its Ticker, Market Capitalization, TTM Sales$ and Employees.

                company_url = def_url + paragraphs[j].text
                if company_url != "http://www.reuters.com/finance/stocks/overview/CRC.N":
                    company.append(company_url)


                '''print paragraphs[j].text
                print paragraphs[j+1].text
                print paragraphs[j+2].text
                print paragraphs[j+3].text
                print paragraphs[j+4].text
                print title'''
                j += 5

    print "-----------------DONE-----------------------------------------"
    convert_to_html(company_info)
    employee = dict()
    r = len(company)

    for link in company:
        print "###########################"
        print r
        print link
        r -= 1
        response = requests.get(link).content

        soup = BeautifulSoup(response, "lxml")
        table_data = soup.findAll('table', attrs={'class': 'dataTable'})

        qTables = []
        for datum in table_data:
            qTable = Table()
            qTable.populate(datum)
            qTables.append(qTable)

        description_table = None
        person_data_table = None

        for qTable in qTables:
            if qTable.type == 'Description Table':
                description_table = qTable
            elif qTable.type == 'Person Data Table':
                person_data_table = qTable

        person_list = []

        if (description_table is not None) or (person_data_table is not None):
            for name in person_data_table.data.keys():
                person_data = person_data_table.data[name]
                person_description = description_table.data[name][1]

                person = Person(name=person_data[0],
                                age=person_data[1],
                                since=person_data[2],
                                current_position=person_data[3],
                                description=person_description)

                person_list.append(person)

        print "PERSON LIST"
        convert_to_html2(person_list)
        print person_list
