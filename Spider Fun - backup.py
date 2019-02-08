from bs4 import BeautifulSoup
import requests
import xlsxwriter

# read SR from txt file
def getSR(path):

    with open(path, 'r') as f:
        SRlist = []
        for sr in f.readlines():
            sr=sr.strip('\n')
            SRlist.append(sr)
    # print(f.read())
    print(SRlist)
    return (SRlist)


def basicInfo(soup):
    tt = soup.select('body>table')
    uu = tt[0].select('tr>td')
    vv = uu[0].select('tr>td')
    ww = vv[0].select('tr')
    name = (ww[0].text).strip()
    company = (ww[1].text).strip()
    level=(ww[2].text).strip()
    level=level[14:]
    basic=[name, company, level]

    print(name)
    print(company)
    print(level)

    return(basic)


def getSRInfo(soup,sr):
    sr_location = soup.find_all('b', text=sr)
    sp= sr_location[0].find_parent("td")
    status = sp.find_next_sibling("td")
    date = status.find_next_sibling("td").find_next_sibling("td")
    owner = date.find_next_sibling("td").find_next_sibling("td")
    summary = owner.find_next_sibling("td").find_next_sibling("td")

    status = status.string
    date = date.text
    owner = owner.text
    summary = summary.string
    print(status)
    print(date)
    print(owner)
    print(summary)
    SRinfo=[summary,status,owner,date]
    return SRinfo


def findSale(soup):

    TSR=soup.find_all('b',text='TSR')
    TSR=TSR[0].find_parent("td")
    TSR=TSR.find_next_sibling("td")
    salesman=TSR.string
    print(salesman)
    return salesman

# get web soup from the sr-specified link
def getSoup(url):
    r = requests.get(url)
    if r.status_code == requests.codes.ok:
        web = r.text
        soup = BeautifulSoup(web, features="html.parser")
        print("Status:OK")
        return soup

def write2file(worksheet,data,row):
    for col in list(range(len(data)-1)):
        if col == 0:
            worksheet.write_url(row,col,data[-1],string=data[0])
        else:
            worksheet.write(row,col,data[col])
    return

def keywordCheck(key,data): ####Return the data of elements containing keywords in key.
    result = list()
    for i in data:
        for j in key:
            if j in i[2]:
                result.append(i)
                print(result)
    return result
# Main file

urlroot="http://force.natinst.com:8000/pls/ebiz/niae_screenpop.main?p_incident_number="
path = 'SR.txt'
srlist=getSR(path)

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('Test2.xlsx')
worksheet = workbook.add_worksheet('All SR')
worksheet_auto = workbook.add_worksheet('Transportation SR')
worksheet.set_column(2,4,20)
worksheet.set_column(5,9,15)
worksheet_auto.set_column(2, 4, 20)
worksheet_auto.set_column(5,9,15)

bold = workbook.add_format({'bold': True})
Header_Index=['A1','B1','C1','D1','E1','F1','G1','H1','I1']
Header_Content=['SR','Name','Company','Level','Summary','Status','Owner','Date','Sales']

for i in list(range(len(Header_Index))):
    worksheet.write(Header_Index[i],Header_Content[i],bold)
    worksheet_auto.write(Header_Index[i], Header_Content[i], bold)


Data=list()
Keyword = ['北京']

for sr in srlist:

    URL = urlroot+sr
    soup = getSoup(URL)

    # #### search sales name
    basic = basicInfo(soup) #### basic=[name, company, level]
    salesman = findSale(soup)
    SRdata = getSRInfo(soup,sr)  ##### SRinfo=[summary,status,owner,date]
    data = [sr]+basic+SRdata+[salesman]+[URL]  ##['SR','Name','Company','Level','Summary','Status','Owner','Date','Sales']
    Data.append(data)

Data_auto = keywordCheck(Keyword,Data)


#  写入excel
for i,v in enumerate(Data):
    write2file(worksheet,v,i+1) ###The first row is the header

for i,v in enumerate(Data_auto):
    write2file(worksheet_auto,v,i+1) ###The first row is the header

workbook.close()

