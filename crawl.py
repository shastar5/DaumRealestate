from bs4 import BeautifulSoup
from urllib.request import urlopen
import xlsxwriter

# 아파트 이름, 위치
def title(key):
    titleurl = 'http://realestate.daum.net/maemul/danji/' + key.__str__() + '/A1A3A4/S/maemulList#t:DanjiInfo&c:A1&s:S'
    titlesoup = BeautifulSoup(urlopen(titleurl), from_encoding='utf-8')
    loc = titlesoup.title.contents[0]
    if(loc[0:2] == '서울'):
        print(loc[0:2])
        return 1
    if(loc[3:5] == '전주'):
        print('전주')
        return 2
    if(loc[3:5] == '완주'):
        print('완주')
        return 3

# 단지 정보
def danji_info(key):
    infourl = 'http://realestate.daum.net/iframe/maemul/DanjiInfo.daum?danjiId=' + key.__str__() + '&mcateCode=A1A3A4&saleTypeCode=S&tabName=info'
    infosoup = BeautifulSoup(urlopen(infourl), from_encoding='utf-8')
    danji_info = []
    for hit in infosoup.find_all('span', attrs={'class':['desc_info', 'tit_info']}):
        print(hit.contents[0].strip())
        danji_info.append(hit.contents[0].strip())
    return danji_info

# 주변 정보
def near_info(key):
    infourl = 'http://realestate.daum.net/iframe/maemul/DanjiInfo.daum?danjiId=' + key.__str__() + '&mcateCode=A1A3A4&saleTypeCode=S&tabName=info'
    infosoup = BeautifulSoup(urlopen(infourl), from_encoding='utf-8')
    near_info = []
    for hit in infosoup.find_all('div', {'id': 'colSurrounding'}):
        for row in hit.findAll('dd'):
            near_info.append(row.text.strip())
            print(row.text.strip())
    return near_info

# 시세
def price_info(key):
    priceurl = 'http://realestate.daum.net/iframe/maemul/DanjiSise.daum?danjiId=' + key.__str__() + '&mcateCode=A1A3A4&saleTypeCode=S&tabName=sise&ptype='
    pricesoup = BeautifulSoup(urlopen(priceurl), from_encoding='utf-8')
    price = []
    table = pricesoup.find('table', {'class':'tbl'})
    for row in table.findAll('tbody'):
        col = row.find_all('td')
        # 10개의 칼럼을 가지고 있음.
        for x in col:
            print(x.string.strip())
            price.append(x.string.strip())
    return price

indexnum = 11719

# Open and create xlsx file
workbook = xlsxwriter.Workbook('data.xlsx')
sheet = []
sheet[0] = workbook.add_worksheet('서울')
sheet[1] = workbook.add_worksheet('전주')
sheet[2] = workbook.add_worksheet('완주')

# Write some data headers.
bold = workbook.add_format({'bold': 1})
for x in range[0:2]:
    # 개요 정보
    sheet[x].write('B1', '주소', bold)
    sheet[x].write('C1', '총세대수', bold)
    sheet[x].write('D1', '총동수', bold)
    sheet[x].write('E1', '준공년월', bold)
    sheet[x].write('F1', '입주년월', bold)
    sheet[x].write('G1', '건설사명', bold)
    sheet[x].write('H1', '최저/최고층', bold)
    sheet[x].write('I1', '난방방식', bold)
    sheet[x].write('J1', '난방연료', bold)
    sheet[x].write('K1', '용적율', bold)
    sheet[x].write('L1', '건폐율', bold)

    # 주변 시설
    sheet[x].write('M1', '지하철', bold)
    sheet[x].write('N1', '버스', bold)
    sheet[x].write('O1', '도로시설', bold)
    sheet[x].write('P1', '공원시설', bold)
    sheet[x].write('Q1', '편의시설', bold)
    sheet[x].write('R1', '교육시설', bold)
    sheet[x].write('S1', '의료시설', bold)


loc = title(indexnum)
danji_info(indexnum)
near_info(indexnum)
price_info(indexnum)

# 서울
if(loc == 1):
    print(2)

# 전주
elif(loc == 2):
    print(2)

# 완주
elif(loc == 3):
    print(2)


workbook.close()