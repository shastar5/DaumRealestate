from bs4 import BeautifulSoup
from urllib.request import urlopen
import xlsxwriter

# 아파트 이름, 위치
def title(key):
    titleurl = 'http://realestate.daum.net/maemul/danji/' + key.__str__() + '/A1A3A4/S/maemulList#t:DanjiInfo&c:A1&s:S'
    titlesoup = BeautifulSoup(urlopen(titleurl), from_encoding='utf-8')
    if(titlesoup.title == None):
        return 0
    loc = titlesoup.title.contents[0]
    if(loc[0:2] == '서울'):
        #print(loc[0:2])
        return 1
    if(loc[3:5] == '전주'):
        #print('전주')
        return 2
    if(loc[3:5] == '완주'):
        #print('완주')
        return 3

# 단지 정보
def danji_info(key):
    infourl = 'http://realestate.daum.net/iframe/maemul/DanjiInfo.daum?danjiId=' + key.__str__() + '&mcateCode=A1A3A4&saleTypeCode=S&tabName=info'
    try:
        infosoup = BeautifulSoup(urlopen(infourl), from_encoding='utf-8')
    except Exception as e:
        return None
    danji_info = []
    for hit in infosoup.find_all('span', attrs={'class':['desc_info', 'tit_info']}):
        print(hit.contents[0].strip())
        danji_info.append(hit.contents[0].strip())
    return danji_info

# 주변 정보
def near_info(key):
    infourl = 'http://realestate.daum.net/iframe/maemul/DanjiInfo.daum?danjiId=' + key.__str__() + '&mcateCode=A1A3A4&saleTypeCode=S&tabName=info'
    try:
        infosoup = BeautifulSoup(urlopen(infourl), from_encoding='utf-8')
    except Exception as e:
        return None
    near_info = []
    for hit in infosoup.find_all('div', {'id': 'colSurrounding'}):
        for row in hit.findAll('dd'):
            near_info.append(row.text.strip())
            #print(row.text.strip())
    return near_info

# 시세
def price_info(key):
    priceurl = 'http://realestate.daum.net/iframe/maemul/DanjiSise.daum?danjiId=' + key.__str__() + '&mcateCode=A1A3A4&saleTypeCode=S&tabName=sise&ptype='
    try:
        pricesoup = BeautifulSoup(urlopen(priceurl), from_encoding='utf-8')

        price = []
        table = pricesoup.find('table', {'class':'tbl'})
        for row in table.findAll('tbody'):
            col = row.find_all('td')
            # 10개의 칼럼을 가지고 있음.
            for x in col:
                #print(x.string.strip())
                price.append(x.string.strip())

    except Exception as e:
        return None

    return price


# Open and create xlsx file
workbook = xlsxwriter.Workbook('data.xlsx')
sheet = [workbook.add_worksheet('서울'), workbook.add_worksheet('전주'), workbook.add_worksheet('완주')]

# Declare excel merging format.
merge_format = workbook.add_format({'bold': 1,
                                    'border': 1,
                                    'align': 'center',
                                    'valign': 'center'})

# Write some data headers.
bold = workbook.add_format({'bold': 1})
for x in range(3):
    # 개요 정보
    sheet[x].write('A1', '아파트이름', bold)
    sheet[x].write('B1', '주소', bold)
    sheet[x].write('C1', '총세대수', bold)
    sheet[x].write('D1', '총동수', bold)
    sheet[x].write('E1', '준공년월', bold)
    sheet[x].write('F1', '입주년월', bold)
    sheet[x].write('G1', '건설사명', bold)
    sheet[x].write('H1', '최저/최고층', bold)
    sheet[x].write('I1', '총 주차대수', bold)
    sheet[x].write('J1', '난방방식', bold)
    sheet[x].write('K1', '난방연료', bold)
    sheet[x].write('L1', '용적율', bold)
    sheet[x].write('M1', '건폐율', bold)

    # 주변 시설
    sheet[x].write('N1', '지하철', bold)
    sheet[x].write('O1', '버스', bold)
    sheet[x].write('P1', '도로시설', bold)
    sheet[x].write('Q1', '공원시설', bold)
    sheet[x].write('R1', '편의시설', bold)
    sheet[x].write('S1', '교육시설', bold)
    sheet[x].write('T1', '의료시설', bold)

    # 아파트 시세
    sheet[x].merge_range('U1:U3', '면적', merge_format)
    # Merge cell
    sheet[x].merge_range('V1:Z1', '매매', merge_format)
    sheet[x].merge_range('AA1:AE1', '전세', merge_format)
    sheet[x].merge_range('V2:W2', '부동산114', merge_format)
    sheet[x].merge_range('X2:Z2', '실거래가', merge_format)
    sheet[x].merge_range('AA2:AB2', '부동산114', merge_format)
    sheet[x].merge_range('AC2:AE2', '실거래가', merge_format)
    sheet[x].write('V3', '최고가', bold)
    sheet[x].write('W3', '최저가', bold)
    sheet[x].write('X3', '최고가', bold)
    sheet[x].write('Y3', '최저가', bold)
    sheet[x].write('Z3', '거래건수', bold)
    sheet[x].write('AA3', '최고가', bold)
    sheet[x].write('AB3', '최저가', bold)
    sheet[x].write('AC3', '최고가', bold)
    sheet[x].write('AD3', '최저가', bold)
    sheet[x].write('AE3', '거래건수', bold)

workbook.close()
seoulrow = 3
jeonjurow = 3
wanjurow = 3

workbook = xlsxwriter.Workbook('data.xlsx')
def crawl(indexnum):
    global seoulrow
    global jeonjurow
    global wanjurow

    loc = title(indexnum)
    danji = danji_info(indexnum)
    near = near_info(indexnum)
    price = price_info(indexnum)

    workbook = xlsxwriter.Workbook('data.xlsx')

    # 서울
    try:
        if(loc == 1):
            for col in range(1, 13):
                sheet[0].write(seoulrow, col, danji[col-1])
            for col in range(13, 20):
                sheet[0].write(seoulrow, col, near[col-13])
            for col in range(20, 31):
                sheet[0].write(seoulrow, col, price[col-20])
            seoulrow += 1
            workbook.close()

        # 전주
        elif(loc == 2):
            for col in range(1, 12):
                sheet[1].write(seoulrow, col, danji[col-1])
            for col in range(12, 19):
                sheet[1].write(seoulrow, col, near[col-12])
            for col in range(19, 30):
                sheet[1].write(seoulrow, col, price[col-19])
            jeonjurow += 1
            workbook.close()

        # 완주
        elif(loc == 3):
            for col in range(1, 12):
                sheet[2].write(seoulrow, col, danji[col-1])
            for col in range(12, 19):
                sheet[2].write(seoulrow, col, near[col-12])
            for col in range(19, 30):
                sheet[2].write(seoulrow, col, price[col-19])
            wanjurow += 1
            workbook.close()

    except Exception as e:
        return

    workbook.close()

for x in range(0, 9000):
    indexnum = 1000 + x
    crawl(indexnum)

workbook.close()