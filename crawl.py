from bs4 import BeautifulSoup
from urllib.request import urlopen
import xlsxwriter
from time import sleep

fname = '0-500000.xlsx'

# 아파트 이름, 위치
def title(key):
    titleurl = 'http://realestate.daum.net/maemul/danji/' + key.__str__() + '/A1A3A4/S/maemulList#t:DanjiInfo&c:A1&s:S'
    titlesoup = BeautifulSoup(urlopen(titleurl), from_encoding='utf-8')

    try:
        if (titlesoup.title == None):
            return 0
        loc = titlesoup.title.contents[0]
        if (loc[0:2] == '경기'):
            # print(loc[0:2])
            return 1
        """
        if (loc[3:5] == '전주'):
            # print('전주')
            return 2
        if (loc[3:5] == '완주'):
            # print('완주')
            return 3
        """
    except Exception as e:
        return None


# 단지 정보
def danji_info(key):
    infourl = 'http://realestate.daum.net/iframe/maemul/DanjiInfo.daum?danjiId=' + key.__str__() + '&mcateCode=A1A3A4&saleTypeCode=S&tabName=info'
    try:
        infosoup = BeautifulSoup(urlopen(infourl), from_encoding='utf-8')
    except Exception as e:
        return None

    danji_info = []
    titlename = infosoup.find_all('h3', {'class': 'fl_le fs_big'})
    for hit in titlename:
        danji_info.append(hit.text)
    for hit in infosoup.find_all('span', attrs={'class': ['desc_info', 'tit_info', 'ico_realestate_v1 address_number']}):
        # print(hit.contents[0].strip())
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
            # print(row.text.strip())
    return near_info


isKB = False
only114 = False


# 시세
def price_info(key):
    global isKB
    global only114

    only114 = True
    isKB = False
    priceurl = 'http://realestate.daum.net/iframe/maemul/DanjiSise.daum?danjiId=' + key.__str__() + '&mcateCode=A1A3A4&saleTypeCode=S&tabName=sise&ptype='
    try:
        pricesoup = BeautifulSoup(urlopen(priceurl), from_encoding='utf-8')
        price = []

        table = pricesoup.find('table', {'class': 'tbl'})

        # KB에서 제공하는지 114에서 제공하는지 검사
        dataSource = pricesoup.find_all('span', {'class': 'desc fR'})
        for hit in dataSource:
            if (hit.text.find('KB') != -1):
                isKB = True
            if '국토' in hit.text:
                only114 = False
        for row in table.findAll('tbody'):
            col = row.find_all('td')
            # 10개의 칼럼을 가지고 있음.
            for x in col:
                # print(x.string.strip())
                price.append(x.string.strip())
        return price
    except Exception as e:
        return None


# Open and create xlsx file
workbook = xlsxwriter.Workbook(fname)
sheet = [workbook.add_worksheet('경기')]

format = workbook.add_format()
format.set_align('center')
format.set_align('vcenter')
format.set_bold(True)

# Declare excel merging format.
merge_format = workbook.add_format({'bold': 1,
                                    'border': 1,
                                    'align': 'center',
                                    'valign': 'center'})

# Write some data headers.
for x in range(1):
    # 개요 정보
    sheet[x].write('A1', '아파트이름', format)
    sheet[x].write('B1', '도로명 주소', format)
    sheet[x].write('C1', '지번 주소', format)

    sheet[x].write('D1', '총세대수', format)
    sheet[x].write('E1', '총동수', format)
    sheet[x].write('F1', '준공년월', format)
    sheet[x].write('G1', '입주년월', format)
    sheet[x].write('H1', '건설사명', format)
    sheet[x].write('I1', '최저/최고층', format)
    sheet[x].write('J1', '총 주차대수', format)
    sheet[x].write('K1', '난방방식', format)
    sheet[x].write('L1', '난방연료', format)
    sheet[x].write('M1', '용적율', format)
    sheet[x].write('N1', '건폐율', format)

    # 주변 시설
    sheet[x].write('O1', '지하철', format)
    sheet[x].write('P1', '버스', format)
    sheet[x].write('Q1', '도로시설', format)
    sheet[x].write('R1', '공원시설', format)
    sheet[x].write('S1', '편의시설', format)
    sheet[x].write('T1', '교육시설', format)
    sheet[x].write('U1', '의료시설', format)

    # 아파트 시세
    sheet[x].merge_range('V1:V3', '면적', merge_format)

    # Merge cell
    sheet[x].merge_range('W1:AA1', '매매', merge_format)
    sheet[x].merge_range('AB1:AF1', '전세', merge_format)
    sheet[x].merge_range('W2:X2', '부동산114', merge_format)
    sheet[x].merge_range('Y2:AA2', '실거래가', merge_format)
    sheet[x].merge_range('AB2:AC2', '부동산114', merge_format)
    sheet[x].merge_range('AD2:AF2', '실거래가', merge_format)
    sheet[x].merge_range('AG1:AI2', '매매', merge_format)
    sheet[x].merge_range('AJ1:AL2', '전세', merge_format)

    # 부동산 정보
    sheet[x].write('W3', '최고가', format)
    sheet[x].write('X3', '최저가', format)
    sheet[x].write('Y3', '최고가', format)
    sheet[x].write('Z3', '최저가', format)
    sheet[x].write('AA3', '거래건수', format)
    sheet[x].write('AB3', '최저가', format)
    sheet[x].write('AC3', '최고가', format)
    sheet[x].write('AD3', '최저가', format)
    sheet[x].write('AE3', '최고가', format)
    sheet[x].write('AF3', '거래건수', format)
    sheet[x].write('AG3', '하위평균가', format)
    sheet[x].write('AH3', '일반평균가', format)
    sheet[x].write('AI3', '상위평균가', format)
    sheet[x].write('AJ3', '하위평균가', format)
    sheet[x].write('AK3', '일반평균가', format)
    sheet[x].write('AL3', '상위평균가', format)

    # Check index number
    sheet[x].write('AM1', 'Index', format)

    # Set Column Size
    sheet[x].set_column('A:B', 20)
    sheet[x].set_column('C:M', 20)
    sheet[x].set_column('N:T', 50)

seoulrow = 3
jeonjurow = 3
wanjurow = 3

format.set_bold(False)


def crawl(indexnum):
    global seoulrow
    global jeonjurow
    global wanjurow

    loc = title(indexnum)
    if (loc == None):
        return
    danji = danji_info(indexnum)
    near = near_info(indexnum)
    price = price_info(indexnum)

    if (price != None):
        numofPrice = len(price)
    iteration = 0
    # 서울/경기
    if (loc == 1):
        sheet[0].write(seoulrow, 38, indexnum, format)
        if (price != None):
            # 정상적인 경우
            if (isKB == False and only114 == False):
                while (iteration <= numofPrice):
                    for col in range(0, 14):
                        sheet[0].write(seoulrow, col, danji[col], format)
                    for col in range(14, 21):
                        sheet[0].write(seoulrow, col, near[col - 14], format)
                    for col in range(21, 32):
                        sheet[0].write(seoulrow, col, price[iteration], format)
                        iteration = iteration + 1
                    numofPrice = numofPrice - 11
                    seoulrow += 1

            # 114정보만 있을때
            # Number of cols == 5
            elif (isKB == False and only114 == True):
                while (True):
                    for col in range(0, 14):
                        sheet[0].write(seoulrow, col, danji[col], format)
                    for col in range(14, 21):
                        sheet[0].write(seoulrow, col, near[col - 14], format)
                    sheet[0].write(seoulrow, 21, price[iteration], format)
                    iteration = iteration + 1
                    sheet[0].write(seoulrow, 22, price[iteration], format)
                    iteration = iteration + 1
                    sheet[0].write(seoulrow, 23, price[iteration], format)
                    iteration = iteration + 1
                    sheet[0].write(seoulrow, 27, price[iteration], format)
                    iteration = iteration + 1
                    sheet[0].write(seoulrow, 28, price[iteration], format)
                    iteration = iteration + 1
                    seoulrow += 1
                    if (iteration >= numofPrice):
                        break

            # KB에서 제공할 경우
            else:
                while (iteration <= numofPrice):
                    for col in range(0, 14):
                        sheet[0].write(seoulrow, col, danji[col], format)
                    for col in range(14, 21):
                        sheet[0].write(seoulrow, col, near[col - 14], format)
                    sheet[0].write(seoulrow, 21, price[iteration], format)
                    iteration = iteration + 1
                    for col in range(32, 38):
                        sheet[0].write(seoulrow, col, price[iteration], format)
                        iteration = iteration + 1
                    numofPrice = numofPrice - 7
                    seoulrow += 1

        else:
            for col in range(0, 14):
                sheet[0].write(seoulrow, col, danji[col], format)
            for col in range(14, 21):
                sheet[0].write(seoulrow, col, near[col - 14], format)
            seoulrow += 1

    # 전주
    if (loc == 2):
        sheet[1].write(jeonjurow, 38, indexnum, format)
        if (price != None):
            # 정상적인 경우
            if (isKB == False and only114 == False):
                while (iteration <= numofPrice):
                    for col in range(0, 14):
                        sheet[1].write(jeonjurow, col, danji[col], format)
                    for col in range(14, 21):
                        sheet[1].write(jeonjurow, col, near[col - 14], format)
                    for col in range(21, 32):
                        sheet[1].write(jeonjurow, col, price[iteration], format)
                        iteration = iteration + 1
                    numofPrice = numofPrice - 11
                    jeonjurow += 1

            # 114정보만 있을때
            # Number of cols == 5
            elif (isKB == False and only114 == True):
                while (True):
                    for col in range(0, 14):
                        sheet[1].write(jeonjurow, col, danji[col], format)
                    for col in range(14, 21):
                        sheet[1].write(jeonjurow, col, near[col - 14], format)
                    sheet[1].write(jeonjurow, 21, price[iteration], format)
                    iteration = iteration + 1
                    sheet[1].write(jeonjurow, 22, price[iteration], format)
                    iteration = iteration + 1
                    sheet[1].write(jeonjurow, 23, price[iteration], format)
                    iteration = iteration + 1
                    sheet[1].write(jeonjurow, 27, price[iteration], format)
                    iteration = iteration + 1
                    sheet[1].write(jeonjurow, 28, price[iteration], format)
                    iteration = iteration + 1
                    jeonjurow += 1
                    if (iteration >= numofPrice):
                        break

            # KB에서 제공할 경우
            else:
                while (iteration <= numofPrice):
                    for col in range(0, 14):
                        sheet[1].write(jeonjurow, col, danji[col], format)
                    for col in range(14, 21):
                        sheet[1].write(jeonjurow, col, near[col - 14], format)
                    sheet[1].write(jeonjurow, 21, price[iteration], format)
                    iteration = iteration + 1
                    for col in range(32, 38):
                        sheet[1].write(jeonjurow, col, price[iteration], format)
                        iteration = iteration + 1
                    numofPrice = numofPrice - 7
                    jeonjurow += 1

        else:
            for col in range(0, 14):
                sheet[1].write(jeonjurow, col, danji[col], format)
            for col in range(14, 21):
                sheet[1].write(jeonjurow, col, near[col - 14], format)
            jeonjurow += 1

    # 완주
    elif (loc == 3):
        sheet[2].write(wanjurow, 38, indexnum, format)
        if (price != None):
            # 정상적인 경우
            if (isKB == False and only114 == False):
                while (iteration <= numofPrice):
                    for col in range(0, 14):
                        sheet[2].write(wanjurow, col, danji[col], format)
                    for col in range(14, 21):
                        sheet[2].write(wanjurow, col, near[col - 14], format)
                    for col in range(21, 32):
                        sheet[2].write(wanjurow, col, price[iteration], format)
                        iteration = iteration + 1
                    numofPrice = numofPrice - 11
                    wanjurow += 1

            # 114정보만 있을때
            # Number of cols == 5
            elif (isKB == False and only114 == True):
                while (True):
                    for col in range(0, 14):
                        sheet[2].write(wanjurow, col, danji[col], format)
                    for col in range(14, 21):
                        sheet[2].write(wanjurow, col, near[col - 14], format)
                    sheet[2].write(wanjurow, 21, price[iteration], format)
                    iteration = iteration + 1
                    sheet[2].write(wanjurow, 22, price[iteration], format)
                    iteration = iteration + 1
                    sheet[2].write(wanjurow, 23, price[iteration], format)
                    iteration = iteration + 1
                    sheet[2].write(wanjurow, 27, price[iteration], format)
                    iteration = iteration + 1
                    sheet[2].write(wanjurow, 28, price[iteration], format)
                    iteration = iteration + 1
                    wanjurow += 1
                    if (iteration >= numofPrice):
                        break

            # KB에서 제공할 경우
            else:
                while (iteration <= numofPrice):
                    for col in range(0, 14):
                        sheet[2].write(wanjurow, col, danji[col], format)
                    for col in range(14, 21):
                        sheet[2].write(wanjurow, col, near[col - 14], format)
                    sheet[2].write(wanjurow, 21, price[iteration], format)
                    iteration = iteration + 1
                    for col in range(32, 38):
                        sheet[2].write(wanjurow, col, price[iteration], format)
                        iteration = iteration + 1
                    numofPrice = numofPrice - 7
                    wanjurow += 1

        else:
            for col in range(0, 14):
                sheet[2].write(wanjurow, col, danji[col], format)
            for col in range(14, 21):
                sheet[2].write(wanjurow, col, near[col - 14], format)
            wanjurow += 1

error = []

def run(idx, idx2):
    for x in range(idx, idx2):
        try:
            if x % 100 == 0:
                print(x)
            crawl(x)
        except Exception as e:
            print(e)
            print(x)
            error.append(x)
            sleep(10)
            run(x, idx2)
            break

        if x % 3000 == 0:
            print(x)
            sleep(10)

run(0, 100)


print(error)


workbook.close()