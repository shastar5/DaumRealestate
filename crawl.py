from bs4 import BeautifulSoup
from urllib.request import urlopen
import lxml.html

# Basic
key = 11719

titleurl = 'http://realestate.daum.net/maemul/danji/' + key.__str__() + '/A1A3A4/S/maemulList#t:DanjiInfo&c:A1&s:S'
titlesoup = BeautifulSoup(urlopen(titleurl), from_encoding='utf-8')

infourl = 'http://realestate.daum.net/iframe/maemul/DanjiInfo.daum?danjiId=' + key.__str__() + '&mcateCode=A1A3A4&saleTypeCode=S&tabName=info'
infosoup = BeautifulSoup(urlopen(infourl), from_encoding='utf-8')

priceurl = 'http://realestate.daum.net/iframe/maemul/DanjiSise.daum?danjiId=' + key.__str__() + '&mcateCode=A1A3A4&saleTypeCode=S&tabName=sise&ptype='
pricesoup = BeautifulSoup(urlopen(priceurl), from_encoding='utf-8')

# 아파트 이름, 위치
City = 0
loc = titlesoup.title.contents[0]
if(loc[0:2] == '서울'):
    print(loc[0:2])
    City = 1
if(loc[3:5] == '전주'):
    print('전주시임')
    City = 2
if(loc[3:5] == '완주'):
    print('완주시임')
    City = 3


# 단지 정보
for hit in infosoup.find_all('span', attrs={'class':['desc_info', 'tit_info']}):
    print(hit.contents[0].strip())

for hit in infosoup.find_all('div', {'id': 'colSurrounding'}):
    print(hit)



# 시세
table = pricesoup.find('table', {'class':'tbl'})
for row in table.findAll('tbody'):
    col = row.find_all('td')
    # 10개의 칼럼을 가지고 있음.
    for x in col:
        print(x.string.strip())