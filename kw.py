#coding:utf-8
import requests
import re
import xlsxwriter
import time

#关键词排名
def location1(kw,page=1):
    adlid = '233019116'
    if page==1:
        new_kw = re.sub('\s', '+', kw.strip())
        url = 'https://www.alibaba.com/trade/search?fsb=y&IndexArea=product_en&SearchText=' + kw
        html = requests.get(url).text
        companys = re.findall('"supplierId":"(\d+)",', html)
        if adlid in companys:
            location = [j for j in range(len(companys)) if companys[j] == adlid]
            return '1.%s'% (location[0])
        else:
            return location1(kw,2)
    if page>=2:
        new_kw = re.sub('\s', '_', kw.strip())
        url = 'https://www.alibaba.com/products/F0/%s/%s.html' % (new_kw,page)
        html = requests.get(url).text
        adlid = '233019116'
        companys = re.findall('"supplierId":"(\d+)",', html)
        if adlid in companys:
           location=[j for j in range(len(companys)) if companys[j]==adlid ]
           return '%d.%s'% (page,location[0])
        else:
            page+=1
            new_kw=kw
            new_kw = re.sub('\s', '_', kw.strip())
            url = 'https://www.alibaba.com/products/F0/%s/%s.html' % (new_kw, page)
            html = requests.get(url).text
            adlid = '233019116'
            companys = re.findall('"supplierId":"(\d+)",', html)
            if adlid in companys:
                location = [j for j in range(len(companys)) if companys[j] == adlid]
                return '%d.%s' % (page, location[0])
            else:
                page += 1
                new_kw = re.sub('\s', '_', kw.strip())
                url = 'https://www.alibaba.com/products/F0/%s/%s.html' % (new_kw, page)
                html = requests.get(url).text
                adlid = '233019116'
                companys = re.findall('"supplierId":"(\d+)",', html)
                if adlid in companys:
                    location = [j for j in range(len(companys)) if companys[j] == adlid]
                    return '%d.%s'% (page,location[0])
                else:
                    page += 1
                if page>=4:
                    return '前4页无排名'

#产品发布人
def who(kw,kw_location):
    page_split=kw_location.split('.')
    page=page_split[0]
    product_location=page_split[1]

    if page=='1':
        new_kw = re.sub('\s', '+', kw.strip())
        url = 'https://www.alibaba.com/trade/search?fsb=y&IndexArea=product_en&SearchText=' + kw
    else:
        new_kw = re.sub('\s', '_', kw.strip())
        url = 'https://www.alibaba.com/products/F0/%s/%s.html' % (new_kw, page)
    html = requests.get(url).text
    contacts = re.findall('"productHref":"(.*?)",', html)
    product_href = 'https:' + contacts[int(product_location)].encode('latin').decode('unicode-escape')
    product_html = requests.get(product_href).text
    contact_name = re.search('"contactName":"(.*?)"', product_html)
    name=re.search('(Mr|Ms)\.\s*[\w]+',str(contact_name))
    if name:
        return contact_name.group(1)
    else:
        return 'unknow'

def num(kw):
    new_kw = re.sub('\s', '+', kw.strip())
    url = 'https://www.alibaba.com/trade/search?fsb=y&IndexArea=product_en&SearchText=' + kw
    html = requests.get(url).text
    product_number = re.search('"num":"(.*?)"', html)
    if type(product_number) != None:
        return product_number.group(1)
    else:
        return '0'

with open('kw.txt','r') as f:
    line = f.readline()
    i=0
    name=int(time.time())
    workbook = xlsxwriter.Workbook('%d.xlsx' %(name))
    worksheet = workbook.add_worksheet()
    worksheet.write(0, 0, '关键词')
    worksheet.write(0, 1, '排名')
    worksheet.write(0, 2, '数量')
    worksheet.write(0, 3, '发布人')
    while line:
        try:
            if(line==''):
                continue
            kw_location=location1(line)
            quantity=num(line)
            if kw_location != '前4页无排名':
                contact=who(line,kw_location)
            else:
                contact='unknow'
            print('正在写入第%d个关键词%s,位置%s' %(i+1,line,kw_location))
            worksheet.write(i+1, 0, line)
            worksheet.write(i+1, 1, kw_location)
            worksheet.write(i+1, 2, quantity)
            worksheet.write(i+1, 3, contact)
            i+=1
            # print(i)
            # print(line)
            # print(kw_location)
            # print(quantity)
            # print(contact)
            # if i>=10:
            #     break
            # print(i)
            # if i>20:
            #     break
            line = f.readline()
        except:
            workbook.close()
    workbook.close()
