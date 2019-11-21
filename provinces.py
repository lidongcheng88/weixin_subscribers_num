import time
import asyncio

import xlwt
from pathlib import Path
from pprint import pprint
from pyppeteer import launch


async def main(sheet, fname='数据表'):
    browser = await launch(
        headless=False, userDataDir='./userdata', args=['--disable-infobars'])
    page = await browser.newPage()
    await page.setViewport(viewport={'width': 1380, 'height': 800})
    await page.goto('https://mp.weixin.qq.com')
    await page.click('a[title=用户分析]')
    time.sleep(2)
    await page.click('a[title=用户属性]')
    time.sleep(2)
    #   page.click('div.weui-desktop-layout__main__hd>div:nth-child(3) li:nth-child(2)')
    diyu = await page.xpath('//a[contains(text(), "地域")]')
    await diyu[0].click()
    tbodys = await page.xpath('//table/tbody')
    # 省份数据
    tbodys[3]
    provinces = {}
    while True:
        trs1 = await tbodys[3].xpath('./tr')
        for tr1 in trs1:
            tds1 = await tr1.xpath('./td')
            province_name = await (
                await tds1[0].getProperty("textContent")).jsonValue()
            provinces[province_name] = await (
                await tds1[1].getProperty("textContent")).jsonValue()
        next_selectors = await page.xpath('//a[contains(text(), "下一页")]')
        try:
            await next_selectors[0].click()
        except Exception as e:
            print(e)
            break
    # await page_datas(tbodys[3], provinces, page, 0)
    pprint(provinces)
    print('省份数据采集成功！')
    # 导出表格
    # 创建工作簿
    workbook = xlwt.Workbook(encoding='utf-8')
    # 省份
    data_sheet1 = workbook.add_sheet(sheet)
    sheet_header1 = ['省份', '用户数']
    for index, value in enumerate(sheet_header1):
        data_sheet1.write(0, index, value)
    i = 0
    for province, subscribers_num in provinces.items():
        i += 1
        data_sheet1.write(i, 0, province)
        data_sheet1.write(i, 1, subscribers_num)
    print('省份数据输出成功！')
    # 保存到：当前路径/fname.xls
    workbook.save(str(Path.cwd() / fname) + '.xls')
    print('文件创建成功')
    await browser.close()


if __name__ == '__main__':
    sheet = '省份数据'
    # path = 'C:/Users/Administrator/Desktop'
    fname = '公众号粉丝数据省份表'
    asyncio.get_event_loop().run_until_complete(main(sheet, fname=fname))
