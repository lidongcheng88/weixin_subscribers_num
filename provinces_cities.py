import time
import asyncio

import xlwt
from pathlib import Path
from pprint import pprint
from pyppeteer import launch


async def main(sheet, fname, path=None):
    browser = await launch(
        headless=False, userDataDir='./userdata', args=['--disable-infobars'])
    page = await browser.newPage()
    await page.setViewport(viewport={'width': 1380, 'height': 800})
    await page.goto('https://mp.weixin.qq.com')
    await page.click('a[title=用户分析]')
    time.sleep(2)
    await page.click('a[title=用户属性]')
    time.sleep(2)
    diyu = await page.xpath('//a[contains(text(), "地域")]')
    await diyu[0].click()
    tbodys = await page.xpath('//table/tbody')
    # 城市数据
    province_cities = {}
    while True:
        cities = {}
        while True:
            trs = await tbodys[4].xpath('./tr')
            for tr in trs:
                tds = await tr.xpath('./td')
                city_name = await (
                    await tds[0].getProperty("textContent")).jsonValue()
                cities[city_name] = await (
                    await tds[1].getProperty("textContent")).jsonValue()
            next_selectors = await page.xpath('//a[contains(text(), "下一页")]')
            try:
                await next_selectors[1].click()
            except Exception as e:
                print(e)
                break
        # print(cities)
        province = await page.querySelector('dt.weui-desktop-form__dropdown__dt')
        province_name = await (
            await province.getProperty("textContent")).jsonValue()
        province_cities[province_name.strip()] = cities
        await province.click()
        next_province = await page.xpath(
            '//li[contains(@class, "checked")]/following-sibling::li[1]')
        if next_province:
            await next_province[0].click()
        else:
            break
    pprint(province_cities)
    print('城市数据采集成功！')
    print('数据采集成功，浏览器关闭')
    await browser.close()
    # 导出表格
    # 创建工作簿
    workbook = xlwt.Workbook(encoding='utf-8')
    # 创建sheet
    data_sheet = workbook.add_sheet(sheet)
    # 表头
    sheet_header = ['省份', '城市', '用户数']
    for index, value in enumerate(sheet_header):
        data_sheet.write(0, index, value)
    # 数据
    i = 0
    for province, cities in province_cities.items():
        for city, subscribers_num in cities.items():
            i += 1
            data_sheet.write(i, 0, province)
            data_sheet.write(i, 1, city)
            data_sheet.write(i, 2, subscribers_num)
    print('城市数据输出成功')
    # 默认：当前路径/fname.xls  自定义： path/fname
    path = (f'{path}/{fname}' if path else str(Path.cwd() / fname)) + '.xls'
    workbook.save(path)
    print('文件创建成功')
    await browser.close()


if __name__ == '__main__':
    sheet = '城市数据'
    # path = 'C:/Users/Administrator/Desktop'
    fname = '公众号粉丝数据表'
    asyncio.get_event_loop().run_until_complete(main(sheet, fname))
