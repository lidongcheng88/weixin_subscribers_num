import time
import asyncio

from pyppeteer import launch


async def start():
    browser = await launch(
        headless=False, userDataDir='./userdata', args=['--disable-infobars'])
    page = await browser.newPage()
    await page.setViewport(viewport={'width': 1380, 'height': 800})
    await page.goto('https://mp.weixin.qq.com')
    time.sleep(5)
    await page.type('input[name=account]', '用户名')
    await page.type('input[name=password]', '密码')
    await page.click('a.btn_login')
    time.sleep(10000)


if __name__ == '__main__':
    asyncio.get_event_loop().run_until_complete(start())
