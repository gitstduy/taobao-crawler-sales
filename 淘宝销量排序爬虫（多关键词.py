import sys
import time
import random
import tkinter
from selenium import webdriver
from selenium.webdriver.common.by import By
from threading import Thread
import re
import pandas


# 定义全局变量
CONST_KEY_WORD = ""
CONST_BEGIN_PAGE = 0
CONST_END_PAGE = 0
gui_text = {}
gui_label_now = {}
gui_label_eta = {}
tbPageVersion = 0
keywords_list = []
all_data = []
keyword = ''

#读取设置信息+读取老的商品信息
def read_settings():
    global CONST_KEY_WORD, CONST_BEGIN_PAGE, CONST_END_PAGE, old_ID_data, keywords_list  # 声明keywords_list为全局变量
    try:
        # 读取设置文件
        with open('./setkeyword.ini', 'r+', encoding='utf8') as f:
            keywords_list = [line.strip() for line in f.readlines()]
        with open('./setpages.ini', 'r+', encoding='utf8') as h:
            lines = h.readlines()
            if len(lines) == 2:
                CONST_BEGIN_PAGE = int(lines[0].strip())
                CONST_END_PAGE = int(lines[1].strip())
            else:
                print("抓取页数设置文件格式不正确")
        CONST_KEY_WORD = keywords_list[0]

        # 读取上周数据库文件文件
        df = pandas.read_excel("商品库.xlsx")
        ID_data = df["商品ID"].astype(int)
        # 存储到一个list
        old_ID_data = ID_data.tolist()

        gui_text['text'] = '上周数据库读取中'
        if old_ID_data:
            gui_text['text'] = '读取成功'
            print(old_ID_data)
        else:
            gui_text['text'] = '读取失败'
            sys.exit()

    except Exception as e:
        print("读取设置文件失败:", e)

# GUI函数
def gui_func():
    global gui_text, gui_label_now, gui_label_eta
    gui = tkinter.Tk()
    gui.title('淘宝搜索页面爬取---MOH电商')
    gui['background'] = '#ffffff'
    gui.geometry("600x100-50+20")
    gui.attributes("-topmost", 1)
    gui_text = tkinter.Label(gui, text='初始化', font=('微软雅黑', '20'))
    gui_text.pack()
    gui_label_now = tkinter.Label(gui, text='暂无信息', font=('微软雅黑', '10'))
    gui_label_now.pack()
    gui_label_eta = tkinter.Label(gui, text='暂无信息', font=('微软雅黑', '10'))
    gui_label_eta.pack()
    gui.mainloop()

# 初始化GUI线程
Gui_thread = Thread(target=gui_func, daemon=True)
Gui_thread.start()
time.sleep(2)


# 启动浏览器
def start_browser():
    try:
        options = webdriver.ChromeOptions()
        options.add_experimental_option('excludeSwitches', ['enable-automation'])
        options.add_argument("--disable-blink-features")
        options.binary_location = "F:\google catch\chrome-win64\chrome.exe"
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_experimental_option('excludeSwitches', ['enable-logging'])

        browser = webdriver.Chrome(executable_path="F:\google catch\chromedriver-win64\chromedriver.exe",
                                   options=options)

        browser.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
            "source": """Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"""
        })

        return browser
    except Exception as e:
        print("启动浏览器失败:", e)
        sys.exit()


def login_taobao(browser):
    try:
        browser.get('https://login.taobao.com/member/login.jhtml')
        browser.maximize_window()

        # GUI提醒登录
        gui_text['background'] = '#ffffff'
        gui_text['text'] = '请尽快扫码登录淘宝，10秒后将尝试自动爬取'
        time.sleep(5)
        gui_text['text'] = '如有滑块认证请辅助滑块认证'
        time.sleep(5)

        # 循环检测是否成功登录
        while browser.title != '我的淘宝':
            gui_text['text'] = '''请确保已经登录淘宝并在【我的淘宝】页面'''
            gui_text['bg'] = 'red'
            time.sleep(5)
            gui_text['text'] = '''即将重新尝试自动爬取'''

    except Exception as e:
        print("登录淘宝失败:", e)
        sys.exit()

def enter_search_page(browser,guanjianci):
    try:
        # 搜索词与页数管理
        gui_text['text'] = '正在操作'
        gui_text['background'] = '#ffffff'
        browser.get(f'https:s.taobao.com/search?q={guanjianci}')
        browser.implicitly_wait(8)
        gui_text['text'] = f'等待完全加载页面中'
        # GUI提醒：验证码拦截
        while browser.title == '验证码拦截':
            gui_text['text'] = f'出错：如有滑块验证请及时验证，程序将于验证后重新尝试爬取'
            gui_text['bg'] = 'red'
            gui_label_eta['text'] = '-'
            gui_label_now['text'] = f'-'
            time.sleep(2)

        # 检测淘宝新老搜索页面
        try:
            # 老版PC淘宝页面
            taobaoPage = browser.find_element(By.CSS_SELECTOR,
                                              '#J_relative > div.sort-row > div > div.pager > ul > li:nth-child(2)').text
            taobaoPage = re.findall('[^/]*$', taobaoPage)[0]
            tbPageVersion = 0

        except:
            # 新版PC淘宝页面
            taobaoPage = browser.find_element(By.CSS_SELECTOR,
                                              '#sortBarWrap > div.SortBar--sortBarWrapTop--VgqKGi6 > '
                                              'div.SortBar--otherSelector--AGGxGw3 > div:nth-child(2) > '
                                              'div.next-pagination.next-small.next-simple.next-no-border > div > span').text
            taobaoPage = re.findall('[^/]*$', taobaoPage)[0]
            tbPageVersion = 1
            sale_desc_button = browser.find_element(By.CSS_SELECTOR,
                                                    '#sortBarWrap > div.SortBar--sortBarWrapTop--VgqKGi6 > div:nth-child(1) > div > div.next-tabs-bar.SortBar--customTab--OpWQmfy > div > div > div > ul > li:nth-child(2)')
            sale_desc_button.click()

        # GUI提醒页面
        gui_text['text'] = f'目录检测到共计{taobaoPage}页'
        gui_text['background'] = '#f35315'
        time.sleep(2)

        # 页数控制
        page_start = int(CONST_BEGIN_PAGE)
        page_end = int(CONST_END_PAGE + 1) if int(CONST_END_PAGE + 1) < int(taobaoPage) else int(taobaoPage)

        # GUI提醒页面
        gui_text['text'] = f'将从{page_start}页爬取到{page_end - 1}页'
        gui_text['background'] = '#ffffff'
        time.sleep(2)

        return browser, page_start, page_end, tbPageVersion
    except Exception as e:
        print("进入搜索页面失败:", e)
        sys.exit()

# 爬取数据
def scrape_data(browser, page_start, page_end, tbPageVersion):
    try:
        output_list = []

        # 循环爬取程序
        for page in range(page_start, page_end):
            gui_text['text'] = f'当前正在获取第{page}页，还有{page_end - page_start - page}页'
            gui_text['bg'] = '#10d269'
            gui_label_now['text'] = '暂无数据'
            gui_label_eta['text'] = '暂无数据'
            month_deals = 0
            # 判断淘宝搜索页面版本
            if tbPageVersion == 0:
                browser.get(
                    f'https://s.taobao.com/search?_input_charset=utf-8&commend=all&ie=utf8&page=3&q={CONST_KEY_WORD}&search_type=item&source=suggest&sourceId=tb.index&spm=a21bo.jianhua.201856-taobao-item.2&ssid=s5-e&suggest=history_1&suggest_query=&wq=&sort=sale-desc&s={(page - 1) * 44} ')
            elif tbPageVersion == 1:
                if page != page_start:
                    next_page_button = browser.find_element(By.CSS_SELECTOR,
                                                            '#sortBarWrap > div.SortBar--sortBarWrapTop--VgqKGi6 > div.SortBar--otherSelector--AGGxGw3 > div:nth-child(2) > div.next-pagination.next-small.next-simple.next-no-border > div > button.next-btn.next-small.next-btn-normal.next-pagination-item.next-next')
                    next_page_button.click()

            # 验证码拦截
            while browser.title == '验证码拦截':
                gui_text['text'] = f'出错：如有滑块验证请及时验证，程序将于验证后重新尝试爬取'
                gui_text['bg'] = 'red'
                gui_label_eta['text'] = '-'
                gui_label_now['text'] = f'-'
                time.sleep(5)

            # GUI显示
            gui_text['text'] = f'当前正在获取{CONST_KEY_WORD}第{page}页，还有{page_end - page_start - page}页'
            gui_text['bg'] = '#10d269'

            # 定位元素
            try:
                # if tbPageVersion == 0:
                #     print('using classic version selector')
                #     goods_arr = browser.find_elements(By.CSS_SELECTOR, '#mainsrp-itemlist > div > div > div:nth-child(1)>div')
                #     goods_length = len(goods_arr)
                #     # 遍历商品
                #     for i, goods in enumerate(goods_arr):
                #         gui_label_now['text'] = f'正在获取第{i}个,共计{goods_length}个'
                #         item_name = goods.find_element(By.CSS_SELECTOR,
                #                                        'div.ctx-box.J_MouseEneterLeave.J_IconMoreNew > div.row.row-2.title>a').text
                #         item_price = goods.find_element(By.CSS_SELECTOR,
                #                                         'div.ctx-box.J_MouseEneterLeave.J_IconMoreNew > div.row.row-1.g-clearfix > div.price.g_price.g_price-highlight > strong').text
                #         item_shop = goods.find_element(By.CSS_SELECTOR,
                #                                        'div.ctx-box.J_MouseEneterLeave.J_IconMoreNew > div.row.row-3.g-clearfix > div.shop > a > span:nth-child(2)').text
                #         month_deals = goods.find_element(By.CSS_SELECTOR,
                #                                          'div.ctx-box.J_MouseEneterLeave.J_IconMoreNew > div.row.row-1.g-clearfix > div.deal-cnt').text.replace(
                #             '人付款', '')
                #         ships_from = goods.find_element(By.CSS_SELECTOR,
                #                                         'div.ctx-box.J_MouseEneterLeave.J_IconMoreNew > div.row.row-3.g-clearfix > div.location').text
                #         shop_link = goods.find_element(By.CSS_SELECTOR,
                #                                        'div.ctx-box.J_MouseEneterLeave.J_IconMoreNew > div.row.row-3.g-clearfix > div.shop > a').get_attribute(
                #             'href')
                #         item_link = goods.find_element(By.CSS_SELECTOR,
                #                                        'div.pic-box.J_MouseEneterLeave.J_PicBox > div > div.pic>a').get_attribute(
                #             'href')
                #         goods_item = {"商品名称": item_name, "商品价格": item_price, "月销售量": month_deals,
                #                       "商品店铺名称": item_shop, "归属地": ships_from, "商品链接": item_link}
                #         output_list += [goods_item]

                if tbPageVersion == 1:
                    print('新页面')
                    time.sleep(1)

                    # # 查找包含验证码的<div>元素
                    # try:
                    #     captcha_div = browser.find_element_by_class_name('captcha-tips')
                    #     # 检查是否找到了验证码
                    #      if captcha_div:
                    #         print("页面包含验证码")
                    #         # 等待用户滑动解锁
                    #         delay_time = random.randint(30, 60)
                    #         for delay in range(delay_time):
                    #             gui_label_now['text'] = '-'
                    #             gui_text['bg'] = '#eeeeee'
                    #             gui_text['text'] = f'出错：如有滑块验证请及时验证，程序将于验证后重新尝试爬取'
                    #             gui_label_eta['text'] = f'等待下次翻页{delay}秒，总共需等待{delay_time}秒'
                    #             time.sleep(1)
                    #
                    # except Exception as e:
                    #             print("爬取数据失败:", e)


                    goods_arr = browser.find_elements(By.CSS_SELECTOR,
                                                      '#root > div > div:nth-child(2) > div.PageContent--contentWrap--mep7AEm > div.LeftLay--leftWrap--xBQipVc > div.LeftLay--leftContent--AMmPNfB > div.Content--content--sgSCZ12 > div>div')
                    while len(goods_arr) ==0:
                        print("抓取页面商品中")
                        goods_arr = browser.find_elements(By.CSS_SELECTOR,
                                                          '#root > div > div:nth-child(2) > div.PageContent--contentWrap--mep7AEm > div.LeftLay--leftWrap--xBQipVc > div.LeftLay--leftContent--AMmPNfB > div.Content--content--sgSCZ12 > div>div')
                        time.sleep(1)

                    goods_length = len(goods_arr)
                    for i, goods in enumerate(goods_arr):
                        i = i + 1
                        gui_label_now['text'] = f'正在获取第{i}个,共计{goods_length}个'

                        item_name = goods.find_element(By.CSS_SELECTOR,
                                                       f'div:nth-child({i})>a>div > div.Card--mainPicAndDesc--wvcDXaK > div.Title--descWrapper--HqxzYq0 > div > span').text
                        item_price_int = goods.find_element(By.CSS_SELECTOR,
                                                            f'div:nth-child({i})>a>div > div.Card--mainPicAndDesc--wvcDXaK > div.Price--priceWrapper--Q0Dn7pN > span.Price--priceInt--ZlsSi_M').text
                        item_price_float = goods.find_element(By.CSS_SELECTOR,
                                                              f'div:nth-child({i})>a>div> div.Card--mainPicAndDesc--wvcDXaK > div.Price--priceWrapper--Q0Dn7pN > span.Price--priceFloat--h2RR0RK').text
                        item_price = item_price_int + item_price_float
                        item_shop = goods.find_element(By.CSS_SELECTOR,
                                                       f'div:nth-child({i})>a>div> div.ShopInfo--shopInfo--ORFs6rK  > div>a').text
                        month_deals = goods.find_element(By.CSS_SELECTOR,
                                                         f'div:nth-child({i}) > a > div > div.Card--mainPicAndDesc--wvcDXaK > div.Price--priceWrapper--Q0Dn7pN > span.Price--realSales--FhTZc7U').text.replace(
                            '+人付款', '').replace('+人收货', '').replace('万', '0000')
                        ships_from_province = goods.find_element(By.CSS_SELECTOR,
                                                                 f'div:nth-child({i}) > a > div > div.Card--mainPicAndDesc--wvcDXaK > div.Price--priceWrapper--Q0Dn7pN > div:nth-child(5) > span').text

                        shop_link = goods.find_element(By.CSS_SELECTOR,
                                                       f'div:nth-child({i})>a>div> div.ShopInfo--shopInfo--ORFs6rK  > div>a').get_attribute(
                            'href')
                        # 商品地址
                        item_link = goods.find_element(By.CSS_SELECTOR,
                                                       f'div:nth-child({i})>a').get_attribute(
                            'href')

                        # 使用正则表达式来获取商品id的值，并转换为整数
                        match = re.search(r"id=(\d+)", item_link)
                        if match:
                            item_id = match.group(1)
                        else:
                            item_id = '0'
                        # 定位城市，由于有些没有城市属性所以需要try-except，但是很慢
                        # try:
                        #     ships_from_city = goods.find_element(By.CSS_SELECTOR,
                        #                                          f'div:nth-child({i}) > a > div > div.Card--mainPicAndDesc--wvcDXaK > div.Price--priceWrapper--Q0Dn7pN > div:nth-child(6) > span').text
                        # except:
                        #     ships_from_city = ''
                        ships_from_city = ''
                        goods_item = {"商品名称": item_name, "商品ID": item_id, "商品价格": item_price,
                                      "月销售量": month_deals,
                                      "商品店铺名称": item_shop, "归属地": ships_from_province + ' ' + ships_from_city,
                                      "商品链接": item_link}
                        # 检查月收货人数是否小于30，如果小于30则提前结束当前抓取
                        if month_deals.isdigit() and int(month_deals) <  50:
                            print(f"月收货人数小于50，提前结束当前关键词抓取: 第{page}页")
                            return output_list

                        output_list += [goods_item]

            except:
                gui_text['text'] = f'本页面定位元素失败，程序将于5秒后重新尝试爬取'
                gui_text['bg'] = 'red'
                gui_label_eta['text'] = '暂无信息'
                gui_label_now['text'] = f'注意:第【{page}】页将跳过如需获取请重新运行程序！'
                print(f'注意:第【{page}】页将跳过如需获取请重新运行程序！')
                time.sleep(5)

            delay_time = random.randint(5, 10)
            for delay in range(delay_time):
                gui_label_now['text'] = '-'
                gui_text['bg'] = '#eeeeee'
                gui_text['text'] = f'{CONST_KEY_WORD}第{page}页，还有{page_end - page_start - page}页'
                gui_label_eta['text'] = f'等待下次翻页{delay}秒，总共需等待{delay_time}秒'
                time.sleep(1)

        return output_list
    except Exception as e:
        print("爬取数据失败:", e)
        return []


# 导出排行榜数据到Excel
def export_to_excel(data):
    try:
        gui_text['text'] = '正在导出xlsx'
        output_dataframe = pandas.DataFrame(data).drop_duplicates()
        output_dataframe.to_excel('淘宝销量榜' + f'{time.strftime("%Y-%m-%d_%H-%M-%S", time.localtime())}' + '.xlsx',
                                  index=False)
        gui_text['text'] = '保存文件完成，准备退出中'
        time.sleep(5)
    except Exception as e:
        print("导出数据失败:", e)


#导出新品数据
def export_to_excel2(data):
    try:
        gui_text['text'] = '正在导出xlsx'
        output_dataframe = pandas.DataFrame(data)
        output_dataframe.to_excel('新品库' + f'{time.strftime("%Y-%m-%d_%H-%M-%S", time.localtime())}' + '.xlsx',
                                  index=False)
        gui_text['text'] = f'保存新品库文件完成，准备退出中'
        time.sleep(5)
    except Exception as e:
        print("导出数据失败:", e)


def find_new_products(output_list, old_ID_data):
    try:
        # 将商品ID列转换为整数
        output_dataframe = pandas.DataFrame(output_list, columns=["商品名称", "商品ID", "商品价格", "月销售量", "商品店铺名称", "归属地", "商品链接"])
        #output_dataframe['商品ID'] = output_dataframe['商品ID'].astype(int)

        # 使用 isin 函数进行比较
        mask = output_dataframe['商品ID'].isin(old_ID_data)

        # 使用布尔索引从新数据中删除匹配的行
        filtered_data = output_dataframe[~mask]

        if not filtered_data.empty:
            gui_text['text'] = '恭喜找到新品'
            return filtered_data, '恭喜找到新品'
        else:
            gui_text['text'] = '寻找新品失败'
            return None, '寻找新品失败'
    except Exception as e:
        return None, f'出现错误: {str(e)}'


# 主函数
if __name__ == '__main__':
    try:

        read_settings()
        #gui_text函数输出显示当下我抓取的关键词有哪些
        gui_label_now['text'] = f'本次抓取的词：{keywords_list}'
        time.sleep(2)
        browser=start_browser()  # 启动浏览器放在循环外部
        login_taobao(browser)

        for keyword in keywords_list:
            CONST_KEY_WORD = keyword  # 设置当前关键词
            browser, page_start, page_end, tbPageVersion = enter_search_page(browser, CONST_KEY_WORD)
            time.sleep(2)  # 等待页面加载
            data = scrape_data(browser, page_start, page_end, tbPageVersion)
            if data:
                print('抓取成功')
                all_data.extend(data)
        if all_data:
            export_to_excel(all_data)
        # 验证抓取到的数据是否有新品\
        # new_products, message = find_new_products(all_data, old_ID_data)
        # if new_products is not None:
        #     print(f'找到新款')
        #     export_to_excel2(new_products)
        # else:
        #     print(f'寻找新款失败 ')

        # export_to_excel(all_data)
    except Exception as e:
        print("程序发生错误:", e)
    finally:
        sys.exit()



