import pandas as pd
import time
from apscheduler.schedulers.blocking import BlockingScheduler
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import requests
from datetime import datetime
from selenium.webdriver.chrome.options import Options

options = Options()
options.add_experimental_option('excludeSwitches', ['enable-logging'])


def log_in_eshop():
    '''
    通过URL判定是否登录成功，比判定元素快，如果网速较慢可能会判定失败
    :param driver:
    :param username: 账号
    :param password: 密码
    :return:
    '''
    url = ''
    brower = webdriver.Chrome(executable_path='./chromedriver', chrome_options=options)
    brower.get(url)

    # 主页url
    index_url = ''
    url = brower.current_url
    while url != index_url:
        time.sleep(3)
        url = brower.current_url
    user_name = WebDriverWait(brower, 10, 1).until(EC.presence_of_element_located((
        By.CSS_SELECTOR,
        "body > div.common-nav-wrapper.common-nav-logout > div.layout-right > div.layout-block.common-nav-back-container > span")))

    # 用户名
    if user_name.text != '':
        print('ILLEGAL_ARGUMENT')
    else:
        print('开始执行擦亮任务')
        return brower


IS_LIGHTING = 0
IS_ALIVEING = 0


def keep_alive(brower):
    print('keep aliving')
    if not IS_LIGHTING:
        try:
            print(' aliving')
            global IS_ALIVEING
            IS_ALIVEING = 1
            creditlife_link = WebDriverWait(brower, 10, 1).until(EC.presence_of_element_located((
                By.CSS_SELECTOR, "body > div.c-wrapper-flow > div.c-wrapper-left > ul > li.m-side-nav-1st.active.unfold > ul > li > a")))
            creditlife_link.click()
            IS_ALIVEING = 0
        except Exception as e:
            print(e)


def click_lighting(brower, goods_msg):
    time.sleep(5)
    table_tr_list = brower.find_elements(By.TAG_NAME, "tr")
    for tr in table_tr_list:
        table_td_list = tr.find_elements(By.TAG_NAME, "td")
        if table_td_list:
            if table_td_list[0].text == goods_msg[0] and table_td_list[1].text[:-3] == goods_msg[1][:-3]:
                if table_td_list[3].text == '已下架':
                    return True, {'success': False, 'goods_name': goods_msg[0], 'msg': '已下架'}
                else:
                    buttons = table_td_list[5].find_elements(By.TAG_NAME, 'button')
                    for button in buttons:
                        if button.text == '擦亮':
                            button.click()
                            time.sleep(3)
                            bts = brower.find_elements(By.TAG_NAME, 'button')
                            # print([bt.text for bt in bts])
                            for bt in bts:
                                if bt.text == '确定':
                                    bt.click()
                                    time.sleep(1)
                                    light_msg = brower.find_element(By.CLASS_NAME, 'el-message__content')
                                    print(light_msg.text)
                                    if light_msg.text == '内容已擦亮':
                                        return True, {'success': True, 'goods_name': goods_msg[0], 'msg': '已擦亮'}
                                    else:
                                        return True, {'success': False, 'goods_name': goods_msg[0], 'msg': light_msg.text}
    return False, {'success': False, 'goods_name': goods_msg[0], 'msg': '商品未找到'}


def send_error_msg(msgs, msg_type='lighting'):
    corporation = '机汤'
    now_time = str(datetime.now())[:16]
    content = ''
    if msg_type == 'lighting':
        opr = '成功' if msgs['success'] else '失败'
        content = '擦亮报告：\n{}在{}的擦亮{}，{}信息是："{}"【{}】'.format(msgs['goods_name'], now_time, opr, opr, msgs['msg'], corporation)
    elif msg_type == 'sys':
        content = f'系统异常警告：\n系统登录出现异常，请重新启动系统【{corporation }】'

    msg_json = {
        "msgtype": "text",
        "text": {"content": content}
    }
    res = requests.post(
        '',
        json=msg_json,
        headers={"Content-Type": "application/json"},
        timeout=2
    )
    res = msg_json
    print(res)
    return res


def lighting(brower, goods_msg):
    print('lingting')
    global IS_LIGHTING
    IS_LIGHTING = 1

    if IS_ALIVEING == 1:
        time.sleep(10)
    try:
        # title = brower.title
        # if title == '芝麻信用商家服务平台':
        next_page = WebDriverWait(brower, 10, 1).until(EC.presence_of_element_located((
            By.CLASS_NAME, "btn-next")))
        while True:
            light = click_lighting(brower, goods_msg)
            finded = light[0]
            msgs = light[1]
            # print(finded)
            if finded:
                break
            if next_page.get_attribute('disabled'):
                light = click_lighting(brower, goods_msg)
                finded = light[0]
                msgs = light[1]
                if not finded:
                    msgs = {'success': False, 'goods_name': goods_msg[0], 'msg': '商品未找到'}
                break
            next_page.click()
        send_error_msg(msgs)
    except:
        # pass
        send_error_msg('', 'sys')
    print('over lighting')
    IS_LIGHTING = 0


scheduler = BlockingScheduler()

brower = log_in_eshop()


df = pd.read_excel('商品擦亮计划.xls')
# 对于每一行，通过列名name访问对应的元素
for row in df.iterrows():
    # print(row[1]['name'], row[1]['create_time'], row[1]['hour'], row[1]['min']) # 输出每一行

    create_time = str(row[1]['创建时间']).replace('/', '-')
    goods_msg = (row[1]['商品名称'], create_time)
    hour = row[1]['hour']
    min = row[1]['min']

    scheduler.add_job(lighting, 'cron', hour=hour, minute=min, args=[brower, goods_msg])


scheduler.add_job(func=keep_alive, trigger='interval', seconds=320, id='ping pang', args=[brower])

jobs = scheduler.get_jobs()
scheduler.start()
