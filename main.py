import dearpygui.dearpygui as dpg
import datetime
import re
import os
import os.path
import json
import configparser
from time import sleep
# from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
from pprint import pprint
import pyperclip

config = configparser.ConfigParser()
config.read('config.ini')

DEFAULT_DATE = datetime.datetime.now().strftime("%d.%m.%Y")
MID_DAY_HOUR = int(config['APP']['MID_DAY_HOUR'])
DRIVER_PATH = config['APP']['DRIVER_PATH']
LOGIN_ACC = config['APP']['LOGIN_ACC']
PASSWORD_ACC = config['APP']['PASSWORD_ACC']

BUFFER = []
RED_COLOR = (255, 0, 0, 255)
GREEN_COLOR = (255, 0, 0, 255)
YELLOW_COLOR = (255, 255, 0, 255)
ELEMENTS = []

dpg.create_context()

with dpg.font_registry():
    with dpg.font("./fonts/JetBrainsMonoNL-Regular.ttf", 20, default_font=True, id="Default font"):
        dpg.add_font_range_hint(dpg.mvFontRangeHint_Cyrillic)

class Query():
    def __init__(self, shop, events):
        self.shop = shop
        self.events = events

    def get_start_date(self):
        min_date = self.events[0]['date']
        for event in self.events:
            if event['date'] < min_date:
                min_date = event['date']
        return min_date
    
    def get_end_date(self):
        max_date = self.get_start_date()
        for event in self.events:
            if event['work_shift'] == 'Day':
                if event['date'] > max_date: max_date = event['date']
            if event['work_shift'] == 'Night':
                if event['date'] + datetime.timedelta(days=1) > max_date:
                    max_date = event['date'] + datetime.timedelta(days=1)
        return max_date

def set_events():
    events = []
    for group in dpg.get_item_children(events_groups, 1):
        event = {}
        counter = 1
        for input in dpg.get_item_children(group, 1):
            if counter == 1:
                event['worker'] = 'ДЮ-' + dpg.get_value(input)
            if counter == 2:
                event['date'] = datetime.datetime.strptime(dpg.get_value(input), '%d.%m.%Y').date()
            if counter == 3:
                event['work_shift'] = dpg.get_value(input)
            counter += 1
        event['status'] = ''
        event['confirm'] = []
        event['delta'] = 0
        events.append(event)
    return events

def start_chrome(driver, shop_number, events, date_in, date_out):
    LOG = []
    service = Service(driver)
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument('--headless=new')
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.get(f'http://{LOGIN_ACC}:{PASSWORD_ACC}@srv-setmon-dcb.dixy.local/TechnoLink/')
    
    radio = driver.find_element(By.ID, 'RadioButtonList1_1')
    radio.click()

    # грузим лог
    shop_input = driver.find_element(By.ID, 'Value_Shop')
    start_date_input = driver.find_element(By.ID, 'Date')
    end_date_input = driver.find_element(By.ID, 'Date2')
    find_on_events_button = driver.find_element(By.ID, 'Events')
    shop_input.clear()
    start_date_input.clear()
    end_date_input.clear()
    shop_input.send_keys(shop_number)
    start_date_input.send_keys(date_in.strftime("%Y-%m-%d") + ' 00:00:00')
    end_date_input.send_keys(date_out.strftime("%Y-%m-%d") + ' 23:59:59')
    find_on_events_button.click()
    try:
        table = WebDriverWait(driver, timeout=2).until(EC.presence_of_element_located((By.ID, 'GW_Result')))
    except Exception:
        return -1
    sleep(1)

    rows = table.find_elements(By.TAG_NAME, 'tr')
    qrlog = []
    for row in rows[1:]:
        tds = row.find_elements(By.TAG_NAME, 'td')
        qrmark = [td.text for td in tds]
        qrlog.append(qrmark)

    labels = ['shop_number', 'tab_number', 'author_qr', 'full_name', 'date_time_mark',
        'create_qr', 'comment', 'pasport', 'error']
        

    normalizelog = [dict(zip(labels, mark)) for mark in qrlog]

    return normalizelog

def parse_log(log, events):
    for event in events:
        if event['work_shift'] == 'Day':
            for mark in log:
                date_time = datetime.datetime.strptime(mark['date_time_mark'], '%d.%m.%Y %H:%M:%S')
                comment = mark['comment']
                error = mark['error']
                if event['worker'] == mark['tab_number'] and date_time.date() == event['date'] and error == ' ':
                    event['confirm'].append({
                        'date_time': date_time,
                        'comment': comment,
                        'error': error
                    })
        if event['work_shift'] == 'Night':
            for mark in log:
                date_time = datetime.datetime.strptime(mark['date_time_mark'], '%d.%m.%Y %H:%M:%S')
                comment = mark['comment']
                error = mark['error']
                if event['worker'] == mark['tab_number'] and (date_time.date() == event['date'] or date_time.date() == event['date'] + datetime.timedelta(days=1)) and error == ' ':
                    event['confirm'].append({
                        'date_time': date_time,
                        'comment': comment,
                        'error': error
                    })

    


def add_event(sender, data):
    group = dpg.add_group(horizontal=True, parent=events_groups)
    dpg.add_input_text(label="Worker number", width=100, parent=group)
    dpg.add_input_text(label="Event date", width=100, default_value=DEFAULT_DATE, parent=group)
    dpg.add_radio_button(['Day', 'Night'], horizontal=True, default_value='Day', parent=group)
    ELEMENTS.append(group)

def destroy_elements(sender, data):
    for element in ELEMENTS:
        dpg.delete_item(element)
    ELEMENTS.clear()

def copy():
    res = ''
    for mess in BUFFER:
        res += mess + '\n'
    pyperclip.copy(res)

def create_new_window(events, shop, log): 
    with dpg.window(label='result', pos=(10, 10), width=760, height=540):
        dpg.add_button(label='COPY TO CLIPBOARD', callback=copy, width=744, height=100)

        dpg.add_separator()
        if log == -1:
            dpg.add_text('События не найдены', color=RED_COLOR)
            return
        for event in events:
            color = (0, 255, 0, 255) if event['status'] == 'succ' else (255, 0, 0, 255)
            dpg.add_text(default_value=f'{shop} по сотруднику {event["worker"]} за {event["date"].strftime("%d.%m.%Y")} ({"Дневная" if event["work_shift"] == "Day" else "Ночная"} смена)', color=color)
            with dpg.group(horizontal=True) as confirm_group:
                dpg.add_text('Отметки      :')
                dpg.add_text(f"({len(event['confirm'])})", color=YELLOW_COLOR)
                with dpg.group(horizontal=False):
                    for mark in event['confirm']:
                        dpg.add_text(f"({mark['date_time']} {mark['comment']} {mark['error']})")
            with dpg.group(horizontal=True) as delta_group:
                dpg.add_text('Отработано   :')
                dpg.add_text(event['delta'])
            with dpg.group(horizontal=True) as status_group:
                dpg.add_text('Статус       :')
                dpg.add_text(event['status'])
            with dpg.group(horizontal=True) as resp_group:
                dpg.add_text('Ответ        :')
                dpg.add_text(event['resp'], wrap=600)
            dpg.add_separator()

def set_state(events):
    for event in events:
        marks = ''
        chek_in = False
        chek_out = False
        if event['work_shift'] == 'Day':
            for time_mark in event['confirm']:
                marks += f'{time_mark["date_time"]}\n'
                if time_mark['date_time'].time().hour < 14 and time_mark['error'] == ' ': chek_in = True
            for time_mark in event['confirm']:
                if time_mark['date_time'].time().hour > 14 and time_mark['error'] == ' ': chek_out = True
            
            
            if chek_in and chek_out: 
                event['status'] = 'succ'
                event['delta'] = event['confirm'][-1]['date_time'] - event['confirm'][0]['date_time']
                return
            if chek_in == False and chek_out == False:
                event['status'] = 'fail_full'
                event['delta'] = 0
                return
            if chek_in == False or chek_out == False:
                event['status'] = 'fail_one'
                event['delta'] = event['confirm'][-1]['date_time'] - event['confirm'][0]['date_time']
                return
            

        if event['work_shift'] == 'Night':
            time_in = 0
            time_out = 0
            for time_mark in event['confirm']:
                if time_mark['date_time'].time().hour > 14 and time_mark['date_time'].date() == event['date'] and time_mark['error'] == ' ': 
                    chek_in = True
                    time_in = time_mark['date_time']
                    marks += f'{time_mark["date_time"]}\n'
            for time_mark in event['confirm']:
                if time_mark['date_time'].time().hour < 14 and time_mark['date_time'].date() == event['date'] + datetime.timedelta(days=1) and time_mark['error'] == ' ': 
                    chek_out = True
                    time_out = time_mark['date_time']
                    marks += f'{time_mark["date_time"]}\n'
            
            
            if chek_in and chek_out: 
                event['status'] = 'succ'
                event['delta'] = time_out - time_in
                return
            if chek_in == False and chek_out == False:
                event['status'] = 'fail_full'
                event['delta'] = 0
                return
            if chek_in == False or chek_out == False:
                event['status'] = 'fail_one'
                event['delta'] = event['confirm'][-1]['date_time'] - event['confirm'][0]['date_time']
                return

    

def set_response(events):
    for event in events:
        if event['status'] == 'succ':
            marks = [mark['date_time'] for mark in event['confirm'] if mark['error'] == ' ']
            marks_str = ''
            for mark in marks:
                marks_str += f'{mark.strftime("%d.%m.%Y %H:%M:%S")}\n'
            event['resp'] = f'По сотруднику {event["worker"]} за {event["date"].strftime("%d.%m.%Y")}.\nВ базе зарегистрированы следующие события:\n{marks_str}Передано на "1-ая линия поддержки 1С: ЗУП"'
            BUFFER.append(event['resp'])

        if event['status'] == 'fail_full':
            event['resp'] = f'В базе Технолинк за {event["date"].strftime("%d.%m.%Y")} данных нет.\nТехническая ошибка не подтверждена.\nДля подтверждения технической проблемы, необходимо вкладывать скриншоты с ошибкой.\nЗапросите подтверждение работы сотрудников от ТР через СВ, далее пишите на электронный адрес "Табель учета рабочего времени магазины все РУ" taburv-allshops@dixy.ru'
            BUFFER.append(event['resp'])

        if event['status'] == 'fail_one':
            marks = [mark['date_time'] for mark in event['confirm'] if mark['error'] == ' ']
            marks_str = ''
            for mark in marks:
                marks_str += f'{mark.strftime("%d.%m.%Y %H:%M:%S")}\n'
            event['resp'] = f'В базе зарегистрированы следующие события:\n{marks_str}Запросите подтверждение работы сотрудников от ТР через СВ, далее пишите на электронный адрес "Табель учета рабочего времени магазины все РУ" taburv-allshops@dixy.ru'
            BUFFER.append(event['resp'])
                

def find(sender, data):
    BUFFER.clear()
    with dpg.window(label='preloader', pos=(250, 200), width=300, height=200, no_move=True, no_close=True, no_resize=True, no_collapse=True, no_title_bar=True, modal=True) as window:        
        dpg.add_loading_indicator(pos=(120, 50))
        dpg.add_text('Загружаем лог...', pos=(90, 120))
    query = Query(dpg.get_value('input_shop'), set_events())
    log = start_chrome(DRIVER_PATH, query.shop, query.events, query.get_start_date(), query.get_end_date())
    if not log == -1:
        parse_log(log, query.events)
        set_state(query.events)
        pprint(query.events)
        set_response(query.events)
    create_new_window(query.events, query.shop, log)

with dpg.window(label="App", tag="main_window"):
    dpg.bind_font("Default font")
    with dpg.group(horizontal=True) as main_group:
        dpg.add_input_text(label="Shop number", tag='input_shop', width=100, default_value = '')
        dpg.add_button(label="ADD EVENT", callback=add_event, width=100)
        dpg.add_button(label="DESTROY", callback=destroy_elements, width=100)
        dpg.add_button(label="FIND", callback=find, width=400)
    dpg.add_separator()
    with dpg.group(horizontal=False) as events_groups:
        with dpg.group(horizontal=True):
            dpg.add_input_text(label="Worker number", tag='input_worker', width=100, default_value = '')
            dpg.add_input_text(label="Event date", tag='input_date', width=100, default_value=DEFAULT_DATE)
            dpg.add_radio_button(['Day', 'Night'], horizontal=True, default_value='Day')

dpg.create_viewport(title='Bio robot', width=800, height=600)
dpg.setup_dearpygui()
dpg.show_viewport()
dpg.set_primary_window("main_window", True)
dpg.start_dearpygui()
dpg.destroy_context()