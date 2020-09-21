from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common import exceptions
import datetime
from tkinter import Tk, messagebox
from itertools import count
from bs4 import BeautifulSoup


EXTRA_INFO_SECTIONS = {"失信被执行", "动产抵押"}

DATE_COL = {
    "变更记录": "变更日期",
    "裁判文书": "公示时间",
    "被执行人": "立案时间",
    "失信被执行": "立案时间",
    "限制高消费": "限消令发布日期",
    "开庭公告": "开庭日期",
    "股权冻结": "起止日期",
    "行政处罚": "处罚日期",
    "环保处罚": "处罚日期",
    "经营异常": "列入日期",
    "税务异常": "发布时间",
    "严重违法失信": "列入日期",
    "动产抵押": "登记日期",
    "股权出质": "登记日期",
    "新闻舆情": "发布日期"
}

VALID_COLS = {
    "变更记录": {"变更事项", "变更前", "变更后"},
    "裁判文书": {"判决时间", "身份", "判决结果"},
    "被执行人": {"案号", "执行标的", "执行法院"},
    "失信被执行": {"案号", "执行依据文号", "执行法院", "被执行人履行情况"},
    "限制高消费": {"案号", "限消法人或组织", "执行法院", "申请执行人"},
    "开庭公告": {"案由", "当事人"},
    "股权冻结": {"被执行人", "股权数额", "执行通知书文号", "起止日期", "类型/状态"},
    "行政处罚": {"决定文书号", "处罚内容", "处罚机关"},
    "环保处罚": {"决定文书号", "处罚内容", "处罚机关"},
    "经营异常": {"做出决定机关", "列入经营异常名录原因", "移出日期", "移出经营异常名录原因"},
    "税务异常": {"纳税人识别号", "欠税税种", "欠税金额", "当前新发生的欠税额"},
    "严重违法失信": {"列入原因", "做出决定机关（列入）", "移出原因", "移出日期", "作出决定机关（移出）"},
    "动产抵押": {"被担保债权种类", "被担保债权数额"},
    "股权出质": {"登记编号", "出质人", "质权人", "状态"}
}

EXTRA_INFO_VALID_COLS = {
    "失信被执行": {"生效法律文书确定的义务", "失信被执行人为具体情形"},
    "动产抵押": {"抵押权人名称"}
}


SCROLL_OFFSET = 150
WAIT_TIMEOUT = 9999


def _date_is_invalid(date_string):
    return (date_string == "-"
        or date_string == "None")
cur_dt = datetime.datetime.now()
target_year, target_month = cur_dt.year, cur_dt.month - 1
if target_month == 0:
    target_year -= 1
    target_month = 12
def _is_in_last_month(date_string):
    if _date_is_invalid(date_string):
        return True
    date = datetime.datetime.fromisoformat(date_string.split()[0]
            .replace("年", "-").replace("月", "-").replace("日", ""))
    return date.year == target_year and date.month == target_month


def set_month_to_fetch(year, month):
    global target_year, target_month
    target_year, target_month = year, month


def _get_first(xs):
    return None if not xs else xs[0]


def _soup(web_element):
    return BeautifulSoup(web_element.get_attribute("outerHTML"), "html.parser")

    
def _scroll_click_wait(driver, btn, wait_condition):
    driver.execute_script('window.scrollTo(0, %d)' % (btn.location['y'] - SCROLL_OFFSET))
    btn.click()
    WebDriverWait(driver, WAIT_TIMEOUT).until(wait_condition)
    
    
def _get_and_wait_for_captcha(driver, url):
    driver.get(url)
    WebDriverWait(driver, WAIT_TIMEOUT).until(EC.invisibility_of_element_located(('class name', "captcha-container")))


class Table:
    def __init__(self):
        self.col_names = []
        self.body = []

    def merge(self, b):
        self.col_names.extend(b.col_names)
        self.body = [row_a + row_b for row_a, row_b in zip(self.body, b.body)]
    
    def _zig_zag_skip_boolean(self, table_entry_with_flag):
        result = [None] * (2 * len(self.col_names))
        result[::2] = (col_name + ": " for col_name in self.col_names)
        result[1::2] = (
            entry + "; " if idx != len(self.col_names) - 1 else entry
            for idx, entry in enumerate(table_entry_with_flag[1:])
        )
        return "".join(result)

    def to_entrys(self):
        dated_entrys = [
            self._zig_zag_skip_boolean(table_entry)
            for table_entry in self.body
            if not table_entry[0]
        ]
        non_dated_entrys = [
            self._zig_zag_skip_boolean(table_entry)
            for table_entry in self.body
            if table_entry[0]
        ]
        return dated_entrys, non_dated_entrys


def login(driver, usern, passw):
    Tk().withdraw()
    messagebox.showinfo("提示", "请不要同时在两处登录同一个启信宝账号。\n"
                        + "运行该爬虫期间，请不要在别处使用同一个账号。"
                        )
    driver.get('https://www.qixin.com/auth/login')
    usern_field = driver.find_element_by_xpath('/html/body/div[2]/div/div/div[2]/div/div[2]/div/div/div/div/div[1]/input')
    passw_field = driver.find_element_by_xpath('/html/body/div[2]/div/div/div[2]/div/div[2]/div/div/div/div/div[2]/input')
    usern_field.send_keys(usern)
    passw_field.send_keys(passw)
    messagebox.showinfo("提示", "请完成人机验证。\n"
                        + "运行该爬虫期间，请不要离开电脑。\n"
                        + "启信宝可能会再次弹出人机验证。\n"
                        + "需要人为操作后，爬虫才能继续运行。"
                        )
    try:
        passw_field.send_keys(webdriver.common.keys.Keys.ENTER)
    except:
        pass
    WebDriverWait(driver, WAIT_TIMEOUT).until(EC.url_changes('https://www.qixin.com/auth/login'))


def get_company_code(driver, company_name):
    for page in count(start = 1):
        _get_and_wait_for_captcha(driver, 'https://www.qixin.com/search?key=%s&page=%d' % (company_name, page))
        search_results_soup = BeautifulSoup(driver.page_source, "html.parser")
        query_res = search_results_soup.find_all("div", class_='company-item')
        if not query_res:
            return None
        for res in query_res:
            res_a = res.find_all("div", recursive=False)[1].div.div.a
            if company_name == res_a.text:
                return res_a['href'].split('/')[-1]


COMPANY_SUBPAGES = {
    '企业概况': 'company',
    '司法涉诉': 'risk',
    '经营预警': 'warn',
    '经营信息': 'operation'
}

def navigate_to_subpage(driver, company_code, subpage_name):
    _get_and_wait_for_captcha(driver, 'https://www.qixin.com/%s/%s' % (COMPANY_SUBPAGES[subpage_name], company_code))
    sections = driver.find_elements_by_class_name('tab-content')
    if subpage_name == "经营信息":
        sections.append(driver.find_element_by_class_name('app-news'))
    
    section_h3s = (_soup(section).h3 for section in sections)
    
    sections = {
        section_h3.text: section
        for section, section_h3 in zip(sections, section_h3s)
        if section_h3 is not None
    }
    return sections


def _table_page_loaded(tab, tab_label, tab_old_text):
    try:
        if tab.get_attribute('class') == 'active':
            return True
        if tab_label.text != str(tab_old_text):
            return True
    except exceptions.StaleElementReferenceException:
        return True
    return False


def _go_to_next_table_page(driver, table_nav, target_page_idx):
    if table_nav is None:
        return False
    for tab in table_nav.find_elements_by_xpath('ul/li')[1:-1]:
        tab_label = tab.find_element_by_xpath('a')
        if tab_label.text == str(target_page_idx):
            _scroll_click_wait(driver, tab_label, lambda _ : _table_page_loaded(tab, tab_label, target_page_idx))
            return True
    return False


def _fetch_extra_info(driver, extra_info_btn, section_name):
    _scroll_click_wait(driver, extra_info_btn, EC.visibility_of_element_located(('class name', 'modal-content')))
    
    extra_info_popup = driver.find_element_by_class_name('modal-content')
    extra_info_entrys = [
        td.text
        for table in _soup(extra_info_popup).find_all('table')
        for td in table.find_all('td')
        if not td.find_all("span", class_="font-f1", recursive=False)
    ]
    extra_info_col_names = []
    extra_info_row = []
    for idx, entry in enumerate(extra_info_entrys):
        if idx % 2 == 0 and (section_name not in EXTRA_INFO_VALID_COLS or entry in EXTRA_INFO_VALID_COLS[section_name]):
            extra_info_col_names.append(entry)
            extra_info_row.append(extra_info_entrys[idx + 1])
    
    extra_info_close_btn = extra_info_popup.find_element_by_xpath('div[1]/div[1]')
    WebDriverWait(driver, WAIT_TIMEOUT).until(lambda _ : extra_info_close_btn.is_enabled())
    extra_info_close_btn.click()
    WebDriverWait(driver, WAIT_TIMEOUT).until(EC.invisibility_of_element_located(('class name', 'modal-content')))
    return extra_info_col_names, extra_info_row


def fetch_section(driver, section, section_name=""):
    table = _get_first(section.find_elements_by_tag_name('table'))
    table_nav = _get_first(section.find_elements_by_tag_name('nav'))
    if table is None:
        return None
    
    get_extra_info = section_name in EXTRA_INFO_SECTIONS or section_name == ""
        
    for next_page_idx in count(start = 2):
        table_soup = _soup(table)
        
        if (next_page_idx == 2):
            main_table = Table()
            main_table.col_names = [
                col_name.text
                for col_name in table_soup.thead.tr.find_all('th', recursive=False)
            ]
            date_col_idx = main_table.col_names.index(DATE_COL[section_name]) if section_name in DATE_COL else -1
            if get_extra_info:
                extra_info_table = Table()
                extra_info_col_idx = next(i for i, col_name in enumerate(main_table.col_names)
                                            if col_name == '详情' or col_name == '操作')
        
        tdss = [
            tr.find_all("td", recursive=False)
            for tr in table_soup.tbody.find_all("tr", recursive=False)
        ]
        
        if date_col_idx != -1:
            date_strings = [
                tds[date_col_idx].text
                for tds in tdss
            ]
        else:
            date_strings = range(len(tdss))
        
        main_table.body.extend(
            ([date_col_idx == -1 or _date_is_invalid(date_string)] +
            [
                td.text
                for col_name, td in zip(main_table.col_names, tds)
                if section_name not in VALID_COLS or col_name in VALID_COLS[section_name]
            ])
            for tds, date_string in zip(tdss, date_strings) if date_col_idx == -1 or _is_in_last_month(date_string)
        )
        
        if get_extra_info:
            extra_infos = [
                _fetch_extra_info(driver, btn, section_name)
                for btn, date_string in zip(table.find_elements_by_xpath(
                                'tbody/tr/td[%d]/a' % (extra_info_col_idx + 1)),
                                date_strings)
                                     if date_col_idx == -1 or _is_in_last_month(date_string)
            ]
            if extra_infos:
                extra_info_table.col_names, _ = extra_infos[0]
                extra_info_table.body = [
                    extra_info_row
                    for _, extra_info_row in extra_infos
                ]
        if not _go_to_next_table_page(driver, table_nav, next_page_idx):
            break
    if not main_table.body: return None
    main_table.col_names = [
        col_name
        for col_name in main_table.col_names
        if section_name not in VALID_COLS or col_name in VALID_COLS[section_name]
    ]
    if get_extra_info:
        main_table.merge(extra_info_table)
    return main_table


def _element_has_class(element, class_name):
    return class_name in element.get_attribute('class').split()


def fetch_tax_section(driver, section):
    section_btns = section.find_elements_by_xpath('div[2]/div')
    if not section_btns: return None, None

    tab1_table = None
    if _element_has_class(section_btns[0], 'selected'):
        tab1_table = fetch_section(driver, section, "税务异常")
        _scroll_click_wait(driver, section_btns[1], lambda _ : _element_has_class(section_btns[1], 'selected'))

    tab2_table = None
    if _element_has_class(section_btns[1], 'selected'):
        tab2_table = fetch_section(driver, section)

    return tab1_table, tab2_table


def fetch_news_section(driver, section):
    recent_3months_btn = section.find_elements_by_tag_name('button')[1]
    _scroll_click_wait(driver, recent_3months_btn, lambda _ : _element_has_class(recent_3months_btn, 'btn-primary'))

    WebDriverWait(driver, WAIT_TIMEOUT).until(EC.invisibility_of_element_located(('id', 'nprogress')))

    table = _get_first(section.find_elements_by_xpath('div[2]'))
    table_nav = _get_first(section.find_elements_by_tag_name('nav'))
    if table is None:
        return None

    news_titles = []
    for next_page_idx in count(start = 2):
        table_soup = _soup(table)
        news_titles.extend(
            (_date_is_invalid(news.div.span.text[5:]), news.div.h4['title'])
            for news in table_soup.find_all('a') if _is_in_last_month(news.div.span.text[5:])
        )
        if not _go_to_next_table_page(driver, table_nav, next_page_idx):
            break
    
    dated_entrys = [
        news_title[1] for news_title in news_titles if not news_title[0]
    ]
    non_dated_entrys = [
        news_title[1] for news_title in news_titles if news_title[0]
    ]
    return dated_entrys, non_dated_entrys
