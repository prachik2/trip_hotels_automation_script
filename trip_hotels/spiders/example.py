# -*- coding: utf-8 -*-
import scrapy
import time
import xlrd
from selenium import webdriver
from selenium.common import exceptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


class TripHotels(scrapy.Spider):
    name = 'trip_hotels'
    allowed_domains = ['trip.com']

    # automation script
    def test(self):
    # def __init__(self, **kwargs):
    #     super().__init__(**kwargs)
        global start_url_list
        start_url_list = []
        loc = "TripReport.xlsx"

        driver = webdriver.Chrome()

        wb = xlrd.open_workbook(loc)
        sheet = wb.sheet_by_index(1)

        sheet.cell_value(0, 0)
        row_data_list = []
        row_data_dict = {}
        for i in range(sheet.nrows - 1):
            row_data_list.append(sheet.row_values(i + 1))
        for j in row_data_list:
            row_data_dict["website"] = j[0]
            row_data_dict["destination"] = j[1]
            row_data_dict["check_in_date"] = xlrd.xldate.xldate_as_datetime(j[2], wb.datemode).date()
            row_data_dict["check_out_date"] = xlrd.xldate.xldate_as_datetime(j[3], wb.datemode).date()
            row_data_dict["room"] = int(j[4])
            row_data_dict["adult"] = int(j[5])
            row_data_dict["children"] = int(j[6])

            driver.get("https://www." + str(row_data_dict["website"]).lstrip())

            driver.maximize_window()

            wait = WebDriverWait(driver, 5)
            driver.implicitly_wait(10)
            wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="hotels-destination"]')))
            destination_box = driver.find_element_by_xpath('//*[@id="hotels-destination"]')
            destination_box.clear()
            destination_box.send_keys(row_data_dict["destination"])
            time.sleep(10)
            checkin = driver.find_element_by_xpath(
                '/html/body/div[1]/div[4]/div[3]/div/div[2]/div/div[2]/div/div/div/ul/li[2]/div/div[1]/input').click()
            select_month_year = driver.find_element_by_xpath(
                '/html/body/div[1]/div[4]/div[3]/div/div[2]/div/div[2]/div/div/div/ul/li[2]/div/div[4]/div/div['
                '1]/div[1]/h3')
            select_month_year.send_keys("Nov 2020")
            WebDriverWait(driver, 10)
            driver.implicitly_wait(10)
            wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'search-btn-wrap')))
            try:
                search_icon = driver.find_element(By.CLASS_NAME, 'search-btn-wrap')
                search_icon.click()
            except exceptions.NoSuchElementException as e:
                pass
            wait = WebDriverWait(driver, 10)
            wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'long-list')))
            # start_url_list.append(driver.current_url)

            try:
                ul_list = driver.find_element_by_class_name('long-list')
                wait.until(EC.presence_of_element_located((By.TAG_NAME, 'li')))
                all_li = ul_list.find_elements_by_tag_name("li")
                for item in all_li:
                    try:
                        wait.until(EC.presence_of_element_located((By.TAG_NAME, 'span')))
                        if item.find_element_by_tag_name("span") and item.is_displayed():
                            WebDriverWait(driver, 10)
                            # get the window handle after the window has opened
                            # window_before = driver.window_handles[0]
            #                 window_before_title = driver.title
            #
                            new_page = item.find_element_by_tag_name("span")
                            new_page.click()
                            win_handle = []
                            win_handle.append(driver.window_handles[0])  #Return a set of window handle
            #                 # get the window handle after a new window has opened
            #                 window_after = driver.window_handles[0]
            #
            #                 # switch on to new child window
            #                 driver.switch_to.window(window_after)
            #                 time.sleep(10)
            #
            #                 window_after_title = driver.title
            #                 winnn = driver.current_url
            #
            #                 # Compare and verify that main window and child window title don't match
            #                 if window_before_title != window_after_title:
            #                     print('Context switched to Twitter, so the title did not match')
            #                 else:
            #                     print('Control did not switch to new window')
                    except exceptions.NoSuchElementException as e:
                        print(e)
                        pass

                    except exceptions.StaleElementReferenceException as e:
                        print(e)
                        pass

            except exceptions.NoSuchElementException as e:
                pass

    def start_requests(self):
        # url = "https://www.trip.com/hotels/list?city=77528&countryId=4&checkin=2020/06/15&checkout=2020/06/16&optionId=11127264&optionType=Hotel&optionName=Bangkok%20Boutique%20Resort&display=Bangkok%20Boutique%20Resort&crn=1&adult=1&children=0&searchBoxArg=t&travelPurpose=0&ctm_ref=ix_sb_dl&domestic=0"
        url ="https://www.trip.com/hotels/london-hotel-detail-2198797/hampton-by-hilton-london-waterloo?checkin=2020" \
             "-12-01&checkout=2020-12-08&city=338&page=1&minprice=164&mincurr=USD&adult=2&children=0&ages=&crn=1" \
             "&display=0&from_page=list&showtotalamt=0&hoteluniquekey" \
             "=H4sIAAAAAAAAAON6y8TFK8Fk8B8GGIWYOBilrjNxbNgwJ9zgEqvFVSdHBhBQmeTgCaIb" \
             "-A47BPBMYpTi7D_0VSOmP8NB8DCIlkpzcGLjODCTUYJlEqMMpqQGI35pJvzSzPilWfBLs" \
             "-KXZoNJz2DsndXGuIKRcSMj2NM5Mx12MDKdYFzCuIBp0YUlnLuYgAo3cUiwHAIyujcDdZxiYrjExHCLieERE8MrJoZPTAy_mKCGNTEzdDEzTGJm5TgtKcEyi5lhETODEAsoXKUUUoyTUi2NzZONUw2NTQzSDJLSLFJNLA2BfNM0yxRLcwVujbuH5nxnNWKwYpZidPNgD2JzMzY3dXGOkuFiDg12EewrLZTmTZBxkALxFGG8JNbUPN2IiIw9rAWMXYxMAoyTGDnBMQnyAgB2HUq37AEAAA((&stand=&stdcode=&shoppingid=&fgt=&module=list&pctoken=d3be937c3e1340f0bf8e491c3e5f9d97-2198797&ltracelogid=&link=title "
        # for url in start_url_list:
        yield scrapy.Request(url, self.parse_hotel_details)

    def parse(self, response):
        time.sleep(20)
        for item in response.css('ul.long-list'):
            li_item = item.css("li > div.with-decorator-wrap")
            if li_item:
                start_url = li_item.css("div.info > div.list-card-tagAndTitle > div.list-card-title ").get()
            else:
                pass

    def parse_hotel_details(self,response):

        for item in response.css("div.hotel-detail_main"):
            data = {
                'hotel_name' : item.css('section.detail-baseinfo_title > h1 ::text').extract_first(),

            }
            count = 0
            for item1 in item.css("section.detail-baseinfo_title > i.detail-baseinfo_title_level"):
                if item1.css("i"):
                    count +=1
                else:
                    pass

                data["rating"] = count


            for item2 in item.css("div.room-list"):
                data["room_type"] = item2.css('div.roomname ::text').extract()
                for amanity in item2.css("div.roomlist-baseroom-card > div.roomcard > div.saleroomlist > div.salecardlist-rooms > div.salecard-frame"):
                    data["breakfast"] = amanity.css("div.salecard-bedfacility  > div.facility > div.des-with-icon > span > span.desc-text underline :: text").extract_first()
                    data["price"] = amanity.css("div.salecard-price > div.salecard-price-panel > div.price :: text").extract_first()

            print(data)
