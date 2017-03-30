import os
import time
import urllib2
import getpass

import pandas as pd
from pandas import read_json
from pandas.io.json import json_normalize
from selenium import webdriver
from selenium.common.exceptions import WebDriverException, NoSuchElementException
from xlwings import Sheet
from xmlutils.xml2json import xml2json

from dcm.dcm_api import DCM_API
from analytics.data.file_io import DataMethods

class URLs(object):

    def __init__(self):
        self.save_path = os.path.join('C:', 'Users', getpass.getuser(), 'Desktop', 'sitemap.xml')
        self.sitemap_url = "https://www.t-mobile.com/sitemap.xml"

    def tmobile_sitemap(self):
        Sheet(DCM_API().SITEMAP).clear()
        s = urllib2.urlopen(self.sitemap_url)
        contents = s.read()

        file = open(self.save_path, 'w')
        file.write(contents)
        file.close()

        json_path = os.path.join(self.save_path[:self.save_path.rindex('\\')], 'sitemap.json')
        xml_convert = xml2json(self.save_path, json_path, encoding='utf-8')
        xml_convert.convert()

        clean_sitemap = read_json(json_path)
        clean_sitemap = \
            json_normalize(clean_sitemap['{http://www.sitemaps.org/schemas/sitemap/0.9}urlset'][
                               '{http://www.sitemaps.org/schemas/sitemap/0.9}url'])

        clean_sitemap.rename(columns=lambda x: x.replace('{http://www.sitemaps.org/schemas/sitemap/0.9}', ''),
                             inplace=True)

        DataMethods().chunk_df(clean_sitemap, DCM_API().SITEMAP, 'A1')
        os.remove(self.save_path)
        Sheet(DCM_API().SITEMAP).autofit()

    @staticmethod
    def list_floodlights_from_urls(urls, driver_path):
        # url_list = raw_input('Enter path to list of URLs (.csv or .xlsx). \n'
        #                      'Must include only one tab with list of URLs in column A: ')
        #
        # url_list = url_list.encode('string-escape')
        #
        # if url_list[-4:] == '.csv':
        #     urls = pd.read_csv(url_list, index_col=None).ix[:, 0].tolist()
        # else:
        #     urls = pd.read_excel(url_list, index_col=None).ix[:, 0].tolist()
        #
        # driver_path = raw_input('The webdriver is saved here: \n'
        #                         'S:\SEA-Media\Analytics\T-Mobile\_\drivers \n'
        #                         'Copy the .exe to another location and paste the full path.')
        #
        os.environ['webdriver.chrome.driver'] = driver_path
        driver = webdriver.Chrome(driver_path)

        floodlights = []

        for url in urls:
            try:
                driver.get(url)
            except WebDriverException:
                pass

            time.sleep(10)
            iframes = driver.find_elements_by_xpath('//iframe')

            fls = []
            for i in iframes:
                src = i.get_attribute('src')
                if 'fls.doubleclick.net' in src:
                    fls.append(src)

            joined = []
            for j in fls:
                split = j.split(';')
                joined.append(';'.join(list(split[0:4])))

            for k in joined:
                if k == 'http://998766.fls.doubleclick.net/activityi':
                    joined.remove(k)

            if url != driver.current_url:
                floodlights.append([url, joined])
            else:
                floodlights.append([url, joined, driver.current_url])

        floodlights = pd.DataFrame(floodlights)

        return floodlights

    @staticmethod
    def url_descriptions():
        url_list = raw_input('Enter path to list of URLs (.csv or .xlsx). \n'
                             'Must include only one tab with list of URLs in column A: ')

        url_list = url_list.encode('string-escape')

        if url_list[-4:] == '.csv':
            urls = pd.read_csv(url_list, index_col=None).ix[:, 0].tolist()
        else:
            urls = pd.read_excel(url_list, index_col=None).ix[:, 0].tolist()

        driver_path = raw_input('The webdriver is saved here: \n'
                                'S:\SEA-Media\Analytics\T-Mobile\_\drivers \n'
                                'Copy the .exe to another location and paste the full path.')

        os.environ['webdriver.chrome.driver'] = driver_path
        driver = webdriver.Chrome(driver_path)

        descriptions = []

        for i in urls:
            try:
                driver.get(i)
                time.sleep(3)

                try:
                    title, description = driver.find_element_by_xpath("//meta[@name='title']").get_attribute('content'), \
                                         driver.find_element_by_xpath("//meta[@name='description']").get_attribute(
                                             'content')
                except NoSuchElementException:
                    pass

                if title is None or '':
                    title = 'Not Found'
                if description is None or '':
                    description = 'Not Found'

            except WebDriverException:
                i = 'Invalid URL'

            descriptions.append(list([i, title, description]))

            url_descriptions = pd.DataFrame(descriptions, columns=['url', 'title', 'description'])

            save_path = os.path.join(url_list[:url_list.rindex('\\')], 'url_floodlights.csv')
            url_descriptions.to_csv(save_path, encoding='utf-8', index=False)