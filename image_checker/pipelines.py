# -*- coding: utf-8 -*-

# Define your item pipelines here
#
# Don't forget to add your pipeline to the ITEM_PIPELINES setting
# See: https://doc.scrapy.org/en/latest/topics/item-pipeline.html
from scrapy import signals
import xlsxwriter
import os

class ImageCheckerPipeline(object):
    def __init__(self):
            #Instantiate API Connection
        self.workbook = None

    @classmethod
    def from_crawler(cls, crawler):
        pipeline = cls()
        crawler.signals.connect(pipeline.spider_opened, signals.spider_opened)
        crawler.signals.connect(pipeline.spider_closed, signals.spider_closed)
        return pipeline


    def spider_opened(self, spider):
        filepath = 'output.xlsx'
        if os.path.isfile(filepath):
            os.remove(filepath)
        self.workbook = xlsxwriter.Workbook(filepath)
        self.sheet = self.workbook.add_worksheet('output.xlsx')
        self.headers = ['Sku', 'Image']
        self.index = 0
        for col, val in enumerate(self.headers):
            self.sheet.write(self.index, col, val)
        self.sheet.write(self.index, col, val)
    def spider_closed(self, spider):
        models = spider.models
        for row, item in enumerate(models):
            for col, key in enumerate(item.keys()):
                self.sheet.write(row+1, col, item[key])
        self.workbook.close()

    def process_item(self, item, spider):
        return item
