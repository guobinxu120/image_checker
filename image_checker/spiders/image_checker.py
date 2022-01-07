# -*- coding: utf-8 -*-
from scrapy import Spider, Request
from collections import OrderedDict
from xlrd import open_workbook
import os, requests, re

def download(url, destfilename):
	if not os.path.exists(destfilename):

		try:
			r = requests.get(url, stream=True)
			with open(destfilename, 'wb') as f:
				for chunk in r.iter_content(chunk_size=1024):
					if chunk:
						f.write(chunk)
						f.flush()
		except:
			pass

def readExcel(path):
    wb = open_workbook(path)
    result = []
    for sheet in wb.sheets():
        number_of_rows = sheet.nrows
        number_of_columns = sheet.ncols
        herders = []
        for row in range(0, number_of_rows):
            values = OrderedDict()
            for col in range(number_of_columns):
                value = (sheet.cell(row,col).value)
                if row == 0:
                    herders.append(value)
                else:

                    values[herders[col]] = value
            if len(values.values()) > 0:
                result.append(values)
        break

    return result


class AngelSpider(Spider):
    name = "image_checker"
    start_urls = ['https://www.tropicmarket.com/']
    count = 0
    use_selenium = False
    models = readExcel("Input_Image_crawler.xlsx")
    #
    # def start_requests(self):
    #     for i, val in enumerate(self.models):
    #         url = val['URL LINKS']
    #         yield Request(url , callback=self.parse1, errback=self.parse, meta={'order_num':i}, dont_filter=True)

    def parse(self, response):
        for i, item in enumerate(self.models):
            yield Request(response.urljoin(item['Sku']), callback=self.parse1, errback=self.err_parse, dont_filter=True, priority=len(self.models)-i, meta={'index':i})

    def parse1(self, response):
        i = response.meta['index']
        sku = self.models[i]['Sku']
        # image_url = response.xpath('//meta [@property="og:image"]/@content').extract_first()
        image_urls = re.findall('"full":(.*),"caption"',str(response.body))
        if len(image_urls) > 0:
            image_url = image_urls[0].replace('"','').replace('\\','')
            image_name = image_url.split('/')[-1]

            filename = "Images/" + image_name
            download(image_url, filename)

            item=OrderedDict()
            item['Sku'] = sku
            item['Image'] = image_name
            self.models[i] = item
            yield item

    def err_parse(self, response):
        pass
        # i = response.request.meta['index']
        # item=OrderedDict()
        # item['Sku'] = self.models[i]
        # item['Image'] = ''
        # self.models[i] = item
        # yield item