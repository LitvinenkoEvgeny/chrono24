# -*- coding: utf-8 -*-
import scrapy
import xlrd
from xlutils.copy import copy


class ChronoSpider(scrapy.Spider):
    name = "chrono"
    allowed_domains = ["www.chrono24.com"]
    start_urls = (
        'http://www.chrono24.com/en/search/index.htm',
    )

    def parse(self, response):
        brend_values = {}
        for option in response.xpath(
                '//select[@name="manufacturerIds"]/option'):
            brend = option.xpath("text()").extract()[0]
            value = option.xpath("@value").extract()[0]
            if len(value) > 0:
                brend_values[brend] = value

        for key in brend_values.keys():
            url = "http://www.chrono24.com/en/search/index.htm?dosearch=true&watchTypes=U&searchexplain=false&manufacturerIds={}&query=&STARTSEARCH=Search&watchCategories=&gender=&caseMaterials=&braceletMaterial=&countryIds=&SEARCH_REGION_ID=&maxAgeInDays=&sortorder=0&priceFrom=&priceTo=&SEARCHSTATS_BRAND_ID=&SEARCHSTATS_MATERIAL_ID=".format(
                brend_values[key])

            request = scrapy.Request(
                url=url,
                meta={"brend": key},
                callback=self.brend_products
            )
            yield request

    def brend_products(self, response):
        brend = response.meta["brend"]

        # loop over items in page
        for item in response.xpath(
                '//a[@class="list-item relative rounded-small clearfix"]/@href'
        ):
            url = response.urljoin(item.extract())
            yield scrapy.Request(url, self.item_page,
                                 meta={"brend": brend})

        try:
            next_page_url = response.xpath(
                '//td[@class="page last"]/div/a/@href')\
                .extract()[0]
            next_page_url = response.urljoin(next_page_url)
            yield scrapy.Request(next_page_url, self.brend_products,
                                 meta={"brend": brend})
        except IndexError:
            # print("all pages was scraped")
            pass

    def item_page(self, response):
        print("=" * 50)
        # print(response.url)
        brend = response.meta["brend"]
        name = response.xpath(
            '//h1[@class="watch-headline sub"]/text()').extract()[0]

        try:
            price = response.xpath(
                '//span[@class="spacing-v-none ad-price"]/text()')\
                .extract()[1]\
                .replace("$", "")
        except IndexError:
            price = "price on request"

        seller_type = response.xpath(
            '//strong[@class="seller-type"]/text()')\
            .extract()

        if len(seller_type) == 0:
            try:
                seller_type = response.xpath(
                    '//div[@class="spacing-v-sm seller-data"]/text()')\
                    .extract()[0].strip()
                seller_type = [seller_type]
            except IndexError:
                try:
                    seller_type = response.xpath(
                        '//strong[@class="private-seller"]/text()')\
                        .extract()[0]
                    seller_type = [seller_type]
                except IndexError:
                    seller_type = response.xpath(
                        '//span[@class="private-seller"]/text()')\
                        .extract()[0]
                    seller_type = [seller_type]

        seller_type = seller_type[0].replace(" |", "")

        try:
            contact_info = response.xpath(
                '//p[@class="hide-simple"]/text()').extract()
        except IndexError:
            contact_info = ["empty"]
        new_contact_info = []

        for info in contact_info:
            info = info.strip()
            if len(info) == 0:
                continue
            new_contact_info.append(info)

        # end variant
        # contact_info = u"\n\t\t".join(new_contact_info)
        contact_info = new_contact_info

        print(u"""
brend:\t\t{}
name:\t\t{}
price:\t\t{}
seller-type:\t{}
contact_info:\t{}
""".format(brend, name, price.replace("$", ""),
           seller_type,
           u"\n\t\t".join(contact_info)))

        print("=" * 50)

        self.write_to_csv([brend, name, price, seller_type,
                           ";".join(contact_info)])

    def write_to(self, data):
        with open("watchs", "ar+") as write_to:
            write_to.write(data.encode("utf-8"))

    def excel_write(self, data):
        rb = xlrd.open_workbook('data.xls', formatting_info=True)
        r_sheet = rb.sheet_by_index(0)
        r = r_sheet.nrows
        wb = copy(rb)
        sheet = wb.get_sheet(0)
        sheet.write(r, 0, data[0])
        sheet.write(r, 1, data[1])
        sheet.write(r, 2, data[2])
        sheet.write(r, 3, data[3])
        for index, info in enumerate(data[4]):
            sheet.write(r, 4 + index, info)
        wb.save('data.xls')

    def write_to_csv(self, data):
        with open("data.csv", "a") as csv:
            data = ";".join(data) + "\n"
            data = data.encode("utf-8")
            csv.write(data)
