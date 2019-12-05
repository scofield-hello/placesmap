# -*- coding: utf-8 -*-
import re
from typing import List
from scrapy import Spider, Request
from scrapy.http import TextResponse
from bs4 import BeautifulSoup, Tag
from placesmap.items import PlacesmapItem


class PlaceSpider(Spider):
    name = 'place'
    allowed_domains = ['placesmap.net', 'placesmap.us']
    start_urls = ['https://placesmap.net/']
    country_name = None
    interests = []

    def __init__(self, country: str, interests: str, **kwargs):
        super().__init__(**kwargs)
        self.country_name = country.strip().replace("-", " ")
        spilit_interests = list(
            filter(lambda i: len(i.strip()) > 0, interests.split(",")))
        self.interests = list(
            map(lambda i: i.replace("-", " "), spilit_interests))
        print(f"目标国家\区域：{self.country_name}")
        print(f"目标数据类型：{self.interests}")

    def parse(self, response: TextResponse) -> None:
        bs4 = BeautifulSoup(response.text, "lxml")
        country_node = bs4.find(name="a", text=self.country_name)
        country_node_href = f"https:{country_node['href']}"
        print(f"目标国家/区域 链接地址:{country_node_href}")
        yield Request(url=country_node_href, callback=self.parse_area_page)

    def parse_area_page(self, response: TextResponse) -> None:
        bs4 = BeautifulSoup(response.text, "lxml")
        area_div_list = bs4.find_all(name="div", class_="four columns")
        print(f"共包含区域个数：{len(area_div_list)}")
        for div in area_div_list:
            a_node = div.find("a")
            if a_node == None or a_node.string == None:
                continue
            a_node_href = f"https:{a_node['href']}"
            print(f"区域：{a_node.string} , 链接：{a_node_href}")
            yield Request(url=a_node_href,
                          meta={"area_name": a_node.string},
                          callback=self.parse_interests_category)

    def parse_interests_category(self, response: TextResponse) -> None:
        bs4 = BeautifulSoup(response.text, "lxml")
        interests = [bs4.find(name="a", text=i) for i in self.interests]
        for interest_node in interests:
            interest_node_href = f"https:{interest_node['href']}"
            yield Request(url=interest_node_href,
                          meta=response.meta,
                          callback=self.parse_place_pagination)

    def parse_place_pagination(self, response: TextResponse) -> None:
        pagination = re.search(
            r"Found: (\d{1,}.?\d{0,}) Places, (\d{1,}.?\d{0,}) Pages",
            response.text)
        if pagination != None:
            max_page = int(pagination.group(2))
            if max_page == 1:
                self.parse_place_item(response)
            else:
                for page_index in range(1, max_page + 1):
                    indexed_url = response.url + f"/{page_index}/"
                    yield Request(url=indexed_url,
                                  meta=response.meta,
                                  callback=self.parse_place_item)

    def parse_place_item(self, response: TextResponse) -> None:
        bs4 = BeautifulSoup(response.text, "lxml")
        div_node_list: List[Tag] = bs4.find_all("div", class_="six columns")
        for div_node in div_node_list:
            p_node_list = div_node.find_all(name="p", recursive=False)
            for p_node in p_node_list:
                if p_node is None:
                    continue
                p_b_node: Tag = p_node.find(name="b")
                if p_b_node is None:
                    continue
                b_a_node: Tag = p_b_node.find(name="a")
                if b_a_node is None:
                    continue
                place_item = PlacesmapItem()
                place_item['F1'] = response.meta['area_name']
                place_item['F2'] = b_a_node.string.strip()
                place_item['F3'] = p_node.contents[3].strip()
                a_node: Tag = p_node.find(name="a", recursive=False)
                cordinate = re.search(r"([0-9\.-]{3,}),\s+([0-9\.-]{3,})",
                                      a_node.string)
                place_item['F4'] = cordinate.group(1).strip()
                place_item['F5'] = cordinate.group(2).strip()
                detail_text = p_node.contents[8]
                m_telphone = re.search(
                    r"Phone:\s?((\(\d{1,4}\)\s?[0-9\s-]{3,30})|([+0-9\s-]{3,30}))",
                    detail_text)
                m_website = re.search(r"\(([a-zA-Z]{1,}.*)\)", detail_text)
                place_item['F6'] = m_telphone.group(
                    1).strip() if m_telphone else ""
                place_item['F7'] = m_website.group(1).replace(
                    " ", "") if m_website else ""
                yield place_item