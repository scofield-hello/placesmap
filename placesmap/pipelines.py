# -*- coding: utf-8 -*-

# Define your item pipelines here
#
# Don't forget to add your pipeline to the ITEM_PIPELINES setting
# See: https://docs.scrapy.org/en/latest/topics/item-pipeline.html
import os, datetime
from logging import getLogger
from typing import List
from placesmap.excel_item_exporter import ExcelItemExporter


class PlacesmapPipeline(object):
    __logger = getLogger('PlacesmapPipeline')

    def __init__(self, export_dir: str, sheet_columns: List[str]):
        if export_dir is None or len(export_dir) == 0:
            self.__export_dir = os.getcwd()
        else:
            self.__export_dir = export_dir
        if not os.path.exists(self.__export_dir):
            os.makedirs(self.__export_dir)
        self.__sheet_columns = sheet_columns

    @classmethod
    def from_crawler(cls, crawler):
        export_dir = crawler.settings.get("EXPORT_DIR")
        sheet_columns = crawler.settings.get("SHEET_COLUMNS")
        return cls(export_dir=export_dir, sheet_columns=sheet_columns)

    def open_spider(self, spider):
        self.__logger.info('初始化 ExcelItemExporter...')
        self.__country_name = spider.country_name
        self.__interests = "_".join(spider.interests).replace(" ", "-")
        self.__file_path = self.__export_dir + self.__interests + ".xlsx"
        self.__exporter = ExcelItemExporter(
            filename=self.__file_path,
            sheet_columns={"地点": self.__sheet_columns})
        self.__exporter.start_exporting()

    def process_item(self, item, spider):
        self.__exporter.export_item(item)