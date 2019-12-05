# -*- coding: utf-8 -*-

# Define your item pipelines here
#
# Don't forget to add your pipeline to the ITEM_PIPELINES setting
# See: https://docs.scrapy.org/en/latest/topics/item-pipeline.html
import os, datetime
from logging import getLogger
from placesmap.excel_item_exporter import ExcelItemExporter


class PlacesmapPipeline(object):
    _logger = getLogger('PlacesmapPipeline')

    def __init__(self, export_dir: str):
        if export_dir is None or len(export_dir) == 0:
            self.__export_dir = os.getcwd()
        else:
            self.__export_dir = export_dir
        if not os.path.exists(self.__export_dir):
            os.makedirs(self.__export_dir)
        timestamp = datetime.datetime.now().strftime(r"%Y%m%d_%H%M%S")
        self.__export_file = self.__export_dir + timestamp + ".xlsx"

    @classmethod
    def from_crawler(cls, crawler):
        return cls(export_dir=crawler.settings.get("EXPORT_DIR"))

    def open_spider(self, spider):
        spider.logger.info('Spider opened...')
        self._logger.info('初始化ExcelItemExporter...')
        self.__country_name = spider.country_name
        sheet_columns = {
            '': ['矫正单位', '矫正人员编号', '姓名', '性别', '身份证号', '定位手机', '个人联系电话',
                        '处理等级', '矫正类别', '入矫时间', '社区矫正开始时间', '社区矫正结束时间', '民族',
                        '是否调查评估', '就业就学情况', '婚姻状况', '现政治面貌', '文化程度', '国籍', '民族', '捕前职业',
                        '有无港澳通行证', '有无台胞证', '有无台湾通行证', '有无港澳通行证', '有无港澳居民往来内地通行证',
                        '是否成年', '是否有精神病', '是否有传染病', '是否三无人员', '是否重点人员', '是否居住地变更迁入',
                        '移交罪犯机关', '是否共同犯罪', '是否有前科', '原判审判机关名称', '原判审判机关类型', '是否累犯',
                        '罪名', '犯罪类型', '是否数罪并罚', '原判判决书字号', '原判刑罚', '原判刑期', '附加刑',
                        "是否'五独'", "是否'五涉", "是否有'四史'", '报道情况', '社区矫正人员接受方式', '是否建立矫正小组',
                        '是否被宣告禁止令', '是否建立矫正方案', '是否采用电子定位管理', '是否矫正中止', '是否脱管',
                        '矫正解除类型', '矫正状态', '有无护照', '有无虚拟身份', '执行通知书日期', '交付执行日期',
                        '居住地地址', '户籍地地址', 'ID', '出生日期', '原政治面貌', '社区矫正决定机关', '判决日期',
                        '文书类型', '文书编号', '文书生效日期', '主要犯罪事实', '原判刑期开始日期', '原判刑期结束日期',
                        '县区社区服刑人员接收日期', '司法所社区服刑人员接收日期', '矫正单位ID'],
            '家庭及社会关系': ['关系', '姓名', '所在单位', '家庭住址', '联系电话', '人员ID', '档案关联人员ID', '档案关联人员姓名'],
            '个人简历信息': ['起始日期', '结束日期', '所在单位', '职务', '人员ID', '档案关联人员ID', '档案关联人员姓名'],
            '共同犯罪信息': ['姓名', '性别', '出生日期', '罪名', '被判处刑罚', '所在监所', '人员ID', '档案关联人员ID', '档案关联人员姓名']
        }

            self._exporter = ExcelItemExporter(filename='{export_dir}{basename}.{ext}'
                                               .format(export_dir=export_dir, basename='矫正档案', ext='xlsx'),
                                               sheet_columns=sheet_columns)
            self._exporter.start_exporting()

    def process_item(self, item, spider):
        if isinstance(item, CategoryItem):
            self._exporter.export_item(item)
            return item
        raise DropItem(f"当前Pipeline不处理{type(item)}类型")