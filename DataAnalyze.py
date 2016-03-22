#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from Rank import Rank
import xlsxwriter
import csv
import datetime

keywords = []
websites = []

class RankData(object):
    """Docstring for RankData. """
    def __init__(self):
        """TODO: to be defined1. """
        self.rawRank = dict()
        self.websitesRankCSV = dict()
        self.websitesRankXLSX = []

    def getRanks(self):
        for i in keywords:
            print('processing keyword [{0}] [{1}]'.format(keywords.index(i), i))
            x = Rank(i)
            self.rawRank[i] = x.getRank(30)

        self.countWebsitesCSV()
        self.countWebsitesXLSX()

    def dumpToFile(self, filename):
        a = list()
        [a.append(i) for i in self.rawRank.values()]
        csvrows = list(zip(*a))
        with open(filename, 'w') as csvfile:
            fieldnames = keywords
            writer = csv.DictWriter(csvfile, fieldnames = fieldnames)
            writer.writeheader()
            for i in csvrows:
                t = dict(zip(fieldnames, i))
                writer.writerow(t)

    def countWebsitesCSV(self):
        for site in websites:
            a={'Name': site}
            for kw in self.rawRank.keys():
                rank = 21
                for i in range(len(self.rawRank[kw])):
                    if self.rawRank[kw][i].find(site) >= 0:
                        rank = i + 1
                a[kw] = 20 if rank > 20 else rank
            self.websitesRankCSV[site] = a

    def dumpWebstitesStasticsToCSV(self, filename):
        fieldnames = ['Name']
        [fieldnames.append(i) for i in keywords]
        [print(i) for i in fieldnames]
        with open(filename, 'w') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames = fieldnames)
            writer.writeheader()
            for v in self.websitesRankCSV.values():
                writer.writerow(v)

    def countWebsitesXLSX(self):

        """keywords title"""
        t = ['']
        [t.append(i) for i in keywords]
        t.append('平均值')
        t.append('综合排名')
        self.websitesRankXLSX.append(t)
        average = []
        for site in websites:
            a = [site]
            for kw in keywords:
                rank = 0
                for i in self.rawRank[kw]:#i is website address
                    rank = rank + 1
                    if i.find(site) >= 0:
                        break
                a.append(20 if rank > 20 else rank)
            ave = round(sum(a[1:])/len(a[1:]), 2)#calculate average rank
            a.append(ave)#calculate average rank
            average.append(ave)
            self.websitesRankXLSX.append(a)

        # order by average
        average.sort()
        for l in self.websitesRankXLSX:
            last = l[len(l)-1]
            if isinstance(last, str):
                continue
            try:
                l.append(average.index(last) + 1)
            except ValueError as e:
                print(e)


    def dumpWebstitesStasticsToXLSX(self, filename):
        workbook = xlsxwriter.Workbook(filename)
        worksheet = workbook.add_worksheet()

        title_format = workbook.add_format({
            'bold':     True,
            'align':    'center',
            'valign':   'vcenter',
            'font_size': 20,
        })

        second_row_format = workbook.add_format({
            'bold':     False,
            'align':    'left',
            'valign':   'vbottom',
        })
        x_format = workbook.add_format({
            'bold':     False,
            'border':   1,
            'diag_type':   2,
            'align':    'center',
            'align':   'vcenter',
            'text_wrap':True,
        })
        c_format = workbook.add_format({
            'text_wrap':True,
        })
        c_format.set_align('center')
        c_format.set_align('vcenter')
        c_format.set_border(1)

        c_red_format = workbook.add_format({
            'text_wrap':True,
            'font_color':'red',
        })
        c_red_format.set_align('center')
        c_red_format.set_align('vcenter')
        c_red_format.set_border(1)

        worksheet.merge_range('A1:U1', '百度快照排名统计表', title_format)
        worksheet.set_row(0, 30)
        nowtime = datetime.datetime.now()
        worksheet.merge_range('A2:U2',
                '统计日期：{0}年{1}月{2}日{3}:{4}'.format(nowtime.year, nowtime.month,
                    nowtime.day, nowtime.hour, nowtime.minute),
                second_row_format)
        worksheet.set_row(1, 20)
        worksheet.set_row(2, 20)
        worksheet.set_column('A:A', 12)
        worksheet.set_column('B:B', 2)

        worksheet.merge_range('A3:A4', '      关键词\n网站   ', x_format)
        worksheet.write('B3', '序', c_format)
        for i in range(len(keywords)):
            worksheet.write(2, 2 + i , (i + 1).__str__(), c_format)

        worksheet.write('B4', '词')

        for row in range(len(self.websitesRankXLSX)):
            r = self.websitesRankXLSX[row]
            last_col = len(r) - 1
            for col in range(len(r)):
                if row == 0 and col == 0:
                    worksheet.write(row + 3, col, r[col], x_format)
                else:
                    worksheet.set_row(row + 3, 20)
                    if col == 0:
                        #worksheet.write(row + 3, col + 1, r[col])
                        worksheet.merge_range(row + 3, col, row + 3, col + 1, r[col], c_format)
                    if col == last_col:
                        worksheet.write(row + 3, col + 1, r[col], c_red_format)
                    else:
                        worksheet.write(row + 3, col + 1, r[col], c_format)

                    if r[col] =='平均值' or r[col] == '综合排名':
                        worksheet.merge_range(row + 3, col + 1, row + 2, col + 1, r[col], c_format)

        worksheet.set_row(3, 40)
        last_row = len(self.websitesRankXLSX) + 5
        worksheet.merge_range('A{0}:U{0}'.format(last_row), '注：此数据统计出现在百度快照头两页的排名按实际排名计入，出现在第三页之后的排名均按排名20位计入。')

        workbook.close()

if __name__ == "__main__":
    a = RankData()
    a.getRanks()
    a.dumpWebstitesStasticsToXLSX('w.xlsx')
